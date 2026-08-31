require('dotenv').config();
const express = require('express');
const cors = require('cors');
const crypto = require('crypto');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(cors());

const LINE_ACCESS_TOKEN = process.env.LINE_ACCESS_TOKEN;
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const TARGET_USER_EMAIL = "kate@cyber-cloud.info"; 

const msalConfig = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET
    }
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getGraphClient() {
    const tokenRequest = { scopes: ['https://graph.microsoft.com/.default'] };
    try {
        const response = await cca.acquireTokenByClientCredential(tokenRequest);
        return Client.init({ authProvider: (done) => { done(null, response.accessToken); } });
    } catch (error) { throw error; }
}

function sanitizePathSegment(value) {
    return String(value).replace(/[<>:"/\\|?*#%]/g, '_').replace(/\s+/g, ' ').trim();
}

async function readJsonFromOneDrive(filePath, defaultData, throwOnNotFound = false) {
    try {
        const graphClient = await getGraphClient();
        const meta = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${filePath}`).get();
        const downloadUrl = meta['@microsoft.graph.downloadUrl'];
        if (!downloadUrl) throw new Error('Graph 未回傳檔案下載網址');
        
        const response = await fetch(downloadUrl);
        if (!response.ok) throw new Error(`下載設定檔失敗 ${response.status}`);
        return await response.json();
    } catch (error) {
        const statusCode = error?.statusCode || error?.status || error?.code;
        if (statusCode === 404 || statusCode === 'itemNotFound') {
            if (throwOnNotFound) throw new Error(`找不到必要設定檔: ${filePath}`);
            return defaultData;
        }
        throw error;
    }
}

async function writeJsonToOneDrive(filePath, data) {
    const graphClient = await getGraphClient();
    await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${filePath}:/content`)
        .put(JSON.stringify(data, null, 2));
}

async function readProjectsFromOneDrive() {
    return await readJsonFromOneDrive('工程專案管理/_系統設定/projects.json', { projects: [] }, true);
}
async function readBindingsFromOneDrive() {
    return await readJsonFromOneDrive('工程專案管理/_系統設定/line-bindings.json', { bindings: [] });
}
async function writeBindingsToOneDrive(config) {
    await writeJsonToOneDrive('工程專案管理/_系統設定/line-bindings.json', config);
}

async function findProjectByName(projectName) {
    const config = await readProjectsFromOneDrive();
    const projects = Array.isArray(config.projects) ? config.projects : [];
    const normalizedName = String(projectName).trim();
    return projects.find(p => p.active === true && String(p.projectName).trim() === normalizedName) || null;
}

async function replyLineMessage(replyToken, text) {
    if (!replyToken) return;
    const response = await fetch('https://api.line.me/v2/bot/message/reply', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${LINE_ACCESS_TOKEN}`
        },
        body: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] })
    });
    if (!response.ok) console.error(`LINE Reply 失敗 ${response.status}`);
}

async function pushLineMessage(targetId, text) {
    const response = await fetch('https://api.line.me/v2/bot/message/push', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${LINE_ACCESS_TOKEN}`
        },
        body: JSON.stringify({ to: targetId, messages: [{ type: 'text', text }] })
    });
    if (!response.ok) throw new Error(`LINE Push 失敗 ${response.status}`);
}

function verifyLineSignature(rawBody, signature) {
    if (!LINE_CHANNEL_SECRET || !signature) return false;
    const expectedSignature = crypto.createHmac('sha256', LINE_CHANNEL_SECRET).update(rawBody).digest('base64');
    const actualBuffer = Buffer.from(signature);
    const expectedBuffer = Buffer.from(expectedSignature);
    if (actualBuffer.length !== expectedBuffer.length) return false;
    return crypto.timingSafeEqual(actualBuffer, expectedBuffer);
}

function getLineTargetId(event) {
    if (event.source?.type === 'group') return event.source.groupId;
    if (event.source?.type === 'room') return event.source.roomId;
    return null;
}

function getTaiwanDateParts() {
    const parts = new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
    }).formatToParts(new Date());
    const values = Object.fromEntries(parts.filter(p => p.type !== 'literal').map(p => [p.type, p.value]));
    return {
        dateStr: `${values.year}-${values.month}-${values.day}`,
        timeStr: `${values.hour}${values.minute}${values.second}`
    };
}

function validateReportData(reportData) {
    const requiredFields = ['projectName', 'contractor', 'workerCount', 'progress', 'materials'];
    const missingFields = requiredFields.filter(field => !reportData[field] || String(reportData[field]).trim() === '');
    return { valid: missingFields.length === 0, missingFields };
}

app.post('/webhook', express.raw({ type: 'application/json' }), async (req, res) => {
    const signature = req.get('x-line-signature');
    if (!verifyLineSignature(req.body, signature)) return res.status(401).send('Invalid signature');
    
    let body;
    try { body = JSON.parse(req.body.toString('utf8')); } 
    catch (error) { return res.status(400).send('Invalid JSON'); }
    
    res.status(200).send('OK');

    for (const event of body.events || []) {
        try {
            const targetId = getLineTargetId(event);
            if (event.type === 'join') {
                await replyLineMessage(event.replyToken, '👷 云說工程小幫手已加入本群組\n\n請輸入：\n設定案場 案場名稱\n\n例如：\n設定案場 台塑大樓');
                continue;
            }
            if (event.type === 'message' && event.message.type === 'text') {
                const text = event.message.text.trim();
                if (text.startsWith('設定案場')) {
                    if (!targetId) {
                        await replyLineMessage(event.replyToken, '⚠️ 請在施工群組內使用「設定案場」指令。');
                        continue;
                    }
                    const match = text.match(/^設定案場\s+(.+)$/);
                    if (!match) {
                        await replyLineMessage(event.replyToken, '⚠️ 指令格式錯誤\n\n正確格式：\n設定案場 台塑大樓');
                        continue;
                    }
                    const projectName = match[1].trim();
                    const project = await findProjectByName(projectName);
                    if (!project) {
                        await replyLineMessage(event.replyToken, `⚠️ 找不到這個案場\n\n輸入名稱：${projectName}\n\n請先在系統建立正式案場。`);
                        continue;
                    }
                    const config = await readBindingsFromOneDrive();
                    const bindings = Array.isArray(config.bindings) ? config.bindings : [];
                    const filteredBindings = bindings.filter(b => b.projectId !== project.projectId && b.groupId !== targetId);
                    filteredBindings.push({
                        projectId: project.projectId, projectName: project.projectName, groupId: targetId,
                        sourceType: event.source.type, active: true, boundAt: new Date().toISOString()
                    });
                    await writeBindingsToOneDrive({ bindings: filteredBindings });
                    await replyLineMessage(event.replyToken, `✅ 案場設定完成\n\n本群組案場：${project.projectName}`);
                }
                else if (text === '查詢案場') {
                    if (!targetId) continue;
                    const config = await readBindingsFromOneDrive();
                    const binding = (Array.isArray(config.bindings) ? config.bindings : []).find(b => b.groupId === targetId && b.active);
                    await replyLineMessage(event.replyToken, binding ? `📍 本群組目前設定案場\n\n${binding.projectName}` : '⚠️ 本群組尚未設定案場\n\n請輸入：\n設定案場 案場名稱');
                }
                else if (text === '解除案場') {
                    if (!targetId) continue;
                    const config = await readBindingsFromOneDrive();
                    const originalCount = config.bindings.length;
                    config.bindings = config.bindings.filter(b => b.groupId !== targetId);
                    if (config.bindings.length !== originalCount) {
                        await writeBindingsToOneDrive(config);
                        await replyLineMessage(event.replyToken, '✅ 已解除本群組的案場設定。');
                    }
                }
            }
        } catch (error) { console.error('LINE 事件處理失敗：', error); }
    }
});

app.use(express.json());

app.get('/', (req, res) => res.send('✅ 伺服器運作中！'));

app.post('/api/submit-report', async (req, res) => {
    try {
        const reportData = req.body || {}; 
        const validation = validateReportData(reportData);
        if (!validation.valid) return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'INVALID_REPORT_DATA', error: `缺少必要欄位：${validation.missingFields.join(', ')}` });

        const project = await findProjectByName(reportData.projectName);
        if (!project) return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'PROJECT_NOT_FOUND', error: '找不到指定案場' });

        const graphClient = await getGraphClient();
        const { dateStr, timeStr } = getTaiwanDateParts();
        const targetFolderPath = `工程專案管理/${sanitizePathSegment(project.projectName)}`;
        const fileName = `${dateStr}_${timeStr}_施工日報.txt`;

        let reportText = `📋 施工日報\n\n日期：${dateStr.replace(/-/g, '/')}\n案場：${project.projectName}\n\n`;
        reportText += `溫度：${reportData.temp}度\n濕度：${reportData.humidity}%\n風速：${reportData.wind}m/s\n\n`;
        reportText += `施工廠商：${reportData.contractor}\n施工人數：${reportData.workerCount}\n\n━━━━━━━━━━━━\n\n`;
        reportText += `今日作業進度：\n${reportData.progress}\n\n今日用料：\n${reportData.materials}\n\n備註：\n${reportData.remarks || '無'}\n\n━━━━━━━━━━━━\n以上為今日進度報告`;

        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`).put(reportText);

        const config = await readBindingsFromOneDrive();
        const binding = (Array.isArray(config.bindings) ? config.bindings : []).find(b => b.projectName === project.projectName && b.active);

        if (!binding) return res.status(200).json({ success: true, archived: true, pushed: false, reason: 'PROJECT_NOT_BOUND' });

        try {
            await pushLineMessage(binding.groupId, reportText);
            return res.status(200).json({ success: true, archived: true, pushed: true, message: '日報已歸檔並發布' });
        } catch (lineError) {
            return res.status(200).json({ success: true, archived: true, pushed: false, reason: 'LINE_PUSH_FAILED' });
        }
    } catch (error) {
        if (!res.headersSent) return res.status(500).json({ success: false, error: error.message });
    }
});

const requiredVars = ['LINE_ACCESS_TOKEN', 'LINE_CHANNEL_SECRET', 'AZURE_CLIENT_ID', 'AZURE_TENANT_ID', 'AZURE_CLIENT_SECRET'];
if (requiredVars.some(v => !process.env[v])) process.exit(1);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 伺服器運作中：http://localhost:${PORT}`);
});
