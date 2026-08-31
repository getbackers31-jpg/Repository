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

// --- 檔案與路徑輔助函式 ---
function sanitizePathSegment(value) {
    return String(value)
        .replace(/[<>:"/\\|?*#%]/g, '_')
        .replace(/\s+/g, ' ')
        .trim();
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
        console.error(`❌ 讀取 ${filePath} 失敗`, { statusCode, message: error.message });
        
        if (statusCode === 404 || statusCode === 'itemNotFound') {
            if (throwOnNotFound) {
                throw new Error(`找不到必要設定檔: ${filePath}`);
            }
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
    
    return projects.find(project => {
        return project.active === true && String(project.projectName).trim() === normalizedName;
    }) || null;
}

// --- LINE API 輔助函式 ---
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
    if (!response.ok) {
        const body = await response.text();
        console.error(`LINE Reply 失敗 ${response.status}: ${body}`);
    }
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
    if (!response.ok) {
        const body = await response.text();
        throw new Error(`LINE Push 失敗 ${response.status}: ${body}`);
    }
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

function maskLineId(value) {
    if (!value || value.length < 10) return 'unknown';
    return `${value.slice(0, 5)}...${value.slice(-4)}`;
}

// --- 取得台灣時間 ---
function getTaiwanDateParts() {
    const parts = new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Asia/Taipei',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    }).formatToParts(new Date());

    const values = Object.fromEntries(
        parts.filter(part => part.type !== 'literal').map(part => [part.type, part.value])
    );

    return {
        year: values.year,
        month: values.month,
        day: values.day,
        dateStr: `${values.year}-${values.month}-${values.day}`,
        timeStr: `${values.hour}${values.minute}${values.second}`
    };
}

// --- API 資料驗證 ---
function validateReportData(reportData) {
    const requiredFields = ['projectName', 'temp', 'humidity', 'wind', 'contractor', 'workerCount', 'progress', 'materials'];
    const missingFields = requiredFields.filter(field => {
        const value = reportData[field];
        return (value === undefined || value === null || String(value).trim() === '');
    });
    return {
        valid: missingFields.length === 0,
        missingFields
    };
}

// ==========================================
// 1. Webhook 路由
// ==========================================
app.post('/webhook', express.raw({ type: 'application/json' }), async (req, res) => {
    const signature = req.get('x-line-signature');
    if (!verifyLineSignature(req.body, signature)) {
        console.warn('⚠️ LINE Webhook 簽章驗證失敗');
        return res.status(401).send('Invalid signature');
    }

    let body;
    try {
        body = JSON.parse(req.body.toString('utf8'));
    } catch (error) {
        console.error('Webhook JSON 解析失敗：', error);
        return res.status(400).send('Invalid JSON');
    }

    res.status(200).send('OK');

    console.log('📥 收到 LINE Webhook', { eventCount: body.events?.length || 0 });

    for (const event of body.events || []) {
        try {
            const targetId = getLineTargetId(event);
            
            console.log('📌 LINE 事件', {
                type: event.type,
                sourceType: event.source?.type,
                hasGroupId: Boolean(targetId),
                maskedId: maskLineId(targetId)
            });

            if (event.type === 'join') {
                await replyLineMessage(
                    event.replyToken,
                    '👷 云說工程小幫手已加入本群組\n\n請輸入：\n設定案場 案場名稱\n\n例如：\n設定案場 台塑大樓'
                );
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
                        await replyLineMessage(event.replyToken, `⚠️ 找不到這個案場\n\n輸入名稱：${projectName}\n\n請先在系統建立正式案場，\n或確認案場名稱是否正確。`);
                        continue;
                    }

                    const config = await readBindingsFromOneDrive();
                    const bindings = Array.isArray(config.bindings) ? config.bindings : [];
                    const filteredBindings = bindings.filter(b => b.projectId !== project.projectId && b.groupId !== targetId);
                    
                    filteredBindings.push({
                        projectId: project.projectId,
                        projectName: project.projectName,
                        groupId: targetId,
                        sourceType: event.source.type,
                        active: true,
                        boundAt: new Date().toISOString()
                    });

                    await writeBindingsToOneDrive({ bindings: filteredBindings });
                    await replyLineMessage(event.replyToken, `✅ 案場設定完成\n\n本群組案場：${project.projectName}\n\n之後該案場的施工日報，\n將自動發布至本群組。`);
                }

                else if (text === '查詢案場') {
                    if (!targetId) {
                        await replyLineMessage(event.replyToken, '請在施工群組內使用「查詢案場」。');
                        continue;
                    }
                    const config = await readBindingsFromOneDrive();
                    const bindings = Array.isArray(config.bindings) ? config.bindings : [];
                    const binding = bindings.find(item => item.groupId === targetId && item.active === true);
                    if (!binding) {
                        await replyLineMessage(event.replyToken, '⚠️ 本群組尚未設定案場\n\n請輸入：\n設定案場 案場名稱');
                    } else {
                        await replyLineMessage(event.replyToken, `📍 本群組目前設定案場\n\n${binding.projectName}`);
                    }
                }

                else if (text === '解除案場') {
                    if (!targetId) {
                        await replyLineMessage(event.replyToken, '請在施工群組內使用「解除案場」。');
                        continue;
                    }
                    const config = await readBindingsFromOneDrive();
                    const bindings = Array.isArray(config.bindings) ? config.bindings : [];
                    const originalCount = bindings.length;
                    config.bindings = bindings.filter(b => b.groupId !== targetId);

                    if (config.bindings.length === originalCount) {
                        await replyLineMessage(event.replyToken, '本群組目前沒有設定任何案場。');
                    } else {
                        await writeBindingsToOneDrive(config);
                        await replyLineMessage(event.replyToken, '✅ 已解除本群組的案場設定\n\n日報仍會歸檔至 OneDrive，\n但不會再發布至本群組。');
                    }
                }
            }
        } catch (error) {
            console.error('LINE 事件處理失敗：', error);
        }
    }
});

// ==========================================
// 2. 一般 API 路由
// ==========================================
app.use(express.json());

app.get('/', (req, res) => {
    res.send('✅ 伺服器運作中 (第一階段：單群組閉環測試版)！');
});

app.post('/api/submit-report', async (req, res) => {
    try {
        console.log("收到日報提交請求！");
        const reportData = req.body || {}; 
        
        const validation = validateReportData(reportData);
        if (!validation.valid) {
            return res.status(400).json({
                success: false, archived: false, pushed: false,
                reason: 'INVALID_REPORT_DATA',
                error: `缺少必要欄位：${validation.missingFields.join(', ')}`
            });
        }

        const project = await findProjectByName(reportData.projectName);
        if (!project) {
            return res.status(400).json({
                success: false, archived: false, pushed: false,
                reason: 'PROJECT_NOT_FOUND',
                error: '找不到指定案場，請確認中央清單'
            });
        }

        const graphClient = await getGraphClient();
        const { year, month, dateStr, timeStr } = getTaiwanDateParts();
        
        // 👉 極簡版路徑：直接存進案場資料夾
        const safeProjectName = sanitizePathSegment(project.projectName);
        const targetFolderPath = `工程專案管理/${safeProjectName}`;
        const fileName = `${dateStr}_${timeStr}_施工日報.txt`;

        let reportText = `📋 施工日報\n\n`;
        reportText += `日期：${dateStr.replace(/-/g, '/')}\n`;
        reportText += `案場：${project.projectName}\n\n`;
        reportText += `溫度：${reportData.temp}度\n`;
        reportText += `濕度：${reportData.humidity}%\n`;
        reportText += `風速：${reportData.wind}m/s\n\n`;
        reportText += `施工廠商：${reportData.contractor}\n`;
        reportText += `施工人數：${reportData.workerCount}\n\n`;
        reportText += `━━━━━━━━━━━━\n\n`;
        reportText += `今日作業進度：\n${reportData.progress}\n\n`;
        reportText += `今日用料：\n${reportData.materials}\n\n`;
        reportText += `備註：\n${reportData.remarks || '無'}\n\n`;
        reportText += `━━━━━━━━━━━━\n以上為今日進度報告`;

        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`).put(reportText);
        console.log("✅ 文字日報寫入 OneDrive 成功！");

        const config = await readBindingsFromOneDrive();
        const bindings = Array.isArray(config.bindings) ? config.bindings : [];
        const binding = bindings.find(b => b.projectName === project.projectName && b.active === true);

        if (!binding) {
            return res.status(200).json({
                success: true, archived: true, pushed: false,
                reason: 'PROJECT_NOT_BOUND', message: '日報已歸檔，但案場尚未設定 LINE 群組'
            });
        }

        try {
            await pushLineMessage(binding.groupId, reportText);
            return res.status(200).json({
                success: true, archived: true, pushed: true,
                message: '日報已歸檔並發布至施工群組'
            });
        } catch (lineError) {
            console.error('LINE 發布失敗：', lineError);
            return res.status(200).json({
                success: true, archived: true, pushed: false,
                reason: 'LINE_PUSH_FAILED', message: '日報已歸檔，但 LINE 發布失敗'
            });
        }
    } catch (error) {
        console.error("❌ 系統處理錯誤:", error);
        if (!res.headersSent) {
            return res.status(500).json({ success: false, error: error.message });
        }
    }
});

// ==========================================
// 3. 環境變數檢查與啟動
// ==========================================
const requiredEnvironmentVariables = [
    'LINE_ACCESS_TOKEN',
    'LINE_CHANNEL_SECRET',
    'AZURE_CLIENT_ID',
    'AZURE_TENANT_ID',
    'AZURE_CLIENT_SECRET'
];

console.log('環境變數檢查：', {
    LINE_ACCESS_TOKEN: Boolean(LINE_ACCESS_TOKEN),
    LINE_CHANNEL_SECRET: Boolean(LINE_CHANNEL_SECRET),
    AZURE_CLIENT_ID: Boolean(process.env.AZURE_CLIENT_ID),
    AZURE_TENANT_ID: Boolean(process.env.AZURE_TENANT_ID),
    AZURE_CLIENT_SECRET: Boolean(process.env.AZURE_CLIENT_SECRET)
});

const missingEnvironmentVariables = requiredEnvironmentVariables.filter(name => !process.env[name]);

if (missingEnvironmentVariables.length > 0) {
    console.error('❌ 缺少必要環境變數：', missingEnvironmentVariables.join(', '));
    process.exit(1); 
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 伺服器運作中：http://localhost:${PORT}`);
});
