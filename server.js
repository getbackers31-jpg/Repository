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
const LIFF_ID = process.env.LIFF_ID || '2011289657-vQgMb0eI';
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

function normalizeProjectName(value) {
    return String(value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
}

function validateProjectName(projectName) {
    const normalizedName = normalizeProjectName(projectName);
    if (!normalizedName) throw new Error('案場名稱不可為空');
    if (normalizedName.length > 80) throw new Error('案場名稱不可超過 80 個字');
    if (/[<>:"/\\|?*#%]/.test(normalizedName)) {
        throw new Error('案場名稱不可包含以下字元：< > : " / \\ | ? * # %');
    }
    return normalizedName;
}

function getProjectRegistrationErrorMessage(error) {
    const safeMessages = [
        '案場名稱不可為空',
        '案場名稱不可超過 80 個字',
        '案場名稱不可包含以下字元'
    ];
    const message = String(error?.message || '');
    const isSafeMessage = safeMessages.some(prefix => message.startsWith(prefix));
    return isSafeMessage ? message : '系統暫時無法建立案場，請稍後再試';
}

let projectWriteQueue = Promise.resolve();
function withProjectWriteLock(task) {
    const result = projectWriteQueue.then(task, task);
    projectWriteQueue = result.catch(() => undefined);
    return result;
}

let bindingWriteQueue = Promise.resolve();
function withBindingWriteLock(task) {
    const result = bindingWriteQueue.then(task, task);
    bindingWriteQueue = result.catch(() => undefined);
    return result;
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
async function writeProjectsToOneDrive(config) {
    await writeJsonToOneDrive('工程專案管理/_系統設定/projects.json', config);
}

async function readBindingsFromOneDrive() {
    return await readJsonFromOneDrive('工程專案管理/_系統設定/line-bindings.json', { bindings: [] });
}
async function writeBindingsToOneDrive(config) {
    await writeJsonToOneDrive('工程專案管理/_系統設定/line-bindings.json', config);
}

async function ensureProjectFolder(projectName) {
    const graphClient = await getGraphClient();
    const safeProjectName = sanitizePathSegment(projectName);
    if (!safeProjectName) throw new Error('案場資料夾名稱不可為空');

    const folderPath = `工程專案管理/2026_工程專案/${safeProjectName}`;

    try {
        const item = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${folderPath}`).get();
        if (!item.folder) throw new Error(`同名項目不是資料夾：${safeProjectName}`);
        return { created: false, folderId: item.id, folderPath };
    } catch (error) {
        const statusCode = error?.statusCode || error?.status || error?.code;
        if (statusCode !== 404 && statusCode !== 'itemNotFound') throw error;
    }

    try {
        const createdFolder = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/工程專案管理/2026_工程專案:/children`).post({
            name: safeProjectName,
            folder: {},
            '@microsoft.graph.conflictBehavior': 'fail'
        });
        return { created: true, folderId: createdFolder.id, folderPath };
    } catch (error) {
        const item = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${folderPath}`).get();
        if (!item.folder) throw error;
        return { created: false, folderId: item.id, folderPath };
    }
}

// 建立子資料夾用的函式
async function ensureChildFolder(graphClient, parentPath, childFolderName) {
    const safeChildName = sanitizePathSegment(childFolderName);
    if (!safeChildName) throw new Error('子資料夾名稱不可為空');
    const childPath = `${parentPath}/${safeChildName}`;

    try {
        const existingItem = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${childPath}`).get();
        if (!existingItem.folder) throw new Error(`同名項目不是資料夾：${childPath}`);
        return { created: false, folderId: existingItem.id, folderPath: childPath };
    } catch (error) {
        const statusCode = error?.statusCode || error?.status || error?.code;
        if (statusCode !== 404 && statusCode !== 'itemNotFound') throw error;
    }

    try {
        const createdFolder = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${parentPath}:/children`).post({
            name: safeChildName,
            folder: {},
            '@microsoft.graph.conflictBehavior': 'fail'
        });
        return { created: true, folderId: createdFolder.id, folderPath: childPath };
    } catch (createError) {
        const existingItem = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${childPath}`).get();
        if (!existingItem.folder) throw createError;
        return { created: false, folderId: existingItem.id, folderPath: childPath };
    }
}

async function findProjectByName(projectName) {
    const config = await readProjectsFromOneDrive();
    const projects = Array.isArray(config.projects) ? config.projects : [];
    const normalizedName = normalizeProjectName(projectName);
    return projects.find(p => p.active === true && normalizeProjectName(p.projectName) === normalizedName) || null;
}

async function findProjectById(projectId) {
    const config = await readProjectsFromOneDrive();
    const projects = Array.isArray(config.projects) ? config.projects : [];
    const normalizedProjectId = String(projectId || '').trim();
    if (!normalizedProjectId) {
        return null;
    }
    return projects.find(project => {
        return project.active === true && project.projectId === normalizedProjectId;
    }) || null;
}

function createProjectId() {
    return `PRJ-${crypto.randomUUID()}`;
}

async function registerProjectByName(projectName) {
    return withProjectWriteLock(async () => {
        const normalizedName = validateProjectName(projectName);

        const config = await readProjectsFromOneDrive();
        const projects = Array.isArray(config.projects) ? config.projects : [];

        let existingProject = projects.find(p => p.active === true && normalizeProjectName(p.projectName) === normalizedName);

        if (existingProject) {
            await ensureProjectFolder(existingProject.projectName);
            return { project: existingProject, created: false };
        }

        const project = {
            projectId: createProjectId(),
            projectName: normalizedName,
            active: true,
            createdAt: new Date().toISOString()
        };

        await ensureProjectFolder(project.projectName);

        projects.push(project);
        await writeProjectsToOneDrive({
            ...config,
            projects,
            updatedAt: new Date().toISOString()
        });

        return { project, created: true };
    });
}

async function replyLineMessage(replyToken, text) {
    if (!replyToken) return;
    const response = await fetch('https://api.line.me/v2/bot/message/reply', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${LINE_ACCESS_TOKEN}` },
        body: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] })
    });
    const responseBody = await response.text();
    if (!response.ok) throw new Error(`LINE Reply 失敗 ${response.status}: ${responseBody}`);
}

async function pushLineMessage(targetId, text) {
    const response = await fetch('https://api.line.me/v2/bot/message/push', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${LINE_ACCESS_TOKEN}` },
        body: JSON.stringify({ to: targetId, messages: [{ type: 'text', text }] })
    });
    const responseBody = await response.text();
    if (!response.ok) throw new Error(`LINE Push 失敗 ${response.status}: ${responseBody}`);
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

app.get('/api/projects', async (req, res) => {
    try {
        const config = await readProjectsFromOneDrive();
        const projects = Array.isArray(config.projects) ? config.projects : [];
        const activeProjects = projects
            .filter(p => p.active === true)
            .map(p => ({ projectId: p.projectId, projectName: p.projectName }))
            .sort((a, b) => a.projectName.localeCompare(b.projectName, 'zh-Hant'));

        return res.status(200).json({ success: true, projects: activeProjects });
    } catch (error) {
        console.error('讀取案場清單失敗：', error);
        return res.status(500).json({ success: false, error: '無法取得案場清單' });
    }
});

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
                const welcomeText = [
                    '👷 歡迎使用「云說工程小幫手」！',
                    '我是負責協助自動化建案與日報歸檔的機器人。將我邀請至施工群組後，請依以下步驟啟用專屬日報：',
                    '',
                    '【📝 日常填寫日報 3 步驟】',
                    '1️⃣ 首次開工請先輸入「設定案場 案場名稱」。',
                    '2️⃣ 將機器人回覆的專屬網址「設為群組置頂公告」。',
                    '3️⃣ 以後只需點擊群組公告，即可自動鎖定案場填寫日報！'
                ].join('\n');
                await replyLineMessage(event.replyToken, welcomeText);
                continue;
            }
            if (event.type === 'message' && event.message.type === 'text') {
                const text = event.message.text.trim();
                
                // 🔍 除錯雷達：在 Render 後台印出 LINE 傳來的所有文字，幫你抓出空白或錯字
                console.log(`[Webhook 收到訊息] 內容: "${text}", 來源類型: ${event.source.type}, targetId: ${targetId}`);

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
                    
                    let registration;
                    try {
                        registration = await registerProjectByName(projectName);
                    } catch (error) {
                        console.error('建立案場失敗：', error);
                        await replyLineMessage(event.replyToken, `⚠️ 無法建立案場\n\n${getProjectRegistrationErrorMessage(error)}`);
                        continue;
                    }

                    const project = registration.project;

                    await withBindingWriteLock(async () => {
                        const config = await readBindingsFromOneDrive();
                        const bindings = Array.isArray(config.bindings) ? config.bindings : [];
                        const filteredBindings = bindings.filter(b => b.projectId !== project.projectId && b.groupId !== targetId);
                        filteredBindings.push({
                            projectId: project.projectId, projectName: project.projectName, groupId: targetId,
                            sourceType: event.source.type, active: true, boundAt: new Date().toISOString()
                        });
                        await writeBindingsToOneDrive({ ...config, bindings: filteredBindings, updatedAt: new Date().toISOString() });
                    });

                    const statusText = registration.created ? '已建立新案場及 OneDrive 資料夾' : '已使用現有案場';
                    const reportUrl = `https://liff.line.me/${LIFF_ID}/?projectId=${encodeURIComponent(project.projectId)}`;
                    await replyLineMessage(
                        event.replyToken,
                        [
                            '✅ 案場設定完成',
                            '',
                            `案場：${project.projectName}`,
                            `狀態：${statusText}`,
                            '',
                            '請使用以下專屬連結填寫施工日報：',
                            reportUrl,
                            '',
                            '此連結已綁定本案場，建議將這則訊息設為群組公告。'
                        ].join('\n')
                    );
                }
                else if (text === '查詢案場' || text === '案場查詢') {
                    if (!targetId) {
                        await replyLineMessage(event.replyToken, '⚠️ 請在施工群組內使用「查詢案場」指令。');
                        continue;
                    }
                    const config = await readBindingsFromOneDrive();
                    const binding = (Array.isArray(config.bindings) ? config.bindings : []).find(b => b.groupId === targetId && b.active);
                    await replyLineMessage(event.replyToken, binding ? `📍 本群組目前設定案場\n\n${binding.projectName}` : '⚠️ 本群組尚未設定案場\n\n請輸入：\n設定案場 案場名稱');
                }
                else if (text === '解除案場') {
                    if (!targetId) {
                        await replyLineMessage(event.replyToken, '⚠️ 請在施工群組內使用「解除案場」指令。');
                        continue;
                    }
                    await withBindingWriteLock(async () => {
                        const config = await readBindingsFromOneDrive();
                        const bindings = Array.isArray(config.bindings) ? config.bindings : [];
                        const filteredBindings = bindings.filter(b => b.groupId !== targetId);
                        
                        if (filteredBindings.length === bindings.length) {
                            await replyLineMessage(event.replyToken, '本群組目前沒有設定任何案場。');
                            return;
                        }

                        await writeBindingsToOneDrive({ ...config, bindings: filteredBindings, updatedAt: new Date().toISOString() });
                        await replyLineMessage(event.replyToken, '✅ 已解除本群組的案場設定。');
                    });
                }
                // 👇 全新的結案指令 👇
                else if (text.startsWith('結案')) {
                    if (!targetId) {
                        await replyLineMessage(event.replyToken, '⚠️ 請在施工群組內使用「結案」指令。');
                        continue;
                    }
                    const match = text.match(/^結案\s+(.+)$/);
                    if (!match) {
                        await replyLineMessage(event.replyToken, '⚠️ 指令格式錯誤\n\n正確格式：\n結案 大安區');
                        continue;
                    }
                    const targetProjectName = match[1].trim();

                    await withProjectWriteLock(async () => {
                        const config = await readProjectsFromOneDrive();
                        const projects = Array.isArray(config.projects) ? config.projects : [];
                        const normalizedTarget = normalizeProjectName(targetProjectName);

                        const projectIndex = projects.findIndex(p => normalizeProjectName(p.projectName) === normalizedTarget);

                        if (projectIndex === -1) {
                            await replyLineMessage(event.replyToken, `⚠️ 找不到名為「${targetProjectName}」的案場，可能已經結案或名稱輸入錯誤。`);
                            return;
                        }

                        // 從陣列中徹底刪除
                        projects.splice(projectIndex, 1);
                        await writeProjectsToOneDrive({ ...config, projects, updatedAt: new Date().toISOString() });

                        // 同步解除相關群組綁定
                        await withBindingWriteLock(async () => {
                            const bindingConfig = await readBindingsFromOneDrive();
                            const bindings = Array.isArray(bindingConfig.bindings) ? bindingConfig.bindings : [];
                            const filteredBindings = bindings.filter(b => b.projectName !== targetProjectName);
                            await writeBindingsToOneDrive({ ...bindingConfig, bindings: filteredBindings, updatedAt: new Date().toISOString() });
                        });

                        await replyLineMessage(
                            event.replyToken,
                            [
                                `✅ 案場「${targetProjectName}」已成功結案！`,
                                '',
                                '系統已自動將此案場從選單中下架，並解除相關的群組綁定。',
                                '您現在可以安心將 OneDrive 中的資料夾移動至「已完工」目錄了。'
                            ].join('\n')
                        );
                    });
                }
                else if (['指令', '說明', '功能', '小幫手', '【點此查看指令說明】'].includes(text)) {
                    const helpText = [
                        '📖 「云說工程小幫手」群組指令說明',
                        '',
                        '🔹 設定案場 案場名稱',
                        '🔹 查詢案場',
                        '🔹 解除案場',
                        '🔹 結案 案場名稱 (徹底下架選單)'
                    ].join('\n');
                    await replyLineMessage(event.replyToken, helpText);
                }
            }
        } catch (error) { console.error('LINE 事件處理失敗：', error); }
    }
});

app.use(express.json());

app.get('/api/projects/:projectId', async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId);
        if (!project) return res.status(404).json({ success: false, reason: 'PROJECT_NOT_FOUND', error: '找不到指定案場' });
        return res.status(200).json({ success: true, project: { projectId: project.projectId, projectName: project.projectName } });
    } catch (error) {
        console.error('讀取指定案場失敗：', error);
        return res.status(500).json({ success: false, reason: 'INTERNAL_ERROR', error: '無法取得案場資料' });
    }
});

app.get('/', (req, res) => res.send('✅ 伺服器運作中！'));

app.post('/api/submit-report', async (req, res) => {
    try {
        const reportData = req.body || {}; 
        const validation = validateReportData(reportData);
        if (!validation.valid) return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'INVALID_REPORT_DATA', error: `缺少必要欄位：${validation.missingFields.join(', ')}` });

        const submittedProjectId = String(reportData.projectId || '').trim();
        let project;
        if (submittedProjectId) {
            project = await findProjectById(submittedProjectId);
        } else {
            project = await findProjectByName(reportData.projectName);
        }
        
        if (!project) {
            return res.status(400).json({
                success: false, archived: false, pushed: false,
                reason: 'PROJECT_NOT_FOUND',
                error: '找不到指定案場，請重新從施工群組的專屬連結開啟表單'
            });
        }

        // ==========================================
        // 👇 資料夾建立與 JSON 結構化落檔邏輯 👇
        // ==========================================
        await ensureProjectFolder(project.projectName);
        const safeProjectName = sanitizePathSegment(project.projectName);
        const graphClient = await getGraphClient();
        
        const projectFolderPath = `工程專案管理/2026_工程專案/${safeProjectName}`;
        const textFolderResult = await ensureChildFolder(graphClient, projectFolderPath, '施工日報');
        const dataFolderResult = await ensureChildFolder(graphClient, projectFolderPath, '結構化資料');
        
        const textFolderPath = textFolderResult.folderPath;
        const dataFolderPath = dataFolderResult.folderPath;

        const { dateStr, timeStr } = getTaiwanDateParts();
        const reportDate = dateStr; 

        const fullSubmissionId = String(reportData.submissionId || crypto.randomUUID()).trim();
        const shortSubmissionId = fullSubmissionId.replace(/[^a-zA-Z0-9]/g, '').substring(0, 8);

        const isNoWork = reportData.isNoWork === true;
        const contractorItems = Array.isArray(reportData.contractorItems) ? reportData.contractorItems : [];
        const workItems = Array.isArray(reportData.workItems) ? reportData.workItems : [];
        const materialItems = Array.isArray(reportData.materialItems) ? reportData.materialItems : [];

        const calculatedTotalWorkerCount = isNoWork ? 0 : contractorItems.reduce((total, item) => {
            const workerCount = Number(item.workerCount || 0);
            return total + (Number.isFinite(workerCount) ? workerCount : 0);
        }, 0);

        const structuredReport = {
            schemaVersion: 1,
            projectId: project.projectId,
            projectName: project.projectName,
            reportDate,
            submissionId: fullSubmissionId,
            submittedAt: new Date().toISOString(),
            submittedDateLocal: dateStr,
            submittedTimeLocal: timeStr,
            isNoWork,
            noWorkReason: isNoWork ? String(reportData.noWorkReason || '') : '',
            weather: {
                temp: reportData.temp,
                humidity: reportData.humidity,
                wind: reportData.wind
            },
            contractorItems: isNoWork ? [] : contractorItems,
            totalWorkerCount: calculatedTotalWorkerCount,
            workItems: isNoWork ? [] : workItems,
            customWorkItem: isNoWork ? '' : String(reportData.customWorkItem || ''),
            workNotes: isNoWork ? '' : String(reportData.workNotes || ''),
            materialItems: isNoWork ? [] : materialItems,
            remarks: String(reportData.remarks || '')
        };

        const baseFileName = `${reportDate}_${shortSubmissionId}`;
        const jsonFileName = `${baseFileName}.json`;
        const txtFileName = `${baseFileName}_施工日報.txt`;
        const jsonFilePath = `${dataFolderPath}/${jsonFileName}`;
        const txtFilePath = `${textFolderPath}/${txtFileName}`;

        try {
            const jsonFileContent = Buffer.from(JSON.stringify(structuredReport, null, 2), 'utf-8');
            await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${jsonFilePath}:/content`).put(jsonFileContent);
            console.log(`[Success] 結構化 JSON 已寫入：${jsonFilePath}`);
        } catch (jsonUploadError) {
            console.error('[Error] 結構化 JSON 寫入失敗：', jsonUploadError);
            throw new Error('結構化日報寫入失敗');
        }

        // ==========================================
        // 👇 TXT 與 LINE 推播邏輯 👇
        // ==========================================
        let reportText = `📋 施工日報\n\n日期：${reportDate.replace(/-/g, '/')}\n案場：${project.projectName}\n\n`;
        reportText += `溫度：${reportData.temp}度\n濕度：${reportData.humidity}%\n風速：${reportData.wind}m/s\n\n`;
        reportText += `施工廠商：${reportData.contractor}\n施工人數：${reportData.workerCount}\n\n━━━━━━━━━━━━\n\n`;
        reportText += `今日作業進度：\n${reportData.progress}\n\n今日用料：\n${reportData.materials}\n\n備註：\n${reportData.remarks || '無'}\n\n━━━━━━━━━━━━\n以上為今日進度報告`;

        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${txtFilePath}:/content`).put(reportText);

        const config = await readBindingsFromOneDrive();
        const binding = (Array.isArray(config.bindings) ? config.bindings : []).find(b => b.projectId === project.projectId && b.active);

        if (!binding) return res.status(200).json({ success: true, archived: true, pushed: false, reason: 'PROJECT_NOT_BOUND' });

        try {
            await pushLineMessage(binding.groupId, reportText);
            return res.status(200).json({ success: true, archived: true, pushed: true, message: '日報已歸檔並發布' });
        } catch (lineError) {
            console.error('LINE 推播失敗：', lineError);
            return res.status(200).json({ success: true, archived: true, pushed: false, reason: 'LINE_PUSH_FAILED' });
        }
    } catch (error) {
        console.error('提交日報失敗：', error);
        if (!res.headersSent) {
            return res.status(500).json({
                success: false, archived: false, pushed: false,
                reason: 'INTERNAL_ERROR',
                error: '系統處理失敗，請稍後再試'
            });
        }
    }
});

const requiredVars = ['LINE_ACCESS_TOKEN', 'LINE_CHANNEL_SECRET', 'AZURE_CLIENT_ID', 'AZURE_TENANT_ID', 'AZURE_CLIENT_SECRET'];
if (requiredVars.some(v => !process.env[v])) process.exit(1);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 伺服器運作中：http://localhost:${PORT}`);
});
