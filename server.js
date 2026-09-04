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
const STATS_API_KEY = process.env.STATS_API_KEY;

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
    const safeMessages = ['案場名稱不可為空', '案場名稱不可超過 80 個字', '案場名稱不可包含以下字元'];
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

async function readProjectsFromOneDrive() { return await readJsonFromOneDrive('工程專案管理/_系統設定/projects.json', { projects: [] }, true); }
async function writeProjectsToOneDrive(config) { await writeJsonToOneDrive('工程專案管理/_系統設定/projects.json', config); }
async function readBindingsFromOneDrive() { return await readJsonFromOneDrive('工程專案管理/_系統設定/line-bindings.json', { bindings: [] }); }
async function writeBindingsToOneDrive(config) { await writeJsonToOneDrive('工程專案管理/_系統設定/line-bindings.json', config); }

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
            name: safeProjectName, folder: {}, '@microsoft.graph.conflictBehavior': 'fail'
        });
        return { created: true, folderId: createdFolder.id, folderPath };
    } catch (error) {
        const item = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${folderPath}`).get();
        if (!item.folder) throw error;
        return { created: false, folderId: item.id, folderPath };
    }
}

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
            name: safeChildName, folder: {}, '@microsoft.graph.conflictBehavior': 'fail'
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
    if (!normalizedProjectId) return null;
    return projects.find(project => project.active === true && project.projectId === normalizedProjectId) || null;
}

function createProjectId() { return `PRJ-${crypto.randomUUID()}`; }

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

        const project = { projectId: createProjectId(), projectName: normalizedName, active: true, createdAt: new Date().toISOString() };
        await ensureProjectFolder(project.projectName);
        projects.push(project);
        await writeProjectsToOneDrive({ ...config, projects, updatedAt: new Date().toISOString() });
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
    if (!response.ok) throw new Error(`LINE Reply 失敗 ${response.status}: ${await response.text()}`);
}

async function pushLineMessage(targetId, text) {
    const response = await fetch('https://api.line.me/v2/bot/message/push', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${LINE_ACCESS_TOKEN}` },
        body: JSON.stringify({ to: targetId, messages: [{ type: 'text', text }] })
    });
    if (!response.ok) throw new Error(`LINE Push 失敗 ${response.status}: ${await response.text()}`);
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
    return { dateStr: `${values.year}-${values.month}-${values.day}`, timeStr: `${values.hour}${values.minute}${values.second}` };
}

function validateReportData(reportData) {
    const requiredFields = ['projectName', 'contractor', 'workerCount', 'progress', 'materials'];
    const missingFields = requiredFields.filter(field => !reportData[field] || String(reportData[field]).trim() === '');
    return { valid: missingFields.length === 0, missingFields };
}

// ==========================================
// 👇 核心統計引擎 (可供 API 與結案指令共用) 👇
// ==========================================
async function generateProjectStats(project) {
    const safeProjectName = sanitizePathSegment(project.projectName);
    const dataFolderPath = `工程專案管理/2026_工程專案/${safeProjectName}/結構化資料`;
    const graphClient = await getGraphClient();

    let requestUrl = `/users/${TARGET_USER_EMAIL}/drive/root:/${dataFolderPath}:/children`;
    const allItems = [];
    
    try {
        while (requestUrl) {
            const result = await graphClient.api(requestUrl).get();
            if (Array.isArray(result.value)) {
                allItems.push(...result.value.filter(f => f.name.endsWith('.json')));
            }
            requestUrl = result['@odata.nextLink'] || null;
        }
    } catch (error) {
        if (error.statusCode === 404 || error.code === 'itemNotFound') {
            return { error: '尚無日報資料或資料夾不存在', dataQuality: null, stats: null };
        }
        throw error;
    }

    const reports = [];
    const invalidFiles = [];
    
    for (const file of allItems) {
        try {
            const downloadUrl = file['@microsoft.graph.downloadUrl'];
            if (!downloadUrl) throw new Error('缺少下載網址');
            const response = await fetch(downloadUrl);
            if (!response.ok) throw new Error(`下載失敗 HTTP ${response.status}`);
            reports.push(await response.json());
        } catch (error) {
            console.error(`統計資料讀取失敗：${file.name}`, error);
            invalidFiles.push({ fileName: file.name, error: String(error.message || '未知錯誤') });
        }
    }

    const latestReportsMap = new Map();
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    let supersededCount = 0;

    for (const report of reports) {
        const date = String(report.reportDate || '').trim();
        const submittedAtTime = Date.parse(report.submittedAt || '');

        if (!dateRegex.test(date) || !Number.isFinite(submittedAtTime)) {
            invalidFiles.push({ fileName: `SubmissionId: ${report.submissionId}`, error: '時間格式錯誤' });
            continue;
        }

        const existing = latestReportsMap.get(date);
        if (!existing) {
            latestReportsMap.set(date, report);
            continue;
        }

        const existingTime = Date.parse(existing.submittedAt || '');
        if (!Number.isFinite(existingTime) || submittedAtTime > existingTime) {
            latestReportsMap.set(date, report);
            supersededCount++;
        } else {
            supersededCount++;
        }
    }

    const validReports = Array.from(latestReportsMap.values());

    const stats = {
        totalDays: validReports.length,
        workDays: 0,
        noWorkDays: 0,
        totalManDays: 0,
        contractorStats: {},
        materialStats: {},
        workItemsStats: {},
        customWorkItemsStats: {},
        noWorkReasons: {}
    };

    for (const report of validReports) {
        if (report.isNoWork) {
            stats.noWorkDays++;
            const reason = report.noWorkReason || '未填寫原因';
            stats.noWorkReasons[reason] = (stats.noWorkReasons[reason] || 0) + 1;
        } else {
            stats.workDays++;
            stats.totalManDays += Number(report.totalWorkerCount) || 0;

            if (Array.isArray(report.contractorItems)) {
                report.contractorItems.forEach(item => {
                    const name = String(item.contractorName || '未知廠商').trim();
                    const count = Number(item.workerCount) || 0;
                    
                    if (!stats.contractorStats[name]) {
                        stats.contractorStats[name] = { manDays: 0, workDays: 0 };
                    }
                    stats.contractorStats[name].manDays += count;
                    stats.contractorStats[name].workDays += 1;
                });
            }

            if (Array.isArray(report.workItems)) {
                const uniqueWorkItems = new Set(report.workItems.map(item => String(item || '').trim()).filter(Boolean));
                for (const item of uniqueWorkItems) {
                    if (item === '其他' && report.customWorkItem) {
                        const customName = String(report.customWorkItem).trim();
                        stats.customWorkItemsStats[customName] = (stats.customWorkItemsStats[customName] || 0) + 1;
                        stats.workItemsStats['其他(自訂)'] = (stats.workItemsStats['其他(自訂)'] || 0) + 1;
                    } else {
                        stats.workItemsStats[item] = (stats.workItemsStats[item] || 0) + 1;
                    }
                }
            }

            if (Array.isArray(report.materialItems)) {
                report.materialItems.forEach(item => {
                    const materialName = String(item.materialName || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
                    const baseUnit = String(item.baseUnit || '').trim();
                    const baseQuantity = Number(item.baseQuantity);

                    if (!materialName || !baseUnit || !Number.isFinite(baseQuantity) || baseQuantity <= 0) return;

                    const key = `${materialName} (${baseUnit})`;
                    stats.materialStats[key] = (stats.materialStats[key] || 0) + baseQuantity;
                });
            }
        }
    }

    const dataQuality = {
        sourceFileCount: allItems.length,
        parsedFileCount: reports.length,
        invalidFileCount: invalidFiles.length,
        effectiveReportCount: validReports.length,
        supersededReportCount: supersededCount
    };

    return { stats, dataQuality, warnings: invalidFiles };
}

function verifyStatsApiKey(req, res, next) {
    const apiKey = req.get('x-api-key');
    if (!STATS_API_KEY || apiKey !== STATS_API_KEY) {
        return res.status(401).json({ success: false, error: '未授權存取' });
    }
    next();
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
                    '我是負責協助自動化建案與日報歸檔的機器人。請依以下步驟啟用專屬日報：',
                    '',
                    '1️⃣ 首次開工請輸入「設定案場 案場名稱」',
                    '2️⃣ 將回覆的專屬網址「設為置頂公告」'
                ].join('\n');
                await replyLineMessage(event.replyToken, welcomeText);
                continue;
            }
            if (event.type === 'message' && event.message.type === 'text') {
                const text = event.message.text.trim();
                
                if (text.startsWith('設定案場')) {
                    if (!targetId) { await replyLineMessage(event.replyToken, '⚠️ 請在施工群組內使用。'); continue; }
                    const match = text.match(/^設定案場\s+(.+)$/);
                    if (!match) { await replyLineMessage(event.replyToken, '⚠️ 格式錯誤\n正確格式：設定案場 大安區'); continue; }
                    const projectName = match[1].trim();
                    
                    let registration;
                    try { registration = await registerProjectByName(projectName); } 
                    catch (error) { await replyLineMessage(event.replyToken, `⚠️ 無法建立案場\n${getProjectRegistrationErrorMessage(error)}`); continue; }

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

                    const reportUrl = `https://liff.line.me/${LIFF_ID}/?projectId=${encodeURIComponent(project.projectId)}`;
                    await replyLineMessage(event.replyToken, `✅ 案場「${project.projectName}」設定完成\n\n請將以下網址設為群組公告：\n${reportUrl}`);
                }
                else if (text === '查詢案場' || text === '案場查詢') {
                    if (!targetId) continue;
                    const config = await readBindingsFromOneDrive();
                    const binding = (Array.isArray(config.bindings) ? config.bindings : []).find(b => b.groupId === targetId && b.active);
                    await replyLineMessage(event.replyToken, binding ? `📍 本群組綁定案場：\n${binding.projectName}` : '⚠️ 尚未設定案場');
                }
                else if (text === '解除案場') {
                    if (!targetId) continue;
                    await withBindingWriteLock(async () => {
                        const config = await readBindingsFromOneDrive();
                        const bindings = Array.isArray(config.bindings) ? config.bindings : [];
                        const filteredBindings = bindings.filter(b => b.groupId !== targetId);
                        if (filteredBindings.length === bindings.length) { await replyLineMessage(event.replyToken, '無綁定紀錄。'); return; }
                        await writeBindingsToOneDrive({ ...config, bindings: filteredBindings, updatedAt: new Date().toISOString() });
                        await replyLineMessage(event.replyToken, '✅ 已解除綁定。');
                    });
                }
                else if (text === '查詢統計' || text === '案場統計') {
                    if (!targetId) {
                        await replyLineMessage(event.replyToken, '⚠️ 請在施工群組內使用「查詢統計」指令。');
                        continue;
                    }

                    const bindingConfig = await readBindingsFromOneDrive();
                    const bindings = Array.isArray(bindingConfig.bindings) ? bindingConfig.bindings : [];
                    const currentBinding = bindings.find(b => b.groupId === targetId && b.active === true);

                    if (!currentBinding) {
                        await replyLineMessage(event.replyToken, '⚠️ 本群組目前沒有綁定案場，無法查詢統計。');
                        continue;
                    }

                    const config = await readProjectsFromOneDrive();
                    const projects = Array.isArray(config.projects) ? config.projects : [];
                    const project = projects.find(p => p.projectId === currentBinding.projectId);

                    if (!project) {
                        await replyLineMessage(event.replyToken, '⚠️ 系統找不到此案場的詳細資料。');
                        continue;
                    }

                    try {
                        const result = await generateProjectStats(project);
                        if (result.error) {
                            await replyLineMessage(event.replyToken, `⚠️ 查詢失敗：${result.error}`);
                            continue;
                        }

                        const stats = result.stats;
                        
                        let msg = `【${project.projectName}】累計統計表\n`;
                        msg += `━━━━━━━━━━━━\n`;
                        msg += `實際工作天：${stats.workDays} 天\n`;
                        msg += `免計工作天：${stats.noWorkDays} 天\n`;
                        msg += `全案總人天：${stats.totalManDays} 人天\n\n`;

                        msg += `[ 各廠商出工統計 ]\n`;
                        if (Object.keys(stats.contractorStats).length === 0) msg += ` • 無紀錄\n`;
                        for (const [name, data] of Object.entries(stats.contractorStats)) {
                            msg += ` • ${name}：${data.workDays} 工作天 (${data.manDays} 人天)\n`;
                        }

                        msg += `\n[ 材料累計消耗 ]\n`;
                        if (Object.keys(stats.materialStats).length === 0) msg += ` • 無紀錄\n`;
                        for (const [name, qty] of Object.entries(stats.materialStats)) {
                            msg += ` • ${name}：共 ${qty}\n`;
                        }

                        msg += `━━━━━━━━━━━━\n`;
                        msg += `* 資料計算至最新一份日報`;

                        if (result.dataQuality) {
                            if (result.dataQuality.invalidFileCount > 0) {
                                msg += `\n⚠️ 注意：發現 ${result.dataQuality.invalidFileCount} 份資料異常，統計可能不完整`;
                            }
                            if (result.dataQuality.supersededReportCount > 0) {
                                msg += `\n* 同日舊版已排除：${result.dataQuality.supersededReportCount} 份`;
                            }
                        }

                        await replyLineMessage(event.replyToken, msg);

                    } catch (err) {
                        console.error('群組查詢統計失敗：', err);
                        await replyLineMessage(event.replyToken, '⚠️ 統計計算過程中發生錯誤，請稍後再試。');
                    }
                }
                else if (text.startsWith('結案')) {
                    if (!targetId) { await replyLineMessage(event.replyToken, '⚠️ 請在施工群組內使用「結案」指令。'); continue; }
                    const match = text.match(/^結案\s+(.+)$/);
                    if (!match) { await replyLineMessage(event.replyToken, '⚠️ 格式錯誤\n正確格式：結案 大安區'); continue; }
                    const targetProjectName = match[1].trim();

                    const bindingConfig = await readBindingsFromOneDrive();
                    const bindings = Array.isArray(bindingConfig.bindings) ? bindingConfig.bindings : [];
                    const currentBinding = bindings.find(b => b.groupId === targetId && b.active === true);

                    if (!currentBinding) {
                        await replyLineMessage(event.replyToken, '⚠️ 本群組目前沒有綁定案場，無法執行結案。');
                        continue;
                    }
                    if (normalizeProjectName(targetProjectName) !== normalizeProjectName(currentBinding.projectName)) {
                        await replyLineMessage(event.replyToken, `⚠️ 結案名稱不符\n本群組案場：${currentBinding.projectName}\n輸入名稱：${targetProjectName}`);
                        continue;
                    }

                    await withProjectWriteLock(async () => {
                        const config = await readProjectsFromOneDrive();
                        const projects = Array.isArray(config.projects) ? config.projects : [];
                        const projectIndex = projects.findIndex(p => p.projectId === currentBinding.projectId);
                        
                        if (projectIndex === -1) {
                            await replyLineMessage(event.replyToken, '⚠️ 系統找不到此案場資料。'); return;
                        }
                        
                        const closingProject = projects[projectIndex];

                        let finalStatsResult;
                        try {
                            finalStatsResult = await generateProjectStats(closingProject);
                            
                            if (finalStatsResult.error || !finalStatsResult.stats) {
                                await replyLineMessage(event.replyToken, ['⚠️ 結案暫停', '', finalStatsResult.error || '目前無法產生結案統計。', '', '案場尚未下架，群組綁定也未解除。'].join('\n'));
                                return;
                            }

                            if (Array.isArray(finalStatsResult.warnings) && finalStatsResult.warnings.length > 0) {
                                await replyLineMessage(event.replyToken, ['⚠️ 結案暫停', '', `發現 ${finalStatsResult.warnings.length} 份異常結構化資料。`, '為避免統計漏算，本次尚未完成結案。', '', '請先檢查 OneDrive 資料或 Render Logs。'].join('\n'));
                                return;
                            }

                            const { dateStr, timeStr } = getTaiwanDateParts();
                            const statsFileName = `結案統計_${dateStr.replace(/-/g, '')}_${timeStr}.json`;
                            const safeProjectName = sanitizePathSegment(closingProject.projectName);
                            const graphClient = await getGraphClient();
                            const statsFilePath = `工程專案管理/2026_工程專案/${safeProjectName}/${statsFileName}`;
                            
                            const statsBuffer = Buffer.from(JSON.stringify({
                                schemaVersion: 1, projectId: closingProject.projectId, projectName: closingProject.projectName,
                                generatedAt: new Date().toISOString(), ...finalStatsResult
                            }, null, 2), 'utf-8');
                            
                            await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${statsFilePath}:/content`).put(statsBuffer);

                        } catch (statErr) {
                            console.error('結案統計產生或儲存失敗：', statErr);
                            await replyLineMessage(event.replyToken, ['⚠️ 結案失敗', '', '系統無法完成最終統計或寫入統計檔。', '案場尚未下架，群組綁定也未解除。', '', '請稍後再試。'].join('\n'));
                            return;
                        }

                        projects.splice(projectIndex, 1);
                        await writeProjectsToOneDrive({ ...config, projects, updatedAt: new Date().toISOString() });

                        await withBindingWriteLock(async () => {
                            const latestBindingConfig = await readBindingsFromOneDrive();
                            const latestBindings = Array.isArray(latestBindingConfig.bindings) ? latestBindingConfig.bindings : [];
                            const filteredBindings = latestBindings.filter(binding => binding.projectId !== closingProject.projectId);
                            await writeBindingsToOneDrive({ ...latestBindingConfig, bindings: filteredBindings, updatedAt: new Date().toISOString() });
                        });

                        await replyLineMessage(event.replyToken, `✅ 案場「${targetProjectName}」已成功結案！\n\n系統已自動產生無異常之最終統計報表，並存入您的 OneDrive 資料夾中。`);
                    });
                }
                else if (['指令', '說明', '功能', '小幫手', '【點此查看指令說明】'].includes(text)) {
                    const helpText = [
                        '📖 「云說工程小幫手」群組指令說明',
                        '',
                        '🔹 設定案場 案場名稱',
                        '🔹 查詢案場',
                        '🔹 查詢統計',
                        '🔹 解除案場',
                        '🔹 結案 案場名稱 (自動結算並下架)'
                    ].join('\n');
                    await replyLineMessage(event.replyToken, helpText);
                }
            }
        } catch (error) { console.error('LINE 事件處理失敗：', error); }
    }
});

app.use(express.json());

app.get('/api/project-stats/:projectId', verifyStatsApiKey, async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId);
        if (!project) return res.status(404).json({ success: false, error: '找不到該案場' });
        
        const result = await generateProjectStats(project);
        if (result.error) return res.status(200).json({ success: true, message: result.error });

        return res.status(200).json({ success: true, projectName: project.projectName, ...result });
    } catch (error) {
        console.error('統計產生失敗：', error);
        return res.status(500).json({ success: false, error: '統計產生失敗' });
    }
});

app.get('/api/projects/:projectId', async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId);
        if (!project) return res.status(404).json({ success: false, reason: 'PROJECT_NOT_FOUND', error: '找不到指定案場' });
        return res.status(200).json({ success: true, project: { projectId: project.projectId, projectName: project.projectName } });
    } catch (error) {
        console.error('讀取案場失敗：', error);
        return res.status(500).json({ success: false, reason: 'INTERNAL_ERROR', error: '無法取得案場' });
    }
});

app.get('/', (req, res) => res.send('✅ 伺服器運作中！'));

app.post('/api/submit-report', async (req, res) => {
    try {
        const reportData = req.body || {}; 
        const validation = validateReportData(reportData);
        if (!validation.valid) return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'INVALID_REPORT_DATA', error: `缺少必要欄位：${validation.missingFields.join(', ')}` });

        const submittedProjectId = String(reportData.projectId || '').trim();
        let project = submittedProjectId ? await findProjectById(submittedProjectId) : await findProjectByName(reportData.projectName);
        
        if (!project) {
            return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'PROJECT_NOT_FOUND', error: '找不到指定案場' });
        }

        const isNoWork = reportData.isNoWork === true;
        let contractorItems = Array.isArray(reportData.contractorItems) ? reportData.contractorItems : [];

        if (!isNoWork) {
            if (contractorItems.length === 0) return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'INVALID_CONTRACTOR_ITEMS', error: '請至少填寫一組有效廠商' });
            for (const item of contractorItems) {
                const contractorName = String(item.contractorName || '').trim();
                const workerCount = Number(item.workerCount);
                if (!contractorName) return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'INVALID_CONTRACTOR_NAME', error: '施工廠商名稱不可為空' });
                if (!Number.isInteger(workerCount) || workerCount <= 0 || workerCount > 200) return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'INVALID_WORKER_COUNT', error: `廠商「${contractorName}」人數格式不正確` });
            }
        } else {
            contractorItems = [];
        }

        const calculatedTotalWorkerCount = isNoWork ? 0 : contractorItems.reduce((total, item) => total + Number(item.workerCount), 0);

        await ensureProjectFolder(project.projectName);
        const safeProjectName = sanitizePathSegment(project.projectName);
        const graphClient = await getGraphClient();
        
        const projectFolderPath = `工程專案管理/2026_工程專案/${safeProjectName}`;
        const textFolderResult = await ensureChildFolder(graphClient, projectFolderPath, '施工日報');
        const dataFolderResult = await ensureChildFolder(graphClient, projectFolderPath, '結構化資料');

        const { dateStr, timeStr } = getTaiwanDateParts();
        
        const submittedSubmissionId = String(reportData.submissionId || '').trim();
        const fullSubmissionId = submittedSubmissionId || crypto.randomUUID();
        const safeSubmissionId = fullSubmissionId.replace(/[^a-zA-Z0-9]/g, '');
        const shortSubmissionId = safeSubmissionId.slice(0, 16);

        if (!shortSubmissionId) throw new Error('無法產生日報識別碼');

        const workItems = Array.isArray(reportData.workItems) ? reportData.workItems : [];
        const materialItems = Array.isArray(reportData.materialItems) ? reportData.materialItems : [];

        const structuredReport = {
            schemaVersion: 1, projectId: project.projectId, projectName: project.projectName, reportDate: dateStr,
            submissionId: fullSubmissionId, submittedAt: new Date().toISOString(), submittedDateLocal: dateStr, submittedTimeLocal: timeStr,
            isNoWork, noWorkReason: isNoWork ? String(reportData.noWorkReason || '') : '',
            weather: { temp: reportData.temp, humidity: reportData.humidity, wind: reportData.wind },
            contractorItems, totalWorkerCount: calculatedTotalWorkerCount, workItems: isNoWork ? [] : workItems,
            customWorkItem: isNoWork ? '' : String(reportData.customWorkItem || ''), workNotes: isNoWork ? '' : String(reportData.workNotes || ''),
            materialItems: isNoWork ? [] : materialItems, remarks: String(reportData.remarks || '')
        };

        const baseFileName = `${dateStr}_${shortSubmissionId}`;
        const jsonFilePath = `${dataFolderResult.folderPath}/${baseFileName}.json`;
        const txtFilePath = `${textFolderResult.folderPath}/${baseFileName}_施工日報.txt`;

        try {
            await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${jsonFilePath}:/content`).put(Buffer.from(JSON.stringify(structuredReport, null, 2), 'utf-8'));
        } catch (jsonUploadError) {
            console.error('[Error] 結構化 JSON 寫入失敗：', jsonUploadError);
            throw new Error('結構化日報寫入失敗');
        }

        let reportText = `📋 施工日報\n\n日期：${dateStr.replace(/-/g, '/')}\n案場：${project.projectName}\n\n溫度：${reportData.temp}度\n濕度：${reportData.humidity}%\n風速：${reportData.wind}m/s\n\n施工廠商：${reportData.contractor}\n施工人數：${reportData.workerCount}\n\n━━━━━━━━━━━━\n\n今日作業進度：\n${reportData.progress}\n\n今日用料：\n${reportData.materials}\n\n備註：\n${reportData.remarks || '無'}\n\n━━━━━━━━━━━━\n以上為今日進度報告`;
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${txtFilePath}:/content`).put(reportText);

        const config = await readBindingsFromOneDrive();
        const binding = (Array.isArray(config.bindings) ? config.bindings : []).find(b => b.projectId === project.projectId && b.active);

        if (!binding) return res.status(200).json({ success: true, archived: true, pushed: false, reason: 'PROJECT_NOT_BOUND' });

        try {
            await pushLineMessage(binding.groupId, reportText);
            return res.status(200).json({ success: true, archived: true, pushed: true, message: '日報已歸檔並發布' });
        } catch (lineError) {
            return res.status(200).json({ success: true, archived: true, pushed: false, reason: 'LINE_PUSH_FAILED' });
        }
    } catch (error) {
        if (!res.headersSent) return res.status(500).json({ success: false, archived: false, pushed: false, reason: 'INTERNAL_ERROR', error: '系統處理失敗' });
    }
});

const requiredVars = ['LINE_ACCESS_TOKEN', 'LINE_CHANNEL_SECRET', 'AZURE_CLIENT_ID', 'AZURE_TENANT_ID', 'AZURE_CLIENT_SECRET', 'STATS_API_KEY'];
if (requiredVars.some(v => !process.env[v])) process.exit(1);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 伺服器運作中：http://localhost:${PORT}`));
