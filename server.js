require('dotenv').config();
const express = require('express');
const cors = require('cors');
const crypto = require('crypto');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const ExcelJS = require('exceljs');

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

let cachedGraphClient = null;
let tokenExpiresAt = 0;
let graphClientPromise = null;

async function getGraphClient() {
    const now = Date.now();
    if (cachedGraphClient && now < tokenExpiresAt) return cachedGraphClient;
    if (graphClientPromise) return graphClientPromise;

    graphClientPromise = (async () => {
        try {
            const response = await cca.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
            if (!response || !response.accessToken) throw new Error('Microsoft Graph Token 取得失敗');
            const expiresAt = response.expiresOn ? response.expiresOn.getTime() : Date.now() + 50 * 60 * 1000;
            tokenExpiresAt = Math.max(Date.now() + 60 * 1000, expiresAt - 5 * 60 * 1000);
            cachedGraphClient = Client.init({ authProvider(done) { done(null, response.accessToken); } });
            return cachedGraphClient;
        } catch (error) {
            cachedGraphClient = null;
            tokenExpiresAt = 0;
            throw error;
        } finally {
            graphClientPromise = null;
        }
    })();
    return graphClientPromise;
}

function sanitizePathSegment(value) { return String(value).replace(/[<>:"/\\|?*#%]/g, '_').replace(/\s+/g, ' ').trim(); }
function normalizeProjectName(value) { return String(value || '').normalize('NFKC').replace(/\s+/g, ' ').trim(); }
function validateProjectName(projectName) {
    const normalizedName = normalizeProjectName(projectName);
    if (!normalizedName) throw new Error('案場名稱不可為空');
    if (normalizedName.length > 80) throw new Error('案場名稱不可超過 80 個字');
    if (/[<>:"/\\|?*#%]/.test(normalizedName)) throw new Error('案場名稱不可包含以下字元：< > : " / \\ | ? * # %');
    return normalizedName;
}
function getProjectRegistrationErrorMessage(error) {
    const safeMessages = ['案場名稱不可為空', '案場名稱不可超過 80 個字', '案場名稱不可包含以下字元'];
    const message = String(error?.message || '');
    return safeMessages.some(prefix => message.startsWith(prefix)) ? message : '系統暫時無法建立案場，請稍後再試';
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
    await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${filePath}:/content`).put(JSON.stringify(data, null, 2));
}

function cloneJsonData(data) { return JSON.parse(JSON.stringify(data)); }

const CACHE_TTL = 30 * 1000;
const configCache = { 
    projects: { data: null, timestamp: 0 }, 
    bindings: { data: null, timestamp: 0 },
    globalMaterials: { data: null, timestamp: 0 },
    projMaterials: {}
};

async function readProjectsFromOneDrive() {
    const cache = configCache.projects;
    if (cache.data && Date.now() - cache.timestamp < CACHE_TTL) return cloneJsonData(cache.data);
    const data = await readJsonFromOneDrive('工程專案管理/_系統設定/projects.json', { projects: [] }, true);
    configCache.projects = { data: cloneJsonData(data), timestamp: Date.now() };
    return cloneJsonData(data);
}

async function writeProjectsToOneDrive(config) {
    await writeJsonToOneDrive('工程專案管理/_系統設定/projects.json', config);
    configCache.projects = { data: cloneJsonData(config), timestamp: Date.now() };
}

async function readBindingsFromOneDrive() {
    const cache = configCache.bindings;
    if (cache.data && Date.now() - cache.timestamp < CACHE_TTL) return cloneJsonData(cache.data);
    const data = await readJsonFromOneDrive('工程專案管理/_系統設定/line-bindings.json', { bindings: [] });
    configCache.bindings = { data: cloneJsonData(data), timestamp: Date.now() };
    return cloneJsonData(data);
}

async function writeBindingsToOneDrive(config) {
    await writeJsonToOneDrive('工程專案管理/_系統設定/line-bindings.json', config);
    configCache.bindings = { data: cloneJsonData(config), timestamp: Date.now() };
}

async function readGlobalMaterials() {
    const cache = configCache.globalMaterials;
    if (cache.data && Date.now() - cache.timestamp < CACHE_TTL) return cloneJsonData(cache.data);
    const data = await readJsonFromOneDrive('工程專案管理/_系統設定/materials.json', { materials: [] }, false);
    configCache.globalMaterials = { data: cloneJsonData(data), timestamp: Date.now() };
    return cloneJsonData(data);
}

async function readProjectMaterials(projectName) {
    const safeName = sanitizePathSegment(projectName);
    const cache = configCache.projMaterials[safeName];
    if (cache && Date.now() - cache.timestamp < CACHE_TTL) return cloneJsonData(cache.data);
    try {
        const data = await readJsonFromOneDrive(`工程專案管理/2026_工程專案/${safeName}/專屬材料.json`, null, false);
        if (data && Array.isArray(data.materials)) {
            configCache.projMaterials[safeName] = { data: cloneJsonData(data), timestamp: Date.now() };
            return cloneJsonData(data);
        }
    } catch (e) {}
    return null;
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
        if (error?.statusCode !== 404 && error?.code !== 'itemNotFound') throw error;
    }
    try {
        const createdFolder = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/工程專案管理/2026_工程專案:/children`).post({ name: safeProjectName, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
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
        if (error?.statusCode !== 404 && error?.code !== 'itemNotFound') throw error;
    }
    try {
        const createdFolder = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${parentPath}:/children`).post({ name: safeChildName, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
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

function normalizeMaterialItems(rawItems, isNoWork) {
    if (isNoWork) return [];
    if (!Array.isArray(rawItems)) return [];
    
    const allowedStockUnits = new Set(['桶', '組', '支', '公斤', '公升', '個', '捲']);
    
    return rawItems.map((item, index) => {
        const materialName = String(item.materialName || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
        const quantity = Number(item.quantity);
        const stockUnit = String(item.stockUnit || '').trim();
        const packageQuantity = item.packageQuantity == null ? null : Number(item.packageQuantity);
        const packageUnit = item.packageUnit == null ? null : String(item.packageUnit).trim();
        
        if (!materialName) throw new Error(`第 ${index + 1} 筆材料名稱不可為空`);
        if (!Number.isFinite(quantity) || quantity <= 0) throw new Error(`第 ${index + 1} 筆材料數量不正確`);
        if (!allowedStockUnits.has(stockUnit)) throw new Error(`第 ${index + 1} 筆材料單位不正確`);
        
        let baseQuantity = quantity;
        let baseUnit = stockUnit;
        
        if (stockUnit === '桶' && packageUnit === '加侖') {
            if (packageQuantity !== 1 && packageQuantity !== 5) {
                throw new Error(`第 ${index + 1} 筆桶裝容量不正確`);
            }
            baseQuantity = quantity * packageQuantity;
            baseUnit = '加侖';
        }
        
        return {
            materialId: null, materialName, quantity, stockUnit,
            packageQuantity, packageUnit, baseQuantity, baseUnit
        };
    });
}

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
            return { error: '尚無日報資料或資料夾不存在', dataQuality: null, stats: null, reports: [] };
        }
        throw error;
    }

    const reports = [];
    const invalidFiles = [];
    const DOWNLOAD_CONCURRENCY = 10;
    
    for (let i = 0; i < allItems.length; i += DOWNLOAD_CONCURRENCY) {
        const chunk = allItems.slice(i, i + DOWNLOAD_CONCURRENCY);
        const chunkResults = await Promise.all(chunk.map(async file => {
            try {
                const downloadUrl = file['@microsoft.graph.downloadUrl'];
                if (!downloadUrl) throw new Error('缺少下載網址');
                const response = await fetch(downloadUrl);
                if (!response.ok) throw new Error(`下載失敗 HTTP ${response.status}`);
                return { success: true, report: await response.json() };
            } catch (error) {
                return { success: false, fileName: file.name, error: String(error.message || '未知錯誤') };
            }
        }));

        for (const result of chunkResults) {
            if (result.success) {
                reports.push(result.report);
            } else {
                console.error(`統計資料讀取失敗：${result.fileName}`, result.error);
                invalidFiles.push({ fileName: result.fileName, error: result.error });
            }
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
    validReports.sort((a, b) => a.reportDate.localeCompare(b.reportDate));

    const stats = {
        totalDays: validReports.length,
        workDays: 0,
        noWorkDays: 0,
        totalManDays: 0,
        contractorStats: {},
        materialStats: {},
        workItemsStats: {},
        customWorkItemsStats: {},
        noWorkReasons: {},
        reporterStats: {} // 💡 新增：紀錄填表人的出工天數
    };

    for (const report of validReports) {
        if (report.isNoWork) {
            stats.noWorkDays++;
            const reason = report.noWorkReason || '未填寫原因';
            stats.noWorkReasons[reason] = (stats.noWorkReasons[reason] || 0) + 1;
        } else {
            stats.workDays++;
            stats.totalManDays += Number(report.totalWorkerCount) || 0;

            // 💡 統計填表人 (帶班主管) 的出工天數
            const reporter = String(report.reporterName || '未紀錄').trim();
            stats.reporterStats[reporter] = (stats.reporterStats[reporter] || 0) + 1;

            if (Array.isArray(report.contractorItems)) {
                const dailyContractors = new Map();
                for (const item of report.contractorItems) {
                    const name = String(item.contractorName || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
                    const count = Number(item.workerCount);
                    if (!name || !Number.isFinite(count) || count <= 0) continue;
                    dailyContractors.set(name, (dailyContractors.get(name) || 0) + count);
                }
                for (const [name, count] of dailyContractors.entries()) {
                    if (!stats.contractorStats[name]) stats.contractorStats[name] = { manDays: 0, workDays: 0 };
                    stats.contractorStats[name].manDays += count;
                    stats.contractorStats[name].workDays += 1;
                }
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

    return { stats, dataQuality, warnings: invalidFiles, reports: validReports };
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

app.get('/api/materials', async (req, res) => {
    const projectId = req.query.projectId;
    try {
        if (projectId) {
            const project = await findProjectById(projectId);
            if (project) {
                const customConfig = await readProjectMaterials(project.projectName);
                if (customConfig && Array.isArray(customConfig.materials)) {
                    return res.status(200).json({ success: true, materials: customConfig.materials, type: 'project' });
                }
            }
        }
        const globalConfig = await readGlobalMaterials();
        const materials = Array.isArray(globalConfig.materials) ? globalConfig.materials : [];
        return res.status(200).json({ success: true, materials, type: 'global' });
    } catch (error) {
        console.error('讀取材料清單失敗：', error);
        return res.status(500).json({ success: false, error: '無法取得材料清單' });
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
                            if (result.dataQuality.invalidFileCount > 0) msg += `\n⚠️ 注意：發現 ${result.dataQuality.invalidFileCount} 份資料異常，統計可能不完整`;
                            if (result.dataQuality.supersededReportCount > 0) msg += `\n* 同日舊版已排除：${result.dataQuality.supersededReportCount} 份`;
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
                            const safeProjectName = sanitizePathSegment(closingProject.projectName);
                            const graphClient = await getGraphClient();
                            
                            const statsFileName = `結案統計_${dateStr.replace(/-/g, '')}_${timeStr}.json`;
                            const statsFilePath = `工程專案管理/2026_工程專案/${safeProjectName}/${statsFileName}`;
                            const statsBuffer = Buffer.from(JSON.stringify({
                                schemaVersion: 1, projectId: closingProject.projectId, projectName: closingProject.projectName,
                                generatedAt: new Date().toISOString(), ...finalStatsResult
                            }, null, 2), 'utf-8');
                            await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${statsFilePath}:/content`).put(statsBuffer);

                            const workbook = new ExcelJS.Workbook();
                            workbook.creator = '云說工程小幫手';
                            workbook.created = new Date();

                            // 💡 Excel 總表新增：開案日期 與 填表人出工統計
                            const summarySheet = workbook.addWorksheet('案場總表');
                            summarySheet.columns = [
                                { header: '統計項目', key: 'item', width: 25 },
                                { header: '數據', key: 'value', width: 40 }
                            ];
                            const startDateStr = finalStatsResult.reports.length > 0 ? finalStatsResult.reports[0].reportDate : '無紀錄';
                            
                            summarySheet.addRow({ item: '案場名稱', value: closingProject.projectName });
                            summarySheet.addRow({ item: '開案日期', value: startDateStr });
                            summarySheet.addRow({ item: '結案日期', value: dateStr });
                            summarySheet.addRow({ item: '累計日曆天', value: `${finalStatsResult.stats.totalDays} 天` });
                            summarySheet.addRow({ item: '實際工作天', value: `${finalStatsResult.stats.workDays} 天` });
                            summarySheet.addRow({ item: '免計工作天', value: `${finalStatsResult.stats.noWorkDays} 天` });
                            summarySheet.addRow({ item: '全案總人天', value: `${finalStatsResult.stats.totalManDays} 人天` });
                            
                            summarySheet.addRow({ item: '', value: '' });
                            summarySheet.addRow({ item: '【各人員填表(帶班)天數】', value: '' });
                            for (const [name, days] of Object.entries(finalStatsResult.stats.reporterStats)) {
                                summarySheet.addRow({ item: ` • ${name}`, value: `${days} 工作天` });
                            }

                            summarySheet.addRow({ item: '', value: '' });
                            summarySheet.addRow({ item: '【各廠商出工統計】', value: '' });
                            for (const [name, data] of Object.entries(finalStatsResult.stats.contractorStats)) {
                                summarySheet.addRow({ item: ` • ${name}`, value: `${data.workDays} 工作天 (${data.manDays} 人天)` });
                            }

                            const inventorySheet = workbook.addWorksheet('材料累計消耗');
                            inventorySheet.columns = [
                                { header: '材料名稱', key: 'name', width: 25 },
                                { header: '單位', key: 'unit', width: 15 },
                                { header: '現場累計消耗', key: 'used', width: 20 },
                                { header: '備註', key: 'remarks', width: 40 }
                            ];
                            for (const [key, qty] of Object.entries(finalStatsResult.stats.materialStats)) {
                                const match = key.match(/(.+?)\s+\((.+)\)/);
                                const name = match ? match[1] : key;
                                const unit = match ? match[2] : '';
                                inventorySheet.addRow({ name: name, unit: unit, used: qty, remarks: '' });
                            }

                            const materialLogSheet = workbook.addWorksheet('材料進出紀錄');
                            materialLogSheet.columns = [
                                { header: '日期', key: 'date', width: 15 },
                                { header: '材料名稱', key: 'name', width: 25 },
                                { header: '單位', key: 'unit', width: 15 },
                                { header: '使用數量', key: 'used_qty', width: 15 },
                                { header: '日報備註', key: 'remarks', width: 50 }
                            ];
                            for (const r of finalStatsResult.reports) {
                                if (!r.isNoWork && Array.isArray(r.materialItems)) {
                                    for (const m of r.materialItems) {
                                        materialLogSheet.addRow({
                                            date: r.reportDate,
                                            name: m.materialName,
                                            unit: m.baseUnit,
                                            used_qty: m.baseQuantity,
                                            remarks: r.remarks || r.workNotes || ''
                                        });
                                    }
                                }
                            }

                            // 💡 Excel 日報明細新增：填表人 欄位
                            const dailyLogSheet = workbook.addWorksheet('日報明細');
                            dailyLogSheet.columns = [
                                { header: '日期', key: 'date', width: 15 },
                                { header: '填表人', key: 'reporter', width: 15 },
                                { header: '出工狀態', key: 'status', width: 15 },
                                { header: '出工人數', key: 'workers', width: 15 },
                                { header: '施作項目', key: 'work_items', width: 40 },
                                { header: '日報備註', key: 'remarks', width: 50 }
                            ];
                            for (const r of finalStatsResult.reports) {
                                let itemsStr = Array.isArray(r.workItems) ? r.workItems.join('、') : '';
                                if (r.customWorkItem) itemsStr += ` (${r.customWorkItem})`;
                                dailyLogSheet.addRow({
                                    date: r.reportDate,
                                    reporter: r.reporterName || '未紀錄',
                                    status: r.isNoWork ? `停工 (${r.noWorkReason})` : '施工',
                                    workers: r.totalWorkerCount || 0,
                                    work_items: itemsStr,
                                    remarks: r.remarks || ''
                                });
                            }

                            workbook.eachSheet((sheet) => {
                                const headerRow = sheet.getRow(1);
                                headerRow.font = { bold: true, color: { arg: 'FFFFFFFF' } };
                                headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { arg: 'FF4F81BD' } };
                                headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
                            });

                            const excelBufferArray = await workbook.xlsx.writeBuffer();
                            const excelBuffer = Buffer.from(excelBufferArray);
                            const excelFileName = `結案總表_${safeProjectName}_${dateStr.replace(/-/g, '')}.xlsx`;
                            const excelFilePath = `工程專案管理/2026_工程專案/${safeProjectName}/${excelFileName}`;
                            await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${excelFilePath}:/content`).put(excelBuffer);

                        } catch (statErr) {
                            console.error('結案報表產生或儲存失敗：', statErr);
                            await replyLineMessage(event.replyToken, ['⚠️ 結案失敗', '', '系統無法完成最終統計或寫入報表檔。', '案場尚未下架，群組綁定也未解除。', '', '請稍後再試。'].join('\n'));
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

                        await replyLineMessage(event.replyToken, `✅ 案場「${targetProjectName}」已成功結案！\n\n系統已自動產生【Excel 結案報表】與統計資料，並存入您的 OneDrive 資料夾中。`);
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
            
            const contractorNameSet = new Set();
            const validContractorItems = [];
            
            for (const item of contractorItems) {
                const contractorName = String(item.contractorName || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
                const workerCount = Number(item.workerCount);
                
                if (!contractorName) return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'INVALID_CONTRACTOR_NAME', error: '施工廠商名稱不可為空' });
                if (!Number.isInteger(workerCount) || workerCount <= 0 || workerCount > 200) {
                    return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'INVALID_WORKER_COUNT', error: `廠商「${contractorName}」人數格式不正確` });
                }
                if (contractorNameSet.has(contractorName)) {
                    return res.status(400).json({ success: false, archived: false, pushed: false, reason: 'DUPLICATE_CONTRACTOR', error: `施工廠商「${contractorName}」重複填寫` });
                }
                
                contractorNameSet.add(contractorName);
                validContractorItems.push({ contractorName, workerCount });
            }
            contractorItems = validContractorItems;
        } else {
            contractorItems = [];
        }

        const calculatedTotalWorkerCount = isNoWork ? 0 : contractorItems.reduce((total, item) => total + Number(item.workerCount), 0);

        const safeProjectName = sanitizePathSegment(project.projectName);
        const graphClient = await getGraphClient();
        const projectFolderPath = `工程專案管理/2026_工程專案/${safeProjectName}`;

        await ensureProjectFolder(project.projectName);
        const [textFolderResult, dataFolderResult] = await Promise.all([
            ensureChildFolder(graphClient, projectFolderPath, '施工日報'),
            ensureChildFolder(graphClient, projectFolderPath, '結構化資料')
        ]);

        const { dateStr, timeStr } = getTaiwanDateParts();
        
        const submittedSubmissionId = String(reportData.submissionId || '').trim();
        const fullSubmissionId = submittedSubmissionId || crypto.randomUUID();
        const safeSubmissionId = fullSubmissionId.replace(/[^a-zA-Z0-9]/g, '');
        const shortSubmissionId = safeSubmissionId.slice(0, 16);

        if (!shortSubmissionId) throw new Error('無法產生日報識別碼');

        const workItems = Array.isArray(reportData.workItems) ? reportData.workItems : [];
        
        let materialItems;
        try {
            materialItems = normalizeMaterialItems(reportData.materialItems, isNoWork);
        } catch (materialError) {
            return res.status(400).json({
                success: false, archived: false, pushed: false,
                reason: 'INVALID_MATERIAL_ITEMS', error: materialError.message
            });
        }

        const structuredReport = {
            schemaVersion: 1, projectId: project.projectId, projectName: project.projectName, reportDate: dateStr,
            submissionId: fullSubmissionId, submittedAt: new Date().toISOString(), submittedDateLocal: dateStr, submittedTimeLocal: timeStr,
            reporterName: String(reportData.reporterName || '未紀錄').trim(), // 💡 新增：紀錄填表人
            isNoWork, noWorkReason: isNoWork ? String(reportData.noWorkReason || '') : '',
            weather: { temp: reportData.temp, humidity: reportData.humidity, wind: reportData.wind },
            contractorItems, totalWorkerCount: calculatedTotalWorkerCount, workItems: isNoWork ? [] : workItems,
            customWorkItem: isNoWork ? '' : String(reportData.customWorkItem || ''), workNotes: isNoWork ? '' : String(reportData.workNotes || ''),
            materialItems, remarks: String(reportData.remarks || '')
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

        // 💡 修改 LINE 回報文字，加入填表人
        const reporterNameStr = reportData.reporterName ? String(reportData.reporterName).trim() : '未紀錄';
        let reportText = `📋 施工日報\n\n日期：${dateStr.replace(/-/g, '/')}\n案場：${project.projectName}\n填表：${reporterNameStr}\n\n溫度：${reportData.temp}度\n濕度：${reportData.humidity}%\n風速：${reportData.wind}m/s\n\n施工廠商：${reportData.contractor}\n施工人數：${reportData.workerCount}\n\n━━━━━━━━━━━━\n\n今日作業進度：\n${reportData.progress}\n\n今日用料：\n${reportData.materials}\n\n備註：\n${reportData.remarks || '無'}\n\n━━━━━━━━━━━━\n以上為今日進度報告`;
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
