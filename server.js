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

const PORT = process.env.PORT || 3000;
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
    const message = String(error?.message || '');
    const safePrefixes = ['案場名稱不可為空', '案場名稱不可超過 80 個字', '案場名稱不可包含以下字元'];
    return safePrefixes.some(prefix => message.startsWith(prefix)) ? message : '系統暫時無法建立案場，請稍後再試';
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

const materialWriteQueues = new Map();
function withMaterialWriteLock(projectId, task) {
    const previous = materialWriteQueues.get(projectId) || Promise.resolve();
    const current = previous.then(task, task);
    materialWriteQueues.set(projectId, current.catch(() => undefined));
    return current;
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
    globalInventory: { data: null, timestamp: 0 },
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

async function readGlobalInventory() {
    const cache = configCache.globalInventory;
    if (cache.data && Date.now() - cache.timestamp < CACHE_TTL) {
        return cloneJsonData(cache.data);
    }
    const data = await readJsonFromOneDrive('工程專案管理/_系統設定/inventory.json', { items: [] }, true);
    if (!data || !Array.isArray(data.items)) {
        throw new Error('inventory.json 格式不正確');
    }
    configCache.globalInventory = { data: cloneJsonData(data), timestamp: Date.now() };
    return cloneJsonData(data);
}

async function readProjectMaterials(projectName) {
    const safeName = sanitizePathSegment(projectName);
    const cache = configCache.projMaterials[safeName];
    if (cache && Date.now() - cache.timestamp < CACHE_TTL) return cloneJsonData(cache.data);
    try {
        const data = await readJsonFromOneDrive(`工程專案管理/2026_工程專案/${safeName}/專屬材料.json`, null, false);
        if (data && (Array.isArray(data.items) || Array.isArray(data.materials))) {
            configCache.projMaterials[safeName] = { data: cloneJsonData(data), timestamp: Date.now() };
            return cloneJsonData(data);
        }
    } catch (e) {}
    return null;
}

// ⭐ [關鍵補回] 建立全域與專屬材料映射表
async function buildInventoryMap(project) {
    const globalInventory = await readGlobalInventory();
    const customInventory = project ? await readProjectMaterials(project.projectName) : null;
    const inventoryMap = {};

    for (const item of globalInventory?.items || []) {
        inventoryMap[item.materialId] = item;
    }
    for (const item of customInventory?.items || customInventory?.materials || []) {
        inventoryMap[item.materialId] = item;
    }
    return inventoryMap;
}

// ⭐ [關鍵補回] 原汁原味的專案資料夾建立功能
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

// ⭐ [關鍵補回] 原汁原味的子資料夾建立功能
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

function createProjectId() {
    return `PRJ-${crypto.randomUUID()}`;
}

async function registerProjectByName(projectName) {
    return withProjectWriteLock(async () => {
        const normalizedName = validateProjectName(projectName);
        const config = await readProjectsFromOneDrive();
        const projects = Array.isArray(config.projects) ? config.projects : [];
        const existingProject = projects.find(project => project.active === true && normalizeProjectName(project.projectName) === normalizedName);

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
    await fetch('https://api.line.me/v2/bot/message/reply', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${LINE_ACCESS_TOKEN}` },
        body: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] })
    });
}

async function pushLineMessage(targetId, text) {
    const response = await fetch('https://api.line.me/v2/bot/message/push', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${LINE_ACCESS_TOKEN}` },
        body: JSON.stringify({ to: targetId, messages: [{ type: 'text', text }] })
    });
    if (!response.ok) {
        throw new Error(`LINE Push失敗：${response.status}`);
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

function getTaiwanDateParts() {
    const parts = new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
    }).formatToParts(new Date());
    const values = Object.fromEntries(parts.filter(p => p.type !== 'literal').map(p => [p.type, p.value]));
    return { dateStr: `${values.year}-${values.month}-${values.day}`, timeStr: `${values.hour}${values.minute}${values.second}` };
}

function validateIssueDate(value, today) {
    const dateValue = String(value || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateValue)) throw new Error('材料進場日期格式不正確');
    const parsedTime = Date.parse(`${dateValue}T00:00:00+08:00`);
    if (!Number.isFinite(parsedTime)) throw new Error('材料進場日期無效');
    if (dateValue > today) throw new Error('材料進場日期不可晚於今天');
    return dateValue;
}

function normalizeMaterialItems(rawItems, isNoWork, inventoryMap) {
    if (isNoWork) return [];
    if (!Array.isArray(rawItems)) return [];
    
    return rawItems.map((item, index) => {
        if (!item.materialId) {
            throw new Error(`第 ${index + 1} 筆材料缺少 materialId`);
        }
        
        const dbItem = inventoryMap[item.materialId];
        if (!dbItem) {
            throw new Error(`第 ${index + 1} 筆找不到指定材料主檔 (ID: ${item.materialId})`);
        }

        const quantity = Number(item.quantity);
        if (!Number.isFinite(quantity) || quantity <= 0) {
            throw new Error(`「${dbItem.materialName}」數量不正確`);
        }
        
        const packageQty = dbItem.packageQuantity ? Number(dbItem.packageQuantity) : 1;
        
        return {
            materialId: dbItem.materialId,
            materialCode: dbItem.materialCode || '無編碼',
            materialName: dbItem.materialName,
            quantity: quantity,
            stockUnit: dbItem.stockUnit,
            packageQuantity: packageQty,
            packageUnit: dbItem.packageUnit || null,
            baseQuantity: quantity * packageQty,
            baseUnit: dbItem.baseUnit || dbItem.stockUnit
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

    const validReports = Array.from(latestReportsMap.values()).sort((a, b) => a.reportDate.localeCompare(b.reportDate));

    const stats = {
        totalDays: validReports.length,
        workDays: 0,
        noWorkDays: 0,
        totalManDays: 0,
        contractorStats: {},
        materialStats: {},
        materialDetails: {}, 
        reporterStats: {}
    };

    for (const report of validReports) {
        if (report.isNoWork) {
            stats.noWorkDays++;
        } else {
            stats.workDays++;
            stats.totalManDays += Number(report.totalWorkerCount) || 0;

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

            if (Array.isArray(report.materialItems)) {
                report.materialItems.forEach(item => {
                    const key = item.materialId || `${item.materialName} (${item.baseUnit})`;
                    stats.materialStats[key] = (stats.materialStats[key] || 0) + Number(item.baseQuantity || 0);
                    stats.materialDetails[key] = { name: item.materialName, unit: item.baseUnit || '' };
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

// Webhook 必須使用 express.raw，維持獨立
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
            
            if (event.type === 'message' && event.message.type === 'text') {
                const text = event.message.text.trim();
                
                if (text.startsWith('設定案場')) {
                    if (!targetId) { await replyLineMessage(event.replyToken, '⚠️ 請在施工群組內使用。'); continue; }
                    const match = text.match(/^設定案場\s+(.+)$/);
                    if (!match) { await replyLineMessage(event.replyToken, '⚠️ 格式錯誤\n正確格式：設定案場 大安區'); continue; }
                    
                    try {
                        const reg = await registerProjectByName(match[1]);
                        await withBindingWriteLock(async () => {
                            const config = await readBindingsFromOneDrive();
                            const bindings = (config.bindings || []).filter(b => b.projectId !== reg.project.projectId && b.groupId !== targetId);
                            bindings.push({ 
                                projectId: reg.project.projectId, projectName: reg.project.projectName, 
                                groupId: targetId, active: true 
                            });
                            await writeBindingsToOneDrive({ ...config, bindings });
                        });
                        await replyLineMessage(event.replyToken, `✅ 案場設定完成\n\n網址：https://liff.line.me/${LIFF_ID}/?projectId=${encodeURIComponent(reg.project.projectId)}`);
                    } catch (error) {
                        await replyLineMessage(event.replyToken, `⚠️ 無法建立案場\n${getProjectRegistrationErrorMessage(error)}`);
                    }
                }
                else if (text === '查詢案場' || text === '案場查詢') {
                    if (!targetId) continue;
                    const config = await readBindingsFromOneDrive();
                    const binding = (config.bindings || []).find(b => b.groupId === targetId && b.active);
                    await replyLineMessage(event.replyToken, binding ? `📍 本群組綁定案場：\n${binding.projectName}` : '⚠️ 尚未設定案場');
                }
                else if (text === '解除案場') {
                    if (!targetId) continue;
                    await withBindingWriteLock(async () => {
                        const config = await readBindingsFromOneDrive();
                        const filtered = (config.bindings || []).filter(b => b.groupId !== targetId);
                        if (filtered.length === (config.bindings || []).length) { await replyLineMessage(event.replyToken, '無綁定紀錄。'); return; }
                        await writeBindingsToOneDrive({ ...config, bindings: filtered });
                        await replyLineMessage(event.replyToken, '✅ 已解除綁定。');
                    });
                }
                else if (text === '查詢統計' || text === '案場統計') {
                    if (!targetId) continue;
                    const bindingConfig = await readBindingsFromOneDrive();
                    const currentBinding = (bindingConfig.bindings || []).find(b => b.groupId === targetId && b.active);
                    if (!currentBinding) { 
                        await replyLineMessage(event.replyToken, '⚠️ 本群組目前沒有綁定案場'); continue; 
                    }

                    const config = await readProjectsFromOneDrive();
                    const project = (config.projects || []).find(p => p.projectId === currentBinding.projectId);
                    if (!project) continue;

                    try {
                        const result = await generateProjectStats(project);
                        if (result.error || !result.stats) { 
                            await replyLineMessage(event.replyToken, '⚠️ 查詢失敗'); continue; 
                        }
                        
                        const stats = result.stats;
                        let msg = `【${project.projectName}】累計統計表\n━━━━━━━━━━━━\n實際工作天：${stats.workDays} 天\n免計工作天：${stats.noWorkDays} 天\n全案總人天：${stats.totalManDays} 人天\n\n[ 各廠商出工統計 ]\n`;
                        
                        for (const [name, data] of Object.entries(stats.contractorStats)) {
                            msg += ` • ${name}：${data.workDays}天 (${data.manDays}人天)\n`;
                        }
                        
                        msg += `\n[ 材料累計消耗 ]\n`;
                        for (const [key, qty] of Object.entries(stats.materialStats)) {
                            const detail = stats.materialDetails[key] || { name: key, unit: '' };
                            msg += ` • ${detail.name}：共 ${qty} ${detail.unit}\n`;
                        }
                        
                        await replyLineMessage(event.replyToken, msg);
                    } catch (err) {
                        await replyLineMessage(event.replyToken, '⚠️ 統計發生錯誤，請稍後再試。');
                    }
                }
                else if (['指令', '說明', '功能', '小幫手'].includes(text)) {
                    await replyLineMessage(event.replyToken, '📖 「云說工程小幫手」指令：\n\n🔹 設定案場 案場名稱\n🔹 查詢案場\n🔹 查詢統計\n🔹 解除案場\n🔹 結案 案場名稱');
                }
                else if (text.startsWith('結案')) {
                    if (!targetId) continue;
                    const match = text.match(/^結案\s+(.+)$/);
                    if (!match) continue;
                    
                    const bindings = await readBindingsFromOneDrive();
                    const b = (bindings.bindings || []).find(x => x.groupId === targetId && x.active);
                    if (!b || normalizeProjectName(b.projectName) !== normalizeProjectName(match[1])) {
                        await replyLineMessage(event.replyToken, `⚠️ 名稱不符或無綁定`); 
                        continue;
                    }

                    await withProjectWriteLock(async () => {
                        const config = await readProjectsFromOneDrive();
                        const pIdx = (config.projects || []).findIndex(p => p.projectId === b.projectId);
                        if (pIdx === -1) {
                            await replyLineMessage(event.replyToken, '⚠️ 系統找不到此案場資料。'); return;
                        }
                        
                        const closingProject = config.projects[pIdx];

                        try {
                            const resExcel = await fetch(`http://localhost:${PORT}/api/projects/${b.projectId}/export-excel`);
                            if (!resExcel.ok) {
                                const errorText = await resExcel.text();
                                throw new Error(`HTTP ${resExcel.status} ${errorText}`);
                            }
                            
                            const contentType = resExcel.headers.get('content-type') || '';
                            if (!contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
                                throw new Error('結案 API 未回傳 Excel 檔案');
                            }
                            
                            const buf = Buffer.from(await resExcel.arrayBuffer());
                            const { dateStr } = getTaiwanDateParts();
                            const gClient = await getGraphClient();
                            
                            await gClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/工程專案管理/2026_工程專案/${sanitizePathSegment(b.projectName)}/結案總表_${sanitizePathSegment(b.projectName)}_${dateStr.replace(/-/g, '')}.xlsx:/content`).put(buf);
                                
                        } catch(e) { 
                            console.error('結案失敗', e);
                            await replyLineMessage(event.replyToken, `⚠️ 結案報表產生異常 (${e.message})\n案場尚未下架。`); 
                            return; 
                        }
                        
                        config.projects.splice(pIdx, 1);
                        await writeProjectsToOneDrive(config);
                        
                        await withBindingWriteLock(async () => {
                            const latestB = await readBindingsFromOneDrive();
                            await writeBindingsToOneDrive({ 
                                ...latestB, 
                                bindings: (latestB.bindings || []).filter(x => x.projectId !== b.projectId) 
                            });
                        });
                        
                        await replyLineMessage(event.replyToken, `✅ 案場「${match[1]}」已成功結案！\n\n系統已自動產生【Excel 結案報表】與統計資料，並存入您的 OneDrive 資料夾中。`);
                    });
                }
            }
        } catch (e) { console.error('Webhook Error', e); }
    }
});

// ⭐ API 路由前必須套用 express.json()
app.use('/api', express.json());

app.get('/api/projects', async (req, res) => {
    try {
        const config = await readProjectsFromOneDrive();
        const activeProjects = (config.projects || [])
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
        let project = null;
        if (projectId) {
            project = await findProjectById(projectId);
        }
        
        const inventoryMap = await buildInventoryMap(project);
        const materials = Object.values(inventoryMap);
        
        return res.status(200).json({ success: true, materials });
    } catch (error) {
        console.error('讀取材料清單失敗：', error);
        return res.status(500).json({ success: false, error: '無法取得材料清單' });
    }
});

app.get('/api/projects/:projectId/material-balances', async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId);
        if (!project) return res.status(404).json({ success: false, error: '找不到指定案場' });

        const safeProjectName = sanitizePathSegment(project.projectName);
        const txPath = `工程專案管理/2026_工程專案/${safeProjectName}/project-material-transactions.json`;
        let txData = { transactions: [] };
        try { txData = await readJsonFromOneDrive(txPath, { transactions: [] }, false); } catch(e){}

        const issuedStats = {};
        (txData.transactions || []).forEach(tx => {
            const key = tx.materialId || `${tx.materialName} (${tx.baseUnit})`;
            issuedStats[key] = (issuedStats[key] || 0) + Number(tx.baseQuantity || 0);
        });

        const statsResult = await generateProjectStats(project);
        const consumedStats = statsResult.stats?.materialStats || {};

        const inventoryMap = await buildInventoryMap(project);

        const balances = Object.values(inventoryMap).map(m => {
            const key = m.materialId;
            const issued = issuedStats[key] || 0;
            const consumed = consumedStats[key] || 0;
            return {
                materialId: m.materialId,
                materialName: m.materialName,
                packageQuantity: m.packageQuantity || 1,
                packageUnit: m.packageUnit || '',
                stockUnit: m.stockUnit,
                baseUnit: m.baseUnit || m.stockUnit,
                issuedBaseQuantity: issued,
                consumedBaseQuantity: consumed,
                remainingBaseQuantity: issued - consumed
            };
        });
        
        return res.status(200).json({ success: true, projectId: project.projectId, balances });
    } catch (error) { 
        return res.status(500).json({ success: false, error: '取得餘額失敗' }); 
    }
});

app.get('/api/projects/:projectId/export-excel', async (req, res) => {
    try {
        const projectId = req.params.projectId;
        const project = await findProjectById(projectId);
        
        if (!project) {
            return res.status(404).json({ success: false, message: '找不到此專案' });
        }
        
        const projectName = project.projectName;
        const projectBasePath = `工程專案管理/2026_工程專案/${projectName}`;

        const inventoryMap = await buildInventoryMap(project);

        let transactionData = await readJsonFromOneDrive(`${projectBasePath}/project-material-transactions.json`, { transactions: [] }, false);
        const { stats, reports, dataQuality } = await generateProjectStats(project);

        const workbook = new ExcelJS.Workbook();
        workbook.creator = '工程專案自動化系統';
        const resolveMaterialCode = (item) => (item.materialId && inventoryMap[item.materialId]) ? inventoryMap[item.materialId].materialCode : (item.materialCode || '無編碼');

        const wsSummary = workbook.addWorksheet('案場總表');
        wsSummary.views = [{ showGridLines: true }];
        
        const startDate = reports.length > 0 ? reports[0].reportDate : '未定';
        const endDate = reports.length > 0 ? reports[reports.length - 1].reportDate : '未定';
        let calDays = 0;
        if(startDate !== '未定' && endDate !== '未定') {
            calDays = Math.floor((Date.parse(`${endDate}T00:00:00+08:00`) - Date.parse(`${startDate}T00:00:00+08:00`)) / 86400000) + 1;
        }
        
        wsSummary.addRow(['案場名稱', projectName]);
        wsSummary.addRow(['開案日期', startDate]);
        wsSummary.addRow(['結案日期', endDate]);
        wsSummary.addRow(['累計日曆天', calDays]);
        wsSummary.addRow(['實際工作天', stats.workDays]);
        wsSummary.addRow(['免計工作天', stats.noWorkDays]);
        wsSummary.addRow(['全案總人天', stats.totalManDays]);
        wsSummary.addRow([]);
        
        wsSummary.addRow(['【各廠商出工統計】']);
        wsSummary.addRow(['廠商名稱', '出工工作天', '累計人天', '平均每日人數', '占全案人天比例']);
        for (const [name, data] of Object.entries(stats.contractorStats)) {
            const ratio = stats.totalManDays > 0 ? ((data.manDays / stats.totalManDays) * 100).toFixed(1) + '%' : '0%';
            wsSummary.addRow([name, data.workDays, data.manDays, (data.manDays / data.workDays).toFixed(1), ratio]);
        }
        wsSummary.addRow([]);
        
        wsSummary.addRow(['【各填表人填報統計】']);
        wsSummary.addRow(['填表人', '出工天數']);
        for (const [name, days] of Object.entries(stats.reporterStats)) {
            wsSummary.addRow([name, days]);
        }

        const wsMaterials = workbook.addWorksheet('材料結案總表');
        wsMaterials.views = [{ showGridLines: true }];
        wsMaterials.addRow(['材料分類編碼', '材料名稱', '包裝規格', '庫存單位', '案場領入數量', '領入換算量', '日報累計耗用', '理論剩餘', '基準單位']);
        
        const materialSummaryMap = {};
        (transactionData.transactions || []).forEach(tx => {
            const pkgSpec = `${tx.packageQuantity||1}${tx.packageUnit||''}/${tx.stockUnit}`;
            const uniqueKey = tx.materialId || `${tx.materialName}_${pkgSpec}`;
            
            if (!materialSummaryMap[uniqueKey]) {
                materialSummaryMap[uniqueKey] = { 
                    materialCode: resolveMaterialCode(tx), 
                    materialName: tx.materialName, 
                    packageSpec: pkgSpec, 
                    stockUnit: tx.stockUnit, 
                    issuedQty: 0, 
                    baseIssuedQty: 0, 
                    consumedBaseQty: 0, 
                    baseUnit: tx.baseUnit 
                };
            }
            materialSummaryMap[uniqueKey].issuedQty += Number(tx.quantity || 0); 
            materialSummaryMap[uniqueKey].baseIssuedQty += Number(tx.baseQuantity || 0);
        });
        
        reports.forEach(report => {
            (report.materialItems || []).forEach(item => {
                const pkgSpec = `${item.packageQuantity||1}${item.packageUnit||''}/${item.stockUnit}`;
                const uniqueKey = item.materialId || `${item.materialName}_${pkgSpec}`;
                
                if (!materialSummaryMap[uniqueKey]) {
                    materialSummaryMap[uniqueKey] = { 
                        materialCode: resolveMaterialCode(item), 
                        materialName: item.materialName, 
                        packageSpec: pkgSpec, 
                        stockUnit: item.stockUnit, 
                        issuedQty: 0, 
                        baseIssuedQty: 0, 
                        consumedBaseQty: 0, 
                        baseUnit: item.baseUnit 
                    };
                }
                materialSummaryMap[uniqueKey].consumedBaseQty += Number(item.baseQuantity || 0);
            });
        });
        
        let matRowIdx = 2;
        Object.values(materialSummaryMap).forEach(m => {
            wsMaterials.addRow([
                m.materialCode, m.materialName, m.packageSpec, m.stockUnit, 
                m.issuedQty, m.baseIssuedQty, m.consumedBaseQty, 
                { formula: `=F${matRowIdx}-G${matRowIdx}`, result: m.baseIssuedQty - m.consumedBaseQty }, 
                m.baseUnit
            ]);
            matRowIdx++;
        });

        const wsTxLog = workbook.addWorksheet('材料進出紀錄');
        wsTxLog.views = [{ showGridLines: true }];
        wsTxLog.addRow(['日期', '異動類型', '材料分類編碼', '材料名稱', '包裝規格', '原始數量', '庫存單位', '換算後數量', '基準單位', '備註']);
        
        const typeMap = { 'OPENING_ISSUE': '開工領入', 'ADDITIONAL_ISSUE': '追加領入' };
        
        (transactionData.transactions || []).forEach(tx => {
            wsTxLog.addRow([
                tx.transactionDate, 
                typeMap[tx.transactionType] || tx.transactionType, 
                resolveMaterialCode(tx), 
                tx.materialName, 
                `${tx.packageQuantity||1}${tx.packageUnit||''}/${tx.stockUnit}`, 
                tx.quantity, tx.stockUnit, tx.baseQuantity, tx.baseUnit, tx.remarks || ''
            ]);
        });
        
        reports.forEach(r => {
            (r.materialItems || []).forEach(item => {
                wsTxLog.addRow([
                    r.reportDate, '施工耗用', 
                    resolveMaterialCode(item), 
                    item.materialName, 
                    `${item.packageQuantity||1}${item.packageUnit||''}/${item.stockUnit}`, 
                    item.quantity, item.stockUnit, item.baseQuantity, item.baseUnit, item.remarks || '日報自動記錄'
                ]);
            });
        });

        const wsDaily = workbook.addWorksheet('日報明細');
        wsDaily.views = [{ showGridLines: true }];
        wsDaily.addRow(['日期', '填表人', '出工狀態', '無出工原因', '施工廠商', '出工人數', '施作項目', '作業補充', '材料使用摘要', '氣溫', '濕度', '風速', '日報備註']);
        
        reports.forEach(r => {
            wsDaily.addRow([
                r.reportDate, r.reporterName || '未紀錄', r.isNoWork ? '無出工' : '施工', r.noWorkReason || '',
                (r.contractorItems || []).map(c => c.contractorName).join(', ') || '', r.totalWorkerCount || 0,
                (r.workItems || []).join(', ') || '', r.workNotes || '',
                (r.materialItems || []).map(m => `${m.materialName} ${m.quantity}${m.stockUnit}`).join(', ') || '',
                r.weather?.temp || '', r.weather?.humidity || '', r.weather?.wind || '', r.remarks || ''
            ]);
        });

        const wsQuality = workbook.addWorksheet('資料品質');
        wsQuality.views = [{ showGridLines: true }];
        wsQuality.addRow(['【本次結案資料品質與健檢摘要】']);
        wsQuality.addRow(['統計項目', '數量／內容']);
        wsQuality.addRow(['原始 JSON 總數', dataQuality?.sourceFileCount || reports.length]);
        wsQuality.addRow(['成功解析並納入日報數', dataQuality?.effectiveReportCount || reports.length]);
        wsQuality.addRow(['異常資料數量 (解析失敗等)', dataQuality?.invalidFileCount || 0]);
        wsQuality.addRow(['同日重複舊版排除數', dataQuality?.supersededReportCount || 0]);
        wsQuality.addRow(['報表產生時間', new Date().toISOString().replace('T', ' ').substring(0, 19)]);
        
        workbook.eachSheet(worksheet => { 
            for (let i = 1; i <= 15; i++) { 
                worksheet.getColumn(i).width = 22; 
                worksheet.getColumn(i).alignment = { vertical: 'middle', wrapText: true }; 
            } 
        });

        const excelBuffer = await workbook.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(`結案總表_${projectName}.xlsx`)}`);
        res.send(excelBuffer);
    } catch (error) { 
        console.error('產出 Excel 發生錯誤', error);
        res.status(500).json({ success: false, message: '產出 Excel 失敗', error: error.message }); 
    }
});

// ⭐ 單一案場查詢 API 補回
app.get('/api/projects/:projectId', async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId);
        if (!project) {
            return res.status(404).json({ success: false, error: '找不到指定案場' });
        }
        
        return res.status(200).json({
            success: true,
            project: {
                projectId: project.projectId,
                projectName: project.projectName
            }
        });
    } catch (error) {
        console.error('取得案場資料失敗：', error);
        return res.status(500).json({ success: false, error: '無法取得案場資料' });
    }
});

app.post('/api/submit-report', async (req, res) => {
    try {
        const reportData = req.body || {}; 
        const formType = reportData.formType || 'daily_report';

        const submittedProjectId = String(reportData.projectId || '').trim();
        let project = submittedProjectId ? await findProjectById(submittedProjectId) : await findProjectByName(reportData.projectName);
        
        if (!project) {
            return res.status(400).json({ success: false, error: '找不到指定案場' });
        }

        const graphClient = await getGraphClient();
        const safeProjectName = sanitizePathSegment(project.projectName);
        const projectFolderPath = `工程專案管理/2026_工程專案/${safeProjectName}`;
        
        // ⭐ 日期宣告邏輯
        const { dateStr, timeStr } = getTaiwanDateParts();
        let reportDate = dateStr;
        if (formType === 'material_issue') {
            try { reportDate = validateIssueDate(reportData.date, dateStr); }
            catch (dateError) { return res.status(400).json({ success: false, error: dateError.message }); }
        }
        const submitDate = reportDate;

        const inventoryMap = await buildInventoryMap(project);

        // ============================
        // 處理材料進場模式
        // ============================
        if (formType === 'material_issue') {
            // ⭐ 後端進場類型驗證
            if (!['OPENING', 'ADDITIONAL'].includes(reportData.issueType)) {
                return res.status(400).json({ success: false, error: '進場類型不正確' });
            }

            if (!reportData.materialItems || reportData.materialItems.length === 0) {
                return res.status(400).json({ success: false, error: '請至少選擇一項進場材料' });
            }

            const submissionId = String(reportData.submissionId || '').trim();
            if (!submissionId) return res.status(400).json({ success: false, error: '材料進場缺少 submissionId' });

            let materialItems;
            try {
                materialItems = normalizeMaterialItems(reportData.materialItems, false, inventoryMap);
            } catch (materialError) {
                return res.status(400).json({ success: false, error: materialError.message });
            }

            let isDuplicateSubmission = false;
            await withMaterialWriteLock(project.projectId, async () => {
                const txPath = `${projectFolderPath}/project-material-transactions.json`;
                let txData = { transactions: [] };
                try {
                    txData = await readJsonFromOneDrive(txPath, { transactions: [] }, false);
                } catch (e) {}
                if (!Array.isArray(txData.transactions)) txData.transactions = [];
                isDuplicateSubmission = txData.transactions.some(tx => tx.submissionId === submissionId);
                if (isDuplicateSubmission) return;

                const issueType = reportData.issueType === 'ADDITIONAL' ? 'ADDITIONAL_ISSUE' : 'OPENING_ISSUE';
                
                materialItems.forEach(m => {
                    txData.transactions.push({
                        submissionId,
                        transactionDate: submitDate,
                        transactionType: issueType,
                        materialId: m.materialId,
                        materialCode: m.materialCode,
                        materialName: m.materialName,
                        quantity: m.quantity,
                        stockUnit: m.stockUnit,
                        packageQuantity: m.packageQuantity,
                        packageUnit: m.packageUnit,
                        baseQuantity: m.baseQuantity,
                        baseUnit: m.baseUnit,
                        remarks: reportData.remarks || ''
                    });
                });

                await ensureProjectFolder(project.projectName);
                await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${txPath}:/content`).put(Buffer.from(JSON.stringify(txData, null, 2), 'utf-8'));
            });
            if (isDuplicateSubmission) {
                return res.status(200).json({ success: true, duplicate: true, pushed: false, message: '此筆材料進場先前已完成歸檔，未重複入帳' });
            }

            const reporterNameStr = reportData.reporterName ? String(reportData.reporterName).trim() : '未紀錄';
            const issueTypeLabel = reportData.issueType === 'ADDITIONAL' ? '追加進場' : '開工首批進場';
            let msg = `📦 材料進場通知\n\n日期：${submitDate.replace(/-/g, '/')}\n案場：${project.projectName}\n填表：${reporterNameStr}\n類型：${issueTypeLabel}\n\n━━━━━━━━━━━━\n[進場明細]\n`;
            
            materialItems.forEach(m => {
                msg += ` • ${m.materialName}：${m.quantity} ${m.stockUnit}\n`;
            });
            if (reportData.remarks) msg += `\n備註：${reportData.remarks}`;

            let pushed = false;
            const config = await readBindingsFromOneDrive();
            const binding = (Array.isArray(config.bindings) ? config.bindings : []).find(b => b.projectId === project.projectId && b.active);
            
            if (binding) {
                try { 
                    await pushLineMessage(binding.groupId, msg); 
                    pushed = true;
                } catch (e) {
                    console.error('LINE推播失敗', e);
                }
            }

            return res.status(200).json({ success: true, pushed: pushed, message: '材料進場紀錄已成功歸檔' });
        }

        // ============================
        // 處理施工日報模式
        // ============================
        const isNoWork = reportData.isNoWork === true;
        let contractorItems = Array.isArray(reportData.contractorItems) ? reportData.contractorItems : [];
        if (!isNoWork) {
            if (contractorItems.length === 0) return res.status(400).json({ success: false, error: '請至少填寫一組施工廠商' });
            const contractorNames = new Set();
            const normalizedContractors = [];
            for (const item of contractorItems) {
                const contractorName = String(item.contractorName || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
                const workerCount = Number(item.workerCount);
                if (!contractorName) return res.status(400).json({ success: false, error: '施工廠商名稱不可為空' });
                if (!Number.isInteger(workerCount) || workerCount <= 0 || workerCount > 200) return res.status(400).json({ success: false, error: `廠商「${contractorName}」施工人數不正確` });
                if (contractorNames.has(contractorName)) return res.status(400).json({ success: false, error: `施工廠商「${contractorName}」重複填寫` });
                contractorNames.add(contractorName);
                normalizedContractors.push({ contractorName, workerCount });
            }
            contractorItems = normalizedContractors;
        } else contractorItems = [];
        const calculatedTotalWorkerCount = isNoWork ? 0 : contractorItems.reduce((total, item) => total + item.workerCount, 0);

        await ensureProjectFolder(project.projectName);
        const [textFolderResult, dataFolderResult] = await Promise.all([
            ensureChildFolder(graphClient, projectFolderPath, '施工日報'),
            ensureChildFolder(graphClient, projectFolderPath, '結構化資料')
        ]);
        
        const submittedSubmissionId = String(reportData.submissionId || '').trim();
        const fullSubmissionId = submittedSubmissionId || crypto.randomUUID();
        const shortSubmissionId = fullSubmissionId.replace(/[^a-zA-Z0-9]/g, '').slice(0, 16);

        const workItems = Array.isArray(reportData.workItems) ? reportData.workItems : [];
        let materialItems;
        try {
            materialItems = normalizeMaterialItems(reportData.materialItems, isNoWork, inventoryMap);
        } catch (materialError) {
            return res.status(400).json({ success: false, error: materialError.message });
        }

        const structuredReport = {
            schemaVersion: 1, projectId: project.projectId, projectName: project.projectName, reportDate: reportDate,
            submissionId: fullSubmissionId, submittedAt: new Date().toISOString(), submittedDateLocal: dateStr, submittedTimeLocal: timeStr,
            reporterName: String(reportData.reporterName || '未紀錄').trim(),
            isNoWork, noWorkReason: isNoWork ? String(reportData.noWorkReason || '') : '',
            weather: { temp: reportData.temp, humidity: reportData.humidity, wind: reportData.wind },
            contractorItems, totalWorkerCount: calculatedTotalWorkerCount, workItems: isNoWork ? [] : workItems,
            customWorkItem: isNoWork ? '' : String(reportData.customWorkItem || ''), workNotes: isNoWork ? '' : String(reportData.workNotes || ''),
            materialItems, remarks: String(reportData.remarks || '')
        };

        const baseFileName = `${reportDate}_${shortSubmissionId}`;
        const jsonFilePath = `${dataFolderResult.folderPath}/${baseFileName}.json`;
        const txtFilePath = `${textFolderResult.folderPath}/${baseFileName}_施工日報.txt`;

        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${jsonFilePath}:/content`).put(Buffer.from(JSON.stringify(structuredReport, null, 2), 'utf-8'));

        let alertMsgs = [];
        if (!isNoWork && materialItems.length > 0) {
            const txData = await readJsonFromOneDrive(`${projectFolderPath}/project-material-transactions.json`, { transactions: [] }, false);
            const issuedStats = {};
            (txData.transactions || []).forEach(tx => {
                const key = tx.materialId;
                issuedStats[key] = (issuedStats[key] || 0) + Number(tx.baseQuantity || 0);
            });
            
            const sRes = await generateProjectStats(project);
            if (sRes.stats) {
                materialItems.forEach(m => {
                    const key = m.materialId;
                    const totCons = sRes.stats.materialStats[key] || 0;
                    const totIss = issuedStats[key] || 0;
                    const bal = totIss - totCons;
                    
                    if (bal < 0) {
                        alertMsgs.push(`⚠️ ${m.materialName}\n • 累計領入: ${totIss} ${m.baseUnit}\n • 累計耗用: ${totCons} ${m.baseUnit}\n • 理論剩餘: ${bal} ${m.baseUnit}`);
                    }
                });
            }
        }

        const reporterNameStr = reportData.reporterName ? String(reportData.reporterName).trim() : '未紀錄';
        let reportText = `📋 施工日報\n\n日期：${reportDate.replace(/-/g, '/')}\n案場：${project.projectName}\n填表：${reporterNameStr}\n\n`;
        
        if (!isNoWork) {
            reportText += `溫度：${reportData.temp}度\n濕度：${reportData.humidity}%\n風速：${reportData.wind}m/s\n\n施工廠商：${reportData.contractor}\n${reportData.workerCount}\n\n━━━━━━━━━━━━\n\n今日進度：\n${reportData.progress}\n\n今日用料：\n${reportData.materials}\n\n備註：\n${reportData.remarks || '無'}\n\n━━━━━━━━━━━━\n以上為今日進度報告`;
        } else {
            reportText += `🛑 今日無出工\n原因：${reportData.noWorkReason}\n備註：${reportData.remarks || '無'}`;
        }
        
        if (alertMsgs.length > 0) {
            reportText += `\n\n🚨 【系統異常警示：材料帳庫存不足】\n\n` + alertMsgs.join('\n\n') + `\n\n💡 請協助確認是否漏登材料進場`;
        }
        
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${txtFilePath}:/content`).put(reportText);

        let pushed = false;
        const config = await readBindingsFromOneDrive();
        const binding = (Array.isArray(config.bindings) ? config.bindings : []).find(b => b.projectId === project.projectId && b.active);

        if (binding) {
            try { 
                await pushLineMessage(binding.groupId, reportText); 
                pushed = true;
            } catch (e) {
                console.error('LINE推播失敗', e);
            }
        }
        
        return res.status(200).json({ success: true, pushed: pushed, message: '日報已歸檔' });

    } catch (error) {
        console.error('提交錯誤：', error);
        return res.status(500).json({ success: false, error: error.message || '系統內部處理失敗' });
    }
});

const requiredVars = ['LINE_ACCESS_TOKEN', 'LINE_CHANNEL_SECRET', 'AZURE_CLIENT_ID', 'AZURE_TENANT_ID', 'AZURE_CLIENT_SECRET'];
const missingVars = requiredVars.filter(name => !process.env[name]);
if (missingVars.length > 0) {
    console.error('缺少必要環境變數：', missingVars.join(', '));
    process.exit(1);
}

app.listen(PORT, () => console.log(`🚀 伺服器運作中：http://localhost:${PORT}`));
