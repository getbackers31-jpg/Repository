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

// ==========================================
// 系統參數與環境變數設定
// ==========================================
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

// ==========================================
// Microsoft Graph API 客戶端與 Token 管理
// ==========================================
let cachedGraphClient = null;
let tokenExpiresAt = 0;
let graphClientPromise = null;

async function getGraphClient() {
    const now = Date.now();
    if (cachedGraphClient && now < tokenExpiresAt) {
        return cachedGraphClient;
    }
    
    if (graphClientPromise) {
        return graphClientPromise;
    }
    
    graphClientPromise = (async () => {
        try {
            const response = await cca.acquireTokenByClientCredential({ 
                scopes: ['https://graph.microsoft.com/.default'] 
            });
            
            if (!response || !response.accessToken) {
                throw new Error('Microsoft Graph Token 取得失敗');
            }
            
            tokenExpiresAt = Math.max(
                Date.now() + 60 * 1000, 
                (response.expiresOn ? response.expiresOn.getTime() : Date.now() + 50 * 60 * 1000) - 5 * 60 * 1000
            );
            
            cachedGraphClient = Client.init({ 
                authProvider(done) { 
                    done(null, response.accessToken); 
                } 
            });
            
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

// ==========================================
// 共用工具與驗證函式
// ==========================================
function sanitizePathSegment(value) { 
    return String(value).replace(/[<>:"/\\|?*#%]/g, '_').replace(/\s+/g, ' ').trim(); 
}

function normalizeProjectName(value) { 
    return String(value || '').normalize('NFKC').replace(/\s+/g, ' ').trim(); 
}

function validateProjectName(projectName) {
    const normalizedName = normalizeProjectName(projectName);
    if (!normalizedName) {
        throw new Error('案場名稱不可為空');
    }
    if (normalizedName.length > 80) {
        throw new Error('案場名稱不可超過 80 個字');
    }
    if (/[<>:"/\\|?*#%]/.test(normalizedName)) {
        throw new Error('名稱不可含特殊字元');
    }
    return normalizedName;
}

function getProjectRegistrationErrorMessage(error) {
    const safeMessages = ['案場名稱不可為空', '案場名稱不可超過 80 個字', '案場名稱不可包含以下字元'];
    const message = String(error?.message || '');
    return safeMessages.some(prefix => message.startsWith(prefix)) ? message : '系統暫時無法建立案場，請稍後再試';
}

// ==========================================
// 寫入鎖機制 (防範併發寫入覆蓋)
// ==========================================
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

// ==========================================
// OneDrive JSON 讀寫與快取機制
// ==========================================
async function readJsonFromOneDrive(filePath, defaultData, throwOnNotFound = false) {
    try {
        const graphClient = await getGraphClient();
        const meta = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${filePath}`).get();
        const downloadUrl = meta['@microsoft.graph.downloadUrl'];
        
        if (!downloadUrl) {
            throw new Error('未回傳下載網址');
        }
        
        const response = await fetch(downloadUrl);
        if (!response.ok) {
            throw new Error(`下載失敗 ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        if ((error?.statusCode === 404 || error?.code === 'itemNotFound') && !throwOnNotFound) {
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

function cloneJsonData(data) { 
    return JSON.parse(JSON.stringify(data)); 
}

const CACHE_TTL = 30 * 1000;
const configCache = { 
    projects: { data: null, timestamp: 0 }, 
    bindings: { data: null, timestamp: 0 }, 
    globalInventory: { data: null, timestamp: 0 }, 
    projMaterials: {} 
};

async function readProjectsFromOneDrive() {
    if (configCache.projects.data && Date.now() - configCache.projects.timestamp < CACHE_TTL) {
        return cloneJsonData(configCache.projects.data);
    }
    const data = await readJsonFromOneDrive('工程專案管理/_系統設定/projects.json', { projects: [] }, true);
    configCache.projects = { data: cloneJsonData(data), timestamp: Date.now() }; 
    return cloneJsonData(data);
}

async function writeProjectsToOneDrive(config) { 
    await writeJsonToOneDrive('工程專案管理/_系統設定/projects.json', config); 
    configCache.projects = { data: cloneJsonData(config), timestamp: Date.now() }; 
}

async function readBindingsFromOneDrive() {
    if (configCache.bindings.data && Date.now() - configCache.bindings.timestamp < CACHE_TTL) {
        return cloneJsonData(configCache.bindings.data);
    }
    const data = await readJsonFromOneDrive('工程專案管理/_系統設定/line-bindings.json', { bindings: [] });
    configCache.bindings = { data: cloneJsonData(data), timestamp: Date.now() }; 
    return cloneJsonData(data);
}

async function writeBindingsToOneDrive(config) { 
    await writeJsonToOneDrive('工程專案管理/_系統設定/line-bindings.json', config); 
    configCache.bindings = { data: cloneJsonData(config), timestamp: Date.now() }; 
}

async function readGlobalInventory() {
    if (configCache.globalInventory.data && Date.now() - configCache.globalInventory.timestamp < CACHE_TTL) {
        return cloneJsonData(configCache.globalInventory.data);
    }
    const data = await readJsonFromOneDrive('工程專案管理/_系統設定/inventory.json', { items: [] }, false);
    configCache.globalInventory = { data: cloneJsonData(data), timestamp: Date.now() }; 
    return cloneJsonData(data);
}

async function readProjectMaterials(projectName) {
    const safeName = sanitizePathSegment(projectName);
    if (configCache.projMaterials[safeName] && Date.now() - configCache.projMaterials[safeName].timestamp < CACHE_TTL) {
        return cloneJsonData(configCache.projMaterials[safeName].data);
    }
    try {
        const data = await readJsonFromOneDrive(`工程專案管理/2026_工程專案/${safeName}/專屬材料.json`, null, false);
        if (data && (Array.isArray(data.items) || Array.isArray(data.materials))) { 
            configCache.projMaterials[safeName] = { data: cloneJsonData(data), timestamp: Date.now() }; 
            return cloneJsonData(data); 
        }
    } catch (e) {
        // 忽略找不到專屬材料的錯誤
    } 
    return null;
}

// 整合版資料夾建立功能，支援建立案場主目錄或子目錄
async function ensureFolder(graphClient, parentPath, childFolderName = null) {
    const fullPath = childFolderName ? `${parentPath}/${sanitizePathSegment(childFolderName)}` : parentPath;
    try {
        const item = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${fullPath}`).get();
        return { folderId: item.id, folderPath: fullPath };
    } catch (error) {
        if (error?.statusCode !== 404 && error?.code !== 'itemNotFound') {
            throw error;
        }
        const parent = childFolderName ? parentPath : '工程專案管理/2026_工程專案';
        const name = childFolderName ? sanitizePathSegment(childFolderName) : parentPath.split('/').pop();
        const created = await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${parent}:/children`)
            .post({ name: name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
        return { folderId: created.id, folderPath: fullPath };
    }
}

async function findProjectByName(projectName) {
    const config = await readProjectsFromOneDrive();
    return (config.projects || []).find(p => p.active && normalizeProjectName(p.projectName) === normalizeProjectName(projectName)) || null;
}

async function findProjectById(projectId) {
    const config = await readProjectsFromOneDrive();
    return (config.projects || []).find(project => project.active && project.projectId === String(projectId || '').trim()) || null;
}

async function registerProjectByName(projectName) {
    return withProjectWriteLock(async () => {
        const normalizedName = validateProjectName(projectName);
        const config = await readProjectsFromOneDrive();
        const projects = Array.isArray(config.projects) ? config.projects : [];
        
        let existing = projects.find(p => p.active && normalizeProjectName(p.projectName) === normalizedName);
        if (existing) { 
            await ensureFolder(await getGraphClient(), `工程專案管理/2026_工程專案/${existing.projectName}`); 
            return { project: existing }; 
        }
        
        const project = { 
            projectId: `PRJ-${crypto.randomUUID()}`, 
            projectName: normalizedName, 
            active: true, 
            createdAt: new Date().toISOString() 
        };
        
        await ensureFolder(await getGraphClient(), `工程專案管理/2026_工程專案/${project.projectName}`);
        projects.push(project);
        await writeProjectsToOneDrive({ ...config, projects, updatedAt: new Date().toISOString() });
        return { project };
    });
}

// ==========================================
// LINE Messaging API 相關
// ==========================================
async function replyLineMessage(replyToken, text) {
    if (!replyToken) return;
    
    await fetch('https://api.line.me/v2/bot/message/reply', { 
        method: 'POST', 
        headers: { 
            'Content-Type': 'application/json', 
            Authorization: `Bearer ${LINE_ACCESS_TOKEN}` 
        }, 
        body: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] }) 
    });
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
        throw new Error(`LINE Push 失敗：${response.status}`);
    }
}

function verifyLineSignature(rawBody, signature) {
    if (!LINE_CHANNEL_SECRET || !signature) {
        return false;
    }
    
    const expectedSignature = crypto.createHmac('sha256', LINE_CHANNEL_SECRET)
        .update(rawBody)
        .digest('base64');
        
    const actualBuffer = Buffer.from(signature);
    const expectedBuffer = Buffer.from(expectedSignature);
    
    if (actualBuffer.length !== expectedBuffer.length) {
        return false;
    }
    
    return crypto.timingSafeEqual(actualBuffer, expectedBuffer);
}

function getLineTargetId(event) {
    if (event.source?.type === 'group') return event.source.groupId;
    if (event.source?.type === 'room') return event.source.roomId;
    return null;
}

// ==========================================
// 日期與材料處理邏輯
// ==========================================
function getTaiwanDateParts() {
    const parts = new Intl.DateTimeFormat('en-CA', { 
        timeZone: 'Asia/Taipei', 
        year: 'numeric', month: '2-digit', day: '2-digit', 
        hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false 
    }).formatToParts(new Date());
    
    const values = Object.fromEntries(parts.filter(p => p.type !== 'literal').map(p => [p.type, p.value]));
    
    return { 
        dateStr: `${values.year}-${values.month}-${values.day}`, 
        timeStr: `${values.hour}${values.minute}${values.second}` 
    };
}

function normalizeMaterialItems(rawItems, inventoryMap) {
    if (!Array.isArray(rawItems)) return [];
    
    return rawItems.map((item, index) => {
        const qty = Number(item.quantity);
        if (!Number.isFinite(qty) || qty <= 0) {
            throw new Error(`第 ${index + 1} 筆數量不正確`);
        }
        
        const dbItem = inventoryMap[item.materialId];
        if (!dbItem) {
            throw new Error(`找不到指定的標準材料 (ID: ${item.materialId})`);
        }
        
        const packageQty = dbItem.packageQuantity ? Number(dbItem.packageQuantity) : 1;
        
        return {
            materialId: dbItem.materialId, 
            materialCode: dbItem.materialCode || '無編碼', 
            materialName: dbItem.materialName,
            quantity: qty, 
            stockUnit: dbItem.stockUnit, 
            packageQuantity: packageQty, 
            packageUnit: dbItem.packageUnit || null,
            baseQuantity: qty * packageQty, 
            baseUnit: dbItem.baseUnit || dbItem.stockUnit
        };
    });
}

function validateReportData(reportData) {
    const requiredFields = ['projectName', 'contractor', 'workerCount', 'progress', 'materials'];
    const missingFields = requiredFields.filter(field => !reportData[field] || String(reportData[field]).trim() === '');
    return { valid: missingFields.length === 0, missingFields };
}

// ==========================================
// 專案統計生成核心邏輯
// ==========================================
async function generateProjectStats(project) {
    const safeProjectName = sanitizePathSegment(project.projectName);
    const graphClient = await getGraphClient();
    let requestUrl = `/users/${TARGET_USER_EMAIL}/drive/root:/工程專案管理/2026_工程專案/${safeProjectName}/結構化資料:/children`;
    const allItems = [];
    
    try {
        while (requestUrl) {
            const result = await graphClient.api(requestUrl).get();
            if (Array.isArray(result.value)) {
                allItems.push(...result.value.filter(f => f.name.endsWith('.json')));
            }
            requestUrl = result['@odata.nextLink'] || null;
        }
    } catch (e) { 
        return { stats: null, reports: [], dataQuality: null }; 
    }

    const reports = [];
    const invalidFiles = [];
    
    for (const file of allItems) {
        try {
            const response = await fetch(file['@microsoft.graph.downloadUrl']);
            if (response.ok) {
                reports.push(await response.json());
            } else {
                invalidFiles.push({ fileName: file.name, error: `HTTP ${response.status}` });
            }
        } catch (err) {
            invalidFiles.push({ fileName: file.name, error: '解析或讀取失敗' });
        }
    }

    const latestReportsMap = new Map();
    let supersededCount = 0;
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    
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
        
        supersededCount++;
        const existingTime = Date.parse(existing.submittedAt || '');
        if (!Number.isFinite(existingTime) || submittedAtTime > existingTime) {
            latestReportsMap.set(date, report);
        }
    }
    
    const validReports = Array.from(latestReportsMap.values()).sort((a, b) => a.reportDate.localeCompare(b.reportDate));

    const stats = { 
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
            
            (report.contractorItems || []).forEach(item => {
                const name = String(item.contractorName || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
                const count = Number(item.workerCount);
                if (name && count > 0) {
                    if (!stats.contractorStats[name]) {
                        stats.contractorStats[name] = { manDays: 0, workDays: 0 };
                    }
                    stats.contractorStats[name].manDays += count; 
                    stats.contractorStats[name].workDays += 1;
                }
            });
            
            (report.materialItems || []).forEach(item => {
                const key = item.materialId || `${item.materialName} (${item.baseUnit})`;
                stats.materialStats[key] = (stats.materialStats[key] || 0) + Number(item.baseQuantity || 0);
                stats.materialDetails[key] = { name: item.materialName, unit: item.baseUnit || '' };
            });
        }
    }
    
    return { 
        stats, 
        reports: validReports,
        dataQuality: {
            sourceFileCount: allItems.length,
            parsedFileCount: reports.length,
            invalidFileCount: invalidFiles.length,
            effectiveReportCount: validReports.length,
            supersededReportCount: supersededCount
        }
    };
}

// ==========================================
// LINE Webhook 處理
// ==========================================
app.post('/webhook', express.raw({ type: 'application/json' }), async (req, res) => {
    const signature = req.get('x-line-signature');
    if (!verifyLineSignature(req.body, signature)) {
        return res.status(401).send('Invalid signature');
    }
    
    let body; 
    try { 
        body = JSON.parse(req.body.toString('utf8')); 
    } catch (e) { 
        return res.status(400).send('Invalid JSON'); 
    }
    
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
                    if (!targetId) continue;
                    const match = text.match(/^設定案場\s+(.+)$/);
                    if (!match) {
                        await replyLineMessage(event.replyToken, '⚠️ 格式錯誤\n正確格式：設定案場 大安區');
                        continue;
                    }
                    
                    try {
                        const reg = await registerProjectByName(match[1]);
                        await withBindingWriteLock(async () => {
                            const config = await readBindingsFromOneDrive();
                            const bindings = (config.bindings || []).filter(b => b.projectId !== reg.project.projectId && b.groupId !== targetId);
                            bindings.push({ 
                                projectId: reg.project.projectId, 
                                projectName: reg.project.projectName, 
                                groupId: targetId, 
                                active: true 
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
                else if (text === '查詢統計' || text === '案場統計') {
                    if (!targetId) continue;
                    const bindingConfig = await readBindingsFromOneDrive();
                    const currentBinding = (bindingConfig.bindings || []).find(b => b.groupId === targetId && b.active);
                    if (!currentBinding) { 
                        await replyLineMessage(event.replyToken, '⚠️ 本群組目前沒有綁定案場'); 
                        continue; 
                    }

                    const config = await readProjectsFromOneDrive();
                    const project = (config.projects || []).find(p => p.projectId === currentBinding.projectId);
                    if (!project) continue;

                    try {
                        const result = await generateProjectStats(project);
                        if (result.error || !result.stats) { 
                            await replyLineMessage(event.replyToken, '⚠️ 查詢失敗'); 
                            continue; 
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
                        console.error('群組查詢統計失敗：', err);
                        await replyLineMessage(event.replyToken, '⚠️ 統計計算過程中發生錯誤，請稍後再試。');
                    }
                }
                else if (text === '解除案場') {
                    if (!targetId) continue;
                    await withBindingWriteLock(async () => {
                        const config = await readBindingsFromOneDrive();
                        const filtered = (config.bindings || []).filter(b => b.groupId !== targetId);
                        if (filtered.length === (config.bindings || []).length) { 
                            await replyLineMessage(event.replyToken, '無綁定紀錄。'); 
                            return; 
                        }
                        await writeBindingsToOneDrive({ ...config, bindings: filtered });
                        await replyLineMessage(event.replyToken, '✅ 已解除綁定。');
                    });
                }
                else if (['指令', '說明', '功能', '小幫手', '【點此查看指令說明】'].includes(text)) {
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
                            await replyLineMessage(event.replyToken, '⚠️ 系統找不到此案場資料。');
                            return;
                        }
                        
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
                            
                            await gClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/工程專案管理/2026_工程專案/${sanitizePathSegment(b.projectName)}/結案總表_${sanitizePathSegment(b.projectName)}_${dateStr.replace(/-/g, '')}.xlsx:/content`)
                                .put(buf);
                                
                        } catch(e) { 
                            console.error('結案失敗', e);
                            await replyLineMessage(event.replyToken, `⚠️ 結案失敗：報表產生異常 (${e.message})\n案場尚未下架。`); 
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
                        
                        await replyLineMessage(event.replyToken, `✅ 「${match[1]}」已結案並產生完整 Excel 報表！`);
                    });
                }
            }
        } catch (e) { 
            console.error('Webhook Error', e); 
        }
    }
});

// ==========================================
// Express API 路由
// ==========================================

// ⭐ [補回] 中介軟體：驗證 API 金鑰
function verifyStatsApiKey(req, res, next) {
    const apiKey = req.get('x-api-key');
    if (!STATS_API_KEY || apiKey !== STATS_API_KEY) {
        return res.status(401).json({ success: false, error: '未授權存取' });
    }
    next();
}

// ⭐ [補回] 健康檢查路由
app.get('/', (req, res) => {
    res.send('✅ 伺服器運作中！');
});

// ⭐ [補回] 外部統計查詢 API
app.get('/api/project-stats/:projectId', verifyStatsApiKey, async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId);
        if (!project) {
            return res.status(404).json({ success: false, error: '找不到該案場' });
        }
        
        const result = await generateProjectStats(project);
        if (result.error) {
            return res.status(200).json({ success: true, message: result.error });
        }

        return res.status(200).json({ success: true, projectName: project.projectName, ...result });
    } catch (error) {
        console.error('統計產生失敗：', error);
        return res.status(500).json({ success: false, error: '統計產生失敗' });
    }
});

app.get('/api/projects', async (req, res) => {
    try {
        const config = await readProjectsFromOneDrive();
        const activeProjects = (config.projects || [])
            .filter(p => p.active)
            .map(p => ({ 
                projectId: p.projectId, 
                projectName: p.projectName 
            }))
            .sort((a, b) => a.projectName.localeCompare(b.projectName, 'zh-Hant'));
            
        return res.status(200).json({ success: true, projects: activeProjects });
    } catch (error) {
        console.error('讀取案場清單失敗：', error);
        return res.status(500).json({ success: false, error: '無法取得案場清單' });
    }
});

app.get('/api/projects/:projectId', async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId);
        if (!project) {
            return res.status(404).json({ success: false, error: '找不到指定案場' });
        }
        return res.status(200).json({ success: true, project });
    } catch (error) {
        console.error('取得案場資訊發生錯誤', error);
        return res.status(500).json({ success: false, error: '取得案場資訊發生錯誤' });
    }
});

app.get('/api/materials', async (req, res) => {
    try {
        const projectId = req.query.projectId;
        if (projectId) {
            const project = await findProjectById(projectId);
            if (project) {
                const customConfig = await readProjectMaterials(project.projectName);
                if (customConfig?.items || customConfig?.materials) {
                    return res.status(200).json({ success: true, materials: customConfig.items || customConfig.materials });
                }
            }
        }
        
        const globalConfig = await readGlobalInventory();
        return res.status(200).json({ success: true, materials: globalConfig.items || globalConfig.materials || [] });
    } catch (error) {
        console.error('讀取材料清單失敗', error);
        return res.status(500).json({ success: false, error: '無法取得材料清單' });
    }
});

app.get('/api/projects/:projectId/material-balances', async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId);
        if (!project) {
            return res.status(404).json({ success: false, error: '找不到指定案場' });
        }

        const txPath = `工程專案管理/2026_工程專案/${sanitizePathSegment(project.projectName)}/project-material-transactions.json`;
        const txData = await readJsonFromOneDrive(txPath, { transactions: [] }, false);
        const issuedStats = {};
        
        (txData.transactions || []).forEach(tx => {
            const key = tx.materialId || `${tx.materialName} (${tx.baseUnit})`;
            issuedStats[key] = (issuedStats[key] || 0) + Number(tx.baseQuantity || 0);
        });

        const statsResult = await generateProjectStats(project);
        const consumedStats = statsResult.stats?.materialStats || {};

        const globalInventory = await readGlobalInventory();
        const customConfig = await readProjectMaterials(project.projectName);
        const inventoryMap = {};
        
        (globalInventory?.items || globalInventory?.materials || []).forEach(m => inventoryMap[m.materialId] = m);
        (customConfig?.items || customConfig?.materials || []).forEach(m => inventoryMap[m.materialId] = m);

        const balances = Object.values(inventoryMap).map(m => {
            const key = m.materialId || `${m.materialName} (${m.baseUnit || m.stockUnit})`;
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
        console.error('取得餘額失敗', error);
        return res.status(500).json({ success: false, error: '取得餘額失敗' }); 
    }
});

app.post('/api/submit-report', async (req, res) => {
    try {
        const reportData = req.body || {}; 
        const formType = reportData.formType || 'daily_report';
        const project = await findProjectById(reportData.projectId);
        
        if (!project) {
            return res.status(400).json({ success: false, error: '案場不存在' });
        }

        const graphClient = await getGraphClient();
        const safeProjectName = sanitizePathSegment(project.projectName);
        const projectPath = `工程專案管理/2026_工程專案/${safeProjectName}`;
        
        const { dateStr, timeStr } = getTaiwanDateParts();
        const reportDate = (reportData.date && /^\d{4}-\d{2}-\d{2}$/.test(reportData.date)) ? reportData.date : dateStr;

        const globalInv = await readGlobalInventory();
        const customInv = await readProjectMaterials(project.projectName);
        const invMap = {};
        
        (globalInv?.items || globalInv?.materials || []).forEach(m => invMap[m.materialId] = m);
        (customInv?.items || customInv?.materials || []).forEach(m => invMap[m.materialId] = m);

        // ============================
        // 處理材料進場模式
        // ============================
        if (formType === 'material_issue') {
            if (!reportData.materialItems?.length) {
                return res.status(400).json({ success: false, error: '請選擇進場材料' });
            }
            
            const mItems = normalizeMaterialItems(reportData.materialItems, invMap);

            await withMaterialWriteLock(project.projectId, async () => {
                const txPath = `${projectPath}/project-material-transactions.json`;
                const txData = await readJsonFromOneDrive(txPath, { transactions: [] }, false);
                const issueType = reportData.issueType === 'OPENING' ? 'OPENING_ISSUE' : 'ADDITIONAL_ISSUE';
                
                mItems.forEach(m => {
                    txData.transactions.push({
                        transactionDate: reportDate, 
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
                
                await ensureFolder(graphClient, projectPath);
                await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${txPath}:/content`)
                    .put(Buffer.from(JSON.stringify(txData, null, 2)));
            });

            const rName = String(reportData.reporterName || '未紀錄').trim();
            let msg = `📦 材料進場通知\n\n日期：${reportDate.replace(/-/g, '/')}\n案場：${project.projectName}\n填表：${rName}\n\n[進場明細]\n`;
            
            mItems.forEach(m => { 
                msg += ` • ${m.materialName}：${m.quantity} ${m.stockUnit}\n`; 
            });
            
            let pushed = false;
            const bindings = await readBindingsFromOneDrive();
            const b = (bindings.bindings || []).find(x => x.projectId === project.projectId && x.active);
            
            if (b) { 
                try { 
                    await pushLineMessage(b.groupId, msg); 
                    pushed = true; 
                } catch(e) {
                    console.error('LINE Push Failed', e);
                } 
            }
            
            return res.status(200).json({ success: true, pushed });
        }

        // ============================
        // 處理施工日報模式
        // ============================
        const isNoWork = reportData.isNoWork === true;
        const mItems = isNoWork || (reportData.materialItems?.length === 0) ? [] : normalizeMaterialItems(reportData.materialItems, invMap);
        const cItems = isNoWork ? [] : (reportData.contractorItems || []);
        
        await ensureFolder(graphClient, projectPath);
        const [txtFolder, jsonFolder] = await Promise.all([
            ensureFolder(graphClient, projectPath, '施工日報'), 
            ensureFolder(graphClient, projectPath, '結構化資料')
        ]);
        
        const subId = crypto.randomUUID();

        const structuredReport = {
            schemaVersion: 1, 
            projectId: project.projectId, 
            projectName: project.projectName, 
            reportDate: reportDate, 
            submissionId: subId, 
            submittedAt: new Date().toISOString(),
            reporterName: String(reportData.reporterName || '未紀錄').trim(), 
            isNoWork, 
            noWorkReason: isNoWork ? reportData.noWorkReason : '',
            weather: reportData.temp ? { temp: reportData.temp, humidity: reportData.humidity, wind: reportData.wind } : {},
            contractorItems: cItems, 
            totalWorkerCount: cItems.reduce((acc, c) => acc + Number(c.workerCount), 0),
            workItems: reportData.workItems || [], 
            workNotes: reportData.workNotes || '', 
            materialItems: mItems, 
            remarks: reportData.remarks || ''
        };

        const baseFileName = `${reportDate}_${subId.replace(/-/g, '').slice(0, 8)}`;
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${jsonFolder.folderPath}/${baseFileName}.json:/content`)
            .put(Buffer.from(JSON.stringify(structuredReport, null, 2)));

        let alertMsgs = [];
        if (!isNoWork && mItems.length > 0) {
            const txData = await readJsonFromOneDrive(`${projectPath}/project-material-transactions.json`, { transactions: [] }, false);
            const issuedStats = {};
            
            (txData.transactions || []).forEach(tx => {
                const key = tx.materialId || `${tx.materialName} (${tx.baseUnit})`;
                issuedStats[key] = (issuedStats[key] || 0) + Number(tx.baseQuantity || 0);
            });
            
            const sRes = await generateProjectStats(project);
            if (sRes.stats) {
                mItems.forEach(m => {
                    const key = m.materialId || `${m.materialName} (${m.baseUnit})`;
                    const totCons = sRes.stats.materialStats[key] || 0;
                    const totIss = issuedStats[key] || 0;
                    const bal = totIss - totCons;
                    
                    if (bal < 0) {
                        alertMsgs.push(`⚠️ ${m.materialName}\n • 累計領入: ${totIss} ${m.baseUnit}\n • 累計耗用: ${totCons} ${m.baseUnit}\n • 理論剩餘: ${bal} ${m.baseUnit}`);
                    }
                });
            }
        }

        let reportText = `📋 施工日報\n\n日期：${reportDate.replace(/-/g, '/')}\n案場：${project.projectName}\n填表：${structuredReport.reporterName}\n\n`;
        
        if (!isNoWork) {
            reportText += `溫度：${reportData.temp}度\n濕度：${reportData.humidity}%\n風速：${reportData.wind}m/s\n\n施工廠商：\n${reportData.contractor}\n${reportData.workerCount}\n\n━━━━━━━━━━━━\n\n今日進度：\n${reportData.progress}\n\n今日用料：\n${reportData.materials}\n\n備註：\n${reportData.remarks || '無'}\n\n━━━━━━━━━━━━\n以上為今日報告`;
        } else {
            reportText += `🛑 今日無出工\n原因：${reportData.noWorkReason}\n備註：${reportData.remarks || '無'}`;
        }
        
        if (alertMsgs.length > 0) {
            reportText += `\n\n🚨 【系統異常警示：材料帳庫存不足】\n\n` + alertMsgs.join('\n\n') + `\n\n💡 請協助確認是否漏登材料進場`;
        }
        
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${txtFolder.folderPath}/${baseFileName}_施工日報.txt:/content`)
            .put(reportText);

        let pushed = false;
        const bindings = await readBindingsFromOneDrive();
        const b = (bindings.bindings || []).find(x => x.projectId === project.projectId && x.active);
        
        if (b) { 
            try { 
                await pushLineMessage(b.groupId, reportText); 
                pushed = true; 
            } catch(e) {
                console.error('LINE Push Failed', e);
            } 
        }

        return res.status(200).json({ success: true, pushed });
    } catch (error) { 
        console.error('Submit report error', error);
        return res.status(500).json({ success: false, error: '系統處理失敗' }); 
    }
});

// ==========================================
// 產出結案 Excel API (五大工作表完整版)
// ==========================================
app.get('/api/projects/:projectId/export-excel', async (req, res) => {
    try {
        const projectId = req.params.projectId;
        const project = await findProjectById(projectId);
        
        if (!project) {
            return res.status(404).json({ success: false, message: '找不到此專案' });
        }
        
        const projectName = project.projectName;
        const projectBasePath = `工程專案管理/2026_工程專案/${projectName}`;

        const globalInventory = await readGlobalInventory();
        const customConfig = await readProjectMaterials(project.projectName);
        const inventoryMap = {};
        
        (globalInventory?.items || globalInventory?.materials || []).forEach(m => inventoryMap[m.materialId] = m);
        (customConfig?.items || customConfig?.materials || []).forEach(m => inventoryMap[m.materialId] = m);

        let transactionData = await readJsonFromOneDrive(`${projectBasePath}/project-material-transactions.json`, { transactions: [] }, false);
        const { stats, reports, dataQuality } = await generateProjectStats(project);

        const workbook = new ExcelJS.Workbook();
        workbook.creator = '工程專案自動化系統';
        const resolveMaterialCode = (item) => (item.materialId && inventoryMap[item.materialId]) ? inventoryMap[item.materialId].materialCode : (item.materialCode || '無編碼');

        // 第一張表：案場總表
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

        // 第二張表：材料結案總表
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

        // 第三張表：材料進出紀錄
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

        // 第四張表：日報明細
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

        // 第五張表：資料品質
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

const requiredVars = ['LINE_ACCESS_TOKEN', 'LINE_CHANNEL_SECRET', 'AZURE_CLIENT_ID', 'AZURE_TENANT_ID', 'AZURE_CLIENT_SECRET'];
if (requiredVars.some(v => !process.env[v])) {
    console.error('缺少必要環境變數');
    process.exit(1);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 伺服器運作中：http://localhost:${PORT}`));
