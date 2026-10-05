require('dotenv').config();
const express = require('express');
const cors = require('cors');
const crypto = require('crypto');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const ExcelJS = require('exceljs');

const APP_VERSION = '6.5.7';
const app = express();
app.use(cors());

const PORT = process.env.PORT || 3000;
const LINE_ACCESS_TOKEN = process.env.LINE_ACCESS_TOKEN;
const LINE_LOGIN_CHANNEL_ID = '2011289657';
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const LIFF_ID = process.env.LIFF_ID || '2011289657-vQgMb0eI';
const WAREHOUSE_LIFF_ID = '2011289657-dSXS9DVe';
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

let warehouseWriteQueue = Promise.resolve();
function withWarehouseWriteLock(task) {
    const current = warehouseWriteQueue.then(task, task);
    warehouseWriteQueue = current.catch(() => undefined);
    return current;
}

const WAREHOUSE_TRANSACTION_PATH = '工程專案管理/_系統設定/warehouse-transactions.json';

async function readWarehouseTransactions() {
    const data = await readJsonFromOneDrive(WAREHOUSE_TRANSACTION_PATH, { schemaVersion: 1, transactions: [] }, false);
    if (!data || !Array.isArray(data.transactions)) {
        throw new Error('倉庫異動檔格式不正確');
    }
    return data;
}

async function writeWarehouseTransactions(data) {
    await writeJsonToOneDrive(WAREHOUSE_TRANSACTION_PATH, data);
}

const MINIMUM_STOCK_THRESHOLD = 20;
function normalizeWarehousePolicyText(value) {
    return String(value || '').normalize('NFKC').toLowerCase().replace(/\s+/g, '');
}
function isNippon9011MinimumStockItem(item) {
    if (!item || typeof item !== 'object') return false;
    const text = normalizeWarehousePolicyText([item.materialName, item.materialCode, item.packageUnit, item.baseUnit].filter(Boolean).join(' '));
    const packageQuantity = Number(item.packageQuantity);
    return text.includes('立邦') && text.includes('90-11') &&
        (text.includes('加侖') || text.includes('gallon')) && [1, 5].includes(packageQuantity);
}
function applyWarehouseMinimumStockPolicy(item) {
    const enabled = isNippon9011MinimumStockItem(item);
    item.minimumStockEnabled = enabled;
    item.lowStockThreshold = enabled ? MINIMUM_STOCK_THRESHOLD : null;
    return item;
}
function calculateWarehouseStatus(item) {
    applyWarehouseMinimumStockPolicy(item);
    const stockQuantity = Number(item.stockQuantity);
    if (!Number.isFinite(stockQuantity) || stockQuantity < 0) return 'REVIEW_REQUIRED';
    if (stockQuantity === 0) return 'OUT_OF_STOCK';
    if (item.minimumStockEnabled === true && stockQuantity < MINIMUM_STOCK_THRESHOLD) return 'LOW_STOCK';
    return 'NORMAL';
}
function enrichWarehouseInventoryItem(item) {
    applyWarehouseMinimumStockPolicy(item);
    item.stockStatus = calculateWarehouseStatus(item);
    item.shortageQuantity = item.minimumStockEnabled === true
        ? Math.max(0, MINIMUM_STOCK_THRESHOLD - Number(item.stockQuantity || 0))
        : null;
    return item;
}

function findLatestWarehouseTransaction(transactions, materialId) {
    for (let index = transactions.length - 1; index >= 0; index--) {
        const transaction = transactions[index];
        if (transaction.materialId === materialId) return transaction;
    }
    return null;
}

async function reconcileWarehouseSnapshotIfLatest(inventoryData, txData, existingTx) {
    const latestTransaction = findLatestWarehouseTransaction(
        txData.transactions || [],
        existingTx.materialId
    );

    const item = (inventoryData.items || []).find(
        inventoryItem => inventoryItem.materialId === existingTx.materialId
    );

    const isLatestTransaction =
        latestTransaction?.submissionId === existingTx.submissionId;

    if (
        item &&
        isLatestTransaction &&
        Number(item.stockQuantity) !== Number(latestTransaction.afterQuantity)
    ) {
        item.stockQuantity = Number(latestTransaction.afterQuantity);
        const packageQuantity = Number(item.packageQuantity || 1);
        item.stockBaseQuantity = Number(
            (item.stockQuantity * packageQuantity).toFixed(4)
        );
        enrichWarehouseInventoryItem(item);
        inventoryData.updatedAt = new Date().toISOString();

        await writeJsonToOneDrive(
            '工程專案管理/_系統設定/inventory.json',
            inventoryData
        );

        configCache.globalInventory = {
            data: cloneJsonData(inventoryData),
            timestamp: Date.now()
        };

        return true;
    }

    return false;
}

function validateWarehouseDate(value, today) {
    const dateValue = String(value || today).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateValue)) {
        throw new Error('倉庫異動日期格式不正確');
    }
    const parsedDate = Date.parse(`${dateValue}T00:00:00+08:00`);
    if (!Number.isFinite(parsedDate)) {
        throw new Error('倉庫異動日期無效');
    }
    if (dateValue > today) {
        throw new Error('倉庫異動日期不可晚於今天');
    }
    return dateValue;
}

async function requireWarehouseAccess(req, res, next) {
    try {
        const authorization = String(req.get('authorization') || '');
        const idToken = authorization.startsWith('Bearer ') ? authorization.slice(7) : '';
        
        if (!idToken) {
            return res.status(401).json({ success: false, error: '缺少登入憑證' });
        }
        
        const response = await fetch('https://api.line.me/oauth2/v2.1/verify', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({ id_token: idToken, client_id: LINE_LOGIN_CHANNEL_ID })
        });
        
        const tokenData = await response.json();
        if (!response.ok || !tokenData.sub) {
            return res.status(401).json({ success: false, error: '登入憑證無效' });
        }
        
        const allowedUserIds = String(process.env.WAREHOUSE_ALLOWED_LINE_USER_IDS || '').split(',').map(v => v.trim()).filter(Boolean);
        if (!allowedUserIds.includes(tokenData.sub)) {
            return res.status(403).json({ success: false, error: '沒有庫存異動權限' });
        }
        
        req.warehouseUserId = tokenData.sub;
        next();
    } catch (error) {
        console.error('倉庫權限驗證失敗：', error);
        return res.status(500).json({ success: false, error: '無法驗證使用者權限' });
    }
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
    data.items = migrateWarehouseInventory(data).items.map(enrichWarehouseInventoryItem);
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

async function findProjectByName(projectName, includeInactive = false) {
    const config = await readProjectsFromOneDrive();
    const projects = Array.isArray(config.projects) ? config.projects : [];
    const normalizedName = normalizeProjectName(projectName);
    return projects.find(project =>
        (includeInactive || project.active !== false) &&
        normalizeProjectName(project.projectName) === normalizedName
    ) || null;
}
async function findProjectById(projectId, includeInactive = false) {
    const config = await readProjectsFromOneDrive();
    const projects = Array.isArray(config.projects) ? config.projects : [];
    const normalizedProjectId = String(projectId || '').trim();
    if (!normalizedProjectId) return null;
    return projects.find(project =>
        (includeInactive || project.active !== false) &&
        project.projectId === normalizedProjectId
    ) || null;
}function createProjectId() {
    return `PRJ-${crypto.randomUUID()}`;
}

async function registerProjectByName(projectName) {
    return withProjectWriteLock(async () => {
        const normalizedName = validateProjectName(projectName);
        const config = await readProjectsFromOneDrive();
        const projects = Array.isArray(config.projects) ? config.projects : [];
        const existingProject = projects.find(project => project.active !== false && normalizeProjectName(project.projectName) === normalizedName);

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

function formatTaiwanDateTime(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return '未紀錄';
    return new Intl.DateTimeFormat('zh-TW', {
        timeZone: 'Asia/Taipei',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
    }).format(date).replace(/\//g, '-');
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

function resolveMaterialSource(tx) {
    if (tx.materialSource === 'WAREHOUSE') {
        return '公司倉庫';
    }
    if (tx.materialSource === 'SUPPLIER_DIRECT') {
        return '供應商直送';
    }
    if (['WAREHOUSE_TRANSFER_IN', 'PROJECT_RETURN_OUT'].includes(tx.transactionType)) {
        return '公司倉庫';
    }
    if (['OPENING_ISSUE', 'ADDITIONAL_ISSUE'].includes(tx.transactionType)) {
        return '供應商直送／舊資料';
    }
    return '其他';
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



async function calculateProjectMaterialBalances(project) {
    const safeProjectName = sanitizePathSegment(project.projectName);
    const txPath = `工程專案管理/2026_工程專案/${safeProjectName}/project-material-transactions.json`;
    let txData = { transactions: [] };
    try { txData = await readJsonFromOneDrive(txPath, { transactions: [] }, false); } catch (error) {}
    if (!Array.isArray(txData.transactions)) txData.transactions = [];
    const issuedStats = {};
    for (const tx of txData.transactions) {
        const key = tx.materialId || `${tx.materialName} (${tx.baseUnit || ''})`;
        issuedStats[key] = (issuedStats[key] || 0) + Number(tx.baseQuantity || 0);
    }
    const statsResult = await generateProjectStats(project);
    const consumedStats = statsResult.stats?.materialStats || {};
    const inventoryMap = await buildInventoryMap(project);
    return Object.values(inventoryMap).map(material => {
        const key = material.materialId;
        const issued = Number(issuedStats[key] || 0);
        const consumed = Number(consumedStats[key] || 0);
        const remaining = Number((issued - consumed).toFixed(4));
        return {
            materialId: material.materialId,
            materialName: material.materialName,
            packageQuantity: Number(material.packageQuantity || 1),
            packageUnit: material.packageUnit || '',
            stockUnit: material.stockUnit,
            baseUnit: material.baseUnit || material.stockUnit,
            issuedBaseQuantity: issued,
            consumedBaseQuantity: consumed,
            remainingBaseQuantity: remaining,
            returnable: remaining > 0.0001
        };
    });
}

const WAREHOUSE_REPORT_ROOT='工程專案管理/倉庫管理/Excel報表';
const WAREHOUSE_REPORT_VERSION='6.5.7';
const WAREHOUSE_TYPE_LABELS={INITIAL_COUNT:'期初盤點',PURCHASE_IN:'採購入庫',WAREHOUSE_ADJUSTMENT:'盤點修正',PROJECT_TRANSFER_OUT:'領至案場',PROJECT_RETURN:'案場退回',SCRAP_DISPOSAL:'報廢處理'};
function reportTaiwanParts(){return Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hour12:false}).formatToParts(new Date()).filter(x=>x.type!=='literal').map(x=>[x.type,x.value]))}
function realDate(v){if(!/^\d{4}-\d{2}-\d{2}$/.test(String(v||'')))return false;const [y,m,d]=v.split('-').map(Number),x=new Date(Date.UTC(y,m-1,d));return x.getUTCFullYear()===y&&x.getUTCMonth()===m-1&&x.getUTCDate()===d}
function reportRange(range,start,end){const p=reportTaiwanParts(),today=`${p.year}-${p.month}-${p.day}`;if(!['all','month','custom','current'].includes(range))throw new Error('不支援的報表範圍');if(range==='all')return{range,label:'全部紀錄',startDate:null,endDate:today};if(range==='month')return{range,label:'本月',startDate:`${p.year}-${p.month}-01`,endDate:today};if(range==='current')return{range,label:'目前庫存',startDate:null,endDate:today};start=String(start||'');end=String(end||'');if(!realDate(start)||!realDate(end))throw new Error('自訂日期格式不正確');if(start>end)throw new Error('開始日期不可晚於結束日期');if(end>today)throw new Error('結束日期不可晚於今天');return{range,label:'自訂日期',startDate:start,endDate:end}}
function txInRange(list,r){if(r.range==='current')return[];return(list||[]).filter(t=>{const d=String(t.transactionDate||'').slice(0,10);return d&&(!r.startDate||d>=r.startDate)&&(!r.endDate||d<=r.endDate)&&t.writeStatus!=='FAILED'})}
function latestTxMap(list){const m=new Map();for(const t of list||[]){if(!t?.materialId||t.writeStatus==='FAILED')continue;const n=Date.parse(t.createdAt||t.transactionDate||0)||0,o=m.get(t.materialId);if(!o||n>=o._n)m.set(t.materialId,{...t,_n:n})}return m}
function twDateTime(v){const d=new Date(v);return Number.isNaN(d.getTime())?'':new Intl.DateTimeFormat('zh-TW',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hour12:false}).format(d)}
function twTime(v){const d=new Date(v);return Number.isNaN(d.getTime())?'':new Intl.DateTimeFormat('zh-TW',{timeZone:'Asia/Taipei',hour:'2-digit',minute:'2-digit',hour12:false}).format(d)}
function widths(ws,a){a.forEach((w,i)=>ws.getColumn(i+1).width=w)}
function setupSheet(ws,title,r,at,count){ws.mergeCells(1,1,1,count);const c=ws.getCell(1,1);c.value=title;c.font={bold:true,size:18,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF087F5B'}};c.alignment={vertical:'middle',horizontal:'left'};ws.getRow(1).height=34;const p=r.range==='all'?'第一筆異動至目前':r.range==='current'?'目前庫存快照':`${r.startDate} 至 ${r.endDate}`;ws.mergeCells(2,1,2,count);ws.getCell(2,1).value=`報表範圍：${r.label}｜${p}`;ws.mergeCells(3,1,3,count);ws.getCell(3,1).value=`匯出時間：${at}｜系統版本：v${WAREHOUSE_REPORT_VERSION}`;for(const n of[2,3]){ws.getRow(n).height=24;ws.getCell(n,1).font={size:11,color:{argb:'FF495057'}};ws.getCell(n,1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF1F3F5'}}}const h=ws.getRow(5);h.height=42;h.font={bold:true,size:12,color:{argb:'FFFFFFFF'}};h.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0CA678'}};h.alignment={vertical:'middle',horizontal:'center',wrapText:false};ws.views=[{state:'frozen',ySplit:5}];ws.autoFilter={from:{row:5,column:1},to:{row:5,column:count}}}
function bodyStyle(ws,noWrap=[]){ws.eachRow((row,n)=>{if(n<6)return;row.height=Math.max(row.height||0,28);row.font={size:11};row.alignment={vertical:'middle',wrapText:true};noWrap.forEach(i=>{row.getCell(i).alignment={vertical:'middle',horizontal:typeof row.getCell(i).value==='number'?'right':'center',wrapText:false}});row.eachCell(cell=>{if(typeof cell.value==='number')cell.numFmt=Number.isInteger(cell.value)?'#,##0':'#,##0.####'})})}
function emptyRow(ws,text){const r=ws.addRow([text]);r.height=30;r.font={italic:true,size:11,color:{argb:'FF6C757D'}}}
async function reportFolder(){const g=await getGraphClient(),p=reportTaiwanParts();await ensureChildFolder(g,'工程專案管理','倉庫管理');await ensureChildFolder(g,'工程專案管理/倉庫管理','Excel報表');await ensureChildFolder(g,WAREHOUSE_REPORT_ROOT,p.year);return(await ensureChildFolder(g,`${WAREHOUSE_REPORT_ROOT}/${p.year}`,p.month)).folderPath}
async function warehouseWorkbook(inv,all,r){const wb=new ExcelJS.Workbook();wb.creator='云說工程小幫手';wb.subject='倉庫庫存報表';wb.description=`系統版本 v${WAREHOUSE_REPORT_VERSION}`;const p=reportTaiwanParts(),at=`${p.year}/${p.month}/${p.day} ${p.hour}:${p.minute}`,filtered=txInRange(all,r).sort((a,b)=>(Date.parse(b.createdAt||b.transactionDate||0)||0)-(Date.parse(a.createdAt||a.transactionDate||0)||0)),latest=latestTxMap(all),items=(inv.items||[]).filter(x=>x.inventoryManaged!==false).map(x=>{const item=enrichWarehouseInventoryItem({...x});item.stockBaseQuantity=Number((Number(item.stockQuantity||0)*Number(item.packageQuantity||1)).toFixed(4));return item}),sl=s=>({NORMAL:'庫存正常',LOW_STOCK:'庫存不足',OUT_OF_STOCK:'缺貨',REVIEW_REQUIRED:'帳面異常'}[s]||s||'');
items.sort((a,b)=>({REVIEW_REQUIRED:0,OUT_OF_STOCK:1,LOW_STOCK:2,NORMAL:3}[a.stockStatus]??9)-({REVIEW_REQUIRED:0,OUT_OF_STOCK:1,LOW_STOCK:2,NORMAL:3}[b.stockStatus]??9)||String(a.materialCode||'').localeCompare(String(b.materialCode||''),'zh-Hant'));
let w=wb.addWorksheet('倉庫庫存總表');w.addRows([[],[],[],[],['材料編碼','材料名稱','目前庫存','庫存單位','包裝容量','包裝單位','換算後總量','基準單位','啟用最低庫存','最低庫存桶數','庫存狀態','尚缺桶數','最後異動日期','最後異動類型','最後操作人','最後異動備註']]);for(const x of items){const t=latest.get(x.materialId)||{};w.addRow([x.materialCode||'',x.materialName||'',+x.stockQuantity||0,x.stockUnit||'',+x.packageQuantity||1,x.packageUnit||'',+x.stockBaseQuantity||0,x.baseUnit||x.stockUnit||'',x.minimumStockEnabled?'是':'否',x.minimumStockEnabled?+x.lowStockThreshold||20:'',sl(x.stockStatus),x.minimumStockEnabled?+x.shortageQuantity||0:'',t.transactionDate||'',WAREHOUSE_TYPE_LABELS[t.transactionType]||t.transactionType||'',t.operatorName||'',t.remarks||''])}setupSheet(w,'云說工程小幫手｜倉庫庫存總表',r,at,16);widths(w,[20,36,14,12,14,14,17,13,17,17,15,14,16,18,16,40]);bodyStyle(w,[1,3,4,5,6,7,8,9,10,11,12,13,14,15]);
w=wb.addWorksheet('倉庫異動明細');w.addRows([[],[],[],[],['異動日期','異動時間','異動類型','材料編碼','材料名稱','數量變化','異動前數量','異動後數量','庫存單位','案場名稱','操作人','備註','建立時間']]);if(r.range==='current')emptyRow(w,'本次報表選擇「目前庫存」，未包含期間異動明細。');else if(!filtered.length)emptyRow(w,'本報表範圍內無倉庫異動紀錄。');else for(const t of filtered)w.addRow([t.transactionDate||'',twTime(t.createdAt),WAREHOUSE_TYPE_LABELS[t.transactionType]||t.transactionType||'',t.materialCode||'',t.materialName||'',+t.quantityChange||0,+t.beforeQuantity||0,+t.afterQuantity||0,t.stockUnit||'',t.projectName||'',t.operatorName||'',t.remarks||'',twDateTime(t.createdAt)]);setupSheet(w,'云說工程小幫手｜倉庫異動明細',r,at,13);widths(w,[14,12,18,21,38,14,15,15,12,28,16,40,22]);bodyStyle(w,[1,2,3,4,6,7,8,9,11,13]);const tr=filtered.filter(t=>['PROJECT_TRANSFER_OUT','PROJECT_RETURN'].includes(t.transactionType));w=wb.addWorksheet('案場轉撥紀錄');w.addRows([[],[],[],[],['異動日期','異動時間','異動方向','案場名稱','材料編碼','材料名稱','領出數量','退回數量','淨轉撥數量','庫存單位','換算總量','基準單位','操作人','備註']]);if(r.range==='current')emptyRow(w,'本次報表選擇「目前庫存」，未包含期間異動明細。');else if(!tr.length)emptyRow(w,'本報表範圍內無案場轉撥紀錄。');else for(const t of tr){const q=Math.abs(+t.quantityChange||0);w.addRow([t.transactionDate||'',twTime(t.createdAt),WAREHOUSE_TYPE_LABELS[t.transactionType],t.projectName||'',t.materialCode||'',t.materialName||'',t.transactionType==='PROJECT_TRANSFER_OUT'?q:0,t.transactionType==='PROJECT_RETURN'?q:0,+t.quantityChange||0,t.stockUnit||'',Math.abs(+t.baseQuantityChange||0),t.baseUnit||'',t.operatorName||'',t.remarks||''])}setupSheet(w,'云說工程小幫手｜案場轉撥紀錄',r,at,14);widths(w,[14,12,18,28,21,38,14,14,15,12,16,13,16,40]);bodyStyle(w,[1,2,3,5,7,8,9,10,11,12,13]);const pr=filtered.filter(t=>['INITIAL_COUNT','WAREHOUSE_ADJUSTMENT'].includes(t.transactionType));w=wb.addWorksheet('盤點差異');w.addRows([[],[],[],[],['盤點日期','盤點時間','盤點類型','材料編碼','材料名稱','盤點前帳面數量','實際盤點數量','盤點差異','庫存單位','包裝容量','換算差異量','基準單位','操作人','備註']]);if(r.range==='current')emptyRow(w,'本次報表選擇「目前庫存」，未包含期間異動明細。');else if(!pr.length)emptyRow(w,'本報表範圍內無盤點差異紀錄。');else for(const t of pr){const d=+t.quantityChange||0;w.addRow([t.transactionDate||'',twTime(t.createdAt),WAREHOUSE_TYPE_LABELS[t.transactionType],t.materialCode||'',t.materialName||'',+t.beforeQuantity||0,+t.afterQuantity||0,d,t.stockUnit||'',+t.packageQuantity||1,+t.baseQuantityChange||0,t.baseUnit||'',t.operatorName||'',t.remarks||''])}setupSheet(w,'云說工程小幫手｜盤點差異',r,at,14);widths(w,[14,12,18,21,38,18,18,14,12,14,17,13,16,40]);bodyStyle(w,[1,2,3,4,6,7,8,9,10,11,12,13]);const al=items.filter(x=>['LOW_STOCK','OUT_OF_STOCK','REVIEW_REQUIRED'].includes(x.stockStatus));w=wb.addWorksheet('庫存警示');w.addRows([[],[],[],[],['警示類型','材料編碼','材料名稱','目前庫存','庫存單位','最低庫存','尚缺數量','包裝容量','換算後總量','基準單位','最後異動日期','最後異動類型','最後操作人','建議處理']]);if(!al.length)emptyRow(w,'目前沒有庫存警示。');else for(const x of al){const t=latest.get(x.materialId)||{},ad=x.stockStatus==='LOW_STOCK'?'建議補貨至至少20桶':x.stockStatus==='OUT_OF_STOCK'?'請確認是否補貨':'請重新盤點並修正';w.addRow([sl(x.stockStatus),x.materialCode||'',x.materialName||'',+x.stockQuantity||0,x.stockUnit||'',x.minimumStockEnabled?+x.lowStockThreshold||20:'',x.minimumStockEnabled?+x.shortageQuantity||0:'',+x.packageQuantity||1,+x.stockBaseQuantity||0,x.baseUnit||x.stockUnit||'',t.transactionDate||'',WAREHOUSE_TYPE_LABELS[t.transactionType]||t.transactionType||'',t.operatorName||'',ad])}setupSheet(w,'云說工程小幫手｜庫存警示',r,at,14);widths(w,[16,21,38,14,12,14,14,14,17,13,16,18,16,34]);bodyStyle(w,[1,2,4,5,6,7,8,9,10,11,12,13]);return wb}

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

                if (text === '我的ID') {
                    const userId = event.source.userId;
                    await replyLineMessage(event.replyToken, `👤 您的專屬 LINE ID 是：\n\n${userId}\n\n👉 請長按複製上方代碼，並傳送給管理員以開通庫存修改權限。`);
                    continue;
                }

                if (text === '查庫存' || text === '庫存總表') {
                    try {
                        const inventoryData = await readGlobalInventory();
                        let totalCount = 0, normalCount = 0, lowCount = 0, outCount = 0, errCount = 0;

                        (inventoryData.items || []).forEach(item => {
                            if (!item.inventoryManaged) return;
                            totalCount++;
                            if (item.stockStatus === 'NORMAL') normalCount++;
                            else if (item.minimumStockEnabled === true && item.stockStatus === 'LOW_STOCK') lowCount++;
                            else if (item.stockStatus === 'OUT_OF_STOCK' || item.stockQuantity === 0) outCount++;
                            else errCount++;
                        });

                        const updateTime = formatTaiwanDateTime(inventoryData.updatedAt);
                        
                        let replyText = `📦【倉庫庫存摘要】\n` +
                                        `更新時間：${updateTime}\n\n` +
                                        `總共 ${totalCount} 項追蹤中物料：\n` +
                                        `🟢 正常庫存：${normalCount} 項\n` +
                                        `🟡 庫存不足：${lowCount} 項\n` +
                                        `⚪ 目前缺貨：${outCount} 項\n`;
                                        
                        if (errCount > 0) replyText += `🚨 帳面異常：${errCount} 項\n`;

                        replyText += `\n👇 點擊下方網址查看【完整庫存總表】\n` +
                                     `https://getbackers31-jpg.github.io/chuanda-frontend-/inventory.html`;

                        await replyLineMessage(event.replyToken, replyText);
                        continue;
                    } catch (error) {
                        console.error('讀取庫存失敗:', error);
                        await replyLineMessage(event.replyToken, '❌ 無法讀取庫存資料，請檢查系統連線。');
                        continue;
                    }
                }

                if (text.startsWith('查庫存 ')) {
                    try {
                        const keyword = text.replace('查庫存', '').trim().toLowerCase();
                        const inventoryData = await readGlobalInventory();
                        
                        const matchedItems = (inventoryData.items || []).filter(item => 
                            item.inventoryManaged && 
                            (item.materialName.toLowerCase().includes(keyword) || (item.materialCode && item.materialCode.toLowerCase().includes(keyword)))
                        );

                        if (matchedItems.length === 0) {
                            await replyLineMessage(event.replyToken, `❌ 找不到包含「${keyword}」的庫存品項。`);
                            continue;
                        }

                        if (matchedItems.length > 1) {
                             const options = matchedItems.map(item => `▪ ${item.materialName}`).join('\n');
                             await replyLineMessage(event.replyToken, `找到多筆符合「${keyword}」的品項，請輸入更完整的名稱：\n${options}`);
                             continue;
                        }

                        const target = matchedItems[0];
                        const statusMap = { 'NORMAL': '正常', 'LOW_STOCK': '⚠️ 庫存不足', 'REVIEW_REQUIRED': '🚨 異常', 'OUT_OF_STOCK': '❌ 缺貨' };
                        const updateTime = formatTaiwanDateTime(inventoryData.updatedAt);
                        
                        const detailText = `📦 ${target.materialName}\n\n` +
                                           `▪ 料號：${target.materialCode || '無'}\n` +
                                           `▪ 庫存：${target.stockQuantity} ${target.stockUnit}\n` +
                                           `▪ 換算：${target.stockBaseQuantity} ${target.baseUnit}\n` +
                                           `▪ 狀態：${statusMap[target.stockStatus] || target.stockStatus}\n` +
                                           (target.minimumStockEnabled === true ? `▪ 最低保有：${target.lowStockThreshold} ${target.stockUnit}\n▪ 尚缺：${target.shortageQuantity || 0} ${target.stockUnit}\n` : '') +
                                           `▪ 更新：${updateTime}`;

                        await replyLineMessage(event.replyToken, detailText);
                        continue;

                    } catch (error) {
                        await replyLineMessage(event.replyToken, '❌ 查詢失敗，請稍後再試。');
                        continue;
                    }
                }
                
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
                    await replyLineMessage(event.replyToken, '📖 「云說工程小幫手」指令：\n\n🔹 設定案場 案場名稱\n🔹 查詢案場\n🔹 查詢統計\n🔹 解除案場\n🔹 結案 案場名稱\n🔹 查庫存\n🔹 查庫存 關鍵字');
                }
                else if (text.startsWith('結案')) {
                    if (!targetId) continue;
                    const match = text.match(/^結案\s+(.+)$/);
                    if (!match) continue;
                    const bindings = await readBindingsFromOneDrive();
                    const binding = (bindings.bindings || []).find(item => item.groupId === targetId && item.active);
                    if (!binding || normalizeProjectName(binding.projectName) !== normalizeProjectName(match[1])) {
                        await replyLineMessage(event.replyToken, '⚠️ 名稱不符或無綁定');
                        continue;
                    }
                    await withProjectWriteLock(async () => {
                        const config = await readProjectsFromOneDrive();
                        const project = (config.projects || []).find(item => item.projectId === binding.projectId);
                        if (!project) {
                            await replyLineMessage(event.replyToken, '⚠️ 系統找不到此案場資料。');
                            return;
                        }
                        let leftoverSummary = '';
                        try {
                            const balances = await calculateProjectMaterialBalances(project);
                            const leftovers = balances.filter(item => item.returnable).map(item => {
                                const quantity = Number((item.remainingBaseQuantity / item.packageQuantity).toFixed(4));
                                return `• ${item.materialName}：${quantity} ${item.stockUnit}`;
                            });
                            if (leftovers.length) {
                                leftoverSummary = `\n\n⚠️ 案場尚有剩餘材料：\n${leftovers.join('\n')}\n\n請至倉庫系統執行「案場退回」。`;
                            }
                        } catch (error) {
                            console.error('計算結案剩料失敗：', error);
                            leftoverSummary = '\n\n⚠️ 剩料摘要計算失敗，請至倉庫系統確認案場餘量。';
                        }
                        try {
                            const response = await fetch(`http://localhost:${PORT}/api/projects/${binding.projectId}/export-excel`);
                            if (!response.ok) throw new Error(`HTTP ${response.status} ${await response.text()}`);
                            const contentType = response.headers.get('content-type') || '';
                            if (!contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) throw new Error('結案 API 未回傳 Excel 檔案');
                            const buffer = Buffer.from(await response.arrayBuffer());
                            const { dateStr } = getTaiwanDateParts();
                            const graphClient = await getGraphClient();
                            await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/工程專案管理/2026_工程專案/${sanitizePathSegment(project.projectName)}/結案總表_${sanitizePathSegment(project.projectName)}_${dateStr.replace(/-/g, '')}.xlsx:/content`).put(buffer);
                        } catch (error) {
                            console.error('結案失敗：', error);
                            await replyLineMessage(event.replyToken, `⚠️ 結案報表產生異常 (${error.message})\n案場尚未結案。`);
                            return;
                        }
                        const now = new Date().toISOString();
                        project.active = false;
                        project.status = 'CLOSED';
                        project.closedAt = now;
                        project.closedBy = event.source?.userId || '';
                        project.updatedAt = now;
                        config.updatedAt = now;
                        await writeProjectsToOneDrive(config);
                        await withBindingWriteLock(async () => {
                            const latest = await readBindingsFromOneDrive();
                            for (const item of latest.bindings || []) {
                                if (item.projectId === binding.projectId && item.active !== false) {
                                    item.active = false;
                                    item.closedAt = now;
                                    item.closeReason = 'PROJECT_CLOSED';
                                }
                            }
                            latest.updatedAt = now;
                            await writeBindingsToOneDrive(latest);
                        });
                        await replyLineMessage(event.replyToken, `✅ 案場「${project.projectName}」已成功結案！\n\n系統已產生 Excel 結案報表並保留案場歷史資料。${leftoverSummary}`);
                    });
                }            }
        } catch (e) { console.error('Webhook Error', e); }
    }
});


// ===== v6.5 材料品項管理核心 =====
const MATERIAL_CODE_STATUS={FORMAL:'FORMAL',TEMPORARY:'TEMPORARY'};
const SCRAP_REASONS=new Set(['材料過期','材料變質','材料結塊','包裝破損','受潮或污染','無法繼續施工使用','盤點確認報廢','其他']);
const DISABLE_REASONS=new Set(['品牌不再使用','停止採購','改用其他品牌','材料規格淘汰','主管決定停用','其他']);
function cleanText(v){return String(v==null?'':v).trim()}
function round4(v){return Number(Number(v).toFixed(4))}
function materialUid(){return `MAT-${crypto.randomUUID().toUpperCase()}`}
function materialCode(v){return cleanText(v).toUpperCase().replace(/\s+/g,'')}
function validateMaterialCode(v){const c=materialCode(v);if(!c)throw new Error('英文代碼不可空白');if(!/^[A-Z0-9-]+$/.test(c))throw new Error('英文代碼只能包含英文字母、數字與連字號');return c}
function migrateWarehouseItem(x){x.materialUid=cleanText(x.materialUid)||materialUid();x.materialCode=cleanText(x.materialCode)||cleanText(x.materialId);x.materialId=cleanText(x.materialId)||x.materialCode;x.codeStatus=x.codeStatus==='TEMPORARY'?'TEMPORARY':'FORMAL';x.active=x.active!==false;x.materialCategories=Array.isArray(x.materialCategories)?[...new Set(x.materialCategories.map(cleanText).filter(Boolean))]:[];x.previousCodes=Array.isArray(x.previousCodes)?x.previousCodes:[];return x}
function migrateWarehouseInventory(data){data.items=(data.items||[]).map(migrateWarehouseItem);return data}
function allMaterialCodes(items,exceptUid=''){const set=new Set();for(const x of items||[]){if(exceptUid&&x.materialUid===exceptUid)continue;[x.materialId,x.materialCode,...(x.previousCodes||[]).map(p=>p.code)].map(materialCode).filter(Boolean).forEach(c=>set.add(c))}return set}
function nextTemporaryMaterialCode(items){const d=getTaiwanDateParts().dateStr.replace(/-/g,''),prefix=`TMP-${d}-`;let max=0;for(const c of allMaterialCodes(items)){if(c.startsWith(prefix)){const n=Number(c.slice(prefix.length));if(Number.isInteger(n))max=Math.max(max,n)}}return `${prefix}${String(max+1).padStart(3,'0')}`}
function findWarehouseItem(data,key){const k=cleanText(key);return(data.items||[]).find(x=>x.materialUid===k||x.materialId===k||x.materialCode===k)}
function validateCategories(v){const a=[...new Set((Array.isArray(v)?v:[]).map(cleanText).filter(Boolean))];if(!a.length)throw new Error('至少選擇一個材料標籤');return a}
function publicWarehouseItem(x){return enrichWarehouseInventoryItem(migrateWarehouseItem({...x}))}
function safeV65Message(e){const m=String(e?.message||'');const allowed=['英文代碼不可空白','英文代碼只能包含英文字母、數字與連字號','英文代碼已存在','中文材料名稱不可空白','至少選擇一個材料標籤','庫存單位不可空白','基準單位不可空白','每一庫存單位的容量必須大於0','找不到該材料主檔','操作人不可空白','修改原因不可空白','更正原因不可空白','報廢數量必須大於0','報廢原因不正確','選擇其他報廢原因時，補充說明必填','停用原因不正確','選擇其他停用原因時，補充說明必填','目前庫存為0，請直接使用停用品項','目前仍有庫存，請先報廢材料或使用全部報廢並停用','重新啟用原因不可空白','已停用品項不可執行報廢'];return allowed.includes(m)||m.startsWith('報廢數量不可超過目前庫存')?m:'材料品項處理失敗'}
async function saveWarehouseInventory(data){data.updatedAt=new Date().toISOString();await writeJsonToOneDrive('工程專案管理/_系統設定/inventory.json',data);configCache.globalInventory={data:cloneJsonData(data),timestamp:Date.now()}}

app.use('/api', express.json());

// 倉庫專屬全域庫存 API
app.get('/api/warehouse/inventory', async (req, res) => {
    try {
        const inventoryData = await readGlobalInventory();
        const items = (inventoryData.items || []).filter(item => item.inventoryManaged !== false).map(publicWarehouseItem);
        return res.status(200).json({ 
            success: true, 
            updatedAt: inventoryData.updatedAt || null,
            items 
        });
    } catch (error) {
        console.error('取得倉庫庫存失敗：', error);
        return res.status(500).json({ success: false, error: '無法取得倉庫庫存' });
    }
});


// v6.3 最近五筆倉庫異動摘要
app.get('/api/warehouse/recent-transactions', requireWarehouseAccess, async (req, res) => {
    try {
        const requestedLimit = Number.parseInt(req.query.limit, 10);
        const limit = Number.isFinite(requestedLimit) ? Math.min(Math.max(requestedLimit, 1), 20) : 5;
        const txData = await readWarehouseTransactions();
        const completed = (txData.transactions || [])
            .filter(tx => tx && tx.writeStatus !== 'FAILED')
            .sort((a, b) => {
                const aTime = Date.parse(a.createdAt || a.transactionDate || 0) || 0;
                const bTime = Date.parse(b.createdAt || b.transactionDate || 0) || 0;
                return bTime - aTime;
            });
        return res.status(200).json({
            success: true,
            version: APP_VERSION,
            total: completed.length,
            items: completed.slice(0, limit).map(tx => ({
                transactionId: tx.transactionId || null,
                transactionDate: tx.transactionDate || null,
                createdAt: tx.createdAt || null,
                transactionType: tx.transactionType || null,
                materialId: tx.materialId || null,
                materialCode: tx.materialCode || null,
                materialName: tx.materialName || '未命名材料',
                quantityChange: Number(tx.quantityChange || 0),
                beforeQuantity: Number(tx.beforeQuantity || 0),
                afterQuantity: Number(tx.afterQuantity || 0),
                stockUnit: tx.stockUnit || '',
                projectId: tx.projectId || null,
                projectName: tx.projectName || null,
                operatorName: tx.operatorName || '',
                remarks: tx.remarks || ''
            }))
        });
    } catch (error) {
        console.error('取得最近倉庫異動失敗：', error);
        return res.status(500).json({ success: false, error: '無法取得最近異動' });
    }
});


// v6.4.1 倉庫 Excel 報表，產生後直接儲存至 OneDrive
app.post('/api/warehouse/export-excel',requireWarehouseAccess,async(req,res)=>{try{const p=req.body||{},r=reportRange(String(p.range||'month'),p.startDate,p.endDate),[inv,tx]=await Promise.all([readGlobalInventory(),readWarehouseTransactions()]),wb=await warehouseWorkbook(inv,tx.transactions||[],r),buf=Buffer.from(await wb.xlsx.writeBuffer()),folder=await reportFolder(),d=reportTaiwanParts(),stamp=`${d.year}${d.month}${d.day}_${d.hour}${d.minute}`;let part=r.label;if(r.range==='custom')part+=`_${r.startDate.replace(/-/g,'')}-${r.endDate.replace(/-/g,'')}`;const fileName=`倉庫庫存報表_${part}_${stamp}_v${WAREHOUSE_REPORT_VERSION}.xlsx`,g=await getGraphClient();await g.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${folder}/${fileName}:/content`).put(buf);return res.json({success:true,version:WAREHOUSE_REPORT_VERSION,fileName,folderPath:folder,range:r})}catch(e){console.error('產生倉庫 Excel 報表失敗：',e);const m=String(e?.message||''),valid=['不支援的報表範圍','自訂日期格式不正確','開始日期不可晚於結束日期','結束日期不可晚於今天'];if(valid.includes(m))return res.status(400).json({success:false,error:m});return res.status(500).json({success:false,error:'無法產生或儲存倉庫 Excel 報表'})}});

// ➕ 採購入庫 / 🔄 盤點修正
app.post('/api/warehouse/transactions', requireWarehouseAccess, async (req, res) => {
    try {
        const payload = req.body;
        const submissionId = String(payload.submissionId || '').trim();
        if (!submissionId) {
            return res.status(400).json({ success: false, error: '缺少 submissionId' });
        }
        
        const materialId = String(payload.materialId || '').trim();
        if (!materialId) {
            return res.status(400).json({ success: false, error: '缺少材料 ID' });
        }
        
        const transactionType = payload.transactionType;
        if (!['PURCHASE_IN', 'WAREHOUSE_ADJUSTMENT'].includes(transactionType)) {
            return res.status(400).json({ success: false, error: '不支援的異動類型' });
        }

        const operatorName = String(payload.operatorName || '').trim();
        if (!operatorName) {
            return res.status(400).json({ success: false, error: '缺少操作人姓名' });
        }

        await withWarehouseWriteLock(async () => {
            const inventoryData = await readGlobalInventory();
            const txData = await readWarehouseTransactions();
            
            const existingTx = txData.transactions.find(tx => tx.submissionId === submissionId);
            if (existingTx) {
                const reconciled = await reconcileWarehouseSnapshotIfLatest(
                    inventoryData,
                    txData,
                    existingTx
                );
                return res.status(200).json({
                    success: true,
                    duplicate: true,
                    reconciled,
                    message: '此筆庫存異動先前已完成'
                });
            }
            
            const itemIndex = (inventoryData.items || []).findIndex(i => i.materialId === materialId);
            if (itemIndex === -1) {
                return res.status(404).json({ success: false, error: '找不到該材料主檔' });
            }
            const item = inventoryData.items[itemIndex];
            if (item.inventoryManaged === false) {
                return res.status(400).json({ success: false, error: '該材料不納入庫存計算' });
            }
            if (item.active === false) return res.status(400).json({success:false,error:'該品項已停用'});
            
            const beforeQuantity = Number(item.stockQuantity || 0);
            let quantityChange = 0;
            let afterQuantity = 0;
            
            if (transactionType === 'WAREHOUSE_ADJUSTMENT') {
                const adjustmentReason=cleanText(payload.adjustmentReason);
                if(!adjustmentReason)return res.status(400).json({success:false,error:'盤點原因必填'});
                if(adjustmentReason==='其他'&&!cleanText(payload.remarks))return res.status(400).json({success:false,error:'選擇其他盤點原因時，補充說明必填'});
            }
            if (transactionType === 'PURCHASE_IN') {
                const qty = Number(payload.quantity);
                if (!Number.isFinite(qty) || qty <= 0) return res.status(400).json({ success: false, error: '入庫數量無效' });
                quantityChange = qty;
                afterQuantity = beforeQuantity + quantityChange;
            } else if (transactionType === 'WAREHOUSE_ADJUSTMENT') {
                const actualQty = Number(payload.actualQuantity);
                if (!Number.isFinite(actualQty) || actualQty < 0) return res.status(400).json({ success: false, error: '實際盤點數量無效' });
                afterQuantity = actualQty;
                quantityChange = afterQuantity - beforeQuantity;
            }
            
            if (quantityChange === 0 && transactionType === 'WAREHOUSE_ADJUSTMENT') {
                return res.status(400).json({ success: false, error: '盤點數量與目前庫存相同，無須調整' });
            }

            const packageQuantity = Number(item.packageQuantity || 1);
            const baseQuantityChange = quantityChange * packageQuantity;
            
            const { dateStr } = getTaiwanDateParts();
            const transactionDate = validateWarehouseDate(payload.transactionDate, dateStr);
            const transactionId = `TX-${crypto.randomUUID()}`;
            const nowIso = new Date().toISOString();
            
            const newTx = {
                transactionId,
                submissionId,
                transferId: null,
                transactionType,
                transactionDate,
                materialId: item.materialId,
                materialCode: item.materialCode,
                materialName: item.materialName,
                quantityChange,
                beforeQuantity,
                afterQuantity,
                stockUnit: item.stockUnit,
                packageQuantity,
                packageUnit: item.packageUnit,
                baseQuantityChange,
                baseUnit: item.baseUnit || item.stockUnit,
                projectId: null,
                projectName: null,
                operatorName,
                remarks: String(payload.remarks || ''),
                adjustmentReason: transactionType==='WAREHOUSE_ADJUSTMENT'?cleanText(payload.adjustmentReason):null,
                createdAt: nowIso,
                writeStatus: "COMPLETED"
            };
            
            txData.transactions.push(newTx);
            txData.updatedAt = nowIso;
            
            item.stockQuantity = afterQuantity;
            item.stockBaseQuantity = afterQuantity * packageQuantity;
            enrichWarehouseInventoryItem(item);
            inventoryData.updatedAt = nowIso;
            
            await writeWarehouseTransactions(txData);
            await writeJsonToOneDrive('工程專案管理/_系統設定/inventory.json', inventoryData);
            configCache.globalInventory = { data: cloneJsonData(inventoryData), timestamp: Date.now() };
            
            res.status(200).json({ 
                success: true, 
                transactionId,
                afterQuantity,
                message: '庫存異動成功'
            });
        });

    } catch (error) {
        console.error('庫存異動失敗:', error);
        const message = String(error?.message || '');
        if ([
            '倉庫異動日期格式不正確',
            '倉庫異動日期無效',
            '倉庫異動日期不可晚於今天'
        ].includes(message)) {
            return res.status(400).json({ success: false, error: message });
        }
        return res.status(500).json({ success: false, error: '系統錯誤，無法更新庫存' });
    }
});

// 🚚 領至案場 (帶有 materialSource: 'WAREHOUSE')
app.post('/api/warehouse/project-transfer', requireWarehouseAccess, async (req, res) => {
    try {
        const payload = req.body;
        const submissionId = String(payload.submissionId || '').trim();
        if (!submissionId) return res.status(400).json({ success: false, error: '缺少 submissionId' });

        const projectId = String(payload.projectId || '').trim();
        const materialId = String(payload.materialId || '').trim();
        const quantity = Number(payload.quantity);
        const operatorName = String(payload.operatorName || '').trim();

        if (!projectId || !materialId || !operatorName || !Number.isFinite(quantity) || quantity <= 0) {
            return res.status(400).json({ success: false, error: '參數不完整或數量無效' });
        }

        const project = await findProjectById(projectId);
        if (!project) return res.status(404).json({ success: false, error: '找不到指定案場' });

        await withWarehouseWriteLock(async () => {
            await withMaterialWriteLock(project.projectId, async () => {
                const inventoryData = await readGlobalInventory();
                const txData = await readWarehouseTransactions();

                const existingTx = txData.transactions.find(tx => tx.submissionId === submissionId);
                if (existingTx) {
                    const reconciled = await reconcileWarehouseSnapshotIfLatest(
                        inventoryData,
                        txData,
                        existingTx
                    );
                    return res.status(200).json({
                        success: true,
                        duplicate: true,
                        reconciled,
                        message: '此筆領料先前已完成'
                    });
                }

                const itemIndex = (inventoryData.items || []).findIndex(i => i.materialId === materialId);
                if (itemIndex === -1) return res.status(404).json({ success: false, error: '找不到該材料主檔' });

                const item = inventoryData.items[itemIndex];
                const beforeQuantity = Number(item.stockQuantity || 0);

                if (beforeQuantity < quantity) {
                    return res.status(400).json({ success: false, error: `倉庫庫存不足，目前只有 ${beforeQuantity} ${item.stockUnit}` });
                }

                const packageQuantity = Number(item.packageQuantity || 1);
                const afterQuantity = beforeQuantity - quantity;
                const baseQuantityChange = -(quantity * packageQuantity);

                const { dateStr } = getTaiwanDateParts();
                const transactionDate = validateWarehouseDate(payload.transactionDate, dateStr);
                const transactionId = `TX-${crypto.randomUUID()}`;
                const nowIso = new Date().toISOString();

                const safeProjectName = sanitizePathSegment(project.projectName);
                const projectTxPath = `工程專案管理/2026_工程專案/${safeProjectName}/project-material-transactions.json`;

                let projTxData = { transactions: [] };
                try {
                    projTxData = await readJsonFromOneDrive(projectTxPath, { transactions: [] }, false);
                } catch(e) {}
                if (!Array.isArray(projTxData.transactions)) projTxData.transactions = [];

                const existingProjectTransaction = projTxData.transactions.find(tx => tx.submissionId === submissionId);
                const transferId = existingProjectTransaction ? existingProjectTransaction.transferId : `TRF-${crypto.randomUUID()}`;

                if (!existingProjectTransaction) {
                    projTxData.transactions.push({
                        submissionId,
                        transferId,
                        materialSource: 'WAREHOUSE',
                        transactionDate,
                        transactionType: 'WAREHOUSE_TRANSFER_IN',
                        materialId: item.materialId,
                        materialCode: item.materialCode,
                        materialName: item.materialName,
                        quantity: quantity,
                        stockUnit: item.stockUnit,
                        packageQuantity: packageQuantity,
                        packageUnit: item.packageUnit,
                        baseQuantity: quantity * packageQuantity,
                        baseUnit: item.baseUnit || item.stockUnit,
                        remarks: String(payload.remarks || '倉庫轉入')
                    });
                    await writeJsonToOneDrive(projectTxPath, projTxData);
                }

                const warehouseTx = {
                    transactionId,
                    submissionId,
                    transferId,
                    transactionType: 'PROJECT_TRANSFER_OUT',
                    transactionDate,
                    materialId: item.materialId,
                    materialCode: item.materialCode,
                    materialName: item.materialName,
                    quantityChange: -quantity,
                    beforeQuantity,
                    afterQuantity,
                    stockUnit: item.stockUnit,
                    packageQuantity,
                    packageUnit: item.packageUnit,
                    baseQuantityChange,
                    baseUnit: item.baseUnit || item.stockUnit,
                    projectId: project.projectId,
                    projectName: project.projectName,
                    operatorName,
                    remarks: String(payload.remarks || ''),
                    createdAt: nowIso,
                    writeStatus: "COMPLETED"
                };

                txData.transactions.push(warehouseTx);
                txData.updatedAt = nowIso;
                await writeWarehouseTransactions(txData);

                item.stockQuantity = afterQuantity;
                item.stockBaseQuantity = afterQuantity * packageQuantity;
                enrichWarehouseInventoryItem(item);
                inventoryData.updatedAt = nowIso;
                await writeJsonToOneDrive('工程專案管理/_系統設定/inventory.json', inventoryData);
                configCache.globalInventory = { data: cloneJsonData(inventoryData), timestamp: Date.now() };

                res.status(200).json({ success: true, transactionId, transferId, afterQuantity, message: '成功領至案場' });
            });
        });
    } catch (error) {
        console.error('領料失敗:', error);
        const message = String(error?.message || '');
        if ([
            '倉庫異動日期格式不正確',
            '倉庫異動日期無效',
            '倉庫異動日期不可晚於今天'
        ].includes(message)) {
            return res.status(400).json({ success: false, error: message });
        }
        return res.status(500).json({ success: false, error: '系統錯誤，無法完成領料' });
    }
});

// ↩️ 案場退回 (帶有 materialSource: 'WAREHOUSE')
app.post('/api/warehouse/project-return', requireWarehouseAccess, async (req, res) => {
    try {
        const payload = req.body || {};
        const submissionId = String(payload.submissionId || '').trim();
        if (!submissionId) {
            return res.status(400).json({ success: false, error: '缺少 submissionId' });
        }

        const projectId = String(payload.projectId || '').trim();
        const materialId = String(payload.materialId || '').trim();
        const requestedQuantity = Number(payload.quantity);
        const operatorName = String(payload.operatorName || '').trim();

        if (
            !projectId ||
            !materialId ||
            !operatorName ||
            !Number.isFinite(requestedQuantity) ||
            requestedQuantity <= 0
        ) {
            return res.status(400).json({ success: false, error: '參數不完整或數量無效' });
        }

        const project = await findProjectById(projectId, true);
        if (!project) {
            return res.status(404).json({ success: false, error: '找不到指定案場' });
        }

        await withWarehouseWriteLock(async () => {
            await withMaterialWriteLock(project.projectId, async () => {
                const inventoryData = await readGlobalInventory();
                const txData = await readWarehouseTransactions();

                const existingWarehouseTx = txData.transactions.find(
                    tx => tx.submissionId === submissionId
                );

                if (existingWarehouseTx) {
                    const reconciled = await reconcileWarehouseSnapshotIfLatest(
                        inventoryData,
                        txData,
                        existingWarehouseTx
                    );
                    return res.status(200).json({
                        success: true,
                        duplicate: true,
                        reconciled,
                        message: '此筆退料先前已完成'
                    });
                }

                const itemIndex = (inventoryData.items || []).findIndex(
                    item => item.materialId === materialId
                );
                if (itemIndex === -1) {
                    return res.status(404).json({ success: false, error: '找不到該材料主檔' });
                }

                const item = inventoryData.items[itemIndex];
                if (item.inventoryManaged === false) {
                    return res.status(400).json({ success: false, error: '該材料不納入庫存計算' });
                }

                const safeProjectName = sanitizePathSegment(project.projectName);
                const projectTxPath =
                    `工程專案管理/2026_工程專案/${safeProjectName}/project-material-transactions.json`;

                let projTxData = { transactions: [] };
                try {
                    projTxData = await readJsonFromOneDrive(
                        projectTxPath,
                        { transactions: [] },
                        false
                    );
                } catch (error) {}
                if (!Array.isArray(projTxData.transactions)) projTxData.transactions = [];

                const existingProjectTransaction = projTxData.transactions.find(
                    tx => tx.submissionId === submissionId
                );

                const packageQuantity = Number(item.packageQuantity || 1);
                let quantity = requestedQuantity;
                let baseQuantityToReturn = quantity * packageQuantity;
                let transferId;
                let transactionDate;

                const { dateStr } = getTaiwanDateParts();

                if (existingProjectTransaction) {
                    if (existingProjectTransaction.transactionType !== 'PROJECT_RETURN_OUT') {
                        return res.status(409).json({
                            success: false,
                            error: '相同 submissionId 已被其他案場交易使用'
                        });
                    }
                    if (existingProjectTransaction.materialId !== materialId) {
                        return res.status(409).json({
                            success: false,
                            error: '重送資料的材料與原交易不一致'
                        });
                    }

                    quantity = Math.abs(Number(existingProjectTransaction.quantity || 0));
                    baseQuantityToReturn = Math.abs(
                        Number(existingProjectTransaction.baseQuantity || 0)
                    );
                    transferId = existingProjectTransaction.transferId;
                    transactionDate = existingProjectTransaction.transactionDate;
                } else {
                    const balances = await calculateProjectMaterialBalances(project);
                    const balanceItem = balances.find(entry => entry.materialId === materialId);
                    const currentBalance = Number(balanceItem?.remainingBaseQuantity || 0);

                    if (currentBalance < baseQuantityToReturn) {
                        return res.status(400).json({
                            success: false,
                            error:
                                `退料超過案場餘額！案場帳面僅剩 ` +
                                `${currentBalance / packageQuantity} ${item.stockUnit}`
                        });
                    }

                    transferId = `TRF-${crypto.randomUUID()}`;
                    transactionDate = validateWarehouseDate(
                        payload.transactionDate,
                        dateStr
                    );

                    projTxData.transactions.push({
                        submissionId,
                        transferId,
                        materialSource: 'WAREHOUSE',
                        transactionDate,
                        transactionType: 'PROJECT_RETURN_OUT',
                        materialId: item.materialId,
                        materialCode: item.materialCode,
                        materialName: item.materialName,
                        quantity: -quantity,
                        stockUnit: item.stockUnit,
                        packageQuantity,
                        packageUnit: item.packageUnit,
                        baseQuantity: -baseQuantityToReturn,
                        baseUnit: item.baseUnit || item.stockUnit,
                        remarks: String(payload.remarks || '案場退回倉庫')
                    });

                    await writeJsonToOneDrive(projectTxPath, projTxData);
                }

                const beforeQuantity = Number(item.stockQuantity || 0);
                const afterQuantity = beforeQuantity + quantity;
                const transactionId = `TX-${crypto.randomUUID()}`;
                const nowIso = new Date().toISOString();

                const warehouseTx = {
                    transactionId,
                    submissionId,
                    transferId,
                    transactionType: 'PROJECT_RETURN',
                    transactionDate,
                    materialId: item.materialId,
                    materialCode: item.materialCode,
                    materialName: item.materialName,
                    quantityChange: quantity,
                    beforeQuantity,
                    afterQuantity,
                    stockUnit: item.stockUnit,
                    packageQuantity,
                    packageUnit: item.packageUnit,
                    baseQuantityChange: baseQuantityToReturn,
                    baseUnit: item.baseUnit || item.stockUnit,
                    projectId: project.projectId,
                    projectName: project.projectName,
                    operatorName,
                    remarks: String(payload.remarks || ''),
                    createdAt: nowIso,
                    writeStatus: 'COMPLETED'
                };

                txData.transactions.push(warehouseTx);
                txData.updatedAt = nowIso;
                await writeWarehouseTransactions(txData);

                item.stockQuantity = afterQuantity;
                item.stockBaseQuantity = Number(
                    (afterQuantity * packageQuantity).toFixed(4)
                );
                enrichWarehouseInventoryItem(item);
                inventoryData.updatedAt = nowIso;

                await writeJsonToOneDrive(
                    '工程專案管理/_系統設定/inventory.json',
                    inventoryData
                );
                configCache.globalInventory = {
                    data: cloneJsonData(inventoryData),
                    timestamp: Date.now()
                };

                return res.status(200).json({
                    success: true,
                    transactionId,
                    transferId,
                    afterQuantity,
                    resumed: Boolean(existingProjectTransaction),
                    message: existingProjectTransaction
                        ? '已接續完成先前未完成的退料'
                        : '成功退回倉庫'
                });
            });
        });
    } catch (error) {
        console.error('退料失敗:', error);
        const message = String(error?.message || '');
        if (['倉庫異動日期格式不正確','倉庫異動日期無效','倉庫異動日期不可晚於今天'].includes(message)) {
            return res.status(400).json({ success: false, error: message });
        }
        return res.status(500).json({ success: false, error: '系統錯誤，無法完成退料' });
    }
});


// v6.5 品項清單
app.get('/api/warehouse/material-items',requireWarehouseAccess,async(req,res)=>{try{const d=await readGlobalInventory();return res.json({success:true,version:APP_VERSION,items:(d.items||[]).map(publicWarehouseItem)})}catch(e){console.error(e);return res.status(500).json({success:false,error:'無法取得材料品項'})}});
app.post('/api/warehouse/material-items',requireWarehouseAccess,async(req,res)=>{try{await withWarehouseWriteLock(async()=>{const d=await readGlobalInventory(),p=req.body||{},status=p.codeStatus==='TEMPORARY'?'TEMPORARY':'FORMAL';let code=status==='TEMPORARY'?nextTemporaryMaterialCode(d.items):validateMaterialCode(p.materialCode);if(allMaterialCodes(d.items).has(code))throw new Error('英文代碼已存在');const name=cleanText(p.materialName),stockUnit=cleanText(p.stockUnit),baseUnit=cleanText(p.baseUnit),capacity=Number(p.packageQuantity);if(!name)throw new Error('中文材料名稱不可空白');if(!stockUnit)throw new Error('庫存單位不可空白');if(!baseUnit)throw new Error('基準單位不可空白');if(!(capacity>0))throw new Error('每一庫存單位的容量必須大於0');const now=new Date().toISOString(),item={materialUid:materialUid(),materialId:code,materialCode:code,materialName:name,codeStatus:status,active:true,materialCategories:validateCategories(p.materialCategories),previousCodes:[],stockQuantity:0,stockBaseQuantity:0,stockUnit,packageQuantity:round4(capacity),packageUnit:cleanText(p.packageUnit)||baseUnit,baseUnit,inventoryManaged:true,minimumStockEnabled:false,createdAt:now,createdBy:cleanText(p.operatorName),createReason:cleanText(p.createReason),remarks:cleanText(p.remarks)};if(!item.createdBy)throw new Error('操作人不可空白');d.items.push(item);await saveWarehouseInventory(d);return res.status(201).json({success:true,item:publicWarehouseItem(item)})})}catch(e){return res.status(400).json({success:false,error:safeV65Message(e)})}});
app.patch('/api/warehouse/material-items/:uid',requireWarehouseAccess,async(req,res)=>{try{await withWarehouseWriteLock(async()=>{const d=await readGlobalInventory(),item=findWarehouseItem(d,req.params.uid);if(!item)throw new Error('找不到該材料主檔');const p=req.body||{},action=cleanText(p.action),operator=cleanText(p.operatorName);if(!operator)throw new Error('操作人不可空白');if(action==='RENAME'){if(!cleanText(p.materialName))throw new Error('中文材料名稱不可空白');if(!cleanText(p.reason))throw new Error('修改原因不可空白');item.materialName=cleanText(p.materialName);item.nameUpdatedAt=new Date().toISOString();item.nameUpdatedBy=operator;item.nameUpdateReason=cleanText(p.reason)}else if(action==='UPDATE_CATEGORIES'){item.materialCategories=validateCategories(p.materialCategories);item.categoriesUpdatedAt=new Date().toISOString();item.categoriesUpdatedBy=operator}else if(action==='RECODE'||action==='FORMALIZE'){const code=validateMaterialCode(p.materialCode);if(allMaterialCodes(d.items,item.materialUid).has(code))throw new Error('英文代碼已存在');const reason=cleanText(p.reason);if(!reason)throw new Error('更正原因不可空白');const old=materialCode(item.materialCode||item.materialId);if(old!==code)item.previousCodes.push({code:old,changedAt:new Date().toISOString(),changedBy:operator,reason});item.materialId=code;item.materialCode=code;item.codeStatus='FORMAL';if(cleanText(p.materialName))item.materialName=cleanText(p.materialName)}else if(action==='DISABLE'){if(Number(item.stockQuantity||0)!==0)throw new Error('目前仍有庫存，請先報廢材料或使用全部報廢並停用');const reason=cleanText(p.disableReason);if(!DISABLE_REASONS.has(reason))throw new Error('停用原因不正確');if(reason==='其他'&&!cleanText(p.remarks))throw new Error('選擇其他停用原因時，補充說明必填');item.active=false;item.disabledAt=new Date().toISOString();item.disabledBy=operator;item.disabledReason=reason}else if(action==='ENABLE'){if(!cleanText(p.reason))throw new Error('重新啟用原因不可空白');item.active=true;item.reenabledAt=new Date().toISOString();item.reenabledBy=operator;item.reenabledReason=cleanText(p.reason)}else throw new Error('不支援的品項操作');await saveWarehouseInventory(d);return res.json({success:true,item:publicWarehouseItem(item)})})}catch(e){return res.status(400).json({success:false,error:safeV65Message(e)})}});
app.post('/api/warehouse/material-items/:uid/scrap',requireWarehouseAccess,async(req,res)=>{try{await withWarehouseWriteLock(async()=>{const d=await readGlobalInventory(),t=await readWarehouseTransactions(),item=findWarehouseItem(d,req.params.uid),p=req.body||{};if(!item)throw new Error('找不到該材料主檔');if(item.active===false)throw new Error('已停用品項不可執行報廢');if((t.transactions||[]).some(x=>x.submissionId===cleanText(p.submissionId)))return res.json({success:true,duplicate:true});const qty=Number(p.quantity),before=Number(item.stockQuantity||0);if(!(qty>0))throw new Error('報廢數量必須大於0');if(qty>before)throw new Error(`報廢數量不可超過目前庫存${before}${item.stockUnit||''}`);if(!SCRAP_REASONS.has(cleanText(p.scrapReason)))throw new Error('報廢原因不正確');if(cleanText(p.scrapReason)==='其他'&&!cleanText(p.remarks))throw new Error('選擇其他報廢原因時，補充說明必填');if(!cleanText(p.operatorName))throw new Error('操作人不可空白');const after=round4(before-qty),now=new Date().toISOString(),tx={transactionId:`TX-${crypto.randomUUID()}`,submissionId:cleanText(p.submissionId),transferId:null,transactionType:'SCRAP_DISPOSAL',transactionDate:validateWarehouseDate(p.transactionDate,getTaiwanDateParts().dateStr),materialUid:item.materialUid,materialId:item.materialId,materialCode:item.materialCode,materialName:item.materialName,quantityChange:-qty,beforeQuantity:before,afterQuantity:after,stockUnit:item.stockUnit,packageQuantity:Number(item.packageQuantity||1),packageUnit:item.packageUnit,baseQuantityChange:round4(-qty*Number(item.packageQuantity||1)),baseUnit:item.baseUnit||item.stockUnit,projectId:null,projectName:null,scrapReason:cleanText(p.scrapReason),operatorName:cleanText(p.operatorName),remarks:cleanText(p.remarks),createdAt:now,writeStatus:'COMPLETED'};item.stockQuantity=after;item.stockBaseQuantity=round4(after*Number(item.packageQuantity||1));if(p.disableAfter===true){if(after!==0)throw new Error('全部報廢並停用必須報廢目前全部庫存');const dr=cleanText(p.disableReason);if(!DISABLE_REASONS.has(dr))throw new Error('停用原因不正確');if(dr==='其他'&&!cleanText(p.disableRemarks||p.remarks))throw new Error('選擇其他停用原因時，補充說明必填');item.active=false;item.disabledAt=now;item.disabledBy=cleanText(p.operatorName);item.disabledReason=dr}enrichWarehouseInventoryItem(item);t.transactions.push(tx);t.updatedAt=now;await writeWarehouseTransactions(t);await saveWarehouseInventory(d);return res.json({success:true,transactionId:tx.transactionId,item:publicWarehouseItem(item)})})}catch(e){return res.status(400).json({success:false,error:safeV65Message(e)})}});

// 其他專案與材料相關 API 路由
app.get('/api/projects', async (req, res) => {
    try {
        const includeInactive = String(req.query.includeInactive || '').toLowerCase() === 'true';
        const returnableOnly = String(req.query.returnableOnly || '').toLowerCase() === 'true';
        const config = await readProjectsFromOneDrive();
        const source = (config.projects || []).filter(project => includeInactive || project.active !== false);
        const projects = [];
        for (const project of source) {
            let hasReturnableMaterials = false;
            let returnableItemCount = 0;
            if (includeInactive || returnableOnly) {
                try {
                    const balances = await calculateProjectMaterialBalances(project);
                    returnableItemCount = balances.filter(item => item.returnable).length;
                    hasReturnableMaterials = returnableItemCount > 0;
                } catch (error) {
                    console.error(`計算案場 ${project.projectName} 餘量失敗：`, error);
                }
            }
            if (returnableOnly && !hasReturnableMaterials) continue;
            projects.push({
                projectId: project.projectId,
                projectName: project.projectName,
                active: project.active !== false,
                status: project.status || (project.active === false ? 'CLOSED' : 'ACTIVE'),
                hasReturnableMaterials,
                returnableItemCount
            });
        }
        projects.sort((a, b) => (Number(b.active) - Number(a.active)) || a.projectName.localeCompare(b.projectName, 'zh-Hant'));
        return res.status(200).json({ success: true, projects });
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

app.get('/api/projects/:projectId/material-balances', async (req, res) => {
    try {
        const project = await findProjectById(req.params.projectId, true);
        if (!project) return res.status(404).json({ success: false, error: '找不到指定案場' });
        const balances = await calculateProjectMaterialBalances(project);
        return res.status(200).json({ success: true, projectId: project.projectId, active: project.active !== false, balances });
    } catch (error) {
        console.error('取得案場材料餘額失敗：', error);
        return res.status(500).json({ success: false, error: '取得餘額失敗' });
    }
});
// Excel 結案報表匯出 (包含完整五張工作表與材料來源欄位)
app.get('/api/projects/:projectId/export-excel', async (req, res) => {
    try {
        const projectId = req.params.projectId;
        const project = await findProjectById(projectId, true);
        
        if (!project) {
            return res.status(404).json({ success: false, message: '找不到此專案' });
        }
        
        const projectName = project.projectName;
        const projectBasePath = `工程專案管理/2026_工程專案/${projectName}`;

        const inventoryMap = await buildInventoryMap(project);

        let transactionData = await readJsonFromOneDrive(`${projectBasePath}/project-material-transactions.json`, { transactions: [] }, false);
        const statsResult = await generateProjectStats(project);
        const reports = Array.isArray(statsResult.reports) ? statsResult.reports : [];
        const stats = statsResult.stats || {
            totalDays: 0,
            workDays: 0,
            noWorkDays: 0,
            totalManDays: 0,
            contractorStats: {},
            materialStats: {},
            materialDetails: {},
            reporterStats: {}
        };
        const dataQuality = statsResult.dataQuality || {
            sourceFileCount: 0,
            parsedFileCount: 0,
            invalidFileCount: 0,
            effectiveReportCount: 0,
            supersededReportCount: 0
        };

        const workbook = new ExcelJS.Workbook();
        workbook.creator = '工程專案自動化系統';
        const resolveMaterialCode = (item) => (item.materialId && inventoryMap[item.materialId]) ? inventoryMap[item.materialId].materialCode : (item.materialCode || '無編碼');

        // 工作表 1：案場總表
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

        // 工作表 2：材料結案總表
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

        // 工作表 3：材料進出紀錄 (含材料來源)
        const wsTxLog = workbook.addWorksheet('材料進出紀錄');
        wsTxLog.views = [{ showGridLines: true }];
        wsTxLog.addRow(['日期', '異動類型', '材料來源', '材料分類編碼', '材料名稱', '包裝規格', '原始數量', '庫存單位', '換算後數量', '基準單位', '備註']);
        
        const typeMap = { 
            'OPENING_ISSUE': '開工首批進場', 
            'ADDITIONAL_ISSUE': '追加進場',
            'WAREHOUSE_TRANSFER_IN': '倉庫領料轉入',
            'PROJECT_RETURN_OUT': '退料回大倉庫'
        };
        
        (transactionData.transactions || []).forEach(tx => {
            wsTxLog.addRow([
                tx.transactionDate, 
                typeMap[tx.transactionType] || tx.transactionType, 
                resolveMaterialSource(tx),
                resolveMaterialCode(tx), 
                tx.materialName, 
                `${tx.packageQuantity||1}${tx.packageUnit||''}/${tx.stockUnit}`, 
                tx.quantity, tx.stockUnit, tx.baseQuantity, tx.baseUnit, tx.remarks || ''
            ]);
        });
        
        reports.forEach(r => {
            (r.materialItems || []).forEach(item => {
                wsTxLog.addRow([
                    r.reportDate, '施工耗用', '-', 
                    resolveMaterialCode(item), 
                    item.materialName, 
                    `${item.packageQuantity||1}${item.packageUnit||''}/${item.stockUnit}`, 
                    item.quantity, item.stockUnit, item.baseQuantity, item.baseUnit, item.remarks || '日報自動記錄'
                ]);
            });
        });

        // 工作表 4：日報明細
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

        // 工作表 5：資料品質
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

// 📋 提交日報與小幫手材料進場 (含第三層防呆與完整一般日報流程)
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
        
        const { dateStr, timeStr } = getTaiwanDateParts();
        let reportDate = dateStr;
        if (formType === 'material_issue') {
            try { reportDate = validateIssueDate(reportData.date, dateStr); }
            catch (dateError) { return res.status(400).json({ success: false, error: dateError.message }); }
        }
        const submitDate = reportDate;

        const inventoryMap = await buildInventoryMap(project);

        if (formType === 'material_issue') {
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

            const isSameNumber = (a, b) => Math.abs(Number(a) - Number(b)) < 0.0001;
            const forceDirectIssue = reportData.forceDirectIssue === true;

            let isDuplicateSubmission = false;
            let suspiciousMatches = [];

            await withMaterialWriteLock(project.projectId, async () => {
                const txPath = `${projectFolderPath}/project-material-transactions.json`;
                let txData = { transactions: [] };
                try {
                    txData = await readJsonFromOneDrive(txPath, { transactions: [] }, false);
                } catch (e) {}
                if (!Array.isArray(txData.transactions)) txData.transactions = [];

                isDuplicateSubmission = txData.transactions.some(tx => tx.submissionId === submissionId);
                if (isDuplicateSubmission) return;

                // 第三層防呆檢查：同日、同材料、同數量且來源為 WAREHOUSE (相容舊資料)
                if (!forceDirectIssue) {
                    for (const material of materialItems) {
                        const matchedTransaction = txData.transactions.find(tx =>
                            tx.transactionType === 'WAREHOUSE_TRANSFER_IN' &&
                            (!tx.materialSource || tx.materialSource === 'WAREHOUSE') &&
                            tx.transactionDate === submitDate &&
                            tx.materialId === material.materialId &&
                            isSameNumber(tx.quantity, material.quantity)
                        );

                        if (matchedTransaction) {
                            suspiciousMatches.push({
                                materialId: material.materialId,
                                materialName: material.materialName,
                                quantity: material.quantity,
                                stockUnit: material.stockUnit,
                                existingTransferId: matchedTransaction.transferId
                            });
                        }
                    }
                }

                if (suspiciousMatches.length > 0 && !forceDirectIssue) {
                    return; 
                }

                const issueType = reportData.issueType === 'ADDITIONAL' ? 'ADDITIONAL_ISSUE' : 'OPENING_ISSUE';
                
                materialItems.forEach(m => {
                    txData.transactions.push({
                        submissionId,
                        transferId: null,
                        materialSource: 'SUPPLIER_DIRECT', // 👈 標記為供應商直送
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

            if (suspiciousMatches.length > 0 && !forceDirectIssue) {
                return res.status(409).json({
                    success: false,
                    requiresConfirmation: true,
                    warningCode: 'POSSIBLE_DUPLICATE_RECEIPT',
                    error: `今天已有相同材料由公司倉庫領入 ${suspiciousMatches[0].quantity} ${suspiciousMatches[0].stockUnit}，請確認這次是否為另一批供應商直送材料。`,
                    matches: suspiciousMatches
                });
            }

            if (isDuplicateSubmission) {
                return res.status(200).json({ success: true, duplicate: true, pushed: false, message: '此筆材料進場先前已完成歸檔，未重複入帳' });
            }

            const reporterNameStr = reportData.reporterName ? String(reportData.reporterName).trim() : '未紀錄';
            const issueTypeLabel = reportData.issueType === 'ADDITIONAL' ? '追加進場' : '開工首批進場';
            
            let msg = `📦 案場材料進場\n\n` +
                      `來源：供應商直接送達\n` +
                      `日期：${submitDate.replace(/-/g, '/')}\n` +
                      `案場：${project.projectName}\n` +
                      `填表：${reporterNameStr}\n` +
                      `類型：${issueTypeLabel}\n\n` +
                      `━━━━━━━━━━━━\n[進場明細]\n`;
            
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

            return res.status(200).json({ success: true, pushed: pushed, message: pushed ? '供應商直送材料紀錄已成功歸檔' : '材料已成功歸檔，但LINE群組發布失敗' });
        }

        // 完整的一般施工日報邏輯
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
            const contractorText = contractorItems.length > 0
                ? contractorItems
                    .map(item => `• ${item.contractorName}：${item.workerCount}人`)
                    .join('\n')
                : '無';

            const workItemText = workItems.length > 0
                ? workItems
                    .map(item => `• ${item}`)
                    .join('\n')
                : '無';

            const customWorkText = String(reportData.customWorkItem || '').trim();
            const workNotesText = String(reportData.workNotes || '').trim();

            const materialText = materialItems.length > 0
                ? materialItems
                    .map(item => `• ${item.materialName}：${item.quantity}${item.stockUnit}`)
                    .join('\n')
                : '今日無用料';

            reportText +=
                `溫度：${reportData.temp || '未紀錄'}度\n` +
                `濕度：${reportData.humidity || '未紀錄'}%\n` +
                `風速：${reportData.wind || '未紀錄'}m/s\n\n` +
                `施工廠商：\n${contractorText}\n\n` +
                `總出工人數：${calculatedTotalWorkerCount}人\n\n` +
                `━━━━━━━━━━━━\n\n` +
                `今日進度：\n${workItemText}\n`;

            if (customWorkText) {
                reportText += `• 其他：${customWorkText}\n`;
            }

            if (workNotesText) {
                reportText += `\n作業補充：\n${workNotesText}\n`;
            }

            reportText +=
                `\n今日用料：\n${materialText}\n\n` +
                `備註：\n${reportData.remarks || '無'}\n\n` +
                `━━━━━━━━━━━━\n` +
                `以上為今日進度報告`;
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
        
        return res.status(200).json({ success: true, pushed: pushed, message: pushed ? '日報已歸檔' : '日報已歸檔，但LINE群組發布失敗' });

    } catch (error) {
        console.error('提交錯誤：', error);
        return res.status(500).json({ success: false, error: error.message || '系統內部處理失敗' });
    }
});

const requiredVars = [
    'LINE_ACCESS_TOKEN', 
    'LINE_CHANNEL_SECRET', 
    'LINE_LOGIN_CHANNEL_ID',
    'WAREHOUSE_ALLOWED_LINE_USER_IDS',
    'AZURE_CLIENT_ID', 
    'AZURE_TENANT_ID', 
    'AZURE_CLIENT_SECRET'
];
const missingVars = requiredVars.filter(name => !process.env[name]);
if (missingVars.length > 0) {
    console.error('缺少必要環境變數：', missingVars.join(', '));
    process.exit(1);
}

if (require.main === module) {
    app.listen(PORT, () => console.log(`🚀 伺服器運作中：http://localhost:${PORT}`));
}
module.exports = app;
