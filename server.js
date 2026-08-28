require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(cors());
app.use(express.json()); // 處理前端傳來的 JSON 資料

const upload = multer({ storage: multer.memoryStorage() });

const msalConfig = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET
    }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getGraphClient() {
    const tokenRequest = {
        scopes: ['https://graph.microsoft.com/.default'],
    };
    try {
        const response = await cca.acquireTokenByClientCredential(tokenRequest);
        return Client.init({
            authProvider: (done) => {
                done(null, response.accessToken);
            }
        });
    } catch (error) {
        console.error("取得 Azure Token 失敗:", error);
        throw error;
    }
}

app.get('/', (req, res) => {
    res.send('✅ 云說工程：LINE 日報後端伺服器已成功啟動！');
});

// 這是前端 (LINE LIFF) 送出日報會呼叫的路由
app.post('/api/submit-report', upload.array('photos'), async (req, res) => {
    try {
        const reportData = req.body; 
        const photos = req.files; 
        const graphClient = await getGraphClient();
        const TARGET_USER_EMAIL = "kate@cyber-cloud.info"; 

        const dateStr = new Date().toISOString().split('T')[0]; 
        // 自動建立「工程專案管理/案場名稱_日期」的專屬資料夾
        const targetFolderPath = `工程專案管理/${reportData.projectName}_${dateStr}`;
        const fileName = `${reportData.workerName}_日報表.txt`;

        const fileContent = `提交人員：${reportData.workerName}\n工程名稱：${reportData.projectName}\n時間：${reportData.date}\n\n施工進度：\n${reportData.content}`;

        // 1. 寫入文字檔
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`).put(fileContent);

        // 2. 將照片一張一張寫入同一個資料夾
        if (photos && photos.length > 0) {
            for (let file of photos) {
                // 為避免檔名重複，加上時間戳記
                const photoName = `${Date.now()}_${file.originalname}`;
                await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${photoName}:/content`).put(file.buffer);
            }
        }

        res.status(200).json({ success: true, message: '日報與照片已成功歸檔' });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});
====================
云說工程 - 施工日報表
====================
提交人員：${reportData.workerName || '未知師傅'}
工程名稱：${reportData.projectName || '未填寫'}
提交時間：${reportData.date || new Date().toLocaleString()}
--------------------
今日用料與施工進度：
${reportData.content || '無內容'}
        `.trim();

        // 3. 設定要寫入 OneDrive 的資料夾與檔名
        const dateStr = new Date().toISOString().split('T')[0]; 
        const fileName = `${reportData.projectName || '未命名案場'}_${reportData.workerName || '師傅'}_${dateStr}.txt`;
        const targetFolderPath = "工程專案管理"; // 您設定的主管資料夾名稱
        
        // 指定主管的微軟信箱
        const TARGET_USER_EMAIL = "otto@cyber-cloud.info"; 

        console.log(`準備寫入 OneDrive: 使用者 [${TARGET_USER_EMAIL}], 資料夾 [${targetFolderPath}], 檔名 [${fileName}]`);

        // 4. 呼叫 Graph API 實際寫入檔案到指定使用者的 OneDrive 根目錄中
        await graphClient
            .api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`)
            .put(fileContent);

        console.log("日報成功寫入 OneDrive！");
        res.status(200).json({ success: true, message: '日報已成功歸檔至 OneDrive' });

        // 4. 呼叫 Graph API 實際寫入檔案到指定使用者的 OneDrive 根目錄中
        await graphClient
            .api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`)
            .put(fileContent);

        console.log("日報成功寫入 OneDrive！");
        res.status(200).json({ success: true, message: '日報已成功歸檔至 OneDrive' });

    } catch (error) {
        console.error('處理日報失敗:', error);
        res.status(500).json({ success: false, error: '伺服器處理失敗: ' + error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 伺服器運作中：http://localhost:${PORT}`);
});
