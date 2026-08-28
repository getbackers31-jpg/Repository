require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(cors());
app.use(express.json());

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
    const tokenRequest = { scopes: ['https://graph.microsoft.com/.default'] };
    try {
        const response = await cca.acquireTokenByClientCredential(tokenRequest);
        return Client.init({
            authProvider: (done) => { done(null, response.accessToken); }
        });
    } catch (error) {
        console.error("取得 Azure Token 失敗:", error);
        throw error;
    }
}

app.get('/', (req, res) => {
    res.send('✅ 云說工程：LINE 日報後端伺服器已成功啟動！');
});

// 這是前端送出圖文日報會呼叫的路由
app.post('/api/submit-report', upload.array('photos'), async (req, res) => {
    try {
        console.log("收到圖文日報提交請求！");
        const reportData = req.body; 
        const photos = req.files; 
        const graphClient = await getGraphClient();
        
        // 🌟 指定主管的微軟信箱 (已為您設定好)
        const TARGET_USER_EMAIL = "kate@cyber-cloud.info"; 

        // 設定日期與自動產生資料夾路徑
        const dateStr = new Date().toISOString().split('T')[0]; 
        const targetFolderPath = `工程專案管理/${reportData.projectName}_${dateStr}`;
        const fileName = `${reportData.workerName}_日報表.txt`;

        const fileContent = `提交人員：${reportData.workerName}\n工程名稱：${reportData.projectName}\n時間：${reportData.date}\n\n施工進度：\n${reportData.content}`;

        // 1. 寫入純文字日報
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`).put(fileContent);
        console.log("✅ 文字日報寫入成功");

        // 2. 將照片一張一張寫入同一個資料夾
        if (photos && photos.length > 0) {
            for (let file of photos) {
                // 處理照片檔名亂碼，並加上時間戳記
                const safeName = Buffer.from(file.originalname, 'latin1').toString('utf8');
                const photoName = `${Date.now()}_${safeName}`;
                await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${photoName}:/content`).put(file.buffer);
            }
            console.log(`✅ ${photos.length} 張照片上傳成功`);
        }

        // 3. 成功回傳
        return res.status(200).json({ success: true, message: '日報與照片已成功歸檔' });

    } catch (error) {
        console.error("❌ 上傳發生錯誤:", error);
        if (!res.headersSent) {
            return res.status(500).json({ success: false, error: error.message });
        }
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 伺服器運作中：http://localhost:${PORT}`);
});
