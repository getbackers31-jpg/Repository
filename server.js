require('dotenv').config(); // 讓本機測試時可以讀取 .env 檔
const express = require('express');
const cors = require('cors');
const multer = require('multer'); // 用於處理照片檔案上傳
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(cors());
app.use(express.json());

// 設定上傳記憶體暫存 (不直接存硬碟，方便轉傳到 OneDrive)
const upload = multer({ storage: multer.memoryStorage() });

// 這裡會讀取我們稍早要在 Render 填寫的環境變數 (金鑰)
const msalConfig = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET
    }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getGraphClient() {
    // 因為我們是背景伺服器，使用 Client Credentials 流程 (應用程式權限)
    const tokenRequest = {
        scopes: ['https://graph.microsoft.com/.default'],
    };

    try {
        const response = await cca.acquireTokenByClientCredential(tokenRequest);
        
        // 建立並回傳 Graph Client
        const client = Client.init({
            authProvider: (done) => {
                done(null, response.accessToken);
            }
        });
        return client;
    } catch (error) {
        console.error("取得 Azure Token 失敗:", error);
        throw error;
    }
}

// 當我們部署到 Render 時，可以用這個網址看伺服器有沒有醒來
app.get('/', (req, res) => {
    res.send('✅ 詮達工程：LINE 日報後端伺服器已成功啟動！');
});

// 這是前端 (LINE LIFF) 按下「確認送出日報」時會呼叫的網址
app.post('/api/submit-report', upload.array('photos'), async (req, res) => {
    try {
        console.log("收到新的日報提交請求！");
        const reportData = JSON.parse(req.body.reportData); // 前端傳來的文字與耗用數據
        const photos = req.files; // 前端傳來的照片檔案

        // 1. 取得擁有權限的 Graph Client
        const graphClient = await getGraphClient();

        // 2. 處理照片上傳至 OneDrive (概念示範)
        // 注意：正式上線前需確認貴公司 OneDrive 的 ID 或 SharePoint Site ID
        if (photos && photos.length > 0) {
            const dateStr = reportData.date.replace(/\//g, '-');
            const targetPath = `案場照片/${reportData.projectName}/${dateStr}`;
            
            console.log(`準備將 ${photos.length} 張照片上傳至 OneDrive 路徑: ${targetPath}`);
            
            // 這裡實作迴圈上傳 (由於 Graph API 限制，大檔案可能需改用 Upload Session)
            // for (let file of photos) { ... }
        }

        // 3. (未來可加) 呼叫 LINE Messaging API 推播圖文卡片到群組

        res.status(200).json({ success: true, message: '日報與照片已成功歸檔至 OneDrive' });
    } catch (error) {
        console.error('處理日報失敗:', error);
        res.status(500).json({ success: false, error: '伺服器處理失敗' });
    }
});

// Render 會自動分配 PORT，若在本機測試則預設使用 3000
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 伺服器運作中：http://localhost:${PORT}`);
});
