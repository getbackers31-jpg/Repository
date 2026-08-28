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
        console.log("收到新的日報提交請求！");
        const reportData = req.body; 

        // 1. 取得擁有權限的 Graph Client
        const graphClient = await getGraphClient();

        // 2. 將前端傳來的資料排版成要存入文字檔的內容
        const fileContent = `
====================
詮達工程 - 施工日報表
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
        const targetFolderPath = "工程專案管理"; // 您可以在這裡改資料夾名稱
        
        // 🌟🌟🌟【請務必修改這裡】🌟🌟🌟
        // 伺服器需要知道存進哪一位使用者的 OneDrive
        // 請將下方引號內的中文，替換成您登入微軟 Azure / OneDrive 的那個 Email 信箱
        const TARGET_USER_EMAIL = "kate@cyber-cloud.info"; 

        if (TARGET_USER_EMAIL === "kate@cyber-cloud.info") {
            return res.status(400).json({ success: false, message: '後端尚未設定微軟信箱，請管理員修改 server.js' });
        }

        console.log(`準備寫入 OneDrive: 使用者 [${TARGET_USER_EMAIL}], 資料夾 [${targetFolderPath}], 檔名 [${fileName}]`);

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
