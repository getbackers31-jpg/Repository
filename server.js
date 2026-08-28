require('dotenv').config();
const express = require('express');
const cors = require('cors');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(cors());
app.use(express.json()); // 專門解析純 JSON 資料

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

app.get('/', (req, res) => { res.send('✅ 伺服器運作中 (純文字高速版)！'); });

app.post('/api/submit-report', async (req, res) => {
    try {
        console.log("收到日報提交請求 (無照片)！");
        const reportData = req.body; 
        const graphClient = await getGraphClient();
        
        // 您的主管信箱
        const TARGET_USER_EMAIL = "kate@cyber-cloud.info"; 

        const dateStr = new Date().toISOString().split('T')[0]; 
        const targetFolderPath = `工程專案管理/${reportData.projectName}_${dateStr}`;
        const fileName = `${reportData.workerName}_日報表.txt`;

        // 整理文字檔內容 (包含所有氣象、用料與備註)
        const fileContent = `====================
詮達工程 - 施工日報表
====================
提交人員：${reportData.workerName}
案場名稱：${reportData.projectName}
提交時間：${reportData.date}

[氣象資訊]
溫度：${reportData.temp}度
濕度：${reportData.humidity}%
風速：${reportData.wind}m/s

[人力資訊]
施工廠商：${reportData.contractor}
施工人數：${reportData.workerCount}
--------------------
[今日作業進度]
${reportData.progress}

[今日用料]
${reportData.materials}

[其他/備註]
${reportData.remarks}
====================`;

        // 寫入純文字檔至 OneDrive
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`).put(fileContent);
        console.log("✅ 文字日報寫入成功！");

        return res.status(200).json({ success: true, message: '日報已成功歸檔' });

    } catch (error) {
        console.error("❌ 上傳錯誤:", error);
        if (!res.headersSent) {
            return res.status(500).json({ success: false, error: error.message });
        }
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 伺服器運作中：http://localhost:${PORT}`); });
