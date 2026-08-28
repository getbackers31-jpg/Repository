require('dotenv').config();
const express = require('express');
const cors = require('cors');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(cors());
app.use(express.json()); 

// 🌟🌟🌟【請務必在這裡貼上您的 Channel access token (long-lived)】🌟🌟🌟
const LINE_ACCESS_TOKEN = "WJwUI8ZuAUkawqYSYRz+lmuZ2sHuAdCF6Ffe9l+oZwOXz4/ZQ0vCulcwQwE7LCeFjgjMwKHSK3CAproDQobNqH+ZIjQIgU7Sxzn2osK9JPZYFreCsoSNOz1L8E1l95C+WGmCWRZfQRO48kAJXhB88wdB04t89/1O/w1cDnyilFU="; 

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

app.get('/', (req, res) => { res.send('✅ 伺服器運作中 (機器人連動版)！'); });

app.post('/api/submit-report', async (req, res) => {
    try {
        console.log("收到日報提交請求！");
        const reportData = req.body; 
        const graphClient = await getGraphClient();
        
        const TARGET_USER_EMAIL = "kate@cyber-cloud.info"; 
        const dateStr = new Date().toISOString().split('T')[0]; 
        const targetFolderPath = `工程專案管理/${reportData.projectName}_${dateStr}`;
        const fileName = `${reportData.workerName}_日報表.txt`;

        // 1. 準備日報的純文字排版
        const dateDisplay = new Date().toLocaleDateString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit' });
        
        let reportText = `${dateDisplay}\n\n`;
        reportText += `溫度：${reportData.temp}度\n`;
        reportText += `濕度：${reportData.humidity}%\n`;
        reportText += `風速：${reportData.wind}m/s\n\n`;
        reportText += `施工廠商：${reportData.contractor}\n`;
        reportText += `施工人數：${reportData.workerCount}\n`;
        reportText += `_____________________\n\n`;
        reportText += `今日作業進度：\n${reportData.progress}\n\n`;
        reportText += `今日用料：\n${reportData.materials}\n\n`;
        reportText += `其他/備註：\n${reportData.remarks}\n`;
        reportText += `_____________________\n`;
        reportText += `以上為今日進度報告\n(由 ${reportData.workerName} 提交)`;

        // 2. 寫入 OneDrive
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`).put(reportText);
        console.log("✅ 文字日報寫入 OneDrive 成功！");

        // 3. 🎯 呼叫「云說工程小幫手」把日報推播到群組
        if (reportData.targetId && LINE_ACCESS_TOKEN !== "WJwUI8ZuAUkawqYSYRz+lmuZ2sHuAdCF6Ffe9l+oZwOXz4/ZQ0vCulcwQwE7LCeFjgjMwKHSK3CAproDQobNqH+ZIjQIgU7Sxzn2osK9JPZYFreCsoSNOz1L8E1l95C+WGmCWRZfQRO48kAJXhB88wdB04t89/1O/w1cDnyilFU=") {
            console.log(`準備將日報推播至 LINE 群組: ${reportData.targetId}`);
            
            const lineResponse = await fetch('https://api.line.me/v2/bot/message/push', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${LINE_ACCESS_TOKEN}`
                },
                body: JSON.stringify({
                    to: reportData.targetId,
                    messages: [{ type: 'text', text: reportText }]
                })
            });

            if (!lineResponse.ok) {
                const errorDetail = await lineResponse.text();
                console.error("❌ LINE 推播失敗:", errorDetail);
            } else {
                console.log("✅ LINE 機器人推播成功！");
            }
        }

        return res.status(200).json({ success: true, message: '日報已成功歸檔與推播' });

    } catch (error) {
        console.error("❌ 系統處理錯誤:", error);
        if (!res.headersSent) {
            return res.status(500).json({ success: false, error: error.message });
        }
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 伺服器運作中：http://localhost:${PORT}`); });
