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
        return Client.init({ authProvider: (done) => { done(null, response.accessToken); } });
    } catch (error) { throw error; }
}

app.get('/', (req, res) => { res.send('✅ 伺服器運作中！'); });

app.post('/api/submit-report', upload.array('photos'), async (req, res) => {
    try {
        const reportData = req.body; 
        const photos = req.files; 
        const graphClient = await getGraphClient();
        
        // 主管信箱
        const TARGET_USER_EMAIL = "kate@cyber-cloud.info"; 

        // 只設定資料夾路徑，不再產生文字檔內容
        const dateStr = new Date().toISOString().split('T')[0]; 
        const targetFolderPath = `工程專案管理/${reportData.projectName}_${dateStr}`;

        // 🎯 這裡只留下上傳照片的迴圈，文字檔上傳已經移除了！
        if (photos && photos.length > 0) {
            for (let file of photos) {
                const safeName = Buffer.from(file.originalname, 'latin1').toString('utf8');
                const photoName = `${Date.now()}_${safeName}`;
                await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${photoName}:/content`).put(file.buffer);
            }
            console.log(`✅ ${photos.length} 張照片成功上傳至 ${targetFolderPath}`);
        }

        return res.status(200).json({ success: true, message: '照片已成功歸檔' });

    } catch (error) {
        console.error("❌ 上傳錯誤:", error);
        if (!res.headersSent) {
            return res.status(500).json({ success: false, error: error.message });
        }
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 伺服器運作中：http://localhost:${PORT}`); });
