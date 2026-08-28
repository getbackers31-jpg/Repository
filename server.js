app.post('/api/submit-report', async (req, res) => {
    try {
        console.log("收到日報提交請求！");
        const reportData = req.body; 
        const graphClient = await getGraphClient();
        
        const TARGET_USER_EMAIL = "kate@cyber-cloud.info"; 
        const dateStr = new Date().toISOString().split('T')[0]; 
        const targetFolderPath = `工程專案管理/${reportData.projectName}_${dateStr}`;
        const fileName = `${reportData.workerName}_日報表.txt`;

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

        // 1. 寫入 OneDrive (核心功能，絕對不受影響)
        await graphClient.api(`/users/${TARGET_USER_EMAIL}/drive/root:/${targetFolderPath}/${fileName}:/content`).put(reportText);
        console.log("✅ 文字日報寫入 OneDrive 成功！");

        // 2. 🎯 LINE 群組推播（加上嚴格的 ID 格式檢查）
        let targetId = reportData.targetId;
        // LINE 群組 ID 通常是以 'C' 或 'R' 開頭的長字串
        if (targetId && (targetId.startsWith('C') || targetId.startsWith('R'))) {
            console.log(`🚀 準備將日報推播至 LINE 群組: ${targetId}`);
            
            const lineResponse = await fetch('https://api.line.me/v2/bot/message/push', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${LINE_ACCESS_TOKEN}`
                },
                body: JSON.stringify({
                    to: targetId,
                    messages: [{ type: 'text', text: reportText }]
                })
            });

            const lineResultText = await lineResponse.text();
            if (!lineResponse.ok) {
                console.error("❌ LINE 推播失敗詳細原因:", lineResultText);
            } else {
                console.log("🎉 LINE 機器人推播成功！");
            }
        } else {
            console.log(`⚠️ 略過 LINE 推播：收到的 ID (${targetId}) 不是群組 ID (需為 C 或 R 開頭)。但 OneDrive 存檔已順利完成！`);
        }

        return res.status(200).json({ success: true, message: '處理完成' });

    } catch (error) {
        console.error("❌ 系統處理錯誤:", error);
        if (!res.headersSent) {
            return res.status(500).json({ success: false, error: error.message });
        }
    }
});
