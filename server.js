<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>云說工程 - 施工日報與材料進場表</title>
  <script charset="utf-8" src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>
  <style>
    body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; background:#f4f7f6; padding:15px; color:#333; margin:0; }
    .container { background:white; padding:20px; border-radius:10px; box-shadow:0 4px 8px rgba(0,0,0,.1); max-width:500px; margin:auto; position: relative; }
    h2 { text-align:center; color:#00B900; margin-top:0; font-size:22px; position: relative; }
    .version-tag { font-size: 10px; color: #ccc; position: absolute; right: 0; bottom: 0; font-weight: normal; }
    label { font-weight:bold; display:block; margin-top:15px; margin-bottom:5px; color:#444; }
    select, textarea, input[type="text"], input[type="number"], input[type="date"] { width:100%; padding:10px; border:1px solid #ccc; border-radius:5px; box-sizing:border-box; font-size:16px; background:#fafafa; margin-bottom:5px; }
    .flex-row { display:flex; gap:10px; margin-bottom:5px; }
    .flex-row > * { flex:1; min-width:0; }
    .btn-weather { background:#17a2b8; color:white; border:0; padding:9px; border-radius:5px; cursor:pointer; font-size:14px; width:100%; margin-bottom:10px; }
    .btn-weather:disabled, .btn-add:disabled, .submit-btn:disabled { background:#ccc; cursor:not-allowed; }
    .checkbox-grid { display:grid; grid-template-columns:1fr 1fr; gap:10px; background:#f9f9f9; padding:10px; border-radius:5px; border:1px solid #eee; }
    .checkbox-label { display:flex; align-items:center; font-weight:normal; margin:0; font-size:15px; }
    .checkbox-label input { margin-right:8px; transform:scale(1.2); }
    .material-group { background:#f9f9f9; padding:10px; border-radius:5px; border:1px solid #eee; margin-bottom:10px; }
    .material-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:5px; }
    .material-title { font-size:14px; font-weight:bold; color:#666; }
    .remove-material-button { color:#d9534f; border:0; background:none; font-weight:bold; cursor:pointer; }
    .btn-add { background:#6c757d; color:white; border:0; padding:9px; border-radius:5px; cursor:pointer; font-size:14px; width:100%; margin-bottom:10px; }
    .submit-btn { width:100%; background:#00B900; color:white; border:0; padding:12px; font-size:18px; border-radius:5px; margin-top:20px; cursor:pointer; font-weight:bold; }
    .section-divider { border-bottom:2px dashed #ddd; margin:20px 0; }
    .hidden { display:none !important; }
    .no-material-option { display:flex; align-items:center; font-weight:bold; color:#d9534f; margin-bottom:15px; background:#fdf5f5; padding:10px; border-radius:5px; border:1px solid #f5c6c6; }
    .no-material-option input { width:auto; margin-right:10px; transform:scale(1.3); }
    .field-help { margin-top: 4px; color: #777; font-size: 13px; line-height: 1.5; }
    .readonly-input { background-color: #e9ecef; color: #495057; cursor: not-allowed; }
    @media (max-width:380px) { .checkbox-grid { grid-template-columns:1fr; } .flex-row { flex-direction:column; gap:4px; } }
  </style>
</head>
<body>
  <div class="container">
    <h2 id="formTitle">👷 施工日報表 <span class="version-tag">v5.4</span></h2>

    <div style="display:flex; gap:10px; margin-bottom:15px;">
        <button type="button" id="btnDailyMode" class="btn btn-primary" style="flex:1; padding:10px; border:none; border-radius:5px; background:#00B900; color:white; font-weight:bold; cursor:pointer;" onclick="switchMode('daily_report')">📝 填寫施工日報</button>
        <button type="button" id="btnIssueMode" class="btn btn-outline-secondary" style="flex:1; padding:10px; border:1px solid #ccc; border-radius:5px; background:#e9ecef; color:#333; font-weight:bold; cursor:pointer;" onclick="switchMode('material_issue')">📦 登記材料進場</button>
    </div>

    <form id="dailyReportForm">
      <input type="hidden" id="formType" value="daily_report">

      <div id="issueExtraSection" class="hidden" style="background:#f8f9fa; padding:12px; border-radius:5px; border-left:4px solid #28a745; margin-bottom:15px;">
        <label for="issueType" style="margin-top:0;">🚚 進場類型 <span style="color:red">*</span></label>
        <select id="issueType" onchange="updateAllBalances(); saveDraft()" style="margin-bottom:10px;">
          <option value="" selected disabled>請選擇進場類型</option>
          <option value="OPENING">🏢 開工首批進場</option>
          <option value="ADDITIONAL">🚚 後續追加進場</option>
        </select>

        <label for="reportDate" style="margin-top:0;">📅 實際進場日期 (可往回補登) <span style="color:red">*</span></label>
        <input type="date" id="reportDate" onchange="saveDraft()">
      </div>
      
      <label for="reporterName">👤 填表人姓名</label>
      <input type="text" id="reporterName" placeholder="請輸入您的姓名 (系統將自動記憶)" oninput="saveDraft()" required>

      <div id="lockedProjectSection" class="hidden">
        <label for="lockedProjectName">📍 案場名稱</label>
        <input type="text" id="lockedProjectName" class="readonly-input" readonly>
      </div>

      <div id="compatibleProjectSection">
        <label for="projectName">📍 案場名稱</label>
        <select id="projectName" onchange="saveDraft()" required disabled>
          <option value="" disabled selected>正在載入案場清單...</option>
        </select>
      </div>

      <input type="hidden" id="projectId" value="">

      <div class="daily-only">
        <label class="no-material-option" style="margin-top:15px; background:#fff3f3; border-color:#ffcccc; color:#cc0000;">
          <input type="checkbox" id="noWorkToday" onchange="handleNoWorkChange(); saveDraft()">今日無出工 (停工/休息)
        </label>
      </div>

      <div id="noWorkReasonSection" class="hidden">
        <select id="noWorkReason" onchange="saveDraft()" style="border-color:#ffcccc; background:#fffcfc;">
          <option value="" disabled selected>請選擇無出工原因</option>
          <option value="天候因素 (下雨/颱風等)">天候因素 (下雨/颱風等)</option>
          <option value="例假日休息">例假日休息</option>
          <option value="現場無施工安排">現場無施工安排</option>
          <option value="停工/待料">停工/待料</option>
          <option value="其他">其他 (請於最下方備註說明)</option>
        </select>
        <div class="section-divider"></div>
      </div>

      <div id="contractorSection" class="daily-only">
        <label>🤝 施工廠商與人數</label>
        <div id="contractorsContainer"></div>
        <button type="button" class="btn-add" onclick="addContractorRow()" style="background:#5bc0de;">➕ 新增一組廠商</button>
        <div class="section-divider"></div>
      </div>

      <div class="daily-only">
        <label>🌤️ 氣象資訊</label>
        <button type="button" class="btn-weather" onclick="getWeather()">一鍵獲取當地天氣</button>
        <div class="flex-row">
          <input type="number" id="temp" placeholder="溫度(度)" step="0.1" oninput="saveDraft()">
          <input type="number" id="humidity" placeholder="濕度(%)" min="0" max="100" step="1" oninput="saveDraft()">
          <input type="number" id="wind" placeholder="風速(m/s)" min="0" step="0.1" oninput="saveDraft()">
        </div>
        <div class="section-divider"></div>
      </div>

      <div id="workAndMaterialSection" class="daily-only">
        <label>🛠️ 今日作業</label>
        <div class="checkbox-grid" id="workItemsGrid"></div>
        <input type="text" id="customWorkItem" class="hidden" placeholder="請輸入其他作業項目" maxlength="100" oninput="saveDraft()" style="margin-top:10px">
        <textarea id="workNotes" rows="2" placeholder="作業補充（選填）" oninput="saveDraft()" style="margin-top:10px"></textarea>
        <div class="section-divider"></div>
      </div>

      <div id="materialSectionContainer">
        <label id="materialLabel">📦 今日用料</label>
        <div id="noMaterialWrapper" class="daily-only">
          <label class="no-material-option">
            <input type="checkbox" id="noMaterial" onchange="handleNoMaterialChange(); saveDraft()">今日無用料
          </label>
        </div>
        <div id="materialsContainer"></div>
        <button type="button" id="addMaterialButton" class="btn-add" onclick="addMaterialRow()">新增一項用料</button>
        <div class="section-divider"></div>
      </div>

      <label for="remarks">📝 備註（選填）</label>
      <textarea id="remarks" rows="2" placeholder="天候影響、材料短缺、人員異動等狀況" oninput="saveDraft()"></textarea>

      <button type="button" id="submitButton" class="submit-btn" onclick="submitForm()">確認送出施工日報</button>
    </form>
  </div>

  <script>
    const LIFF_ID = '2011289657-vQgMb0eI';
    const API_BASE_URL = 'https://bot.cyber-cloud.info';
    
    const API_URL = `${API_BASE_URL}/api/submit-report`;
    const PROJECTS_API_URL = `${API_BASE_URL}/api/projects`;
    const MATERIALS_API_URL = `${API_BASE_URL}/api/materials`;
    const MATERIAL_BALANCES_API_URL = `${API_BASE_URL}/api/projects`;
    
    const MAX_MATERIAL_ROWS = 10;
    const MAX_CONTRACTOR_ROWS = 10;
    const CONTRACTORS_KEY = 'savedContractors';

    const workItemsList = [
      '材料進場', '現場放樣', '鷹架搭設', '鷹架拆除', '水刀除鏽',
      '手動工具除鏽', '底漆施工', '中塗漆施工', '面漆施工', '底漆修補',
      '中塗漆修補', '面漆修補', '底漆膜厚檢測', '中塗漆膜厚檢測',
      '面漆膜厚檢測', '缺失改善', '現場清潔', '等待驗收', '完成驗收', '其他'
    ];

    let MATERIAL_ITEMS_MAP = new Map(); 
    let MATERIAL_OPTIONS = []; 
    let projectMaterialBalances = {}; 

    let isSubmitting = false;
    let isRestoringDraft = false;
    let reportArchived = false;
    let projectMode = 'LOADING';
    let currentSubmissionId = '';

    function getTaiwanDateString() {
      return new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Asia/Taipei',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
      }).format(new Date());
    }

    function switchMode(mode) {
        document.getElementById('formType').value = mode;
        const btnDaily = document.getElementById('btnDailyMode');
        const btnIssue = document.getElementById('btnIssueMode');
        const dailyOnlyElements = document.querySelectorAll('.daily-only');
        const issueExtraSection = document.getElementById('issueExtraSection');
        const formTitle = document.getElementById('formTitle');
        const submitButton = document.getElementById('submitButton');
        const materialLabel = document.getElementById('materialLabel');

        if (mode === 'daily_report') {
            btnDaily.style.background = '#00B900';
            btnDaily.style.color = 'white';
            btnDaily.style.border = 'none';
            btnIssue.style.background = '#e9ecef';
            btnIssue.style.color = '#333';
            btnIssue.style.border = '1px solid #ccc';
            
            formTitle.innerHTML = '👷 施工日報表 <span class="version-tag">v5.4</span>';
            submitButton.textContent = '確認送出並發布至群組';
            submitButton.style.background = '#00B900';
            materialLabel.textContent = '📦 今日用料';

            issueExtraSection.classList.add('hidden');
            dailyOnlyElements.forEach(el => el.classList.remove('hidden'));
        } else {
            btnDaily.style.background = '#e9ecef';
            btnDaily.style.color = '#333';
            btnDaily.style.border = '1px solid #ccc';
            btnIssue.style.background = '#28a745';
            btnIssue.style.color = 'white';
            btnIssue.style.border = 'none';

            formTitle.innerHTML = '📦 材料進場登記表 <span class="version-tag">v5.4</span>';
            submitButton.textContent = '確認送出材料進場紀錄';
            submitButton.style.background = '#28a745';
            materialLabel.textContent = '📦 進場材料清單';

            issueExtraSection.classList.remove('hidden');
            dailyOnlyElements.forEach(el => el.classList.add('hidden'));
        }
        updateAllBalances();
        saveDraft();
    }

    function updateAllBalances() {
        document.querySelectorAll('#materialsContainer .material-group').forEach(group => {
            updateBalanceDisplay(group);
        });
    }

    function createSubmissionId() {
      if (window.crypto && typeof window.crypto.randomUUID === 'function') {
        return window.crypto.randomUUID();
      }
      return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2);
    }

    function ensureSubmissionId() {
      if (!currentSubmissionId) currentSubmissionId = createSubmissionId();
      return currentSubmissionId;
    }

    function getSelectedProjectName() {
      if (projectMode === 'LOCKED') return document.getElementById('lockedProjectName').value.trim();
      return document.getElementById('projectName').value.trim();
    }

    function getDraftKey() {
      const projectId = document.getElementById('projectId').value.trim();
      if (projectId) return `dailyReportDraft:${projectId}`;
      return 'dailyReportDraft';
    }

    function getProjectIdFromUrl() {
      const params = new URLSearchParams(window.location.search);
      return String(params.get('projectId') || '').trim();
    }

    async function fetchMaterials(projectId = '') {
      try {
        const url = projectId ? `${MATERIALS_API_URL}?projectId=${projectId}` : MATERIALS_API_URL;
        const res = await fetch(url);
        const data = await res.json();
        if (data.success && Array.isArray(data.materials) && data.materials.length > 0) {
            MATERIAL_OPTIONS = data.materials;
            MATERIAL_ITEMS_MAP.clear();
            MATERIAL_OPTIONS.forEach(m => {
                if (m.materialId) MATERIAL_ITEMS_MAP.set(m.materialId, m);
            });
        }
      } catch (err) { console.warn('材料清單讀取失敗', err); }
    }

    async function fetchMaterialBalances(projectId) {
      try {
        const url = `${MATERIAL_BALANCES_API_URL}/${encodeURIComponent(projectId)}/material-balances`;
        const res = await fetch(url);
        const data = await res.json();
        if (data.success && data.balances) {
            data.balances.forEach(b => projectMaterialBalances[b.materialId] = b);
        }
      } catch (err) { console.warn('餘額讀取失敗', err); }
    }

    window.addEventListener('load', async () => {
      initializeStaticForm();
      const todayStr = getTaiwanDateString();
      document.getElementById('reportDate').value = todayStr;

      try {
        await liff.init({ liffId: LIFF_ID });
        const profile = await liff.getProfile();
        if (profile) document.getElementById('reporterName').value = profile.displayName;
      } catch (error) { console.warn('LIFF 初始化跳過：', error); }

      const projectIdFromUrl = getProjectIdFromUrl();
      try {
        if (projectIdFromUrl) {
          await loadLockedProject(projectIdFromUrl);
          await fetchMaterials(projectIdFromUrl);
          await fetchMaterialBalances(projectIdFromUrl); 
        } else {
          await loadProjects();
          await fetchMaterials('');
        }
        loadDraft();
      } catch (error) {
        projectMode = 'ERROR';
        document.getElementById('submitButton').disabled = true;
        alert('案場載入失敗，請重新從施工群組開啟表單');
      }
    });

    async function loadLockedProject(projectId) {
      const response = await fetch(`${PROJECTS_API_URL}/${encodeURIComponent(projectId)}`);
      const responseText = await response.text();
      let result;
      try { result = responseText ? JSON.parse(responseText) : {}; } catch { throw new Error(`HTTP ${response.status}`); }
      if (!response.ok || result.success !== true || !result.project) throw new Error(result.error || '找不到指定案場');

      document.getElementById('projectId').value = result.project.projectId;
      document.getElementById('lockedProjectName').value = result.project.projectName;
      document.getElementById('compatibleProjectSection').classList.add('hidden');
      document.getElementById('lockedProjectSection').classList.remove('hidden');
      projectMode = 'LOCKED';
    }

    async function loadProjects() {
      const projectSelect = document.getElementById('projectName');
      const response = await fetch(PROJECTS_API_URL);
      const responseText = await response.text();
      let result;
      try { result = responseText ? JSON.parse(responseText) : {}; } catch { throw new Error(`HTTP ${response.status}`); }
      if (!response.ok || result.success !== true || !Array.isArray(result.projects)) throw new Error(result.error || '無法取得案場清單');

      projectSelect.innerHTML = '';
      const placeholder = new Option('請選擇正式案場', '');
      placeholder.disabled = true; placeholder.selected = true;
      projectSelect.add(placeholder);

      result.projects.forEach(project => projectSelect.add(new Option(project.projectName, project.projectName)));
      projectSelect.disabled = false;
      projectMode = 'COMPATIBLE';
    }

    function safeReadArray(key) {
      try {
        const value = JSON.parse(localStorage.getItem(key) || '[]');
        return Array.isArray(value) ? value : [];
      } catch { return []; }
    }

    function getSavedContractorsList() {
      let savedStr = localStorage.getItem(CONTRACTORS_KEY);
      if (!savedStr) {
        const defaultList = ['容煜', '詮達'];
        localStorage.setItem(CONTRACTORS_KEY, JSON.stringify(defaultList));
        return defaultList;
      }
      return safeReadArray(CONTRACTORS_KEY);
    }

    function initializeStaticForm() {
      const grid = document.getElementById('workItemsGrid');
      workItemsList.forEach(item => {
        const label = document.createElement('label');
        label.className = 'checkbox-label';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.name = 'workItem';
        checkbox.value = item;
        checkbox.addEventListener('change', handleWorkItemChange);
        label.append(checkbox, document.createTextNode(` ${item}`));
        grid.appendChild(label);
      });
    }

    function handleWorkItemChange() {
      if (isRestoringDraft) return;
      const otherChecked = [...document.querySelectorAll('input[name="workItem"]:checked')].some(cb => cb.value === '其他');
      const input = document.getElementById('customWorkItem');
      input.classList.toggle('hidden', !otherChecked);
      if (!otherChecked) input.value = '';
      saveDraft();
    }

    function handleNoWorkChange() {
      const isNoWork = document.getElementById('noWorkToday').checked;
      document.getElementById('noWorkReasonSection').classList.toggle('hidden', !isNoWork);
      document.getElementById('contractorSection').classList.toggle('hidden', isNoWork);
    }

    function handleNoMaterialChange() {
      const noMaterial = document.getElementById('noMaterial').checked;
      const container = document.getElementById('materialsContainer');
      const addButton = document.getElementById('addMaterialButton');
      container.style.display = noMaterial ? 'none' : 'block';
      addButton.style.display = noMaterial ? 'none' : 'block';
      if (!noMaterial && container.children.length === 0) addMaterialRow();
    }

    function addContractorRow(name = '', customName = '', workers = '') {
      const container = document.getElementById('contractorsContainer');
      if (container.children.length >= MAX_CONTRACTOR_ROWS) { alert(`最多可填寫 ${MAX_CONTRACTOR_ROWS} 組廠商`); return; }

      const div = document.createElement('div');
      div.className = 'contractor-group material-group';
      const header = document.createElement('div');
      header.className = 'material-header';
      const title = document.createElement('div');
      title.className = 'material-title';
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'remove-material-button';
      removeBtn.textContent = '移除';
      removeBtn.onclick = () => { div.remove(); refreshContractorTitles(); saveDraft(); };
      header.append(title, removeBtn);

      const row = document.createElement('div');
      row.className = 'flex-row';

      const nameSelect = document.createElement('select');
      nameSelect.className = 'contractor-name';
      nameSelect.add(new Option('請選擇廠商', ''));
      nameSelect.options[0].disabled = true;
      getSavedContractorsList().forEach(c => nameSelect.add(new Option(c, c)));
      nameSelect.add(new Option('其他 (自行輸入)', '其他'));

      const workerInput = document.createElement('input');
      workerInput.type = 'number';
      workerInput.className = 'contractor-workers';
      workerInput.placeholder = '人數';
      workerInput.min = '1'; workerInput.step = '1';
      
      row.append(nameSelect, workerInput);

      const customInput = document.createElement('input');
      customInput.type = 'text';
      customInput.className = 'custom-contractor-name hidden';
      customInput.placeholder = '請輸入新廠商名稱';
      div.append(header, row, customInput);

      nameSelect.onchange = () => {
        const isCustom = nameSelect.value === '其他';
        customInput.classList.toggle('hidden', !isCustom);
        if (isCustom) customInput.focus(); else customInput.value = '';
        saveDraft();
      };
      customInput.oninput = saveDraft;
      workerInput.oninput = saveDraft;

      const isCustomVal = name === '其他' || (name && !getSavedContractorsList().includes(name));
      nameSelect.value = isCustomVal ? '其他' : name;
      if (isCustomVal) { customInput.value = name === '其他' ? customName : name; customInput.classList.remove('hidden'); }
      if (workers !== '') workerInput.value = workers;

      container.appendChild(div);
      refreshContractorTitles();
      if (!isRestoringDraft) saveDraft();
    }

    function refreshContractorTitles() {
      document.querySelectorAll('.contractor-group').forEach((g, i) => g.querySelector('.material-title').textContent = `施工廠商 ${i + 1}`);
    }

    function addMaterialRow(materialId = '', qty = '') {
      const container = document.getElementById('materialsContainer');
      if (container.children.length >= MAX_MATERIAL_ROWS) { alert(`最多 ${MAX_MATERIAL_ROWS} 筆用料`); return; }
      const div = document.createElement('div');
      div.className = 'material-group';

      const header = document.createElement('div');
      header.className = 'material-header';
      const title = document.createElement('div');
      title.className = 'material-title';
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'remove-material-button';
      removeBtn.textContent = '移除';
      removeBtn.onclick = () => { div.remove(); refreshMaterialTitles(); saveDraft(); };
      header.append(title, removeBtn);

      const nameSelect = document.createElement('select');
      nameSelect.className = 'mat-name';
      nameSelect.add(new Option('請選擇標準材料...', ''));
      nameSelect.options[0].disabled = true;
      
      MATERIAL_OPTIONS.forEach(m => {
          const specStr = m.packageQuantity && m.packageUnit && m.stockUnit && m.packageUnit !== m.stockUnit ? ` (${m.packageQuantity}${m.packageUnit}/${m.stockUnit})` : '';
          nameSelect.add(new Option(`${m.materialName}${specStr}`, m.materialId));
      });

      const row = document.createElement('div');
      row.className = 'flex-row';
      const qtyInput = document.createElement('input');
      qtyInput.type = 'number';
      qtyInput.className = 'mat-qty';
      qtyInput.placeholder = '數量';
      qtyInput.step = '0.1'; qtyInput.min = '0.1';

      const unitDisplay = document.createElement('div');
      unitDisplay.className = 'mat-unit-display';
      unitDisplay.style.cssText = "padding:10px; background:#e9ecef; border:1px solid #ccc; border-radius:5px; color:#555; text-align:center; min-width:60px;";
      unitDisplay.textContent = "單位";

      row.append(qtyInput, unitDisplay);
      
      const balanceInfo = document.createElement('div');
      balanceInfo.className = 'material-balance-info hidden';
      balanceInfo.style.cssText = "font-size:13px; color:#555; background:#eef5ee; padding:8px; border-left:3px solid #00B900; margin-top:5px; border-radius:4px;";

      div.append(header, nameSelect, row, balanceInfo);

      nameSelect.onchange = () => {
          const sId = nameSelect.value;
          if (sId && MATERIAL_ITEMS_MAP.has(sId)) {
              unitDisplay.textContent = MATERIAL_ITEMS_MAP.get(sId).stockUnit;
          }
          updateBalanceDisplay(div);
          saveDraft();
      };
      qtyInput.oninput = () => { updateBalanceDisplay(div); saveDraft(); };

      if (materialId && MATERIAL_ITEMS_MAP.has(materialId)) {
          nameSelect.value = materialId;
          unitDisplay.textContent = MATERIAL_ITEMS_MAP.get(materialId).stockUnit;
      }
      if (qty !== '') qtyInput.value = qty;

      container.appendChild(div);
      refreshMaterialTitles();
      if (!isRestoringDraft) saveDraft();
    }

    function refreshMaterialTitles() {
      const groups = document.querySelectorAll('#materialsContainer .material-group');
      groups.forEach((group, index) => group.querySelector('.material-title').textContent = `品項 ${index + 1}`);
      const addButton = document.getElementById('addMaterialButton');
      addButton.disabled = groups.length >= MAX_MATERIAL_ROWS;
    }

    function updateBalanceDisplay(groupDiv) {
        const select = groupDiv.querySelector('.mat-name');
        const qtyInput = groupDiv.querySelector('.mat-qty');
        const infoDiv = groupDiv.querySelector('.material-balance-info');

        const materialId = select.value;
        if (!materialId || !projectMaterialBalances[materialId]) {
            infoDiv.classList.add('hidden');
            return;
        }

        const balanceData = projectMaterialBalances[materialId];
        const inputQty = Number(qtyInput.value) || 0;
        const itemMap = MATERIAL_ITEMS_MAP.get(materialId);
        const pkgQty = itemMap && itemMap.packageQuantity ? Number(itemMap.packageQuantity) : 1;
        const baseUnit = itemMap ? (itemMap.baseUnit || itemMap.stockUnit) : '';

        const currentUsageBaseQty = inputQty * pkgQty;
        const currentRemaining = balanceData.remainingBaseQuantity;
        const expectedRemaining = currentRemaining - currentUsageBaseQty;
        const formType = document.getElementById('formType').value;

        infoDiv.classList.remove('hidden');

        if (formType === 'material_issue') {
            const issueType = document.getElementById('issueType').value;
            const issueLabel = issueType === 'OPENING' ? '本次開工首批進場' : '本次追加進場';
            infoDiv.style.borderLeftColor = '#00B900';
            infoDiv.style.backgroundColor = '#eef5ee';
            infoDiv.innerHTML = `
                <strong>目前案場剩餘：</strong> ${currentRemaining} ${baseUnit}<br>
                <span style="color: #00B900;"><strong>${issueLabel}：</strong> +${currentUsageBaseQty} ${baseUnit}</span><br>
                <strong>登記後預計剩餘：</strong> ${currentRemaining + currentUsageBaseQty} ${baseUnit}
            `;
        } else {
            if (expectedRemaining < 0 && inputQty > 0) {
                infoDiv.style.borderLeftColor = '#dc3545';
                infoDiv.style.backgroundColor = '#fff3f3';
                infoDiv.innerHTML = `
                    <strong>目前案場剩餘：</strong> ${currentRemaining} ${baseUnit}<br>
                    <strong>本次預計耗用：</strong> -${currentUsageBaseQty} ${baseUnit}<br>
                    <strong>送出後預計剩餘：</strong> <span style="color: #dc3545; font-weight: bold;">${expectedRemaining} ${baseUnit}</span><br>
                    <div style="margin-top: 4px; font-size: 12px; color: #dc3545;">⚠️ 送出後請確認是否漏登進場單據</div>
                `;
            } else {
                infoDiv.style.borderLeftColor = '#00B900';
                infoDiv.style.backgroundColor = '#eef5ee';
                infoDiv.innerHTML = `
                    <strong>案場累計進場：</strong> ${balanceData.issuedBaseQuantity} ${baseUnit}<br>
                    <strong>先前已耗用：</strong> ${balanceData.consumedBaseQuantity} ${baseUnit}<br>
                    <strong>目前理論剩餘：</strong> ${currentRemaining} ${baseUnit}<br>
                    <span style="color: #007bff;"><strong>送出後預計剩餘：</strong> ${expectedRemaining} ${baseUnit}</span>
                `;
            }
        }
    }

    async function getWeather() {
      if (!navigator.geolocation) { alert('此裝置不支援定位'); return; }
      navigator.geolocation.getCurrentPosition(async position => {
        try {
          const lat = position.coords.latitude, lon = position.coords.longitude;
          const res = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,relative_humidity_2m,wind_speed_10m&wind_speed_unit=ms&timezone=Asia%2FTaipei`);
          if (!res.ok) throw new Error('氣象服務錯誤');
          const data = await res.json();
          document.getElementById('temp').value = Number(data.current.temperature_2m).toFixed(1);
          document.getElementById('humidity').value = Math.round(data.current.relative_humidity_2m);
          document.getElementById('wind').value = Number(data.current.wind_speed_10m).toFixed(1);
          saveDraft();
        } catch (error) {} 
      }, () => {}, { timeout: 10000 });
    }

    function saveDraft() {
      if (isRestoringDraft || reportArchived || (projectMode !== 'LOCKED' && projectMode !== 'COMPATIBLE')) return;
      
      const materialsList = [...document.querySelectorAll('#materialsContainer .material-group')].map(group => ({
        materialId: group.querySelector('.mat-name').value,
        qty: group.querySelector('.mat-qty').value
      }));

      const contractorsList = [...document.querySelectorAll('#contractorsContainer .contractor-group')].map(group => ({
        name: group.querySelector('.contractor-name').value,
        customName: group.querySelector('.custom-contractor-name')?.value.trim() || '',
        workers: group.querySelector('.contractor-workers').value
      }));

      const draft = {
        formType: document.getElementById('formType').value,
        issueType: document.getElementById('issueType').value,
        submissionId: ensureSubmissionId(),
        projectId: document.getElementById('projectId').value.trim(),
        reporterName: document.getElementById('reporterName').value.trim(), 
        reportDate: document.getElementById('reportDate').value,
        isNoWork: document.getElementById('noWorkToday').checked,
        noWorkReason: document.getElementById('noWorkReason').value,
        contractorsList, 
        workItems: [...document.querySelectorAll('input[name="workItem"]:checked')].map(cb => cb.value),
        customWorkItem: document.getElementById('customWorkItem').value,
        workNotes: document.getElementById('workNotes').value,
        noMaterial: document.getElementById('noMaterial').checked,
        materialsList,
        remarks: document.getElementById('remarks').value,
        temp: document.getElementById('temp').value,
        humidity: document.getElementById('humidity').value,
        wind: document.getElementById('wind').value
      };
      localStorage.setItem(getDraftKey(), JSON.stringify(draft));
    }

    function loadDraft() {
      isRestoringDraft = true;
      try {
        const saved = localStorage.getItem(getDraftKey());
        if (!saved) { 
          currentSubmissionId = createSubmissionId();
          addContractorRow(); addMaterialRow(); return; 
        }
        
        const draft = JSON.parse(saved);
        if (draft.formType) switchMode(draft.formType);
        if (draft.issueType) document.getElementById('issueType').value = draft.issueType;
        if (draft.reportDate) document.getElementById('reportDate').value = draft.reportDate;

        currentSubmissionId = String(draft.submissionId || createSubmissionId());
        
        if (draft.isNoWork) {
            document.getElementById('noWorkToday').checked = true;
            document.getElementById('noWorkReason').value = draft.noWorkReason || '';
            handleNoWorkChange();
        }

        document.getElementById('temp').value = draft.temp || '';
        document.getElementById('humidity').value = draft.humidity || '';
        document.getElementById('wind').value = draft.wind || '';
        document.getElementById('workNotes').value = draft.workNotes || '';
        document.getElementById('remarks').value = draft.remarks || '';

        // ⭐ 完全補回草稿恢復機制
        document.getElementById('reporterName').value = draft.reporterName || document.getElementById('reporterName').value;

        document.getElementById('contractorsContainer').innerHTML = '';
        if (Array.isArray(draft.contractorsList) && draft.contractorsList.length > 0) {
            draft.contractorsList.forEach(c => addContractorRow(c.name, c.customName, c.workers));
        } else {
            addContractorRow();
        }

        if (Array.isArray(draft.workItems)) {
            document.querySelectorAll('input[name="workItem"]').forEach(checkbox => {
                checkbox.checked = draft.workItems.includes(checkbox.value);
            });
            const hasOther = draft.workItems.includes('其他');
            const customWorkInput = document.getElementById('customWorkItem');
            customWorkInput.classList.toggle('hidden', !hasOther);
            customWorkInput.value = hasOther ? (draft.customWorkItem || '') : '';
        }

        const noMaterial = draft.noMaterial === true;
        document.getElementById('noMaterial').checked = noMaterial;
        
        if (noMaterial) {
          handleNoMaterialChange();
        } else {
          document.getElementById('materialsContainer').style.display = 'block';
          if (Array.isArray(draft.materialsList) && draft.materialsList.length > 0) {
            document.getElementById('materialsContainer').innerHTML = '';
            draft.materialsList.forEach(m => addMaterialRow(m.materialId || '', m.qty || ''));
          } else { addMaterialRow(); }
        }
      } catch (error) {
        localStorage.removeItem(getDraftKey());
        document.getElementById('contractorsContainer').innerHTML = '';
        addContractorRow(); addMaterialRow();
      } finally {
        isRestoringDraft = false;
      }
    }

    function buildMaterialItems() {
      const formType = document.getElementById('formType').value;
      const isNoWork = formType === 'daily_report' && document.getElementById('noWorkToday').checked;
      const isNoMaterial = formType === 'daily_report' && document.getElementById('noMaterial').checked;
      if (isNoWork || isNoMaterial) return [];

      return [...document.querySelectorAll('#materialsContainer .material-group')]
        .map(group => {
          const selectedId = group.querySelector('.mat-name').value;
          if (!selectedId || !MATERIAL_ITEMS_MAP.has(selectedId)) return null;

          const item = MATERIAL_ITEMS_MAP.get(selectedId);
          const quantity = Number(group.querySelector('.mat-qty').value) || 0;

          return {
            materialId: item.materialId,
            materialCode: item.materialCode,
            materialName: item.materialName,
            quantity,
            stockUnit: item.stockUnit,
            packageQuantity: item.packageQuantity,
            packageUnit: item.packageUnit,
            baseQuantity: quantity * (item.packageQuantity || 1),
            baseUnit: item.baseUnit || item.stockUnit
          };
        })
        .filter(item => item && item.quantity > 0);
    }

    function validateProject() {
      if (projectMode !== 'LOCKED' && projectMode !== 'COMPATIBLE') { alert('案場資料尚未載入'); return false; }
      if (!getSelectedProjectName()) { alert('請選擇案場'); return false; }
      return true;
    }

    function validateForm() {
      const reporterName = document.getElementById('reporterName').value.trim();
      if (!reporterName) { alert('請填寫您的填表人姓名'); return false; }

      const formType = document.getElementById('formType').value;
      const mGroups = [...document.querySelectorAll('#materialsContainer .material-group')];
      const selectedMatIds = new Set();

      if (formType === 'material_issue') {
          const issueType = document.getElementById('issueType').value;
          const reportDate = document.getElementById('reportDate').value;

          if (!issueType) { alert('請選擇進場類型'); return false; }
          if (!reportDate) { alert('請選擇實際進場日期'); return false; }

          let validCount = 0;
          for (let i = 0; i < mGroups.length; i++) {
              const selectedId = mGroups[i].querySelector('.mat-name').value;
              const qty = Number(mGroups[i].querySelector('.mat-qty').value);
              
              if (!selectedId && !qty) continue;
              if (!selectedId) { alert(`請選擇進場材料 ${i + 1} 的品項`); return false; }
              if (!qty || qty <= 0) { alert(`請輸入進場材料 ${i + 1} 的正確數量`); return false; }
              
              if (selectedMatIds.has(selectedId)) {
                  alert(`材料重複選取，請將相同材料合併為一筆數量計算`);
                  return false;
              }
              selectedMatIds.add(selectedId);
              validCount++;
          }
          if (validCount === 0) { alert('請至少填寫一筆進場材料'); return false; }
          return true;
      }

      const isNoWork = document.getElementById('noWorkToday').checked;
      if (isNoWork) {
         if (!document.getElementById('noWorkReason').value) { alert('請選擇今日無出工原因'); return false; }
      } else {
          const cGroups = document.querySelectorAll('#contractorsContainer .contractor-group');
          if (cGroups.length === 0) { alert('請至少填寫一組施工廠商'); return false; }
          
          const contractorNames = new Set(); 
          for (let i = 0; i < cGroups.length; i++) {
              const group = cGroups[i];
              const name = group.querySelector('.contractor-name').value;
              const custom = group.querySelector('.custom-contractor-name')?.value.trim() || '';
              const workers = group.querySelector('.contractor-workers').value;

              if (!name) { alert(`請選擇施工廠商 ${i + 1} 的名稱`); return false; }
              if (name === '其他' && !custom) { alert(`請輸入施工廠商 ${i + 1} 的自訂名稱`); return false; }
              
              const workerCount = Number(workers);
              if (!Number.isInteger(workerCount) || workerCount <= 0 || workerCount > 200) { 
                  alert(`請輸入施工廠商 ${i + 1} 的正確整數人數`); 
                  return false; 
              }

              const finalContractorName = name === '其他' ? custom : name;
              const normalizedContractorName = finalContractorName.normalize('NFKC').replace(/\s+/g, ' ').trim();
              
              if (contractorNames.has(normalizedContractorName)) {
                  alert(`施工廠商「${finalContractorName}」重複填寫，請合併人數`);
                  return false;
              }
              contractorNames.add(normalizedContractorName);
          }

          const selectedWorkItems = [...document.querySelectorAll('input[name="workItem"]:checked')];
          const workNotes = document.getElementById('workNotes').value.trim();
          
          if (selectedWorkItems.length === 0 && !workNotes) { 
              alert('請選擇施工作業項目或填寫作業內容補充'); return false; 
          }
          if (selectedWorkItems.some(cb => cb.value === '其他') && !document.getElementById('customWorkItem').value.trim()) { 
              alert('請輸入其他作業項目名稱'); return false; 
          }
          
          if (!document.getElementById('noMaterial').checked) {
              let matValidCount = 0;
              for (let i = 0; i < mGroups.length; i++) {
                  const selectedId = mGroups[i].querySelector('.mat-name').value;
                  const qty = Number(mGroups[i].querySelector('.mat-qty').value);
                  
                  if (!selectedId && !qty) continue;
                  if (!selectedId) { alert(`請選擇用料 ${i + 1} 的品項`); return false; }
                  if (!qty || qty <= 0) { alert(`請輸入用料 ${i + 1} 的正確數量`); return false; }
                  
                  if (selectedMatIds.has(selectedId)) {
                      alert(`材料重複選取，請將相同材料合併為一筆數量計算`);
                      return false;
                  }
                  selectedMatIds.add(selectedId);
                  matValidCount++;
              }
              if (matValidCount === 0) { alert('請選擇今日用料，或勾選「今日無用料」'); return false; }
          }
      }
      return true;
    }

    async function submitForm() {
      if (isSubmitting) return;
      if (!validateProject()) return;
      if (!validateForm()) return;

      const submitButton = document.getElementById('submitButton');
      isSubmitting = true;
      submitButton.disabled = true;
      submitButton.textContent = '⏳ 正在送出，請稍候...';

      const formType = document.getElementById('formType').value;
      let payload = {
        formType: formType,
        projectId: document.getElementById('projectId').value.trim(),
        projectName: getSelectedProjectName(),
        submissionId: ensureSubmissionId(),
        reporterName: document.getElementById('reporterName').value.trim(),
        remarks: document.getElementById('remarks').value.trim() || '無'
      };

      if (formType === 'material_issue') {
          payload.issueType = document.getElementById('issueType').value;
          payload.date = document.getElementById('reportDate').value;
          payload.materialItems = buildMaterialItems();
      } else {
          const isNoWork = document.getElementById('noWorkToday').checked;
          if (isNoWork) {
              payload.isNoWork = true;
              payload.noWorkReason = document.getElementById('noWorkReason').value;
              payload.contractor = '無出工'; payload.workerCount = '0'; payload.progress = '本日無出工'; payload.materials = '無';
              payload.contractorItems = [];
          } else {
              let totalWorkers = 0;
              const contractorItems = [];
              const contractorListStr = [];
              
              document.querySelectorAll('#contractorsContainer .contractor-group').forEach(group => {
                  const select = group.querySelector('.contractor-name').value;
                  const name = select === '其他' ? group.querySelector('.custom-contractor-name').value.trim() : select;
                  const workers = Number(group.querySelector('.contractor-workers').value);
                  
                  if (name && Number.isInteger(workers) && workers > 0) {
                      contractorItems.push({ contractorName: name, workerCount: workers });
                      contractorListStr.push(`• ${name}（${workers}人）`);
                      totalWorkers += workers;
                  }
              });

              payload.isNoWork = false;
              payload.contractorItems = contractorItems;
              payload.contractor = '\n' + contractorListStr.join('\n');
              payload.workerCount = `共 ${totalWorkers} 人`;
              
              const workItems = [...document.querySelectorAll('input[name="workItem"]:checked')].map(i => i.value);
              const customWork = document.getElementById('customWorkItem').value.trim();
              payload.workItems = workItems;
              payload.customWorkItem = customWork;
              payload.workNotes = document.getElementById('workNotes').value.trim();
              
              let progressStr = workItems.map(w => `• ${w === '其他' ? customWork : w}`).join('\n');
              if (payload.workNotes) {
                  progressStr += (progressStr ? '\n' : '') + `(補充: ${payload.workNotes})`;
              }
              payload.progress = progressStr || '未填寫施作項目';
              
              payload.materialItems = buildMaterialItems();
              payload.materials = document.getElementById('noMaterial').checked ? '無' : payload.materialItems.map(m => `${m.materialName}：${m.quantity}${m.stockUnit}`).join('\n');
          }
      }

      try {
        const response = await fetch(API_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        
        const responseText = await response.text();
        let result = {};
        try {
            result = responseText ? JSON.parse(responseText) : {};
        } catch {
            throw new Error(`後端回傳格式錯誤，HTTP ${response.status}`);
        }
        
        if (!response.ok || result.success !== true) {
            throw new Error(result.error || result.message || `HTTP ${response.status}`);
        }

        localStorage.removeItem(getDraftKey());
        alert(result.message || '✅ 資料已成功送出與歸檔');
        if (window.liff && liff.isInClient()) liff.closeWindow();
        
      } catch (error) {
        console.error('送出失敗完整資訊：', error);
        alert(
            `⚠️ 資料送出失敗\n\n` +
            `原因：${error.message || '未知錯誤'}\n\n` +
            `表單內容仍保留，請勿重新填寫。`
        );
      } finally {
        isSubmitting = false;
        submitButton.disabled = false;
        submitButton.textContent = formType === 'material_issue' ? '確認送出材料進場紀錄' : '確認送出施工日報';
      }
    }
  </script>
</body>
</html>
