<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>云說工程 - 施工日報表</title>
  <script charset="utf-8" src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>
  <style>
    body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; background:#f4f7f6; padding:15px; color:#333; margin:0; }
    .container { background:white; padding:20px; border-radius:10px; box-shadow:0 4px 8px rgba(0,0,0,.1); max-width:500px; margin:auto; }
    h2 { text-align:center; color:#00B900; margin-top:0; font-size:22px; }
    label { font-weight:bold; display:block; margin-top:15px; margin-bottom:5px; color:#444; }
    select, textarea, input[type="text"], input[type="number"] { width:100%; padding:10px; border:1px solid #ccc; border-radius:5px; box-sizing:border-box; font-size:16px; background:#fafafa; margin-bottom:5px; }
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
    <h2>👷 施工日報表</h2>
    <form id="dailyReportForm">
      
      <!-- 💡 新增：填表人欄位 -->
      <label for="reporterName">👤 填表人姓名</label>
      <input type="text" id="reporterName" placeholder="請輸入您的姓名 (將自動記憶)" oninput="saveDraft()" required>

      <div id="lockedProjectSection" class="hidden">
        <label for="lockedProjectName">📍 案場名稱</label>
        <input type="text" id="lockedProjectName" class="readonly-input" readonly>
        <div class="field-help" style="color: #666; font-size: 0.9em; margin-top: 4px;">
          🔒 此連結已綁定以下案場，案場名稱不可修改
        </div>
      </div>

      <div id="compatibleProjectSection">
        <label for="projectName">📍 案場名稱</label>
        <select id="projectName" onchange="saveDraft()" required disabled>
          <option value="" disabled selected>正在載入案場清單...</option>
        </select>
        <div class="field-help" style="color: #666; font-size: 0.9em; margin-top: 4px;">
          新案場請先在施工群組輸入「設定案場 案場名稱」
        </div>
      </div>

      <input type="hidden" id="projectId" value="">

      <!-- 今日無出工選項 -->
      <label class="no-material-option" style="margin-top:15px; background:#fff3f3; border-color:#ffcccc; color:#cc0000;">
        <input type="checkbox" id="noWorkToday" onchange="handleNoWorkChange(); saveDraft()">今日無出工 (停工/休息)
      </label>

      <!-- 無出工原因 -->
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

      <div id="contractorSection">
        <label>🤝 施工廠商與人數</label>
        <div id="contractorsContainer"></div>
        <button type="button" class="btn-add" onclick="addContractorRow()" style="background:#5bc0de;">➕ 新增一組廠商</button>
        <button type="button" class="btn-add" onclick="resetContractorsCache()" style="background:#d9534f; margin-top:5px;">🗑️ 清除所有自訂廠商紀錄</button>
        <div class="section-divider"></div>
      </div>

      <label>🌤️ 氣象資訊</label>
      <button type="button" class="btn-weather" onclick="getWeather()">一鍵獲取當地天氣</button>
      <div class="flex-row">
        <input type="number" id="temp" placeholder="溫度(度)" step="0.1" oninput="saveDraft()">
        <input type="number" id="humidity" placeholder="濕度(%)" min="0" max="100" step="1" oninput="saveDraft()">
        <input type="number" id="wind" placeholder="風速(m/s)" min="0" step="0.1" oninput="saveDraft()">
      </div>
      <div class="section-divider"></div>

      <div id="workAndMaterialSection">
        <label>🛠️ 今日作業</label>
        <div class="checkbox-grid" id="workItemsGrid"></div>
        <input type="text" id="customWorkItem" class="hidden" placeholder="請輸入其他作業項目" maxlength="100" oninput="saveDraft()" style="margin-top:10px">
        <textarea id="workNotes" rows="2" placeholder="作業補充（選填）" oninput="saveDraft()" style="margin-top:10px"></textarea>
        <div class="section-divider"></div>

        <label>📦 今日用料</label>
        <label class="no-material-option">
          <input type="checkbox" id="noMaterial" onchange="handleNoMaterialChange(); saveDraft()">今日無用料
        </label>
        <div id="materialsContainer"></div>
        <button type="button" id="addMaterialButton" class="btn-add" onclick="addMaterialRow()">新增一項用料</button>

        <div class="section-divider"></div>
      </div>

      <label for="remarks">📝 備註（選填）</label>
      <textarea id="remarks" rows="2" placeholder="天候影響、材料短缺、人員異動等狀況" oninput="saveDraft()"></textarea>

      <button type="button" id="submitButton" class="submit-btn" onclick="submitForm()">確認送出並發布至群組</button>
    </form>
  </div>

  <script>
    const LIFF_ID = '2011289657-vQgMb0eI';
    const API_URL = 'https://gong-cheng-linechuan-jie.onrender.com/api/submit-report';
    const PROJECTS_API_URL = 'https://gong-cheng-linechuan-jie.onrender.com/api/projects';
    const MATERIALS_API_URL = 'https://gong-cheng-linechuan-jie.onrender.com/api/materials';
    
    const MAX_MATERIAL_ROWS = 10;
    const MAX_CONTRACTOR_ROWS = 10;
    const CONTRACTORS_KEY = 'savedContractors';

    const workItemsList = [
      '材料進場', '現場放樣', '鷹架搭設', '鷹架拆除', '水刀除鏽',
      '手動工具除鏽', '底漆施工', '中塗漆施工', '面漆施工', '底漆修補',
      '中塗漆修補', '面漆修補', '底漆膜厚檢測', '中塗漆膜厚檢測',
      '面漆膜厚檢測', '缺失改善', '現場清潔', '等待驗收', '完成驗收', '其他'
    ];

    let MATERIAL_OPTIONS = [];

    const UNIT_OPTIONS = [
      '桶 (1加侖)', '桶 (5加侖)', '桶', '組', '支', '公斤', '公升', '個', '捲'
    ];

    let isSubmitting = false;
    let isRestoringDraft = false;
    let reportArchived = false;
    let projectMode = 'LOADING';
    let currentSubmissionId = '';

    function createSubmissionId() {
      if (window.crypto && typeof window.crypto.randomUUID === 'function') {
        return window.crypto.randomUUID();
      }
      return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2);
    }

    function ensureSubmissionId() {
      if (!currentSubmissionId) {
        currentSubmissionId = createSubmissionId();
      }
      return currentSubmissionId;
    }

    function getSelectedProjectName() {
      if (projectMode === 'LOCKED') {
        return document.getElementById('lockedProjectName').value.trim();
      }
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
        } else {
            MATERIAL_OPTIONS = ['立邦9011', '虹牌MIO', '虹牌優利面UP-04', '矽利康'];
        }
      } catch (err) {
        console.warn('材料清單讀取失敗，使用預設值', err);
        MATERIAL_OPTIONS = ['立邦9011', '虹牌MIO', '虹牌優利面UP-04', '矽利康'];
      }
    }

    window.addEventListener('load', async () => {
      initializeStaticForm();

      try {
        await liff.init({ liffId: LIFF_ID });
      } catch (error) {
        console.error('LIFF 初始化失敗：', error);
      }

      const projectIdFromUrl = getProjectIdFromUrl();

      try {
        if (projectIdFromUrl) {
          await loadLockedProject(projectIdFromUrl);
          await fetchMaterials(projectIdFromUrl);
        } else {
          await loadProjects();
          await fetchMaterials('');
        }
        loadDraft();
      } catch (error) {
        console.error('案場初始化失敗：', error);
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

      if (result.projects.length === 0) {
        projectSelect.innerHTML = '';
        const emptyOption = new Option('目前沒有可用案場', '');
        emptyOption.disabled = true;
        emptyOption.selected = true;
        projectSelect.add(emptyOption);
        projectSelect.disabled = true;
        projectMode = 'ERROR';
        throw new Error('目前沒有可用案場');
      }

      projectSelect.innerHTML = '';
      const placeholder = new Option('請選擇正式案場', '');
      placeholder.disabled = true;
      placeholder.selected = true;
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

    function resetContractorsCache() {
      if (confirm('確定要清除所有自行輸入的廠商名稱紀錄嗎？(容煜、詮達等預設廠商將保留)')) {
        localStorage.removeItem(CONTRACTORS_KEY);
        document.querySelectorAll('.contractor-name').forEach(select => {
            const currentVal = select.value;
            select.innerHTML = '';
            select.add(new Option('請選擇廠商', ''));
            select.options[0].disabled = true;
            getSavedContractorsList().forEach(c => select.add(new Option(c, c)));
            select.add(new Option('其他 (自行輸入)', '其他'));
            const exists = [...select.options].some(opt => opt.value === currentVal);
            select.value = exists ? currentVal : (currentVal ? '其他' : '');
            toggleCustomContractorRow(select);
        });
        saveDraft();
        alert('已清除自訂廠商紀錄！');
      }
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
      document.getElementById('workAndMaterialSection').classList.toggle('hidden', isNoWork);
    }

    function handleNoMaterialChange() {
      const noMaterial = document.getElementById('noMaterial').checked;
      const container = document.getElementById('materialsContainer');
      const addButton = document.getElementById('addMaterialButton');
      container.style.display = noMaterial ? 'none' : 'block';
      addButton.style.display = noMaterial ? 'none' : 'block';
      if (!noMaterial && container.children.length === 0) addMaterialRow();
    }

    function appendOption(select, value, text = value) { select.add(new Option(text, value)); }

    function addContractorRow(name = '', customName = '', workers = '') {
      const container = document.getElementById('contractorsContainer');
      if (container.children.length >= MAX_CONTRACTOR_ROWS) {
        alert(`最多可填寫 ${MAX_CONTRACTOR_ROWS} 組廠商`); return;
      }

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
      workerInput.placeholder = '施工人數';
      workerInput.min = '1';
      workerInput.step = '1';            
      workerInput.inputMode = 'numeric'; 
      workerInput.style.maxWidth = '100px';

      row.append(nameSelect, workerInput);

      const customInput = document.createElement('input');
      customInput.type = 'text';
      customInput.className = 'custom-contractor-name hidden';
      customInput.placeholder = '請輸入新廠商名稱';
      customInput.maxLength = 50;
      customInput.style.marginTop = '5px';

      div.append(header, row, customInput);

      nameSelect.onchange = () => { toggleCustomContractorRow(nameSelect); saveDraft(); };
      customInput.oninput = saveDraft;
      workerInput.oninput = saveDraft;

      const isCustomVal = name === '其他' || (name && !getSavedContractorsList().includes(name));
      nameSelect.value = isCustomVal ? '其他' : name;
      if (isCustomVal) {
         customInput.value = name === '其他' ? customName : name;
         customInput.classList.remove('hidden');
      }
      if (workers !== '') workerInput.value = workers;

      container.appendChild(div);
      refreshContractorTitles();
      if (!isRestoringDraft) saveDraft();
    }

    function refreshContractorTitles() {
      const groups = document.querySelectorAll('.contractor-group');
      groups.forEach((group, index) => group.querySelector('.material-title').textContent = `施工廠商 ${index + 1}`);
    }

    function toggleCustomContractorRow(select) {
      const input = select.closest('.contractor-group').querySelector('.custom-contractor-name');
      const isCustom = select.value === '其他';
      input.classList.toggle('hidden', !isCustom);
      if (isCustom) input.focus(); else input.value = '';
    }

    function saveNewContractorsFromForm() {
       const groups = document.querySelectorAll('.contractor-group');
       let savedList = getSavedContractorsList();
       let updated = false;
       groups.forEach(group => {
          const selectVal = group.querySelector('.contractor-name').value;
          if (selectVal === '其他') {
             const custom = group.querySelector('.custom-contractor-name').value.trim();
             if (custom && !savedList.includes(custom)) {
                savedList.push(custom);
                updated = true;
             }
          }
       });
       if (updated) localStorage.setItem(CONTRACTORS_KEY, JSON.stringify(savedList));
    }

    function addMaterialRow(name = '', customName = '', qty = '', unit = '') {
      const container = document.getElementById('materialsContainer');
      if (container.children.length >= MAX_MATERIAL_ROWS) {
        alert(`最多可填寫 ${MAX_MATERIAL_ROWS} 筆用料`); return;
      }
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
      appendOption(nameSelect, '', '請選擇品項');
      nameSelect.options[0].disabled = true;
      MATERIAL_OPTIONS.forEach(option => appendOption(nameSelect, option));
      appendOption(nameSelect, '其他');

      const customInput = document.createElement('input');
      customInput.type = 'text';
      customInput.className = 'custom-mat-name hidden';
      customInput.placeholder = '請輸入材料名稱';
      customInput.maxLength = 50;

      const row = document.createElement('div');
      row.className = 'flex-row';
      const qtyInput = document.createElement('input');
      qtyInput.type = 'number';
      qtyInput.className = 'mat-qty';
      qtyInput.placeholder = '數量';
      qtyInput.step = '0.5';
      qtyInput.min = '0.5';

      const unitSelect = document.createElement('select');
      unitSelect.className = 'mat-unit';
      appendOption(unitSelect, '', '單位');
      unitSelect.options[0].disabled = true;
      UNIT_OPTIONS.forEach(option => appendOption(unitSelect, option));
      row.append(qtyInput, unitSelect);
      div.append(header, nameSelect, customInput, row);

      nameSelect.onchange = () => { toggleCustomMaterial(nameSelect); saveDraft(); };
      customInput.oninput = saveDraft;
      qtyInput.oninput = saveDraft;
      unitSelect.onchange = saveDraft;

      const isCustom = name === '其他' || (name && !MATERIAL_OPTIONS.includes(name));
      nameSelect.value = isCustom ? '其他' : name;
      if (isCustom) {
        customInput.value = name === '其他' ? customName : name;
        customInput.classList.remove('hidden');
      }
      if (qty !== '') qtyInput.value = qty;
      if (unit) unitSelect.value = unit;

      container.appendChild(div);
      refreshMaterialTitles();
      if (!isRestoringDraft) saveDraft();
    }

    function refreshMaterialTitles() {
      const groups = document.querySelectorAll('#materialsContainer .material-group');
      groups.forEach((group, index) => group.querySelector('.material-title').textContent = `用料 ${index + 1}`);
      const addButton = document.getElementById('addMaterialButton');
      addButton.disabled = groups.length >= MAX_MATERIAL_ROWS;
      addButton.textContent = groups.length >= MAX_MATERIAL_ROWS ? `已達上限，共 ${MAX_MATERIAL_ROWS} 筆` : '新增一項用料';
    }

    function toggleCustomMaterial(select) {
      const input = select.closest('.material-group').querySelector('.custom-mat-name');
      const isCustom = select.value === '其他';
      input.classList.toggle('hidden', !isCustom);
      if (isCustom) input.focus(); else input.value = '';
    }

    function formatMaterial(name, qty, unit) {
      if (unit === '桶 (1加侖)') return `${name}：${qty}桶，每桶1加侖`;
      if (unit === '桶 (5加侖)') return `${name}：${qty}桶，每桶5加侖`;
      return `${name}：${qty}${unit}`;
    }

    async function getWeather() {
      if (!navigator.geolocation) { alert('此裝置不支援定位，請手動輸入氣象'); return; }
      const button = document.querySelector('.btn-weather');
      button.disabled = true;
      button.textContent = '正在取得天氣...';
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
        } catch (error) {
          alert('氣象抓取失敗，請手動輸入');
        } finally {
          button.disabled = false; button.textContent = '一鍵獲取當地天氣';
        }
      }, () => {
        button.disabled = false; button.textContent = '一鍵獲取當地天氣';
        alert('請允許網頁取用位置，或手動輸入氣象');
      }, { timeout: 10000 });
    }

    function saveDraft() {
      if (isRestoringDraft || reportArchived || (projectMode !== 'LOCKED' && projectMode !== 'COMPATIBLE')) return;
      
      const materialsList = [...document.querySelectorAll('#materialsContainer .material-group')].map(group => ({
        name: group.querySelector('.mat-name').value,
        customName: group.querySelector('.custom-mat-name').value,
        qty: group.querySelector('.mat-qty').value,
        unit: group.querySelector('.mat-unit').value
      }));

      const contractorsList = [...document.querySelectorAll('#contractorsContainer .contractor-group')].map(group => ({
        name: group.querySelector('.contractor-name').value,
        customName: group.querySelector('.custom-contractor-name').value,
        workers: group.querySelector('.contractor-workers').value
      }));

      const draft = {
        submissionId: ensureSubmissionId(),
        projectId: document.getElementById('projectId').value.trim(),
        projectName: getSelectedProjectName(),
        reporterName: document.getElementById('reporterName').value.trim(), // 💡 儲存填表人
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
        
        // 💡 載入填表人記憶
        const lastReporter = localStorage.getItem('lastReporterName') || '';

        if (!saved) { 
          currentSubmissionId = createSubmissionId();
          document.getElementById('reporterName').value = lastReporter;
          addContractorRow(); 
          addMaterialRow(); 
          return; 
        }
        
        const draft = JSON.parse(saved);
        currentSubmissionId = String(draft.submissionId || createSubmissionId());
        
        document.getElementById('reporterName').value = draft.reporterName || lastReporter;

        const currentProjectId = document.getElementById('projectId').value.trim();
        
        if (projectMode === 'LOCKED') {
          if ((draft.projectId && draft.projectId !== currentProjectId) || (!draft.projectId && draft.projectName && draft.projectName !== getSelectedProjectName())) {
            console.warn('草稿不符合目前鎖定案場，已略過');
            addContractorRow(); addMaterialRow(); return;
          }
        }

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

        document.getElementById('contractorsContainer').innerHTML = '';
        if (Array.isArray(draft.contractorsList) && draft.contractorsList.length > 0) {
            draft.contractorsList.forEach(c => addContractorRow(c.name, c.customName, c.workers));
        } else if (draft.contractor) {
            addContractorRow(draft.contractor, draft.customContractor || '', draft.workerCount || '');
        } else {
            addContractorRow();
        }

        if (Array.isArray(draft.workItems)) {
          document.querySelectorAll('input[name="workItem"]').forEach(cb => cb.checked = draft.workItems.includes(cb.value));
          const hasOtherWork = draft.workItems.includes('其他');
          const customWorkInput = document.getElementById('customWorkItem');
          customWorkInput.classList.toggle('hidden', !hasOtherWork);
          customWorkInput.value = hasOtherWork ? (draft.customWorkItem || '') : '';
        }

        const noMaterial = draft.noMaterial === true;
        document.getElementById('noMaterial').checked = noMaterial;
        
        if (noMaterial) {
          document.getElementById('materialsContainer').innerHTML = '';
          handleNoMaterialChange();
        } else {
          document.getElementById('materialsContainer').style.display = 'block';
          document.getElementById('addMaterialButton').style.display = 'block';
          if (Array.isArray(draft.materialsList) && draft.materialsList.length > 0) {
            document.getElementById('materialsContainer').innerHTML = '';
            draft.materialsList.forEach(m => addMaterialRow(m.name || '', m.customName || '', m.qty || '', m.unit || ''));
          } else {
            addMaterialRow();
          }
        }

        if (projectMode === 'COMPATIBLE' && draft.projectName) {
          const projectSelect = document.getElementById('projectName');
          if ([...projectSelect.options].some(option => option.value === draft.projectName)) {
            projectSelect.value = draft.projectName;
          }
        }
      } catch (error) {
        localStorage.removeItem(getDraftKey());
        document.getElementById('contractorsContainer').innerHTML = '';
        document.getElementById('materialsContainer').innerHTML = '';
        addContractorRow(); addMaterialRow();
      } finally {
        isRestoringDraft = false;
      }
    }

    function buildProgressString() {
      const customWorkItem = document.getElementById('customWorkItem').value.trim();
      const items = [...document.querySelectorAll('input[name="workItem"]:checked')].map(item =>
        `• ${item.value === '其他' ? customWorkItem : item.value}`
      );
      const notes = document.getElementById('workNotes').value.trim();
      return items.join('\n') + (notes ? `\n\n作業補充：\n${notes}` : '');
    }

    function buildMaterialString() {
      if (document.getElementById('noMaterial').checked) return '無';
      const materials = [];
      document.querySelectorAll('#materialsContainer .material-group').forEach(group => {
        const selected = group.querySelector('.mat-name').value;
        const name = selected === '其他' ? group.querySelector('.custom-mat-name').value.trim() : selected;
        const qty = group.querySelector('.mat-qty').value;
        const unit = group.querySelector('.mat-unit').value;
        if (name && qty && unit) materials.push(formatMaterial(name, qty, unit));
      });
      return materials.join('\n');
    }

    function buildContractorItems() {
      if (document.getElementById('noWorkToday').checked) {
        return [];
      }
      return [...document.querySelectorAll('#contractorsContainer .contractor-group')]
        .map(group => {
          const selected = group.querySelector('.contractor-name').value;
          const contractorName = selected === '其他' 
            ? group.querySelector('.custom-contractor-name').value.trim() 
            : selected;
          const workerCount = Number(group.querySelector('.contractor-workers').value) || 0;
          return { contractorName, workerCount };
        })
        .filter(item => {
          return item.contractorName && item.workerCount > 0;
        });
    }

    function buildWorkItems() {
      if (document.getElementById('noWorkToday').checked) {
        return [];
      }
      return [...document.querySelectorAll('input[name="workItem"]:checked')]
        .map(item => item.value);
    }

    function buildMaterialItems() {
      const isNoWork = document.getElementById('noWorkToday').checked;
      const isNoMaterial = document.getElementById('noMaterial').checked;
      if (isNoWork || isNoMaterial) {
        return [];
      }

      return [...document.querySelectorAll('#materialsContainer .material-group')]
        .map(group => {
          const selected = group.querySelector('.mat-name').value;
          const materialName = selected === '其他' 
            ? group.querySelector('.custom-mat-name').value.trim() 
            : selected;
          const quantity = Number(group.querySelector('.mat-qty').value) || 0;
          const selectedUnit = group.querySelector('.mat-unit').value;

          let stockUnit = selectedUnit;
          let packageQuantity = null;
          let packageUnit = null;
          let baseQuantity = quantity;
          let baseUnit = selectedUnit;

          if (selectedUnit === '桶 (1加侖)') {
            stockUnit = '桶';
            packageQuantity = 1;
            packageUnit = '加侖';
            baseQuantity = quantity;
            baseUnit = '加侖';
          } else if (selectedUnit === '桶 (5加侖)') {
            stockUnit = '桶';
            packageQuantity = 5;
            packageUnit = '加侖';
            baseQuantity = quantity * 5;
            baseUnit = '加侖';
          }

          return {
            materialId: null,
            materialName,
            quantity,
            stockUnit,
            packageQuantity,
            packageUnit,
            baseQuantity,
            baseUnit
          };
        })
        .filter(item => {
          return item.materialName && item.quantity > 0;
        });
    }

    function validateProject() {
      if (projectMode !== 'LOCKED' && projectMode !== 'COMPATIBLE') { alert('案場資料尚未成功載入'); return false; }
      if (!getSelectedProjectName()) { alert('請選擇案場'); return false; }
      if (projectMode === 'LOCKED' && !document.getElementById('projectId').value.trim()) { alert('案場識別資料不完整'); return false; }
      return true;
    }

    function validateForm() {
      const reporterName = document.getElementById('reporterName').value.trim();
      if (!reporterName) { alert('請填寫您的填表人姓名'); return false; }

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
              const custom = group.querySelector('.custom-contractor-name').value.trim();
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
          if (!selectedWorkItems.length) { alert('請至少勾選一項今日作業'); return false; }
          if (selectedWorkItems.some(cb => cb.value === '其他') && !document.getElementById('customWorkItem').value.trim()) { alert('請輸入其他作業項目名稱'); return false; }

          if (!document.getElementById('noMaterial').checked) {
            let completed = 0;
            const groups = [...document.querySelectorAll('#materialsContainer .material-group')];
            for (let i = 0; i < groups.length; i++) {
              const group = groups[i];
              const selected = group.querySelector('.mat-name').value;
              const custom = group.querySelector('.custom-mat-name').value.trim();
              const qty = Number(group.querySelector('.mat-qty').value);
              const unit = group.querySelector('.mat-unit').value;
              
              if (!selected && !custom && !qty && !unit) continue;
              if (!selected) { alert(`請選擇用料 ${i + 1} 的品項`); return false; }
              if (selected === '其他' && !custom) { alert(`請輸入用料 ${i + 1} 的名稱`); return false; }
              if (!qty || qty <= 0) { alert(`請輸入用料 ${i + 1} 的正確數量`); return false; }
              if (!unit) { alert(`請選擇用料 ${i + 1} 的單位`); return false; }
              completed++;
            }
            if (!completed) { alert('請至少填寫一筆今日用料，或勾選「今日無用料」'); return false; }
          }
      }

      const temp = document.getElementById('temp').value.trim();
      const hum = document.getElementById('humidity').value.trim();
      const wind = document.getElementById('wind').value.trim();
      if (temp && !Number.isFinite(Number(temp))) { alert('溫度必須是數字'); return false; }
      if (hum && (!Number.isFinite(Number(hum)) || Number(hum) < 0 || Number(hum) > 100)) { alert('濕度必須介於 0 至 100'); return false; }
      if (wind && (!Number.isFinite(Number(wind)) || Number(wind) < 0)) { alert('風速必須是大於或等於 0 的數字'); return false; }

      return true;
    }

    async function submitForm() {
      if (isSubmitting) return;
      if (reportArchived) { alert('這份日報已歸檔，請勿重複送出。'); return; }
      if (!validateProject()) return;
      if (!validateForm()) return;

      const submitButton = document.getElementById('submitButton');
      isSubmitting = true;
      submitButton.disabled = true;
      submitButton.textContent = '正在送出...';

      // 💡 送出成功後，永久記憶這次填寫的名字
      const reporterName = document.getElementById('reporterName').value.trim();
      localStorage.setItem('lastReporterName', reporterName);

      let finalContractorStr, finalWorkerCountStr, finalProgressStr, finalMaterialStr;
      
      const isNoWork = document.getElementById('noWorkToday').checked;
      if (isNoWork) {
          const reason = document.getElementById('noWorkReason').value;
          finalContractorStr = '無出工';
          finalWorkerCountStr = '0';
          finalProgressStr = `今日無出工 (停工/休息)\n原因：${reason}`;
          finalMaterialStr = '無';
      } else {
          saveNewContractorsFromForm();

          let totalWorkers = 0;
          const contractorList = [];
          document.querySelectorAll('#contractorsContainer .contractor-group').forEach(group => {
              const select = group.querySelector('.contractor-name').value;
              const name = select === '其他' ? group.querySelector('.custom-contractor-name').value.trim() : select;
              const workers = Number(group.querySelector('.contractor-workers').value);
              contractorList.push(`• ${name}（${workers}人）`);
              totalWorkers += workers;
          });

          finalContractorStr = '\n' + contractorList.join('\n');
          finalWorkerCountStr = `共 ${totalWorkers} 人`;
          finalProgressStr = buildProgressString();
          finalMaterialStr = buildMaterialString();
      }

      const contractorItems = buildContractorItems();
      const workItems = buildWorkItems();
      const materialItems = buildMaterialItems();
      const totalWorkerCount = contractorItems.reduce((total, item) => total + item.workerCount, 0);

      const submissionIdStr = ensureSubmissionId();

      const payload = {
        projectId: document.getElementById('projectId').value.trim(),
        projectName: getSelectedProjectName(),
        submissionId: submissionIdStr,
        reporterName: reporterName, // 💡 把填表人包進去送給伺服器
        
        temp: document.getElementById('temp').value || '未填寫',
        humidity: document.getElementById('humidity').value || '未填寫',
        wind: document.getElementById('wind').value || '未填寫',
        contractor: finalContractorStr,
        workerCount: finalWorkerCountStr,
        progress: finalProgressStr,
        materials: finalMaterialStr,
        remarks: document.getElementById('remarks').value.trim() || '無',
        
        isNoWork,
        noWorkReason: isNoWork ? document.getElementById('noWorkReason').value : '',
        contractorItems,
        totalWorkerCount,
        workItems,
        customWorkItem: isNoWork ? '' : document.getElementById('customWorkItem').value.trim(),
        workNotes: isNoWork ? '' : document.getElementById('workNotes').value.trim(),
        materialItems
      };

      try {
        const response = await fetch(API_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        
        const responseText = await response.text();
        let result;
        try { result = responseText ? JSON.parse(responseText) : {}; } catch { throw new Error(`HTTP ${response.status}`); }
        if (!response.ok || result.success === false) throw new Error(result.error || result.message || `系統處理失敗`);

        if (result.archived) reportArchived = true;
        if (result.archived && result.pushed) {
          localStorage.removeItem(getDraftKey());
          alert('日報已歸檔並發布至 LINE 群組');
          if (window.liff && liff.isInClient()) liff.closeWindow();
          return;
        }
        if (result.archived && result.reason === 'PROJECT_NOT_BOUND') {
          submitButton.textContent = '日報已歸檔，尚未發布';
          alert('日報已歸檔至 OneDrive，但案場尚未綁定 LINE。');
          return;
        }
        if (result.archived && result.reason === 'LINE_PUSH_FAILED') {
          submitButton.textContent = '日報已歸檔，發布失敗';
          alert('日報已歸檔至 OneDrive，但 LINE 發布失敗。');
          return;
        }
        alert(result.message || '日報處理完成');
      } catch (error) {
        alert(`傳送失敗：${error.message}`);
      } finally {
        isSubmitting = false;
        if (!reportArchived) { submitButton.disabled = false; submitButton.textContent = '確認送出並發布至群組'; }
      }
    }
  </script>
</body>
</html>
