/* app.js 完整修正版 - 解決語法錯誤與重複定義 */
const ENDPOINT = window.CONFIG?.GAS_ENDPOINT || window.GAS_ENDPOINT;
const $ = (id) => document.getElementById(id);
const statusEl = $("status");
const whoEl = $("who");
const locEl = $("loc");

// --- 1. 核心通訊 API ---
async function callApi(payload) {
  if (!ENDPOINT) throw new Error("缺少 GAS_ENDPOINT");
  const res = await fetch(ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  const txt = await res.text();
  try { return JSON.parse(txt); } 
  catch (e) { throw new Error("伺服器回傳格式錯誤: " + txt); }
}

function setStatus(msg, ok) {
  statusEl.innerHTML = msg;
  statusEl.className = "status " + (ok ? "ok" : "bad");
}

function getUser() {
  return { 
    userId: localStorage.getItem("employeeId"), 
    displayName: localStorage.getItem("employeeName") 
  };
}

function getLocation(force) {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      if (force) return reject(new Error("瀏覽器不支援定位"));
      return resolve({ lat: "", lng: "" });
    }
    navigator.geolocation.getCurrentPosition(
      (pos) => resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
      (err) => (force ? reject(err) : resolve({ lat: "", lng: "" })),
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 }
    );
  });
}

function showPanel(type) {
  ["panelClock", "panelOuting", "panelLeave", "panelOvertime"].forEach(id => {
    $(id).style.display = "none";
  });
  if (type === "clock") { $("panelClock").style.display = "block"; locEl.textContent = "需定位 (送出時抓取)"; }
  else if (type === "outing") { $("panelOuting").style.display = "block"; locEl.textContent = "申請免定位 / 打卡需定位"; }
  else if (type === "leave") { $("panelLeave").style.display = "block"; locEl.textContent = "免定位"; }
  else if (type === "overtime") { $("panelOvertime").style.display = "block"; locEl.textContent = "免定位"; }
}

// --- 2. 自動計算時數函式 (綁定在 window 以便 HTML onchange 呼叫) ---

// 計算請假時數
window.calcLeaveHours = function() {
  const s = $("leaveStart").value;
  const e = $("leaveEnd").value;
  if (!s || !e) return;
  const start = new Date(s), end = new Date(e);
  
  if (end <= start) { 
    alert("❌ 結束時間不能早於或等於開始時間！"); 
    $("leaveEnd").value = ""; 
    $("leaveTotalHours").textContent = "0.0";
    return; 
  }
  
  const diff = (end - start) / (36e5); // 毫秒轉小時
  $("leaveTotalHours").textContent = diff.toFixed(1);
};

// 計算加班時數 (0.5小時為單位)
window.calcOtHours = function() {
  const d = $("otDate").value;
  const s = $("otStart").value;
  const e = $("otEnd").value;
  if (!s || !e) return;
  
  // 為了計算方便，假設同一天 (跨日需另外處理或填兩張單)
  const baseDate = d || new Date().toISOString().split('T')[0];
  const start = new Date(`${baseDate}T${s}`);
  const end = new Date(`${baseDate}T${e}`);

  if (end <= start) { 
    alert("❌ 加班結束時間不能早於開始時間"); 
    $("otEnd").value = ""; 
    $("otTotalHours").textContent = "0.0";
    return; 
  }
  
  let h = (end - start) / (36e5);
  // 無條件捨去到 0.5 (例如 1.8 -> 1.5)
  $("otTotalHours").textContent = (Math.floor(h * 2) / 2).toFixed(1);
};

// 計算外出時數
window.calcOutingHours = function() {
  const d = $("outDate").value;
  const s = $("outStart").value;
  const e = $("outEnd").value;
  if (!s || !e) return;

  const baseDate = d || new Date().toISOString().split('T')[0];
  const start = new Date(`${baseDate}T${s}`);
  const end = new Date(`${baseDate}T${e}`);

  if (end <= start) { 
    alert("❌ 結束時間不能早於開始時間"); 
    $("outEnd").value = ""; 
    $("outTotalHours").textContent = "0.0";
    return; 
  }
  
  $("outTotalHours").textContent = ((end - start) / (36e5)).toFixed(1);
};

// --- 3. 統一送出資料邏輯 ---
async function submitRecord({ action, dataObj, requireGps }) {
  const { userId, displayName } = getUser();
  if (!userId) {
    setStatus("❌ 請先至 <a href='setup.html'>setup.html</a> 設定員工身份", false);
    return;
  }

  const buttons = document.querySelectorAll("button");
  buttons.forEach(b => b.disabled = true);
  setStatus("送出中...", true);

  try {
    let gps = { lat: "", lng: "" };
    if (requireGps) {
      try { gps = await getLocation(true); } 
      catch (e) { throw new Error("無法取得定位，請確認已授權 GPS。"); }
    }

    whoEl.textContent = `${displayName} (${userId})`;
    if (requireGps) locEl.textContent = `${gps.lat}, ${gps.lng}`;

    const payload = {
      action: action,
      userId,
      displayName,
      lat: gps.lat,
      lng: gps.lng,
      data: dataObj
    };

    const res = await callApi(payload);

    if (res.ok) {
      setStatus(`✅ 成功：${res.message}`, true);
      
      // 打卡成功跳出時間提示
      if (action.includes("clock")) {
        const now = new Date();
        const timeStr = now.getHours().toString().padStart(2,'0') + ":" + now.getMinutes().toString().padStart(2,'0');
        alert(`打卡成功！時間：${timeStr}`);
      }
      
      // 如果是外出申請，清空欄位
      if (action === "create_outing") {
        $("outDest").value = ""; $("outReason").value = "";
        // 重新整理外出單清單
        await loadApprovedOutings();
      }
      
    } else {
      setStatus(`❌ 失敗：${res.message}`, false);
    }

  } catch (err) {
    setStatus(`❌ 錯誤：${err.message}`, false);
  } finally {
    buttons.forEach(b => b.disabled = false);
  }
}

// 讀取我的外出單 (用於下拉選單)
async function loadApprovedOutings() {
  const { userId } = getUser();
  if(!userId) return;
  try {
    const res = await callApi({ action: "get_my_outings", userId });
    const sel = $("approvedOutingSelect");
    sel.innerHTML = "";
    if (res.ok && res.list && res.list.length > 0) {
      res.list.forEach(item => {
        const opt = document.createElement("option");
        opt.value = item.requestId;
        opt.textContent = `${item.date} ${item.destination} (${item.status})`;
        sel.appendChild(opt);
      });
    } else {
      sel.innerHTML = "<option>無已核准的外出單</option>";
    }
  } catch(e) { console.error(e); }
}

// --- 4. 事件綁定 (bindEvents) ---
function bindEvents() {
  // 切換面板
  $("actionType").addEventListener("change", (e) => showPanel(e.target.value));

  // A. 上下班打卡
  const handleClock = (action) => {
    submitRecord({ action, requireGps: true, dataObj: {} });
  };
  $("btnClockIn").onclick = () => handleClock("clock_in");
  $("btnClockOut").onclick = () => handleClock("clock_out");

  // B. 外出申請送出
  $("btnOutApply").onclick = () => {
    if($("outTotalHours").textContent === "0.0") return alert("請確認時間與時數");
    if(!$("outDate").value) return alert("請選擇日期");
    if(!$("outDest").value) return alert("請填寫目的地");
    if(!$("outReason").value) return alert("請填寫原因");

    const d = $("outDate").value;
    submitRecord({
      action: "create_outing", 
      requireGps: false,
      dataObj: {
        start_full: `${d} ${$("outStart").value}`, 
        end_full: `${d} ${$("outEnd").value}`,
        hours: $("outTotalHours").textContent,
        destination: $("outDest").value,
        reason: $("outReason").value
      }
    });
  };

  // C. 外出打卡 (需選單號)
  const getOutReq = () => ({ requestId: $("approvedOutingSelect").value });
  $("btnOutIn").onclick = () => submitRecord({ action: "clock_in", requireGps: true, dataObj: { ...getOutReq(), isOuting: true } });
  $("btnOutOut").onclick = () => submitRecord({ action: "clock_out", requireGps: true, dataObj: { ...getOutReq(), isOuting: true } });

  // D. 請假送出
  $("btnLeaveSubmit").onclick = () => {
    if($("leaveTotalHours").textContent === "0.0") return alert("請確認時間與時數");
    if(!$("leaveReason").value) return alert("請填寫原因");
    
    submitRecord({
      action: "create_leave", requireGps: false,
      dataObj: {
        type: $("leaveKind").value,
        start: $("leaveStart").value.replace("T", " "), // 去掉 T 讓格式好看
        end: $("leaveEnd").value.replace("T", " "),
        hours: $("leaveTotalHours").textContent,
        reason: $("leaveReason").value
      }
    });
  };

  // E. 加班送出
  $("btnOtSubmit").onclick = () => {
    if($("otTotalHours").textContent === "0.0") return alert("請確認時間與時數");
    if(!$("otDate").value) return alert("請填寫日期");
    if(!$("otReason").value) return alert("請填寫事由");

    const d = $("otDate").value;
    submitRecord({
      action: "create_ot", requireGps: false,
      dataObj: {
        start_full: `${d} ${$("otStart").value}`, 
        end_full: `${d} ${$("otEnd").value}`,
        hours: $("otTotalHours").textContent,
        reason: $("otReason").value
      }
    });
  };
}

// --- 5. 初始化 ---
function init() {
  if (!ENDPOINT) return setStatus("❌ 未設定 GAS_ENDPOINT (請檢查 config.js)", false);
  
  const user = getUser();
  if (!user.userId) {
    setStatus("⚠️ 請先執行 <a href='setup.html'>setup.html</a> 設定身份", false);
  } else {
    whoEl.textContent = `${user.displayName} (${user.userId})`;
    setStatus("✅ 系統就緒", true);
  }
  
  showPanel($("actionType").value);
  bindEvents();
  loadApprovedOutings();
}

init();
