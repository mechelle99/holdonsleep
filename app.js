/* app.js (Updated: Fetch API version) */
const ENDPOINT = window.CONFIG?.GAS_ENDPOINT || window.GAS_ENDPOINT;

const $ = (id) => document.getElementById(id);
const statusEl = $("status");
const whoEl = $("who");
const locEl = $("loc");

// --- 核心通訊函式 (取代原本的 jsonp) ---
async function callApi(payload) {
  if (!ENDPOINT) throw new Error("缺少 GAS_ENDPOINT");
  
  // 使用 fetch POST，這是跨網域最標準的做法
  const res = await fetch(ENDPOINT, {
    method: "POST",
    // 必須用 text/plain 以避免 Google 的 CORS 預檢 (preflight) 失敗
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  
  const txt = await res.text();
  try {
    return JSON.parse(txt);
  } catch (e) {
    throw new Error("伺服器回傳格式錯誤: " + txt);
  }
}

function setStatus(msg, ok) {
  statusEl.innerHTML = msg;
  statusEl.className = "status " + (ok ? "ok" : "bad");
}

// --- 使用者與定位 ---
function getUser() {
  const userId = localStorage.getItem("employeeId");
  const displayName = localStorage.getItem("employeeName");
  return { userId, displayName };
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

// --- UI 切換 ---
function showPanel(type) {
  ["panelClock", "panelOuting", "panelLeave", "panelOvertime"].forEach(id => {
    $(id).style.display = "none";
  });

  if (type === "clock") {
    $("panelClock").style.display = "block";
    locEl.textContent = "需定位 (送出時抓取)";
  } else if (type === "outing") {
    $("panelOuting").style.display = "block";
    locEl.textContent = "外出申請免定位 / 外出打卡需定位";
  } else {
    // leave, overtime
    const map = { leave: "panelLeave", overtime: "panelOvertime" };
    $(map[type]).style.display = "block";
    locEl.textContent = "不需定位";
  }
}

// --- 統一送出函式 ---
async function submitRecord({ action, dataObj, requireGps }) {
  const { userId, displayName } = getUser();
  if (!userId) {
    setStatus("❌ 請先至 <a href='setup.html'>setup.html</a> 設定員工身份", false);
    return;
  }

  // 鎖定按鈕避免重複按
  const buttons = document.querySelectorAll("button");
  buttons.forEach(b => b.disabled = true);
  setStatus("送出中...", true);

  try {
    let gps = { lat: "", lng: "" };
    if (requireGps) {
      try {
        gps = await getLocation(true);
      } catch (e) {
        throw new Error("無法取得定位，請確認已授權 GPS。");
      }
    }

    whoEl.textContent = `${displayName} (${userId})`;
    if (requireGps) locEl.textContent = `${gps.lat}, ${gps.lng}`;

    // 呼叫後端
    const payload = {
      action: action, // 告訴後端要做什麼
      userId,
      displayName,
      lat: gps.lat,
      lng: gps.lng,
      data: dataObj
    };

    const res = await callApi(payload);

    if (res.ok) {
      setStatus(`✅ 成功：${res.message || "已送出"}`, true);
      // 如果是外出申請成功，重新整理清單
      if (action === "create_outing") await loadApprovedOutings();
    } else {
      setStatus(`❌ 失敗：${res.message || "未知錯誤"}`, false);
    }

  } catch (err) {
    setStatus(`❌ 錯誤：${err.message}`, false);
  } finally {
    buttons.forEach(b => b.disabled = false);
  }
}

// --- 功能綁定 ---
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

function bindEvents() {
  $("actionType").addEventListener("change", (e) => showPanel(e.target.value));

  // 1. 上下班打卡
  $("btnClockIn").onclick = () => submitRecord({ 
    action: "clock_in", requireGps: true, dataObj: {} 
  });
  $("btnClockOut").onclick = () => submitRecord({ 
    action: "clock_out", requireGps: true, dataObj: {} 
  });

  // 2. 外出申請
  $("btnOutApply").onclick = () => {
    submitRecord({
      action: "create_outing", requireGps: false,
      dataObj: {
        date: $("outDate").value,
        start: $("outStart").value,
        end: $("outEnd").value,
        destination: $("outDest").value,
        reason: $("outReason").value
      }
    });
  };

  // 3. 外出打卡
  const getOutReq = () => ({ requestId: $("approvedOutingSelect").value });
  $("btnOutIn").onclick = () => submitRecord({ 
    action: "clock_in", requireGps: true, dataObj: { ...getOutReq(), isOuting: true } 
  });
  $("btnOutOut").onclick = () => submitRecord({ 
    action: "clock_out", requireGps: true, dataObj: { ...getOutReq(), isOuting: true } 
  });

  // 4. 請假
  $("btnLeaveSubmit").onclick = () => {
    submitRecord({
      action: "create_leave", requireGps: false,
      dataObj: {
        type: $("leaveKind").value,
        start: $("leaveStart").value,
        end: $("leaveEnd").value,
        reason: $("leaveReason").value
      }
    });
  };

  // 5. 加班
  $("btnOtSubmit").onclick = () => {
    submitRecord({
      action: "create_ot", requireGps: false,
      dataObj: {
        date: $("otDate").value,
        hours: $("otHours").value,
        reason: $("otReason").value
      }
    });
  };
}

// 初始化
function init() {
  if (!ENDPOINT) return setStatus("❌ 未設定 GAS_ENDPOINT", false);
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
