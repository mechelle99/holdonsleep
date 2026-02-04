/* app.js - 自動扣除午休 1小時版 */
const ENDPOINT = window.CONFIG?.GAS_ENDPOINT || window.GAS_ENDPOINT;
const $ = (id) => document.getElementById(id);
const statusEl = $("status");
const whoEl = $("who");
const locEl = $("loc");

// 通訊 API
async function callApi(payload) {
  if (!ENDPOINT) throw new Error("缺少 GAS_ENDPOINT");
  const res = await fetch(ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  const txt = await res.text();
  try { return JSON.parse(txt); } 
  catch (e) { throw new Error("伺服器回傳格式錯誤"); }
}

function setStatus(msg, ok) {
  statusEl.innerHTML = msg;
  statusEl.className = "status " + (ok ? "ok" : "bad");
  statusEl.style.display = "block";
  setTimeout(() => { statusEl.style.display = "none"; }, 3000);
}

function getUser() {
  return { 
    userId: localStorage.getItem("employeeId"), 
    displayName: localStorage.getItem("employeeName") 
  };
}

window.logout = function() {
  if(confirm("確定要登出嗎？")) {
    localStorage.removeItem("employeeId");
    localStorage.removeItem("employeeName");
    location.href = "login.html";
  }
}

// 載入儀表板數據
async function loadDashboard() {
  const { userId, displayName } = getUser();
  if (!userId) return;
  $("dispAnnualLeft").textContent = "...";
  $("dispCompLeft").textContent = "...";
  try {
    const res = await callApi({ action: "get_dashboard", userId, displayName });
    if (res.ok && res.data) {
      $("dispAnnualLeft").textContent = res.data.annual.left + " 天";
      $("dispAnnualTotal").textContent = res.data.annual.total;
      $("dispAnnualUsed").textContent = res.data.annual.used;
      $("dispCompLeft").textContent = res.data.comp.left + " 時";
      $("dispCompTotal").textContent = res.data.comp.total;
      $("dispCompUsed").textContent = res.data.comp.used;
    }
  } catch (e) { console.error("儀表板錯誤", e); }
}

function getLocation(force) {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      if (force) return reject(new Error("瀏覽器不支援定位"));
      return resolve({ lat: "", lng: "" });
    }
    const options = { enableHighAccuracy: false, timeout: 20000, maximumAge: 60000 };
    navigator.geolocation.getCurrentPosition(
      (pos) => resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
      (err) => {
        if (!force) return resolve({ lat: "", lng: "" }); 
        reject(new Error("定位失敗 (請確認權限或訊號)"));
      },
      options
    );
  });
}

function showPanel(type) {
  ["panelClock", "panelOuting", "panelLeave", "panelOvertime"].forEach(id => {
    $(id).style.display = "none";
  });
  if (type === "clock") { $("panelClock").style.display = "block"; locEl.textContent = "需定位"; }
  else if (type === "outing") { $("panelOuting").style.display = "block"; locEl.textContent = "免定位"; }
  else if (type === "leave") { $("panelLeave").style.display = "block"; locEl.textContent = "免定位"; }
  else if (type === "overtime") { $("panelOvertime").style.display = "block"; locEl.textContent = "免定位"; }
}

// --- 💡 修改重點：自動扣除午休 1 小時 ---
window.calcLeaveHours = function() {
  const s = $("leaveStart").value;
  const e = $("leaveEnd").value;
  if (!s || !e) return;
  
  const start = new Date(s);
  const end = new Date(e);
  
  if (end <= start) { 
    alert("結束時間不能早於開始時間"); 
    $("leaveEnd").value = ""; 
    return; 
  }

  // 1. 算出原始時數
  let hours = (end - start) / (36e5); // 毫秒轉小時
  
  // 2. 自動扣除午休規則
  // 如果時數超過 4 小時 (代表跨越上午下午)，我們假設有午休，自動扣 1 小時
  // 例如：09:00 ~ 18:00 = 原始9小時 -> 自動變 8 小時
  if (hours > 4) {
    hours = hours - 1;
  }
  
  // 顯示結果
  $("leaveTotalHours").textContent = hours.toFixed(1);
};

// 加班時數計算 (加班通常是下班後，所以不扣午休，維持原樣)
window.calcOtHours = function() {
  const d = $("otDate").value, s = $("otStart").value, e = $("otEnd").value;
  if (!d || !s || !e) return;
  const start = new Date(`${d}T${s}`), end = new Date(`${d}T${e}`);
  if (end <= start) { alert("結束錯誤"); $("otEnd").value=""; return; }
  let h = (end - start)/(36e5);
  // 加班如果超過 4 小時通常也有休息，看你們規定，目前先不扣
  $("otTotalHours").textContent = (Math.floor(h * 2) / 2).toFixed(1);
};

window.calcOutingHours = function() {
  const s = $("outStart").value, e = $("outEnd").value;
  if (!s || !e) return;
  const today = new Date().toISOString().split('T')[0];
  const start = new Date(`${today}T${s}`), end = new Date(`${today}T${e}`);
  if (end <= start) { alert("結束錯誤"); $("outEnd").value=""; return; }
  let h = (end - start)/(36e5);
  // 外出如果含午休也要扣嗎？通常外出比較彈性，這裡先設為扣除
  if (h > 4) h = h - 1; 
  $("outTotalHours").textContent = h.toFixed(1);
};

async function submitRecord({ action, dataObj, requireGps }) {
  const { userId, displayName } = getUser();
  if (!userId) { location.href = "login.html"; return; }
  const buttons = document.querySelectorAll("button");
  buttons.forEach(b => b.disabled = true);
  setStatus("處理中...", true);

  try {
    let gps = { lat: "", lng: "" };
    if (requireGps) {
      setStatus("📡 定位中...", true);
      try { gps = await getLocation(true); } catch (e) { throw e; }
    }

    setStatus("送出資料...", true);
    const payload = { action, userId, displayName, lat: gps.lat, lng: gps.lng, data: dataObj };
    const res = await callApi(payload);
    
    if (res.ok) {
      setStatus(`✅ ${res.message}`, true);
      if (action.includes("clock")) alert(`打卡成功！${new Date().toTimeString().slice(0,5)}`);
      if (action.includes("create")) {
        $("leaveReason").value=""; $("otReason").value=""; 
        await loadDashboard(); 
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
    } else { sel.innerHTML = "<option>無單據</option>"; }
  } catch(e) {}
}

function bindEvents() {
  $("actionType").addEventListener("change", (e) => showPanel(e.target.value));
  $("btnClockIn").onclick = () => submitRecord({ action: "clock_in", requireGps: true, dataObj: {} });
  $("btnClockOut").onclick = () => submitRecord({ action: "clock_out", requireGps: true, dataObj: {} });
  
  $("btnOutApply").onclick = () => {
    if($("outTotalHours").textContent === "0.0") return alert("請確認時間");
    const d=$("outDate").value;
    submitRecord({ action: "create_outing", requireGps: false, dataObj: {
      start_full: `${d} ${$("outStart").
