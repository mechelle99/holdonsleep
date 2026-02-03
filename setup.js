/* setup.js */
const $ = (id) => document.getElementById(id);

function setStatus(msg, ok){
  const el = $("status");
  el.classList.remove("ok","bad");
  if(ok===true) el.classList.add("ok");
  if(ok===false) el.classList.add("bad");
  el.innerHTML = msg;
}
function safe(s){ return String(s||"").replace(/[<>&]/g,c=>({ "<":"&lt;",">":"&gt;","&":"&amp;" }[c])); }

function load(){
  const employeeId = localStorage.getItem("employeeId") || "";
  const employeeName = localStorage.getItem("employeeName") || "";
  const managerId = localStorage.getItem("managerId") || "";

  $("employeeId").value = employeeId;
  $("employeeName").value = employeeName;
  $("managerId").value = managerId;

  setStatus(
    `目前設定：<br/>
     - employeeId：<span class="mono">${safe(employeeId)||"(未設定)"}</span><br/>
     - employeeName：<span class="mono">${safe(employeeName)||"(未設定)"}</span><br/>
     - managerId：<span class="mono">${safe(managerId)||"(未設定)"}</span>`,
    true
  );
}

function saveEmployee(){
  const id = $("employeeId").value.trim();
  const name = $("employeeName").value.trim();
  if(!id) return setStatus("❌ employeeId 必填", false);
  if(!name) return setStatus("❌ employeeName 必填", false);

  localStorage.setItem("employeeId", id);
  localStorage.setItem("employeeName", name);

  setStatus(`✅ 已儲存員工身份：<span class="mono">${safe(name)} (${safe(id)})</span>`, true);
}

function saveManager(){
  const mid = $("managerId").value.trim();
  if(!mid){
    localStorage.removeItem("managerId");
    return setStatus("✅ 已清除主管身份（managerId）", true);
  }
  localStorage.setItem("managerId", mid);
  setStatus(`✅ 已儲存主管身份：<span class="mono">${safe(mid)}</span><br/>（記得 GAS 的 MANAGER_USER_IDS 要包含它）`, true);
}

function clearAll(){
  localStorage.removeItem("employeeId");
  localStorage.removeItem("employeeName");
  localStorage.removeItem("managerId");
  load();
  setStatus("✅ 已清除全部身份設定", true);
}

$("btnSaveEmployee").addEventListener("click",(e)=>{ e.preventDefault(); saveEmployee(); }, {passive:false});
$("btnSaveManager").addEventListener("click",(e)=>{ e.preventDefault(); saveManager(); }, {passive:false});
$("btnLoad").addEventListener("click",(e)=>{ e.preventDefault(); load(); }, {passive:false});
$("btnClear").addEventListener("click",(e)=>{ e.preventDefault(); clearAll(); }, {passive:false});

load();
