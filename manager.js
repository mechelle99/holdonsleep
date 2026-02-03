/* manager.js */

// ✅ 統一：只用 window.GAS_ENDPOINT
if (!window.GAS_ENDPOINT) {
  throw new Error("❌ GAS_ENDPOINT not loaded. Did you include config.js first?");
}

const ENDPOINT = window.GAS_ENDPOINT;
const $ = (id) => document.getElementById(id);
const statusEl = $("status");
const listEl = $("list");
const midEl = $("mid");

function setStatus(msg, ok){
  statusEl.innerHTML = msg;
  statusEl.classList.remove("ok","bad");
  if(ok===true) statusEl.classList.add("ok");
  if(ok===false) statusEl.classList.add("bad");
}

function safeText(s){
  return String(s||'').replace(/[<>&]/g,c=>({
    '<':'&lt;',
    '>':'&gt;',
    '&':'&amp;'
  }[c]));
}

// ===== JSONP =====
function jsonp(url, params, timeoutMs=12000){
  return new Promise((resolve)=>{
    const cb="cb_"+Math.random().toString(16).slice(2);
    let done=false;

    function finish(d){
      if(done) return;
      done=true;
      resolve(d);
      cleanup();
    }

    window[cb]=(d)=>finish(d);

    const qs=new URLSearchParams({ ...params, callback: cb });
    const sc=document.createElement("script");
    sc.src=url+(url.includes("?")?"&":"?")+qs.toString();
    sc.async=true;
    sc.onerror=()=>finish({ok:false,error:"jsonp_network_error"});

    const t=setTimeout(()=>finish({ok:false,error:"jsonp_timeout"}), timeoutMs);

    function cleanup(){
      clearTimeout(t);
      try{ delete window[cb]; }catch(_){ window[cb]=undefined; }
      if(sc.parentNode) sc.parentNode.removeChild(sc);
    }

    document.body.appendChild(sc);
  });
}

// ===== Manager 身分（暫用 localStorage）=====
function getManagerId(){
  return localStorage.getItem("managerId")
      || localStorage.getItem("userId")
      || "demo_manager";
}

// ===== 載入待核准 =====
async function loadPending(){
  const managerId = getManagerId();
  midEl.textContent = managerId;

  setStatus("載入待核准清單中…", true);

  const res = await jsonp(ENDPOINT, {
    action:"list_outing_pending",
    managerId
  });

  if(!res?.ok){
    setStatus(
      `❌ 載入失敗：<span class="mono">${safeText(res?.error || "unknown")}</span><br>
       （若顯示 not_manager，請確認 GAS 裡 MANAGER_USER_IDS）`,
      false
    );
    listEl.innerHTML="";
    return;
  }

  const items = res.items || [];
  if(items.length===0){
    setStatus("✅ 目前沒有待核准外出單", true);
    listEl.innerHTML="";
    return;
  }

  setStatus(`✅ 待核准 ${items.length} 筆`, true);
  renderList(items, managerId);
}

// ===== 畫面 =====
function renderList(items, managerId){
  listEl.innerHTML="";

  items.forEach(it=>{
    const div=document.createElement("div");
    div.className="item";
    div.innerHTML=`
      <div class="title">${safeText(it.displayName)}（${safeText(it.userId)}）</div>
      <div class="meta">
        單號：<span class="mono">${safeText(it.requestId)}</span><br>
        時段：${safeText(it.date)} ${safeText(it.start)} - ${safeText(it.end)}<br>
        目的地：${safeText(it.destination)}<br>
        原因：${safeText(it.reason)}
      </div>

      <textarea placeholder="主管備註（可留空）" data-note="${safeText(it.requestId)}"></textarea>

      <div class="row">
        <button data-act="APPROVED" data-id="${safeText(it.requestId)}">核准</button>
        <button class="secondary" data-act="REJECTED" data-id="${safeText(it.requestId)}">駁回</button>
      </div>
    `;
    listEl.appendChild(div);
  });

  listEl.querySelectorAll("button[data-id]").forEach(btn=>{
    btn.addEventListener("click", async ()=>{
      const requestId = btn.dataset.id;
      const decision = btn.dataset.act;
      const ta = listEl.querySelector(`textarea[data-note="${CSS.escape(requestId)}"]`);
      const managerNote = ta ? ta.value.trim() : "";

      btn.disabled=true;
      setStatus(`送出 ${decision} 中…`, true);

      const res = await jsonp(ENDPOINT,{
        action:"approve_outing",
        managerId,
        requestId,
        decision,
        managerNote
      });

      if(!res?.ok){
        setStatus(`❌ 操作失敗：${safeText(res?.error||"unknown")}`, false);
        btn.disabled=false;
        return;
      }

      setStatus(`✅ 已${decision==="APPROVED"?"核准":"駁回"}`, true);
      loadPending();
    });
  });
}

// ===== 事件 =====
$("btnRefresh")?.addEventListener("click",(e)=>{
  e.preventDefault();
  loadPending();
});

// ===== 啟動 =====
loadPending();
