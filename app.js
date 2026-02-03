/* app.js */

// ✅ 正確讀法
const ENDPOINT = window.GAS_ENDPOINT;

const $ = (id) => document.getElementById(id);
const statusEl = $("status");
const whoEl = $("who");
const locEl = $("loc");

function setStatus(msg, ok) {
  statusEl.innerHTML = msg;
  statusEl.classList.remove("ok", "bad");
  if (ok === true) statusEl.classList.add("ok");
  if (ok === false) statusEl.classList.add("bad");
}

function safeText(s){
  return String(s||'').replace(/[<>&]/g,c=>({
    '<':'&lt;','>':'&gt;','&':'&amp;'
  }[c]));
}

function jsonp(url, params, timeoutMs = 12000){
  return new Promise((resolve)=>{
    const cb = "cb_" + Math.random().toString(16).slice(2);
    let done = false;

    function finish(d){
      if(done) return;
      done = true;
      resolve(d);
      cleanup();
    }

    window[cb] = (d)=>finish(d);

    const qs = new URLSearchParams({ ...params, callback: cb });
    const sc = document.createElement("script");
    sc.src = url + (url.includes("?") ? "&" : "?") + qs.toString();
    sc.async = true;
    sc.onerror = ()=>finish({ ok:false, error:"jsonp_network_error" });

    const t = setTimeout(()=>finish({ ok:false, error:"jsonp_timeout" }), timeoutMs);

    function cleanup(){
      clearTimeout(t);
      try{ delete window[cb]; }catch{ window[cb]=undefined; }
      sc.remove();
    }

    document.body.appendChild(sc);
  });
}

async function postOrJsonp(payload){
  if(!ENDPOINT) return { ok:false, error:"missing_endpoint" };

  try{
    const r = await fetch(ENDPOINT,{
      method:"POST",
      headers:{ "Content-Type":"application/json" },
      body: JSON.stringify(payload)
    });
    const txt = await r.text();
    try{ return JSON.parse(txt); }
    catch{ return { ok:r.ok }; }
  }catch{
    return await jsonp(ENDPOINT,{
      action:"submit",
      ...payload
    });
  }
}

function getUser(){
  return {
    userId: localStorage.getItem("userId") || "demo_user",
    displayName: localStorage.getItem("displayName") || "Demo"
  };
}

async function submitRecord({ type, requireGps, note, dataObj }){
  if(!ENDPOINT){
    setStatus("❌ GAS_ENDPOINT 未載入", false);
    return;
  }

  const { userId, displayName } = getUser();
  whoEl.textContent = `${displayName} (${userId})`;

  setStatus("送出中…", true);

  const payload = {
    type,
    userId,
    displayName,
    note: note || "",
    data: JSON.stringify(dataObj || {})
  };

  const res = await postOrJsonp(payload);

  if(!res?.ok){
    setStatus(`❌ 送出失敗：${safeText(res?.error)}`, false);
  }else{
    setStatus("✅ 已成功送出", true);
  }
}

function init(){
  if(!ENDPOINT){
    setStatus("❌ 缺少 GAS_ENDPOINT（config.js）", false);
    return;
  }

  setStatus("⏳ 系統初始化中…", true);

  jsonp(ENDPOINT,{ action:"ping" }).then(r=>{
    if(r?.ok) setStatus("✅ 系統就緒", true);
    else setStatus("⚠️ GAS 無回應", false);
  });

  $("btnClockIn")?.addEventListener("click", ()=>submitRecord({
    type:"clockin", requireGps:true
  }));

  $("btnClockOut")?.addEventListener("click", ()=>submitRecord({
    type:"clockout", requireGps:true
  }));
}

init();
