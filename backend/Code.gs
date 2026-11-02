/*************************************************
 * StuDy v2 — Google Apps Script backend
 * Data lives in the attached Google Sheet with tabs:
 *   Users | Attempts | Payments | Settings | Notifications | OTPs
 * (Run setup() once from the editor to create them.)
 * Deploy: Deploy → New deployment → Web app → Execute as: Me,
 * Access: Anyone. Paste the Web App URL into BACKEND_URL in index.html.
 * Admins: ADMIN_PHONES below; each admin's password is whatever that account
 * was registered with. Register that number normally first.
 *************************************************/
const ADMIN_PHONES = ["9588293680", "9628187101"]; // owner + partner (secondary admin)
const SHEET_NAMES = ["Users","Attempts","Payments","Settings","Notifications","OTPs","UserNotifs","Cards","Claims"];

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name){ let s = ss().getSheetByName(name); if(!s){ s = ss().insertSheet(name); } return s; }
function setup(){
  const headers = {
    Users:["phone","name","gmail","dob","pass","stars","lang","purchases","ledger","seen","createdAt","session","theme","subUntil"],
    Attempts:["id","phone","cat","sub","type","score","total","right","wrong","skip","first","at","timeUsed"],
    Payments:["id","phone","amount","utr","status","at","insta"],
    UserNotifs:["phone","title","body","at","link"],
    Cards:["code","phone","name","points","tests","avg","acc","rank","at"],
    Settings:["key","json"],
    Notifications:["id","title","body","at"],
    OTPs:["gmail","code","exp","phone"],
    Claims:["id","phone","amount","kind","at","used"]
  };
  SHEET_NAMES.forEach(n=>{ const s=sheet(n); if(s.getLastRow()===0) s.appendRow(headers[n]); });
}
function rows(name){
  const s = sheet(name); const v = s.getDataRange().getValues();
  if(v.length<2 && s.getLastRow()===0) return [];
  const head = v[0]; return v.slice(1).map(r=>{ const o={}; head.forEach((h,i)=>o[h]=r[i]); return o; });
}
function findUser(phone){ return rows("Users").find(u=>String(u.phone)===String(phone)); }
function normDob(d){
  if(d===null || d===undefined || d==="") return "";
  if(Object.prototype.toString.call(d)==="[object Date]") return Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
  let s = String(d).trim();
  if(s.charAt(0)==='"' && s.charAt(s.length-1)==='"'){ try{ s = JSON.parse(s); }catch(e){ s = s.slice(1,-1); } }
  s = String(s).trim();
  if(s.indexOf("T")>0){ const dt = new Date(s); if(!isNaN(dt)) return Utilities.formatDate(dt, "Asia/Kolkata", "yyyy-MM-dd"); }
  const m = s.match(/^\d{4}-\d{2}-\d{2}/);
  return m ? m[0] : s;
}
function saveUser(u){
  const s = sheet("Users"); const v = s.getDataRange().getValues(); const head = v[0];
  const rowOf = o => head.map(h=>{ let val = (h==="dob") ? normDob(o[h]) : o[h]; if(typeof val==="object") val=JSON.stringify(val); return val===undefined?"":val; });
  for(let i=1;i<v.length;i++){ if(String(v[i][0])===String(u.phone)){ const r=rowOf(u); r.forEach((val,j)=> s.getRange(i+1,j+1).setValue(val)); return; } }
  s.appendRow(rowOf(u));
}
function newSession(){ return "s" + Date.now().toString(36) + Math.floor(Math.random()*1e12).toString(36); }
function pubUser(u){
  let p = {tests:{},mocks:{},notes:{}}; try{ const x = JSON.parse(u.purchases||"{}"); if(x && x.tests) p = x; }catch(e){}
  let led = []; try{ led = JSON.parse(u.ledger||"[]"); }catch(e){}
  return { phone:String(u.phone), name:u.name||"", gmail:u.gmail, dob:normDob(u.dob), stars:+u.stars||0, lang:u.lang||"en", purchases:p, ledger:led, session:u.session||"", theme:u.theme||"", subUntil:+u.subUntil||0 };
}
function getSettings(){
  const r = rows("Settings").find(x=>x.key==="main");
  if(r) return JSON.parse(r.json);
  return { prices:{test:5,mock:10,notes:30,bundle:50,sub:1000},
    help:{en:"See Help in the app.",hi:"See Help in the app."}, contact:{en:"Telegram: @dj3680",hi:"Telegram: @dj3680"} };
}
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- Weekly Rank-1 prize: hold #1 for 7 continuous days -> 100 stars auto-credit ---------- */
function istDayStr(offset){
  return Utilities.formatDate(new Date(Date.now() + (offset||0)*86400000), "Asia/Kolkata", "yyyy-MM-dd");
}
function prizeLogSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName("PrizeLog");
  if(!s){ s = ss.insertSheet("PrizeLog"); s.appendRow(["date","topPhone","streak","credited"]); }
  return s;
}
function weeklyTopCalc(){
  const since = Date.now() - 7*24*3600*1000;
  const agg = {};
  rows("Attempts").forEach(r=>{
    if((r.first===true || String(r.first)==="TRUE") && (+r.at||0) >= since){
      const k = String(r.phone);
      agg[k] = agg[k] || {phone:k, score:0, tests:0};
      agg[k].score += (+r.score||0); agg[k].tests += 1;
    }
  });
  const all = Object.values(agg).sort((a,b)=> (b.score - a.score) || (a.tests - b.tests));
  return all.length ? all[0] : null;
}
function dailyPrizeCheck(){
  try{
    const props = PropertiesService.getScriptProperties();
    if(!props.getProperty("prizeTriggerV1")){
      try{ ScriptApp.newTrigger("dailyPrizeCheck").timeBased().everyDays(1).atHour(6).create(); }catch(e){}
      props.setProperty("prizeTriggerV1", "1");
    }
  }catch(e){}
  const top = weeklyTopCalc(); if(!top) return;
  const s = prizeLogSheet();
  const vals = s.getDataRange().getValues();
  const today = istDayStr(0), yest = istDayStr(-1);
  for(let i=1;i<vals.length;i++){ if(String(vals[i][0])===today) return; }
  let streak = 1, credited = "";
  if(vals.length > 1){
    const last = vals[vals.length-1];
    if(String(last[0])===yest && String(last[1])===String(top.phone)){ streak = (+last[2]||0) + 1; credited = String(last[3]||""); }
  }
  if(streak >= 7 && !credited){
    const u = findUser(top.phone);
    if(u){
      u.stars = (+u.stars||0) + 100;
      const led = JSON.parse(u.ledger||"[]");
      led.unshift({at:Date.now(), amt:100, label:"Weekly Rank 1 prize"});
      u.ledger = JSON.stringify(led);
      saveUser(u);
      sheet("UserNotifs").appendRow(["ntf_"+Date.now(), String(top.phone), "You held Rank #1 for a full week! A ₹100 bonus has been credited to your StuDy wallet.", Date.now()]);
      credited = "paid:" + today;
    }
  }
  s.appendRow([today, String(top.phone), streak, credited]);
}
function pruneNotifs(){
  const cut = Date.now() - 5*24*3600*1000;
  [["Notifications",3],["UserNotifs",3]].forEach(function(cfg){
    const s = sheet(cfg[0]); const v = s.getDataRange().getValues();
    for(let i = v.length - 1; i >= 1; i--){ const at = +v[i][cfg[1]] || 0; if(at && at < cut) s.deleteRow(i + 1); }
  });
}
/* Professional OTP email (plain text + HTML), sent under the name "StuDy" */
function otpMail(to, code, kind){
  const isReg = kind === "register";
  const subject = isReg ? "StuDy - Your Registration OTP" : "StuDy - Password Reset OTP";
  const intro = isReg
    ? "Welcome to StuDy! Please use the one-time password (OTP) below to verify your Gmail and complete your registration."
    : "We received a request to reset the password of your StuDy account. Please use the one-time password (OTP) below to set a new password.";
  const ignore = isReg
    ? "If you did not request this OTP, please ignore this email - no account will be created without verification."
    : "If you did not request this OTP, please ignore this email - your password will remain unchanged.";
  const thanks = isReg ? "Thank you for joining StuDy. We wish you the very best for your preparation!" : "Thank you for learning with StuDy.";
  const body = "Hello,\n\n" + intro + "\n\nYour OTP is: " + code + "\n\nThis OTP is valid for 10 minutes.\n\nFor your security, never share this OTP with anyone. Team StuDy will never call, message or email you asking for your OTP.\n\n" + ignore + "\n\n" + thanks + "\n\nWarm regards,\nTeam StuDy\nOwner: Dinesh";
  const html = `<div style="margin:0;padding:24px 12px;background:#f6f4ee;font-family:Arial,Helvetica,sans-serif">
  <div style="max-width:520px;margin:0 auto;background:#ffffff;border:1px solid #e8e2d4;border-radius:14px;overflow:hidden">
    <div style="background:#111111;padding:18px 24px">
      <span style="color:#ffffff;font-size:22px;font-weight:bold">Stu<span style="color:#e5b95c">D</span>y</span>
      <div style="color:#cbbd93;font-size:12px;margin-top:3px">Tests \u2022 Mocks \u2022 Notes - UPSC \u00b7 RPSC \u00b7 SSC \u00b7 NEET \u00b7 JEE \u00b7 REET</div>
    </div>
    <div style="padding:24px;color:#222222;font-size:14.5px;line-height:1.6">
      <p style="margin:0 0 12px">Hello,</p>
      <p style="margin:0 0 6px">${intro}</p>
      <div style="text-align:center;margin:22px 0">
        <span style="display:inline-block;background:#fffbf0;border:1.5px dashed #e5b95c;border-radius:12px;padding:13px 30px;font-size:30px;font-weight:bold;letter-spacing:8px;color:#111111">${code}</span>
      </div>
      <p style="margin:0 0 12px">This OTP is valid for <b>10 minutes</b>.</p>
      <p style="margin:0 0 12px;color:#666666;font-size:13px">For your security, never share this OTP with anyone. Team StuDy will never call, message or email you asking for your OTP.</p>
      <p style="margin:0 0 12px;color:#666666;font-size:13px">${ignore}</p>
      <p style="margin:0">${thanks}</p>
      <p style="margin:20px 0 0">Warm regards,<br><b>Team StuDy</b><br><span style="color:#666666;font-size:13px">Owner: Dinesh</span></p>
    </div>
  </div>
  <div style="max-width:520px;margin:10px auto 0;text-align:center;color:#999999;font-size:11.5px">This is an automated email from the StuDy app. Please do not reply.</div>
</div>`;
  MailApp.sendEmail({to: to, subject: subject, body: body, htmlBody: html, name: "StuDy"});
}

function doPost(e){
  let req; try{ req = JSON.parse(e.postData.contents); }catch(err){ return json({ok:false,error:"badjson"}); }
  if(req && req.update_id !== undefined){ try{ tgHandle_(req); }catch(err){} return json({ok:true}); }
  const a = req.action;
  try{
    if(a==="bootstrap"){
      try{ pruneNotifs(); }catch(e){}
      return json({ok:true, settings:getSettings(), notifs: rows("Notifications").map(n=>({id:String(n.id),title:n.title,body:n.body,at:+n.at})).sort((x,y)=>y.at-x.at) });
    }
    if(a==="sendOtp"){
      if(findUser(req.phone)) return json({ok:false,error:"exists"});
      const code = String(Math.floor(100000+Math.random()*900000));
      sheet("OTPs").appendRow([req.gmail, code, Date.now()+10*60*1000, req.phone]);
      otpMail(req.gmail, code, "register");
      return json({ok:true});
    }
    if(a==="register"){
      if(findUser(req.phone)) return json({ok:false,error:"exists"});
      const ok = rows("OTPs").some(o=>String(o.gmail)===String(req.gmail) && String(o.code)===String(req.code) && +o.exp>Date.now());
      if(!ok) return json({ok:false,error:"badotp"});
      const u = { phone:String(req.phone), name:req.name||"", gmail:req.gmail, dob:req.dob, pass:req.pass, stars:0, lang:"en",
                  purchases: JSON.stringify({tests:{},mocks:{},notes:{}}), ledger:"[]", seen:"[]", createdAt:Date.now(), session:newSession(), theme:"", subUntil:0 };
      saveUser(u);
      return json({ok:true, user:pubUser(u)});
    }
    if(a==="sendResetOtp"){
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      const code = String(Math.floor(100000+Math.random()*900000));
      sheet("OTPs").appendRow([u.gmail, code, Date.now()+10*60*1000, req.phone]);
      otpMail(u.gmail, code, "reset");
      return json({ok:true});
    }
    if(a==="resetPassword"){
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      if(!req.newPass || String(req.newPass).length<6) return json({ok:false,error:"weakpass"});
      const ok = rows("OTPs").some(o=>String(o.gmail)===String(u.gmail) && String(o.code)===String(req.code) && +o.exp>Date.now());
      if(!ok) return json({ok:false,error:"badotp"});
      u.pass = String(req.newPass); u.session = ""; saveUser(u);
      return json({ok:true});
    }
    if(a==="login"){
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      if(String(u.pass)!==String(req.pass)) return json({ok:false,error:"badpass"});
      u.session = newSession(); saveUser(u);
      return json({ok:true, user:pubUser(u)});
    }
    if(a==="me"){ const u = findUser(req.phone); if(!u) return json({ok:false});
      if(u.session){ if(String(req.session||"") !== String(u.session)) return json({ok:false,error:"session"}); }
      else if(req.session){ u.session = String(req.session); saveUser(u); } const un = rows("UserNotifs").filter(n=>String(n.phone)===String(req.phone)).map(n=>({title:n.title, body:n.body, at:+n.at, link:n.link||""})).sort((x,y)=>y.at-x.at); const pu = pubUser(u); pu._unotifs = un; return json({ok:true, user:pu}); }
    if(a==="leaderboard"){
      const names = {}; rows("Users").forEach(u=> names[String(u.phone)] = u.name || String(u.phone));
      const agg = {};
      rows("Attempts").forEach(r=>{ if(r.first===true || String(r.first)==="TRUE"){ const k = String(r.phone); agg[k] = agg[k] || {name: names[k] || k, score:0, tests:0, phone: "XXXXX" + k.slice(-5), _raw:k}; agg[k].score += (+r.score||0); agg[k].tests++; } });
      const all = Object.values(agg).sort((x,y)=>y.score-x.score);
      let me = null;
      if(req.phone){ const i = all.findIndex(r=>r._raw===String(req.phone)); if(i>=0) me = {rank:i+1, score:all[i].score, tests:all[i].tests}; }
      return json({ok:true, rows: all.slice(0,10).map(r=>({name:r.name, score:r.score, tests:r.tests, phone:r.phone})), me});
    }
    if(a==="resetSeen"){
      const u = findUser(req.phone); if(!u) return json({ok:false});
      const drop = {}; (req.ids||[]).forEach(i=>{ drop[i]=1; });
      const seen = JSON.parse(u.seen||"[]").filter(i=>!drop[i]);
      u.seen = JSON.stringify(seen); saveUser(u); return json({ok:true});
    }
    if(a==="weeklyTop"){
      dailyPrizeCheck();
      const top = weeklyTopCalc();
      if(!top) return json({ok:true, top:null});
      let streak = 0;
      try{ const vv = prizeLogSheet().getDataRange().getValues(); if(vv.length>1){ const last = vv[vv.length-1]; if(String(last[1])===String(top.phone)) streak = +last[2]||0; } }catch(e){}
      const u = findUser(top.phone);
      return json({ok:true, top:{name:(u&&u.name)||"", score:top.score, tests:top.tests, streak}});
    }
    if(a==="setLang"){ const u=findUser(req.phone); if(u){ u.lang=req.lang; saveUser(u);} return json({ok:true}); }
    if(a==="setTheme"){ const u=findUser(req.phone); if(u){ u.theme = (req.theme==="dark" ? "dark" : "light"); saveUser(u);} return json({ok:true}); }
    if(a==="updateDob"){ const u=findUser(req.phone); if(!u) return json({ok:false,error:"nouser"}); if(!/^\d{4}-\d{2}-\d{2}$/.test(String(req.dob||""))) return json({ok:false,error:"baddob"}); u.dob = String(req.dob); saveUser(u); return json({ok:true, dob:normDob(u.dob)}); }
    if(a==="addSeen"){
      const u = findUser(req.phone); if(!u) return json({ok:false});
      const seen = new Set(JSON.parse(u.seen||"[]")); (req.ids||[]).forEach(i=>seen.add(i));
      u.seen = JSON.stringify([...seen]); saveUser(u); return json({ok:true});
    }
    if(a==="submitAttempt"){
      sheet("Attempts").appendRow([req.id, String(req.phone), req.cat, req.sub, req.type, req.score, req.total, req.right, req.wrong, req.skip, req.first, req.at, req.timeUsed||0]);
      return json({ok:true});
    }
    if(a==="requestDeposit"){
      if(rows("Payments").some(p=>String(p.utr)===String(req.utr))) return json({ok:false,error:"dup"});
      const pid = req.id ? String(req.id) : "p" + Date.now() + Math.floor(Math.random()*10000);
      const pat = +req.at || Date.now();
      sheet("Payments").appendRow([pid, String(req.phone), req.amount, String(req.utr), "pending", pat, req.insta ? 1 : 0]);
      try{ tgDepositAlert_({ id:pid, phone:String(req.phone), amount:+req.amount||0, utr:String(req.utr), at:pat, insta:req.insta ? 1 : 0 }); }catch(err){}
      return json({ok:true});
    }
    if(a==="spend"){
      const u = findUser(req.phone); if(!u) return json({ok:false});
      if((+u.stars||0) < req.amount) return json({ok:false,error:"funds"});
      u.stars = (+u.stars||0) - req.amount;
      const led = JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(), amt:-req.amount, label:req.label}); u.ledger = JSON.stringify(led);
      saveUser(u); return json({ok:true, stars:u.stars});
    }
    if(a==="buy"){
      const u = findUser(req.phone); if(!u) return json({ok:false});
      let p = {tests:{},mocks:{},notes:{}}; try{ const x=JSON.parse(u.purchases||"{}"); if(x&&x.tests) p=x; }catch(e){}
      const bucket = req.kind==="test" ? "tests" : req.kind==="mock" ? "mocks" : "notes";
      p[bucket][req.key] = Date.now(); u.purchases = JSON.stringify(p); saveUser(u);
      return json({ok:true});
    }
    if(a==="payClaim"){
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      const kind = req.kind==="sub" ? "sub" : "wallet";
      const amount = kind==="sub" ? 1000 : (+req.amount||0);
      if(kind==="wallet" && [100,200,500].indexOf(amount)<0) return json({ok:false,error:"amount"});
      const cid = "c" + Date.now() + Math.floor(Math.random()*10000);
      sheet("Claims").appendRow([cid, String(req.phone), amount, kind, Date.now(), 0]);
      return json({ok:true, id:cid});
    }
    if(a==="gatewayCredit"){
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      if(u.session && String(req.session||"") !== String(u.session)) return json({ok:false,error:"session"});
      const s = sheet("Claims"); const v = s.getDataRange().getValues();
      for(let i=1;i<v.length;i++){
        if(String(v[i][0])===String(req.claimId)){
          if(String(v[i][1])!==String(req.phone)) return json({ok:false,error:"notyours"});
          if(+v[i][5]) return json({ok:false,error:"used"});
          if(Date.now() - (+v[i][4]||0) > 2*3600*1000) return json({ok:false,error:"expired"});
          s.getRange(i+1,6).setValue(1);
          const amount = +v[i][2]||0, kind = String(v[i][3]);
          const led = JSON.parse(u.ledger||"[]");
          if(kind==="sub"){ const base = Math.max(Date.now(), +u.subUntil||0); u.subUntil = base + 365*24*3600*1000; led.unshift({at:Date.now(), amt:0, label:"StuDy Premium (1 year) - online payment"}); }
          else { u.stars = (+u.stars||0) + amount; led.unshift({at:Date.now(), amt:amount, label:"Wallet top-up Rs."+amount+" (online payment)"}); }
          u.ledger = JSON.stringify(led); saveUser(u);
          sheet("UserNotifs").appendRow([String(req.phone), kind==="sub" ? "Premium activated" : "Payment successful", kind==="sub" ? "Your StuDy Premium subscription is active for 1 year. Everything is unlocked!" : ("Rs."+amount+" has been added to your StuDy wallet."), Date.now(), ""]);
          return json({ok:true, user:pubUser(u)});
        }
      }
      return json({ok:false,error:"notfound"});
    }
    if(a==="buySub"){
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      if(u.session && String(req.session||"") !== String(u.session)) return json({ok:false,error:"session"});
      if((+u.subUntil||0) > Date.now()) return json({ok:false,error:"active"});
      const subPrice = +(((getSettings()||{}).prices||{}).sub) || 1000;
      if((+u.stars||0) < subPrice) return json({ok:false,error:"funds"});
      u.stars = (+u.stars||0) - subPrice;
      u.subUntil = Date.now() + 365*24*3600*1000;
      const led = JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(), amt:-subPrice, label:"StuDy Premium (1 year)"}); u.ledger = JSON.stringify(led);
      saveUser(u); return json({ok:true, user:pubUser(u)});
    }
    const adminOk = () => { const u = findUser(req.adminPhone); return ADMIN_PHONES.indexOf(String(req.adminPhone)) >= 0 && u && String(u.pass)===String(req.adminPass); };
    if(a==="adminStats"){
      if(!adminOk()) return json({ok:false,error:"forbidden"});
      const users = rows("Users");
      let revenue = 0;
      users.forEach(u=>{ try{ JSON.parse(u.ledger||"[]").forEach(l=>{ const amt = +l.amt || 0; if(amt < 0) revenue += -amt; }); }catch(e){} });
      const names = {}; users.forEach(u=>{ names[String(u.phone)] = u.name || ""; });
      const pending = rows("Payments").filter(p=>String(p.status)==="pending").map(p=>({ id:String(p.id), phone:String(p.phone), name:names[String(p.phone)] || "", amount:+p.amount || 0, utr:String(p.utr || ""), at:+p.at || 0, insta:(+p.insta === 1) }));
      const notifs = rows("Notifications").map(n=>({ id:String(n.id), title:n.title || "", body:n.body || "", at:+n.at || 0 })).sort((x,y)=>y.at-x.at);
      return json({ok:true, users:users.length, attempts:rows("Attempts").length, revenue, pending, notifs});
    }
    if(a==="adminDecidePayment"){
      if(!adminOk()) return json({ok:false,error:"forbidden"});
      let utr = req.utr ? String(req.utr) : "";
      if(!utr && req.id){ const row = rows("Payments").find(p=>String(p.id)===String(req.id)); if(row) utr = String(row.utr); }
      if(!utr) return json({ok:false,error:"notfound"});
      const r = decidePayment_(utr, !!req.approve, req.credit!=null ? +req.credit : null);
      return r.ok ? json({ok:true}) : json({ok:false,error:"notfound"});
    }
    if(a==="issueCard"){
      if(!findUser(req.phone)) return json({ok:false,error:"nouser"});
      if(!rows("Cards").some(c=>String(c.code)===String(req.code)))
        sheet("Cards").appendRow([String(req.code), String(req.phone), req.name||"", +req.points||0, +req.tests||0, +req.avg||0, +req.acc||0, String(req.rank), +req.at||Date.now()]);
      return json({ok:true});
    }
    if(a==="verifyCard"){
      const c = rows("Cards").find(x=>String(x.code)===String(req.code||"").toUpperCase());
      if(!c) return json({ok:false});
      return json({ok:true, card:{phone:String(c.phone), name:c.name, points:+c.points, tests:+c.tests, avg:+c.avg, acc:+c.acc, rank:String(c.rank), at:+c.at, code:String(c.code)}});
    }
    if(a==="adminAddStars"){
      if(!adminOk()) return json({ok:false,error:"forbidden"});
      if(!findUser(req.phone)) return json({ok:false,error:"nouser"});
      const amt = +req.amount || 0; if(amt < 1) return json({ok:false,error:"amount"});
      addStars_(String(req.phone), amt, req.message ? String(req.message) : "");
      return json({ok:true});
    }
    if(a==="adminSetSub"){
      if(!adminOk()) return json({ok:false,error:"forbidden"});
      if(!findUser(req.phone)) return json({ok:false,error:"nouser"});
      const days = +req.days || 365;
      const until = giveSub_(String(req.phone), days);
      return json({ok:true, subUntil:until});
    }
    if(a==="adminSaveSettings"){ if(!adminOk()) return json({ok:false,error:"forbidden"});
      const s=sheet("Settings"); const v=s.getDataRange().getValues();
      for(let i=1;i<v.length;i++){ if(v[i][0]==="main"){ s.getRange(i+1,2).setValue(JSON.stringify(req.settings)); return json({ok:true}); } }
      s.appendRow(["main", JSON.stringify(req.settings)]); return json({ok:true}); }
    if(a==="adminAddNotif"){ if(!adminOk()) return json({ok:false,error:"forbidden"});
      postNotif_(req.title, req.body); return json({ok:true}); }
    if(a==="adminDeleteNotif"){ if(!adminOk()) return json({ok:false,error:"forbidden"});
      return deleteNotif_(req.id) ? json({ok:true}) : json({ok:false}); }
    return json({ok:false,error:"unknown action"});
  }catch(err){ return json({ok:false,error:String(err)}); }
}
function doGet(){ return json({ok:true, app:"StuDy backend", hint:"POST actions only"}); }

/* ================= Shared admin helpers (used by the app Admin Panel AND the Telegram bot) ================= */
function decidePayment_(utr, approve, creditOrNull){
  const s = sheet("Payments"); const v = s.getDataRange().getValues();
  for(let i=1;i<v.length;i++){ if(String(v[i][4])==="pending" && String(v[i][3])===String(utr)){
    s.getRange(i+1,5).setValue(approve?"approved":"rejected");
    const payer = String(v[i][1]), amt = +v[i][2], insta = +v[i][6] === 1;
    let credit = 0;
    if(approve){ credit = creditOrNull!=null ? Math.min(+creditOrNull, insta ? Math.round(amt*1.1) : amt) : (insta ? Math.round(amt*1.1) : amt); const u=findUser(payer); if(u){ u.stars=(+u.stars||0)+credit; const led=JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(),amt:credit,label:"Deposit approved Rs."+amt}); u.ledger=JSON.stringify(led); saveUser(u); }
      sheet("UserNotifs").appendRow([payer, "Payment approved ✅", "Your Rs."+amt+" deposit (UTR "+utr+") is approved. "+credit+" rupees have been added to your StuDy wallet." + (credit>amt ? " This includes your 10% Instagram bonus." : ""), Date.now(), ""]);
    } else {
      let body = "Your Rs."+amt+" deposit (UTR "+utr+") could not be approved. Please check the UTR and amount, or contact us on Telegram."; let link = "";
      if(insta){ body = "Your Instagram-discount deposit of Rs."+amt+" (UTR "+utr+") was not approved - we could not verify your follow of @dinesh.jangir_. If you have followed, DM us on Instagram; after checking, we will add your balance."; link = "https://ig.me/m/dinesh.jangir_"; }
      sheet("UserNotifs").appendRow([payer, "Payment not approved", body, Date.now(), link]);
    }
    return {ok:true, phone:payer, amount:amt, credit:credit, insta:insta};
  } }
  return {ok:false, error:"notfound"};
}
function addStars_(phone, amt, message){
  const u = findUser(phone); if(!u) return false;
  u.stars = (+u.stars||0) + amt; const led = JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(), amt:amt, label:"Admin credit"}); u.ledger = JSON.stringify(led); saveUser(u);
  sheet("UserNotifs").appendRow([String(phone), "🎁 Balance added", String(message || (amt+" rupees added to your StuDy wallet.")), Date.now(), ""]);
  return true;
}
function giveSub_(phone, days){
  const u = findUser(phone); if(!u) return 0;
  const base = Math.max(Date.now(), +u.subUntil||0); u.subUntil = base + days*24*3600*1000; saveUser(u);
  sheet("UserNotifs").appendRow([String(phone), "Premium activated", "StuDy Premium has been activated on your account for "+days+" days. Everything is unlocked!", Date.now(), ""]);
  return u.subUntil;
}
function postNotif_(title, body){
  try{ pruneNotifs(); }catch(e){}
  const id = "n"+Date.now();
  sheet("Notifications").appendRow([id, title, body, Date.now()]);
  return id;
}
function deleteNotif_(id){
  const s=sheet("Notifications"); const v=s.getDataRange().getValues();
  for(let i=1;i<v.length;i++){ if(String(v[i][0])===String(id)){ s.deleteRow(i+1); return true; } }
  return false;
}
function statsData_(){
  const users = rows("Users");
  let revenue = 0;
  users.forEach(u=>{ try{ JSON.parse(u.ledger||"[]").forEach(l=>{ const amt = +l.amt || 0; if(amt < 0) revenue += -amt; }); }catch(e){} });
  const names = {}; users.forEach(u=>{ names[String(u.phone)] = u.name || ""; });
  const pending = rows("Payments").filter(p=>String(p.status)==="pending").map(p=>({ id:String(p.id), phone:String(p.phone), name:names[String(p.phone)] || "", amount:+p.amount || 0, utr:String(p.utr || ""), at:+p.at || 0, insta:(+p.insta === 1) }));
  const notifs = rows("Notifications").map(n=>({ id:String(n.id), title:n.title || "", body:n.body || "", at:+n.at || 0 })).sort((x,y)=>y.at-x.at);
  return { users:users, usersCount:users.length, attempts:rows("Attempts").length, revenue:revenue, pending:pending, notifs:notifs, names:names };
}

/* ================= Telegram admin bot =================
 * No hosting needed: alerts are sendMessage calls from this script, and button
 * taps / commands arrive here as webhook POSTs (same Web App URL).
 * Setup (once, in Apps Script): Project Settings -> Script Properties ->
 *   TG_TOKEN = bot token from BotFather
 *   TG_OWNER = owner's Telegram user id
 *   TG_MANAGER = manager's Telegram user id
 * Then deploy a New version and run tgSetWebhook() once from the editor.
 * Roles: owner = master (everything, instant). Manager = payments, balance,
 * notifications, premium — but every manager action first goes to the owner
 * as an Allow/Deny request and runs only after the owner allows it.        */
function tgProp_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ""; }
function tgOwner_(){ return String(tgProp_("TG_OWNER")); }
function tgManager_(){ return String(tgProp_("TG_MANAGER")); }
function tgRole_(id){ const s = String(id); if(s && s===tgOwner_()) return "owner"; if(s && s===tgManager_()) return "manager"; return null; }
function esc_(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
function tgApi_(method, payload){
  const token = tgProp_("TG_TOKEN"); if(!token) return null;
  try{
    const r = UrlFetchApp.fetch("https://api.telegram.org/bot"+token+"/"+method, { method:"post", contentType:"application/json", payload:JSON.stringify(payload||{}), muteHttpExceptions:true });
    return JSON.parse(r.getContentText());
  }catch(e){ return null; }
}
function tgSend_(chatId, text, keyboard){
  const p = { chat_id:chatId, text:text, parse_mode:"HTML", disable_web_page_preview:true };
  if(keyboard) p.reply_markup = { inline_keyboard: keyboard };
  const r = tgApi_("sendMessage", p);
  return (r && r.ok && r.result) ? r.result.message_id : null;
}
function tgEdit_(chatId, msgId, text, keyboard){
  const p = { chat_id:chatId, message_id:msgId, text:text, parse_mode:"HTML", disable_web_page_preview:true };
  p.reply_markup = { inline_keyboard: keyboard || [] };
  return tgApi_("editMessageText", p);
}
function tgAnswer_(cbId, text){ return tgApi_("answerCallbackQuery", { callback_query_id:cbId, text:text||"" }); }
function tgCache_(){ return CacheService.getScriptCache(); }

/* ----- payment alerts ----- */
function tgPayText_(p){
  const t = p.at ? Utilities.formatDate(new Date(+p.at), "Asia/Kolkata", "d MMM, h:mm a") : "";
  return "💰 <b>New payment request</b>\n👤 " + esc_(p.name || "") + " (" + esc_(p.phone) + ")\n💵 <b>₹" + p.amount + "</b>" + (p.insta ? "  🎁 Instagram bonus ticked" : "") + "\n🧾 UTR: <code>" + esc_(p.utr) + "</code>" + (t ? "\n🕐 " + t : "");
}
function tgPayKeyboard_(p){
  const amt = +p.amount || 0;
  const row1 = [{ text:"✅ Approve ₹"+amt, callback_data:"pay:a:"+p.utr }];
  if(p.insta) row1.push({ text:"🎁 Bonus ₹"+Math.round(amt*1.1), callback_data:"pay:b:"+p.utr });
  return [ row1, [{ text:"❌ Reject", callback_data:"pay:r:"+p.utr }] ];
}
function tgSendPayAlert_(chatId, roleKey, p){
  const mid = tgSend_(chatId, tgPayText_(p), tgPayKeyboard_(p));
  if(mid){
    const key = "tgpay_"+p.utr;
    let m = {}; try{ m = JSON.parse(tgCache_().get(key) || "{}"); }catch(e){}
    m[roleKey] = mid; m.text = tgPayText_(p);
    tgCache_().put(key, JSON.stringify(m), 21600);
  }
  return mid;
}
function tgDepositAlert_(p){
  if(!tgProp_("TG_TOKEN")) return;
  const u = findUser(p.phone);
  const full = { id:p.id, phone:p.phone, name:(u && u.name) || p.name || "", amount:p.amount, utr:p.utr, at:p.at, insta:p.insta ? 1 : 0 };
  if(tgOwner_()) tgSendPayAlert_(tgOwner_(), "o", full);
  if(tgManager_()) tgSendPayAlert_(tgManager_(), "m", full);
}
function tgFinalizePay_(utr, line){
  const key = "tgpay_"+utr;
  let m = null; try{ m = JSON.parse(tgCache_().get(key) || "null"); }catch(e){}
  if(!m || !m.text) return;
  const text = m.text + "\n\n" + line;
  if(m.o) tgEdit_(tgOwner_(), m.o, text);
  if(m.m) tgEdit_(tgManager_(), m.m, text);
  tgCache_().remove(key);
}

/* ----- action execution (shared by owner taps, owner commands, allowed manager proposals) ----- */
function execAction_(act){
  if(act.kind==="pay"){
    const credit = act.mode==="a" ? (function(){ const row = rows("Payments").find(p=>String(p.utr)===String(act.utr)); return row ? (+row.amount||0) : null; })() : null;
    const r = decidePayment_(act.utr, act.mode!=="r", credit);
    if(!r.ok) return { ok:false, text:"That payment was already decided (or not found)." };
    const nm = (findUser(r.phone)||{}).name || r.phone;
    if(act.mode==="r"){ tgFinalizePay_(act.utr, "❌ <b>Rejected</b> by " + act.byLabel); return { ok:true, text:"Payment of ₹"+r.amount+" from "+nm+" rejected." }; }
    tgFinalizePay_(act.utr, "✅ <b>Approved</b> — ₹"+r.credit+" credited by " + act.byLabel);
    return { ok:true, text:"₹"+r.credit+" credited to "+nm+" ("+r.phone+") for the ₹"+r.amount+" deposit." };
  }
  if(act.kind==="addstars"){
    if(!addStars_(act.phone, act.amount, act.message||"")) return { ok:false, text:"User "+act.phone+" not found." };
    const nm = (findUser(act.phone)||{}).name || "";
    return { ok:true, text:"₹"+act.amount+" added to "+nm+" ("+act.phone+")." };
  }
  if(act.kind==="notif"){
    postNotif_(act.title, act.body);
    return { ok:true, text:"Notification posted: “"+act.title+"”" };
  }
  if(act.kind==="delnotif"){
    return deleteNotif_(act.id) ? { ok:true, text:"Notification deleted." } : { ok:false, text:"That notification is already gone." };
  }
  if(act.kind==="premium"){
    if(!findUser(act.phone)) return { ok:false, text:"User "+act.phone+" not found." };
    giveSub_(act.phone, 365);
    const nm = (findUser(act.phone)||{}).name || "";
    return { ok:true, text:"1-year Premium given to "+nm+" ("+act.phone+")." };
  }
  return { ok:false, text:"Unknown action." };
}
function actSummary_(act){
  if(act.kind==="pay"){
    const row = rows("Payments").find(p=>String(p.utr)===String(act.utr));
    const nm = row ? ((findUser(row.phone)||{}).name || row.phone) : act.utr;
    const amt = row ? (+row.amount||0) : 0;
    if(act.mode==="r") return "Reject the ₹"+amt+" payment of "+nm+" (UTR "+act.utr+")";
    if(act.mode==="b") return "Approve the ₹"+amt+" payment of "+nm+" WITH 10% Instagram bonus (UTR "+act.utr+")";
    return "Approve the ₹"+amt+" payment of "+nm+" (UTR "+act.utr+")";
  }
  if(act.kind==="addstars") return "Add ₹"+act.amount+" balance to "+(((findUser(act.phone)||{}).name)||"")+" ("+act.phone+")" + (act.message ? " with message: “"+act.message+"”" : "");
  if(act.kind==="notif") return "Post notification: “"+act.title+"” — "+act.body;
  if(act.kind==="delnotif") return "Delete a posted notification (id "+act.id+")";
  if(act.kind==="premium") return "Give 1-year Premium to "+(((findUser(act.phone)||{}).name)||"")+" ("+act.phone+")";
  return act.kind;
}
function tgPropose_(act, fromChatId){
  const token = Math.random().toString(36).slice(2,9);
  act.byLabel = "Anuj (owner approved)";
  tgCache_().put("tgappr_"+token, JSON.stringify({ act:act, mgrChat:String(fromChatId) }), 3600);
  tgSend_(tgOwner_(), "🙋 <b>Anuj requests:</b>\n" + esc_(actSummary_(act)), [[{ text:"✅ Allow", callback_data:"ap:y:"+token }, { text:"❌ Deny", callback_data:"ap:n:"+token }]]);
  tgSend_(fromChatId, "⏳ Sent to the owner for approval:\n" + esc_(actSummary_(act)));
}

/* ----- update handling ----- */
function tgHandle_(upd){
  if(upd.callback_query) return tgHandleCb_(upd.callback_query);
  const msg = upd.message;
  if(!msg || !msg.chat || msg.chat.type!=="private" || !msg.text) return;
  const chatId = msg.chat.id, fromId = msg.from && msg.from.id;
  const role = tgRole_(fromId);
  const text = String(msg.text).trim();
  if(!role){
    if(/^\/(start|id)\b/.test(text)) tgSend_(chatId, "Your Telegram ID is <code>"+fromId+"</code>. Share it with the StuDy owner to get admin access.");
    return;
  }
  tgHandleText_(chatId, role, text);
}
function tgHandleCb_(cb){
  const fromId = cb.from && cb.from.id;
  const role = tgRole_(fromId);
  const data = String(cb.data || "");
  const chatId = cb.message && cb.message.chat && cb.message.chat.id;
  const msgId = cb.message && cb.message.message_id;
  if(!role){ tgAnswer_(cb.id, "Not authorized."); return; }
  if(data.indexOf("pay:")===0){
    const parts = data.split(":"); const mode = parts[1], utr = parts.slice(2).join(":");
    const row = rows("Payments").find(p=>String(p.utr)===utr && String(p.status)==="pending");
    if(!row){ tgAnswer_(cb.id, "Already decided."); tgFinalizePay_(utr, "ℹ️ Already decided."); return; }
    if(role==="owner"){
      const r = execAction_({ kind:"pay", mode:mode, utr:utr, byLabel:"Owner" });
      tgAnswer_(cb.id, r.ok ? "Done ✅" : r.text);
      if(!r.ok) tgFinalizePay_(utr, "ℹ️ Already decided.");
    } else {
      tgPropose_({ kind:"pay", mode:mode, utr:utr }, chatId);
      tgAnswer_(cb.id, "Sent to owner for approval.");
    }
    return;
  }
  if(data.indexOf("ap:")===0){
    if(role!=="owner"){ tgAnswer_(cb.id, "Only the owner can allow this."); return; }
    const parts = data.split(":"); const yes = parts[1]==="y"; const token = parts.slice(2).join(":");
    let rec = null; try{ rec = JSON.parse(tgCache_().get("tgappr_"+token) || "null"); }catch(e){}
    if(!rec){ tgAnswer_(cb.id, "Request expired."); if(chatId && msgId) tgEdit_(chatId, msgId, "⌛ This request expired."); return; }
    tgCache_().remove("tgappr_"+token);
    const summary = actSummary_(rec.act);
    if(yes){
      const r = execAction_(rec.act);
      if(chatId && msgId) tgEdit_(chatId, msgId, "🙋 <b>Anuj requested:</b> " + esc_(summary) + "\n\n✅ <b>Allowed &amp; done.</b> " + esc_(r.text));
      tgSend_(rec.mgrChat, (r.ok ? "✅ Owner approved — done: " : "⚠️ Owner approved but it failed: ") + r.text);
      tgAnswer_(cb.id, r.ok ? "Done ✅" : "Failed");
    } else {
      if(chatId && msgId) tgEdit_(chatId, msgId, "🙋 <b>Anuj requested:</b> " + esc_(summary) + "\n\n❌ <b>Denied by owner.</b>");
      tgSend_(rec.mgrChat, "❌ Owner denied your request: " + summary);
      tgAnswer_(cb.id, "Denied");
    }
    return;
  }
  if(data.indexOf("del:")===0){
    const id = data.slice(4);
    if(role==="owner"){
      const gone = deleteNotif_(id);
      if(chatId && msgId) tgEdit_(chatId, msgId, gone ? "🗑 Notification deleted." : "ℹ️ That notification was already gone.");
      tgAnswer_(cb.id, gone ? "Deleted" : "Already gone");
    } else {
      tgPropose_({ kind:"delnotif", id:id }, chatId);
      tgAnswer_(cb.id, "Sent to owner for approval.");
    }
    return;
  }
  tgAnswer_(cb.id, "");
}
function tgHelpText_(role){
  const head = role==="owner" ? "👑 <b>StuDy Admin — Owner (master)</b>" : "🧑‍💼 <b>StuDy Admin — Manager</b>\nYour actions run after the owner's approval.";
  return head + "\n\n" +
    "/pending — pending payments (with Approve / Bonus / Reject buttons)\n" +
    "/stats — users, attempts, revenue, pending\n" +
    "/users — user list with balances (/users 2 for next page)" + (role==="owner" ? "" : " — owner only") + "\n" +
    "/addbalance — add balance to a user (step by step, or: /addbalance 98XXXXXXXX 200 message)\n" +
    "/premium — give 1-year Premium (/premium 98XXXXXXXX)\n" +
    "/notify — post an app notification (step by step)\n" +
    "/delnotif — delete a posted notification\n" +
    "/cancel — cancel the current step\n" +
    "/id — your Telegram ID";
}
function tgHandleText_(chatId, role, text){
  const cache = tgCache_();
  const stateKey = "tgstate_"+chatId;
  const cmd = (text.split(/\s+/)[0] || "").split("@")[0].toLowerCase();
  if(cmd==="/cancel"){ cache.remove(stateKey); tgSend_(chatId, "Cancelled."); return; }
  let st = null; try{ st = JSON.parse(cache.get(stateKey) || "null"); }catch(e){}
  if(st && text.charAt(0)!=="/"){ tgStateStep_(chatId, role, st, text); return; }
  if(text.charAt(0)==="/") cache.remove(stateKey);
  if(cmd==="/start" || cmd==="/help"){ tgSend_(chatId, tgHelpText_(role)); return; }
  if(cmd==="/id"){ tgSend_(chatId, "Your Telegram ID is <code>"+chatId+"</code>"); return; }
  if(cmd==="/stats"){
    const d = statsData_();
    const pendSum = d.pending.reduce((s,p)=>s+(+p.amount||0),0);
    tgSend_(chatId, "📊 <b>StuDy stats</b>\n👥 Users: "+d.usersCount+"\n📝 Attempts: "+d.attempts+"\n💰 Revenue (spent by users): ₹"+d.revenue+"\n⏳ Pending payments: "+d.pending.length+" (₹"+pendSum+")\n🔔 Live notifications: "+d.notifs.length);
    return;
  }
  if(cmd==="/pending"){
    const d = statsData_();
    if(!d.pending.length){ tgSend_(chatId, "✅ No pending payments."); return; }
    tgSend_(chatId, "⏳ <b>"+d.pending.length+" pending payment(s)</b> — decide below:" + (role==="manager" ? "\n(Your taps go to the owner for approval first.)" : ""));
    d.pending.slice(0,10).forEach(p=> tgSendPayAlert_(chatId, role==="owner" ? "o" : "m", p));
    return;
  }
  if(cmd==="/users"){
    if(role!=="owner"){ tgSend_(chatId, "⛔ The full user list is owner-only."); return; }
    const parts = text.split(/\s+/); const page = Math.max(1, parseInt(parts[1]||"1",10) || 1);
    const all = rows("Users").map(u=>({ name:u.name||"", phone:String(u.phone), stars:+u.stars||0, sub:(+u.subUntil||0)>Date.now() }));
    const size = 15, slice = all.slice((page-1)*size, page*size);
    if(!slice.length){ tgSend_(chatId, "No users on page "+page+"."); return; }
    const lines = slice.map((u,i)=> ((page-1)*size+i+1)+". "+esc_(u.name)+" — "+u.phone+" — ₹"+u.stars+(u.sub?" ⭐Premium":""));
    tgSend_(chatId, "👥 <b>Users</b> (page "+page+" of "+Math.ceil(all.length/size)+", total "+all.length+")\n"+lines.join("\n")+"\n\nNext page: /users "+(page+1));
    return;
  }
  if(cmd==="/delnotif"){
    const d = statsData_();
    if(!d.notifs.length){ tgSend_(chatId, "No live notifications."); return; }
    tgSend_(chatId, "🔔 <b>Posted notifications</b> — tap 🗑 to delete" + (role==="manager" ? " (goes to owner for approval)" : "") + ":");
    d.notifs.slice(0,10).forEach(n=> tgSend_(chatId, "🔔 <b>"+esc_(n.title)+"</b>\n"+esc_(n.body), [[{ text:"🗑 Delete", callback_data:"del:"+n.id }]]));
    return;
  }
  if(cmd==="/addbalance"){
    const parts = text.split(/\s+/);
    if(parts.length>=3 && /^\d{10}$/.test(parts[1]) && (+parts[2]||0) >= 1){
      const act = { kind:"addstars", phone:parts[1], amount:Math.floor(+parts[2]), message:parts.slice(3).join(" ") };
      if(!findUser(act.phone)){ tgSend_(chatId, "User "+act.phone+" not found."); return; }
      if(role==="owner"){ const r = execAction_({ kind:"addstars", phone:act.phone, amount:act.amount, message:act.message, byLabel:"Owner" }); tgSend_(chatId, (r.ok?"✅ ":"⚠️ ") + esc_(r.text)); }
      else tgPropose_(act, chatId);
      return;
    }
    cache.put(stateKey, JSON.stringify({ cmd:"addbalance", step:"phone" }), 1800);
    tgSend_(chatId, "💳 Add balance — send the user's 10-digit phone number (or /cancel).");
    return;
  }
  if(cmd==="/premium"){
    const parts = text.split(/\s+/);
    if(parts.length>=2 && /^\d{10}$/.test(parts[1])){
      if(!findUser(parts[1])){ tgSend_(chatId, "User "+parts[1]+" not found."); return; }
      if(role==="owner"){ const r = execAction_({ kind:"premium", phone:parts[1], byLabel:"Owner" }); tgSend_(chatId, (r.ok?"✅ ":"⚠️ ") + esc_(r.text)); }
      else tgPropose_({ kind:"premium", phone:parts[1] }, chatId);
      return;
    }
    cache.put(stateKey, JSON.stringify({ cmd:"premium", step:"phone" }), 1800);
    tgSend_(chatId, "⭐ Give Premium — send the user's 10-digit phone number (or /cancel).");
    return;
  }
  if(cmd==="/notify"){
    cache.put(stateKey, JSON.stringify({ cmd:"notify", step:"title" }), 1800);
    tgSend_(chatId, "🔔 New notification — send the TITLE (or /cancel).");
    return;
  }
  tgSend_(chatId, "I didn't get that. Here is what I can do:\n\n" + tgHelpText_(role));
}
function tgStateStep_(chatId, role, st, text){
  const cache = tgCache_(); const stateKey = "tgstate_"+chatId;
  const done = ()=> cache.remove(stateKey);
  if(st.cmd==="addbalance"){
    if(st.step==="phone"){
      const ph = text.replace(/\D/g,"");
      if(!/^\d{10}$/.test(ph) || !findUser(ph)){ tgSend_(chatId, "User not found. Send a valid registered 10-digit number, or /cancel."); return; }
      st.phone = ph; st.step = "amount"; cache.put(stateKey, JSON.stringify(st), 1800);
      tgSend_(chatId, "User: "+esc_((findUser(ph)||{}).name||"")+" ("+ph+"). Send the AMOUNT in ₹.");
      return;
    }
    if(st.step==="amount"){
      const amt = Math.floor(+text.replace(/[^\d]/g,""));
      if(!(amt>=1)){ tgSend_(chatId, "Send a valid amount (1 or more), or /cancel."); return; }
      st.amount = amt; st.step = "message"; cache.put(stateKey, JSON.stringify(st), 1800);
      tgSend_(chatId, "Send the message to show the user (e.g. “Payment received, thank you!”), or /skip.");
      return;
    }
    if(st.step==="message"){
      const act = { kind:"addstars", phone:st.phone, amount:st.amount, message:(text==="/skip" ? "" : text) };
      done();
      if(role==="owner"){ const r = execAction_({ kind:"addstars", phone:act.phone, amount:act.amount, message:act.message, byLabel:"Owner" }); tgSend_(chatId, (r.ok?"✅ ":"⚠️ ") + esc_(r.text)); }
      else tgPropose_(act, chatId);
      return;
    }
  }
  if(st.cmd==="premium" && st.step==="phone"){
    const ph = text.replace(/\D/g,"");
    if(!/^\d{10}$/.test(ph) || !findUser(ph)){ tgSend_(chatId, "User not found. Send a valid registered 10-digit number, or /cancel."); return; }
    done();
    if(role==="owner"){ const r = execAction_({ kind:"premium", phone:ph, byLabel:"Owner" }); tgSend_(chatId, (r.ok?"✅ ":"⚠️ ") + esc_(r.text)); }
    else tgPropose_({ kind:"premium", phone:ph }, chatId);
    return;
  }
  if(st.cmd==="notify"){
    if(st.step==="title"){ st.title = text; st.step = "body"; cache.put(stateKey, JSON.stringify(st), 1800); tgSend_(chatId, "Now send the BODY text."); return; }
    if(st.step==="body"){
      const act = { kind:"notif", title:st.title, body:text };
      done();
      if(role==="owner"){ const r = execAction_({ kind:"notif", title:act.title, body:act.body, byLabel:"Owner" }); tgSend_(chatId, (r.ok?"✅ ":"⚠️ ") + esc_(r.text)); }
      else tgPropose_(act, chatId);
      return;
    }
  }
  done();
  tgSend_(chatId, "That flow ended. Use /help to see commands.");
}

/* ----- one-time setup helpers (run manually from the Apps Script editor) ----- */
function tgSetWebhook(){
  const url = ScriptApp.getService().getUrl();
  const r = tgApi_("setWebhook", { url:url, allowed_updates:["message","callback_query"] });
  Logger.log("setWebhook -> " + JSON.stringify(r));
}
function tgWebhookInfo(){
  Logger.log(JSON.stringify(tgApi_("getWebhookInfo", {})));
}
function tgSelfTest(){
  if(tgOwner_()) tgSend_(tgOwner_(), "✅ <b>StuDy admin bot connected.</b> You are the OWNER (master). Send /help for commands.");
  if(tgManager_()) tgSend_(tgManager_(), "✅ <b>StuDy admin bot connected.</b> You are the MANAGER — your actions run after the owner's approval. Send /help for commands.");
}
