/*************************************************
 * StuDy v2 — Google Apps Script backend
 * Data lives in the attached Google Sheet with tabs:
 *   Users | Attempts | Payments | Settings | Notifications | OTPs
 * (Run setup() once from the editor to create them.)
 * Deploy: Deploy → New deployment → Web app → Execute as: Me,
 * Access: Anyone. Paste the Web App URL into BACKEND_URL in index.html.
 * Admin: ADMIN_PHONE below; its password is whatever that account
 * was registered with. Register that number normally first.
 *************************************************/
const ADMIN_PHONE = "9588293680";
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
  return { prices:{test:5,mock:10,notes:30,bundle:50},
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
function doPost(e){
  let req; try{ req = JSON.parse(e.postData.contents); }catch(err){ return json({ok:false,error:"badjson"}); }
  const a = req.action;
  try{
    if(a==="bootstrap"){
      return json({ok:true, settings:getSettings(), notifs: rows("Notifications").map(n=>({id:String(n.id),title:n.title,body:n.body,at:+n.at})).sort((x,y)=>y.at-x.at) });
    }
    if(a==="sendOtp"){
      if(findUser(req.phone)) return json({ok:false,error:"exists"});
      const code = String(Math.floor(100000+Math.random()*900000));
      sheet("OTPs").appendRow([req.gmail, code, Date.now()+10*60*1000, req.phone]);
      MailApp.sendEmail(req.gmail, "StuDy - your OTP", "Your StuDy registration OTP is: " + code + "\n\nValid for 10 minutes. If you did not request this, ignore this email.");
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
      MailApp.sendEmail(u.gmail, "StuDy - password reset OTP", "Your StuDy password reset OTP is: " + code + "\n\nValid for 10 minutes. If you did not request this, ignore this email.");
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
      if((+u.stars||0) < 1000) return json({ok:false,error:"funds"});
      u.stars = (+u.stars||0) - 1000;
      u.subUntil = Date.now() + 365*24*3600*1000;
      const led = JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(), amt:-1000, label:"StuDy Premium (1 year)"}); u.ledger = JSON.stringify(led);
      saveUser(u); return json({ok:true, user:pubUser(u)});
    }
    const adminOk = () => { const u = findUser(ADMIN_PHONE); return String(req.adminPhone)===ADMIN_PHONE && u && String(u.pass)===String(req.adminPass); };
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
      const s = sheet("Payments"); const v = s.getDataRange().getValues();
      for(let i=1;i<v.length;i++){ const isTarget = String(v[i][4])==="pending" && ( (req.id && String(v[i][0])===String(req.id)) || (req.utr && String(v[i][3])===String(req.utr)) ); if(isTarget){
        s.getRange(i+1,5).setValue(req.approve?"approved":"rejected");
        const payer = String(v[i][1]), amt = +v[i][2], utr = String(v[i][3]), insta = +v[i][6] === 1;
        if(req.approve){ const credit = req.credit!=null ? Math.min(+req.credit, insta ? Math.round(amt*1.1) : amt) : (insta ? Math.round(amt*1.1) : amt); const u=findUser(payer); if(u){ u.stars=(+u.stars||0)+credit; const led=JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(),amt:credit,label:"Deposit approved Rs."+amt}); u.ledger=JSON.stringify(led); saveUser(u); }
          sheet("UserNotifs").appendRow([payer, "Payment approved \u2705", "Your Rs."+amt+" deposit (UTR "+utr+") is approved. "+credit+" rupees have been added to your StuDy wallet." + (credit>amt ? " This includes your 10% Instagram bonus." : ""), Date.now(), ""]);
        } else {
          let body = "Your Rs."+amt+" deposit (UTR "+utr+") could not be approved. Please check the UTR and amount, or contact us on Telegram."; let link = "";
          if(insta){ body = "Your Instagram-discount deposit of Rs."+amt+" (UTR "+utr+") was not approved - we could not verify your follow of @dinesh.jangir_. If you have followed, DM us on Instagram; after checking, we will add your balance."; link = "https://ig.me/m/dinesh.jangir_"; }
          sheet("UserNotifs").appendRow([payer, "Payment not approved", body, Date.now(), link]);
        }
        return json({ok:true}); } }
      return json({ok:false,error:"notfound"});
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
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      const amt = +req.amount || 0; if(amt < 1) return json({ok:false,error:"amount"});
      u.stars = (+u.stars||0) + amt; const led = JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(), amt, label:"Admin credit"}); u.ledger = JSON.stringify(led); saveUser(u);
      sheet("UserNotifs").appendRow([String(req.phone), "\u{1F381} Balance added", String(req.message || (amt+" rupees added to your StuDy wallet.")), Date.now(), ""]);
      return json({ok:true});
    }
    if(a==="adminSetSub"){
      if(!adminOk()) return json({ok:false,error:"forbidden"});
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      const days = +req.days || 365;
      const base = Math.max(Date.now(), +u.subUntil||0); u.subUntil = base + days*24*3600*1000; saveUser(u);
      sheet("UserNotifs").appendRow([String(req.phone), "Premium activated", "StuDy Premium has been activated on your account for "+days+" days. Everything is unlocked!", Date.now(), ""]);
      return json({ok:true, subUntil:u.subUntil});
    }
    if(a==="adminSaveSettings"){ if(!adminOk()) return json({ok:false,error:"forbidden"});
      const s=sheet("Settings"); const v=s.getDataRange().getValues();
      for(let i=1;i<v.length;i++){ if(v[i][0]==="main"){ s.getRange(i+1,2).setValue(JSON.stringify(req.settings)); return json({ok:true}); } }
      s.appendRow(["main", JSON.stringify(req.settings)]); return json({ok:true}); }
    if(a==="adminAddNotif"){ if(!adminOk()) return json({ok:false,error:"forbidden"});
      sheet("Notifications").appendRow(["n"+Date.now(), req.title, req.body, Date.now()]); return json({ok:true}); }
    if(a==="adminDeleteNotif"){ if(!adminOk()) return json({ok:false,error:"forbidden"});
      const s=sheet("Notifications"); const v=s.getDataRange().getValues();
      for(let i=1;i<v.length;i++){ if(String(v[i][0])===String(req.id)){ s.deleteRow(i+1); return json({ok:true}); } }
      return json({ok:false}); }
    return json({ok:false,error:"unknown action"});
  }catch(err){ return json({ok:false,error:String(err)}); }
}
function doGet(){ return json({ok:true, app:"StuDy backend", hint:"POST actions only"}); }
