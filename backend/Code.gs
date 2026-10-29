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
const SHEET_NAMES = ["Users","Attempts","Payments","Settings","Notifications","OTPs","UserNotifs","Cards"];

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name){ let s = ss().getSheetByName(name); if(!s){ s = ss().insertSheet(name); } return s; }
function setup(){
  const headers = {
    Users:["phone","name","gmail","dob","pass","stars","lang","purchases","ledger","seen","createdAt"],
    Attempts:["id","phone","cat","sub","type","score","total","right","wrong","skip","first","at","timeUsed"],
    Payments:["id","phone","amount","utr","status","at","insta"],
    UserNotifs:["phone","title","body","at","link"],
    Cards:["code","phone","name","points","tests","avg","acc","rank","at"],
    Settings:["key","json"],
    Notifications:["id","title","body","at"],
    OTPs:["gmail","code","exp","phone"]
  };
  SHEET_NAMES.forEach(n=>{ const s=sheet(n); if(s.getLastRow()===0) s.appendRow(headers[n]); });
}
function rows(name){
  const s = sheet(name); const v = s.getDataRange().getValues();
  if(v.length<2 && s.getLastRow()===0) return [];
  const head = v[0]; return v.slice(1).map(r=>{ const o={}; head.forEach((h,i)=>o[h]=r[i]); return o; });
}
function findUser(phone){ return rows("Users").find(u=>String(u.phone)===String(phone)); }
function saveUser(u){
  const s = sheet("Users"); const v = s.getDataRange().getValues(); const head = v[0];
  const rowOf = o => head.map(h=>{ let val=o[h]; if(typeof val==="object") val=JSON.stringify(val); return val===undefined?"":val; });
  for(let i=1;i<v.length;i++){ if(String(v[i][0])===String(u.phone)){ const r=rowOf(u); r.forEach((val,j)=> s.getRange(i+1,j+1).setValue(val)); return; } }
  s.appendRow(rowOf(u));
}
function pubUser(u){
  let p = {tests:{},mocks:{},notes:{}}; try{ const x = JSON.parse(u.purchases||"{}"); if(x && x.tests) p = x; }catch(e){}
  let led = []; try{ led = JSON.parse(u.ledger||"[]"); }catch(e){}
  return { phone:String(u.phone), name:u.name||"", gmail:u.gmail, dob:u.dob, stars:+u.stars||0, lang:u.lang||"en", purchases:p, ledger:led };
}
function getSettings(){
  const r = rows("Settings").find(x=>x.key==="main");
  if(r) return JSON.parse(r.json);
  return { prices:{test:5,mock:10,notes:30,bundle:50},
    help:{en:"See Help in the app.",hi:"See Help in the app."}, contact:{en:"Telegram: @dj3680",hi:"Telegram: @dj3680"} };
}
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

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
      const ok = rows("OTPs").some(o=>String(o.gmail)===String(req.gmail) && String(o.code)===String(req.code) && +o.exp>Date.now());
      if(!ok) return json({ok:false,error:"badotp"});
      const u = { phone:String(req.phone), name:req.name||"", gmail:req.gmail, dob:req.dob, pass:req.pass, stars:0, lang:"en",
                  purchases: JSON.stringify({tests:{},mocks:{},notes:{}}), ledger:"[]", seen:"[]", createdAt:Date.now() };
      saveUser(u);
      return json({ok:true, user:{phone:u.phone, name:u.name, gmail:u.gmail, dob:u.dob, stars:0, lang:"en", purchases:{tests:{},mocks:{},notes:{}}, ledger:[]}});
    }
    if(a==="login"){
      const u = findUser(req.phone); if(!u) return json({ok:false,error:"nouser"});
      if(String(u.pass)!==String(req.pass)) return json({ok:false,error:"badpass"});
      return json({ok:true, user:pubUser(u)});
    }
    if(a==="me"){ const u = findUser(req.phone); if(!u) return json({ok:false}); const un = rows("UserNotifs").filter(n=>String(n.phone)===String(req.phone)).map(n=>({title:n.title, body:n.body, at:+n.at, link:n.link||""})).sort((x,y)=>y.at-x.at); const pu = pubUser(u); pu._unotifs = un; return json({ok:true, user:pu}); }
    if(a==="leaderboard"){
      const names = {}; rows("Users").forEach(u=> names[String(u.phone)] = u.name || String(u.phone));
      const agg = {};
      rows("Attempts").forEach(r=>{ if(r.first===true || String(r.first)==="TRUE"){ const k = String(r.phone); agg[k] = agg[k] || {name: names[k] || k, score:0, tests:0, phone: "XXXXX" + k.slice(-5)}; agg[k].score += (+r.score||0); agg[k].tests++; } });
      return json({ok:true, rows: Object.values(agg).sort((x,y)=>y.score-x.score).slice(0,10)});
    }
    if(a==="setLang"){ const u=findUser(req.phone); if(u){ u.lang=req.lang; saveUser(u);} return json({ok:true}); }
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
      sheet("Payments").appendRow([req.id, String(req.phone), req.amount, String(req.utr), "pending", req.at, req.insta ? 1 : 0]);
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
    const adminOk = () => { const u = findUser(ADMIN_PHONE); return String(req.phone)===ADMIN_PHONE && u && String(u.pass)===String(req.adminPass); };
    if(a==="adminDecidePayment"){
      if(!adminOk()) return json({ok:false,error:"forbidden"});
      const s = sheet("Payments"); const v = s.getDataRange().getValues();
      for(let i=1;i<v.length;i++){ if(String(v[i][0])===String(req.id)){
        s.getRange(i+1,5).setValue(req.approve?"approved":"rejected");
        const payer = String(v[i][1]), amt = +v[i][2], utr = String(v[i][3]), insta = +v[i][6] === 1;
        if(req.approve){ const credit = req.credit!=null ? Math.min(+req.credit, insta ? Math.round(amt*1.1) : amt) : (insta ? Math.round(amt*1.1) : amt); const u=findUser(payer); if(u){ u.stars=(+u.stars||0)+credit; const led=JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(),amt:credit,label:"Deposit approved Rs."+amt}); u.ledger=JSON.stringify(led); saveUser(u); }
          sheet("UserNotifs").appendRow([payer, "Payment approved \u2705", "Your Rs."+amt+" deposit (UTR "+utr+") is approved. "+credit+" stars have been added to your D Coins wallet." + (credit>amt ? " This includes your 10% Instagram bonus." : ""), Date.now(), ""]);
        } else {
          let body = "Your Rs."+amt+" deposit (UTR "+utr+") could not be approved. Please check the UTR and amount, or contact us on Telegram."; let link = "";
          if(insta){ body = "Your Instagram-discount deposit of Rs."+amt+" (UTR "+utr+") was not approved - we could not verify your follow of @dinesh.jangir_. If you have followed, DM us on Instagram; after checking, we will add your stars."; link = "https://ig.me/m/dinesh.jangir_"; }
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
      u.stars = (+u.stars||0) + amt; const led = JSON.parse(u.ledger||"[]"); led.unshift({at:Date.now(), amt, label:"Admin stars"}); u.ledger = JSON.stringify(led); saveUser(u);
      sheet("UserNotifs").appendRow([String(req.phone), "\u2B50 Stars added", String(req.message || (amt+" stars added to your wallet.")), Date.now(), ""]);
      return json({ok:true});
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
