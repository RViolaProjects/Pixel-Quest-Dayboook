// Keep the existing MOSSLIGHT_SECRET and MOSSLIGHT_EMAIL property names for compatibility.
// Use a 24+ character secret and your own recipient address. Never share the secret.
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = PropertiesService.getScriptProperties();
    const input = JSON.parse(e.postData.contents);
    if (!p.getProperty('MOSSLIGHT_SECRET') || input.secret !== p.getProperty('MOSSLIGHT_SECRET')) throw Error('Unauthorized');
    if (input.email !== p.getProperty('MOSSLIGHT_EMAIL')) throw Error('Recipient must match MOSSLIGHT_EMAIL');
    if (!Array.isArray(input.reminders) || input.reminders.length > 200) throw Error('Maximum 200 reminders');
    const reminders = input.reminders.map(r => {
      if (typeof r.id !== 'string' || typeof r.title !== 'string' || !Number.isFinite(Date.parse(r.at))) throw Error('Invalid reminder');
      return {id:r.id.slice(0,100),title:r.title.slice(0,160),at:r.at};
    });
    const sent=readSent(p);
    // Chunk properties to respect Apps Script's 9 KB per-value limit.
    for(let i=0;i<20;i++)p.deleteProperty('schedule_'+i);
    for(let i=0;i<reminders.length;i+=10)p.setProperty('schedule_'+Math.floor(i/10),JSON.stringify(reminders.slice(i,i+10)));
    const active = {}; reminders.forEach(r=>{const key=r.id+'|'+r.at;if(sent[key])active[key]=sent[key]});
    writeSent(p,active);
    return json({ok:true});
  } catch(err) { return json({ok:false,error:err.message}); }
  finally { if(lock.hasLock())lock.releaseLock(); }
}
function json(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON)}
function installReminders(){
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='sendReminders').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendReminders').timeBased().everyMinutes(5).create();
}
function sendReminders(){
  const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return;
  try{
    const p=PropertiesService.getScriptProperties(),email=p.getProperty('MOSSLIGHT_EMAIL');if(!email)return;
    let reminders=[];for(let i=0;i<20;i++)reminders=reminders.concat(JSON.parse(p.getProperty('schedule_'+i)||'[]'));
    const sent=readSent(p),now=Date.now();
    const due=reminders.filter(r=>!sent[r.id+'|'+r.at]&&Date.parse(r.at)<=now&&now-Date.parse(r.at)<86400000);
    if(!due.length)return;
    MailApp.sendEmail(email,'Pixel Quest Daybook · Your quest reminders',due.map(r=>r.title+' — '+new Date(r.at).toISOString()).join('\n')+'\n\nA little progress, every day.');
    due.forEach(r=>sent[r.id+'|'+r.at]=now);writeSent(p,sent);
  }finally{lock.releaseLock()}
}

function readSent(p){let result={};for(let i=0;i<20;i++)Object.assign(result,JSON.parse(p.getProperty("sent_"+i)||"{}"));return result}
function writeSent(p,v){const entries=Object.entries(v);for(let i=0;i<20;i++)p.deleteProperty("sent_"+i);for(let i=0;i<entries.length;i+=10)p.setProperty("sent_"+Math.floor(i/10),JSON.stringify(Object.fromEntries(entries.slice(i,i+10))))}
