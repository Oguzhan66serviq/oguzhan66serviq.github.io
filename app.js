const SUPABASE_URL = 'https://rkhptbohniykwxhwhjiz.supabase.co';
const SUPABASE_KEY = 'sb_publishable_I3Pz0TnaVqRt6GdszYq_Jw_BETkoJAk';
const AI_URL = `${SUPABASE_URL}/functions/v1/serviq-ai`;
const SERVIQ_URL = 'https://serviq-catering-ai-ogu.oguzhan-yoeruerer.chatgpt.site';
const CALLBACK_URL = `${location.origin}/auth-callback-v2.html`;
const TOKEN_KEY = 'serviq_access_token';
const REFRESH_KEY = 'serviq_refresh_token';
const state = { token: localStorage.getItem(TOKEN_KEY) || '', refreshToken: localStorage.getItem(REFRESH_KEY) || '', user: null, membership: null, mail: null, company: {}, rules: {}, catalog: [] };
const $ = id => document.getElementById(id);
const euro = value => new Intl.NumberFormat('de-DE',{style:'currency',currency:'EUR'}).format(value);
const html = value => String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));

function show(id){ ['signinView','importView','unsupportedView'].forEach(x => $(x).classList.toggle('hidden', x !== id)); }
function message(text,type=''){ $('message').textContent=text; $('message').className=`message ${type}`; }
function friendlyError(error){
  const raw=String(error?.message||error);
  try{ const data=JSON.parse(raw); return data.error||raw; }catch(_){ return raw; }
}
function storeSession(session){
  state.token=session.access_token||''; state.refreshToken=session.refresh_token||'';
  if(state.token)localStorage.setItem(TOKEN_KEY,state.token); else localStorage.removeItem(TOKEN_KEY);
  if(state.refreshToken)localStorage.setItem(REFRESH_KEY,state.refreshToken); else localStorage.removeItem(REFRESH_KEY);
}
function clearSession(){storeSession({access_token:'',refresh_token:''});state.user=null;state.membership=null;}
async function refreshSession(){
  if(!state.refreshToken)throw new Error('Bitte erneut anmelden.');
  const response=await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=refresh_token`,{method:'POST',headers:{apikey:SUPABASE_KEY,'Content-Type':'application/json'},body:JSON.stringify({refresh_token:state.refreshToken})});
  const result=await response.json().catch(()=>({}));
  if(!response.ok)throw new Error(result.error_description||result.msg||'Die Anmeldung ist abgelaufen.');
  storeSession(result);return result;
}
async function api(path,options={},retried=false){
  const response=await fetch(`${SUPABASE_URL}${path}`,{...options,headers:{apikey:SUPABASE_KEY,Authorization:`Bearer ${state.token}`,'Content-Type':'application/json',...(options.headers||{})}});
  if(response.status===401&&!retried&&state.refreshToken){await refreshSession();return api(path,options,true);}
  if(!response.ok) throw new Error(await response.text()||`HTTP ${response.status}`);
  if(response.status===204) return null;
  return response.json();
}
async function ai(operation,payload){
  const response=await fetch(AI_URL,{method:'POST',headers:{apikey:SUPABASE_KEY,Authorization:`Bearer ${state.token}`,'Content-Type':'application/json'},body:JSON.stringify({operation,...payload,consent_to_openai:true})});
  const result=await response.json().catch(()=>({}));
  if(!response.ok) throw new Error(result.error||'Die KI-Verarbeitung ist fehlgeschlagen.');
  return result.data;
}

async function loadAccount(){
  state.user=await api('/auth/v1/user');
  const rows=await api(`/rest/v1/organization_members?user_id=eq.${encodeURIComponent(state.user.id)}&status=eq.active&select=organization_id,display_name,role,organizations(name)&limit=1`);
  if(!rows.length) throw new Error('Für dieses Microsoft-Konto gibt es noch keinen Serviq-Arbeitsbereich. Lade die Uni-Adresse zuerst in Serviq unter „Team“ ein.');
  state.membership=rows[0];
  const org=encodeURIComponent(state.membership.organization_id);
  const [settings,catalog]=await Promise.all([
    api(`/rest/v1/company_settings?organization_id=eq.${org}&select=profile,calculation_rules&limit=1`),
    api(`/rest/v1/service_catalog_items?organization_id=eq.${org}&active=eq.true&select=external_id,name,detail,category,unit,default_qty,price,default_selected&order=created_at.asc`)
  ]);
  state.company=settings[0]?.profile||{};
  state.rules=settings[0]?.calculation_rules||{};
  state.catalog=catalog||[];
  if(!state.catalog.length) throw new Error('Im Serviq-Arbeitsbereich ist noch kein aktiver Leistungskatalog hinterlegt.');
  $('accountName').textContent=state.membership.display_name||state.user.email;
  $('orgName').textContent=state.membership.organizations?.name||'Serviq-Arbeitsbereich';
}

function readMail(){
  return new Promise((resolve,reject)=>{
    const item=Office.context.mailbox?.item;
    if(!item) return reject(new Error('Keine geöffnete Outlook-Nachricht gefunden.'));
    item.body.getAsync(Office.CoercionType.Text,result=>{
      if(result.status!==Office.AsyncResultStatus.Succeeded) return reject(new Error('Der Nachrichtentext konnte nicht gelesen werden.'));
      resolve({subject:item.subject||'Catering-Anfrage',from:item.from?.emailAddress||'',fromName:item.from?.displayName||'',body:result.value||'',messageId:item.internetMessageId||item.itemId||''});
    });
  });
}

async function initialize(){
  if(!state.token){show('signinView');return;}
  try{
    await loadAccount(); state.mail=await readMail();
    $('mailSubject').textContent=state.mail.subject;
    $('mailFrom').textContent=state.mail.fromName?`${state.mail.fromName} · ${state.mail.from}`:state.mail.from||'Unbekannt';
    show('importView');
  }catch(error){
    if(/JWT|token|session|401|abgelaufen/i.test(error.message)){clearSession();show('signinView');}
    else{show('importView');message(friendlyError(error),'error');$('importBtn').disabled=true;}
  }
}

async function beginSignIn(event){
  event.preventDefault();
  const button=$('signInBtn');const status=$('signInMessage');
  button.disabled=true;button.textContent='Anmeldung läuft …';status.textContent='';status.className='message';
  try{
    const response=await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`,{method:'POST',headers:{apikey:SUPABASE_KEY,'Content-Type':'application/json'},body:JSON.stringify({email:$('loginEmail').value.trim(),password:$('loginPassword').value})});
    const result=await response.json().catch(()=>({}));
    if(!response.ok)throw new Error(result.error_description||result.msg||result.error||'Anmeldung fehlgeschlagen.');
    storeSession(result);$('loginPassword').value='';await initialize();
  }catch(error){status.textContent=friendlyError(error);status.className='message error';}
  finally{button.disabled=false;button.textContent='Mit Serviq anmelden';}
}

function beginMicrosoftSignIn(){
  const button=$('microsoftSignInBtn');const status=$('signInMessage');
  button.disabled=true;status.textContent='Microsoft-Anmeldung wird geöffnet …';status.className='message';
  const startUrl=`${location.origin}/auth-start.html?callback=${encodeURIComponent(CALLBACK_URL)}`;
  Office.context.ui.displayDialogAsync(startUrl,{height:70,width:40,displayInIframe:false},result=>{
    if(result.status!==Office.AsyncResultStatus.Succeeded){
      button.disabled=false;status.textContent=`Anmeldefenster konnte nicht geöffnet werden: ${result.error.message}`;status.className='message error';return;
    }
    const dialog=result.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived,async event=>{
      let payload={};
      try{payload=JSON.parse(event.message);}catch(_){payload={error:'Ungültige Antwort der Anmeldung.'};}
      dialog.close();button.disabled=false;
      if(payload.error){status.textContent=payload.error;status.className='message error';return;}
      storeSession(payload);await initialize();
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived,()=>{button.disabled=false;});
  });
}

function calculateLines(proposal,guestCount){
  const suggestions=new Map((proposal.items||[]).map(item=>[item.id,item]));
  return state.catalog.map(item=>{
    const suggestion=suggestions.get(item.external_id);
    const fallback=item.unit==='Personen'?guestCount:Number(item.default_qty)||1;
    const quantity=Math.max(0,Number(suggestion?.quantity??(item.default_selected?fallback:0)));
    return {id:item.external_id,name:item.name,detail:item.detail||'',unit:item.unit,quantity,price:Number(item.price)||0,total:quantity*(Number(item.price)||0),rationale:suggestion?.rationale||''};
  }).filter(item=>item.quantity>0);
}

function pdfSafe(value){
  return String(value??'').replace(/[äÄ]/g,'ae').replace(/[öÖ]/g,'oe').replace(/[üÜ]/g,'ue').replace(/ß/g,'ss').replace(/[^ -~]/g,'').replace(/([\\()])/g,'\\$1');
}
function wrap(text,max=82){
  const words=pdfSafe(text).split(/\s+/); const lines=[]; let line='';
  words.forEach(word=>{const next=line?`${line} ${word}`:word;if(next.length>max&&line){lines.push(line);line=word;}else line=next;});
  if(line)lines.push(line); return lines;
}
function makePdf(quote){
  const lines=[]; const add=(text,size=10,bold=false,gap=15)=>lines.push({text,size,bold,gap});
  add(state.company.name||state.membership.organizations?.name||'Catering-Unternehmen',20,true,28);
  add(`Angebot ${quote.number}`,16,true,24);
  add(`Kunde: ${quote.request.customerCompany||quote.request.contactName||state.mail.fromName||state.mail.from}`);
  add(`Veranstaltung: ${quote.request.eventName||state.mail.subject}`);
  add(`Termin: ${quote.request.eventDate||'nach Abstimmung'} · Gaeste: ${quote.request.guestCount||'offen'}`,10,false,22);
  wrap(quote.introduction).forEach(t=>add(t,10,false,13)); add('',10,false,8);
  add('Leistungen',12,true,20);
  quote.lines.forEach(item=>{add(`${item.name} — ${item.quantity} ${item.unit} x ${euro(item.price)} = ${euro(item.total)}`,9,false,13);});
  add('',10,false,8); add(`Netto: ${euro(quote.net)}`,11,true,16); add(`MwSt. ${quote.vatRate}%: ${euro(quote.tax)}`); add(`Gesamt: ${euro(quote.gross)}`,13,true,22);
  wrap(quote.closing).forEach(t=>add(t,10,false,13));
  if(state.company.paymentTerms)add(`Zahlungsbedingungen: ${state.company.paymentTerms}`,9,false,13);
  const commands=[]; let y=790;
  for(const row of lines){ if(y<55)break; commands.push(`BT /F${row.bold?2:1} ${row.size} Tf 50 ${y} Td (${pdfSafe(row.text)}) Tj ET`); y-=row.gap; }
  const stream=commands.join('\n');
  const objects=[
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R /F2 6 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>'
  ];
  let pdf='%PDF-1.4\n'; const offsets=[0];
  objects.forEach((object,index)=>{offsets[index+1]=pdf.length;pdf+=`${index+1} 0 obj\n${object}\nendobj\n`;});
  const xref=pdf.length; pdf+=`xref\n0 ${objects.length+1}\n0000000000 65535 f \n`;
  for(let i=1;i<=objects.length;i++)pdf+=String(offsets[i]).padStart(10,'0')+' 00000 n \n';
  pdf+=`trailer << /Size ${objects.length+1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return btoa(pdf);
}

function replyHtml(quote){
  const name=quote.request.contactName||state.mail.fromName||'Damen und Herren';
  const rows=quote.lines.map(item=>`<tr><td style="padding:6px;border-bottom:1px solid #ddd">${html(item.name)}</td><td style="padding:6px;text-align:right;border-bottom:1px solid #ddd">${html(item.quantity)} ${html(item.unit)}</td><td style="padding:6px;text-align:right;border-bottom:1px solid #ddd">${html(euro(item.total))}</td></tr>`).join('');
  return `<p>Guten Tag ${html(name)},</p><p>${html(quote.introduction)}</p><table style="border-collapse:collapse;width:100%"><tbody>${rows}<tr><td colspan="2" style="padding:8px;text-align:right"><b>Gesamt inkl. MwSt.</b></td><td style="padding:8px;text-align:right"><b>${html(euro(quote.gross))}</b></td></tr></tbody></table><p>Das vollständige Angebot finden Sie als PDF im Anhang.</p><p>${html(quote.closing)}</p><p>Freundliche Grüße<br><b>${html(state.company.name||state.membership.organizations?.name||'Ihr Catering-Team')}</b></p>`;
}
function openReply(quote,pdfBase64){
  return new Promise((resolve,reject)=>{
    if(!Office.context.requirements.isSetSupported('Mailbox','1.15')) return reject(new Error('Diese Outlook-Version unterstützt PDF-Anhänge im Antwortentwurf noch nicht. Bitte Outlook im Web oder die aktuelle Desktop-Version verwenden.'));
    Office.context.mailbox.item.displayReplyFormAsync({htmlBody:replyHtml(quote),attachments:[{base64file:pdfBase64,name:`Angebot-${quote.number}.pdf`,type:Office.MailboxEnums.AttachmentType.Base64,inLine:false}]},result=>{
      if(result.status===Office.AsyncResultStatus.Succeeded)resolve(); else reject(new Error(result.error?.message||'Der Outlook-Antwortentwurf konnte nicht geöffnet werden.'));
    });
  });
}

async function saveQuote(requestData,proposal,lines,totals,number){
  const requestPayload={organization_id:state.membership.organization_id,source_type:'outlook',source_name:state.mail.subject,raw_text:state.mail.body,customer_company:requestData.customerCompany||'',contact_name:requestData.contactName||state.mail.fromName,contact_email:requestData.contactEmail||state.mail.from,event_name:requestData.eventName||state.mail.subject,event_date:requestData.eventDate||null,event_time:requestData.eventTime||'',venue:requestData.venue||'',guest_count:Number(requestData.guestCount)||1,preferences:requestData.dietary||'',notes:$('note').value.trim()||requestData.notes||'',extracted_data:{...requestData,message_id:state.mail.messageId},missing_fields:requestData.missing_fields||[],confidence:{uncertain_fields:requestData.uncertain_fields||[]},status:'quoted',created_by:state.user.id};
  const requests=await api('/rest/v1/customer_requests',{method:'POST',headers:{Prefer:'return=representation'},body:JSON.stringify(requestPayload)});
  const offers=await api('/rest/v1/offers',{method:'POST',headers:{Prefer:'return=representation'},body:JSON.stringify({organization_id:state.membership.organization_id,request_id:requests[0].id,offer_number:number,status:'needs_review',net_total:totals.net,tax_total:totals.tax,gross_total:totals.gross,created_by:state.user.id})});
  await api('/rest/v1/offer_versions',{method:'POST',headers:{Prefer:'return=minimal'},body:JSON.stringify({organization_id:state.membership.organization_id,offer_id:offers[0].id,version_no:1,snapshot:{request:requestData,items:lines,source:{type:'outlook',message_id:state.mail.messageId},warnings:proposal.warnings||[]},offer_text:`${proposal.introduction}\n\n${proposal.closing}`,change_note:'Automatisch aus Outlook erstellt – Mitarbeiterprüfung erforderlich',net_total:totals.net,tax_total:totals.tax,gross_total:totals.gross,created_by:state.user.id})});
}

async function createOfferReply(){
  $('importBtn').disabled=true; $('importBtn').textContent='E-Mail wird analysiert …'; message('Serviq liest die Anfrage aus und kalkuliert mit dem Leistungskatalog …');
  try{
    const request=await ai('analyze_email',{text:state.mail.body});
    const guestCount=Math.max(1,Number(request.guestCount)||1);
    $('importBtn').textContent='Angebot wird kalkuliert …';
    const proposal=await ai('generate_quote',{request,items:state.catalog.map(item=>({id:item.external_id,name:item.name,detail:item.detail,category:item.category,unit:item.unit,defaultQty:item.default_qty,price:item.price,defaultSelected:item.default_selected})),company:state.company,rules:state.rules});
    const lines=calculateLines(proposal,guestCount);
    if(!lines.length)throw new Error('Die KI konnte keine passende Leistung sicher auswählen. Bitte den Leistungskatalog in Serviq prüfen.');
    const net=lines.reduce((sum,item)=>sum+item.total,0); const vatRate=Number(state.company.vatRate??19); const tax=net*vatRate/100;
    const number=`AN-${new Date().toISOString().slice(0,10).replaceAll('-','')}-${String(Date.now()).slice(-5)}`;
    const quote={number,request:{...request,guestCount},lines,net,tax,gross:net+tax,vatRate,introduction:proposal.introduction||'vielen Dank für Ihre Anfrage. Gern unterbreiten wir Ihnen folgendes Angebot.',closing:proposal.closing||'Für Rückfragen und Anpassungswünsche stehen wir Ihnen gerne zur Verfügung.'};
    $('importBtn').textContent='PDF und Antwort werden erstellt …';
    await saveQuote(request,proposal,lines,{net,tax,gross:net+tax},number);
    await openReply(quote,makePdf(quote));
    message('Antwortentwurf mit Angebots-PDF geöffnet. Bitte vor dem Versand fachlich prüfen.','ok');
  }catch(error){message(friendlyError(error),'error');}
  finally{$('importBtn').disabled=false;$('importBtn').textContent='Angebot erstellen';}
}

$('signInForm').addEventListener('submit',beginSignIn);
$('microsoftSignInBtn').addEventListener('click',beginMicrosoftSignIn);
$('forgotPasswordBtn').addEventListener('click',()=>Office.context.ui.openBrowserWindow(`${SERVIQ_URL}/?auth=recover`));
$('signOutBtn').addEventListener('click',()=>{clearSession();show('signinView');});
$('importBtn').addEventListener('click',createOfferReply);
$('openServiqBtn').addEventListener('click',()=>Office.context.ui.openBrowserWindow(SERVIQ_URL));
Office.onReady(info=>{if(info.host===Office.HostType.Outlook)initialize();else show('unsupportedView');});
