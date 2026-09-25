const SUPABASE_URL = 'https://rkhptbohniykwxhwhjiz.supabase.co';
const SUPABASE_KEY = 'sb_publishable_I3Pz0TnaVqRt6GdszYq_Jw_BETkoJAk';
const AI_URL = `${SUPABASE_URL}/functions/v1/serviq-ai`;
const PDF_ATTACHMENT_URL = `${SUPABASE_URL}/functions/v1/outlook-pdf-attachment`;
const SERVIQ_URL = 'https://serviq-catering-ai-ogu.oguzhan-yoeruerer.chatgpt.site';
const CALLBACK_URL = `${location.origin}/auth-callback-v3.html`;
const HANDOFF_URL = `${SUPABASE_URL}/functions/v1/outlook-auth-handoff`;
const TOKEN_KEY = 'serviq_access_token';
const REFRESH_KEY = 'serviq_refresh_token';
const OAUTH_BRIDGE_KEY = 'serviq_oauth_bridge_v2';
const OAUTH_CHANNEL = 'serviq_oauth_channel_v2';
const state = { token: localStorage.getItem(TOKEN_KEY) || '', refreshToken: localStorage.getItem(REFRESH_KEY) || '', user: null, membership: null, mail: null, company: {}, rules: {}, catalog: [], pendingQuote: null };
const $ = id => document.getElementById(id);
const euro = value => new Intl.NumberFormat('de-DE',{style:'currency',currency:'EUR'}).format(value);
const html = value => String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
const DEFAULT_TEMPLATE={
  title:'Freibleibendes Angebot',
  introduction:'Guten Tag {{ansprechpartner}},\n\nvielen Dank für Ihre Anfrage. Für {{veranstaltung}} am {{datum}} haben wir ein Cateringkonzept für {{gaeste}} Gäste zusammengestellt.',
  closing:'Für Rückfragen und Anpassungswünsche stehen wir Ihnen gerne zur Verfügung.',
  disclaimer:'Dieses Angebot ist freibleibend und unverbindlich. Es steht unter dem Vorbehalt der Verfügbarkeit sowie der abschließenden Prüfung und Bestätigung durch den Anbieter. Änderungen der Gästezahl oder des Leistungsumfangs können zu Preisanpassungen führen.',
  emailGreeting:'Guten Tag {{ansprechpartner}},',
  emailBody:'vielen Dank für Ihre Anfrage. Im Anhang finden Sie unser freibleibendes Angebot {{angebotsnummer}}. Bitte prüfen Sie die enthaltenen Leistungen und Konditionen.',
  emailClosing:'Freundliche Grüße\n{{unternehmen}}'
};
const dateDE=value=>{if(!value)return'nach Vereinbarung';const date=new Date(`${value}T12:00:00`);return Number.isNaN(date.getTime())?String(value):new Intl.DateTimeFormat('de-DE').format(date);};
function templateContext(request,number){return{ansprechpartner:request.contactName||state.mail?.fromName||'Damen und Herren',veranstaltung:request.eventName||state.mail?.subject||'Ihre Veranstaltung',datum:dateDE(request.eventDate),gaeste:Number(request.guestCount)||'',angebotsnummer:number||'',unternehmen:state.company.name||state.membership?.organizations?.name||'Ihr Catering-Team'};}
function renderTemplate(value,context){return String(value||'').replace(/{{\s*(ansprechpartner|veranstaltung|datum|gaeste|angebotsnummer|unternehmen)\s*}}/gi,(_,key)=>String(context[key.toLowerCase()]??''));}
function template(){return{...DEFAULT_TEMPLATE,...(state.company.offerTemplate||{})};}
const paragraphs=value=>String(value||'').split(/\n{2,}/).map(part=>`<p>${html(part).replace(/\n/g,'<br>')}</p>`).join('');
const fromBase64Url=value=>Uint8Array.from(atob(value.replace(/-/g,'+').replace(/_/g,'/').padEnd(Math.ceil(value.length/4)*4,'=')),c=>c.charCodeAt(0));
const randomToken=(size=32)=>{const bytes=crypto.getRandomValues(new Uint8Array(size));return btoa(String.fromCharCode(...bytes)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');};
async function decryptHandoff(ciphertext,iv,key){
  const cryptoKey=await crypto.subtle.importKey('raw',fromBase64Url(key),'AES-GCM',false,['decrypt']);
  const clear=await crypto.subtle.decrypt({name:'AES-GCM',iv:fromBase64Url(iv)},cryptoKey,fromBase64Url(ciphertext));
  return JSON.parse(new TextDecoder().decode(clear));
}

function show(id){ ['signinView','importView','reviewView','unsupportedView'].forEach(x => $(x).classList.toggle('hidden', x !== id)); }
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
  const body=await response.text();
  return body?JSON.parse(body):null;
}
async function ai(operation,payload){
  const response=await fetch(AI_URL,{method:'POST',headers:{apikey:SUPABASE_KEY,Authorization:`Bearer ${state.token}`,'Content-Type':'application/json'},body:JSON.stringify({operation,...payload,consent_to_openai:true})});
  const result=await response.json().catch(()=>({}));
  if(!response.ok) throw new Error(result.error||'Die KI-Verarbeitung ist fehlgeschlagen.');
  return result.data;
}

async function loadAccount(){
  state.user=await api('/auth/v1/user');
  const rows=await api(`/rest/v1/organization_members?user_id=eq.${encodeURIComponent(state.user.id)}&select=organization_id,display_name,role,organizations(name)&limit=1`);
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
  localStorage.removeItem(OAUTH_BRIDGE_KEY);
  const channel=randomToken();const handoffKey=randomToken();
  const callback=`${CALLBACK_URL}?channel=${encodeURIComponent(channel)}&key=${encodeURIComponent(handoffKey)}`;
  const startUrl=`${location.origin}/auth-start.html?callback=${encodeURIComponent(callback)}`;
  Office.context.ui.displayDialogAsync(startUrl,{height:70,width:40,displayInIframe:false},result=>{
    if(result.status!==Office.AsyncResultStatus.Succeeded){
      button.disabled=false;status.textContent=`Anmeldefenster konnte nicht geöffnet werden: ${result.error.message}`;status.className='message error';return;
    }
    const dialog=result.value;
    let finished=false;
    let bridgeChannel=null;
    let bridgePoll=null;
    let handoffPoll=null;
    let handoffTimeout=null;
    const cleanup=()=>{
      if(bridgePoll)clearInterval(bridgePoll);
      if(bridgeChannel)bridgeChannel.close();
      if(handoffPoll)clearInterval(handoffPoll);
      if(handoffTimeout)clearTimeout(handoffTimeout);
      localStorage.removeItem(OAUTH_BRIDGE_KEY);
    };
    const acceptPayload=async raw=>{
      if(finished)return;
      let payload={};
      try{payload=typeof raw==='string'?JSON.parse(raw):raw;}catch(_){payload={error:'Ungültige Antwort der Anmeldung.'};}
      if(!payload?.access_token&&!payload?.error)return;
      finished=true;cleanup();dialog.close();button.disabled=false;
      if(payload.error){status.textContent=payload.error;status.className='message error';return;}
      storeSession(payload);await initialize();
    };
    dialog.addEventHandler(Office.EventType.DialogMessageReceived,event=>{
      acceptPayload(event.message);
    });
    try{
      bridgeChannel=new BroadcastChannel(OAUTH_CHANNEL);
      bridgeChannel.onmessage=event=>acceptPayload(event.data);
    }catch(_){/* Polling bleibt als kompatibler Fallback aktiv. */}
    bridgePoll=setInterval(()=>{
      const bridged=localStorage.getItem(OAUTH_BRIDGE_KEY);
      if(bridged)acceptPayload(bridged);
    },300);
    const pollHandoff=async()=>{
      try{
        const response=await fetch(HANDOFF_URL,{method:'POST',headers:{apikey:SUPABASE_KEY,'Content-Type':'application/json'},body:JSON.stringify({action:'take',channel})});
        if(response.status===202)return;
        const result=await response.json().catch(()=>({}));
        if(!response.ok)throw new Error(result.error||'Anmeldesitzung konnte nicht übernommen werden.');
        if(result.ciphertext&&result.iv)await acceptPayload(await decryptHandoff(result.ciphertext,result.iv,handoffKey));
      }catch(error){
        if(!finished){status.textContent=friendlyError(error);status.className='message error';}
      }
    };
    handoffPoll=setInterval(pollHandoff,1000);pollHandoff();
    handoffTimeout=setTimeout(()=>{if(!finished){cleanup();button.disabled=false;status.textContent='Die Anmeldung ist abgelaufen. Bitte erneut versuchen.';status.className='message error';}},5*60*1000);
    dialog.addEventHandler(Office.EventType.DialogEventReceived,()=>{if(!finished)status.textContent='Microsoft-Anmeldung wird übernommen …';});
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

function replyHtml(quote){
  const rows=quote.lines.map(item=>`<tr><td style="padding:6px;border-bottom:1px solid #ddd">${html(item.name)}</td><td style="padding:6px;text-align:right;border-bottom:1px solid #ddd">${html(item.quantity)} ${html(item.unit)}</td><td style="padding:6px;text-align:right;border-bottom:1px solid #ddd">${html(euro(item.total))}</td></tr>`).join('');
  return `${paragraphs(quote.emailGreeting)}${paragraphs(quote.emailBody)}<table style="border-collapse:collapse;width:100%"><tbody>${rows}<tr><td colspan="2" style="padding:8px;text-align:right"><b>Gesamt inkl. MwSt.</b></td><td style="padding:8px;text-align:right"><b>${html(euro(quote.gross))}</b></td></tr></tbody></table>${paragraphs(quote.emailClosing)}<p><small>${html(quote.disclaimer)}</small></p>`;
}
async function preparePdfAttachment(quote){
  const response=await fetch(PDF_ATTACHMENT_URL,{method:'POST',headers:{apikey:SUPABASE_KEY,Authorization:`Bearer ${state.token}`,'Content-Type':'application/json'},body:JSON.stringify({quote,company:state.company,organizationName:state.membership.organizations?.name||'',filename:`Angebot-${quote.number}.pdf`})});
  const result=await response.json().catch(()=>({}));
  if(!response.ok||!result.url)throw new Error(result.error||'Die Angebots-PDF konnte nicht für Outlook bereitgestellt werden.');
  return result.url;
}
function openReply(quote,pdfUrl){
  return new Promise((resolve,reject)=>{
    if(!Office.context.requirements.isSetSupported('Mailbox','1.9')) return reject(new Error('Diese Outlook-Version unterstützt das Öffnen des Antwortentwurfs noch nicht. Bitte Outlook im Web oder die aktuelle Desktop-Version verwenden.'));
    if(typeof pdfUrl!=='string'||!pdfUrl.startsWith('https://')) return reject(new Error('Die Download-Adresse der Angebots-PDF ist ungültig.'));
    const attachment={url:pdfUrl,name:`Angebot-${quote.number}.pdf`,type:'file',inLine:false};
    try{
      Office.context.mailbox.item.displayReplyFormAsync({htmlBody:replyHtml(quote),attachments:[attachment]},result=>{
        if(result.status===Office.AsyncResultStatus.Succeeded)resolve(); else reject(new Error(result.error?.message||'Der Outlook-Antwortentwurf konnte nicht geöffnet werden.'));
      });
    }catch(error){reject(new Error(`Outlook konnte den Antwortentwurf nicht öffnen: ${error?.message||'unbekannter Fehler'}`));}
  });
}

async function saveQuote(requestData,proposal,lines,totals,number){
  const requestPayload={organization_id:state.membership.organization_id,source_type:'outlook',source_name:state.mail.subject,raw_text:state.mail.body,customer_company:requestData.customerCompany||'',contact_name:requestData.contactName||state.mail.fromName,contact_email:requestData.contactEmail||state.mail.from,event_name:requestData.eventName||state.mail.subject,event_date:requestData.eventDate||null,event_time:requestData.eventTime||'',venue:requestData.venue||'',guest_count:Number(requestData.guestCount)||1,preferences:requestData.dietary||'',notes:$('note').value.trim()||requestData.notes||'',extracted_data:{...requestData,message_id:state.mail.messageId},missing_fields:requestData.missing_fields||[],confidence:{uncertain_fields:requestData.uncertain_fields||[]},status:'quoted',created_by:state.user.id};
  const requests=await api('/rest/v1/customer_requests',{method:'POST',headers:{Prefer:'return=representation'},body:JSON.stringify(requestPayload)});
  const offers=await api('/rest/v1/offers',{method:'POST',headers:{Prefer:'return=representation'},body:JSON.stringify({organization_id:state.membership.organization_id,request_id:requests[0].id,offer_number:number,status:'needs_review',net_total:totals.net,tax_total:totals.tax,gross_total:totals.gross,created_by:state.user.id})});
  await api('/rest/v1/offer_versions',{method:'POST',headers:{Prefer:'return=minimal'},body:JSON.stringify({organization_id:state.membership.organization_id,offer_id:offers[0].id,version_no:1,snapshot:{request:requestData,items:lines,adjustment:state.pendingQuote?.adjustment||null,source:{type:'outlook',message_id:state.mail.messageId},warnings:proposal.warnings||[]},offer_text:`${proposal.introduction}\n\n${proposal.closing}`,change_note:'In Outlook erstellt und vor der PDF-Erstellung manuell geprüft',net_total:totals.net,tax_total:totals.tax,gross_total:totals.gross,created_by:state.user.id})});
}

function recalculatePendingQuote(){
  const quote=state.pendingQuote;if(!quote)return;
  quote.lines.forEach(line=>{line.quantity=Math.max(0,Number(line.quantity)||0);line.price=Math.max(0,Number(line.price)||0);line.total=line.quantity*line.price;});
  const base=quote.lines.reduce((sum,line)=>sum+line.total,0);
  const kind=$('reviewAdjustmentType').value;const value=Math.max(0,Number($('reviewAdjustmentValue').value)||0);
  const adjustment=kind==='discount'?-Math.min(base,base*value/100):kind==='surcharge'?base*value/100:0;
  quote.adjustment={type:kind,value,amount:adjustment};quote.net=Math.max(0,base+adjustment);quote.tax=quote.net*quote.vatRate/100;quote.gross=quote.net+quote.tax;
  $('reviewGross').textContent=euro(quote.gross);
}

function renderQuoteReview(){
  const quote=state.pendingQuote;if(!quote)return;
  $('reviewLines').innerHTML=quote.lines.map((line,index)=>`<div class="review-line" data-index="${index}"><input data-field="name" value="${html(line.name)}" aria-label="Bezeichnung"/><div><input data-field="quantity" type="number" min="0" step="0.1" value="${line.quantity}" aria-label="Menge"/><input data-field="price" type="number" min="0" step="0.01" value="${line.price}" aria-label="Preis netto"/><button data-remove type="button" aria-label="Position entfernen">×</button></div></div>`).join('');
  document.querySelectorAll('.review-line').forEach(row=>{
    const index=Number(row.dataset.index);row.querySelectorAll('[data-field]').forEach(input=>input.addEventListener('input',()=>{state.pendingQuote.lines[index][input.dataset.field]=input.dataset.field==='name'?input.value:Number(input.value);recalculatePendingQuote();}));
    row.querySelector('[data-remove]').addEventListener('click',()=>{state.pendingQuote.lines.splice(index,1);renderQuoteReview();});
  });
  recalculatePendingQuote();show('reviewView');
}

function addReviewLine(event){
  event.preventDefault();if(!state.pendingQuote)return;
  const name=$('reviewAddName').value.trim();const quantity=Math.max(0,Number($('reviewAddQty').value)||0);const price=Math.max(0,Number($('reviewAddPrice').value)||0);
  if(!name)return;
  state.pendingQuote.lines.push({id:`manual-${Date.now()}`,name,detail:'Manuell ergänzt',unit:'pauschal',quantity,price,total:quantity*price,rationale:'Manuell ergänzt'});
  event.currentTarget.reset();$('reviewAddQty').value='1';renderQuoteReview();
}

async function finalizeOffer(){
  const quote=state.pendingQuote;if(!quote||!quote.lines.length){$('reviewMessage').textContent='Bitte mindestens eine Position ergänzen.';$('reviewMessage').className='message error';return;}
  recalculatePendingQuote();$('finalizeOfferBtn').disabled=true;$('finalizeOfferBtn').textContent='PDF und Antwort werden erstellt …';
  try{
    await saveQuote(quote.request,quote.proposal,quote.lines,{net:quote.net,tax:quote.tax,gross:quote.gross},quote.number);
    const pdfUrl=await preparePdfAttachment(quote);await openReply(quote,pdfUrl);
    $('reviewMessage').textContent='Antwortentwurf mit Angebots-PDF geöffnet. Bitte vor dem Versand prüfen.';$('reviewMessage').className='message ok';
  }catch(error){$('reviewMessage').textContent=friendlyError(error);$('reviewMessage').className='message error';}
  finally{$('finalizeOfferBtn').disabled=false;$('finalizeOfferBtn').textContent='Antwort mit PDF öffnen';}
}

async function createOfferReply(){
  $('importBtn').disabled=true; $('importBtn').textContent='E-Mail wird analysiert …'; message('Serviq liest die Anfrage aus und kalkuliert mit dem Leistungskatalog …');
  let stage='Analyse der E-Mail';
  try{
    const request=await ai('analyze_email',{text:state.mail.body});
    const guestCount=Math.max(1,Number(request.guestCount)||1);
    $('importBtn').textContent='Angebot wird kalkuliert …';
    stage='Kalkulation des Angebots';
    const proposal=await ai('generate_quote',{request,items:state.catalog.map(item=>({id:item.external_id,name:item.name,detail:item.detail,category:item.category,unit:item.unit,defaultQty:item.default_qty,price:item.price,defaultSelected:item.default_selected})),company:state.company,rules:state.rules});
    const lines=calculateLines(proposal,guestCount);
    if(!lines.length)throw new Error('Die KI konnte keine passende Leistung sicher auswählen. Bitte den Leistungskatalog in Serviq prüfen.');
    const net=lines.reduce((sum,item)=>sum+item.total,0); const vatRate=Number(state.company.vatRate??19); const tax=net*vatRate/100;
    const number=`AN-${new Date().toISOString().slice(0,10).replaceAll('-','')}-${String(Date.now()).slice(-5)}`;
    const context=templateContext({...request,guestCount},number);const textTemplate=template();
    const quote={number,request:{...request,guestCount},lines,net,tax,gross:net+tax,vatRate,title:renderTemplate(textTemplate.title,context),introduction:renderTemplate(textTemplate.introduction,context)||proposal.introduction,closing:renderTemplate(textTemplate.closing,context)||proposal.closing,disclaimer:renderTemplate(textTemplate.disclaimer,context),emailGreeting:renderTemplate(textTemplate.emailGreeting,context),emailBody:renderTemplate(textTemplate.emailBody,context),emailClosing:renderTemplate(textTemplate.emailClosing,context)};
    proposal.introduction=quote.introduction;proposal.closing=quote.closing;quote.proposal=proposal;
    state.pendingQuote=quote;renderQuoteReview();
  }catch(error){message(`${stage}: ${friendlyError(error)}`,'error');}
  finally{$('importBtn').disabled=false;$('importBtn').textContent='Angebot erstellen';}
}

$('signInForm').addEventListener('submit',beginSignIn);
$('microsoftSignInBtn').addEventListener('click',beginMicrosoftSignIn);
$('forgotPasswordBtn').addEventListener('click',()=>Office.context.ui.openBrowserWindow(`${SERVIQ_URL}/?auth=recover`));
$('signOutBtn').addEventListener('click',()=>{clearSession();show('signinView');});
$('importBtn').addEventListener('click',createOfferReply);
$('reviewAddForm').addEventListener('submit',addReviewLine);
$('reviewAdjustmentType').addEventListener('change',recalculatePendingQuote);
$('reviewAdjustmentValue').addEventListener('input',recalculatePendingQuote);
$('finalizeOfferBtn').addEventListener('click',finalizeOffer);
$('cancelReviewBtn').addEventListener('click',()=>show('importView'));
$('openServiqBtn').addEventListener('click',()=>Office.context.ui.openBrowserWindow(SERVIQ_URL));
Office.onReady(info=>{if(info.host===Office.HostType.Outlook)initialize();else show('unsupportedView');});
