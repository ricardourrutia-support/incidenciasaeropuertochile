/**
 * ============================================================================
 *  SEGUIMIENTO DE VENTAS AEROPUERTO  ·  Google Apps Script Web App
 * ============================================================================
 *  Lee la Base de Comisiones + Turnos 360 (hojas "Coordinadores Mayo",
 *  "Agentes Junio" [agentes = ejecutivos de venta], "Anfitriones ...",
 *  "Supervisores ..."). Cada celda de turno es "05:00 - 16:00" / "21:00 - 08:00"
 *  / "Libre".
 *
 *  Reglas:
 *   - ROL del vendedor: se deduce de en que hoja de Turnos aparece (o del Roster).
 *     Secciones separadas: EJECUTIVOS (sus ventas) y COORDINADORES (comision).
 *   - COMISION coordinadores: por sus HORAS ACTIVAS (turno menos horas
 *     administrativas/colacion del Cronograma V77). 2 activos -> 50/50 (N -> 1/N).
 *   - EQUIPO Diurno/Nocturno = turno del EJECUTIVO vendedor (cruza medianoche =
 *     Nocturno). NO depende de la franja: una misma hora puede tener ventas
 *     diurnas y nocturnas a la vez (un ejecutivo diurno y uno nocturno).
 *   - "Otras Ventas": ventas hechas por NO-ejecutivos (supervisor/coordinador/
 *     anfitrion/otro).
 *   - Correos mal ingresados: atributo ds_user_email (solo van compartida).
 *   - METAS: meta grupal mensual + metas personales (ejecutivos y coordinadores),
 *     con grafico de avance.
 * ============================================================================
 */

var CONFIG = {
  VENTAS_SPREADSHEET_ID: '1i5TjTE34M8jeGYKJ7jBbw8KqS6k5YpOTMr5mEO-jR00',
  VENTAS_SHEET: 'Results',

  TURNOS_SPREADSHEET_ID: '1xepmv4-ocTNZ-RXBa7pFBXp4KM56ZTrAgeO5M0B7yCg',
  ANIO: 2026,
  PREFIJO_AGENTES: 'Agentes',           // = ejecutivos de venta
  PREFIJO_COORDINADORES: 'Coordinadores',
  PREFIJO_ANFITRIONES: 'Anfitriones',
  PREFIJO_SUPERVISORES: 'Supervisores',

  ROSTER_SHEET: 'Roster',
  // El Roster es OPCIONAL: si la pestana no existe, la app resuelve todo desde
  // Turnos 360 (rol por mes, nombre y equipo) cruzando por el correo. El Roster,
  // si existe, actua como refuerzo para correos que no sigan ninguna convencion.
  // Alias manual (opcional) correo -> nombre tal como aparece en Turnos 360:
  ALIAS_CORREOS: { /* 'correo.raro@cabify.com': 'Nombre En Turnos' */ },
  // Libro donde esta la pestana Roster. Si la dejas '', usa el libro donde esta
  // instalado este script (o, si no, el de ventas). Si tu Roster esta en OTRO
  // libro (p.ej. "[CL] AIRPORT SALES"), pega aqui su ID (lo sacas de la URL).
  ROSTER_SPREADSHEET_ID: '',

  COL_FECHA_VENTA: 'tm_start_local_at',
  PRODUCTO_COMPARTIDA: 'van_compartida',
  PRODUCTO_EXCLUSIVA:  'van_exclusive',
  FINISH_OK: 'FINISH_REASON_DROPOFF',
  SOLO_CONCRETADAS: true,

  // Cronograma V77: horas administrativas/colacion por hora de ingreso (no comisionan).
  ADMIN_POR_INGRESO: { 5:[11,12,13], 10:[10,14,15], 13:[17,18,19], 21:[5,6,7] },

  META_PROP: 'METAS_VENTAS_JSON',

  // --- Desempeno de agentes (Tableau crosstab) ---
  TABLEAU_SERVER:      'https://tableau.cabify-data.com',
  TABLEAU_API_VERSION: '3.19',
  TABLEAU_SITE:        'cabify',
  TABLEAU_PAT_NAME:    'B2B Support',
  TABLEAU_PAT_SECRET:  'KI/Fpf9RTVaA8F2oH/vqKQ==:kL07mNjXKGkJcnPJT4v4BjB0jXxmOFuI',
  // contentUrl de la vista: workbook/sheets/vista (de tu link PerformanceSupportCrosstab / Sheet1)
  TABLEAU_VIEW_CONTENT_URL: 'PerformanceSupportCrosstab/sheets/Sheet1',
  PERF_SHEET: 'Performance',            // pestana cache donde se vuelca la crosstab
  PERF_PROP:  'PERF_ULTIMA_ACTUALIZACION',
  FIRT_UMBRAL_H: 24,    // cumple si la primera respuesta es antes de 24 horas
  FURT_UMBRAL_H: 120,   // cumple si la resolucion es antes de 120 horas

  // --- Datos maestros (planilla BUK de colaboradores activos) ---
  MAESTROS_SHEET: 'Maestros',           // pestana cache sincronizada desde la planilla
  MAESTROS_PROP:  'MAESTROS_SYNC'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate().setTitle('Seguimiento de Ventas · Aeropuerto')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ------------------------------- utilidades --------------------------------
function ss_(id){ return SpreadsheetApp.openById(id || CONFIG.VENTAS_SPREADSHEET_ID); }
function norm_(s){ return (s==null?'':String(s)).normalize('NFKD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase(); }
function pad2_(n){ return (n<10?'0':'')+n; }
function dateKey_(d){ return d.getFullYear()+'-'+pad2_(d.getMonth()+1)+'-'+pad2_(d.getDate()); }
function addDays_(d,n){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()+n); }
function round_(n,d){ d=(d==null?0:d); var f=Math.pow(10,d); return Math.round((Number(n)||0)*f)/f; }
function firstOfMonth_(){ var n=new Date(); return new Date(n.getFullYear(), n.getMonth(), 1); }
function endOfDay_(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23,59,59); }
function getColMap_(h){ var m={}; h.forEach(function(x,i){ m[String(x).trim()]=i; }); return m; }
function nameKey_(s){ var t=norm_(s).split(/\s+/).filter(Boolean); return !t.length?'':(t.length===1?t[0]:t[0]+' '+t[t.length-1]); }
function keyFromEmail_(e){ var t=String(e).split('@')[0].split(/[._]/).filter(Boolean); return !t.length?'':(t.length===1?norm_(t[0]):norm_(t[0])+' '+norm_(t[t.length-1])); }

function parseDate_(v){
  if (v instanceof Date && !isNaN(v)) return v;
  if (v==null || v==='') return null;
  var s=String(v).trim();
  var m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (m) return new Date(+m[3],+m[2]-1,+m[1], m[4]?+m[4]:0, m[5]?+m[5]:0);
  m=s.toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g,'').match(/^(\d{1,2}) de ([a-z]+) de (\d{4})$/);
  if (m && MES_FULL_[m[2]]!=null) return new Date(+m[3], MES_FULL_[m[2]], +m[1]);
  var d=new Date(s); return isNaN(d)?null:d;
}
function parseTurno_(v){
  if (v==null) return {libre:true};
  var s=String(v).trim().toLowerCase();
  if (!s||s==='libre'||s==='.'||s==='-'||s.indexOf('vacac')>=0) return {libre:true};
  var m=s.match(/(\d{1,2})[:\.]?\d{0,2}\s*[-a]\s*(\d{1,2})[:\.]?\d{0,2}/);
  if (!m) return {libre:true};
  var st=+m[1], en=+m[2]; if (st>23||en>24) return {libre:true};
  return { libre:false, start:st, end:en, cross:(en<=st) };
}
var MES_={ene:0,feb:1,mar:2,abr:3,may:4,jun:5,jul:6,ago:7,sep:8,set:8,oct:9,nov:10,dic:11};
function parseHeaderDate_(v){
  if (v instanceof Date && !isNaN(v)) return v;
  if (v==null) return null;
  var s=String(v).trim().toLowerCase();
  var m=s.match(/^(\d{1,2})[\-\/ ]([a-záéíóú]{3,})/);
  if (m && MES_[m[2].substring(0,3)]!=null) return new Date(CONFIG.ANIO, MES_[m[2].substring(0,3)], +m[1]);
  m=s.match(/^(\d{1,2})[\-\/](\d{1,2})$/);
  if (m) return new Date(CONFIG.ANIO, (+m[2])-1, +m[1]);
  return null;
}

// ------------------------------- validez email -----------------------------
var VALID_DOM_={'gmail.com':1,'hotmail.com':1,'hotmail.es':1,'hotmail.cl':1,'outlook.com':1,'outlook.es':1,'outlook.cl':1,'yahoo.com':1,'yahoo.es':1,'yahoo.cl':1,'yahoo.com.ar':1,'yahoo.com.br':1,'icloud.com':1,'live.com':1,'live.cl':1,'me.com':1,'mail.com':1,'msn.com':1,'vtr.net':1,'uc.cl':1,'usach.cl':1,'udd.cl':1,'miuandes.cl':1,'ug.uchile.cl':1,'fen.uchile.cl':1,'mayor.cl':1,'latam.com':1,'duocuc.cl':1,'santotomas.cl':1,'unab.cl':1,'protonmail.com':1,'aol.com':1,'proton.me':1,'flesan.cl':1};
var TYPO_DOM_={'gamil.com':1,'gmial.com':1,'gmai.com':1,'gmail.om':1,'gmailc.om':1,'gmaill.com':1,'gmail.cmo':1,'gail.com':1,'gmil.com':1,'gnail.com':1,'gmal.com':1,'gmail.co':1,'gmail.cl':1,'gmaul.com':1,'gmali.com':1,'hotmial.com':1,'hotmai.com':1,'hotmal.com':1,'outlok.com':1,'yaho.com':1,'notiene.cl':1,'gmail.es':1};
var PLACEHOLDER_={'test':1,'prueba':1,'na':1,'no':1,'notiene':1,'nocorreo':1,'sincorreo':1,'xxx':1,'asdf':1,'qwerty':1,'none':1,'nomail':1,'noemail':1,'sin':1,'correo':1,'x':1,'aaa':1};
var TLD_OK_={'com':1,'cl':1,'es':1,'net':1,'org':1,'ar':1,'br':1,'pe':1,'co':1,'io':1,'edu':1,'gob':1,'gov':1};
function emailOk_(e){
  e=(e==null?'':String(e)).trim().toLowerCase();
  if(!e||e==='nan'||e.indexOf('@')<0) return false;
  if(!/^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$/.test(e)) return false;
  var p=e.split('@'),loc=p[0],dom=p[1];
  if(TYPO_DOM_[dom]) return false;
  if(!VALID_DOM_[dom]){ if(/(gm|gn|ga|hotm|outl|yah)/.test(dom)) return false; if(!TLD_OK_[dom.split('.').pop()]) return false; }
  if(/^\d+$/.test(loc)) return false;
  var u={}; for(var i=0;i<loc.length;i++) u[loc[i]]=1; if(Object.keys(u).length===1) return false;
  if(loc.length<=2) return false;
  if(PLACEHOLDER_[loc]) return false;
  if(/^[a-z]+$/.test(loc)&&loc.length<=6&&loc.indexOf('.')<0) return false;
  return true;
}

// ------------------------------- carga -------------------------------------
function rosterBook_(){
  if(CONFIG.ROSTER_SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.ROSTER_SPREADSHEET_ID);
  try{ var a=SpreadsheetApp.getActiveSpreadsheet(); if(a) return a; }catch(e){}
  return ss_(CONFIG.VENTAS_SPREADSHEET_ID);
}
function loadRoster_(){
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.ROSTER_SHEET)||book.getSheetByName('Roster seed')||book.getSheetByName('Roster_seed');
  var byEmail={}, keyByEmail={};
  if(!sh) return {byEmail:byEmail, keyByEmail:keyByEmail};
  var v=sh.getDataRange().getValues(); if(v.length<2) return {byEmail:byEmail, keyByEmail:keyByEmail};
  // Si el CSV quedo pegado en una sola columna (no dividido por comas), dividirlo.
  if(v[0].length<2 || (String(v[0][0]).indexOf(',')>=0 && (''+(v[0][1]||'')).trim()==='')){
    v=v.map(function(r){ return String(r[0]==null?'':r[0]).split(','); });
  }
  // Buscar la fila de encabezado (por si hay filas arriba): la que contenga 'Email_Cabify'.
  var hr=-1; for(var k=0;k<Math.min(8,v.length);k++){
    if(v[k].some(function(x){return String(x).trim()==='Email_Cabify';})){ hr=k; break; } }
  if(hr<0) hr=0;
  var c=getColMap_(v[hr]);
  if(c['Email_Cabify']==null) return {byEmail:byEmail, keyByEmail:keyByEmail};
  for(var i=hr+1;i<v.length;i++){
    var email=String(v[i][c['Email_Cabify']]||'').trim().toLowerCase(); if(!email||email.indexOf('@')<0) continue;
    var nombre=v[i][c['Nombre']]||email;
    byEmail[email]={nombre:nombre, rol:(v[i][c['Rol']]||'').toString().trim(), cargo:v[i][c['Cargo']]||'', supervisor:v[i][c['Supervisor']]||''};
    keyByEmail[email]=nameKey_(nombre);
  }
  return {byEmail:byEmail, keyByEmail:keyByEmail};
}

function loadVentas_(){
  var sh=ss_(CONFIG.VENTAS_SPREADSHEET_ID).getSheetByName(CONFIG.VENTAS_SHEET);
  if(!sh) throw new Error('No existe la pestana de ventas: '+CONFIG.VENTAS_SHEET);
  var v=sh.getDataRange().getValues(), c=getColMap_(v[0]);
  var iD=c[CONFIG.COL_FECHA_VENTA], iA=c['ds_agent_email'], iU=c['ds_user_email'],
      iM=c['qt_price_local'], iP=c['ds_product_name'], iF=c['finishReason'];
  var rows=[];
  for(var i=1;i<v.length;i++){
    var d=parseDate_(v[i][iD]); if(!d) continue;
    rows.push({date:d, hour:d.getHours(), dateKey:dateKey_(d),
      agent:String(v[i][iA]||'').trim().toLowerCase(), user:String(v[i][iU]||'').trim().toLowerCase(),
      amount:Number(v[i][iM])||0, product:String(v[i][iP]||'').trim(), fin:String(v[i][iF]||'').trim()});
  }
  return rows;
}

function loadTurnosPrefijo_(prefijo){
  var ssT=ss_(CONFIG.TURNOS_SPREADSHEET_ID||CONFIG.VENTAS_SPREADSHEET_ID);
  var pref=norm_(prefijo), people={};
  ssT.getSheets().forEach(function(sh){
    if(norm_(sh.getName()).indexOf(pref)!==0) return;
    var v=sh.getDataRange().getValues(); if(!v.length) return;
    var dr=-1, best=0;
    for(var r=0;r<Math.min(8,v.length);r++){ var cnt=0;
      for(var c=1;c<v[r].length;c++) if(parseHeaderDate_(v[r][c])) cnt++;
      if(cnt>best){ best=cnt; dr=r; } }
    if(dr<0||best<3) return;
    var dateCols=[]; for(var c2=1;c2<v[dr].length;c2++) if(parseHeaderDate_(v[dr][c2])) dateCols.push(c2);
    for(var r2=dr+1;r2<v.length;r2++){
      var nm=v[r2][0]; if(!nm) continue; var s=String(nm).trim();
      if(!s||s==='.'||s==='Nombre'||s==='Cargo'||s==='Supervisor') continue;
      var key=nameKey_(s); var rec=people[key]||(people[key]={name:s, shifts:{}});
      dateCols.forEach(function(dc){
        var t=parseTurno_(v[r2][dc]); if(t.libre) return;
        var hd=parseHeaderDate_(v[dr][dc]); if(!hd) return;
        rec.shifts[dateKey_(hd)]=t;
      });
    }
  });
  return people;
}

function shiftHours_(start,end){ var out=[],h=start,off=0,g=0;
  while(h!==end && g<26){ out.push({off:off,h:h}); h++; if(h===24){h=0;off=1;} g++; } return out; }
function adminHoras_(start){ return CONFIG.ADMIN_POR_INGRESO[start] || [(start+4)%24,(start+5)%24,(start+6)%24]; }

// --------------------- DESEMPENO (Tableau crosstab) -------------------------
function tableauUrl_(path){ return CONFIG.TABLEAU_SERVER+'/api/'+CONFIG.TABLEAU_API_VERSION+path; }

function tableauSignIn_(){
  var resp=UrlFetchApp.fetch(tableauUrl_('/auth/signin'),{
    method:'post', contentType:'application/json',
    headers:{Accept:'application/json'},
    payload:JSON.stringify({credentials:{
      personalAccessTokenName:CONFIG.TABLEAU_PAT_NAME,
      personalAccessTokenSecret:CONFIG.TABLEAU_PAT_SECRET,
      site:{contentUrl:CONFIG.TABLEAU_SITE}}}),
    muteHttpExceptions:true});
  if(resp.getResponseCode()>=300) throw new Error('Tableau signin fallo ('+resp.getResponseCode()+'): '+resp.getContentText().substring(0,300));
  var cred=JSON.parse(resp.getContentText()).credentials;
  return {token:cred.token, siteId:cred.site.id};
}

function tableauViewId_(auth){
  var filtro=encodeURIComponent('contentUrl:eq:'+CONFIG.TABLEAU_VIEW_CONTENT_URL);
  var resp=UrlFetchApp.fetch(tableauUrl_('/sites/'+auth.siteId+'/views?filter='+filtro),{
    headers:{'X-Tableau-Auth':auth.token, Accept:'application/json'}, muteHttpExceptions:true});
  if(resp.getResponseCode()>=300) throw new Error('Tableau views fallo: '+resp.getContentText().substring(0,300));
  var vs=JSON.parse(resp.getContentText());
  var arr=(vs.views&&vs.views.view)||[];
  if(!arr.length) throw new Error('No se encontro la vista con contentUrl='+CONFIG.TABLEAU_VIEW_CONTENT_URL);
  return arr[0].id;
}

function parseNumLoc_(v){
  if(v==null||v==='') return null;
  var s=String(v).trim(); if(!s) return null;
  s=s.replace(/%/g,'').replace(/\s/g,'');
  if(/^-?\d{1,3}(\.\d{3})*,\d+$/.test(s)) s=s.replace(/\./g,'').replace(',','.');
  else if(/^-?\d+,\d+$/.test(s)) s=s.replace(',','.');
  var n=Number(s); return isNaN(n)?null:n;
}

/** Descarga la crosstab de Tableau y la vuelca en la pestana PERF_SHEET. */
function actualizarPerformance(){
  var auth=tableauSignIn_();
  var viewId=tableauViewId_(auth);
  var resp=UrlFetchApp.fetch(tableauUrl_('/sites/'+auth.siteId+'/views/'+viewId+'/data?maxAge=5'),{
    headers:{'X-Tableau-Auth':auth.token}, muteHttpExceptions:true});
  if(resp.getResponseCode()>=300) throw new Error('Tableau data fallo: '+resp.getContentText().substring(0,300));
  var txt=resp.getContentText().replace(/^\uFEFF/,'');
  var firstLine=txt.split(/\r?\n/)[0]||'';
  var delim=(firstLine.split(';').length>firstLine.split(',').length)?';':',';
  var rows=Utilities.parseCsv(txt, delim);
  if(rows.length<2) throw new Error('La crosstab llego vacia.');
  // mapear encabezados (tolerante a espacios y variantes)
  var H=rows[0].map(function(h){return String(h).trim().toLowerCase();});
  function col(part){ for(var i=0;i<H.length;i++) if(H[i].indexOf(part)>=0) return i; return -1; }
  var iE=col('email'), iN=col('fullname'),
      iFi=col('first reply'), iFu=col('full resolution'), iCs=col('csat'), iNp=col('nps');
  // puede haber varias columnas con 'date' (ej. "Día de Date Time" y "Date Time"):
  // elegir la primera cuyos valores realmente parseen como fecha.
  var iD=-1;
  for(var ci=0;ci<H.length;ci++){
    if(H[ci].indexOf('date')<0 && H[ci].indexOf('fecha')<0) continue;
    for(var rr=1;rr<Math.min(rows.length,6);rr++){
      if(parseDate_(rows[rr][ci])){ iD=ci; break; }
    }
    if(iD>=0) break;
  }
  if(iE<0) throw new Error('La crosstab no trae columna Email. Encabezados: '+rows[0].join(' | '));
  // Unidades: si el encabezado dice (Min) se convierte a horas
  var fFi=(iFi>=0 && H[iFi].indexOf('(min')>=0)?(1/60):1;
  var fFu=(iFu>=0 && H[iFu].indexOf('(min')>=0)?(1/60):1;
  var out=[['Email','Nombre','Fecha','Firt_h','Furt_h','CSAT','NPS']];
  for(var r=1;r<rows.length;r++){
    var em=String(rows[r][iE]||'').trim().toLowerCase(); if(!em||em.indexOf('@')<0) continue;
    var fecha='';                                  // sin columna de fecha = agregado de la vista
    if(iD>=0){ var d=parseDate_(rows[r][iD]); if(!d) continue; fecha=dateKey_(d); }
    var fi=iFi>=0?parseNumLoc_(rows[r][iFi]):null; if(fi!=null) fi=round_(fi*fFi,2);
    var fu=iFu>=0?parseNumLoc_(rows[r][iFu]):null; if(fu!=null) fu=round_(fu*fFu,2);
    var cs=iCs>=0?parseNumLoc_(rows[r][iCs]):null; if(cs!=null&&cs>1.5) cs=cs/100;   // 80 -> 0.80
    var np=iNp>=0?parseNumLoc_(rows[r][iNp]):null;
    out.push([em, iN>=0?rows[r][iN]:'', fecha, fi, fu, cs, np]);
  }
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.PERF_SHEET)||book.insertSheet(CONFIG.PERF_SHEET);
  sh.clearContents();
  sh.getRange(1,1,out.length,7).setValues(out);
  var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Santiago','dd-MM-yyyy HH:mm');
  PropertiesService.getScriptProperties().setProperty(CONFIG.PERF_PROP, stamp);
  return {filas:out.length-1, actualizado:stamp};
}

/** Lee la pestana Performance (cache). */
function loadPerformance_(){
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.PERF_SHEET);
  if(!sh) return [];
  var v=sh.getDataRange().getValues(); if(v.length<2) return [];
  var rows=[];
  for(var i=1;i<v.length;i++){
    if(!v[i][0]) continue;
    var d=(v[i][2]!==''&&v[i][2]!=null)?parseDate_(v[i][2]):null;   // null = agregado sin fecha
    rows.push({email:String(v[i][0]).trim().toLowerCase(), nombre:v[i][1], date:d,
      firt:(v[i][3]===''||v[i][3]==null)?null:Number(v[i][3]),
      furt:(v[i][4]===''||v[i][4]==null)?null:Number(v[i][4]),
      csat:(v[i][5]===''||v[i][5]==null)?null:Number(v[i][5]),
      nps:(v[i][6]===''||v[i][6]==null)?null:Number(v[i][6])});
  }
  return rows;
}

// --------------------- DATOS MAESTROS (planilla BUK) ------------------------
function rolDesdeEspecialidad_(esp){
  var e=norm_(esp);
  if(e.indexOf('EJECUTIV')>=0)    return 'Ejecutivo';
  if(e.indexOf('COORDINADOR')>=0) return 'Coordinador';
  if(e.indexOf('SUPERVISOR')>=0)  return 'Supervisor';
  if(e.indexOf('ANFITRI')>=0)     return 'Anfitrion';
  return 'Otro';
}
function tc_(s){ return String(s||'').toLowerCase().replace(/(^|\s)\S/g,function(c){return c.toUpperCase();}).trim(); }
function supCorto_(full){
  var t=String(full||'').trim().split(/\s+/).filter(Boolean);
  if(t.length>=4) return tc_(t[0]+' '+t[2]);   // 2 nombres + 2 apellidos
  if(t.length===3) return tc_(t[0]+' '+t[1]);
  return tc_(t.join(' '));
}

/** Recibe los registros parseados de la planilla BUK (hoja Trabajador) y
 *  sincroniza la pestana Maestros. regs: [{nombre, ap1, ap2, especialidad, supervisor}] */
function sincronizarMaestros(regs){
  if(!regs||!regs.length) throw new Error('No llegaron registros de la planilla.');
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.MAESTROS_SHEET)||book.insertSheet(CONFIG.MAESTROS_SHEET);
  var prev={}, pv=sh.getDataRange().getValues();
  for(var i=1;i<pv.length;i++) if(pv[i][0]) prev[pv[i][0]]=1;
  var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Santiago','dd-MM-yyyy HH:mm');
  var out=[['NameKey','Nombre_Completo','Rol','Supervisor','Supervisor_Corto','Sincronizado']];
  var nuevos=0, vistos={};
  regs.forEach(function(r){
    var nombre=String(r.nombre||'').trim(), ap1=String(r.ap1||'').trim(), ap2=String(r.ap2||'').trim();
    if(!nombre||!ap1) return;
    var full=(nombre+' '+ap1+(ap2?' '+ap2:'')).replace(/\s+/g,' ');
    var key=nameKey_(nombre.split(/\s+/)[0]+' '+ap1);     // PRIMER nombre + PRIMER apellido = clave Turnos
    if(vistos[key]) return; vistos[key]=1;
    if(!prev[key]) nuevos++;
    out.push([key, tc_(full), rolDesdeEspecialidad_(r.especialidad), tc_(r.supervisor), supCorto_(r.supervisor), stamp]);
  });
  if(out.length<2) throw new Error('La planilla no trae filas validas (Nombre / Primer Apellido).');
  sh.clearContents();
  sh.getRange(1,1,out.length,6).setValues(out);
  PropertiesService.getScriptProperties().setProperty(CONFIG.MAESTROS_PROP, stamp);
  return {total:out.length-1, nuevos:nuevos, actualizado:stamp};
}

/** Lee la pestana Maestros con estructuras para casar correos por nombre completo. */
function loadMaestros_(){
  var book=rosterBook_(), sh=book.getSheetByName(CONFIG.MAESTROS_SHEET);
  if(!sh) return [];
  var v=sh.getDataRange().getValues(), rows=[];
  for(var i=1;i<v.length;i++){
    if(!v[i][0]) continue;
    var toks=norm_(v[i][1]).split(/\s+/).filter(Boolean);
    var tokSet={}, pairSet={};
    toks.forEach(function(t){tokSet[t]=1;});
    for(var a=0;a<toks.length;a++) for(var b=0;b<toks.length;b++) if(a!==b) pairSet[toks[a]+toks[b]]=1;
    rows.push({key:String(v[i][0]), nombre:v[i][1], rol:v[i][2], supervisor:v[i][3], supCorto:v[i][4],
               tokSet:tokSet, pairSet:pairSet});
  }
  return rows;
}

// ------------------------------- METAS -------------------------------------
function getMetas(){
  var raw=PropertiesService.getScriptProperties().getProperty(CONFIG.META_PROP);
  if(!raw) return {grupo:0, metaAgente:0};
  try{ var o=JSON.parse(raw); return {grupo:Number(o.grupo)||0, metaAgente:Number(o.metaAgente)||0}; }
  catch(e){ return {grupo:0, metaAgente:0}; }
}
/** grupo: meta mensual total; metaAgente: meta personal (misma para todos los agentes). */
function saveMetas(grupo, metaAgente){
  var obj={grupo:Number(grupo)||0, metaAgente:Number(metaAgente)||0};
  PropertiesService.getScriptProperties().setProperty(CONFIG.META_PROP, JSON.stringify(obj));
  return obj;
}

// ------------------------------- DASHBOARD ---------------------------------
var MES_FULL_={enero:0,febrero:1,marzo:2,abril:3,mayo:4,junio:5,julio:6,agosto:7,septiembre:8,setiembre:8,octubre:9,noviembre:10,diciembre:11};
function sheetMonthYm_(name){
  var n=String(name).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g,'');
  for(var k in MES_FULL_){ if(n.indexOf(k)>=0) return CONFIG.ANIO+'-'+pad2_(MES_FULL_[k]+1); }
  return null;
}
/** roleKM["nameKey|YYYY-MM"] = rol, tomado de la hoja cuyo TITULO es ese mes.
 *  Prioridad ante conflicto del mismo mes: Supervisor > Coordinador > Ejecutivo > Anfitrion. */
function buildRoles_(){
  var ssT=ss_(CONFIG.TURNOS_SPREADSHEET_ID||CONFIG.VENTAS_SPREADSHEET_ID);
  var roleKM={}, nameByKey={};
  var RANK={Anfitrion:1, Ejecutivo:2, Coordinador:3, Supervisor:4};
  var PREF=[[CONFIG.PREFIJO_AGENTES,'Ejecutivo'],[CONFIG.PREFIJO_ANFITRIONES,'Anfitrion'],
            [CONFIG.PREFIJO_COORDINADORES,'Coordinador'],[CONFIG.PREFIJO_SUPERVISORES,'Supervisor']];
  ssT.getSheets().forEach(function(sh){
    var name=sh.getName(), rol=null;
    for(var i=0;i<PREF.length;i++){ if(norm_(name).indexOf(norm_(PREF[i][0]))===0){ rol=PREF[i][1]; break; } }
    if(!rol) return;
    var ym=sheetMonthYm_(name); if(!ym) return;
    var v=sh.getDataRange().getValues(); if(!v.length) return;
    var dr=-1,best=0; for(var r=0;r<Math.min(8,v.length);r++){ var cnt=0;
      for(var c=1;c<v[r].length;c++) if(parseHeaderDate_(v[r][c])) cnt++; if(cnt>best){best=cnt;dr=r;} }
    if(dr<0||best<3) return;
    for(var r2=dr+1;r2<v.length;r2++){
      var nm=v[r2][0]; if(!nm) continue; var s=String(nm).trim();
      if(!s||s==='.'||s==='Nombre'||s==='Cargo'||s==='Supervisor') continue;
      var has=false; for(var c3=1;c3<v[r2].length && !has;c3++){ if(!parseTurno_(v[r2][c3]).libre) has=true; }
      if(!has) continue;
      var key=nameKey_(s); nameByKey[key]=s;
      var cur=roleKM[key+'|'+ym];
      if(!cur || RANK[rol]>RANK[cur]) roleKM[key+'|'+ym]=rol;
    }
  });
  return {roleKM:roleKM, nameByKey:nameByKey};
}

function getDashboard(params){
  params=params||{};
  var from=params.from?parseDate_(params.from):firstOfMonth_();
  var to=params.to?endOfDay_(parseDate_(params.to)):endOfDay_(new Date());
  var prodFilter=params.product||'todas', teamFilter=params.team||'todos';

  var R=loadRoster_(), roster=R.byEmail, keyByEmail=R.keyByEmail;
  var ventas=loadVentas_();
  var agentPeople=loadTurnosPrefijo_(CONFIG.PREFIJO_AGENTES);
  var coordPeople=loadTurnosPrefijo_(CONFIG.PREFIJO_COORDINADORES);

  // ROL POR MES (temporal): el rol de cada mes se toma de la hoja cuyo TITULO es
  // ese mes (ej. el rol de mayo viene de "Agentes/Anfitriones/... Mayo"), no del
  // derrame de fechas de otra hoja. Asi un cambio de rol entre meses queda correcto.
  var RB=buildRoles_(), roleKM=RB.roleKM, nameByKey=RB.nameByKey;

  // Indices para casar correos SIN depender del Roster:
  //  - concatIdx: "CARLOSPALACIOS" -> "CARLOS PALACIOS" (nombre+apellido pegados en el correo)
  //  - firstIdx:  nombre de pila -> claves de Turnos (se usa solo si es UNICO, ej. "maribel")
  var concatIdx={}, firstIdx={};
  Object.keys(nameByKey).forEach(function(k){
    var p=k.split(' ');
    (firstIdx[p[0]]=firstIdx[p[0]]||[]).push(k);
    if(p.length>=2) concatIdx[p[0]+p[p.length-1]]=k;
  });

  // franjaCoords["date|hour"] = [coordName]  (horas activas, con wrap)
  var franjaCoords={};
  Object.keys(coordPeople).forEach(function(key){ var p=coordPeople[key];
    Object.keys(p.shifts).forEach(function(dk){ var t=p.shifts[dk]; var base=parseDate_(dk); var admin=adminHoras_(t.start);
      shiftHours_(t.start,t.end).forEach(function(x){ if(admin.indexOf(x.h)>=0) return;
        var d=x.off?dateKey_(addDays_(base,1)):dk; (franjaCoords[d+'|'+x.h]=franjaCoords[d+'|'+x.h]||[]).push(p.name); }); }); });

  // Equipo del EJECUTIVO por su TURNO. Dos mapas:
  //  - agentDayTeam[nk|fechaInicioTurno] = equipo del turno que EMPIEZA ese dia
  //    (cubre TODAS sus ventas de ese dia, sin importar la hora exacta).
  //  - agentCover[nk|fecha|hora] = equipo que cubre esa hora (para la madrugada
  //    que pertenece a un turno nocturno iniciado el dia anterior).
  // cruza medianoche -> Noche (todo el turno); mismo dia -> Dia.
  var agentDayTeam={}, agentCover={};
  Object.keys(agentPeople).forEach(function(key){ var p=agentPeople[key];
    Object.keys(p.shifts).forEach(function(dk){ var t=p.shifts[dk]; var base=parseDate_(dk); var team=t.cross?'Noche':'Dia';
      agentDayTeam[key+'|'+dk]=team;
      shiftHours_(t.start,t.end).forEach(function(x){ var d=x.off?dateKey_(addDays_(base,1)):dk; agentCover[key+'|'+d+'|'+x.h]=team; }); }); });

  var COMP=CONFIG.PRODUCTO_COMPARTIDA, EXCL=CONFIG.PRODUCTO_EXCLUSIVA;
  var horas=[]; for(var h=0;h<24;h++) horas.push({hora:h,comp:0,excl:0,total:0,n:0});
  var dowSum=[0,0,0,0,0,0,0], dowN=[0,0,0,0,0,0,0];  // 0=Dom..6=Sab
  var kpis={totalVentas:0,nViajes:0,compMonto:0,exclMonto:0,compN:0,exclN:0,diaMonto:0,nocheMonto:0,sinMonto:0};
  var ejec={}, coordCom={}, otras={}, malPorEjec={}, malCorreos=[];

  // DATOS MAESTROS (planilla BUK): permite casar CUALQUIER correo contra el
  // nombre completo (incluye 2dos nombres y 2dos apellidos) y da rol/supervisor de respaldo.
  var maestros=loadMaestros_(), mmCache={};
  function matchMaestros_(email){
    if(email in mmCache) return mmCache[email];
    var toks=norm_(String(email).split('@')[0]).split(/[._]/).filter(Boolean);
    var hit=null;
    if(toks.length){
      var hits=maestros.filter(function(m){
        return toks.every(function(t){ return m.tokSet[t]||m.pairSet[t]; });
      });
      if(hits.length===1) hit=hits[0];
    }
    mmCache[email]=hit; return hit;
  }

  var venCache={};
  function infoVendedor(email, ym){
    var ck=email+'|'+ym; if(venCache[ck]) return venCache[ck];
    var rinfo=roster[email];
    var mm=matchMaestros_(email);
    var t0=norm_(String(email).split('@')[0].replace(/_/g,'.').split('.').filter(Boolean)[0]||'');
    var cands=[];
    if(CONFIG.ALIAS_CORREOS[email]) cands.push(nameKey_(CONFIG.ALIAS_CORREOS[email])); // alias manual
    if(keyByEmail[email]) cands.push(keyByEmail[email]);                                // Roster (si existe)
    if(mm) cands.push(mm.key);                                                          // Maestros (planilla BUK)
    cands.push(keyFromEmail_(email));                                                   // nombre.apellido del correo
    if(concatIdx[t0]) cands.push(concatIdx[t0]);                                        // nombreapellido pegados
    if(firstIdx[t0] && firstIdx[t0].length===1) cands.push(firstIdx[t0][0]);            // nombre de pila unico
    var seen={}; cands=cands.filter(function(k){ if(!k||seen[k])return false; seen[k]=1; return true; });
    // nk para el equipo: el primer candidato que EXISTA en Turnos
    var nk=null; for(var i=0;i<cands.length;i++){ if(nameByKey[cands[i]]){ nk=cands[i]; break; } }
    if(!nk) nk=cands[0];
    // rol del periodo: Turnos del mes (cualquier candidato) -> Maestros -> Roster -> 'Otro'
    var rol=''; for(var j=0;j<cands.length;j++){ if(roleKM[cands[j]+'|'+ym]){ rol=roleKM[cands[j]+'|'+ym]; break; } }
    if(!rol) rol=(mm&&mm.rol)||(rinfo&&rinfo.rol)||'Otro';
    var nombre=(mm&&mm.nombre)||(rinfo&&rinfo.nombre)||nameByKey[nk]||email;
    var sup=(mm&&mm.supCorto)||(rinfo&&rinfo.supervisor)||'';
    var res={nk:nk, rol:rol, nombre:nombre, sup:sup};
    venCache[ck]=res; return res;
  }

  ventas.forEach(function(s){
    if(s.date<from||s.date>to) return;
    if(CONFIG.SOLO_CONCRETADAS && s.fin!==CONFIG.FINISH_OK) return;
    var isComp=(s.product===COMP), isExcl=(s.product===EXCL);
    if(prodFilter==='compartida'&&!isComp) return;
    if(prodFilter==='exclusiva'&&!isExcl) return;

    var ven=infoVendedor(s.agent, s.dateKey.substring(0,7));
    var esEjec=(ven.rol==='Ejecutivo');

    // equipo Diurno/Nocturno SOLO aplica a ejecutivos (por su turno del dia).
    // Los no-ejecutivos no tienen equipo: no suman a Dia/Noche/Sin turno ni al grafico de franja.
    var team=null;
    if(esEjec){
      team=agentDayTeam[ven.nk+'|'+s.dateKey] || agentCover[ven.nk+'|'+s.dateKey+'|'+s.hour] || 'Sin turno';
    }
    if(teamFilter!=='todos' && team!==teamFilter) return;  // el filtro de equipo deja solo ejecutivos

    // Tarjetas / totales: TODAS las ventas (ejecutivos + otras)
    kpis.totalVentas+=s.amount; kpis.nViajes++;
    if(isComp){kpis.compMonto+=s.amount;kpis.compN++;}
    if(isExcl){kpis.exclMonto+=s.amount;kpis.exclN++;}
    // Franja horaria (TODAS las ventas) por compartida/exclusiva
    var hb=horas[s.hour]; hb.total+=s.amount; hb.n++;
    if(isComp)hb.comp+=s.amount; if(isExcl)hb.excl+=s.amount;
    // Dia de la semana (TODAS las ventas)
    var wd=s.date.getDay(); dowSum[wd]+=s.amount; dowN[wd]++;
    // Dia/Noche: SOLO ejecutivos (cards de equipo y tabla de ejecutivos)
    if(esEjec){
      if(team==='Noche')kpis.nocheMonto+=s.amount; else if(team==='Dia')kpis.diaMonto+=s.amount; else kpis.sinMonto+=s.amount;
    }

    // EJECUTIVOS (vendedor ejecutivo) u OTRAS VENTAS (vendedor no ejecutivo)
    if(esEjec){
      var e=ejec[ven.nk]||(ejec[ven.nk]={nombre:ven.nombre,sup:ven.sup,monto:0,n:0,comp:0,excl:0,dia:0,noche:0});
      e.monto+=s.amount; e.n++; if(isComp)e.comp+=s.amount; if(isExcl)e.excl+=s.amount;
      if(team==='Noche')e.noche+=s.amount; else if(team==='Dia')e.dia+=s.amount;
    } else {
      var o=otras[s.agent]||(otras[s.agent]={nombre:ven.nombre,rol:ven.rol,monto:0,n:0});
      o.monto+=s.amount; o.n++;
    }

    // COMISION COORDINADORES por franja activa (50/50 si hay 2)
    var coords=franjaCoords[s.dateKey+'|'+s.hour]||[];
    if(coords.length){ var shr=s.amount/coords.length;
      coords.forEach(function(c){ var cc=coordCom[c]||(coordCom[c]={nombre:c,dia:0,noche:0,total:0,n:0});
        cc.total+=shr; cc.n+=1/coords.length; if(team==='Noche')cc.noche+=shr; else if(team==='Dia')cc.dia+=shr; }); }

    // CORREOS MAL INGRESADOS (ds_user_email, solo compartida, vendedor ejecutivo)
    if(isComp&&esEjec){
      var mp=malPorEjec[ven.nk]||(malPorEjec[ven.nk]={nombre:ven.nombre,total:0,malos:0});
      mp.total++;
      if(!emailOk_(s.user)){ mp.malos++;
        if(malCorreos.length<3000) malCorreos.push({ejecutivo:ven.nombre,correo:s.user,fecha:s.dateKey,monto:s.amount}); }
    }
  });

  var metas=getMetas();
  var metaAg=metas.metaAgente||0;

  var ejecList=Object.keys(ejec).map(function(k){ var e=ejec[k];
    return {nombre:e.nombre, sup:e.sup||'', ventas:round_(e.monto), n:e.n, comp:round_(e.comp), excl:round_(e.excl),
            dia:round_(e.dia), noche:round_(e.noche), meta:metaAg,
            avance: metaAg? round_(e.monto/metaAg,3):null}; })
    .sort(function(a,b){return b.ventas-a.ventas;});

  var coordList=Object.keys(coordCom).map(function(c){ var x=coordCom[c];
    return {coordinador:c, dia:round_(x.dia), noche:round_(x.noche), total:round_(x.total), n:round_(x.n,1)}; })
    .sort(function(a,b){return b.total-a.total;});

  var otrasList=Object.keys(otras).map(function(e){ return {email:e,nombre:otras[e].nombre,rol:otras[e].rol,
    monto:round_(otras[e].monto),n:otras[e].n}; }).sort(function(a,b){return b.monto-a.monto;});

  var malList=Object.keys(malPorEjec).map(function(k){ var m=malPorEjec[k]; return {nombre:m.nombre,
    total:m.total,malos:m.malos,pctOk:m.total?round_(1-(m.malos/m.total),3):null}; })
    .sort(function(a,b){return b.malos-a.malos;});

  // DESEMPENO (Tableau cache) agregado por agente dentro del rango
  var perf=loadPerformance_(), perfAgg={}, perfSinFechas=false;
  perf.forEach(function(p){
    if(p.date){ if(p.date<from||p.date>to) return; } else perfSinFechas=true;
    var a=perfAgg[p.email]||(perfAgg[p.email]={n:0,fiOk:0,fin:0,fuOk:0,fun:0,cs:0,csn:0,np:0,npn:0});
    a.n++;
    if(p.firt!=null){a.fin++; if(p.firt<=CONFIG.FIRT_UMBRAL_H)a.fiOk++;}
    if(p.furt!=null){a.fun++; if(p.furt<=CONFIG.FURT_UMBRAL_H)a.fuOk++;}
    if(p.csat!=null){a.cs+=p.csat;a.csn++;}
    if(p.nps!=null){a.np+=p.nps;a.npn++;}
  });
  var ymRef=dateKey_(to).substring(0,7);
  var perfList=Object.keys(perfAgg).map(function(em){ var a=perfAgg[em]; var ven=infoVendedor(em, ymRef);
    return {nombre:ven.nombre, rol:ven.rol, tickets:a.n,
      csat:a.csn?round_(a.cs/a.csn,3):null, nps:a.npn?round_(a.np/a.npn):null,
      firt:a.fin?round_(a.fiOk/a.fin,3):null,    // % cumplimiento primera respuesta <= 24h
      furt:a.fun?round_(a.fuOk/a.fun,3):null};   // % cumplimiento resolucion <= 120h
    })
    .sort(function(a,b){return (b.csat==null?-1:b.csat)-(a.csat==null?-1:a.csat);});
  var perfStamp=PropertiesService.getScriptProperties().getProperty(CONFIG.PERF_PROP)||'';

  // Promedio por dia de semana: suma del dia / cantidad de ese dia en el rango (hasta hoy)
  var diaCount=[0,0,0,0,0,0,0], hoy=new Date(), fin=(to<hoy?to:hoy);
  for(var dcur=new Date(from.getFullYear(),from.getMonth(),from.getDate()); dcur<=fin; dcur=addDays_(dcur,1)) diaCount[dcur.getDay()]++;
  var NOMBRE_DOW=['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  var ORDEN=[1,2,3,4,5,6,0]; // Lunes..Domingo
  var semana=ORDEN.map(function(wd){ return {dia:NOMBRE_DOW[wd], total:round_(dowSum[wd]),
    dias:diaCount[wd], n:dowN[wd], promedio: diaCount[wd]? round_(dowSum[wd]/diaCount[wd]):0}; });

  return {
    rango:{from:dateKey_(from),to:dateKey_(to)},
    kpis:{totalVentas:round_(kpis.totalVentas),nViajes:kpis.nViajes,compMonto:round_(kpis.compMonto),
      exclMonto:round_(kpis.exclMonto),compN:kpis.compN,exclN:kpis.exclN,
      diaMonto:round_(kpis.diaMonto),nocheMonto:round_(kpis.nocheMonto),sinMonto:round_(kpis.sinMonto)},
    franjas:horas.map(function(x){return {hora:x.hora,comp:round_(x.comp),excl:round_(x.excl),total:round_(x.total),n:x.n};}),
    semana:semana,
    desempeno:{lista:perfList, actualizado:perfStamp, sinFechas:perfSinFechas},
    maestros:{total:maestros.length, actualizado:PropertiesService.getScriptProperties().getProperty(CONFIG.MAESTROS_PROP)||''},
    ejecutivos:ejecList, coordinadores:coordList, otrasVentas:otrasList,
    correosMal:{detalle:malCorreos, porEjecutivo:malList},
    metas:{grupo:metas.grupo||0, metaAgente:metaAg}
  };
}

// ------------------------------- export CSV --------------------------------
function exportCSV(tipo,params){
  var d=getDashboard(params), rows=[];
  if(tipo==='ejecutivos'){ rows.push(['Ejecutivo','Supervisor','Ventas','Viajes','Compartida','Exclusiva','Dia','Noche','Meta','%_Avance']);
    d.ejecutivos.forEach(function(e){rows.push([e.nombre,e.sup,e.ventas,e.n,e.comp,e.excl,e.dia,e.noche,e.meta,e.avance]);}); }
  else if(tipo==='coordinador'){ rows.push(['Coordinador','Ventas_Dia','Ventas_Noche','Ventas_Total','Viajes']);
    d.coordinadores.forEach(function(c){rows.push([c.coordinador,c.dia,c.noche,c.total,c.n]);}); }
  else if(tipo==='grupo'){ rows.push(['Grupo','Ventas','% del total']);
    var tot=(d.kpis.diaMonto+d.kpis.nocheMonto+d.kpis.sinMonto)||1;
    rows.push(['Equipo Diurno',d.kpis.diaMonto,round_(d.kpis.diaMonto/tot,3)]);
    rows.push(['Equipo Nocturno',d.kpis.nocheMonto,round_(d.kpis.nocheMonto/tot,3)]);
    rows.push(['Sin turno',d.kpis.sinMonto,round_(d.kpis.sinMonto/tot,3)]);
    rows.push([]); rows.push(['Compartida',d.kpis.compMonto,d.kpis.compN]); rows.push(['Exclusiva',d.kpis.exclMonto,d.kpis.exclN]);
    rows.push([]); rows.push(['Meta grupal',d.metas.grupo,d.metas.grupo?round_(d.kpis.totalVentas/d.metas.grupo,3):'']); }
  else if(tipo==='franja'){ rows.push(['Hora','Compartida','Exclusiva','Total','Viajes']);
    d.franjas.forEach(function(f){rows.push([f.hora,f.comp,f.excl,f.total,f.n]);}); }
  else if(tipo==='semana'){ rows.push(['Dia','Ventas_promedio','Total','Dias_en_periodo','Viajes']);
    d.semana.forEach(function(s){rows.push([s.dia,s.promedio,s.total,s.dias,s.n]);}); }
  else if(tipo==='desempeno'){ rows.push(['Agente','Rol','Tickets','%_CSAT','NPS','%_FIRT_24h','%_FURT_120h']);
    d.desempeno.lista.forEach(function(p){rows.push([p.nombre,p.rol,p.tickets,p.csat,p.nps,p.firt,p.furt]);}); }
  else if(tipo==='otras'){ rows.push(['Email','Nombre','Rol','Ventas','Viajes']);
    d.otrasVentas.forEach(function(o){rows.push([o.email,o.nombre,o.rol,o.monto,o.n]);}); }
  else if(tipo==='correos'){ rows.push(['Ejecutivo','Correo_pasajero','Fecha','Monto']);
    d.correosMal.detalle.forEach(function(m){rows.push([m.ejecutivo,m.correo,m.fecha,m.monto]);}); }
  return rows.map(function(r){return r.map(csvCell_).join(',');}).join('\n');
}
function csvCell_(v){ v=(v==null?'':String(v)); return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; }

// --------- DIAGNOSTICO (ejecutar desde el editor de Apps Script) ------------
// Ejecuta diagnostico() y revisa Registros (Ver > Registros). Te dice cuantas
// filas leyo del Roster y como queda clasificado un correo de ejemplo.
function diagnostico(){
  var book=rosterBook_();
  Logger.log('Libro del Roster: "'+book.getName()+'"  (hojas: '+book.getSheets().map(function(s){return s.getName();}).join(', ')+')');
  var R=loadRoster_();
  var emails=Object.keys(R.byEmail);
  Logger.log('Roster: '+emails.length+' correos leidos.');
  Logger.log('¿maribel.estefani@cabify.com en Roster? '+(R.byEmail['maribel.estefani@cabify.com']?('SI -> '+R.byEmail['maribel.estefani@cabify.com'].rol):'NO'));
  Logger.log('¿carlospalacios.alvarez@cabify.com en Roster? '+(R.byEmail['carlospalacios.alvarez@cabify.com']?('SI -> '+R.byEmail['carlospalacios.alvarez@cabify.com'].rol):'NO'));
  var ag=loadTurnosPrefijo_(CONFIG.PREFIJO_AGENTES);
  Logger.log('Agentes en Turnos 360: '+Object.keys(ag).length+' -> '+Object.keys(ag).slice(0,30).join(', '));
  return {rosterCount:emails.length, agentes:Object.keys(ag)};
}
