let sapRows=[],sap=[],planRows=[],plan=[],planesCargados={},charts={}, sapArchivo="Pendiente", planArchivo="Pendiente", ultimaCarga="Pendiente", githubArchivos=[];
const ESTADOS_API_URL='https://script.google.com/macros/s/AKfycbwryr1x44hIG9wUeU_8kc0ZmOjBwXT7TXjeAhluQh46EdidOjg1pDLJIkCl61VJVovt/exec';
let estadosTerreno={},planActualPorAviso={},avisoEstadoAbierto='';
let encargadoPlanFiltrado='';
const GITHUB_OWNER="hardycofre-commits", GITHUB_REPO="dashboard-hh-mantencion-sap", GITHUB_BRANCH="main", GITHUB_DATA_API=`https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/datos?ref=${GITHUB_BRANCH}`, GITHUB_COMMITS_API=`https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/commits`;
const FALLBACK_DATA_FILES=[
  'EXPORT - 2026-07-03T090738.938.xlsx','EXPORT - 2026-07-03T102715.587.xlsx',
  'EXPORT - 2026-07-06T094609.530.xlsx','EXPORT - 2026-07-06T144233.767.xlsx','EXPORT - 2026-07-06T145430.956.xlsx',
  'EXPORT - 2026-07-07T095437.397.xlsx','EXPORT - 2026-07-07T163348.545.xlsx','EXPORT - 2026-07-08T170719.540.xlsx',
  'EXPORT - 2026-07-10T083450.471.XLSX','EXPORT - 2026-07-13T114745.863.xlsx','EXPORT - 2026-07-14T090216.157.xlsx',
  'EXPORT - 2026-07-15T170932.723.xlsx','EXPORT - 2026-07-20T070443.584.xlsx','EXPORT - 2026-07-21T133102.836.xlsx',
  'EXPORT - 2026-07-22T083156.325.xlsx','EXPORT - 2026-07-23T144449.081.XLSX','EXPORT - 2026-07-24T114636.912.xlsx',
  'EXPORT - 2026-07-27T104903.863.xlsx','EXPORT - 2026-07-29T171504.804.xlsx','EXPORT - 2026-07-30T160251.425.xlsx',
  'EXPORT - 2026-08-03T112534.786.xlsx','sap.xlsx','plan_semanal.xlsx','semana27.xlsx','semana28.xlsx',
  'Semana29.xlsx','Semana30.xlsx','Semana31.xlsx','Semana32.xlsx'
];
const COLORS={blue:'#0b3a78',sky:'#38a3e8',green:'#16a34a',red:'#dc2626',orange:'#f59e0b',pink:'#f45b85',purple:'#6d45c9',gray:'#64748b'};
const CLASS_INFO={ZM01:'Correctiva',ZM02:'Mantención preventiva',ZM05:'Proyecto'};
const META_HH_MENSUAL=350;
const $=id=>document.getElementById(id); const fmt=n=>Number(n||0).toLocaleString('es-CL',{minimumFractionDigits:1,maximumFractionDigits:1});
function norm(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim()}
function findCol(row,keys){const cols=Object.keys(row||{});return cols.find(c=>keys.some(k=>norm(c).includes(k)))||null}
function toNum(v){
  if(typeof v==='number') return isFinite(v)?v:0;
  let raw=String(v??'').trim();
  raw=raw.replace(/[\u200E\u200F\u202A-\u202E\u2066-\u2069\uFEFF]/g,'').replace(/\u00A0/g,' ').trim();
  raw=raw.replace(/[−–—]/g,'-');
  let neg=/-\s*$/.test(raw)||/^\s*-/.test(raw)||/^\(.*\)$/.test(raw);
  let s=raw.replace(/^\s*-/, '').replace(/-\s*$/, '').replace(/[()]/g,'').trim();
  s=s.replace(/[^0-9.,]/g,'');
  if(!s) return 0;
  if(s.includes(',')){
    s=s.replace(/\./g,'').replace(',', '.');
  }else{
    const parts=s.split('.');
    if(parts.length>2){
      const dec=parts.pop();
      s=parts.join('')+'.'+dec;
    }
  }
  let n=parseFloat(s)||0;
  return neg?-Math.abs(n):n;
}
function excelDate(v){if(v instanceof Date)return v.toISOString().slice(0,10); if(typeof v==='number'&&v>30000){let d=XLSX.SSF.parse_date_code(v);return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`} let s=String(v||'').trim(); let m=s.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/); if(m)return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; return s}
function showDate(s){if(!s)return''; let m=String(s).match(/(\d{4})-(\d{2})-(\d{2})/); return m?`${m[3]}-${m[2]}-${m[1]}`:s}
function showMes(s){if(!s)return''; let m=String(s).match(/(\d{4})-(\d{2})/); if(!m)return s; const names=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']; return names[Number(m[2])-1]+' '+m[1]}
function periodoMesesLabel(desde,hasta){let a=(desde||'').slice(0,7), b=(hasta||'').slice(0,7); if(!a&&!b)return 'Sin período'; if(a===b)return showMes(a); return showMes(a)+' a '+showMes(b)}
function llenarSelectorAnios(){
  const anios=unique(sap.map(x=>(x.fecha||'').slice(0,4)).filter(x=>/^\d{4}$/.test(x))).sort().reverse();
  const sel=$('anioPeriodo'),actual=sel.value;
  sel.innerHTML='';
  (anios.length?anios:[String(new Date().getFullYear())]).forEach(anio=>{const o=document.createElement('option');o.value=anio;o.textContent=anio;sel.appendChild(o)});
  if(anios.includes(actual))sel.value=actual;
}
function cambiarTipoPeriodo(){
  const tipo=$('tipoPeriodo').value;
  $('controlSemana').classList.toggle('hidden',tipo!=='semanal');
  $('controlMes').classList.toggle('hidden',tipo!=='mensual');
  $('controlAnio').classList.toggle('hidden',tipo!=='anual');
  document.querySelectorAll('.periodSwitch [data-periodo]').forEach(btn=>btn.classList.toggle('active',btn.dataset.periodo===tipo));
  aplicarPeriodoSeleccionado();
}
function seleccionarTipoPeriodo(tipo){
  $('tipoPeriodo').value=tipo;
  cambiarTipoPeriodo();
}
function rangoSemanaIso(valor){
  const m=String(valor||'').match(/^(\d{4})-W(\d{2})$/);
  if(!m)return null;
  const anio=Number(m[1]),semana=Number(m[2]);
  const cuatroEnero=new Date(Date.UTC(anio,0,4));
  const lunesSemanaUno=new Date(cuatroEnero);
  lunesSemanaUno.setUTCDate(cuatroEnero.getUTCDate()-(cuatroEnero.getUTCDay()||7)+1);
  const lunes=new Date(lunesSemanaUno);
  lunes.setUTCDate(lunesSemanaUno.getUTCDate()+(semana-1)*7);
  const domingo=new Date(lunes);
  domingo.setUTCDate(lunes.getUTCDate()+6);
  const iso=d=>d.toISOString().slice(0,10);
  return {desde:iso(lunes),hasta:iso(domingo)};
}
function semanaIsoDesdeFecha(fecha){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(String(fecha||'')))return'';
  const d=new Date(fecha+'T00:00:00Z');
  d.setUTCDate(d.getUTCDate()+4-(d.getUTCDay()||7));
  const inicio=new Date(Date.UTC(d.getUTCFullYear(),0,1));
  const semana=Math.ceil((((d-inicio)/86400000)+1)/7);
  return `${d.getUTCFullYear()}-W${String(semana).padStart(2,'0')}`;
}
function aplicarPeriodoSeleccionado(actualizar=true){
  if($('tipoPeriodo').value==='semanal'){
    const rango=rangoSemanaIso($('semanaPeriodo').value);
    if(!rango)return;
    $('desde').value=rango.desde;
    $('hasta').value=rango.hasta;
  }else if($('tipoPeriodo').value==='anual'){
    const anio=$('anioPeriodo').value;
    if(!/^\d{4}$/.test(anio))return;
    $('desde').value=anio+'-01-01';
    $('hasta').value=anio+'-12-31';
  }else{
    const mes=$('mesPeriodo').value;
    if(!/^\d{4}-\d{2}$/.test(mes))return;
    const [anio,numeroMes]=mes.split('-').map(Number);
    $('desde').value=mes+'-01';
    $('hasta').value=`${mes}-${String(new Date(anio,numeroMes,0).getDate()).padStart(2,'0')}`;
  }
  const selectorPlan=$('planSelector');
  if(selectorPlan && planesCargados && Object.keys(planesCargados).length){
    selectorPlan.value='__periodo__';
    plan=Object.values(planesCargados).flat();
    planArchivo='Planes consolidados del período';
  }
  if(actualizar)render();
}
function extractOT(v){let m=String(v||'').match(/\b(?:1|2|5)\d{7}\b/g);return m?m[0]:''}
function cellValueHybrid(cell){
  if(!cell) return '';
  const w=String(cell.w??'').trim();
  if(/-\s*$/.test(w)) return w;
  return cell.v ?? w ?? '';
}
function sheetAOAHybrid(ws){
  const range=XLSX.utils.decode_range(ws['!ref']||'A1:A1');
  const aoa=[];
  for(let r=range.s.r;r<=range.e.r;r++){
    const row=[];
    for(let c=range.s.c;c<=range.e.c;c++){
      row.push(cellValueHybrid(ws[XLSX.utils.encode_cell({r,c})]));
    }
    aoa.push(row);
  }
  return aoa;
}
function parseWB(buffer){
  let wb=XLSX.read(buffer,{type:'array',cellText:true,cellNF:true,cellDates:false});
  let ws=wb.Sheets[wb.SheetNames[0]];
  let aoa=sheetAOAHybrid(ws);
  return detectTable(aoa);
}
async function readWB(file){return new Promise((res,rej)=>{const r=new FileReader();r.onload=e=>{try{res(parseWB(e.target.result))}catch(err){rej(err)}};r.onerror=rej;r.readAsArrayBuffer(file)})}
function arrayBufferToBase64(buffer){let bytes=new Uint8Array(buffer),binary='';for(let i=0;i<bytes.byteLength;i++)binary+=String.fromCharCode(bytes[i]);return btoa(binary)}
function base64ToArrayBuffer(base64){let binary=atob(base64),bytes=new Uint8Array(binary.length);for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);return bytes.buffer}
async function guardarArchivoLocal(tipo,file){return new Promise((res,rej)=>{const r=new FileReader();r.onload=e=>{try{localStorage.setItem('hh_'+tipo+'_name',file.name);localStorage.setItem('hh_'+tipo+'_data',arrayBufferToBase64(e.target.result));res()}catch(err){rej(err)}};r.onerror=rej;r.readAsArrayBuffer(file)})}
function cargarGuardado(tipo){let data=localStorage.getItem('hh_'+tipo+'_data');if(!data)return null;return {name:localStorage.getItem('hh_'+tipo+'_name')||'archivo guardado',rows:parseWB(base64ToArrayBuffer(data))}}
function limpiarArchivos(){localStorage.removeItem('hh_sap_name');localStorage.removeItem('hh_sap_data');localStorage.removeItem('hh_plan_name');localStorage.removeItem('hh_plan_data');sap=[];plan=[];sapRows=[];planRows=[];$('sapFile').value='';$('planFile').value='';render()}
function detectTable(aoa){
  let best=0,bestScore=-1;
  for(let i=0;i<Math.min(30,aoa.length);i++){
    let text=aoa[i].map(norm).join(' | ');
    let score=0;
    ['orden','aviso','fecha','trabajo','hh','hora','clase','operacion','notificacion','personal'].forEach(k=>{if(text.includes(k))score++});
    if(score>bestScore){bestScore=score;best=i}
  }
  let headers=(aoa[best]||[]).map((h,i)=>String(h||('Columna '+(i+1))).trim()||('Columna '+(i+1)));
  let seen={};headers=headers.map((h,i)=>{let key=h||('Columna '+(i+1)); if(seen[key]){seen[key]++; key=key+'_'+seen[key]}else seen[key]=1; return key});
  return aoa.slice(best+1).filter(r=>r && r.some(v=>String(v||'').trim()!=='')).map(r=>{let o={};headers.forEach((h,i)=>o[h]=r[i]??'');return o});
}
function findColExact(row,keys){const cols=Object.keys(row||{});return cols.find(c=>keys.some(k=>norm(c)===norm(k)))||cols.find(c=>keys.some(k=>norm(c).includes(norm(k))))||null}
function findHHCol(row){
  const cols=Object.keys(row||{});
  const bad=/h\.in\.real|horfinreal|hora\s*inicio|hora\s*fin|inicio|fin/i;
  const preferred=['Trabajo real','HH','Horas reales','Duracion real','Duración real','Actual work'];
  let c=cols.find(col=>!bad.test(norm(col)) && preferred.some(k=>norm(col)===norm(k))) ||
        cols.find(col=>!bad.test(norm(col)) && preferred.some(k=>norm(col).includes(norm(k))));
  if(c) return c;
  return cols.find(col=>!bad.test(norm(col)) && /trabajo\s*real|hh|horas\s*reales|actual\s*work/i.test(norm(col)))||null;
}
function mapSAP(rows){return rows.map(r=>{
  let joined=Object.values(r).join(' ');
  let cOT=findColExact(r,['Numero de orden','Número de orden','Orden','Orden de trabajo','OT','Order']);
  let cHH=findHHCol(r);
  let cClase=findColExact(r,['Clase de orden','Clase orden','Clase','Tipo de orden','ClOrd']);
  let cFecha=findColExact(r,['Fecha','Fecha contabilización','Fecha contab.','Fecha de contabilización','Fecha notificación','Fecha inicio','Día','FecFinReal']);
  let cTxt=findColExact(r,['Texto breve','Texto','Descripción','Descripcion','Trabajo','Operación','Operacion']);
  let cResp=findColExact(r,['Técnico','Tecnico','Responsable','Nombre','Personal','Ejecutor','Empleado','Creado por']);
  let cEstado=findColExact(r,['Estado','Status','Notificación','Notificacion','Estado sistema']);
  let ot=extractOT(cOT?r[cOT]:joined);
  let hh=toNum(cHH?r[cHH]:0);
  let est=String(cEstado?r[cEstado]:'');
  let notificada=!!ot || hh!==0 || /noti|notif|cnf|conf|cerr|final|iw41|real/i.test(est);
  return{ot,hh,clase:cClase?String(r[cClase]||'Sin clase').trim():'Sin clase',fecha:excelDate(cFecha?r[cFecha]:''),trabajo:cTxt?r[cTxt]:'',resp:cResp?r[cResp]:'Sin responsable',estado:notificada?'Notificada':'Sin notificar'}
}).filter(x=>x.ot)}
function mapPlan(rows){return rows.map(r=>{let cF=findColExact(r,['Fecha']);let cA=findColExact(r,['Número de aviso','Numero de aviso','Aviso']);let cO=findColExact(r,['Número de orden','Numero de orden','Orden','OT']);let cT=findColExact(r,['Trabajo','Descripción','Descripcion']);let cE=findColExact(r,['Encargado','Responsable']);let cTu=findColExact(r,['Turno']);let ot=extractOT(cO?r[cO]:Object.values(r).join(' '));return{fecha:excelDate(cF?r[cF]:''),aviso:cA?r[cA]:'',ot,trabajo:cT?r[cT]:'',encargado:cE?r[cE]:'',turno:cTu?r[cTu]:'',estado:'Pendiente'}}).filter(x=>x.ot)}
function destroy(id){if(charts[id])charts[id].destroy()}
function chart(id,type,labels,data,label,colors){
  destroy(id);
  const isPie=type==='pie'||type==='doughnut';
  const total=data.reduce((a,b)=>a+Number(b||0),0);
  charts[id]=new Chart($(id),{
    type,
    data:{labels,datasets:[{label,data,borderWidth:1,backgroundColor:colors||[COLORS.sky,COLORS.pink,COLORS.orange,COLORS.green,COLORS.purple,COLORS.gray],borderColor:'#ffffff'}]},
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:isPie,position:'top',labels:{usePointStyle:true,boxWidth:10}},
        tooltip:{callbacks:{label:(ctx)=>{let v=Number(ctx.raw||0);let pct=total?Math.round(v/total*100):0;return ` ${ctx.label}: ${v.toLocaleString('es-CL')} ${label||''} (${pct}%)`;}}}
      },
      scales:isPie?{}:{y:{beginAtZero:true}}
    },
    plugins:isPie?[piePercentPlugin]:[]
  });
}
const piePercentPlugin={id:'piePercentPlugin',afterDatasetsDraw(chart){const {ctx}=chart;const ds=chart.data.datasets[0];const total=ds.data.reduce((a,b)=>a+Number(b||0),0);if(!total)return;ctx.save();ctx.font='bold 13px Arial';ctx.fillStyle='#ffffff';ctx.textAlign='center';ctx.textBaseline='middle';chart.getDatasetMeta(0).data.forEach((arc,i)=>{const val=Number(ds.data[i]||0);const pct=Math.round(val/total*100);if(pct<5)return;const p=arc.tooltipPosition();ctx.fillText(pct+'%',p.x,p.y);});ctx.restore();}};
function group(rows,key,val){let m={};rows.forEach(r=>{let k=r[key]||'Sin dato';m[k]=(m[k]||0)+(val?Number(r[val]||0):1)});return m}
function aggregateSAPByOrden(rows){
  const m={};
  rows.forEach(r=>{
    if(!r.ot)return;
    if(!m[r.ot])m[r.ot]={...r,hh:0,_dates:[],_rows:0};
    m[r.ot].hh += Number(r.hh||0);
    m[r.ot]._rows++;
    if(r.fecha)m[r.ot]._dates.push(r.fecha);
    if(!m[r.ot].clase || m[r.ot].clase==='Sin clase')m[r.ot].clase=r.clase;
    if(!m[r.ot].trabajo && r.trabajo)m[r.ot].trabajo=r.trabajo;
    if(!m[r.ot].resp && r.resp)m[r.ot].resp=r.resp;
    if(r.estado==='Notificada')m[r.ot].estado='Notificada';
  });
  return Object.values(m).map(o=>{
    if(o._dates.length)o.fecha=o._dates.sort().slice(-1)[0];
    delete o._dates;
    return o;
  });
}
function unique(a){return [...new Set(a.filter(Boolean))]}

function mesesPeriodo(desde,hasta){
  if(!desde||!hasta)return 1;
  const d=new Date(desde+'T00:00:00');
  const h=new Date(hasta+'T00:00:00');
  if(isNaN(d)||isNaN(h))return 1;
  let a=d<=h?d:h, b=d<=h?h:d;
  return (b.getFullYear()-a.getFullYear())*12+(b.getMonth()-a.getMonth())+1;
}
async function listarArchivosDatos(){
  try{
    const resp=await fetch(GITHUB_DATA_API+'&v='+Date.now(),{cache:'no-store'});
    if(!resp.ok)throw new Error('HTTP '+resp.status);
    const files=await resp.json();
    return files.filter(f=>f.type==='file' && /\.xlsx$/i.test(f.name));
  }catch(e){
    console.warn('API de GitHub no disponible; se usará la lista de respaldo.',e);
    return FALLBACK_DATA_FILES.map(name=>({
      name,
      type:'file',
      download_url:'datos/'+encodeURIComponent(name)
    }));
  }
}
function scoreFechaNombre(nombre){
  const n=nombre.replace(/_/g,' ');
  let m=n.match(/(20\d{2})[-. ]?(\d{2})[-. ]?(\d{2})(?:[T_ -]?(\d{2})?[:.]?(\d{2})?[:.]?(\d{2})?)?/);
  if(m){return Number(`${m[1]}${m[2]}${m[3]}${m[4]||'00'}${m[5]||'00'}${m[6]||'00'}`);}
  m=n.match(/(\d{2})[-. ](\d{2})[-. ](20\d{2})/);
  if(m){return Number(`${m[3]}${m[2]}${m[1]}000000`);}
  return 0;
}
function scoreSemanaNombre(nombre){
  const n=nombre.toLowerCase();
  let m=n.match(/semana\s*0*(\d{1,2})/i);
  let semana=m?Number(m[1]):0;
  let fecha=scoreFechaNombre(nombre);
  return fecha || semana;
}
function obtenerCandidatos(files,tipo){
  let candidatos=files.filter(f=>{
    const n=f.name.toLowerCase();
    if(tipo==='sap') return (n.includes('export') || n.includes('sap')) && !n.includes('semana') && !n.includes('plan');
    return n.includes('semana') || n.includes('plan');
  });
  candidatos.sort((a,b)=>{
    const sa=tipo==='sap'?scoreFechaNombre(a.name):scoreSemanaNombre(a.name);
    const sb=tipo==='sap'?scoreFechaNombre(b.name):scoreSemanaNombre(b.name);
    if(sb!==sa)return sb-sa;
    return b.name.localeCompare(a.name,'es',{numeric:true});
  });
  return candidatos;
}
function elegirArchivo(files,tipo){
  const candidatos=obtenerCandidatos(files,tipo);
  if(!candidatos.length)throw new Error('No se encontró archivo '+(tipo==='sap'?'SAP/EXPORT':'Plan Semanal')+' en la carpeta datos.');
  return candidatos[0];
}

async function fechaUltimaCargaGithub(file){
  try{
    const path='datos/'+file.name;
    const url=GITHUB_COMMITS_API+'?path='+encodeURIComponent(path)+'&sha='+encodeURIComponent(GITHUB_BRANCH)+'&per_page=1&v='+Date.now();
    const resp=await fetch(url,{cache:'no-store'});
    if(!resp.ok)return 0;
    const commits=await resp.json();
    const fecha=commits?.[0]?.commit?.committer?.date || commits?.[0]?.commit?.author?.date || '';
    const ts=Date.parse(fecha);
    return Number.isFinite(ts)?ts:0;
  }catch(e){
    console.warn('No se pudo consultar la fecha de carga de',file.name,e);
    return 0;
  }
}
async function elegirSapMasReciente(files){
  const candidatos=obtenerCandidatos(files,'sap');
  if(!candidatos.length)throw new Error('No se encontró archivo SAP/EXPORT en la carpeta datos.');
  // obtenerCandidatos ya ordena los EXPORT por la fecha y hora de su nombre.
  // Esto evita elegir un archivo antiguo solo porque fue vuelto a subir después.
  return candidatos[0];
}
function llenarSelectorPlanes(files, seleccionado){
  const sel=$('planSelector');
  if(!sel)return;
  const planes=obtenerCandidatos(files,'plan');
  sel.innerHTML='';
  if(!planes.length){sel.innerHTML='<option value="">Sin planes disponibles</option>';return;}
  const consolidado=document.createElement('option');
  consolidado.value='__periodo__';
  consolidado.textContent='Consolidado del período seleccionado';
  sel.appendChild(consolidado);
  planes.forEach((f,i)=>{
    const opt=document.createElement('option');
    opt.value=f.name;
    opt.textContent=(i===0?'Último: ':'')+f.name;
    sel.appendChild(opt);
  });
  sel.value=seleccionado||'__periodo__';
}
function filtrarAvisosPorEncargado(encargado){
  encargadoPlanFiltrado=encargadoPlanFiltrado===encargado?'':encargado;
  aplicarFiltroEncargado();
}
function aplicarFiltroEncargado(){
  document.querySelectorAll('#tablaPlan tbody tr').forEach(tr=>{
    const coincide=!encargadoPlanFiltrado || tr.children[4]?.textContent.trim()===encargadoPlanFiltrado;
    tr.classList.toggle('hidden',!coincide);
  });
  const titulo=$('tituloDetallePlan');
  if(titulo)titulo.textContent=encargadoPlanFiltrado?`Avisos de ${encargadoPlanFiltrado}`:'Detalle del Plan del Período';
  document.querySelectorAll('#encargadoFiltros button').forEach(btn=>{
    const activo=btn.dataset.encargado===encargadoPlanFiltrado;
    btn.classList.toggle('active',activo);
    btn.setAttribute('aria-pressed',String(activo));
  });
}
function mostrarBotonesEncargados(encargados){
  const contenedor=$('encargadoFiltros');
  if(!contenedor)return;
  contenedor.innerHTML='';
  encargados.forEach(encargado=>{
    const btn=document.createElement('button');
    btn.type='button';
    btn.dataset.encargado=encargado;
    btn.textContent=encargado;
    btn.setAttribute('aria-pressed','false');
    btn.addEventListener('click',()=>filtrarAvisosPorEncargado(encargado));
    contenedor.appendChild(btn);
  });
  aplicarFiltroEncargado();
}
function activarFiltroGraficoEncargados(){
  const grafico=charts.chartEnc;
  if(!grafico)return;
  grafico.options.onClick=(evento,elementos)=>{
    let indice=elementos?.[0]?.index;
    if(indice==null && grafico.scales.x && evento.x>=grafico.chartArea.left && evento.x<=grafico.chartArea.right){
      indice=Math.round(grafico.scales.x.getValueForPixel(evento.x));
    }
    const encargado=grafico.data.labels[indice];
    if(encargado)filtrarAvisosPorEncargado(encargado);
  };
  grafico.options.onHover=(evento)=>{evento.native.target.style.cursor='pointer'};
  grafico.update();
}
function seleccionarUltimoPlanDisponible(actualizar=true,sincronizarPeriodo=false){
  const ultimo=obtenerCandidatos(githubArchivos,'plan')[0];
  if(!ultimo || !planesCargados[ultimo.name])return;
  const sel=$('planSelector');
  if(sel)sel.value=ultimo.name;
  plan=planesCargados[ultimo.name];
  planArchivo='datos/'+ultimo.name;
  if(sincronizarPeriodo){
    const ultimaFecha=plan.map(p=>p.fecha).filter(f=>/^\d{4}-\d{2}-\d{2}$/.test(f)).sort().pop();
    const semana=semanaIsoDesdeFecha(ultimaFecha);
    if(semana){
      $('tipoPeriodo').value='semanal';
      $('semanaPeriodo').value=semana;
      $('controlSemana').classList.remove('hidden');
      $('controlMes').classList.add('hidden');
      $('controlAnio').classList.add('hidden');
      document.querySelectorAll('.periodSwitch [data-periodo]').forEach(btn=>btn.classList.toggle('active',btn.dataset.periodo==='semanal'));
      const rango=rangoSemanaIso(semana);
      $('desde').value=rango.desde;
      $('hasta').value=rango.hasta;
    }
  }
  if(actualizar)render();
}
async function cambiarPlanSemanal(){
  const sel=$('planSelector');
  if(!sel||!sel.value||!githubArchivos.length)return;
  if(sel.value==='__periodo__'){
    plan=Object.values(planesCargados).flat();
    planArchivo='Planes consolidados del período';
    render();
    return;
  }
  const file=githubArchivos.find(f=>f.name===sel.value);
  if(!file)return;
  try{
    $('estadoCarga').innerHTML=`Cargando plan seleccionado <b>${file.name}</b>... <span class="gray pill">procesando</span>`;
    planRows=await leerExcelGithub(file);
    plan=mapPlan(planRows).map(p=>({...p,origen:file.name}));
    planArchivo='datos/'+file.name;
    ultimaCarga=new Date().toLocaleString('es-CL');
    render();
  }catch(e){
    console.error(e);
    $('estadoCarga').innerHTML=`No se pudo cargar el plan seleccionado. <span class="bad pill">${e.message}</span>`;
  }
}
async function cargarPlanesConsolidados(files){
  const candidatos=obtenerCandidatos(files,'plan');
  const resultados=await Promise.allSettled(candidatos.map(async file=>{
    const rows=await leerExcelGithub(file);
    return [file.name,mapPlan(rows).map(p=>({...p,origen:file.name}))];
  }));
  planesCargados={};
  resultados.forEach(r=>{if(r.status==='fulfilled')planesCargados[r.value[0]]=r.value[1]});
  plan=Object.values(planesCargados).flat();
  planRows=[];
  planArchivo='Planes consolidados del período';
}
async function leerExcelGithub(file){
  const base=(file.download_url || file);
  const url=base+(base.includes('?')?'&':'?')+'v='+Date.now();
  const resp=await fetch(url,{cache:'no-store'});
  if(!resp.ok)throw new Error('No se pudo leer '+(file.name||file)+' (HTTP '+resp.status+')');
  const buffer=await resp.arrayBuffer();
  return parseWB(buffer);
}
async function cargarDatosGithub(){
  try{
    $('estadoCarga').innerHTML='Buscando el último archivo SAP cargado en GitHub... <span class="gray pill">esperando</span>';
    const archivos=await listarArchivosDatos();
    githubArchivos=archivos;
    const sapFile=await elegirSapMasReciente(archivos);
    // Cada carga/actualización parte desde la semana más reciente disponible.
    // El usuario aún puede revisar semanas anteriores usando el selector.
    const planFile=elegirArchivo(archivos,'plan');
    llenarSelectorPlanes(archivos, planFile.name);
    sapArchivo='datos/'+sapFile.name;
    planArchivo='Planes consolidados del período';
    $('estadoCarga').innerHTML=`Cargando <b>${sapFile.name}</b> y <b>${planFile.name}</b>... <span class="gray pill">procesando</span>`;
    sapRows=await leerExcelGithub(sapFile);
    sap=mapSAP(sapRows);
    llenarSelectorAnios();
    const ultimaFecha=sap.map(x=>x.fecha).filter(x=>/^\d{4}-\d{2}-\d{2}$/.test(x)).sort().pop();
    if(ultimaFecha){$('mesPeriodo').value=ultimaFecha.slice(0,7);$('anioPeriodo').value=ultimaFecha.slice(0,4);aplicarPeriodoSeleccionado(false)}
    await cargarPlanesConsolidados(archivos);
    seleccionarUltimoPlanDisponible(false,$('plan').classList.contains('active'));
    ultimaCarga=new Date().toLocaleString('es-CL');
    render();
  }catch(e){
    console.error(e);
    render();
    $('estadoCarga').innerHTML=`No se pudieron cargar los datos. Verifica que los Excel estén dentro de <b>datos/</b>. <span class="bad pill">${e.message}</span>`;
  }
}
function render(){let desde=$('desde').value,hasta=$('hasta').value,metaMensual=META_HH_MENSUAL,meses=mesesPeriodo(desde,hasta),meta=metaMensual*meses;let sapF=sap.filter(x=>(!desde||!x.fecha||x.fecha>=desde)&&(!hasta||!x.fecha||x.fecha<=hasta));let clases=unique(sap.map(x=>x.clase));let sel=$('claseFiltro');let current=sel.value;if(sel.options.length<=1){clases.forEach(c=>{let o=document.createElement('option');o.textContent=c;sel.appendChild(o)})} if(current&&current!=='Todas')sapF=sapF.filter(x=>x.clase===current);
 let sapCalc=aggregateSAPByOrden(sapF);
 let hh=sapCalc.reduce((a,b)=>a+b.hh,0); let ot=unique(sapCalc.filter(x=>x.estado==='Notificada').map(x=>x.ot)).length; let prom=ot?hh/ot:0; let pct=meta?hh/meta*100:0; let desv=hh-meta; $('hhReal').textContent=fmt(hh);$('hhMeta').textContent=fmt(meta);$('hhMeta').title='Meta mensual '+fmt(metaMensual)+' × '+meses+' mes(es)';$('desvHH').textContent=fmt(desv);$('desvPct').textContent=fmt(meta?desv/meta*100:0)+'%';$('cumplHH').textContent=fmt(pct)+'%';$('topCumpl').textContent=fmt(pct)+'%';$('topMeta').textContent=pct>=100?'Sobre la meta':'Bajo la meta';$('topCumpl').style.color=pct>=100?COLORS.green:COLORS.red;$('otEjecutadas').textContent=ot;$('promHH').textContent=fmt(prom);$('periodoTxt').textContent=`${showDate(desde)} al ${showDate(hasta)}`;$('periodoTxt2').textContent=$('periodoTxt').textContent;$('updateTxt').textContent=new Date().toLocaleString('es-CL');$('estadoCarga').innerHTML=`SAP: <b>${sap.length}</b> registros leídos <span class="gray pill">${sapArchivo}</span> | Plan semanal: <b>${plan.length}</b> OT leídas <span class="gray pill">${planArchivo}</span> | Última lectura: <span class="ok pill">${ultimaCarga}</span>`;
 let dias={};sapCalc.forEach(x=>{let f=x.fecha||'Sin fecha'; if(!dias[f])dias[f]={hh:0,ots:new Set()};dias[f].hh+=x.hh;dias[f].ots.add(x.ot)});let labels=Object.keys(dias).sort();let metaDia=metaMensual/30;
 let periodoMes=periodoMesesLabel(desde,hasta);
 if($('tituloChartMes'))$('tituloChartMes').textContent='HH Real vs Meta HH - '+periodoMes;
 if($('tituloChartTipo'))$('tituloChartTipo').firstChild.textContent='Órdenes por tipo de mantenimiento - '+periodoMes;
 if($('tituloChartAcum'))$('tituloChartAcum').textContent='Cumplimiento acumulado del período - '+periodoMes;
 let mesesGraf={};sapCalc.forEach(x=>{let k=(x.fecha&&/^\d{4}-\d{2}/.test(x.fecha))?x.fecha.slice(0,7):'Sin fecha'; if(!mesesGraf[k])mesesGraf[k]={hh:0};mesesGraf[k].hh+=x.hh});let labelsMes=Object.keys(mesesGraf).sort();destroy('chartDiario');charts.chartDiario=new Chart($('chartDiario'),{type:'bar',data:{labels:labelsMes.map(showMes),datasets:[{label:'HH Real',data:labelsMes.map(f=>mesesGraf[f].hh),borderWidth:1,backgroundColor:COLORS.blue,borderColor:COLORS.blue},{label:'Meta HH',data:labelsMes.map(()=>metaMensual),borderWidth:1,backgroundColor:COLORS.green,borderColor:COLORS.green}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top'}},scales:{y:{beginAtZero:true,title:{display:true,text:'HH'}}}}});let tb=$('tablaDia').querySelector('tbody');tb.innerHTML='';labels.forEach(f=>{let h=dias[f].hh;let dif=h-metaDia;let cls=dif>=0?'positive':'negative';tb.insertAdjacentHTML('beforeend',`<tr><td>${showDate(f)}</td><td>${fmt(h)}</td><td>${fmt(metaDia)}</td><td class="right ${cls}">${fmt(dif)}</td><td>${fmt(metaDia?h/metaDia*100:0)}%</td><td>${dias[f].ots.size}</td></tr>`)});
 let claseRaw=group(sapCalc,'clase','hh');let ordenClases=['ZM01','ZM02','ZM05'];let claseLabels=[],claseData=[],claseColors=[];ordenClases.concat(Object.keys(claseRaw).filter(k=>!ordenClases.includes(k))).forEach(k=>{if(claseRaw[k]){claseLabels.push(CLASS_INFO[k]?`${k} - ${CLASS_INFO[k]}`:k);claseData.push(claseRaw[k]);claseColors.push(k==='ZM01'?COLORS.sky:k==='ZM02'?COLORS.pink:k==='ZM05'?COLORS.orange:COLORS.gray)}});chart('chartTipo','pie',claseLabels,claseData,'HH',claseColors);let acumHH=[],acumMeta=[],a=0,m=0;labels.forEach(f=>{a+=dias[f].hh;m+=metaDia;acumHH.push(a);acumMeta.push(m)});destroy('chartAcum');charts.chartAcum=new Chart($('chartAcum'),{type:'line',data:{labels:labels.map(showDate),datasets:[{label:'HH Real acumulado',data:acumHH,borderColor:COLORS.blue,backgroundColor:COLORS.blue,borderWidth:3,borderDash:[],pointRadius:0,pointHoverRadius:4,spanGaps:true,tension:.25},{label:'Meta HH acumulada',data:acumMeta,borderColor:COLORS.green,backgroundColor:COLORS.green,borderWidth:3,borderDash:[],pointRadius:0,pointHoverRadius:4,spanGaps:true,tension:.25}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top'}},scales:{y:{beginAtZero:true}}}});
 let sapByOT={};sap.forEach(s=>{if(!sapByOT[s.ot])sapByOT[s.ot]={not:false}; if(s.estado==='Notificada')sapByOT[s.ot].not=true});
 let planPeriodo=plan.filter(p=>p.fecha && (!desde||p.fecha>=desde) && (!hasta||p.fecha<=hasta));
 let planUnico={};planPeriodo.forEach(p=>{if(!planUnico[p.ot])planUnico[p.ot]={...p,origenes:[]};if(p.origen&&!planUnico[p.ot].origenes.includes(p.origen))planUnico[p.ot].origenes.push(p.origen);if(p.fecha<planUnico[p.ot].fecha)planUnico[p.ot].fecha=p.fecha});
 let planCalc=Object.values(planUnico).sort((a,b)=>(a.fecha||'').localeCompare(b.fecha||'')||a.ot.localeCompare(b.ot)).map(p=>{const terreno=estadosTerreno[String(p.aviso||'').trim()];return{...p,estado:sapByOT[p.ot]?.not?'Notificada':(terreno?.estado_terreno||'Pendiente')}});planActualPorAviso={};planCalc.forEach(p=>{if(p.aviso)planActualPorAviso[String(p.aviso).trim()]=p}); let total=planCalc.length,notif=planCalc.filter(x=>x.estado==='Notificada').length,pend=total-notif,pp=total?Math.round(notif/total*100):0; $('planTotal').textContent=total;$('planNotif').textContent=notif;$('planPend').textContent=pend;$('planPct').textContent=pp+'%';$('planBar').style.width=pp+'%';chart('chartPlan','doughnut',['Notificadas','Pendientes'],[notif,pend],'OT',[COLORS.green,COLORS.red]);let enc={};planCalc.forEach(p=>{let k=p.encargado||'Sin encargado';if(!enc[k])enc[k]={t:0,n:0};enc[k].t++;if(p.estado==='Notificada')enc[k].n++});const encargados=Object.keys(enc);chart('chartEnc','bar',encargados,Object.values(enc).map(x=>x.t?Math.round(x.n/x.t*100):0),'% cumplimiento');mostrarBotonesEncargados(encargados);activarFiltroGraficoEncargados();
 let tp=$('tablaPlan').querySelector('tbody');tp.innerHTML='';planCalc.forEach(p=>{const estadoClase=p.estado==='Notificada'?'ok':p.estado==='Realizado'?'realizado':p.estado==='En ejecución'?'ejecucion':'bad';tp.insertAdjacentHTML('beforeend',`<tr><td>${showDate(p.fecha)}</td><td><span class="copiable-sap" data-copy="${p.aviso}" data-tipo="Aviso" title="Clic para copiar aviso">${p.aviso}</span></td><td><b class="copiable-sap" data-copy="${p.ot}" data-tipo="Orden" title="Clic para copiar orden">${p.ot}</b></td><td>${p.trabajo}</td><td>${p.encargado}</td><td>${p.turno}</td><td><button type="button" class="pill estadoBtn ${estadoClase}" data-aviso="${encodeURIComponent(String(p.aviso||''))}" title="Abrir actualización del aviso">${p.estado}</button></td></tr>`)});aplicarFiltroEncargado();}
function showView(id){document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));$(id).classList.add('active');document.querySelectorAll('.side button[data-view]').forEach(b=>b.classList.toggle('active',b.dataset.view===id));if(id==='plan')seleccionarUltimoPlanDisponible(true,true)}
function exportarCSV(){let rows=[['Fecha','Aviso','OT','Trabajo','Encargado','Turno','Estado']];document.querySelectorAll('#tablaPlan tbody tr').forEach(tr=>rows.push([...tr.children].map(td=>td.innerText)));let csv=rows.map(r=>r.map(x=>'"'+String(x).replace(/"/g,'""')+'"').join(';')).join('\n');let a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download='cumplimiento_plan_semanal.csv';a.click()}

async function copiarTextoSAP(texto,tipo,elemento){
  const valor=String(texto||'').trim();
  if(!valor)return;
  try{
    if(navigator.clipboard && window.isSecureContext){
      await navigator.clipboard.writeText(valor);
    }else{
      const area=document.createElement('textarea');
      area.value=valor;
      area.style.position='fixed';
      area.style.opacity='0';
      document.body.appendChild(area);
      area.select();
      document.execCommand('copy');
      area.remove();
    }
    const original=elemento.textContent;
    elemento.classList.add('copiado');
    elemento.textContent='✓ '+valor;
    setTimeout(()=>{
      elemento.textContent=original;
      elemento.classList.remove('copiado');
    },900);
  }catch(e){
    console.error(e);
    alert('No se pudo copiar. Selecciona el número y usa Ctrl + C.');
  }
}
document.addEventListener('click',e=>{
  const el=e.target.closest('.copiable-sap');
  if(el)copiarTextoSAP(el.dataset.copy,el.dataset.tipo,el);
  const estadoBtn=e.target.closest('.estadoBtn');
  if(estadoBtn)abrirEstadoAviso(decodeURIComponent(estadoBtn.dataset.aviso||''));
});

function fechaInput(valor){
  if(!valor)return'';
  const d=new Date(valor);
  if(Number.isNaN(d.getTime()))return'';
  const local=new Date(d.getTime()-d.getTimezoneOffset()*60000);
  return local.toISOString().slice(0,16);
}
function ahoraInput(){return fechaInput(new Date().toISOString())}
function actualizarFechasEstado(){
  const estado=$('estadoTerreno').value;
  if(estado==='Pendiente'){$('estadoInicio').value='';$('estadoTermino').value=''}
  if(estado==='En ejecución'&&!$('estadoInicio').value)$('estadoInicio').value=ahoraInput();
  if(estado==='Realizado'){
    if(!$('estadoInicio').value)$('estadoInicio').value=ahoraInput();
    if(!$('estadoTermino').value)$('estadoTermino').value=ahoraInput();
  }
}
function sincronizarEstadoPorFechas(){
  const inicio=$('estadoInicio').value;
  const termino=$('estadoTermino').value;
  $('estadoInicio').setCustomValidity('');
  $('estadoTermino').setCustomValidity('');
  if(termino&&!inicio){
    const mensaje='Debe completar la fecha y hora de inicio antes de informar el término.';
    $('estadoInicio').setCustomValidity(mensaje);
    $('estadoMensaje').textContent='⚠ '+mensaje;
    return false;
  }
  if(inicio&&termino&&new Date(termino)<new Date(inicio)){
    const mensaje='La fecha y hora de término no puede ser anterior al inicio.';
    $('estadoTermino').setCustomValidity(mensaje);
    $('estadoMensaje').textContent='⚠ '+mensaje;
    return false;
  }
  if(inicio&&termino)$('estadoTerreno').value='Realizado';
  else if(inicio)$('estadoTerreno').value='En ejecución';
  else $('estadoTerreno').value='Pendiente';
  $('estadoMensaje').textContent=inicio&&termino?'Estado actualizado automáticamente a Realizado.':inicio?'Estado actualizado automáticamente a En ejecución.':'';
  return true;
}
function abrirEstadoAviso(aviso){
  const planItem=planActualPorAviso[String(aviso||'').trim()];
  if(!planItem)return;
  avisoEstadoAbierto=String(aviso).trim();
  const guardado=estadosTerreno[avisoEstadoAbierto]||{};
  const notificada=planItem.estado==='Notificada';
  $('estadoAviso').textContent=avisoEstadoAbierto||'Sin aviso';
  $('estadoOrden').textContent=planItem.ot||'—';
  $('estadoTrabajo').textContent=planItem.trabajo||'Sin descripción';
  $('estadoTerreno').value=notificada?'Realizado':(guardado.estado_terreno||'Pendiente');
  $('estadoInicio').value=fechaInput(guardado.fecha_inicio);
  $('estadoTermino').value=fechaInput(guardado.fecha_termino);
  $('estadoDetalle').value=guardado.detalle_realizado||'';
  $('estadoSapInfo').classList.toggle('hidden',!notificada);
  ['estadoTerreno','estadoInicio','estadoTermino','estadoDetalle'].forEach(id=>$(id).disabled=notificada);
  $('guardarEstadoBtn').classList.toggle('hidden',notificada);
  $('estadoMensaje').textContent='';
  $('modalEstado').classList.add('open');
  document.body.classList.add('modalOpen');
}
function cerrarEstadoAviso(){
  $('modalEstado').classList.remove('open');
  document.body.classList.remove('modalOpen');
  avisoEstadoAbierto='';
}
function actualizarEstadosVisibles(){
  document.querySelectorAll('.estadoBtn').forEach(btn=>{
    const aviso=decodeURIComponent(btn.dataset.aviso||'');
    const planItem=planActualPorAviso[aviso];
    if(!planItem||planItem.estado==='Notificada')return;
    const nuevoEstado=estadosTerreno[aviso]?.estado_terreno||'Pendiente';
    if(btn.textContent===nuevoEstado)return;
    planItem.estado=nuevoEstado;
    btn.textContent=nuevoEstado;
    btn.classList.remove('ok','bad','ejecucion','realizado');
    btn.classList.add(nuevoEstado==='Realizado'?'realizado':nuevoEstado==='En ejecución'?'ejecucion':'bad');
  });
}
async function cargarEstadosTerreno(silencioso=false){
  try{
    const r=await fetch(`${ESTADOS_API_URL}?t=${Date.now()}`,{cache:'no-store',redirect:'follow'});
    const data=await r.json();
    if(!data.ok)throw new Error(data.error||'No fue posible leer los estados.');
    const nuevos={};(data.items||[]).forEach(item=>{if(item.aviso)nuevos[String(item.aviso).trim()]=item});
    const huboCambios=JSON.stringify(nuevos)!==JSON.stringify(estadosTerreno);
    estadosTerreno=nuevos;
    if(huboCambios)actualizarEstadosVisibles();
  }catch(error){if(!silencioso)console.error('Estados de terreno:',error)}
}
async function guardarEstadoAviso(){
  const btn=$('guardarEstadoBtn');
  if(!sincronizarEstadoPorFechas()){
    $('estadoInicio').reportValidity();
    $('estadoTermino').reportValidity();
    return;
  }
  btn.disabled=true;$('estadoMensaje').textContent='Guardando en Google Drive…';
  try{
    const body=new URLSearchParams({aviso:avisoEstadoAbierto,estado_terreno:$('estadoTerreno').value,fecha_inicio:$('estadoInicio').value,fecha_termino:$('estadoTermino').value,detalle_realizado:$('estadoDetalle').value.trim()});
    const r=await fetch(ESTADOS_API_URL,{method:'POST',body,redirect:'follow'});
    const data=await r.json();
    if(!data.ok)throw new Error(data.error||'No fue posible guardar.');
    estadosTerreno[avisoEstadoAbierto]=data.item;
    $('estadoMensaje').textContent='✓ Cambio guardado y sincronizado.';
    actualizarEstadosVisibles();
    setTimeout(cerrarEstadoAviso,650);
  }catch(error){$('estadoMensaje').textContent='No se pudo guardar: '+error.message}
  finally{btn.disabled=false}
}

cargarDatosGithub().then(()=>cargarEstadosTerreno());
setInterval(()=>cargarEstadosTerreno(true),20000);


function safeText(id){const el=$(id); return el?el.innerText.trim():''}
function canvasImg(id, cls='pdfChart', images={}){
  const src=images[id];
  if(!src || src==='data:,') return '<div style="font-size:10px;color:#64748b">Gráfico no disponible</div>';
  return `<img class="${cls}" src="${src}">`;
}
function nextFrame(){return new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))}
async function capturarGraficosInforme(){
  render();
  const tamanosPdf={
    chartDiario:[500,270],
    chartAcum:[500,270],
    chartTipo:[460,230],
    chartPlan:[460,220],
    chartEnc:[500,220]
  };
  await nextFrame();
  const images={};
  Object.entries(tamanosPdf).forEach(([id,[ancho,alto]])=>{
    let copia;
    try{
      const origen=charts[id];
      if(!origen)throw new Error('Gráfico no disponible');
      const tipo=origen.config.type;
      const circular=tipo==='pie'||tipo==='doughnut';
      const canvas=document.createElement('canvas');
      canvas.width=ancho*2;
      canvas.height=alto*2;
      copia=new Chart(canvas,{
        type:tipo,
        data:JSON.parse(JSON.stringify(origen.data)),
        options:{
          responsive:false,animation:false,maintainAspectRatio:false,devicePixelRatio:2,
          layout:{padding:circular?12:8},
          plugins:{
            legend:{display:true,position:'top',labels:{usePointStyle:circular,boxWidth:14,font:{size:13,weight:'bold'}}},
            tooltip:{enabled:false}
          },
          scales:circular?{}:{
            x:{ticks:{font:{size:11},maxRotation:45,minRotation:0},grid:{display:false}},
            y:{beginAtZero:true,ticks:{font:{size:11}},grid:{color:'#dbe4ef'}}
          },
          elements:{line:{borderWidth:4},point:{radius:0}},
          cutout:tipo==='doughnut'?'52%':undefined
        },
        plugins:circular?[piePercentPlugin]:[]
      });
      copia.resize(ancho,alto);
      copia.update('none');
      images[id]=copia.toBase64Image('image/png',1);
    }catch(e){
      console.error('No se pudo preparar '+id+' para el informe',e);
      images[id]='';
    }finally{
      if(copia)copia.destroy();
    }
  });
  return images;
}
function buildPdfTablePlan(){
  const rows=[...document.querySelectorAll('#tablaPlan tbody tr')];
  let html='<table class="pdfPlanTable"><thead><tr><th>Fecha</th><th>Aviso</th><th>Orden</th><th>Trabajo</th><th>Encargado</th><th>Turno</th><th>Estado</th></tr></thead><tbody>';
  rows.slice(0,28).forEach(tr=>{
    const t=[...tr.children].map(td=>td.innerText.trim());
    const estado=(t[6]||'Pendiente').toLowerCase();
    const badge=estado.includes('notificada')?`<span class="pdfBadgeOk">${t[6]}</span>`:`<span class="pdfBadgeBad">${t[6]||'Pendiente'}</span>`;
    html+=`<tr><td>${t[0]||''}</td><td>${t[1]||''}</td><td>${t[2]||''}</td><td>${t[3]||''}</td><td>${t[4]||''}</td><td>${t[5]||''}</td><td>${badge}</td></tr>`;
  });
  if(rows.length>28){html+=`<tr><td colspan="7">Se muestran las primeras 28 OT de ${rows.length}. El detalle completo permanece disponible en el dashboard.</td></tr>`}
  html+='</tbody></table>'; return html;
}
function buildPdfTableDia(rows=[...document.querySelectorAll('#tablaDia tbody tr')]){
  let html='<table class="pdfDailyTable"><thead><tr><th>Fecha</th><th>HH Real</th><th>Meta HH</th><th>Desviación</th><th>Cumplimiento</th><th>Órdenes</th></tr></thead><tbody>';
  rows.forEach(tr=>{
    const t=[...tr.children].map(td=>td.innerText.trim());
    const desviacion=Number((t[3]||'0').replace(/\./g,'').replace(',','.'));
    const cls=desviacion>=0?'pdfPositive':'pdfNegative';
    html+=`<tr><td>${t[0]||''}</td><td>${t[1]||''}</td><td>${t[2]||''}</td><td class="${cls}">${t[3]||''}</td><td>${t[4]||''}</td><td>${t[5]||''}</td></tr>`;
  });
  if(!rows.length)html+='<tr><td colspan="6">No hay registros diarios para el período seleccionado.</td></tr>';
  html+='</tbody></table>';
  return html;
}
function abrirSelectorInforme(){$('modalInforme').classList.add('open')}
function cerrarSelectorInforme(){$('modalInforme').classList.remove('open')}
function generarInformeSeleccionado(){
  const secciones=[...document.querySelectorAll('input[name="paginaInforme"]:checked')].map(x=>x.value);
  if(!secciones.length){alert('Selecciona al menos una sección para generar el informe.');return}
  cerrarSelectorInforme();
  generarInformePDF(true,secciones);
}
document.addEventListener('keydown',e=>{if(e.key==='Escape')cerrarSelectorInforme()});
async function generarInformePDF(imprimir=true,secciones=['resumen','diario','plan']){
  const chartImages=await capturarGraficosInforme();
  const now=new Date().toLocaleString('es-CL');
  const logo=document.querySelector('.logo img')?.src||'';
  const periodo=safeText('periodoTxt');
  const hh=safeText('hhReal'), meta=safeText('hhMeta'), desv=safeText('desvHH'), cumpl=safeText('cumplHH'), ots=safeText('otEjecutadas'), prom=safeText('promHH');
  const pt=safeText('planTotal'), pn=safeText('planNotif'), pp=safeText('planPend'), pc=safeText('planPct');
  const comentario1=`Durante el período analizado se registraron ${hh} HH frente a una meta de ${meta} HH, alcanzando un cumplimiento de ${cumpl}. Se ejecutaron ${ots} órdenes de mantenimiento, con un promedio de ${prom} HH por orden.`;
  const comentario2=`Durante la semana se programaron ${pt} órdenes de trabajo, de las cuales ${pn} fueron notificadas. Permanecen ${pp} OT pendientes, alcanzando un cumplimiento del Plan Semanal de ${pc}.`;
  const filasDiarias=[...document.querySelectorAll('#tablaDia tbody tr')];
  const bloquesDiarios=[];
  if(!filasDiarias.length)bloquesDiarios.push([]);else for(let i=0;i<filasDiarias.length;i+=26)bloquesDiarios.push(filasDiarias.slice(i,i+26));
  const paginasDiarias=bloquesDiarios.map((filas,i)=>`
  <section class="pdfPage pdfPageDiaria" data-seccion="diario" style="position:relative">
    <div class="pdfHead"><img class="pdfLogo" src="${logo}"><div class="pdfTitle"><h1>INFORME EJECUTIVO HH MANTENCIÓN SAP</h1><h2>Resumen diario del período${bloquesDiarios.length>1?` - Parte ${i+1} de ${bloquesDiarios.length}`:''}</h2></div><div class="pdfMeta"><b>Período:</b><br>${periodo}<br><br><b>Emisión:</b><br>${now}</div></div>
    <div class="pdfKpis">
      <div class="pdfKpi"><b>HH Reales</b><span>${hh}</span></div><div class="pdfKpi"><b>Meta HH</b><span>${meta}</span></div><div class="pdfKpi"><b>Desviación</b><span>${desv}</span></div><div class="pdfKpi"><b>Cumplimiento</b><span>${cumpl}</span></div><div class="pdfKpi"><b>Órdenes</b><span>${ots}</span></div><div class="pdfKpi"><b>Prom. HH/OT</b><span>${prom}</span></div>
    </div>
    <div class="pdfBox"><h3>Detalle diario de HH y cumplimiento</h3>${buildPdfTableDia(filas)}</div>
    <div class="pdfFooter"><span>Dashboard HH Mantención SAP – Piscicultura Lago Verde</span><span></span></div>
  </section>`).join('');
  const rep=$('printReport');
  rep.innerHTML=`
  <section class="pdfPage pdfPageResumen" data-seccion="resumen" style="position:relative">
    <div class="pdfHead"><img class="pdfLogo" src="${logo}"><div class="pdfTitle"><h1>INFORME EJECUTIVO HH MANTENCIÓN SAP</h1><h2>Piscicultura Lago Verde</h2></div><div class="pdfMeta"><b>Período:</b><br>${periodo}<br><br><b>Emisión:</b><br>${now}</div></div>
    <div class="pdfKpis">
      <div class="pdfKpi"><b>HH Reales</b><span>${hh}</span></div><div class="pdfKpi"><b>Meta HH</b><span>${meta}</span></div><div class="pdfKpi"><b>Desviación</b><span>${desv}</span></div><div class="pdfKpi"><b>Cumplimiento</b><span>${cumpl}</span></div><div class="pdfKpi"><b>Órdenes</b><span>${ots}</span></div><div class="pdfKpi"><b>Prom. HH/OT</b><span>${prom}</span></div>
    </div>
    <div class="pdfBox"><h3>Comentario ejecutivo</h3><div class="pdfComment">${comentario1}</div></div>
    <div class="pdfGrid2"><div class="pdfBox"><h3>HH Real vs Meta</h3>${canvasImg('chartDiario','pdfChartTall',chartImages)}</div><div class="pdfBox"><h3>Cumplimiento acumulado</h3>${canvasImg('chartAcum','pdfChartTall',chartImages)}</div></div>
    <div class="pdfBox"><h3>Órdenes por tipo de mantenimiento</h3>${canvasImg('chartTipo','pdfChartPie',chartImages)}</div>
    <div class="pdfFooter"><span>Dashboard HH Mantención SAP – Piscicultura Lago Verde</span><span>Página 1 de 3</span></div>
  </section>
  ${paginasDiarias}
  <section class="pdfPage pdfPagePlan" data-seccion="plan" style="position:relative">
    <div class="pdfHead"><img class="pdfLogo" src="${logo}"><div class="pdfTitle"><h1>INFORME EJECUTIVO HH MANTENCIÓN SAP</h1><h2>Cumplimiento Plan Semanal</h2></div><div class="pdfMeta"><b>Período:</b><br>${periodo}<br><br><b>Emisión:</b><br>${now}</div></div>
    <div class="pdfKpis" style="grid-template-columns:repeat(4,1fr)">
      <div class="pdfKpi"><b>OT Planificadas</b><span>${pt}</span></div><div class="pdfKpi"><b>OT Notificadas</b><span>${pn}</span></div><div class="pdfKpi"><b>OT Pendientes</b><span>${pp}</span></div><div class="pdfKpi"><b>Cumplimiento Plan</b><span>${pc}</span></div>
    </div>
    <div class="pdfGrid2"><div class="pdfBox"><h3>Cumplimiento Plan Semanal</h3>${canvasImg('chartPlan','pdfChart',chartImages)}</div><div class="pdfBox"><h3>Cumplimiento por encargado</h3>${canvasImg('chartEnc','pdfChart',chartImages)}</div></div>
    <div class="pdfBox"><h3>Resumen Plan Semanal</h3><div class="pdfComment">${comentario2}</div></div>
    <div class="pdfBox"><h3>Detalle Plan Semanal</h3>${buildPdfTablePlan()}</div>
    <div class="pdfFooter"><span>Dashboard HH Mantención SAP – Piscicultura Lago Verde</span><span>Página 3 de 3</span></div>
  </section>`;
  rep.querySelectorAll('.pdfPage').forEach(p=>{if(!secciones.includes(p.dataset.seccion))p.remove()});
  const paginas=[...rep.querySelectorAll('.pdfPage')];
  paginas.forEach((p,i)=>{const numero=p.querySelector('.pdfFooter span:last-child');if(numero)numero.textContent=`Página ${i+1} de ${paginas.length}`});
  if(imprimir)setTimeout(()=>window.print(),300);
}
