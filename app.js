// ============================================================
// KPI Dashboard System - app.js v4.0
// กลุ่มภารกิจด้านการพยาบาล โรงพยาบาลชลบุรี
// พัฒนาโดย พว.วิษณุกรณ์ โอยา (Charge Nurse)
// ============================================================
'use strict';

// ═══ CONFIG ═══════════════════════════════════════════════
const GAS_URL = 'https://script.google.com/macros/s/AKfycbzVDJGjjyVQ2Fmicx5iFC_xFRgjZ6TTytT1i-16W-t7O8iLmlpurziONzoCq0MEMfRZ/exec';
const USE_MOCK = false;
const API_TIMEOUT_MS = 30000; // 30 วินาที

// ═══ STATE ════════════════════════════════════════════════
const State = {
  token:null, user:null, units:[], groups:[], kpiMaster:[],
  dashboardData:null, currentPage:null, adminTab:'users',
  entry:{ month:null, year:null, unitId:null },
  charts:{ trend:null, pie:null }, wardCharts:[],
};

// ═══ CONSTANTS ════════════════════════════════════════════
const TH_MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
const TH_MONTHS_FULL = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
const WARD_QUOTES = [
  {e:'💚',t:'<strong>พยาบาลคือหัวใจของโรงพยาบาล</strong> — ทุกความเอาใจใส่ที่ท่านมอบให้ผู้ป่วย คือความดีงามที่ยิ่งใหญ่'},
  {e:'🌟',t:'<strong>ทุกตัวเลขในรายงานนี้ คือชีวิตที่ท่านดูแล</strong> — ขอบคุณสำหรับทุกเวรที่ท่านทุ่มเท'},
  {e:'🏥',t:'วิชาชีพพยาบาลคือ<strong>พันธกิจแห่งความรัก</strong> — ผลงานของท่านสร้างความเชื่อมั่นให้ผู้ป่วย'},
  {e:'✨',t:'<strong>คุณภาพการพยาบาลที่ดี</strong> เริ่มต้นจากหัวใจที่มุ่งมั่น เช่นเดียวกับทุกคนในทีม'},
  {e:'🌱',t:'ตัวชี้วัดที่ดีขึ้นทุกเดือน คือ<strong>การเติบโตของทีมเรา</strong> — ภาคภูมิใจในสิ่งที่ทำร่วมกัน'},
  {e:'💫',t:'<strong>ความปลอดภัยของผู้ป่วย</strong> คือพันธสัญญาที่เราให้กันทุกวัน — ท่านทำได้ดีมาก'},
];
const CHART_COLORS=[
  {s:'#359286',l:'rgba(53,146,134,0.12)'},{s:'#f59e0b',l:'rgba(245,158,11,0.12)'},
  {s:'#3b82f6',l:'rgba(59,130,246,0.12)'},{s:'#8b5cf6',l:'rgba(139,92,246,0.12)'},
  {s:'#ef4444',l:'rgba(239,68,68,0.12)'},{s:'#ec4899',l:'rgba(236,72,153,0.12)'},
];
const MIN_BE=2560, MAX_BE=2585;
let dashView='table', wardView='table';
let dashBarChart=null, dashLineChart=null, wardBarChart=null, wardLineChart=null;

// ═══ INIT ═════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded',()=>{
  initDateTime(); initEntryMonth(); checkStoredSession();
  document.getElementById('login-password')?.addEventListener('keydown',e=>{if(e.key==='Enter')handleLogin();});
  document.getElementById('login-username')?.addEventListener('keydown',e=>{if(e.key==='Enter')handleLogin();});
});
function initDateTime(){
  const up=()=>{
    const el=document.getElementById('topbar-datetime'); if(!el)return;
    const n=new Date();
    el.innerHTML=n.toLocaleDateString('th-TH',{weekday:'short',year:'numeric',month:'short',day:'numeric'})
               +'<br>'+n.toLocaleTimeString('th-TH',{hour:'2-digit',minute:'2-digit'});
  };
  up(); setInterval(up,30000);
}
function initEntryMonth(){ const n=new Date(); State.entry.month=n.getMonth(); State.entry.year=n.getFullYear(); updateMonthDisplay(); }
function checkStoredSession(){
  const t=sessionStorage.getItem('kpi_token'),u=sessionStorage.getItem('kpi_user');
  if(t&&u){try{State.token=t;State.user=JSON.parse(u);onLoginSuccess();}catch (e) {sessionStorage.clear();}}
}

// ═══ API — with timeout & retry ═══════════════════════════
async function api(action, data={}, retries=1) {
  if (USE_MOCK) return mockApi(action, data);

  const attemptFetch = async () => {
    const ctrl = new AbortController();
    const timer = setTimeout(()=>ctrl.abort(), API_TIMEOUT_MS);
    try {
      const r = await fetch(GAS_URL, {
        method:'POST',
        headers:{'Content-Type':'text/plain;charset=utf-8'},
        body: JSON.stringify({action, data, token: State.token||''}),
        signal: ctrl.signal
      });
      clearTimeout(timer);
      const text = await r.text();
      if (!text) return {success:false, message:'GAS ส่ง response ว่าง'};
      try { return JSON.parse(text); }
      catch (e) { return {success:false, message:'Parse error: '+text.slice(0,200)}; }
    } catch(e) {
      clearTimeout(timer);
      if (e.name==='AbortError') throw new Error('Request timeout ('+API_TIMEOUT_MS/1000+'s)');
      throw e;
    }
  };

  for (let i=0; i<=retries; i++) {
    try {
      const j = await attemptFetch();
      if (j.code===401){showToast('หมดเวลาการใช้งาน กรุณา Login ใหม่','warning');handleLogout();}
      return j;
    } catch(e) {
      if (i===retries) {
        const msg = e.message.includes('timeout')
          ? `GAS ตอบช้าเกิน ${API_TIMEOUT_MS/1000} วินาที — กรุณากด Refresh`
          : 'เชื่อมต่อ GAS ไม่ได้: '+e.message;
        console.error('[API]',action,e);
        return {success:false, message:msg};
      }
      await new Promise(r=>setTimeout(r,1000)); // รอ 1s แล้ว retry
    }
  }
}

// ═══ MOCK DATA ════════════════════════════════════════════
function mockApi(action,data){
  const U=[
    {Group_ID:'g01',Group_Name:'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม',Unit_ID:'13',Unit_Name:'หอผู้ป่วย สก.3'},
    {Group_ID:'g02',Group_Name:'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม',Unit_ID:'34',Unit_Name:'หอผู้ป่วยเคมีบำบัด'},
    {Group_ID:'g07',Group_Name:'กลุ่มงานการพยาบาลผู้ป่วยหนัก',Unit_ID:'20',Unit_Name:'หอผู้ป่วยหนัก SICU'},
    {Group_ID:'g08',Group_Name:'กลุ่มงานการพยาบาลผู้ป่วยนอก',Unit_ID:'g081',Unit_Name:'OPD อายุรกรรม'},
  ];
  const G=[
    {Group_ID:'g02',Group_Name:'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม'},
    {Group_ID:'g07',Group_Name:'กลุ่มงานการพยาบาลผู้ป่วยหนัก'},
  ];
  const K=[
    {KPI_ID:'KPI001',KPI_Name:'อัตราการพลัดตกหกล้มในหอผู้ป่วย',KPI_Category:'ความปลอดภัยผู้ป่วย',Unit_ID:'34',Target:'<1',Calc_Type:'x1000',Status:'Active'},
    {KPI_ID:'KPI002',KPI_Name:'อัตราการติดเชื้อทางหลอดเลือดดำส่วนกลาง (CLABSI)',KPI_Category:'ความปลอดภัยผู้ป่วย',Unit_ID:'34',Target:'<1',Calc_Type:'x1000',Status:'Active'},
    {KPI_ID:'KPI003',KPI_Name:'อัตราความพึงพอใจของผู้รับบริการ',KPI_Category:'ความพึงพอใจ',Unit_ID:'34',Target:'≥80',Calc_Type:'x100',Status:'Active'},
    {KPI_ID:'KPI004',KPI_Name:'จำนวนครั้งที่เกิดความคลาดเคลื่อนทางยา',KPI_Category:'ความปลอดภัยผู้ป่วย',Unit_ID:'34',Target:'=0',Calc_Type:'จำนวนครั้ง',Status:'Active'},
    {KPI_ID:'KPI005',KPI_Name:'อัตราการบันทึกทางการพยาบาลถูกต้องครบถ้วน',KPI_Category:'คุณภาพการพยาบาล',Unit_ID:'34',Target:'≥90',Calc_Type:'x100',Status:'Active'},
  ];
  const cy=new Date().getFullYear(),be=cy+543;
  const D=[];
  TH_MONTHS.forEach((m,mi)=>{
    K.forEach(k=>{
      const v=k.Calc_Type==='x100'?+(72+Math.random()*22).toFixed(1):k.Calc_Type==='x1000'?+(Math.random()*1.2).toFixed(3):Math.floor(Math.random()*2);
      const n=k.Calc_Type!=='จำนวนครั้ง'?Math.floor(Math.random()*50+10):'';
      const dn=k.Calc_Type!=='จำนวนครั้ง'?Math.floor(Math.random()*80+40):'';
      D.push({Record_ID:`R${mi}_${k.KPI_ID}`,Month_Year:`${m}${be}`,Unit_ID:'34',KPI_ID:k.KPI_ID,Result_Value:String(v),Numerator:String(n),Denominator:String(dn),Recorded_By:'พว.ตัวอย่าง',Last_Updated:new Date(cy,mi,15).toISOString()});
    });
  });
  return new Promise(r=>setTimeout(()=>{
    if(action==='login'){
      const map={
        admin:{r:'Admin',u:''},nso:{r:'NSO',u:''},
        g02:{r:'Head',u:'g02'},
        spec34:{r:'Spec',u:'34'},
        '34':{r:'User',u:'34'},'22':{r:'User',u:'22'}
      };
      const pw={admin:'admin1234',nso:'nso1234',g02:'head1234',spec34:'spec1234','34':'Password1234','22':'Password1234'};
      const un=data.username;
      if(map[un]&&pw[un]===data.password){
        const unit=U.find(x=>x.Unit_ID===map[un].u);
        r({success:true,token:btoa(JSON.stringify({userId:'USR0X',role:map[un].r,unitId:map[un].u,exp:Date.now()+8*3600*1000})),
          user:{userId:'USR0X',username:un,role:map[un].r,unitId:map[un].u,unitName:unit?.Unit_Name||'',groupId:unit?.Group_ID||'',groupName:unit?.Group_Name||''}});
      }else r({success:false,message:'Demo: admin/admin1234, 34/Password1234'});
    }
    else if(action==='getUnits'){const g={};U.forEach(u=>{if(!g[u.Group_ID])g[u.Group_ID]={groupId:u.Group_ID,groupName:u.Group_Name,units:[]};g[u.Group_ID].units.push({unitId:u.Unit_ID,unitName:u.Unit_Name});});r({success:true,units:U,groups:Object.values(g)});}
    else if(action==='getKPIMaster'){let k=K;if(data.unitId)k=k.filter(x=>x.Unit_ID===data.unitId||String(x.Unit_ID).toLowerCase().startsWith('common'));r({success:true,kpis:k});}
    else if(action==='getKPIData'){let d=[...D];if(data.unitId)d=d.filter(x=>x.Unit_ID===data.unitId);if(data.monthYear)d=d.filter(x=>x.Month_Year===data.monthYear);r({success:true,records:d});}
    else if(action==='getDashboardData'){
      const yrCE=parseInt(data.year)||cy; const be2=yrCE+543; const months=TH_MONTHS.map(m=>`${m}${be2}`);
      let recs=D.filter(x=>x.Month_Year.endsWith(String(be2)));
      if(data.unitId)recs=recs.filter(x=>x.Unit_ID===data.unitId);
      if(data.monthYear)recs=recs.filter(x=>x.Month_Year===data.monthYear);
      const kpis=data.unitId?K.filter(k=>k.Unit_ID===data.unitId):K;
      const summary=kpis.map(k=>{const rec=data.monthYear?recs.find(x=>x.KPI_ID===k.KPI_ID):recs.filter(x=>x.KPI_ID===k.KPI_ID).sort((a,b)=>b.Month_Year>a.Month_Year?1:-1)[0];return{kpiId:k.KPI_ID,kpiName:k.KPI_Name,category:k.KPI_Category,target:k.Target,calcType:k.Calc_Type,latestValue:rec?rec.Result_Value:null,latestMonth:rec?rec.Month_Year:null};});
      const trend=kpis.slice(0,4).map(k=>({kpiId:k.KPI_ID,kpiName:k.KPI_Name,target:k.Target,calcType:k.Calc_Type,data:months.map(m=>{const x=D.find(d=>d.KPI_ID===k.KPI_ID&&d.Month_Year===m);return{month:m,value:x?parseFloat(x.Result_Value):null};})}));
      r({success:true,summary,trend,months});
    }
    else if(action==='getUsers')r({success:true,users:[{User_ID:'USR001',Username:'admin',Role:'Admin',Unit_ID:'',Created_Date:'2025-01-01'}]});
    else r({success:true,message:'OK (Demo)'});
  },300));
}

// ═══ AUTH ═════════════════════════════════════════════════
async function handleLogin(){
  const un=document.getElementById('login-username').value.trim();
  const pw=document.getElementById('login-password').value;
  const err=document.getElementById('login-error'),btn=document.getElementById('login-btn');
  if(!un||!pw){err.textContent='กรุณากรอก Username และรหัสผ่าน';err.classList.remove('d-none');return;}
  err.classList.add('d-none');
  btn.innerHTML='<span class="spinner"></span> กำลังเข้าสู่ระบบ...'; btn.disabled=true;
  const res=await api('login',{username:un,password:pw});
  btn.innerHTML='<i class="fa-solid fa-right-to-bracket"></i> เข้าสู่ระบบ'; btn.disabled=false;
  if(res.success){
    State.token=res.token; State.user=res.user;
    sessionStorage.setItem('kpi_token',res.token);
    sessionStorage.setItem('kpi_user',JSON.stringify(res.user));
    onLoginSuccess();
  }else{err.textContent=res.message||'เข้าสู่ระบบไม่สำเร็จ';err.classList.remove('d-none');}
}
async function onLoginSuccess(){
  const{role,username,unitName,groupName}=State.user;
  document.getElementById('login-page').style.display='none';
  document.getElementById('app-shell').classList.add('visible');
  document.getElementById('sidebar-username').textContent=username;
  document.getElementById('sidebar-role').textContent=role;
  ['nav-nso-section','nav-user-section','nav-admin-section'].forEach(id=>document.getElementById(id)?.classList.add('d-none'));
  const navMap = { Admin:'nav-admin-section', NSO:'nav-nso-section', Head:'nav-head-section', Spec:'nav-spec-section', User:'nav-user-section' };
  document.getElementById(navMap[role] || 'nav-user-section')?.classList.remove('d-none');
  const badge = unitName  ? `หน่วยงาน: ${unitName}`
              : groupName ? `กลุ่มงาน: ${groupName}`
              : role==='Admin' ? 'Admin (ทุกหน่วยงาน)'
              : role==='NSO'   ? 'ผู้บริหาร (ทุกหน่วยงาน)'
              : role==='Spec'  ? `เฉพาะ: ${unitName||'หน่วยงานของฉัน'}`
              : 'ผู้ใช้งาน';
  document.getElementById('topbar-unit-text').textContent=badge;
  document.getElementById('sidebar-unitname').textContent=badge;
  await loadMasterData();
  navigate((role==='User'||role==='Spec')?'myreport':'dashboard');
}
function handleLogout(){
  State.token=null; State.user=null; sessionStorage.clear();
  if(State.charts.trend){State.charts.trend.destroy();State.charts.trend=null;}
  if(State.charts.pie){State.charts.pie.destroy();State.charts.pie=null;}
  State.wardCharts.forEach(c=>{try{c.destroy();}catch (e) {}});State.wardCharts=[];
  document.getElementById('app-shell').classList.remove('visible');
  document.getElementById('login-page').style.display='flex';
  document.getElementById('login-password').value='';
  document.getElementById('login-error').classList.add('d-none');
  closeSidebar();
}
function togglePassword(){const p=document.getElementById('login-password'),e=document.getElementById('pw-eye');p.type=p.type==='password'?'text':'password';e.className=p.type==='password'?'fa-solid fa-eye':'fa-solid fa-eye-slash';}

// ═══ MASTER DATA ══════════════════════════════════════════
async function loadMasterData(){
  const[ur,kr]=await Promise.all([api('getUnits'),api('getKPIMaster')]);
  if(ur.success){State.units=ur.units||[];State.groups=ur.groups||[];populateGroupDropdowns();populateUnitDropdowns();}
  else console.error('[Master] getUnits:',ur.message);
  if(kr.success)State.kpiMaster=kr.kpis||[];
  else console.error('[Master] getKPIMaster:',kr.message);
}

// ═══ ROUTING ══════════════════════════════════════════════
function navigate(page){
  document.querySelectorAll('.page-view').forEach(p=>p.classList.add('d-none'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  document.querySelector(`.nav-item[onclick*="'${page}'"]`)?.classList.add('active');
  document.getElementById(`page-${page}`)?.classList.remove('d-none');
  State.currentPage=page;
  const T={dashboard:'Dashboard ภาพรวม KPI',entry:'บันทึก KPI รายเดือน',admin:'Admin Panel',myreport:'Dashboard หน่วยงาน',audit:'Audit Trail'};
  document.getElementById('topbar-page-title').textContent=T[page]||'';
  if(page==='dashboard')loadDashboard();
  else if(page==='admin')initAdminPanel();
  else if(page==='myreport')loadMyReport();
  else if(page==='entry')initEntryPage();
  else if(page==='audit')loadAuditLog();
  if(window.innerWidth<=992)closeSidebar();
}
function toggleSidebar(){document.getElementById('sidebar').classList.toggle('open');document.getElementById('sidebar-overlay').classList.toggle('open');}
function closeSidebar(){document.getElementById('sidebar').classList.remove('open');document.getElementById('sidebar-overlay').classList.remove('open');}

// ═══ YEAR/MONTH BUILDERS ══════════════════════════════════
function buildYearOptions(selId){
  const sel=document.getElementById(selId); if(!sel)return;
  const cur=sel.value;
  sel.innerHTML='';
  const cy=new Date().getFullYear();
  for(let be=MAX_BE;be>=MIN_BE;be--){
    const o=document.createElement('option');
    o.value=String(be-543); o.textContent=`พ.ศ. ${be}`;
    sel.appendChild(o);
  }
  sel.value=(cur&&sel.querySelector(`option[value="${cur}"]`))?cur:String(cy);
}
function buildMonthOptions(selId,includeAll=true){
  const sel=document.getElementById(selId); if(!sel)return;
  const cur=sel.value; sel.innerHTML='';
  if(includeAll){const o=document.createElement('option');o.value='';o.textContent='— ทั้งปี —';sel.appendChild(o);}
  TH_MONTHS_FULL.forEach((m,i)=>{const o=document.createElement('option');o.value=String(i);o.textContent=m;sel.appendChild(o);});
  if(cur!==null&&cur!==undefined&&sel.querySelector(`option[value="${cur}"]`))sel.value=cur;
}
function readPeriod(yearId,monthId){
  const yearCE=parseInt(document.getElementById(yearId)?.value||new Date().getFullYear());
  const mVal=document.getElementById(monthId)?.value??'';
  const be=yearCE+543;
  const mIdx=mVal!==''?parseInt(mVal):null;
  const monthYear=mIdx!==null?`${TH_MONTHS[mIdx]}${be}`:'';
  const label=mIdx!==null?`${TH_MONTHS_FULL[mIdx]} พ.ศ. ${be}`:`ทั้งปี พ.ศ. ${be}`;
  return{yearCE,be,mIdx,monthYear,label};
}

// ═══ CASCADING DROPDOWNS ══════════════════════════════════
function populateGroupDropdowns(){
  ['#dash-group-filter','#entry-group-select','#kpi-unit'].forEach(s=>{
    const el=document.querySelector(s);if(!el)return;
    const cur=el.value,first=el.options[0];el.innerHTML='';el.appendChild(first);
    State.groups.forEach(g=>{const o=document.createElement('option');o.value=g.groupId;o.textContent=`${g.groupId} — ${g.groupName}`;el.appendChild(o);});
    if(s==='#kpi-unit'){const d=document.createElement('option');d.disabled=true;d.textContent='── หน่วยงาน ──';el.appendChild(d);State.units.forEach(u=>{const o=document.createElement('option');o.value=u.Unit_ID;o.textContent=`${u.Unit_ID} — ${u.Unit_Name}`;el.appendChild(o);});}
    if(cur)el.value=cur;
  });
  const uu=document.getElementById('user-unit');
  if(uu){const f=uu.options[0];uu.innerHTML='';uu.appendChild(f);State.units.forEach(u=>{const o=document.createElement('option');o.value=u.Unit_ID;o.textContent=`${u.Unit_ID} — ${u.Unit_Name}`;uu.appendChild(o);});}
}
function populateUnitDropdowns(filtered){
  const units=filtered||State.units;
  ['#dash-unit-filter','#entry-unit-select'].forEach(s=>{
    const el=document.querySelector(s);if(!el)return;
    const f=el.options[0];el.innerHTML='';el.appendChild(f);
    units.forEach(u=>{const o=document.createElement('option');o.value=u.Unit_ID;o.textContent=`${u.Unit_ID} — ${u.Unit_Name}`;el.appendChild(o);});
  });
}
function onGroupFilterChange(ctx){
  const gId=document.getElementById(ctx==='dash'?'dash-group-filter':'entry-group-select').value;
  const filtered=gId?State.units.filter(u=>u.Group_ID===gId):State.units;
  const tgt=ctx==='dash'?'#dash-unit-filter':'#entry-unit-select';
  const el=document.querySelector(tgt);if(!el)return;
  const f=el.options[0];el.innerHTML='';el.appendChild(f);
  filtered.forEach(u=>{const o=document.createElement('option');o.value=u.Unit_ID;o.textContent=`${u.Unit_ID} — ${u.Unit_Name}`;el.appendChild(o);});
  if(ctx==='dash')loadDashboard();
}

// ═══════════════════════════════════════════════════════════
// DASHBOARD (Admin / NSO)
// ═══════════════════════════════════════════════════════════
async function loadDashboard(){
  const{role}=State.user;
  buildYearOptions('dash-year-filter');
  buildMonthOptions('dash-month-filter',true);

  const orgRow=document.getElementById('dash-org-filters');
  if(orgRow) orgRow.style.display=(role!=='User')?'flex':'none';

  const{yearCE,be,mIdx,monthYear,label}=readPeriod('dash-year-filter','dash-month-filter');
  const groupId=document.getElementById('dash-group-filter')?.value||'';
  const unitId=document.getElementById('dash-unit-filter')?.value||'';

  const sub=document.getElementById('dash-subtitle');
  if(sub)sub.textContent=`ข้อมูลประจำ ${label}`;
  updateDashChips(role,groupId,unitId,mIdx,be);

  const tbody=document.getElementById('dash-kpi-tbody');
  tbody.innerHTML=`<tr><td colspan="7" class="text-center py-5">
    <div class="spinner spinner-lg" style="margin:0 auto 12px"></div>
    <div style="color:var(--text-muted);font-size:.85rem">กำลังโหลดข้อมูล KPI...</div>
  </td></tr>`;
  ['stat-total','stat-pass','stat-fail','stat-pending'].forEach(id=>{
    const el=document.getElementById(id); if(el)el.innerHTML='<span class="spinner" style="width:14px;height:14px;border-width:2px"></span>';
  });

  const callUnitId = role==='User' ? State.user.unitId : unitId;
  const res = await api('getDashboardData',{
    unitId:callUnitId, groupId,
    year:String(yearCE), monthYear
  });

  if(!res.success){
    tbody.innerHTML=`<tr><td colspan="7">
      <div class="empty-state" style="padding:48px 20px">
        <div class="empty-state-icon">⚠️</div>
        <div class="empty-state-title">โหลดข้อมูลไม่สำเร็จ</div>
        <div class="empty-state-desc">${res.message}</div>
        <button class="btn btn-primary btn-sm mt-3" onclick="loadDashboard()">
          <i class="fa-solid fa-rotate"></i> ลองใหม่
        </button>
      </div>
    </td></tr>`;
    ['stat-total','stat-pass','stat-fail','stat-pending'].forEach(id=>{
      const el=document.getElementById(id); if(el)el.textContent='—';
    });
    return;
  }

  State.dashboardData=res;
  const{summary=[],trend=[],months=[]}=res;
  const isAgg = summary.length > 0 && summary[0].isAggregate === true;

  let pass, fail, pending;
  if (isAgg) {
    pass    = summary.reduce((s,k)=>s+(k.passCount||0),0);
    fail    = summary.reduce((s,k)=>s+(k.failCount||0),0);
    pending = summary.reduce((s,k)=>s+(k.pendingCount||0),0);
  } else {
    pass    = summary.filter(k=>k.latestValue!==null&&evaluateTarget(k.latestValue,k.target)).length;
    fail    = summary.filter(k=>k.latestValue!==null&&!evaluateTarget(k.latestValue,k.target)).length;
    pending = summary.filter(k=>k.latestValue===null).length;
  }

  document.getElementById('stat-total').textContent  = summary.length;
  document.getElementById('stat-pass').textContent   = pass;
  document.getElementById('stat-fail').textContent   = fail;
  document.getElementById('stat-pending').textContent= pending;

  const metaPass    = document.querySelector('#stat-pass ~ .kpi-stat-content .kpi-stat-meta, #stat-pass + .kpi-stat-content .kpi-stat-meta');
  const metaFail    = document.querySelector('#stat-fail ~ .kpi-stat-content .kpi-stat-meta');
  const metaPending = document.querySelector('#stat-pending ~ .kpi-stat-content .kpi-stat-meta');
  if (isAgg) {
    document.querySelectorAll('.kpi-stat-meta').forEach((el,i)=>{
      if(i===1) el.textContent='คู่หน่วยงาน-KPI ที่ผ่าน';
      if(i===2) el.textContent='คู่หน่วยงาน-KPI ที่ไม่ผ่าน';
      if(i===3) el.textContent='คู่หน่วยงาน-KPI ยังไม่บันทึก';
    });
  } else {
    document.querySelectorAll('.kpi-stat-meta').forEach((el,i)=>{
      if(i===1) el.textContent='KPI ที่บรรลุเป้าหมาย';
      if(i===2) el.textContent='ต้องพัฒนาปรับปรุง';
      if(i===3) el.textContent='รอการบันทึกข้อมูล';
    });
  }

  const thead = document.querySelector('#dash-kpi-table thead tr');
  if (thead) {
    if (isAgg) {
      thead.innerHTML=`<th>#</th><th>ตัวชี้วัด (KPI)</th><th>หมวดหมู่</th><th>เป้าหมาย</th><th style="text-align:center">✅ ผ่าน</th><th style="text-align:center">❌ ไม่ผ่าน</th><th style="text-align:center">⏳ ว่าง</th><th style="text-align:center">อัตราผ่าน</th>`;
    } else {
      thead.innerHTML=`<th>#</th><th>ตัวชี้วัด (KPI)</th><th>หมวดหมู่</th><th>เป้าหมาย</th><th>ผลลัพธ์ล่าสุด</th><th>เดือน</th><th>สถานะ</th>`;
    }
  }

  if (isAgg) {
    const sub=document.getElementById('dash-subtitle');
    if(sub) sub.innerHTML=`ข้อมูลประจำ ${label} &nbsp;<span style="font-size:.75rem;background:#e0f2fe;color:#0369a1;padding:2px 8px;border-radius:99px;font-weight:700">📊 โหมดภาพรวมองค์กร</span>`;
  }

  if(!summary.length){
    const colspan = isAgg ? '8' : '7';
    tbody.innerHTML=`<tr><td colspan="${colspan}"><div class="empty-state"><div class="empty-state-icon">📊</div><div class="empty-state-title">ไม่มีข้อมูล KPI ในช่วงเวลานี้</div></div></td></tr>`;
  } else if (isAgg) {
    tbody.innerHTML=summary.map((k,i)=>{
      const col = k.passRate>=80?'var(--success)':k.passRate>=50?'#f59e0b':'var(--danger)';
      const bar = `<div style="display:inline-block;width:60px;height:6px;background:#e5e7eb;border-radius:3px;vertical-align:middle;margin-left:6px"><div style="width:${k.passRate}%;height:100%;background:${col};border-radius:3px"></div></div>`;
      return `<tr>
        <td style="color:var(--text-muted);font-weight:600">${i+1}</td>
        <td style="font-weight:600;max-width:280px">${k.kpiName}</td>
        <td><span style="font-size:.78rem;background:var(--bg-main);padding:2px 7px;border-radius:5px">${k.category||'—'}</span></td>
        <td class="target-display">${k.target||'—'}</td>
        <td style="text-align:center;font-weight:700;color:var(--success)">${k.passCount}</td>
        <td style="text-align:center;font-weight:700;color:var(--danger)">${k.failCount}</td>
        <td style="text-align:center;color:var(--text-muted)">${k.pendingCount}</td>
        <td style="text-align:center">
          <span style="font-weight:800;color:${col}">${k.passRate}%</span>
          ${bar}
          <div style="font-size:.72rem;color:var(--text-muted);margin-top:2px">${k.reportedCount}/${k.totalUnits} หน่วยรายงาน</div>
        </td>
      </tr>`;
    }).join('');
  } else {
    tbody.innerHTML=summary.map((k,i)=>{
      const hv=k.latestValue!==null&&k.latestValue!=='';
      const ok=hv&&evaluateTarget(k.latestValue,k.target);
      const badge=hv?`<span class="status-badge ${ok?'pass':'fail'}">${ok?'✅ ผ่าน':'❌ ไม่ผ่าน'}</span>`:`<span class="status-badge pending">⏳ ยังไม่บันทึก</span>`;
      return `<tr>
        <td style="color:var(--text-muted);font-weight:600">${i+1}</td>
        <td style="font-weight:600;max-width:280px">${k.kpiName}</td>
        <td><span style="font-size:.78rem;background:var(--bg-main);padding:2px 7px;border-radius:5px">${k.category||'—'}</span></td>
        <td class="target-display">${k.target||'—'}</td>
        <td style="font-weight:700;font-size:1rem;color:${ok?'var(--success)':hv?'var(--danger)':'var(--text-muted)'}">${hv?formatValue(k.latestValue,k.calcType):'—'}</td>
        <td style="font-size:.82rem;color:var(--text-muted)">${k.latestMonth||'—'}</td>
        <td>${badge}</td>
      </tr>`;
    }).join('');
  }

  renderTrendChart(trend,months);
  renderPieChart(pass,fail,pending);
  renderProgressBars(summary);

  const ts=document.getElementById('trend-kpi-select');
  if(ts)ts.innerHTML=trend.map(t=>`<option value="${t.kpiId}">${t.kpiName.slice(0,40)}${t.kpiName.length>40?'...':''}</option>`).join('');

  const bSel=document.getElementById('dash-bar-kpi-select'),lSel=document.getElementById('dash-line-kpi-select');
  if(bSel)bSel.innerHTML=''; if(lSel)lSel.innerHTML='';
  if(dashBarChart){dashBarChart.destroy();dashBarChart=null;}
  if(dashLineChart){dashLineChart.destroy();dashLineChart=null;}
  if(dashView!=='table')setTimeout(()=>switchDashView(dashView),50);
}

// ═══ CHARTS (Dashboard) ═══════════════════════════════════
function ttOpts(){return{backgroundColor:'white',titleColor:'#1a2e2c',bodyColor:'#4a6a67',borderColor:'#d1ede9',borderWidth:1,padding:12,titleFont:{family:'Sarabun',weight:'700',size:13},bodyFont:{family:'Sarabun',size:12}};}
function renderTrendChart(trend,months){
  const ctx=document.getElementById('trend-chart');if(!ctx)return;
  if(State.charts.trend){State.charts.trend.destroy();State.charts.trend=null;}
  if(!trend?.length)return;
  const kpi=trend[0],target=parseFloat((kpi.target||'').replace(/[^0-9.]/g,''));
  State.charts.trend=new Chart(ctx,{type:'line',data:{labels:months,datasets:[
    {label:kpi.kpiName,data:kpi.data.map(d=>d.value),borderColor:'#359286',backgroundColor:'rgba(53,146,134,0.08)',borderWidth:2.5,pointBackgroundColor:'#359286',pointRadius:4,pointHoverRadius:6,tension:0.4,fill:true,spanGaps:true},
    ...(!isNaN(target)?[{label:'เป้าหมาย',data:months.map(()=>target),borderColor:'#ef4444',borderWidth:2,borderDash:[6,3],pointRadius:0,fill:false}]:[])
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{intersect:false,mode:'index'},plugins:{legend:{display:true,position:'top',labels:{font:{family:'Sarabun',size:12},usePointStyle:true,boxWidth:10}},tooltip:ttOpts()},scales:{x:{grid:{color:'#f0faf9'},ticks:{font:{family:'Sarabun',size:11},color:'#4a6a67',maxRotation:45}},y:{grid:{color:'#f0faf9'},ticks:{font:{family:'Sarabun',size:11},color:'#4a6a67'},beginAtZero:true}}}});
}
function renderPieChart(pass,fail,pending){
  const ctx=document.getElementById('pie-chart');if(!ctx)return;
  if(State.charts.pie){State.charts.pie.destroy();State.charts.pie=null;}
  State.charts.pie=new Chart(ctx,{type:'doughnut',data:{labels:['ผ่านเกณฑ์','ไม่ผ่านเกณฑ์','ยังไม่บันทึก'],datasets:[{data:[pass,fail,pending],backgroundColor:['#10b981','#ef4444','#f59e0b'],borderColor:['#fff','#fff','#fff'],borderWidth:3,hoverOffset:8}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{font:{family:'Sarabun',size:12},usePointStyle:true,padding:16}},tooltip:ttOpts()},cutout:'65%'}});
}
function updateTrendChart(){
  const id=document.getElementById('trend-kpi-select')?.value;
  if(!State.dashboardData||!id||!State.charts.trend)return;
  const kpi=State.dashboardData.trend.find(t=>t.kpiId===id);if(!kpi)return;
  State.charts.trend.data.datasets[0].label=kpi.kpiName;
  State.charts.trend.data.datasets[0].data=kpi.data.map(d=>d.value);
  const tgt=parseFloat((kpi.target||'').replace(/[^0-9.]/g,''));
  if(!isNaN(tgt)&&State.charts.trend.data.datasets[1])State.charts.trend.data.datasets[1].data=State.dashboardData.months.map(()=>tgt);
  State.charts.trend.update();
}

// ═══ WARD DASHBOARD ═══════════════════════════════════════
async function loadMyReport(){
  const{unitId,unitName,username}=State.user;
  const root=document.getElementById('ward-dashboard-root');
  const hist=document.getElementById('ward-history-section');

  buildYearOptions('ward-year-sel');
  buildMonthOptions('ward-month-sel',true);

  const{yearCE,be,mIdx,monthYear,label}=readPeriod('ward-year-sel','ward-month-sel');
  const labelEl=document.getElementById('ward-time-label');
  if(labelEl){
    labelEl.style.display='inline-flex';
    labelEl.innerHTML=`<i class="fa-solid fa-clock" style="font-size:.75rem"></i> ${label} <span class="chip-close" onclick="resetWardTime()" title="รีเซ็ต">✕</span>`;
  }

  if(root)root.innerHTML=wardSkeleton();
  if(hist)hist.innerHTML=`<div class="ward-chart-card"><div class="ward-chart-card-body" style="padding:24px;text-align:center;color:var(--text-muted);font-size:.88rem"><div class="spinner" style="margin:0 auto 10px"></div>กำลังโหลดประวัติการบันทึก...</div></div>`;
  State.wardCharts.forEach(c=>{try{c.destroy();}catch (e) {}});State.wardCharts=[];

  console.log('[Ward] unitId:', unitId, 'yearCE:', yearCE, 'be:', be, 'monthYear:', monthYear||'ทั้งปี');
  const dashRes=await api('getDashboardData',{unitId,year:String(yearCE),monthYear});
  console.log('[Ward] dashRes:', dashRes.success, '| summary:', dashRes.summary?.length, '| debug:', JSON.stringify(dashRes._debug||{}));
  buildWardHero(root,dashRes,unitName,username,label,mIdx);

  if(hist){
    try{
      const histRes=await api('getKPIData',{unitId});
      if(histRes.success)buildWardHistory(hist,histRes.records||[]);
      else hist.innerHTML=`<div class="empty-state" style="padding:32px"><div class="empty-state-icon">⚠️</div><div class="empty-state-title">${histRes.message}</div></div>`;
    }catch(e){
      hist.innerHTML=`<div class="empty-state" style="padding:32px"><div class="empty-state-icon">⚠️</div><div class="empty-state-title">โหลดประวัติไม่สำเร็จ</div></div>`;
    }
  }
}
function onWardTimeChange(){loadMyReport();}
function resetWardTime(){
  const y=document.getElementById('ward-year-sel'),m=document.getElementById('ward-month-sel');
  if(y)y.value=String(new Date().getFullYear());
  if(m)m.value='';
  loadMyReport();
}

function buildWardHero(root,dashRes,unitName,username,label,mIdx){
  if(!root)return;
  if(!dashRes.success){
    const errMsg = dashRes.message || 'ไม่ทราบสาเหตุ';
    root.innerHTML=`<div class="empty-state" style="padding:60px 20px">
      <div class="empty-state-icon">⚠️</div>
      <div class="empty-state-title">โหลดข้อมูลไม่สำเร็จ</div>
      <div class="empty-state-desc" style="max-width:400px;word-break:break-word">${errMsg}</div>
      <button class="btn btn-primary btn-sm" style="margin-top:16px" onclick="loadMyReport()">
        <i class="fa-solid fa-rotate"></i> ลองใหม่
      </button>
    </div>`;
    console.error('[Ward Hero] API failed:', errMsg);
    return;
  }
  if(dashRes._debug) console.table(dashRes._debug);
  const{summary=[],trend=[],months=[]}=dashRes;
  const wv=summary.filter(k=>k.latestValue!==null&&k.latestValue!=='');
  const pc=wv.filter(k=>evaluateTarget(k.latestValue,k.target)).length;
  const fc=wv.length-pc,pdc=summary.length-wv.length;
  const pr=summary.length>0?Math.round((pc/summary.length)*100):0;
  const q=WARD_QUOTES[new Date().getDate()%WARD_QUOTES.length];
  const a=pr>=80?{cls:'excellent',e:'🏆',t:'ยอดเยี่ยม! หน่วยงานของท่านทำได้ดีมาก',s:`ผ่านเกณฑ์ ${pc}/${summary.length} — ระดับดีเลิศ`}
          :pr>=50?{cls:'good',e:'⭐',t:'ดี! กำลังพัฒนาไปในทิศทางที่ถูกต้อง',s:`ผ่านเกณฑ์ ${pc}/${summary.length}`}
          :{cls:'improve',e:'💪',t:'ร่วมกันพัฒนาต่อไป',s:`ผ่านเกณฑ์ ${pc}/${summary.length} — ทุกก้าวคือความสำเร็จ`};
  root.innerHTML=`
    <div class="ward-hero" style="animation:fadeSlideIn .5s ease both">
      <div class="ward-hero-blobs"><div class="ward-blob ward-blob-1"></div><div class="ward-blob ward-blob-2"></div><div class="ward-blob ward-blob-3"></div></div>
      <div class="ward-hero-content">
        <div class="ward-hero-greeting">🌿 สวัสดี คุณ${username} — ${new Date().toLocaleDateString('th-TH',{day:'numeric',month:'long',year:'numeric'})}</div>
        <div class="ward-hero-name">${unitName||'หน่วยงานของฉัน'} <span>✦</span></div>
        <div class="ward-hero-sub"><span>${label}</span><span class="dot"></span><span>${summary.length} ตัวชี้วัด</span></div>
        <div class="ward-hero-stats">
          <div class="ward-hero-stat"><div class="ward-hero-stat-val mint">${pr}%</div><div class="ward-hero-stat-label">อัตราผ่านเกณฑ์</div></div>
          <div class="ward-hero-stat"><div class="ward-hero-stat-val gold">${pc}</div><div class="ward-hero-stat-label">ผ่านเกณฑ์</div></div>
          <div class="ward-hero-stat"><div class="ward-hero-stat-val coral">${fc}</div><div class="ward-hero-stat-label">ต้องพัฒนา</div></div>
          <div class="ward-hero-stat"><div class="ward-hero-stat-val" style="color:rgba(255,255,255,.75)">${pdc}</div><div class="ward-hero-stat-label">ยังไม่บันทึก</div></div>
        </div>
      </div>
    </div>
    <div class="ward-quote-bar"><div class="ward-quote-icon">${q.e}</div><div class="ward-quote-text">${q.t}</div></div>
    <div class="ward-achievement ${a.cls}"><div class="ward-achievement-emoji">${a.e}</div><div class="ward-achievement-text"><div class="ward-achievement-title">${a.t}</div><div class="ward-achievement-sub">${a.s}</div></div></div>
    <div class="ward-section-title">🎯 ตัวชี้วัดคุณภาพ — ${label}</div>
    <div class="ward-kpi-grid" id="ward-kpi-grid">${kpiRingCards(summary)}</div>
    ${mIdx===null?`
      <div class="ward-section-title">📈 กราฟแนวโน้มรายเดือน</div>
      <div id="ward-trend-charts">${trendCardHTML(trend)}</div>
      <div class="ward-section-title">🗓️ ภาพรวมการบันทึก 12 เดือน</div>
      <div class="ward-chart-card" style="animation:fadeSlideIn .5s ease .5s both">
        <div class="ward-chart-card-header"><div class="ward-chart-card-title">📊 สถานะ KPI ทั้ง 12 เดือน</div><div style="font-size:.74rem;color:#9ca3af">🟢ผ่าน 🔴ไม่ผ่าน ⬜ไม่มีข้อมูล</div></div>
        <div class="ward-chart-card-body">${heatmapHTML(trend,months)}</div>
      </div>`:''}
    <div class="ward-spirit-card"><div class="ward-spirit-avatar">🤍</div><div class="ward-spirit-text"><div class="ward-spirit-title">ขอบคุณทุกการดูแลที่ท่านมอบให้</div><div class="ward-spirit-sub">พัฒนาโดย <strong>พว.วิษณุกรณ์ โอยา</strong> | กลุ่มงานวิจัยและพัฒนาการพยาบาล โรงพยาบาลชลบุรี</div></div></div>`;
  setTimeout(()=>animateRings(summary),80);
  if(mIdx===null)setTimeout(()=>drawWardCharts(trend,months),150);
  const toggle=document.getElementById('ward-view-toggle');
  if(toggle)toggle.style.display='';
  if(dashRes.success)State.dashboardData=dashRes;
  wardView='table';
  ['table','bar','line'].forEach(v=>{const btn=document.getElementById(`ward-view-btn-${v}`);if(btn)btn.classList.toggle('active',v==='table');});
  if(wardBarChart){wardBarChart.destroy();wardBarChart=null;}
  if(wardLineChart){wardLineChart.destroy();wardLineChart=null;}
}

// ═══ WARD HISTORY ══════════════════════════════════════════
function buildWardHistory(container,allRecords){
  const km={};State.kpiMaster.forEach(k=>{km[k.KPI_ID]=k;});
  const byMonth={};
  allRecords.forEach(r=>{const key=r.Month_Year||'ไม่ระบุ';if(!byMonth[key])byMonth[key]=[];byMonth[key].push(r);});
  const sorted=Object.keys(byMonth).sort((a,b)=>{
    const p=s=>{const mi=TH_MONTHS.findIndex(m=>s.startsWith(m));const y=parseInt(s.replace(/[^0-9]/g,''))||0;return y*100+(mi>=0?mi:0);};
    return p(b)-p(a);
  });
  if(!sorted.length){
    container.innerHTML=`<div class="ward-section-title">📋 ประวัติการบันทึก KPI</div><div class="ward-chart-card"><div class="ward-chart-card-body"><div class="empty-state" style="padding:40px 0"><div class="empty-state-icon">📭</div><div class="empty-state-title">ยังไม่มีประวัติการบันทึก</div></div></div></div>`;
    return;
  }
  let html=`<div class="ward-section-title" style="margin-top:8px">📋 ประวัติการบันทึก KPI ทั้งหมด (${sorted.length} เดือน)</div><div style="display:flex;flex-direction:column;gap:14px">`;
  sorted.forEach((my,idx)=>{
    const recs=byMonth[my];
    const pr=recs.filter(r=>{const k=km[r.KPI_ID];return k&&r.Result_Value!==''&&r.Result_Value!==null&&evaluateTarget(r.Result_Value,k.Target);}).length;
    const rate=recs.length>0?Math.round((pr/recs.length)*100):0;
    const col=rate>=80?'#10b981':rate>=50?'#f59e0b':'#ef4444';
    const sid=`hist_${my.replace(/[^a-zA-Z0-9]/g,'_')}`;
    const open=idx===0;
    html+=`<div class="ward-chart-card" style="animation:fadeSlideIn .35s ease ${idx*.05}s both;overflow:hidden">
      <div class="ward-chart-card-header" style="cursor:pointer;user-select:none" onclick="toggleHist('${sid}')">
        <div class="ward-chart-card-title"><div class="ward-chart-dot" style="background:${col}"></div>📅 ${my}<span style="font-size:.75rem;font-weight:400;color:#9ca3af;margin-left:6px">${recs.length} ตัวชี้วัด</span></div>
        <div style="display:flex;align-items:center;gap:12px">
          <div style="text-align:right"><div style="font-size:.72rem;color:#9ca3af">ผ่านเกณฑ์</div><div style="font-size:1.1rem;font-weight:800;color:${col}">${pr}/${recs.length}</div></div>
          <div style="width:60px;background:#f3f4f6;border-radius:99px;height:7px;overflow:hidden"><div style="width:${rate}%;height:100%;background:${col};border-radius:99px"></div></div>
          <i class="fa-solid fa-chevron-down" id="chev_${sid}" style="color:#9ca3af;font-size:.8rem;transition:transform .25s;${open?'transform:rotate(180deg)':''}"></i>
        </div>
      </div>
      <div id="${sid}" style="${open?'':'display:none;'}overflow-x:auto">
        <table class="data-table">
          <thead><tr><th>#</th><th>ตัวชี้วัด KPI</th><th>หมวดหมู่</th><th>เป้าหมาย</th><th>ตัวเศษ</th><th>ตัวส่วน</th><th>ผลลัพธ์</th><th>สูตร</th><th>สถานะ</th><th>บันทึกโดย</th><th>อัปเดต</th></tr></thead>
          <tbody>${recs.map((r,i)=>{
            const k=km[r.KPI_ID]||{};
            const hv=r.Result_Value!==''&&r.Result_Value!==null&&r.Result_Value!==undefined;
            const ok=hv&&evaluateTarget(r.Result_Value,k.Target);
            const badge=hv?`<span class="status-badge ${ok?'pass':'fail'}">${ok?'✅ ผ่าน':'❌ ไม่ผ่าน'}</span>`:`<span class="status-badge na">—</span>`;
            const upd=r.Last_Updated?new Date(r.Last_Updated).toLocaleDateString('th-TH',{day:'numeric',month:'short',year:'2-digit',hour:'2-digit',minute:'2-digit'}):'—';
            return `<tr>
              <td style="color:var(--text-muted);font-size:.78rem">${i+1}</td>
              <td style="font-weight:600;max-width:220px;font-size:.85rem">${k.KPI_Name||r.KPI_ID}</td>
              <td><span style="font-size:.72rem;background:var(--bg-main);padding:2px 7px;border-radius:5px">${k.KPI_Category||'—'}</span></td>
              <td class="target-display">${k.Target||'—'}</td>
              <td style="font-size:.85rem">${r.Numerator||'—'}</td>
              <td style="font-size:.85rem">${r.Denominator||'—'}</td>
              <td style="font-weight:700;font-size:.95rem;color:${ok?'var(--success)':hv?'var(--danger)':'var(--text-muted)'}">${hv?formatValue(r.Result_Value,k.Calc_Type):'—'}</td>
              <td><span style="font-size:.72rem;background:var(--primary-xlight);color:var(--primary);padding:2px 7px;border-radius:5px;font-weight:700">${k.Calc_Type||'—'}</span></td>
              <td>${badge}</td>
              <td style="font-size:.78rem;color:var(--text-muted)">${r.Recorded_By||'—'}</td>
              <td style="font-size:.72rem;color:var(--text-muted);white-space:nowrap">${upd}</td>
            </tr>`;
          }).join('')}</tbody>
        </table>
      </div>
    </div>`;
  });
  html+='</div>';
  container.innerHTML=html;
}
function toggleHist(id){const el=document.getElementById(id),chv=document.getElementById(`chev_${id}`);if(!el)return;const open=el.style.display!=='none';el.style.display=open?'none':'block';if(chv)chv.style.transform=open?'':'rotate(180deg)';}

// ═══ KPI RINGS ════════════════════════════════════════════
function kpiRingCards(summary){
  if(!summary.length)return'<div class="empty-state"><div class="empty-state-icon">📋</div><div class="empty-state-title">ยังไม่มีตัวชี้วัด</div></div>';
  const delays=['ward-animate-1','ward-animate-2','ward-animate-3','ward-animate-4','ward-animate-5','ward-animate-6'];
  return summary.map((k,i)=>{
    const hv=k.latestValue!==null&&k.latestValue!=='';
    const ok=hv&&evaluateTarget(k.latestValue,k.target);
    const cls=hv?(ok?'pass':'fail'):'empty';
    const v=parseFloat(k.latestValue),t=parseFloat((k.target||'').replace(/[^0-9.]/g,''));
    const pct=hv?((!isNaN(v)&&!isNaN(t)&&t>0)?Math.min(100,(v/t)*100):ok?100:30):0;
    const unit=k.calcType==='x100'?'%':k.calcType==='x1000'?'‰':'ครั้ง';
    return `<div class="ward-kpi-ring-card ${cls} fadeSlideIn ${delays[i%6]}">
      <div class="ward-ring-wrap">
        <svg class="ward-ring-svg" viewBox="0 0 90 90">
          <circle class="ward-ring-bg" cx="45" cy="45" r="35"/>
          <circle class="ward-ring-progress ${cls}" cx="45" cy="45" r="35" stroke-dasharray="220" stroke-dashoffset="220" id="rp_${i}"/>
        </svg>
        <div class="ward-ring-center">
          <div class="ward-ring-val ${cls}">${hv?formatValue(k.latestValue,k.calcType):'—'}</div>
          ${hv?`<div class="ward-ring-unit">${unit}</div>`:''}
        </div>
      </div>
      <div class="ward-kpi-name">${k.kpiName}</div>
      <div class="ward-kpi-target">🎯 เป้า: ${k.target||'—'}</div>
      <div class="ward-kpi-status-pill ${cls}">${cls==='pass'?'✅ ผ่านเกณฑ์':cls==='fail'?'❌ ต้องพัฒนา':'○ ยังไม่บันทึก'}</div>
    </div>`;
  }).join('');
}
function animateRings(summary){
  summary.forEach((k,i)=>{
    if(k.latestValue===null||k.latestValue==='')return;
    const el=document.getElementById(`rp_${i}`);if(!el)return;
    const v=parseFloat(k.latestValue),t=parseFloat((k.target||'').replace(/[^0-9.]/g,''));
    const pct=(!isNaN(v)&&!isNaN(t)&&t>0)?Math.min(100,(v/t)*100):evaluateTarget(k.latestValue,k.target)?100:30;
    setTimeout(()=>{el.style.strokeDashoffset=220-(220*pct/100);},60+i*55);
  });
}

// ═══ WARD TREND CHARTS ════════════════════════════════════
function trendCardHTML(trend){
  if(!trend.length)return'<div class="empty-state"><div class="empty-state-icon">📈</div><div class="empty-state-title">ยังไม่มีข้อมูลแนวโน้ม</div></div>';
  return`<div class="row g-3">${trend.map((kpi,i)=>{const c=CHART_COLORS[i%CHART_COLORS.length];return`<div class="col-12 col-lg-6"><div class="ward-chart-card" style="animation:fadeSlideIn .5s ease ${.1+i*.08}s both"><div class="ward-chart-card-header"><div class="ward-chart-card-title"><div class="ward-chart-dot" style="background:${c.s}"></div>${kpi.kpiName.length>45?kpi.kpiName.slice(0,45)+'...':kpi.kpiName}</div><div style="font-size:.75rem;color:#9ca3af">เป้า: ${kpi.target}</div></div><div class="ward-chart-card-body" style="padding-top:8px"><div style="height:180px;position:relative"><canvas id="wc_${i}"></canvas></div></div></div></div>`;}).join('')}</div>`;
}
function drawWardCharts(trend,months){
  State.wardCharts.forEach(c=>{try{c.destroy();}catch (e) {}});State.wardCharts=[];
  trend.forEach((kpi,i)=>{
    const ctx=document.getElementById(`wc_${i}`);if(!ctx)return;
    const c=CHART_COLORS[i%CHART_COLORS.length];
    const tgt=parseFloat((kpi.target||'').replace(/[^0-9.]/g,''));
    const ch=new Chart(ctx,{type:'bar',data:{labels:months,datasets:[
      {label:'ผลลัพธ์',data:kpi.data.map(d=>d.value),backgroundColor:kpi.data.map(d=>d.value===null?'rgba(200,200,200,.2)':evaluateTarget(d.value,kpi.target)?'rgba(16,185,129,.65)':'rgba(239,68,68,.65)'),borderRadius:5,barPercentage:.65,spanGaps:true},
      ...(!isNaN(tgt)?[{type:'line',label:'เป้าหมาย',data:months.map(()=>tgt),borderColor:'#f59e0b',borderWidth:2,borderDash:[4,3],pointRadius:0,fill:false}]:[])
    ]},options:{responsive:true,maintainAspectRatio:false,spanGaps:true,plugins:{legend:{display:false},tooltip:{...ttOpts(),callbacks:{label:ctx=>{const v=ctx.raw;if(ctx.dataset.label==='เป้าหมาย')return`เป้า: ${v}`;if(!v&&v!==0)return'ยังไม่บันทึก';return`ผล: ${v} ${evaluateTarget(v,kpi.target)?'✅':'❌'}`;}}}},scales:{x:{grid:{display:false},ticks:{font:{family:'Sarabun',size:9},color:'#9ca3af',maxRotation:60}},y:{grid:{color:'#f8f9fa'},ticks:{font:{family:'Sarabun',size:10},color:'#9ca3af'},beginAtZero:true}}}});
    State.wardCharts.push(ch);
  });
}
function heatmapHTML(trend,months){
  if(!trend.length)return'<div style="color:#9ca3af;padding:8px">ยังไม่มีข้อมูล</div>';
  const rows=trend.slice(0,6).map(kpi=>{
    const cells=months.map((_,mi)=>{const v=kpi.data[mi]?.value;const l=v===null||v===undefined?'level-0':evaluateTarget(v,kpi.target)?'level-3':'level-fail';return`<div class="ward-heatmap-cell ${l}" title="${months[mi]}: ${v!==null&&v!==undefined?v:'ไม่มีข้อมูล'}"></div>`;}).join('');
    const n=kpi.kpiName.length>30?kpi.kpiName.slice(0,30)+'...':kpi.kpiName;
    return`<div style="display:flex;align-items:center;gap:10px;margin-bottom:6px"><div style="width:160px;font-size:.72rem;color:#555;font-weight:600;text-align:right;flex-shrink:0">${n}</div><div class="ward-heatmap" style="flex:1">${cells}</div></div>`;
  }).join('');
  return rows+`<div style="display:flex;align-items:center;gap:10px;margin-top:8px"><div style="width:160px;flex-shrink:0"></div><div class="ward-heatmap-labels" style="flex:1">${months.map(m=>`<div class="ward-heatmap-label">${m.split('.')[0]}</div>`).join('')}</div></div>`;
}
function wardSkeleton(){return`<div class="ward-hero ward-skeleton" style="height:200px;background:linear-gradient(135deg,#d1ede9,#e8f7f5)"></div><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:14px;margin:20px 0">${[1,2,3,4].map(()=>`<div class="ward-skeleton" style="height:160px;border-radius:20px"></div>`).join('')}</div><div class="ward-skeleton" style="height:55px;border-radius:14px;margin-bottom:14px"></div><div class="ward-skeleton" style="height:250px;border-radius:20px"></div>`;}

// ═══ VIEW SWITCHER (Dashboard) ════════════════════════════
function switchDashView(view){
  dashView=view;
  ['table','bar','line'].forEach(v=>{document.getElementById(`dash-view-btn-${v}`)?.classList.toggle('active',v===view);document.getElementById(`dash-view-${v}`)?.classList.toggle('d-none',v!==view);});
  if(view==='bar')buildDashBarChart();
  if(view==='line')buildDashLineChart();
}
function buildDashBarChart(){
  const{summary=[]}=State.dashboardData||{};if(!summary.length)return;
  const sel=document.getElementById('dash-bar-kpi-select');
  if(sel&&!sel.children.length){const all=document.createElement('option');all.value='__all__';all.textContent='— ทุกตัวชี้วัด —';sel.appendChild(all);summary.forEach(k=>{const o=document.createElement('option');o.value=k.kpiId;o.textContent=k.kpiName.length>50?k.kpiName.slice(0,50)+'...':k.kpiName;sel.appendChild(o);});}
  renderDashBarChart(summary);
  const lbl=document.getElementById('dash-bar-label');if(lbl)lbl.textContent=document.getElementById('dash-subtitle')?.textContent||'';
}
function renderDashBarChart(summary){
  const ctx=document.getElementById('dash-bar-chart');if(!ctx)return;
  if(dashBarChart){dashBarChart.destroy();dashBarChart=null;}
  dashBarChart=new Chart(ctx,{type:'bar',data:{labels:summary.map((_,i)=>`KPI ${i+1}`),datasets:[{label:'ผลลัพธ์',data:summary.map(k=>k.latestValue!==null?parseFloat(k.latestValue):null),backgroundColor:summary.map(k=>{if(k.latestValue===null)return'rgba(200,200,200,.35)';return evaluateTarget(k.latestValue,k.target)?'rgba(16,185,129,.75)':'rgba(239,68,68,.75)';}),borderRadius:6,barPercentage:.65}]},options:{responsive:true,maintainAspectRatio:false,spanGaps:true,plugins:{legend:{display:false},tooltip:{...ttOpts(),callbacks:{title:items=>summary[items[0].dataIndex]?.kpiName||'',label:item=>{const k=summary[item.dataIndex];if(!k||item.raw===null)return'ยังไม่บันทึก';const ok=evaluateTarget(item.raw,k.target);return[`ผล: ${formatValue(item.raw,k.calcType)}`,`เป้า: ${k.target}`,ok?'✅ ผ่านเกณฑ์':'❌ ไม่ผ่าน'];}}}},scales:{x:{grid:{display:false},ticks:{font:{family:'Sarabun',size:11},color:'#4a6a67'}},y:{grid:{color:'#f0faf9'},ticks:{font:{family:'Sarabun',size:11},color:'#4a6a67'},beginAtZero:true}}}});
  renderProgressBars(summary);
}
function updateBarChart(){const{summary=[]}=State.dashboardData||{};const sel=document.getElementById('dash-bar-kpi-select');if(!sel)return;const id=sel.value;renderDashBarChart(id==='__all__'?summary:summary.filter(k=>k.kpiId===id));}
function buildDashLineChart(){
  const{trend=[],months=[]}=State.dashboardData||{};if(!trend.length)return;
  const sel=document.getElementById('dash-line-kpi-select');
  if(sel&&!sel.children.length){trend.forEach((t,i)=>{const o=document.createElement('option');o.value=i;o.textContent=t.kpiName.length>50?t.kpiName.slice(0,50)+'...':t.kpiName;sel.appendChild(o);});}
  renderDashLineChart(0);
}
function renderDashLineChart(idx){
  const{trend=[],months=[]}=State.dashboardData||{};
  const ctx=document.getElementById('dash-line-chart');if(!ctx)return;
  if(dashLineChart){dashLineChart.destroy();dashLineChart=null;}
  if(!trend[idx])return;
  const kpi=trend[idx],target=parseFloat((kpi.target||'').replace(/[^0-9.]/g,''));
  const values=kpi.data.map(d=>d.value);
  dashLineChart=new Chart(ctx,{type:'line',data:{labels:months,datasets:[
    {label:kpi.kpiName,data:values,borderColor:'#359286',backgroundColor:'rgba(53,146,134,0.08)',borderWidth:2.5,pointBackgroundColor:values.map(v=>v===null?'#d1d5db':evaluateTarget(v,kpi.target)?'#10b981':'#ef4444'),pointRadius:5,pointHoverRadius:8,tension:0.35,fill:true,spanGaps:true},
    ...(!isNaN(target)?[{label:`เป้าหมาย (${kpi.target})`,data:months.map(()=>target),borderColor:'#ef4444',borderWidth:2,borderDash:[6,4],pointRadius:0,fill:false}]:[])
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{intersect:false,mode:'index'},spanGaps:true,plugins:{legend:{display:true,position:'top',labels:{font:{family:'Sarabun',size:12},usePointStyle:true,boxWidth:10}},tooltip:{...ttOpts(),callbacks:{label:item=>{if(item.dataset.label?.includes('เป้าหมาย'))return`เป้าหมาย: ${item.raw}`;const v=item.raw;if(!v&&v!==0)return'ยังไม่บันทึก';return`ผล: ${formatValue(v,kpi.calcType||'')} ${evaluateTarget(v,kpi.target)?'✅':'❌'}`;}}}},scales:{x:{grid:{color:'#f0faf9'},ticks:{font:{family:'Sarabun',size:11},color:'#4a6a67',maxRotation:45}},y:{grid:{color:'#f0faf9'},ticks:{font:{family:'Sarabun',size:11},color:'#4a6a67'},beginAtZero:true}}}});
}
function updateLineChart(){const sel=document.getElementById('dash-line-kpi-select');if(!sel)return;renderDashLineChart(parseInt(sel.value)||0);}

// ═══ VIEW SWITCHER (Ward) ═════════════════════════════════
function switchWardView(view){
  wardView=view;
  ['table','bar','line'].forEach(v=>{document.getElementById(`ward-view-btn-${v}`)?.classList.toggle('active',v===view);});
  const{summary=[],trend=[],months=[]}=State.dashboardData||{};
  const grid=document.getElementById('ward-kpi-grid');
  if(view==='table'){if(grid){grid.innerHTML=kpiRingCards(summary);grid.style.display='';} setTimeout(()=>animateRings(summary),80);}
  else if(view==='bar')renderWardBarView(summary);
  else if(view==='line')renderWardLineView(trend,months);
}
function renderWardBarView(summary){
  const grid=document.getElementById('ward-kpi-grid');if(!grid)return;
  if(wardBarChart){wardBarChart.destroy();wardBarChart=null;}
  grid.style.display='block';
  grid.innerHTML=`<div class="ward-chart-card" style="animation:fadeSlideIn .4s ease both"><div class="ward-chart-card-header"><div class="ward-chart-card-title">📊 ผลลัพธ์ KPI เทียบกับเป้าหมาย</div></div><div class="ward-chart-card-body"><div style="height:${Math.max(280,summary.length*42)}px;position:relative"><canvas id="ward-bar-canvas"></canvas></div></div></div>`;
  setTimeout(()=>{
    const ctx=document.getElementById('ward-bar-canvas');if(!ctx)return;
    wardBarChart=new Chart(ctx,{type:'bar',data:{labels:summary.map(k=>k.kpiName.length>30?k.kpiName.slice(0,30)+'...':k.kpiName),datasets:[{label:'ผลลัพธ์',data:summary.map(k=>k.latestValue!==null?parseFloat(k.latestValue):null),backgroundColor:summary.map(k=>{if(k.latestValue===null)return'rgba(200,200,200,.35)';return evaluateTarget(k.latestValue,k.target)?'rgba(16,185,129,.75)':'rgba(239,68,68,.75)';}),borderRadius:5,barPercentage:.6}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,spanGaps:true,plugins:{legend:{display:false},tooltip:{...ttOpts(),callbacks:{label:item=>{const k=summary[item.dataIndex];if(!k||item.raw===null)return'ยังไม่บันทึก';const ok=evaluateTarget(item.raw,k.target);return[`ผล: ${formatValue(item.raw,k.calcType)}`,`เป้า: ${k.target}`,ok?'✅ ผ่าน':'❌ ไม่ผ่าน'];}}}},scales:{x:{grid:{color:'#f0faf9'},ticks:{font:{family:'Sarabun',size:11},color:'#9ca3af'},beginAtZero:true},y:{grid:{display:false},ticks:{font:{family:'Sarabun',size:11},color:'#555'}}}}});
  },50);
}
function renderWardLineView(trend,months){
  const grid=document.getElementById('ward-kpi-grid');if(!grid)return;
  if(wardLineChart){wardLineChart.destroy();wardLineChart=null;}
  if(!trend.length){grid.innerHTML='<div class="empty-state"><div class="empty-state-icon">📈</div><div class="empty-state-title">ยังไม่มีข้อมูลแนวโน้ม</div></div>';return;}
  grid.style.display='block';
  grid.innerHTML=`<div class="ward-chart-card" style="animation:fadeSlideIn .4s ease both"><div class="ward-chart-card-header" style="flex-wrap:wrap;gap:8px"><div class="ward-chart-card-title">📈 แนวโน้ม KPI รายเดือน</div><select id="ward-line-sel" class="form-control form-select" style="width:auto;font-size:.82rem" onchange="updateWardLineChart()">${trend.map((t,i)=>`<option value="${i}">${t.kpiName.length>50?t.kpiName.slice(0,50)+'...':t.kpiName}</option>`).join('')}</select></div><div class="ward-chart-card-body"><div style="height:320px;position:relative"><canvas id="ward-line-canvas"></canvas></div></div></div>`;
  setTimeout(()=>drawWardLineChart(trend,months,0),50);
}
function drawWardLineChart(trend,months,idx){
  const ctx=document.getElementById('ward-line-canvas');if(!ctx)return;
  if(wardLineChart){wardLineChart.destroy();wardLineChart=null;}
  const kpi=trend[idx];if(!kpi)return;
  const target=parseFloat((kpi.target||'').replace(/[^0-9.]/g,''));
  const values=kpi.data.map(d=>d.value);
  wardLineChart=new Chart(ctx,{type:'line',data:{labels:months,datasets:[
    {label:kpi.kpiName,data:values,borderColor:'#359286',backgroundColor:'rgba(53,146,134,0.08)',borderWidth:2.5,pointBackgroundColor:values.map(v=>v===null?'#d1d5db':evaluateTarget(v,kpi.target)?'#10b981':'#ef4444'),pointRadius:5,pointHoverRadius:8,tension:0.35,fill:true,spanGaps:true},
    ...(!isNaN(target)?[{label:`เป้าหมาย (${kpi.target})`,data:months.map(()=>target),borderColor:'#ef4444',borderWidth:2,borderDash:[6,4],pointRadius:0,fill:false}]:[])
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{intersect:false,mode:'index'},spanGaps:true,plugins:{legend:{display:true,position:'top',labels:{font:{family:'Sarabun',size:12},usePointStyle:true,boxWidth:10}},tooltip:ttOpts()},scales:{x:{grid:{color:'#f0faf9'},ticks:{font:{family:'Sarabun',size:11},color:'#4a6a67',maxRotation:45}},y:{grid:{color:'#f0faf9'},ticks:{font:{family:'Sarabun',size:11},color:'#4a6a67'},beginAtZero:true}}}});
}
function updateWardLineChart(){const{trend=[],months=[]}=State.dashboardData||{};const idx=parseInt(document.getElementById('ward-line-sel')?.value)||0;drawWardLineChart(trend,months,idx);}

// ═══ PROGRESS BARS ════════════════════════════════════════
function renderProgressBars(summary){
  const container=document.getElementById('dash-progress-bars');if(!container)return;
  if(!summary.length){container.innerHTML='<div class="empty-state" style="padding:30px 0"><div class="empty-state-icon">📊</div></div>';return;}
  const isAgg = summary[0].isAggregate === true;
  container.innerHTML=summary.map((k,i)=>{
    let pct, col, valText, targetText;
    if (isAgg) {
      pct = k.passRate;
      col = pct>=80?'#10b981':pct>=50?'#f59e0b':'#ef4444';
      valText = `${k.passRate}% (${k.passCount}/${k.totalUnits})`;
      targetText = `เป้า: ${k.target||'—'}`;
    } else {
      const hv=k.latestValue!==null&&k.latestValue!=='';
      const ok=hv&&evaluateTarget(k.latestValue,k.target);
      const val=hv?parseFloat(k.latestValue):null;
      const tgt=parseFloat((k.target||'').replace(/[^0-9.]/g,''));
      pct=(hv&&!isNaN(val)&&!isNaN(tgt)&&tgt>0)?Math.min(100,Math.round((val/tgt)*100)):hv?100:0;
      col=!hv?'#e5e7eb':ok?'#10b981':'#ef4444';
      valText=hv?formatValue(k.latestValue,k.calcType):'—';
      targetText=`เป้า: ${k.target||'—'}`;
    }
    return `<div class="kpi-progress-item">
      <div class="kpi-progress-header">
        <div class="kpi-progress-name" title="${k.kpiName}">${k.kpiName}</div>
        <div class="kpi-progress-val" style="color:${col}">${valText}</div>
        <div class="kpi-progress-target">${targetText}</div>
      </div>
      <div class="progress-wrap"><div class="progress-bar" style="width:${pct}%;background:${col};transition:width .8s ease ${i*.04}s"></div></div>
    </div>`;
  }).join('');
}

// ═══ BATCH ENTRY ══════════════════════════════════════════
function initEntryPage(){
  const{role,unitId}=State.user;
  if(role==='User'){document.getElementById('entry-group-col')?.classList.add('d-none');document.getElementById('entry-unit-col')?.classList.add('d-none');State.entry.unitId=unitId;}
  else{document.getElementById('entry-group-col')?.classList.remove('d-none');document.getElementById('entry-unit-col')?.classList.remove('d-none');}
  updateMonthDisplay();
  document.getElementById('batch-table-section')?.classList.add('d-none');
}
function changeMonth(d){State.entry.month+=d;if(State.entry.month>11){State.entry.month=0;State.entry.year++;}if(State.entry.month<0){State.entry.month=11;State.entry.year--;}updateMonthDisplay();}
function updateMonthDisplay(){document.getElementById('entry-month-display').textContent=`${TH_MONTHS[State.entry.month]}${State.entry.year+543}`;}
function getMonthYearString(){return`${TH_MONTHS[State.entry.month]}${State.entry.year+543}`;}
function onEntryUnitChange(){State.entry.unitId=document.getElementById('entry-unit-select').value;}

async function loadBatchEntryData(){
  const{role}=State.user;
  const unitId=role==='User'?State.user.unitId:document.getElementById('entry-unit-select').value;
  const monthYear=getMonthYearString();
  if(!unitId){showToast('กรุณาเลือกหน่วยงาน','warning');return;}
  State.entry.unitId=unitId;
  document.getElementById('batch-table-section')?.classList.remove('d-none');
  const tbody=document.getElementById('batch-entry-tbody');
  tbody.innerHTML=`<tr><td colspan="8" class="text-center py-4"><div class="spinner"></div><div style="margin-top:8px;color:var(--text-muted);font-size:.85rem">กำลังโหลด KPI...</div></td></tr>`;
  const unit=State.units.find(u=>u.Unit_ID===unitId)||State.units.find(u=>String(u.Unit_ID)===String(unitId));
  document.getElementById('batch-unit-label').textContent=unit?.Unit_Name||unitId;
  document.getElementById('batch-month-label').textContent=monthYear;
  const[kr,dr]=await Promise.all([api('getKPIMaster',{unitId}),api('getKPIData',{unitId,monthYear})]);
  const kpis=kr.kpis||[],records=dr.records||[];
  if(!kpis.length){
    tbody.innerHTML=`<tr><td colspan="8"><div class="empty-state"><div class="empty-state-icon">📋</div><div class="empty-state-title">ไม่พบตัวชี้วัด KPI สำหรับหน่วยงานนี้</div><div class="empty-state-desc">Unit_ID ที่ใช้: ${unitId}</div></div></td></tr>`;
    return;
  }
  tbody.innerHTML=kpis.map((kpi,i)=>{
    const ex=records.find(r=>r.KPI_ID===kpi.KPI_ID)||records.find(r=>String(r.KPI_ID)===String(kpi.KPI_ID));
    const isR=kpi.Calc_Type==='x100'||kpi.Calc_Type==='x1000';
    const num=ex?.Numerator||'',den=ex?.Denominator||'';
    const res=(isR&&num&&den)?calcFromRatio(num,den,kpi.Calc_Type):(ex?.Result_Value||'');
    return`<tr data-kpi-id="${kpi.KPI_ID}" data-calc-type="${kpi.Calc_Type}">
      <td style="color:var(--text-muted);font-weight:600;font-size:.82rem">${i+1}</td>
      <td class="kpi-name-cell">${kpi.KPI_Name}<small>${kpi.KPI_Category||''}</small></td>
      <td class="target-display">${kpi.Target||'—'}</td>
      <td><input type="number" step="any" class="batch-input ${!isR?'locked':''}" id="num-${kpi.KPI_ID}" value="${num}" ${!isR?'disabled placeholder="—"':'placeholder="ตัวเศษ"'} oninput="onBatchInputChange('${kpi.KPI_ID}','${kpi.Calc_Type}')"/></td>
      <td><input type="number" step="any" class="batch-input ${!isR?'locked':''}" id="den-${kpi.KPI_ID}" value="${den}" ${!isR?'disabled placeholder="—"':'placeholder="ตัวส่วน"'} oninput="onBatchInputChange('${kpi.KPI_ID}','${kpi.Calc_Type}')"/></td>
      <td><input type="number" step="any" class="batch-input ${isR?'locked':''}" id="res-${kpi.KPI_ID}" value="${res}" ${isR?'disabled placeholder="คำนวณอัตโนมัติ"':'placeholder="กรอกผลลัพธ์"'} oninput="updateBatchSummary()"/></td>
      <td><span style="font-size:.78rem;background:var(--primary-xlight);color:var(--primary);padding:3px 8px;border-radius:6px;font-weight:700">${kpi.Calc_Type}</span></td>
      <td><span class="status-badge ${ex?'pass':'na'}" id="status-${kpi.KPI_ID}">${ex?'✅ มีข้อมูล':'○ ว่าง'}</span></td>
    </tr>`;
  }).join('');
  updateBatchSummary();
}
function onBatchInputChange(id,t){if(t!=='x100'&&t!=='x1000')return;const n=document.getElementById(`num-${id}`)?.value,d=document.getElementById(`den-${id}`)?.value,r=document.getElementById(`res-${id}`);if(!r)return;r.value=(n&&d&&parseFloat(d)>0)?calcFromRatio(n,d,t):'';updateBatchSummary();}
function calcFromRatio(n,d,t){const nv=parseFloat(n),dv=parseFloat(d);if(isNaN(nv)||isNaN(dv)||dv===0)return'';return((nv/dv)*(t==='x1000'?1000:100)).toFixed(4);}
function updateBatchSummary(){
  const rows=document.querySelectorAll('#batch-entry-tbody tr[data-kpi-id]');let f=0;
  rows.forEach(r=>{const id=r.dataset.kpiId;if(document.getElementById(`res-${id}`)?.value?.trim()){f++;const s=document.getElementById(`status-${id}`);if(s&&s.textContent.includes('ว่าง')){s.className='status-badge pending';s.textContent='✏️ แก้ไขแล้ว';}}});
  document.getElementById('batch-filled-count').textContent=f;
  document.getElementById('batch-total-count').textContent=rows.length;
  document.getElementById('batch-last-update').textContent=new Date().toLocaleTimeString('th-TH');
}
function clearBatchForm(){
  document.querySelectorAll('#batch-entry-tbody tr[data-kpi-id]').forEach(r=>{const id=r.dataset.kpiId;const n=document.getElementById(`num-${id}`),d=document.getElementById(`den-${id}`),res=document.getElementById(`res-${id}`);if(n&&!n.disabled)n.value='';if(d&&!d.disabled)d.value='';if(res)res.value='';});
  updateBatchSummary();showToast('ล้างข้อมูลทั้งหมดแล้ว','info');
}
async function saveBatchAll(){
  const{unitId}=State.entry,monthYear=getMonthYearString();
  if(!unitId){showToast('กรุณาเลือกหน่วยงาน','warning');return;}
  const entries=[];
  document.querySelectorAll('#batch-entry-tbody tr[data-kpi-id]').forEach(r=>{const id=r.dataset.kpiId;const num=document.getElementById(`num-${id}`)?.value||'';const den=document.getElementById(`den-${id}`)?.value||'';const res=document.getElementById(`res-${id}`)?.value||'';if(num||den||res)entries.push({kpiId:id,resultValue:res,numerator:num,denominator:den});});
  if(!entries.length){showToast('ไม่มีข้อมูลที่จะบันทึก','warning');return;}
  const btn=document.getElementById('batch-save-btn');btn.innerHTML='<span class="spinner"></span> กำลังบันทึก...';btn.disabled=true;
  const r=await api('saveBatchKPI',{monthYear,unitId,entries});
  btn.innerHTML='<i class="fa-solid fa-floppy-disk"></i> บันทึกทั้งหมด (Save All)';btn.disabled=false;
  if(r.success){showToast(r.message,'success');document.querySelectorAll('#batch-entry-tbody tr[data-kpi-id]').forEach(row=>{const id=row.dataset.kpiId;if(document.getElementById(`res-${id}`)?.value){const s=document.getElementById(`status-${id}`);if(s){s.className='status-badge pass';s.textContent='✅ มีข้อมูล';}}});}
  else showToast(r.message,'error');
}

// ═══ ADMIN ════════════════════════════════════════════════
function initAdminPanel(){if(State.user?.role!=='Admin'){navigate('dashboard');return;}switchAdminTab('users');}
function switchAdminTab(tab){
  State.adminTab=tab;
  const tabs=['users','units','kpis'];
  document.querySelectorAll('.admin-tab').forEach((b,i)=>b.classList.toggle('active',tabs[i]===tab));
  document.querySelectorAll('.admin-tab-content').forEach(e=>e.classList.add('d-none'));
  document.getElementById(`admin-tab-${tab}`)?.classList.remove('d-none');
  if(tab==='users')loadAdminUsers();else if(tab==='units')loadAdminUnits();else loadAdminKPIs();
}
async function loadAdminUsers(){
  const tb=document.getElementById('admin-users-tbody');tb.innerHTML='<tr><td colspan="6" class="text-center py-4"><div class="spinner"></div></td></tr>';
  const r=await api('getUsers');if(!r.success||!r.users?.length){tb.innerHTML='<tr><td colspan="6"><div class="empty-state"><div class="empty-state-icon">👥</div><div class="empty-state-title">ไม่พบผู้ใช้งาน</div></div></td></tr>';return;}
  const rc={Admin:'#359286',NSO:'#3b82f6',User:'#6b7280'};
  tb.innerHTML=r.users.map(u=>{const unit=State.units.find(x=>String(x.Unit_ID)===String(u.Unit_ID));return`<tr><td style="font-size:.78rem;color:var(--text-muted)">${u.User_ID}</td><td style="font-weight:700">${u.Username}</td><td><span class="status-badge" style="background:${rc[u.Role]}20;color:${rc[u.Role]}">${u.Role}</span></td><td style="font-size:.85rem">${unit?unit.Unit_Name:(u.Unit_ID||'—')}</td><td style="font-size:.78rem;color:var(--text-muted)">${u.Created_Date||'—'}</td><td><div class="d-flex gap-1"><button class="btn btn-outline btn-sm btn-icon" onclick='openUserModal(${JSON.stringify(u).replace(/'/g,"&#39;")})'><i class="fa-solid fa-pen"></i></button><button class="btn btn-outline btn-sm btn-icon" onclick="confirmResetPassword('${u.User_ID}','${u.Username.replace(/'/g,"\\'")}')"><i class="fa-solid fa-key"></i></button><button class="btn btn-danger btn-sm btn-icon" onclick="confirmDelete('user','${u.User_ID}','${u.Username.replace(/'/g,"\\'")}')"><i class="fa-solid fa-trash"></i></button></div></td></tr>`;}).join('');
}
function openUserModal(u=null){document.getElementById('user-edit-id').value=u?.User_ID||'';document.getElementById('user-username').value=u?.Username||'';document.getElementById('user-role').value=u?.Role||'User';document.getElementById('user-unit').value=u?.Unit_ID||'';document.getElementById('user-password').value=u?'':'Password1234';document.getElementById('user-pw-section').style.display=u?'none':'';document.getElementById('modal-user-title').textContent=u?`แก้ไข: ${u.Username}`:'เพิ่มผู้ใช้งาน';openModal('modal-user');}
async function submitUserForm(){const id=document.getElementById('user-edit-id').value;const d={Username:document.getElementById('user-username').value.trim(),Role:document.getElementById('user-role').value,Unit_ID:document.getElementById('user-unit').value,password:document.getElementById('user-password').value};if(!d.Username){showToast('กรุณากรอก Username','warning');return;}const r=id?await api('updateUser',{...d,User_ID:id}):await api('createUser',d);if(r.success){showToast(r.message,'success');closeModal('modal-user');loadAdminUsers();}else showToast(r.message,'error');}
async function confirmResetPassword(uid,un){openConfirm(`รีเซ็ตรหัสผ่านของ "${un}" เป็น "Password1234"?`,async()=>{const r=await api('resetPassword',{userId:uid,newPassword:'Password1234'});showToast(r.message,r.success?'success':'error');});}
async function loadAdminUnits(){
  const tb=document.getElementById('admin-units-tbody');tb.innerHTML='<tr><td colspan="5" class="text-center py-4"><div class="spinner"></div></td></tr>';
  const r=await api('getUnits');if(!r.success||!r.units?.length){tb.innerHTML='<tr><td colspan="5"><div class="empty-state"><div class="empty-state-icon">🏢</div><div class="empty-state-title">ไม่พบหน่วยงาน</div></div></td></tr>';return;}
  tb.innerHTML=r.units.map(u=>`<tr><td style="font-weight:700;color:var(--primary)">${u.Group_ID}</td><td>${u.Group_Name}</td><td style="font-weight:700">${u.Unit_ID}</td><td>${u.Unit_Name}</td><td><div class="d-flex gap-1"><button class="btn btn-outline btn-sm btn-icon" onclick='openUnitModal(${JSON.stringify(u).replace(/'/g,"&#39;")})'><i class="fa-solid fa-pen"></i></button><button class="btn btn-danger btn-sm btn-icon" onclick="confirmDelete('unit','${u.Unit_ID}','${u.Unit_Name.replace(/'/g,"\\'")}')"><i class="fa-solid fa-trash"></i></button></div></td></tr>`).join('');
}
function openUnitModal(u=null){document.getElementById('unit-edit-id').value=u?.Unit_ID||'';document.getElementById('unit-group-id').value=u?.Group_ID||'';document.getElementById('unit-group-name').value=u?.Group_Name||'';document.getElementById('unit-unit-id').value=u?.Unit_ID||'';document.getElementById('unit-unit-name').value=u?.Unit_Name||'';document.getElementById('unit-unit-id').readOnly=!!u;document.getElementById('modal-unit-title').textContent=u?`แก้ไข: ${u.Unit_Name}`:'เพิ่มหน่วยงาน';openModal('modal-unit');}
async function submitUnitForm(){const id=document.getElementById('unit-edit-id').value;const d={Group_ID:document.getElementById('unit-group-id').value.trim(),Group_Name:document.getElementById('unit-group-name').value.trim(),Unit_ID:document.getElementById('unit-unit-id').value.trim(),Unit_Name:document.getElementById('unit-unit-name').value.trim()};if(!d.Group_ID||!d.Unit_ID||!d.Unit_Name){showToast('กรุณากรอกข้อมูลให้ครบ','warning');return;}const r=id?await api('updateUnit',d):await api('createUnit',d);if(r.success){showToast(r.message,'success');closeModal('modal-unit');loadAdminUnits();await loadMasterData();}else showToast(r.message,'error');}
async function loadAdminKPIs(){
  const tb=document.getElementById('admin-kpis-tbody');tb.innerHTML='<tr><td colspan="8" class="text-center py-4"><div class="spinner"></div></td></tr>';
  const r=await api('getKPIMaster',{});if(!r.success||!r.kpis?.length){tb.innerHTML='<tr><td colspan="8"><div class="empty-state"><div class="empty-state-icon">📋</div><div class="empty-state-title">ไม่พบ KPI</div></div></td></tr>';return;}
  tb.innerHTML=r.kpis.map(k=>{const u=State.units.find(x=>String(x.Unit_ID)===String(k.Unit_ID));const g=State.groups.find(x=>x.groupId===k.Unit_ID);return`<tr><td style="font-size:.75rem;color:var(--text-muted)">${k.KPI_ID}</td><td style="font-weight:600;max-width:220px">${k.KPI_Name}</td><td><span style="font-size:.78rem;background:var(--bg-main);padding:2px 7px;border-radius:5px">${k.KPI_Category||'—'}</span></td><td style="font-size:.82rem">${u?.Unit_Name||g?.groupName||k.Unit_ID||'—'}</td><td class="target-display">${k.Target||'—'}</td><td><span style="font-size:.75rem;background:var(--primary-xlight);color:var(--primary);padding:2px 7px;border-radius:5px;font-weight:700">${k.Calc_Type}</span></td><td><span class="status-badge ${k.Status==='Active'?'pass':'na'}">${k.Status||'Active'}</span></td><td><div class="d-flex gap-1"><button class="btn btn-outline btn-sm btn-icon" onclick='openKPIModal(${JSON.stringify(k).replace(/'/g,"&#39;")})'><i class="fa-solid fa-pen"></i></button><button class="btn btn-danger btn-sm btn-icon" onclick="confirmDelete('kpi','${k.KPI_ID}','${k.KPI_Name.replace(/'/g,"\\'")}')"><i class="fa-solid fa-trash"></i></button></div></td></tr>`;}).join('');
}
function openKPIModal(k=null){document.getElementById('kpi-edit-id').value=k?.KPI_ID||'';document.getElementById('kpi-name').value=k?.KPI_Name||'';document.getElementById('kpi-category').value=k?.KPI_Category||'';document.getElementById('kpi-unit').value=k?.Unit_ID||'';document.getElementById('kpi-target').value=k?.Target||'';document.getElementById('kpi-calc-type').value=k?.Calc_Type||'x100';document.getElementById('kpi-status').value=k?.Status||'Active';document.getElementById('modal-kpi-title').textContent=k?'แก้ไข KPI':'เพิ่มตัวชี้วัด KPI';openModal('modal-kpi');}
async function submitKPIForm(){const id=document.getElementById('kpi-edit-id').value;const d={KPI_ID:id,KPI_Name:document.getElementById('kpi-name').value.trim(),KPI_Category:document.getElementById('kpi-category').value.trim(),Unit_ID:document.getElementById('kpi-unit').value,Target:document.getElementById('kpi-target').value.trim(),Calc_Type:document.getElementById('kpi-calc-type').value,Status:document.getElementById('kpi-status').value};if(!d.KPI_Name||!d.Target){showToast('กรุณากรอกชื่อ KPI และเป้าหมาย','warning');return;}const r=id?await api('updateKPIMaster',d):await api('createKPIMaster',d);if(r.success){showToast(r.message,'success');closeModal('modal-kpi');loadAdminKPIs();State.kpiMaster=[];}else showToast(r.message,'error');}

// ═══ CONFIRM / CHIPS / HELPERS ════════════════════════════
function confirmDelete(type,id,name){const m={user:`ลบผู้ใช้ "${name}"?`,unit:`ลบหน่วยงาน "${name}"?`,kpi:`ลบตัวชี้วัด "${name}"?`};openConfirm(m[type],async()=>{const a={user:'deleteUser',unit:'deleteUnit',kpi:'deleteKPIMaster'};const k={user:'userId',unit:'unitId',kpi:'kpiId'};const r=await api(a[type],{[k[type]]:id});if(r.success){showToast(r.message,'success');if(type==='user')loadAdminUsers();else if(type==='unit'){loadAdminUnits();loadMasterData();}else loadAdminKPIs();}else showToast(r.message,'error');});}
function openConfirm(msg,onOk){document.getElementById('confirm-message').textContent=msg;const b=document.getElementById('confirm-ok-btn');b.onclick=()=>{closeModal('modal-confirm');onOk();};openModal('modal-confirm');}
function openChangePassword(){['cpw-old','cpw-new','cpw-confirm'].forEach(id=>{document.getElementById(id).value='';});openModal('modal-change-pw');}
async function submitChangePassword(){const o=document.getElementById('cpw-old').value,n=document.getElementById('cpw-new').value,c=document.getElementById('cpw-confirm').value;if(!o||!n||!c){showToast('กรุณากรอกข้อมูลให้ครบ','warning');return;}if(n!==c){showToast('รหัสผ่านใหม่ไม่ตรงกัน','error');return;}if(n.length<8){showToast('รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร','warning');return;}const r=await api('changePassword',{oldPassword:o,newPassword:n});if(r.success){showToast(r.message,'success');closeModal('modal-change-pw');}else showToast(r.message,'error');}
function updateDashChips(role,groupId,unitId,mIdx,be){
  const container=document.getElementById('dash-active-chips');if(!container)return;container.innerHTML='';
  if(groupId&&role!=='User'){const g=State.groups.find(x=>x.groupId===groupId);container.appendChild(makeChip(`<i class="fa-solid fa-sitemap"></i> ${g?.groupName||groupId}`,()=>{document.getElementById('dash-group-filter').value='';onGroupFilterChange('dash');}));}
  if(unitId&&role!=='User'){const u=State.units.find(x=>x.Unit_ID===unitId);container.appendChild(makeChip(`<i class="fa-solid fa-hospital"></i> ${u?.Unit_Name||unitId}`,()=>{document.getElementById('dash-unit-filter').value='';loadDashboard();}));}
  if(mIdx!==null){container.appendChild(makeChip(`<i class="fa-solid fa-calendar-day"></i> ${TH_MONTHS_FULL[mIdx]} พ.ศ. ${be}`,()=>{document.getElementById('dash-month-filter').value='';loadDashboard();}));}
}
function makeChip(html,onClose){const chip=document.createElement('div');chip.className='filter-active-chip';chip.innerHTML=html+' <span class="chip-close">✕</span>';chip.querySelector('.chip-close').onclick=e=>{e.stopPropagation();onClose();};return chip;}
function exportDashboardData(){
  if(!State.dashboardData?.summary){showToast('ไม่มีข้อมูลสำหรับ Export','warning');return;}
  const isAgg = State.dashboardData.summary.length > 0 && State.dashboardData.summary[0].isAggregate === true;
  let rows, csv;
  if (isAgg) {
    rows=[['ตัวชี้วัด','หมวดหมู่','เป้าหมาย','หน่วยทั้งหมด','ผ่านเกณฑ์','ไม่ผ่าน','ยังไม่บันทึก','อัตราผ่าน(%)']];
    State.dashboardData.summary.forEach(k=>{
      rows.push([k.kpiName,k.category,k.target,k.totalUnits,k.passCount,k.failCount,k.pendingCount,k.passRate]);
    });
  } else {
    rows=[['KPI_ID','ชื่อตัวชี้วัด','หมวดหมู่','เป้าหมาย','ผลลัพธ์','เดือน','สถานะ']];
    State.dashboardData.summary.forEach(k=>{
      rows.push([k.kpiId,k.kpiName,k.category,k.target,k.latestValue||'',k.latestMonth||'',k.latestValue!==null?(evaluateTarget(k.latestValue,k.target)?'ผ่าน':'ไม่ผ่าน'):'ยังไม่บันทึก']);
    });
  }
  csv=rows.map(r=>r.map(c=>`"${c}"`).join(',')).join('\n');
  const b=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'});
  const u=URL.createObjectURL(b);const a=document.createElement('a');a.href=u;
  a.download=`KPI_${isAgg?'ภาพรวม_':''}${new Date().toISOString().slice(0,10)}.csv`;
  a.click();URL.revokeObjectURL(u);showToast('Export สำเร็จ','success');
}

// ═══ EVALUATE / FORMAT ════════════════════════════════════
function evaluateTarget(v,t){
  if(!t||v===null||v===''||v===undefined)return false;
  const n=parseFloat(v);if(isNaN(n))return false;
  t=String(t).trim();
  if(t.startsWith('≥')||t.startsWith('>='))return n>=parseFloat(t.replace(/^[≥>]=?/,'').trim());
  if(t.startsWith('≤')||t.startsWith('<='))return n<=parseFloat(t.replace(/^[≤<]=?/,'').trim());
  if(t.startsWith('>'))return n>parseFloat(t.slice(1).trim());
  if(t.startsWith('<'))return n<parseFloat(t.slice(1).trim());
  if(t.startsWith('='))return n===parseFloat(t.slice(1).trim());
  const nt=parseFloat(t);return!isNaN(nt)&&n>=nt;
}
function formatValue(v,t){const n=parseFloat(v);if(isNaN(n))return String(v);if(t==='x100')return n.toFixed(2)+'%';if(t==='x1000')return n.toFixed(2)+'‰';return n%1===0?String(n):n.toFixed(2);}

// ═══ MODAL / TOAST ════════════════════════════════════════
function openModal(id){document.getElementById(id)?.classList.add('open');}
function closeModal(id){document.getElementById(id)?.classList.remove('open');}
document.addEventListener('click',e=>{if(e.target.classList.contains('modal-overlay'))e.target.classList.remove('open');});
function showToast(msg,type='info',dur=4000){
  const icons={success:'✅',error:'❌',warning:'⚠️',info:'ℹ️'};
  const t=document.createElement('div');t.className=`toast ${type}`;
  t.innerHTML=`<span class="toast-icon">${icons[type]}</span><span class="toast-message">${msg}</span><span class="toast-close" onclick="this.parentElement.remove()">✕</span>`;
  document.getElementById('toast-container').appendChild(t);
  setTimeout(()=>{t.style.cssText+='opacity:0;transform:translateX(30px);transition:all .3s ease';setTimeout(()=>t.remove(),300);},dur);
}

// ═══ DIAGNOSTIC (Admin) ═══════════════════════════════════
async function runApiDiagnostic(){
  showToast('กำลัง diagnose...','info',10000);
  const results={};
  const u=await api('getUnits');
  results.getUnits=u.success?`✅ ${(u.units||[]).length} units`:`❌ ${u.message}`;
  const k=await api('getKPIMaster',{unitId:State.user?.unitId||''});
  results.getKPIMaster=k.success?`✅ ${(k.kpis||[]).length} KPIs`:`❌ ${k.message}`;
  const d=await api('getKPIData',{unitId:State.user?.unitId||''});
  results.getKPIData=d.success?`✅ ${(d.records||[]).length} records`:`❌ ${d.message}`;
  const dd=await api('getDashboardData',{unitId:State.user?.unitId||'',year:String(new Date().getFullYear()),monthYear:''});
  results.getDashboardData=dd.success?`✅ summary=${(dd.summary||[]).length} months=${(dd.months||[]).slice(0,3).join(',')}`:`❌ ${dd.message}`;
  if(State.user?.role==='Admin'){const dbg=await api('debugInfo',{});if(dbg.success&&dbg.debug){Object.entries(dbg.debug).forEach(([k,v])=>{results['Sheet:'+k]=typeof v==='object'?`rows=${v.rows}`:(String(v).slice(0,80));});}}
  const msg=Object.entries(results).map(([k,v])=>`${k}:\n  ${v}`).join('\n\n');
  console.table(results);
  alert('=== API Diagnostic ===\n\n'+msg);
}

// ============================================================
// AUDIT TRAIL
// ============================================================

const AUDIT_ACTION_STYLES = {
  'LOGIN':           { bg:'#dbeafe', color:'#1e40af', icon:'🔑' },
  'LOGOUT':          { bg:'#f3f4f6', color:'#6b7280', icon:'🚪' },
  'BATCH_SAVE':      { bg:'#d1fae5', color:'#065f46', icon:'💾' },
  'CREATE':          { bg:'#d1fae5', color:'#065f46', icon:'➕' },
  'UPDATE':          { bg:'#fef3c7', color:'#92400e', icon:'✏️' },
  'DELETE':          { bg:'#fee2e2', color:'#991b1b', icon:'🗑️' },
  'RESET_PASSWORD':  { bg:'#ede9fe', color:'#5b21b6', icon:'🔐' },
  'CHANGE_PASSWORD': { bg:'#ede9fe', color:'#5b21b6', icon:'🔒' },
};

const TABLE_LABELS = {
  'Users':          '👤 Users',
  'Units':          '🏢 Units',
  'KPI_Master':     '📋 KPI Master',
  'KPI_Data_Entry': '📊 KPI Data',
};

let _auditLogs = []; 

async function loadAuditLog() {
  const tbody = document.getElementById('audit-tbody');
  if (!tbody) return;

  tbody.innerHTML = `<tr><td colspan="6" class="text-center py-5">
    <div class="spinner spinner-lg" style="margin:0 auto 12px"></div>
    <div style="color:var(--text-muted);font-size:.85rem">กำลังโหลด Audit Log...</div>
  </td></tr>`;

  const filterAction = document.getElementById('audit-filter-action')?.value || '';
  const filterTable  = document.getElementById('audit-filter-table')?.value  || '';
  const filterUser   = document.getElementById('audit-filter-user')?.value   || '';
  const limit        = document.getElementById('audit-limit')?.value         || '100';

  const res = await api('getAuditLog', { filterAction, filterTable, filterUser, limit });

  if (!res.success) {
    tbody.innerHTML = `<tr><td colspan="6"><div class="empty-state" style="padding:48px">
      <div class="empty-state-icon">⚠️</div>
      <div class="empty-state-title">${res.message}</div>
    </div></td></tr>`;
    return;
  }

  populateAuditFilters(res);

  const totalEl = document.getElementById('audit-total-label');
  if (totalEl) totalEl.textContent = `แสดง ${res.logs.length} / ${res.total} รายการ`;

  _auditLogs = res.logs;

  if (!res.logs.length) {
    tbody.innerHTML = `<tr><td colspan="6"><div class="empty-state" style="padding:48px">
      <div class="empty-state-icon">📭</div>
      <div class="empty-state-title">ไม่พบ Audit Log</div>
      <div class="empty-state-desc">ลองปรับตัวกรองใหม่</div>
    </div></td></tr>`;
    return;
  }

  tbody.innerHTML = res.logs.map((log, i) => {
    const style  = AUDIT_ACTION_STYLES[log.action] || { bg:'#f3f4f6', color:'#374151', icon:'📌' };
    const tLabel = TABLE_LABELS[log.table] || log.table;
    const recId  = log.recordId.length > 30 ? log.recordId.slice(0, 30) + '...' : log.recordId;

    return `<tr style="animation:fadeSlideIn .3s ease ${i * 0.02}s both">
      <td style="color:var(--text-muted);font-size:.78rem;font-weight:600">${i+1}</td>
      <td style="font-size:.82rem;white-space:nowrap;color:var(--text-secondary)">${log.timestamp}</td>
      <td>
        <div style="font-weight:700;font-size:.88rem">${log.username}</div>
        <div style="font-size:.72rem;color:var(--text-muted)">${log.userId}</div>
      </td>
      <td>
        <span style="display:inline-flex;align-items:center;gap:5px;background:${style.bg};color:${style.color};padding:4px 10px;border-radius:99px;font-size:.78rem;font-weight:700">
          ${style.icon} ${log.action}
        </span>
      </td>
      <td style="font-size:.85rem;color:var(--text-secondary)">${tLabel}</td>
      <td style="font-size:.75rem;color:var(--text-muted);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${log.recordId}">${recId}</td>
    </tr>`;
  }).join('');
}

function populateAuditFilters(res) {
  const selectors = [
    { id:'audit-filter-action', items: res.actions || [] },
    { id:'audit-filter-table',  items: res.tables  || [] },
    { id:'audit-filter-user',   items: res.users   || [] },
  ];
  selectors.forEach(({ id, items }) => {
    const sel = document.getElementById(id);
    if (!sel || sel.children.length > 1) return;
    items.forEach(item => {
      const o = document.createElement('option');
      o.value = item;
      o.textContent = TABLE_LABELS[item] || item;
      sel.appendChild(o);
    });
  });
}

function exportAuditLog() {
  if (!_auditLogs.length) { showToast('ไม่มีข้อมูลสำหรับ Export','warning'); return; }
  const rows = [['#','เวลา','ผู้ใช้งาน','User_ID','Action','ตาราง','Record_ID']];
  _auditLogs.forEach((log, i) => {
    rows.push([i+1, log.timestamp, log.username, log.userId, log.action, log.table, log.recordId]);
  });
  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob = new Blob(['\ufeff'+csv], {type:'text/csv;charset=utf-8'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `AuditLog_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('Export Audit Log สำเร็จ','success');
}