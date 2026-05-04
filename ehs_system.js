// ============================================================
// SHAREPOINT REST API — 不需要 MSAL，利用已登入的 session cookie
// ============================================================

// 自動從當前頁面 URL 偵測 SP 站台與資料夾路徑
function spDetect(){
  var loc = window.location;
  // 站台 URL: https://xxxx.sharepoint.com/sites/EHS209
  var siteMatch = loc.pathname.match(/^(\/sites\/[^\/]+|\/teams\/[^\/]+)/);
  var siteRoot = siteMatch ? (loc.origin + siteMatch[1]) : loc.origin;
  // 當前 HTML 所在的資料夾（server-relative）
  var dir = loc.pathname.substring(0, loc.pathname.lastIndexOf('/'));
  return { siteRoot: siteRoot, currentDir: dir };
}

function spStatus(msg){ var el=document.getElementById('sp-status'); if(el) el.textContent=msg; }

// 列出某個 server-relative 路徑下的子資料夾與 CSV 檔案
async function spListFolder(siteRoot, serverRelPath){
  // 編碼路徑中的空格與特殊字元
  var encoded = serverRelPath.replace(/'/g,"''");
  var url = siteRoot + "/_api/web/GetFolderByServerRelativeUrl('"+ encoded +"')?$expand=Folders,Files&$select=Folders/Name,Folders/ServerRelativeUrl,Files/Name,Files/ServerRelativeUrl,Files/Length";
  var r = await fetch(url, {
    headers:{'Accept':'application/json;odata=verbose'},
    credentials:'include'
  });
  if(!r.ok) throw new Error('HTTP '+r.status+' - 請確認在 SharePoint 內開啟此檔案');
  var d = await r.json();
  var result = d.d || {};
  return {
    folders: (result.Folders && result.Folders.results) || [],
    files:   (result.Files   && result.Files.results)   || []
  };
}

// 遞迴搜尋所有 CSV 檔案（最多 2 層深）
async function spFindAllCSV(siteRoot, baseDir){
  var all = [];
  async function scanDir(path, depth){
    try{
      var res = await spListFolder(siteRoot, path);
      // 過濾 CSV 檔案
      res.files.filter(function(f){ return f.Name && f.Name.toLowerCase().endsWith('.csv'); })
               .forEach(function(f){ all.push({name:f.Name, url:f.ServerRelativeUrl, folder:path.split('/').pop()}); });
      // 遞迴進入子資料夾
      if(depth < 2){
        for(var i=0;i<res.folders.length;i++){
          var sub = res.folders[i];
          if(sub.Name && sub.Name.charAt(0)!=='.') await scanDir(sub.ServerRelativeUrl, depth+1);
        }
      }
    }catch(e){ /* 略過無法存取的子資料夾 */ }
  }
  await scanDir(baseDir, 0);
  return all;
}

async function spReadFile(siteRoot, serverRelUrl){
  var url = siteRoot + "/_api/web/GetFileByServerRelativeUrl('"+ serverRelUrl.replace(/'/g,"''") +"')/$value";
  var r = await fetch(url, {credentials:'include'});
  if(!r.ok) throw new Error('HTTP '+r.status);
  return await r.text();
}

// 主入口：點「瀏覽 CSV 檔案」
async function spBrowse(){
  var tree = document.getElementById('sp-tree');
  tree.style.display='block';
  tree.innerHTML='<div style="color:#475569;font-size:11px;padding:4px">⏳ 掃描資料夾中...</div>';
  spStatus('⏳ 連線中...');
  try{
    var sp = spDetect();
    var files = await spFindAllCSV(sp.siteRoot, sp.currentDir);
    if(!files.length){
      tree.innerHTML='<div style="color:#ef4444;font-size:11px;padding:4px">找不到 CSV 檔案（請確認在 SharePoint 內開啟）</div>';
      spStatus('找不到 CSV 檔案');
      return;
    }
    // 依資料夾分組顯示
    var byFolder = {};
    files.forEach(function(f){ if(!byFolder[f.folder]) byFolder[f.folder]=[]; byFolder[f.folder].push(f); });
    var html = '';
    Object.keys(byFolder).forEach(function(fld){
      html += '<div style="color:#22d3ee;font-size:9px;font-weight:700;margin-top:6px;margin-bottom:3px;letter-spacing:.08em">📁 '+esc(fld)+'</div>';
      byFolder[fld].forEach(function(f){
        var isEnv = f.name.toLowerCase().includes('env');
        var isRisk = f.name.toLowerCase().includes('risk');
        var col = isEnv?'#4ade80':isRisk?'#60a5fa':'#94a3b8';
        html += '<div class="sp-file-item" data-url="'+esc(f.url)+'" data-name="'+esc(f.name)+'" style="padding:4px 6px;border-radius:4px;cursor:pointer;font-size:10px;color:'+col+';transition:background .12s;margin-bottom:2px">'
          +(isEnv?'🌿':isRisk?'🛡️':'📄')+' '+esc(f.name)+'</div>';
      });
    });
    tree.innerHTML=html;
    // 為每個檔案加 hover + click（用事件委派）
    spStatus('✅ 找到 '+files.length+' 個 CSV，點擊載入');
  }catch(e){
    tree.innerHTML='<div style="color:#ef4444;font-size:11px;padding:4px">❌ '+esc(e.message)+'</div>';
    spStatus('失敗：'+e.message);
  }
}

// 事件委派：點擊 SP 樹狀清單中的檔案
document.addEventListener('click', function(e2){
  var item = e2.target.closest('.sp-file-item');
  if(!item) return;
  var url = item.dataset.url, name = item.dataset.name;
  if(!url) return;
  item.style.background='rgba(34,211,238,.15)';
  spStatus('⏳ 載入: '+name);
  var sp = spDetect();
  spReadFile(sp.siteRoot, url).then(function(text){
    var isEnv = name.toLowerCase().includes('env');
    processCSVText(text, isEnv?'env':'risk', name);
    spStatus('✅ 已載入: '+name);
  }).catch(function(err){
    spStatus('❌ '+err.message);
    toast('讀取失敗: '+err.message,'err');
    item.style.background='';
  });
});

// ============================================================
// STATE
// ============================================================
var S = { data:[], selectedId:null, tab:'entry' };

// ============================================================
// DICTIONARIES
// ============================================================
var ENV_CAT={
  '環保_空':['酸排','鹼排','一般排氣','VOC','油氣','臭(異)味'],
  '環保_水':['生活污水','冷卻水','氨氮廢水','酸性廢水','鹼性廢水'],
  '環保_廢棄物':['廢污泥','一般事業廢棄物','有害事業廢棄物','廢木材','廢 Wafer','生活垃圾'],
  '環保_毒化物':['毒化物','毒化物(BF3)','毒化物(PH3)'],
  '環保_能資源':['自來水','電力','柴油','LPG','氫氣','氮氣'],
  '環保_其他':['週界噪音','意外災害(天災、水災、火災、地震、颱風等)','環境用藥','其他'],
  '環保_客戶端':['依客戶規定處理']
};
var SAF_CAT={
  '物理性':['墜落,滾落','跌倒,滑倒','衝撞,被撞','夾,捲,壓傷','切,割,擦傷','溺斃','與高、低溫接觸','噪音過高','照明不足','通風不良','粉塵暴露'],
  '化學性':['爆炸','與有害物接觸','化學品洩漏(含廢液)','毒氣洩漏','噴濺腐蝕','吸入腐蝕','缺氧,窒息'],
  '人因工學':['操作高度、空間不適造成傷害','人工搬運超過荷重造成傷害','不適宜之工作姿勢造成傷害','重複性操作造成傷害','人為不當動作'],
  '生物性':['生物危害'],
  '其他':['過負荷','職場不法侵害','母性健康危害之虞','交通事故','未歸類者','其他']
};

// ============================================================
// UTILS
// ============================================================
function calcEnvSEA(ofq,impact){var m={1:{A:'S1',B:'S1',C:'S3'},2:{A:'S1',B:'S1',C:'S3'},3:{A:'S1',B:'S2',C:'S4'},4:{A:'S2',B:'S3',C:'S4'}};return(m[ofq]||{})[impact]||'S4';}
function calcEnvRisk(sea,prob,vio){if(vio)return'H';var m={1:{S1:'H',S2:'H',S3:'M',S4:'L'},2:{S1:'M',S2:'M',S3:'L',S4:'L'},3:{S1:'M',S2:'M',S3:'L',S4:'L'}};return(m[prob]||{})[sea]||'L';}
function calcSafClass(fr,s,vio){if(vio)return'Class 1';var m={100:{A:'Class 1',B:'Class 1',C:'Class 2',D:'Class 3',E:'Class 4'},80:{A:'Class 1',B:'Class 1',C:'Class 2',D:'Class 3',E:'Class 4'},50:{A:'Class 1',B:'Class 1',C:'Class 2',D:'Class 3',E:'Class 4'},40:{A:'Class 1',B:'Class 2',C:'Class 3',D:'Class 4',E:'Class 4'},25:{A:'Class 1',B:'Class 2',C:'Class 3',D:'Class 4',E:'Class 4'},10:{A:'Class 2',B:'Class 2',C:'Class 3',D:'Class 4',E:'Class 4'},8:{A:'Class 2',B:'Class 3',C:'Class 4',D:'Class 4',E:'Class 4'},5:{A:'Class 2',B:'Class 3',C:'Class 4',D:'Class 4',E:'Class 4'},1:{A:'Class 2',B:'Class 3',C:'Class 4',D:'Class 4',E:'Class 4'}};return(m[fr]||{})[s]||'Class 4';}
function parseMats(s){if(!s||!s.trim())return['無'];var a=s.split(/[、,，\s]+/).filter(Boolean);return a.length?a:['無'];}
function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2);}
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function opts(list,sel){return list.map(function(v){return'<option value="'+esc(v)+'"'+(v===sel?' selected':'')+'>'+esc(v)+'</option>';}).join('');}

var toastTimer;
function toast(msg,type){
  var el=document.getElementById('toast');
  el.className='show t-'+(type||'ok');
  el.textContent=msg;
  clearTimeout(toastTimer);
  toastTimer=setTimeout(function(){el.className='';},4000);
}
function closeModal(){document.getElementById('modal-overlay').className='';}

function getYear(){return document.getElementById('sel-year').value;}
function getPlant(){return document.getElementById('sel-plant').value;}
function getDept(){return document.getElementById('inp-dept').value;}

// ============================================================
// EVENT DELEGATION — THE FIX FOR SHAREPOINT CSP
// ============================================================
document.addEventListener('click', function(e){
  var el = e.target.closest('[data-action]');
  if(!el) {
    // close modal if clicking overlay
    if(e.target.id==='modal-overlay') closeModal();
    return;
  }
  var a=el.dataset.action, d=el.dataset;
  switch(a){
    case 'switch-tab': switchTab(d.tab); break;
    case 'show-format': showFormatModal(); break;
    case 'close-modal': closeModal(); break;
    case 'load-demo': loadDemo(); break;
    case 'create-proc': createProc(); break;
    case 'delete-proc': deleteProc(d.pid); break;
    case 'select-proc': selectProc(d.pid); break;
    case 'export-csv': exportCSV(); break;
    case 'add-step': addStep(d.pid); break;
    case 'delete-step': deleteStep(d.pid,d.sid); break;
    case 'add-env': addEnv(d.pid,d.sid); break;
    case 'delete-env': deleteEnv(d.pid,d.sid,d.eid); break;
    case 'add-saf': addSaf(d.pid,d.sid); break;
    case 'delete-saf': deleteSaf(d.pid,d.sid,d.sfid); break;
    case 'open-env-file': document.getElementById('inp-env').click(); break;
    case 'open-risk-file': document.getElementById('inp-risk').click(); break;
    case 'sp-browse': spBrowse(); break;
  }
});

document.addEventListener('change', function(e){
  var el=e.target, a=el.dataset.action;
  if(!a) {
    // year/plant/dept changes → update hint
    if(el.id==='sel-year'||el.id==='sel-plant'||el.id==='inp-dept') updateHint();
    return;
  }
  if(a==='csv-env'){handleCSV(el,'env');return;}
  if(a==='csv-risk'){handleCSV(el,'risk');return;}
  var d=el.dataset, pid=d.pid, sid=d.sid, eid=d.eid, sfid=d.sfid, field=d.field;
  switch(a){
    case 'proc-field': setProcField(pid,field,el.value); break;
    case 'set-common': setCommon(pid,sid,field,el.value); break;
    case 'set-common-mat': setCommonMat(pid,sid,el.value); break;
    case 'set-step-name': setStepName(pid,sid,el.value); break;
    case 'env-field': setEnvField(pid,sid,eid,field,el.value); break;
    case 'env-cat': setEnvCat(pid,sid,eid,el.value); break;
    case 'env-calc': setEnvCalc(pid,sid,eid,field,el.value); break;
    case 'env-vio': setEnvVio(pid,sid,eid,field,el.checked); break;
    case 'saf-field': setSafField(pid,sid,sfid,field,el.value); break;
    case 'saf-cat': setSafCat(pid,sid,sfid,el.value); break;
    case 'saf-calc': setSafCalc(pid,sid,sfid,field,el.value); break;
    case 'saf-vio': setSafVio(pid,sid,sfid,field,el.checked); break;
  }
});

// ============================================================
// SHAREPOINT FILE ACCESS
// ============================================================
// ============================================================
// CSV PARSE
// ============================================================
function handleCSV(input, type){
  var f=input.files[0]; if(!f) return;
  var r=new FileReader();
  r.onload=function(e){ processCSVText(e.target.result, type, f.name); };
  r.readAsText(f,'UTF-8');
  input.value='';
}

function processCSVText(text, type, fname){
  try{
    text=text.replace(/^\uFEFF/,'');
    var lines=text.split('\n').filter(function(l){return l.trim();});
    if(lines.length<2) throw new Error('檔案為空或格式錯誤');
    var header=lines[0].toLowerCase();
    var isEnv = type==='env' || header.includes('環境ofq') || header.includes('環境impact');
    var map={};
    for(var i=1;i<lines.length;i++){
      var cols=lines[i].split(',');
      if(cols.length<14) continue;
      var dept=(cols[0]||'').trim(), proc=(cols[1]||'').trim(), task=(cols[2]||'').trim(), step=(cols[3]||'').trim();
      if(!dept&&!proc) continue;
      var key=dept+'__'+proc+'__'+task;
      if(!map[key]) map[key]={id:uid(),department:dept,processName:proc,taskName:task,steps:[]};
      var st=null;
      for(var j=0;j<map[key].steps.length;j++){if(map[key].steps[j].stepName===step){st=map[key].steps[j];break;}}
      if(!st){
        st={id:uid(),stepName:step,
          common:{status:cols[4]||'例行',time:cols[5]||'現在',worker:cols[6]||'漢民員工',equipment:cols[7]||'',engControl:cols[8]||'',adminControl:cols[9]||'',ppe:cols[10]||'',material:cols[11]||''},
          envs:[],safeties:[]};
        map[key].steps.push(st);
      }
      if(isEnv){
        st.envs.push({id:uid(),material:cols[12]||'',category:cols[13]||'環保_空',envImpact:cols[14]||'',desc:cols[15]||'',
          ofq:parseInt(cols[16])||4,impact:cols[17]||'C',prob:parseInt(cols[18])||3,
          sea:cols[19]||'S4',riskLevel:(cols[20]||'L').trim(),
          v1:false,v2:false,v3:false,treatment:cols[21]||'',controlPlan:cols[22]||''});
      } else {
        st.safeties.push({id:uid(),material:cols[12]||'',riskCat:cols[13]||'物理性',hazardDesc:cols[14]||'',consequence:cols[15]||'',
          of:parseInt(cols[16])||1,p:parseInt(cols[17])||1,fr:parseInt(cols[18])||1,s:cols[19]||'E',
          classLevel:(cols[20]||'Class 4').trim(),isOpp:cols[21]||'',treatment:cols[22]||'',treatmentDesc:cols[23]||'',
          v1:false,v2:false,v3:false});
      }
    }
    var parsed=Object.values(map);
    // merge
    parsed.forEach(function(np){
      var ex=null;
      for(var i=0;i<S.data.length;i++){
        if(S.data[i].processName===np.processName&&S.data[i].taskName===np.taskName&&S.data[i].department===np.department){ex=S.data[i];break;}
      }
      if(ex){
        np.steps.forEach(function(ns){
          var es=null;
          for(var i=0;i<ex.steps.length;i++){if(ex.steps[i].stepName===ns.stepName){es=ex.steps[i];break;}}
          if(es){if(isEnv)es.envs=ns.envs;else es.safeties=ns.safeties;}
          else ex.steps.push(ns);
        });
      } else { S.data.push(np); }
    });
    if(!S.selectedId&&S.data.length) S.selectedId=S.data[0].id;
    renderAll();
    toast('✅ '+(isEnv?'Env':'Risk')+' 匯入成功：'+parsed.length+' 筆 ('+fname+')','ok');
  }catch(err){ toast('❌ '+err.message,'err'); }
}

// ============================================================
// DEMO DATA
// ============================================================
function loadDemo(){
  S.data=[
    {id:'p1',department:'FAC',processName:'空調水系統',taskName:'加藥系統',steps:[
      {id:'s1',stepName:'藥液添加作業',
        common:{status:'非例行',time:'現在',worker:'承攬商',equipment:'手工具',material:'腐蝕結垢抑制劑、管路殺菌滅藻劑',engControl:'防液堤',adminControl:'SDS',ppe:'耐酸鹼手套、面罩'},
        envs:[{id:'e1',material:'腐蝕結垢抑制劑',category:'環保_廢棄物',envImpact:'一般事業廢棄物',desc:'藥劑空桶處理',ofq:4,impact:'C',prob:3,sea:'S4',riskLevel:'L',v1:false,v2:false,v3:false,treatment:'',controlPlan:''}],
        safeties:[{id:'sf1',material:'管路殺菌滅藻劑',riskCat:'化學性',hazardDesc:'添加時遭藥液噴濺',consequence:'與有害物接觸',of:5,p:1,s:'E',fr:5,classLevel:'Class 4',v1:false,v2:false,v3:false,isOpp:'',treatment:'',treatmentDesc:''}]
      }]
    },
    {id:'p2',department:'Mplan',processName:'機台組裝',taskName:'乙醇分裝',steps:[
      {id:'s2',stepName:'管制區分裝',
        common:{status:'例行',time:'現在',worker:'漢民員工',equipment:'無',material:'乙醇',engControl:'通風設備',adminControl:'化學品管制',ppe:'活性碳口罩'},
        envs:[{id:'e2',material:'乙醇',category:'環保_空',envImpact:'VOC',desc:'乙醇分裝時揮發',ofq:2,impact:'B',prob:2,sea:'S1',riskLevel:'M',v1:false,v2:false,v3:false,treatment:'建立管理方案',controlPlan:'建立排氣檢測機制'}],
        safeties:[{id:'sf2',material:'乙醇',riskCat:'化學性',hazardDesc:'分裝時吸入揮發性有機氣體',consequence:'與有害物接觸',of:5,p:1,s:'A',fr:5,classLevel:'Class 2',v1:false,v2:false,v3:false,isOpp:'',treatment:'建立目標方案',treatmentDesc:'定期空氣品質監測'}]
      }]
    }]
  ];
  S.selectedId='p1';
  renderAll();
  toast('已載入範例資料','info');
}

// ============================================================
// EXPORT
// ============================================================
function exportCSV(){
  if(!S.data.length){toast('沒有資料可匯出','err');return;}
  var year=getYear(),plant=getPlant();
  var dept=(S.data[0].department||'Dept').replace(/[\/\\?%*:|"<>]/g,'-');
  var cl=function(v){return(v||'').replace(/,/g,'，');};
  var envH='部門,作業流程,作業名稱,作業步驟,狀態,時間,工作者,設備/工具,工程控制,管理控制,個人防護具,共通使用物料,評估物料/化學品,類別,環境衝擊,說明,環境OFQ,環境Impact,環境機率(P),重大環境考量面(SEA),環境風險等級,風險處理措施,對策說明\n';
  var riskH='部門,作業流程,作業名稱,作業步驟,狀態,時間,工作者,設備/工具,工程控制,管理控制,個人防護具,共通使用物料,評估物料/化學品,風險類別,危害因子說明,安衛後果影響,作業頻率(OF),發生機率(P),可能發生性(FR),後果嚴重程度(S),安全風險等級,改善機會(Y/N),風險處理措施,目標方案說明\n';
  var eC='',rC='';
  S.data.forEach(function(proc){
    proc.steps.forEach(function(step){
      var base=[proc.department,proc.processName,proc.taskName,step.stepName,
        step.common.status,step.common.time,step.common.worker,
        cl(step.common.equipment),cl(step.common.engControl),cl(step.common.adminControl),cl(step.common.ppe),cl(step.common.material)].join(',');
      step.envs.forEach(function(e){eC+=base+','+[cl(e.material),cl(e.category),cl(e.envImpact),cl(e.desc),e.ofq,e.impact,e.prob,e.sea,e.riskLevel,cl(e.treatment),cl(e.controlPlan)].join(',')+'\n';});
      step.safeties.forEach(function(s){rC+=base+','+[cl(s.material),cl(s.riskCat),cl(s.hazardDesc),cl(s.consequence),s.of,s.p,s.fr,s.s,s.classLevel,s.isOpp||'',cl(s.treatment),cl(s.treatmentDesc)].join(',')+'\n';});
    });
  });
  dlCSV('\uFEFF'+envH+eC,year+'_Env_'+plant+'_'+dept+'.csv');
  setTimeout(function(){dlCSV('\uFEFF'+riskH+rC,year+'_Risk_'+plant+'_'+dept+'.csv');},600);
  toast('已匯出 '+year+'_Env/Risk_'+plant+'_'+dept+'.csv','ok');
}
function dlCSV(content,fn){
  var a=document.createElement('a');
  a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(content);
  a.download=fn; document.body.appendChild(a); a.click(); document.body.removeChild(a);
}

// ============================================================
// PROC CRUD
// ============================================================
function createProc(){
  var p={id:uid(),department:getDept()||'部門',processName:'新作業流程',taskName:'新作業名稱',
    steps:[{id:uid(),stepName:'步驟一',common:{status:'例行',time:'現在',worker:'漢民員工',equipment:'',material:'',engControl:'',adminControl:'',ppe:''},envs:[],safeties:[]}]};
  S.data.push(p); S.selectedId=p.id; renderAll();
}
function deleteProc(id){
  if(!confirm('確定要刪除這個作業流程嗎？')) return;
  S.data=S.data.filter(function(p){return p.id!==id;});
  S.selectedId=S.data[0]?S.data[0].id:null;
  renderAll(); toast('已刪除作業流程','info');
}
function getProc(id){for(var i=0;i<S.data.length;i++){if(S.data[i].id===id)return S.data[i];}return null;}
function selectProc(id){S.selectedId=id; renderSidebar(); renderEntry();}
function switchTab(t){
  S.tab=t;
  document.getElementById('tab-entry').className='nav-tab'+(t==='entry'?' active':'');
  document.getElementById('tab-dash').className='nav-tab'+(t==='dash'?' active':'');
  document.getElementById('entry-view').style.display=t==='entry'?'block':'none';
  document.getElementById('dash-view').style.display=t==='dash'?'block':'none';
  document.getElementById('empty-state').style.display=(t==='entry'&&!S.selectedId)?'flex':'none';
  if(t==='dash') renderDash();
}
function updateHint(){
  document.getElementById('export-hint').textContent=getYear()+'_Env/Risk_'+getPlant()+'_'+(S.data[0]?S.data[0].department:'Dept')+'.csv';
}

// ============================================================
// RENDER ALL
// ============================================================
function renderAll(){renderSidebar();if(S.tab==='entry')renderEntry();else renderDash();updateNavCount();}
function updateNavCount(){
  document.getElementById('nav-count').textContent='已載入 '+S.data.length+' 筆流程';
  document.getElementById('proc-count').textContent=S.data.length+' 筆';
  updateHint();
}

// ============================================================
// SIDEBAR
// ============================================================
function renderSidebar(){
  var list=document.getElementById('proc-list');
  if(!S.data.length){list.innerHTML='<div style="text-align:center;padding:20px 12px;color:#334155;font-size:11px;border:1px dashed #334155;border-radius:6px;margin:0 12px">請匯入 CSV 或載入範例資料</div>';return;}
  list.innerHTML=S.data.map(function(p){
    return '<div class="proc-item'+(p.id===S.selectedId?' active':'')+'" data-action="select-proc" data-pid="'+p.id+'">'
      +'<div class="proc-name">'+esc(p.processName)+'</div>'
      +'<div class="proc-sub">▶ '+esc(p.taskName)+'</div>'
      +'<div class="proc-tags"><span class="ptag ptag-d">'+esc(p.department)+'</span><span class="ptag ptag-s">'+p.steps.length+' 步驟</span></div>'
      +'</div>';
  }).join('');
}

// ============================================================
// ENTRY VIEW
// ============================================================
function renderEntry(){
  var ev=document.getElementById('entry-view'), es=document.getElementById('empty-state');
  if(!S.selectedId||!S.data.length){ev.style.display='none';es.style.display='flex';return;}
  ev.style.display='block';es.style.display='none';
  var proc=getProc(S.selectedId); if(!proc){ev.innerHTML='';return;}
  ev.innerHTML=renderProcHTML(proc);
}

function renderProcHTML(proc){
  var tE=0,tS=0; proc.steps.forEach(function(s){tE+=s.envs.length;tS+=s.safeties.length;});
  var year=getYear(),plant=getPlant();
  return '<div class="proc-header-card">'
    +'<div class="proc-meta">'
    +'<span class="meta-badge badge-year">'+esc(year)+'</span>'
    +'<span class="meta-badge badge-plant">'+esc(plant)+'</span>'
    +'<input class="badge-dept-input" value="'+esc(proc.department)+'" data-action="proc-field" data-pid="'+proc.id+'" data-field="department"/>'
    +'</div>'
    +'<input class="edit-h1" value="'+esc(proc.processName)+'" placeholder="作業流程名稱（來自A表）..." data-action="proc-field" data-pid="'+proc.id+'" data-field="processName"/>'
    +'<input class="edit-h2" value="'+esc(proc.taskName)+'" placeholder="作業名稱（來自A表）..." data-action="proc-field" data-pid="'+proc.id+'" data-field="taskName"/>'
    +'<div class="proc-header-foot">'
    +'<div class="proc-stats"><span>📌 <strong>'+proc.steps.length+'</strong> 步驟</span><span>🌿 環境 <strong>'+tE+'</strong> 筆</span><span>🛡️ 安衛 <strong>'+tS+'</strong> 筆</span></div>'
    +'<button class="btn btn-danger" data-action="delete-proc" data-pid="'+proc.id+'">🗑️ 刪除流程</button>'
    +'</div></div>'
    +proc.steps.map(function(s,i){return renderStepHTML(proc.id,s,i);}).join('')
    +'<button class="add-step-btn" data-action="add-step" data-pid="'+proc.id+'">＋ 新增作業步驟（來自 A 表）</button>';
}

function renderStepHTML(pid,step,idx){
  var mats=parseMats(step.common.material);
  return '<div class="step-card" id="step-'+step.id+'">'
    +'<div class="step-head">'
    +'<div class="step-num">'+(idx+1)+'</div>'
    +'<input class="step-name-input" value="'+esc(step.stepName)+'" placeholder="步驟名稱（來自A表）..." data-action="set-step-name" data-pid="'+pid+'" data-sid="'+step.id+'"/>'
    +'<button class="btn" style="background:rgba(220,38,38,.15);color:#fca5a5;border:1px solid rgba(220,38,38,.2);margin-left:auto;flex-shrink:0" data-action="delete-step" data-pid="'+pid+'" data-sid="'+step.id+'">🗑️ 刪除</button>'
    +'</div>'
    +'<div style="padding:0 0 14px">'
    +renderCommonBlock(pid,step)
    +renderEnvBlock(pid,step,mats)
    +renderSafBlock(pid,step,mats)
    +'</div></div>';
}

function renderCommonBlock(pid,step){
  var c=step.common;
  return '<div class="block" style="border-color:#fed7aa">'
    +'<div class="block-head" style="background:linear-gradient(135deg,#ed7d31,#f97316)"><span>📋 基本資訊 &amp; 作業管制（共通欄位）</span></div>'
    +'<div class="block-body">'
    +'<div class="form-grid g4">'
    +fSelect('狀態',['例行','非例行','緊急'],c.status,'set-common',pid,step.id,'status')
    +fSelect('時間',['現在','過去','未來'],c.time,'set-common',pid,step.id,'time')
    +fSelect('工作者',['漢民員工','承攬商','共同作業','派遣'],c.worker,'set-common',pid,step.id,'worker')
    +fInput('設備 / 工具',c.equipment,'設備...','set-common',pid,step.id,'equipment')
    +'</div>'
    +'<div style="margin-bottom:10px"><label class="form-label" style="color:#ea580c">⚗️ 使用物料 / 化學品 <span style="font-weight:400;color:#94a3b8">(逗號分隔，連動下方)</span></label>'
    +'<input class="form-input" style="border-color:#fed7aa;background:#fffbeb" value="'+esc(c.material)+'" placeholder="例：冷媒、異丙醇（若無填「無」）" data-action="set-common-mat" data-pid="'+pid+'" data-sid="'+step.id+'"/></div>'
    +'<div class="form-grid g3" style="padding-top:10px;border-top:1px solid #f1f5f9;margin-bottom:0">'
    +fInput('工程控制',c.engControl,'例：局部排氣','set-common',pid,step.id,'engControl')
    +fInput('管理控制 / SOP',c.adminControl,'例：動火許可','set-common',pid,step.id,'adminControl')
    +fInput('個人防護具 (PPE)',c.ppe,'例：護目鏡','set-common',pid,step.id,'ppe')
    +'</div></div></div>';
}

function fInput(label,val,ph,action,pid,sid,field){
  return '<div><label class="form-label">'+esc(label)+'</label>'
    +'<input class="form-input" value="'+esc(val)+'" placeholder="'+esc(ph)+'" data-action="'+action+'" data-pid="'+pid+'" data-sid="'+sid+'" data-field="'+field+'"/></div>';
}
function fSelect(label,list,sel,action,pid,sid,field){
  return '<div><label class="form-label">'+esc(label)+'</label>'
    +'<select class="form-select" data-action="'+action+'" data-pid="'+pid+'" data-sid="'+sid+'" data-field="'+field+'">'+opts(list,sel)+'</select></div>';
}

function renderEnvBlock(pid,step,mats){
  var rows=step.envs.map(function(e){return renderEnvRow(pid,step.id,e,mats);}).join('');
  return '<div class="block" style="border-color:#bbf7d0">'
    +'<div class="block-head" style="background:linear-gradient(135deg,#15803d,#16a34a)"><span>🌿 環境考量面鑑別</span><span style="font-size:11px;background:rgba(255,255,255,.2);padding:2px 8px;border-radius:12px">'+step.envs.length+' 筆</span></div>'
    +'<div class="block-body"><div id="env-rows-'+step.id+'">'+rows+'</div>'
    +'<button class="add-row-btn add-env-btn" data-action="add-env" data-pid="'+pid+'" data-sid="'+step.id+'">＋ 新增環境考量面評估</button>'
    +'</div></div>';
}

function renderEnvRow(pid,sid,e,mats){
  var matOpts=[].concat(mats); if(e.material&&matOpts.indexOf(e.material)<0) matOpts.unshift(e.material);
  var impacts=ENV_CAT[e.category]||[];
  var isHigh=['S1','S2'].includes(e.sea)||['H','M'].includes(e.riskLevel);
  var seaCls=['S1','S2'].includes(e.sea)?'b-H':'b-sea';
  var rCls=e.riskLevel==='H'?'b-H':e.riskLevel==='M'?'b-M':'b-L';
  return '<div class="env-row" id="env-'+e.id+'">'
    +'<button class="del-btn" data-action="delete-env" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'">✕</button>'
    +'<div class="form-grid g3" style="margin-bottom:8px;padding-right:28px">'
    +'<div><label class="form-label" style="color:#16a34a">物料（連動自上方）</label>'
    +'<select class="form-select" style="border-color:#86efac;background:#f0fdf4" data-action="env-field" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="material">'
    +'<option value="">請選擇...</option>'+matOpts.map(function(m){return'<option'+(m===e.material?' selected':'')+'>'+esc(m)+'</option>';}).join('')+'</select></div>'
    +'<div><label class="form-label">類別</label>'
    +'<select class="form-select" data-action="env-cat" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'">'
    +Object.keys(ENV_CAT).map(function(k){return'<option'+(k===e.category?' selected':'')+'>'+k+'</option>';}).join('')+'</select></div>'
    +'<div><label class="form-label">環境衝擊</label>'
    +'<select class="form-select" data-action="env-field" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="envImpact">'
    +impacts.map(function(m){return'<option'+(m===e.envImpact?' selected':'')+'>'+esc(m)+'</option>';}).join('')+'</select></div></div>'
    +'<div style="margin-bottom:8px;padding-right:28px"><label class="form-label">說明（來源 / 行為 / 衝擊）</label>'
    +'<textarea class="form-textarea" data-action="env-field" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="desc" placeholder="詳細說明環境衝擊來源...">'+esc(e.desc)+'</textarea></div>'
    +'<div class="form-grid g3 score-box env-score" style="margin-bottom:8px">'
    +'<div><label class="form-label">頻率 (OFQ)</label>'
    +'<select class="form-select" data-action="env-calc" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="ofq">'
    +[['1','1-持續'],['2','2-經常'],['3','3-偶而'],['4','4-很少']].map(function(v){return'<option value="'+v[0]+'"'+(v[0]==e.ofq?' selected':'')+'>'+v[1]+'</option>';}).join('')+'</select></div>'
    +'<div><label class="form-label">衝擊程度 (I)</label>'
    +'<select class="form-select" data-action="env-calc" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="impact">'
    +[['A','A-嚴重'],['B','B-中度'],['C','C-輕微']].map(function(v){return'<option value="'+v[0]+'"'+(v[0]===e.impact?' selected':'')+'>'+v[1]+'</option>';}).join('')+'</select></div>'
    +'<div><label class="form-label">機率 (P)</label>'
    +'<select class="form-select" data-action="env-calc" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="prob">'
    +[['1','1-高度'],['2','2-中度'],['3','3-低度']].map(function(v){return'<option value="'+v[0]+'"'+(v[0]==e.prob?' selected':'')+'>'+v[1]+'</option>';}).join('')+'</select></div></div>'
    +'<div class="result-row"><div class="viol-row">'
    +'<label class="viol-label"><input type="checkbox"'+(e.v1?' checked':'')+' data-action="env-vio" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="v1"/>近一年異常(強制H)</label>'
    +'<label class="viol-label"><input type="checkbox"'+(e.v2?' checked':'')+' data-action="env-vio" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="v2"/>近三年裁罰(強制H)</label>'
    +'<label class="viol-label"><input type="checkbox"'+(e.v3?' checked':'')+' data-action="env-vio" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="v3"/>不符法規(強制H)</label>'
    +'</div><div class="result-badges"><span class="badge '+seaCls+'">SEA: '+e.sea+'</span><span class="badge '+rCls+'">風險: '+e.riskLevel+'</span></div></div>'
    +(isHigh?'<div class="treat-box treat-orange"><label class="treat-label">⚠️ 請選擇風險處理措施（M/H 或 SEA S1/S2 觸發）</label>'
      +'<select class="form-select" style="margin-bottom:8px" data-action="env-field" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="treatment">'
      +'<option value="">請選擇...</option>'+opts(['建立作業管制','建立行動措施','建立管理方案'],e.treatment)+'</select>'
      +'<label class="treat-label">對策說明</label>'
      +'<input class="form-input" value="'+esc(e.controlPlan)+'" placeholder="請簡述建立的作業管制、行動措施或管理方案..." data-action="env-field" data-pid="'+pid+'" data-sid="'+sid+'" data-eid="'+e.id+'" data-field="controlPlan"/></div>':'')
    +'</div>';
}

function renderSafBlock(pid,step,mats){
  var rows=step.safeties.map(function(s){return renderSafRow(pid,step.id,s,mats);}).join('');
  return '<div class="block" style="border-color:#bfdbfe">'
    +'<div class="block-head" style="background:linear-gradient(135deg,#1d4ed8,#2563eb)"><span>🛡️ 安全衛生風險登錄</span><span style="font-size:11px;background:rgba(255,255,255,.2);padding:2px 8px;border-radius:12px">'+step.safeties.length+' 筆</span></div>'
    +'<div class="block-body"><div id="saf-rows-'+step.id+'">'+rows+'</div>'
    +'<button class="add-row-btn add-saf-btn" data-action="add-saf" data-pid="'+pid+'" data-sid="'+step.id+'">＋ 新增安全衛生風險評估</button>'
    +'</div></div>';
}

function renderSafRow(pid,sid,s,mats){
  var matOpts=[].concat(mats); if(s.material&&matOpts.indexOf(s.material)<0) matOpts.unshift(s.material);
  var impacts=SAF_CAT[s.riskCat]||[];
  var isC12=['Class 1','Class 2'].includes(s.classLevel);
  var isC3=s.classLevel==='Class 3';
  var cls=s.classLevel==='Class 1'?'b-C1':s.classLevel==='Class 2'?'b-C2':'b-C3';
  return '<div class="saf-row" id="saf-'+s.id+'">'
    +'<button class="del-btn" data-action="delete-saf" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'">✕</button>'
    +'<div class="form-grid g3" style="margin-bottom:8px;padding-right:28px">'
    +'<div><label class="form-label" style="color:#2563eb">物料（連動自上方）</label>'
    +'<select class="form-select" style="border-color:#93c5fd;background:#eff6ff" data-action="saf-field" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="material">'
    +'<option value="">請選擇...</option>'+matOpts.map(function(m){return'<option'+(m===s.material?' selected':'')+'>'+esc(m)+'</option>';}).join('')+'</select></div>'
    +'<div><label class="form-label">風險類別</label>'
    +'<select class="form-select" data-action="saf-cat" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'">'
    +Object.keys(SAF_CAT).map(function(k){return'<option'+(k===s.riskCat?' selected':'')+'>'+k+'</option>';}).join('')+'</select></div>'
    +'<div><label class="form-label">安衛後果影響</label>'
    +'<select class="form-select" data-action="saf-field" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="consequence">'
    +impacts.map(function(m){return'<option'+(m===s.consequence?' selected':'')+'>'+esc(m)+'</option>';}).join('')+'</select></div></div>'
    +'<div style="margin-bottom:8px;padding-right:28px"><label class="form-label">危害因子說明</label>'
    +'<textarea class="form-textarea" data-action="saf-field" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="hazardDesc" placeholder="詳細說明危害因子...">'+esc(s.hazardDesc)+'</textarea></div>'
    +'<div class="form-grid g3 score-box saf-score" style="margin-bottom:8px">'
    +'<div><label class="form-label">頻率 (OF)</label>'
    +'<select class="form-select" data-action="saf-calc" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="of">'
    +[['10','10-持續'],['8','8-經常'],['5','5-偶而'],['1','1-少有']].map(function(v){return'<option value="'+v[0]+'"'+(v[0]==s.of?' selected':'')+'>'+v[1]+'</option>';}).join('')+'</select></div>'
    +'<div><label class="form-label">機率 (P)</label>'
    +'<select class="form-select" data-action="saf-calc" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="p">'
    +[['10','10-高度'],['5','5-中度'],['1','1-低度']].map(function(v){return'<option value="'+v[0]+'"'+(v[0]==s.p?' selected':'')+'>'+v[1]+'</option>';}).join('')+'</select></div>'
    +'<div><label class="form-label">嚴重度 (S)</label>'
    +'<select class="form-select" data-action="saf-calc" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="s">'
    +[['A','A-死亡/罰單'],['B','B-住院'],['C','C-就醫'],['D','D-限制工作'],['E','E-護理']].map(function(v){return'<option value="'+v[0]+'"'+(v[0]===s.s?' selected':'')+'>'+v[1]+'</option>';}).join('')+'</select></div></div>'
    +'<div class="result-row"><div class="viol-row">'
    +'<label class="viol-label"><input type="checkbox"'+(s.v1?' checked':'')+' data-action="saf-vio" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="v1"/>近一年裁罰或職災(強制C1)</label>'
    +'<label class="viol-label"><input type="checkbox"'+(s.v2?' checked':'')+' data-action="saf-vio" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="v2"/>近三年永久失能(強制C1)</label>'
    +'<label class="viol-label"><input type="checkbox"'+(s.v3?' checked':'')+' data-action="saf-vio" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="v3"/>不符合法規(強制C1)</label>'
    +'</div><div class="result-badges">'
    +'<span class="badge" style="background:#f8fafc;color:#64748b;border-color:#e2e8f0">FR: '+s.fr+'</span>'
    +'<span class="badge '+cls+'">'+s.classLevel+'</span></div></div>'
    +(isC3?'<div class="treat-box treat-blue"><label class="treat-label">💡 請選擇是否為改善機會（Class 3 觸發）</label>'
      +'<select class="form-select" data-action="saf-field" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="isOpp">'
      +'<option value="">請選擇 Y / N</option><option value="Y"'+(s.isOpp==='Y'?' selected':'')+'>Y（是改善機會）</option><option value="N"'+(s.isOpp==='N'?' selected':'')+'>N（非）</option></select></div>':'')
    +(isC12?'<div class="treat-box treat-red"><label class="treat-label">⚠️ 請選擇風險處理措施（'+esc(s.classLevel)+' 觸發）</label>'
      +'<select class="form-select" style="margin-bottom:8px" data-action="saf-field" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="treatment">'
      +'<option value="">請選擇處置方案</option>'+opts(['建立目標方案','納入現有過程管理，維持零工傷'],s.treatment)+'</select>'
      +'<label class="treat-label">說明</label>'
      +'<input class="form-input" value="'+esc(s.treatmentDesc)+'" placeholder="請簡述建立的目標方案..." data-action="saf-field" data-pid="'+pid+'" data-sid="'+sid+'" data-sfid="'+s.id+'" data-field="treatmentDesc"/></div>':'')
    +'</div>';
}

// ============================================================
// MUTATIONS — all re-render targeted rows
// ============================================================
function setProcField(pid,field,val){var p=getProc(pid);if(p){p[field]=val;}renderSidebar();updateHint();}
function setCommon(pid,sid,field,val){var p=getProc(pid);var s=p&&p.steps.find(function(x){return x.id===sid;});if(s)s.common[field]=val;}
function setCommonMat(pid,sid,val){var p=getProc(pid);var s=p&&p.steps.find(function(x){return x.id===sid;});if(s){s.common.material=val;renderEntry();}}
function setStepName(pid,sid,val){var p=getProc(pid);var s=p&&p.steps.find(function(x){return x.id===sid;});if(s)s.stepName=val;}

function getEnv(pid,sid,eid){var p=getProc(pid);var s=p&&p.steps.find(function(x){return x.id===sid;});return s&&s.envs.find(function(x){return x.id===eid;});}
function getSaf(pid,sid,sfid){var p=getProc(pid);var s=p&&p.steps.find(function(x){return x.id===sid;});return s&&s.safeties.find(function(x){return x.id===sfid;});}

function setEnvField(pid,sid,eid,field,val){var e=getEnv(pid,sid,eid);if(e)e[field]=val;}
function setEnvCat(pid,sid,eid,val){var e=getEnv(pid,sid,eid);if(!e)return;e.category=val;e.envImpact=(ENV_CAT[val]||[])[0]||'';reRenderEnvRow(pid,sid,eid);}
function setEnvCalc(pid,sid,eid,field,val){
  var e=getEnv(pid,sid,eid);if(!e)return;
  e[field]=field==='ofq'||field==='prob'?parseInt(val):val;
  e.sea=calcEnvSEA(e.ofq,e.impact);
  e.riskLevel=calcEnvRisk(e.sea,e.prob,e.v1||e.v2||e.v3);
  reRenderEnvRow(pid,sid,eid);
}
function setEnvVio(pid,sid,eid,field,val){
  var e=getEnv(pid,sid,eid);if(!e)return;
  e[field]=val; e.riskLevel=calcEnvRisk(e.sea,e.prob,e.v1||e.v2||e.v3);
  reRenderEnvRow(pid,sid,eid);
}
function setSafField(pid,sid,sfid,field,val){var s=getSaf(pid,sid,sfid);if(s)s[field]=val;}
function setSafCat(pid,sid,sfid,val){var s=getSaf(pid,sid,sfid);if(!s)return;s.riskCat=val;s.consequence=(SAF_CAT[val]||[])[0]||'';reRenderSafRow(pid,sid,sfid);}
function setSafCalc(pid,sid,sfid,field,val){
  var s=getSaf(pid,sid,sfid);if(!s)return;
  s[field]=field==='s'?val:parseInt(val);
  s.fr=s.of*s.p; s.classLevel=calcSafClass(s.fr,s.s,s.v1||s.v2||s.v3);
  reRenderSafRow(pid,sid,sfid);
}
function setSafVio(pid,sid,sfid,field,val){
  var s=getSaf(pid,sid,sfid);if(!s)return;
  s[field]=val; s.classLevel=calcSafClass(s.fr,s.s,s.v1||s.v2||s.v3);
  reRenderSafRow(pid,sid,sfid);
}

function reRenderEnvRow(pid,sid,eid){
  var el=document.getElementById('env-'+eid);if(!el)return;
  var p=getProc(pid),s=p&&p.steps.find(function(x){return x.id===sid;}),e=s&&s.envs.find(function(x){return x.id===eid;});
  if(!e)return;
  var mats=parseMats(s.common.material);
  var tmp=document.createElement('div');tmp.innerHTML=renderEnvRow(pid,sid,e,mats);
  el.replaceWith(tmp.firstElementChild);
}
function reRenderSafRow(pid,sid,sfid){
  var el=document.getElementById('saf-'+sfid);if(!el)return;
  var p=getProc(pid),s=p&&p.steps.find(function(x){return x.id===sid;}),sf=s&&s.safeties.find(function(x){return x.id===sfid;});
  if(!sf)return;
  var mats=parseMats(s.common.material);
  var tmp=document.createElement('div');tmp.innerHTML=renderSafRow(pid,sid,sf,mats);
  el.replaceWith(tmp.firstElementChild);
}

// Add/Delete
function addStep(pid){var p=getProc(pid);if(!p)return;p.steps.push({id:uid(),stepName:'新步驟',common:{status:'例行',time:'現在',worker:'漢民員工',equipment:'',material:'',engControl:'',adminControl:'',ppe:''},envs:[],safeties:[]});renderEntry();}
function deleteStep(pid,sid){var p=getProc(pid);if(!p)return;p.steps=p.steps.filter(function(s){return s.id!==sid;});renderEntry();}
function addEnv(pid,sid){var p=getProc(pid),s=p&&p.steps.find(function(x){return x.id===sid;});if(!s)return;s.envs.push({id:uid(),material:'',category:'環保_空',envImpact:'酸排',desc:'',ofq:4,impact:'C',prob:3,sea:'S4',riskLevel:'L',v1:false,v2:false,v3:false,treatment:'',controlPlan:''});renderEntry();}
function deleteEnv(pid,sid,eid){var p=getProc(pid),s=p&&p.steps.find(function(x){return x.id===sid;});if(!s)return;s.envs=s.envs.filter(function(e){return e.id!==eid;});renderEntry();}
function addSaf(pid,sid){var p=getProc(pid),s=p&&p.steps.find(function(x){return x.id===sid;});if(!s)return;s.safeties.push({id:uid(),material:'',riskCat:'物理性',hazardDesc:'',consequence:'墜落,滾落',of:1,p:1,fr:1,s:'E',classLevel:'Class 4',v1:false,v2:false,v3:false,isOpp:'',treatment:'',treatmentDesc:''});renderEntry();}
function deleteSaf(pid,sid,sfid){var p=getProc(pid),s=p&&p.steps.find(function(x){return x.id===sid;});if(!s)return;s.safeties=s.safeties.filter(function(x){return x.id!==sfid;});renderEntry();}

// ============================================================
// DASHBOARD
// ============================================================
function renderDash(){
  var c1=0,c2=0,eH=0,eM=0;
  var tS={'Class 1':{'建立目標方案':0,'納入現有過程管理，維持零工傷':0,'未填寫':0},'Class 2':{'建立目標方案':0,'納入現有過程管理，維持零工傷':0,'未填寫':0},'H':{'建立作業管制':0,'建立行動措施':0,'建立管理方案':0,'未填寫':0},'M':{'建立作業管制':0,'建立行動措施':0,'建立管理方案':0,'未填寫':0}};
  var hi=[];
  S.data.forEach(function(proc){proc.steps.forEach(function(step){
    step.envs.forEach(function(e){if(['H','M'].includes(e.riskLevel)){if(e.riskLevel==='H')eH++;else eM++;var t=e.treatment||'未填寫';if(tS[e.riskLevel][t]!==undefined)tS[e.riskLevel][t]++;hi.push({dept:proc.department,proc:proc.processName,task:proc.taskName,step:step.stepName,mat:e.material,type:'Env',level:e.riskLevel,treatment:e.treatment,desc:e.controlPlan});}});
    step.safeties.forEach(function(s){if(['Class 1','Class 2'].includes(s.classLevel)){if(s.classLevel==='Class 1')c1++;else c2++;var t=s.treatment||'未填寫';if(tS[s.classLevel][t]!==undefined)tS[s.classLevel][t]++;hi.push({dept:proc.department,proc:proc.processName,task:proc.taskName,step:step.stepName,mat:s.material,type:'Safety',level:s.classLevel,fr:s.fr,treatment:s.treatment,desc:s.treatmentDesc});}});
  });});
  var sc=function(lbl,sub,val,col){return'<div class="stat-card" style="border-top-color:'+col+'"><div style="font-size:10px;font-weight:700;color:#94a3b8;margin-bottom:2px">'+lbl+'</div><div style="font-size:13px;font-weight:900;color:'+col+';margin-bottom:6px">'+sub+'</div><div style="font-size:38px;font-weight:900;color:#0f172a">'+val+'</div></div>';};
  var tBox=function(key,col,bg,bd,items){return'<div class="treat-stat-box" style="background:'+bg+';border:1px solid '+bd+'"><div style="font-weight:800;font-size:11px;color:'+col+';margin-bottom:8px;padding-bottom:7px;border-bottom:1px solid '+bd+'">'+key+' 處置統計</div>'+items.map(function(item){return'<div style="display:flex;justify-content:space-between;margin-bottom:5px;font-size:11px"><span style="color:'+(item==='未填寫'?'#ef4444':'#475569')+';font-weight:'+(item==='未填寫'?700:400)+'">'+(item.length>12?item.slice(0,12)+'…':item)+'</span><span style="font-weight:700;color:'+(item==='未填寫'&&tS[key][item]>0?'#ef4444':'#0f172a')+'">'+tS[key][item]+' 件</span></div>';}).join('')+'</div>';};
  document.getElementById('dash-view').innerHTML='<div style="max-width:1400px">'
    +'<div style="margin-bottom:20px"><h1 style="font-weight:900;font-size:20px;color:#0f172a;margin-bottom:3px">全廠 EHS 統計儀表板</h1><p style="font-size:12px;color:#64748b">共 '+S.data.length+' 筆流程</p></div>'
    +'<div class="dash-grid4">'+sc('重大安全風險','Class 1',c1,'#dc2626')+sc('重大安全風險','Class 2',c2,'#ea580c')+sc('環境高風險','Level H',eH,'#be123c')+sc('環境中風險','Level M',eM,'#d97706')+'</div>'
    +'<div style="background:#fff;border-radius:10px;border:1px solid #e2e8f0;padding:18px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,.04)">'
    +'<h2 style="font-weight:800;font-size:14px;color:#0f172a;margin-bottom:14px">🎯 風險處理措施落實統計</h2>'
    +'<div class="treat-grid">'
    +tBox('Class 1','#dc2626','#fef2f2','#fecaca',['建立目標方案','納入現有過程管理，維持零工傷','未填寫'])
    +tBox('Class 2','#ea580c','#fff7ed','#fed7aa',['建立目標方案','納入現有過程管理，維持零工傷','未填寫'])
    +tBox('H','#be123c','#fff1f2','#fecdd3',['建立作業管制','建立行動措施','建立管理方案','未填寫'])
    +tBox('M','#d97706','#fffbeb','#fde68a',['建立作業管制','建立行動措施','建立管理方案','未填寫'])
    +'</div></div>'
    +'<div style="background:#fff;border-radius:10px;border:1px solid #e2e8f0;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.04)">'
    +'<div style="padding:12px 18px;border-bottom:1px solid #f1f5f9;background:#f8fafc"><h2 style="font-weight:800;font-size:14px;color:#0f172a;margin:0">📋 優先管控作業清單（'+hi.length+' 件）</h2></div>'
    +(hi.length===0?'<div style="padding:32px;text-align:center;color:#94a3b8">目前無高風險管控項目</div>'
      :'<div style="overflow-x:auto"><table><thead><tr><th>部門</th><th>作業流程 / 步驟</th><th>物料 / 類型</th><th>風險等級</th><th>處置方案</th></tr></thead><tbody>'
      +hi.map(function(item){return'<tr><td><strong style="color:#ea580c">'+esc(item.dept)+'</strong></td><td><div style="font-weight:700;color:#0f172a">'+esc(item.proc)+'</div><div style="font-size:10px;color:#64748b;margin-top:2px">'+esc(item.task)+' → '+esc(item.step)+'</div></td><td><div style="font-weight:600">'+esc(item.mat||'—')+'</div><span style="font-size:9px;padding:1px 6px;border-radius:3px;font-weight:700;background:'+(item.type==='Env'?'#f0fdf4':'#eff6ff')+';color:'+(item.type==='Env'?'#16a34a':'#2563eb')+';display:inline-block;margin-top:3px">'+(item.type==='Env'?'環境':'安全衛生')+'</span></td><td><span class="badge '+((['Class 1','H'].includes(item.level))?'b-H':'b-M')+'">'+esc(item.level)+(item.fr?' (FR:'+item.fr+')':'')+'</span></td><td>'+(item.treatment?'<div style="font-size:12px;font-weight:700;color:#0f172a">'+esc(item.treatment)+'</div>'+(item.desc?'<div style="font-size:10px;color:#64748b;margin-top:2px">'+esc(item.desc)+'</div>':''):'<span style="font-size:11px;color:#d1d5db;font-style:italic">尚未填寫</span>')+'</td></tr>';}).join('')
      +'</tbody></table></div>')
    +'</div></div>';
}

// ============================================================
// FORMAT MODAL
// ============================================================
function showFormatModal(){
  document.getElementById('modal-box').innerHTML='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:18px">'
    +'<h2 style="font-weight:900;font-size:17px;color:#0f172a">📄 匯入 CSV 資料格式說明</h2>'
    +'<button class="btn btn-ghost" data-action="close-modal" style="font-size:12px">✕ 關閉</button></div>'
    +'<div style="padding:12px;background:#fef9f0;border:1px solid #fed7aa;border-radius:8px;margin-bottom:16px;font-size:12px">'
    +'<strong style="color:#92400e">⚠️ 重要說明</strong><br>'
    +'<ul style="margin:6px 0 0 16px;color:#78350f;line-height:1.9">'
    +'<li>請分別匯入兩個檔案：<strong>環境考量面（Env）</strong> 與 <strong>安全衛生風險（Risk）</strong></li>'
    +'<li>建議檔名：<code style="background:#fef3c7;padding:1px 4px;border-radius:3px">2025_Env_ZY_FAC.csv</code></li>'
    +'<li>第一行為標題列，第二行起為資料</li>'
    +'<li>系統以 <strong>部門 + 作業流程 + 作業名稱</strong> 作為識別鍵</li>'
    +'<li>作業流程/作業/步驟名稱來源為 <strong>A表（流程總覽）</strong></li>'
    +'<li>☁️ SharePoint 直接讀取：需在 SP 內開啟此 HTML，點「登入 Microsoft」即可自動讀取資料夾</li>'
    +'</ul></div>'
    +'<h3 style="font-weight:800;font-size:13px;color:#0f172a;margin-bottom:8px">🌿 Env CSV — 欄位（共 23 欄）</h3>'
    +'<div style="overflow-x:auto;margin-bottom:16px"><table class="ft"><thead><tr><th>#</th><th>欄位</th><th>範例</th><th>說明</th></tr></thead><tbody>'
    +[['1','部門','FAC',''],['2','作業流程','空調水系統','A表'],['3','作業名稱','加藥系統','A表'],['4','作業步驟','藥液添加作業','A表'],['5','狀態','非例行','例行/非例行/緊急'],['6','時間','現在','現在/過去/未來'],['7','工作者','承攬商',''],['8','設備/工具','手工具',''],['9','工程控制','防液堤',''],['10','管理控制','SDS',''],['11','個人防護具','耐酸鹼手套',''],['12','共通使用物料','腐蝕結垢抑制劑','頓號或逗號分隔'],['13','評估物料','腐蝕結垢抑制劑','單一物料'],['14','類別','環保_廢棄物',''],['15','環境衝擊','一般事業廢棄物',''],['16','說明','藥劑空桶處理',''],['17','環境OFQ','4','1~4'],['18','環境Impact','C','A/B/C'],['19','環境機率(P)','3','1~3'],['20','SEA','S4','可留空自動計算'],['21','環境風險等級','L','可留空自動計算'],['22','風險處理措施','',''],['23','對策說明','','']].map(function(r){return'<tr><td style="color:#7c3aed;font-weight:700">'+r[0]+'</td><td style="font-weight:600">'+r[1]+'</td><td style="color:#059669">'+r[2]+'</td><td style="color:#64748b">'+r[3]+'</td></tr>';}).join('')
    +'</tbody></table></div>'
    +'<h3 style="font-weight:800;font-size:13px;color:#0f172a;margin-bottom:8px">🛡️ Risk CSV — 欄位（共 24 欄）</h3>'
    +'<div style="overflow-x:auto;margin-bottom:16px"><table class="ft"><thead><tr><th>#</th><th>欄位</th><th>範例</th><th>說明</th></tr></thead><tbody>'
    +[['1~12','（同Env前12欄）','','完全相同'],['13','評估物料','管路殺菌滅藻劑',''],['14','風險類別','化學性',''],['15','危害因子說明','添加時遭噴濺',''],['16','安衛後果影響','與有害物接觸',''],['17','作業頻率(OF)','5','10/8/5/1'],['18','發生機率(P)','1','10/5/1'],['19','可能發生性(FR)','5','OF×P'],['20','後果嚴重程度(S)','E','A/B/C/D/E'],['21','安全風險等級','Class 4','可留空自動計算'],['22','改善機會(Y/N)','','Y或N'],['23','風險處理措施','',''],['24','目標方案說明','','']].map(function(r){return'<tr><td style="color:#7c3aed;font-weight:700">'+r[0]+'</td><td style="font-weight:600">'+r[1]+'</td><td style="color:#059669">'+r[2]+'</td><td style="color:#64748b">'+r[3]+'</td></tr>';}).join('')
    +'</tbody></table></div>'
    +'<div style="padding:12px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;font-size:12px;color:#166534">💡 可直接使用本系統匯出的 CSV 作為下年度匯入來源。</div>';
  document.getElementById('modal-overlay').className='show';
}

// ============================================================
// INIT
// ============================================================
renderSidebar();
// ============================================================
// 方案 A：衝突偵測警告（Lock 機制）
// Lock 檔案存在 SharePoint: ISO_Risk_Env_Review/_ehs_lock.json
// ============================================================

var LOCK_FILE = '_ehs_lock.json';
var LOCK_EXPIRE_MIN = 30;
var lockCheckTimer = null;
var currentLockFile = null;  // 目前本機正在編輯的檔案名稱

// 取得鎖定狀態（從 SharePoint 讀取 _ehs_lock.json）
async function lockGet(){
  try{
    var sd = await getSiteDrive();
    var path = '/drives/'+sd.driveId+'/root:/'+SP_FOLDER+'/'+LOCK_FILE+':/content';
    var r = await fetch('https://graph.microsoft.com/v1.0'+path, {
      headers:{'Authorization':'Bearer '+graphToken}
    });
    if(r.status===404) return null;  // 沒有鎖定
    if(!r.ok) return null;
    var text = await r.text();
    return JSON.parse(text);
  }catch(e){ return null; }
}

// 寫入鎖定（覆蓋 _ehs_lock.json）
async function lockSet(fileName, userName){
  try{
    var sd = await getSiteDrive();
    var path = '/drives/'+sd.driveId+'/root:/'+SP_FOLDER+'/'+LOCK_FILE+':/content';
    var payload = JSON.stringify({
      file: fileName,
      user: userName,
      time: new Date().toISOString()
    });
    await fetch('https://graph.microsoft.com/v1.0'+path, {
      method:'PUT',
      headers:{'Authorization':'Bearer '+graphToken,'Content-Type':'application/json'},
      body: payload
    });
  }catch(e){ console.warn('Lock set failed:', e); }
}

// 釋放鎖定（刪除 _ehs_lock.json）
async function lockRelease(){
  try{
    var sd = await getSiteDrive();
    var path = '/drives/'+sd.driveId+'/root:/'+SP_FOLDER+'/'+LOCK_FILE+':/content';
    // 寫入空值表示釋放
    await fetch('https://graph.microsoft.com/v1.0'+path, {
      method:'PUT',
      headers:{'Authorization':'Bearer '+graphToken,'Content-Type':'application/json'},
      body: JSON.stringify({file:null,user:null,time:null})
    });
    currentLockFile = null;
    stopLockRefresh();
    hideLockBanner();
  }catch(e){ console.warn('Lock release failed:', e); }
}

// 嘗試取得鎖定（載入 CSV 後呼叫）
async function lockAcquire(fileName){
  var userName = document.getElementById('user-chip')
    ? document.getElementById('user-chip').textContent.replace('👤 ','').trim()
    : '使用者';

  var existing = await lockGet();

  // 已被他人鎖定且未過期
  if(existing && existing.file && existing.file === fileName && existing.user !== userName){
    var lockTime = new Date(existing.time);
    var now = new Date();
    var diffMin = (now - lockTime) / 60000;
    if(diffMin < LOCK_EXPIRE_MIN){
      var leftMin = Math.ceil(LOCK_EXPIRE_MIN - diffMin);
      showLockBanner(existing.user, existing.file, leftMin);
      return false;  // 無法取得鎖定
    }
  }

  // 可以取得鎖定
  await lockSet(fileName, userName);
  currentLockFile = fileName;
  hideLockBanner();
  startLockRefresh(fileName, userName);
  return true;
}

// 每 10 分鐘更新一次 lock 時間（防止過期）
function startLockRefresh(fileName, userName){
  stopLockRefresh();
  lockCheckTimer = setInterval(function(){
    lockSet(fileName, userName);
  }, 10 * 60 * 1000);
}

function stopLockRefresh(){
  if(lockCheckTimer){ clearInterval(lockCheckTimer); lockCheckTimer=null; }
}

// 顯示衝突警告橫幅
function showLockBanner(user, file, leftMin){
  var b = document.getElementById('lock-banner');
  if(!b){
    b = document.createElement('div');
    b.id = 'lock-banner';
    b.style.cssText = 'position:fixed;top:52px;left:0;right:0;z-index:500;background:#7c3aed;color:white;padding:10px 20px;display:flex;align-items:center;justify-content:space-between;font-size:13px;font-weight:700;box-shadow:0 2px 12px rgba(0,0,0,.2)';
    document.body.appendChild(b);
  }
  b.innerHTML = '<span>⚠️ <strong>'+esc(user)+'</strong> 正在編輯 <strong>'+esc(file)+'</strong>，建議等候（約剩 '+leftMin+' 分鐘）或以唯讀方式瀏覽</span>'
    +'<button data-action="force-unlock" style="background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.4);color:white;padding:4px 12px;border-radius:5px;font-size:11px;font-weight:700;cursor:pointer;font-family:inherit">強制解鎖</button>';
  b.style.display='flex';
  // push content down
  document.getElementById('main').style.marginTop='44px';
}

function hideLockBanner(){
  var b = document.getElementById('lock-banner');
  if(b) b.style.display='none';
  var m = document.getElementById('main');
  if(m) m.style.marginTop='';
}

// 強制解鎖（覆蓋別人的鎖定）
async function forceUnlock(){
  if(!confirm('確定要強制解鎖？這會中斷另一位使用者的編輯，請先確認對方已離開。')) return;
  await lockRelease();
  toast('已強制解鎖，現在可以開始編輯','info');
}

// 頁面關閉時自動釋放鎖定
window.addEventListener('beforeunload', function(){
  if(currentLockFile && graphToken){
    // 用 sendBeacon 送出（beforeunload 時 fetch 不可靠）
    var sd_driveId = _driveId;
    if(sd_driveId){
      var url = 'https://graph.microsoft.com/v1.0/drives/'+sd_driveId+'/root:/'+SP_FOLDER+'/'+LOCK_FILE+':/content';
      var payload = JSON.stringify({file:null,user:null,time:null});
      // sendBeacon 不支援自訂 header，改用同步 XHR
      var xhr = new XMLHttpRequest();
      xhr.open('PUT', url, false);
      xhr.setRequestHeader('Authorization','Bearer '+graphToken);
      xhr.setRequestHeader('Content-Type','application/json');
      try{ xhr.send(payload); }catch(e){}
    }
  }
});
