import {SCHEMA, KINDS, defaults, uid, loadBundle, saveBundle, validatePreset, validateBundle, mergeBundle, describe, withoutPreset, movePreset, SHAPE_SIZE_UNIT, displayLength, storedLength} from './core.js?v=0.4.1';
import {ExcelAdapter} from './office.js?v=0.4.1';
import {DemoAdapter} from './demo.js?v=0.4.1';
import {askConfirmation} from './confirmation.js?v=0.4.1';

const $=id=>document.getElementById(id);
const demo=new URLSearchParams(location.search).get('demo')==='1';
let adapter=null, connected=false, busy=false, bundle, storedError=null;
let editorKind=null, editingId=null, statusTimer;
const lengthUnit=SHAPE_SIZE_UNIT;
try{bundle=loadBundle(localStorage);}catch(error){bundle=defaults();storedError=error;}
const TITLES={shapeSize:'図形サイズ',color:'カラー',font:'フォント',fontSize:'文字サイズ',grid:'行高・列幅',border:'罫線',numberFormat:'表示形式'};
const FORM={
  shapeSize:[['width','幅（mm）','number',40],['height','高さ（mm）','number',14]],
  color:[['color','カラー（#RRGGBB）','text','#24876A']],
  font:[['name','フォント名','text','Yu Gothic']],
  fontSize:[['size','文字サイズ（pt）','number',11]],
  grid:[['rowHeight','行高（pt・空欄なら変更しない）','number',''],['columnWidth','列幅（pt・空欄なら変更しない）','number','']],
  border:[['pattern','罫線の位置','select','outer',{outer:'外枠',all:'格子',bottom:'下線',none:'罫線なし'}],['style','線種','select','Continuous',{Continuous:'実線',Dash:'破線',Dot:'点線',Double:'二重線'}],['weight','太さ','select','Thin',{Hairline:'極細',Thin:'細',Medium:'中',Thick:'太'}],['color','カラー（#RRGGBB）','text','#A3B3AA']],
  numberFormat:[['code','表示形式コード（例：#,##0）','text','#,##0']]
};
function el(tag,cls,content){const node=document.createElement(tag);if(cls)node.className=cls;if(content!==undefined)node.textContent=content;return node;}
function status(message,error=false){clearTimeout(statusTimer);$('status').textContent=message;$('status').className=error?'error':'';$('status').hidden=false;statusTimer=setTimeout(()=>$('status').hidden=true,error?15000:5000);}
async function action(fn){
  if(busy)return;
  busy=true;document.body.classList.add('busy');document.querySelectorAll('button').forEach(b=>b.disabled=true);
  try{await fn();}catch(error){status(`${error.message||'操作に失敗しました。'}${error.code?` (${error.code})`:''}`,true);}
  finally{busy=false;document.body.classList.remove('busy');document.querySelectorAll('button').forEach(b=>b.disabled=b.dataset.boundary==='true');}
}
function requireConnection(){if(!connected)throw new Error('Excelからアドインを開いてください。ブラウザーでの見た目確認は ?demo=1 を指定してください。');}
function options(){return {target:$('style-target').value,colorRole:$('color-role').value,resizeMode:$('resize-mode').value,shapes:{mode:'native'}};}
function ask(message){return askConfirmation(document,$('editor').open?$('editor-fields'):$('confirmation-host'),message);}
function commit(next){bundle=saveBundle(localStorage,next);renderPresets();}
function renderPresets(){
  for(const kind of KINDS){
    const container=$(`list-${kind}`);container.replaceChildren();
    const presets=bundle.presets.filter(p=>p.kind===kind);
    if(!presets.length)container.append(el('div','empty','「＋ 登録」から追加してください。'));
    for(const [index,p] of presets.entries()){
      const row=el('div','preset'), apply=el('button','apply-preset');apply.type='button';apply.title=`${p.name}を適用`;
      if(kind==='color'){const dot=el('span','color-dot');dot.style.backgroundColor=p.value.color;apply.append(dot);}
      const label=el('span','preset-text');label.append(el('span','preset-name',p.name),el('span','preset-desc',describe(p,lengthUnit)));apply.append(label);
      apply.addEventListener('click',()=>action(async()=>{requireConnection();const result=await adapter.apply(p,options());const target=result?.target==='shapes'?`図形「${result.names.join('、')}」`:'選択セル';status(`${demo?'デモ：':''}${target}に「${p.name}」を適用しました。`);}));
      const edit=el('button','preset-menu','編集');edit.type='button';edit.title=`${p.name}を編集・削除`;edit.setAttribute('aria-label',edit.title);edit.addEventListener('click',()=>openEditor(kind,p));
      const remove=el('button','preset-delete','削除');remove.type='button';remove.title=`${p.name}を削除`;remove.setAttribute('aria-label',remove.title);remove.addEventListener('click',()=>{openEditor(kind,p);removePreset(p);});
      const actions=el('span','preset-actions');actions.append(edit,remove);row.append(apply,actions);container.append(row);
      if(kind==='shapeSize'){
        const reorder=el('span','preset-order');reorder.setAttribute('role','group');reorder.setAttribute('aria-label',`${p.name}の表示順`);
        for(const [direction,symbol,label] of [[-1,'↑','上へ'],[1,'↓','下へ']]){
          const button=el('button','preset-move',symbol);button.type='button';button.title=`${p.name}を${label}移動`;
          button.setAttribute('aria-label',button.title);button.setAttribute('data-move',p.id);button.setAttribute('data-direction',String(direction));
          button.disabled=direction===-1?index===0:index===presets.length-1;button.setAttribute('data-boundary',String(button.disabled));
          button.addEventListener('click',()=>action(async()=>{
            commit(movePreset(bundle,p.id,direction));
            const buttons=Array.from(document.querySelectorAll('[data-move]')).filter(b=>b.dataset.move===p.id);
            (buttons.find(b=>b.dataset.direction===String(direction)&&!b.disabled)||buttons.find(b=>!b.disabled))?.focus();
            status(`「${p.name}」を${label}移動しました。`);
          }));
          reorder.append(button);
        }
        row.append(reorder);
      }
    }
  }
}
function openEditor(kind,p=null){
  editorKind=kind;editingId=p?.id||null;$('editor-title').textContent=`${TITLES[kind]}を${p?'編集':'登録'}`;
  $('preset-form').elements.presetName.value=p?.name||'';$('editor-fields').replaceChildren();$('editor-error').hidden=true;
  for(const [key,title,type,fallback,choices]of FORM[kind]){
    const label=el('label','',kind==='shapeSize'?`${key==='width'?'幅':'高さ'}（インチ）`:title);const field=el(type==='select'?'select':'input');field.name=key;
    if(type==='select')for(const [value,name]of Object.entries(choices)){const option=el('option','',name);option.value=value;field.append(option);}
    else {field.type=type;if(type==='number'){field.step='any';field.min=kind==='shapeSize'?String(displayLength(0.1,lengthUnit)):'0.1';if(kind==='shapeSize')field.max=String(displayLength(2000,lengthUnit));}else field.maxLength=key==='code'?255:100;field.required=kind!=='grid';}
    const raw=p?.value[key]??fallback;field.value=kind==='shapeSize'?Number(displayLength(raw,lengthUnit).toFixed(6)):raw;label.append(field);$('editor-fields').append(label);
  }
  $('delete-preset')?.remove();
  if(p){const button=el('button','text-button','このプリセットを削除');button.id='delete-preset';button.type='button';button.addEventListener('click',()=>removePreset(p));$('editor-fields').append(button);}
  $('editor').showModal();
}
function removePreset(p){return action(async()=>{if(!await ask(`「${p.name}」を削除しますか？`))return;try{commit(withoutPreset(bundle,p.id));$('editor').close();status('プリセットを削除しました。');}catch(error){editorError(error);}});}
function editorError(error){$('editor-error').textContent=error.message;$('editor-error').hidden=false;}
$('preset-form').addEventListener('submit',event=>{
  event.preventDefault();
  try{
    const form=$('preset-form'), value={};
    for(const [key,,type]of FORM[editorKind]){const raw=form.elements[key].value;if(type==='number'){if(raw==='')continue;value[key]=editorKind==='shapeSize'?storedLength(Number(raw),lengthUnit):Number(raw);}else value[key]=raw;}
    const preset=validatePreset({id:editingId||uid(),name:form.elements.presetName.value,kind:editorKind,value});
    const next={...bundle,presets:editingId?bundle.presets.map(p=>p.id===editingId?preset:p):[...bundle.presets,preset]};
    commit(next);$('editor').close();status('プリセットを保存しました。');
  }catch(error){editorError(error);}
});
$('capture').addEventListener('click',()=>action(async()=>{
  try{requireConnection();const values=await adapter.capture(editorKind,options());for(const [key,val]of Object.entries(values)){if(val===null||val===undefined||val==='')throw new Error('設定を読み取れませんでした。値を手入力してください。');const field=$('preset-form').elements[key];if(field)field.value=editorKind==='shapeSize'?Number(displayLength(val,lengthUnit).toFixed(6)):val;} $('editor-error').hidden=true;}catch(error){editorError(error);}
}));
$('close-editor').addEventListener('click',()=>$('editor').close());
document.querySelectorAll('[data-add]').forEach(button=>button.addEventListener('click',()=>openEditor(button.dataset.add)));
document.querySelectorAll('[data-tab]').forEach(button=>button.addEventListener('click',()=>{
  document.querySelectorAll('[data-tab]').forEach(b=>b.removeAttribute('aria-current'));button.setAttribute('aria-current','page');
  document.querySelectorAll('.tab').forEach(section=>section.hidden=section.id!==button.dataset.tab);
}));
function download(){
  const blob=new Blob([JSON.stringify({...bundle,shapeSizeStorageUnit:'mm',exportedAt:new Date().toISOString()},null,2)],{type:'application/json'});
  const url=URL.createObjectURL(blob),a=el('a');a.href=url;a.download=`excel-style-kit-${new Date().toISOString().slice(0,10)}.json`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);
}
$('export').addEventListener('click',()=>{try{download();status('JSONの書き出しを開始しました。保存先を確認してください。');}catch(error){status(error.message,true);}});
$('import-file').addEventListener('change',()=>action(async()=>{
  const file=$('import-file').files[0];if(!file)return;
  try{
    if(file.size>2*1024*1024)throw new Error('2MB以下のJSONファイルを選んでください。');
    const incoming=validateBundle(JSON.parse(await file.text())),mode=$('import-mode').value;
    const next=mergeBundle(bundle,incoming,mode);
    if(!await ask(`${incoming.presets.length}件を${mode==='replace'?'読み込み、現在の設定をすべて置き換えます':'追加・更新します'}。${mode==='replace'?'必要な設定は先に書き出してください。':''}実行しますか？`))return;
    commit(next);status(`${incoming.presets.length}件のプリセットを読み込みました。`);
  }finally{$('import-file').value='';}
}));
$('restore-defaults').addEventListener('click',()=>action(async()=>{if(await ask('初期プリセットを追加・更新しますか？')){commit(mergeBundle(bundle,defaults()));status('初期プリセットを追加・更新しました。');}}));
$('connect-native').addEventListener('click',()=>action(async()=>{
  requireConnection();if(demo)throw new Error('図形連携は実際のExcel内で接続してください。');
  $('bridge-status').textContent='接続しています…';
  try {const result=await adapter.connectNative();$('native-support').textContent='Excelで選択した1個または複数の図形に適用し、各図形の中心位置を保ちます。';$('bridge-status').textContent=`接続済み v${result.version}：サイズ登録の反映・整列はExcelの複数選択に適用されます。`;status('図形連携に接続しました。');}
  catch(error){$('bridge-status').textContent='未接続：補助アドインを確認して、再接続してください。';throw error;}
}));
document.querySelectorAll('[data-arrange]').forEach(button=>button.addEventListener('click',()=>action(async()=>{
  requireConnection();status('処理中です。図形の選択を変えずにお待ちください。');const result=await adapter.arrange(button.dataset.arrange);status(`${result.count||result.names?.length||0}個の図形を変更しました。`);
})));
renderPresets();

if(storedError)status('保存済み設定を読み込めませんでした。元のデータはまだ保持しています。JSONのバックアップを確認してから登録してください。',true);
function connect(){
  if(demo){adapter=new DemoAdapter();connected=true;$('connection').textContent='画面デモ ／ Excelのデータは変更しません';$('connection').classList.add('preview');$('diagnostics').textContent='デモモード。図形はサンプル、セルの操作はシミュレーションです。';return;}
  if(!globalThis.Office?.onReady){$('connection').textContent='Excel未接続 ／ プリセットの編集は利用できます';$('connection').classList.add('preview');$('diagnostics').textContent='ブラウザーで開いています。Excel内から読み込んでください。';return;}
  Office.onReady(info=>{
    if(info.host!==Office.HostType.Excel){$('connection').textContent='Excel内から開いてください';return;}
    const supported=Office.context.requirements.isSetSupported('ExcelApi','1.9');
    connected=supported;adapter=new ExcelAdapter();$('connection').textContent=supported?'● Excelに接続済み':'ExcelApi 1.9に対応したExcelへ更新してください';
    const native=supported&&adapter.supportsNativeSelection();
    $('diagnostics').textContent=`Excel / ${info.platform} / ExcelApi 1.9: ${supported?'対応':'未対応'} / Excel上の図形1個の取得: ${native?'対応':'未対応（ExcelApi 1.19が必要）'}`;
    if(!native){$('style-target').value='cells';$('native-support').textContent='このExcelは図形の直接選択に未対応です。ExcelApi 1.19対応版に更新してください。';}
  }).catch(error=>status(error.message,true));
  setTimeout(()=>{if(!connected){$('connection').textContent='Excel未接続 ／ Excel内から開いてください';$('connection').classList.add('preview');}},6000);
}
connect();
