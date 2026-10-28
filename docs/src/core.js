// Host-independent preset model and geometry. Reusable by a PowerPoint adapter.
export const SCHEMA = 'excel-style-kit';
export const VERSION = 1;
export const KINDS = ['shapeSize', 'color', 'font', 'fontSize', 'grid', 'border', 'numberFormat'];
export const mmToPt = value => value * 72 / 25.4;
export const ptToMm = value => value * 25.4 / 72;
// Verified against the user's Excel Format Shape pane (2.97", 2.63").
export const SHAPE_SIZE_UNIT = 'in';
export const LENGTH_UNITS = {cm:10, mm:1, in:25.4, pt:25.4/72};
export function displayLength(mm, unit='cm') {
  choice(unit,Object.keys(LENGTH_UNITS),'図形サイズの単位');
  return mm/LENGTH_UNITS[unit];
}
export function storedLength(value, unit='cm') {
  choice(unit,Object.keys(LENGTH_UNITS),'図形サイズの単位');
  if(!Number.isFinite(value))throw new Error('図形サイズは数値で入力してください。');
  return value*LENGTH_UNITS[unit];
}
export function uid() {
  return globalThis.crypto?.randomUUID?.() || `p-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}
function text(value, label, max = 80) {
  if (typeof value !== 'string' || !value.trim() || value.length > max) throw new Error(`${label}を1〜${max}文字で入力してください。`);
  return value.trim();
}
function number(value, label, min, max) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) throw new Error(`${label}は${min}〜${max}で入力してください。`);
  return value;
}
function choice(value, values, label) {
  if (!values.includes(value)) throw new Error(`${label}が不正です。`);
  return value;
}
export function hex(value) {
  if (typeof value !== 'string' || !/^#[0-9a-f]{6}$/i.test(value)) throw new Error('カラーは #RRGGBB で入力してください。');
  return value.toUpperCase();
}
export function validatePreset(p) {
  if (!p || typeof p !== 'object' || Array.isArray(p)) throw new Error('プリセットが不正です。');
  const kind = choice(p.kind, KINDS, '種類');
  const out = {id: text(p.id, 'ID', 120), kind, name: text(p.name, '名前')};
  const v = p.value;
  if (!v || typeof v !== 'object' || Array.isArray(v)) throw new Error('設定値が不正です。');
  switch (kind) {
    case 'shapeSize': out.value = {width: number(v.width, '幅 mm', 0.1, 2000), height: number(v.height, '高さ mm', 0.1, 2000)}; break;
    case 'color': out.value = {color: hex(v.color)}; break;
    case 'font': out.value = {name: text(v.name, 'フォント名', 100)}; break;
    case 'fontSize': out.value = {size: number(v.size, '文字サイズ pt', 1, 409)}; break;
    case 'grid': {
      out.value = {};
      if (v.rowHeight !== undefined && v.rowHeight !== null) out.value.rowHeight = number(v.rowHeight, '行高 pt', 1, 409);
      if (v.columnWidth !== undefined && v.columnWidth !== null) out.value.columnWidth = number(v.columnWidth, '列幅 pt', 1, 2000);
      if (!Object.keys(out.value).length) throw new Error('行高か列幅のいずれかを入力してください。');
      break;
    }
    case 'border': out.value = {
      pattern: choice(v.pattern, ['outer', 'all', 'bottom', 'none'], '罫線位置'),
      style: choice(v.style, ['Continuous', 'Dash', 'Dot', 'Double'], '線種'),
      weight: choice(v.weight, ['Hairline', 'Thin', 'Medium', 'Thick'], '線の太さ'), color: hex(v.color)
    }; break;
    case 'numberFormat': out.value = {code: text(v.code, '表示形式コード', 255)}; break;
  }
  return out;
}
export function validateBundle(data) {
  if (!data || data.schema !== SCHEMA || data.version !== VERSION || !Array.isArray(data.presets)) throw new Error('Excel Style Kit v1のJSONを選択してください。');
  if (data.presets.length > 1000) throw new Error('プリセットは最大1,000件です。');
  const presets = data.presets.map(validatePreset);
  if (new Set(presets.map(p => p.id)).size !== presets.length) throw new Error('IDが重複しています。');
  return {schema: SCHEMA, version: VERSION, presets};
}
export function mergeBundle(current, incoming, mode = 'merge') {
  const next = validateBundle(incoming);
  if (mode === 'replace') return next;
  choice(mode, ['merge'], '読み込み方法');
  const map = new Map(validateBundle(current).presets.map(p => [p.id, p]));
  next.presets.forEach(p => map.set(p.id, p));
  return validateBundle({schema: SCHEMA, version: VERSION, presets: [...map.values()]});
}
export function saveBundle(storage, bundle) {
  const valid = validateBundle(bundle);
  // Write before committing UI state: quota/security errors never silently lose data.
  storage.setItem(SCHEMA, JSON.stringify(valid));
  return valid;
}
export function withoutPreset(bundle,id) {
  const valid=validateBundle(bundle);
  if(!valid.presets.some(p=>p.id===id))throw new Error('削除するプリセットが見つかりません。');
  return {...valid,presets:valid.presets.filter(p=>p.id!==id)};
}
export function loadBundle(storage) {
  const raw = storage.getItem(SCHEMA);
  return raw === null ? defaults() : validateBundle(JSON.parse(raw));
}
export function defaults() {
  const items = [
    ['shapeSize','工程ボックス',{width:40,height:14}], ['shapeSize','注釈ボックス',{width:60,height:10}],
    ['color','ネイビー',{color:'#193E36'}], ['color','グリーン',{color:'#24876A'}],
    ['color','ライトグレー',{color:'#F1F4F2'}], ['color','ブラック',{color:'#222B27'}],
    ['font','游ゴシック',{name:'Yu Gothic'}], ['font','メイリオ',{name:'Meiryo'}], ['font','Arial',{name:'Arial'}],
    ...[9,10,11,12,14,18].map(size=>['fontSize',`${size} pt`,{size}]),
    ['grid','標準の行高',{rowHeight:18}], ['grid','ゆったりした行高',{rowHeight:30}],
    ['grid','標準の列幅',{columnWidth:72}], ['grid','広めの列幅',{columnWidth:120}],
    ['border','外枠',{pattern:'outer',style:'Continuous',weight:'Thin',color:'#A3B3AA'}],
    ['border','格子',{pattern:'all',style:'Continuous',weight:'Thin',color:'#A3B3AA'}],
    ['border','下線',{pattern:'bottom',style:'Continuous',weight:'Thin',color:'#193E36'}],
    ['numberFormat','桁区切り',{code:'#,##0'}], ['numberFormat','金額（円）',{code:'"¥"#,##0'}],
    ['numberFormat','割合',{code:'0.0%'}], ['numberFormat','日付',{code:'yyyy/mm/dd'}]
  ];
  return validateBundle({schema:SCHEMA,version:VERSION,presets:items.map(([kind,name,value],i)=>({id:`default-${i}`,kind,name,value}))});
}
export function resizePlan(shapes, preset, mode) {
  const {width,height} = validatePreset({id:'tmp',name:'tmp',kind:'shapeSize',value:preset}).value;
  choice(mode,['both','width','height','fit'],'サイズ適用方法');
  return shapes.map(s => {
    if (!(s.width > 0 && s.height > 0)) throw new Error('幅・高さが0の図形はサイズ変更できません。');
    const w=mmToPt(width), h=mmToPt(height);
    if(!Number.isFinite(s.left)||!Number.isFinite(s.top))throw new Error('図形の位置を読み取れませんでした。');
    let nextWidth=w,nextHeight=h;
    if(mode==='width')nextHeight=s.height;
    if(mode==='height')nextWidth=s.width;
    if(mode==='fit'){const k=Math.min(w/s.width,h/s.height);nextWidth=s.width*k;nextHeight=s.height*k;}
    // Resize about each shape's own center, independently of other shapes.
    return {id:s.id,width:nextWidth,height:nextHeight,left:s.left+(s.width-nextWidth)/2,top:s.top+(s.height-nextHeight)/2};
  });
}
export function geometryPlan(shapes, action) {
  const actions=['left','center','right','top','middle','bottom','distributeH','distributeV','size','width','height'];
  choice(action,actions,'配置方法');
  if(shapes.length < (action.startsWith('distribute')?3:2)) throw new Error(action.startsWith('distribute')?'等間隔配置は3個以上の図形をチェックしてください。':'図形を2個以上チェックしてください。');
  shapes.forEach(s=>['left','top','width','height'].forEach(k=>{if(!Number.isFinite(s[k]) || (['width','height'].includes(k)&&s[k]<0)) throw new Error('図形の座標またはサイズが不正です。');}));
  const first=shapes[0];
  if(['size','width','height'].includes(action)) return shapes.map(s=>({id:s.id,...(action!=='height'?{width:first.width}:{}),...(action!=='width'?{height:first.height}:{})}));
  if(action.startsWith('distribute')) {
    const horizontal=action==='distributeH', pos=horizontal?'left':'top', size=horizontal?'width':'height';
    const sorted=[...shapes].sort((a,b)=>a[pos]-b[pos]);
    const start=sorted[0][pos], end=Math.max(...sorted.map(s=>s[pos]+s[size]));
    const gap=(end-start-sorted.reduce((sum,s)=>sum+s[size],0))/(sorted.length-1);
    if(gap < -0.001) throw new Error('配置範囲が狭いため図形が重なります。両端の図形を離してから実行してください。');
    let cursor=start;
    return sorted.map(s=>{const p={id:s.id,[pos]:cursor};cursor+=s[size]+gap;return p;});
  }
  const l=Math.min(...shapes.map(s=>s.left)), r=Math.max(...shapes.map(s=>s.left+s.width));
  const t=Math.min(...shapes.map(s=>s.top)), b=Math.max(...shapes.map(s=>s.top+s.height));
  return shapes.map(s=>({id:s.id,...({
    left:{left:l},center:{left:(l+r-s.width)/2},right:{left:r-s.width},
    top:{top:t},middle:{top:(t+b-s.height)/2},bottom:{top:b-s.height}
  }[action])}));
}
export function describe(p,unit=SHAPE_SIZE_UNIT) {
  const v=p.value;
  switch(p.kind){
    case 'shapeSize':return `${displayLength(v.width,unit).toFixed(2)} × ${displayLength(v.height,unit).toFixed(2)} ${unit==='in'?'インチ':unit}`;
    case 'color':return v.color;
    case 'font':return v.name;
    case 'fontSize':return `${v.size} pt`;
    case 'grid':return [v.rowHeight!==undefined?`行 ${v.rowHeight} pt`:'',v.columnWidth!==undefined?`列 ${v.columnWidth} pt`:''].filter(Boolean).join(' / ');
    case 'border':return `${({outer:'外枠',all:'格子',bottom:'下線',none:'罫線なし'})[v.pattern]} · ${v.color}`;
    case 'numberFormat':return v.code;
  }
}
