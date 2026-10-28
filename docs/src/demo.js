// Explicit opt-in only. Never used as a silent fallback for failed Excel writes.
import {defaults,geometryPlan,resizePlan} from './core.js?v=0.2.0';
export class DemoAdapter {
  constructor(){this.items=[{id:'a',name:'工程 01',type:'GeometricShape',left:10,top:10,width:110,height:36},{id:'b',name:'工程 02',type:'GeometricShape',left:180,top:20,width:100,height:42},{id:'c',name:'工程 03',type:'GeometricShape',left:360,top:40,width:120,height:36}];}
  async listShapes(){return {sheetId:'demo',sheetName:'サンプルシート',shapes:this.items.map(s=>({...s}))};}
  supportsNativeSelection(){return true;}
  get(target){if(target?.mode!=='list')return [this.items[0]];if(!target?.ids?.length)throw new Error('補助一覧で対象をチェックしてください。');return target.ids.map(id=>this.items.find(s=>s.id===id));}
  async apply(p,options){const shapeTarget=p.kind==='shapeSize'||(['auto','shapes','listedShapes'].includes(options.target)&&['color','font','fontSize'].includes(p.kind));const shapes=shapeTarget?this.get(options.target==='listedShapes'?{...options.shapes,mode:'list'}:p.kind==='shapeSize'?options.shapes:{mode:'native'}):[];if(p.kind==='shapeSize')resizePlan(shapes,p.value,options.resizeMode).forEach(v=>Object.assign(this.items.find(s=>s.id===v.id),v));return {target:shapeTarget?'shapes':'cells',names:shapes.map(s=>s.name)};}
  async arrange(action,target){if(target?.mode!=='list')throw new Error('整列は補助一覧で対象を指定してください。');geometryPlan(this.get(target),action).forEach(v=>Object.assign(this.items.find(s=>s.id===v.id),v));}
  async capture(kind,options){if(kind==='shapeSize')this.get(options.shapes);return defaults().presets.find(p=>p.kind===kind).value;}
}
