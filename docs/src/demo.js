// Explicit opt-in only. Never used as a silent fallback for failed Excel writes.
import {defaults,geometryPlan,resizePlan} from './core.js';
export class DemoAdapter {
  constructor(){this.items=[{id:'a',name:'工程 01',type:'GeometricShape',left:10,top:10,width:110,height:36},{id:'b',name:'工程 02',type:'GeometricShape',left:180,top:20,width:100,height:42},{id:'c',name:'工程 03',type:'GeometricShape',left:360,top:40,width:120,height:36}];}
  async listShapes(){return {sheetId:'demo',sheetName:'サンプルシート',shapes:this.items.map(s=>({...s}))};}
  get(target){if(!target?.ids?.length)throw new Error('図形一覧で対象をチェックしてください。');return target.ids.map(id=>this.items.find(s=>s.id===id));}
  async apply(p,options){if(p.kind==='shapeSize'){const shapes=this.get(options.shapes);resizePlan(shapes,p.value,options.resizeMode).forEach(v=>Object.assign(this.items.find(s=>s.id===v.id),v));}else if(options.target==='shapes'&&['color','font','fontSize'].includes(p.kind))this.get(options.shapes);}
  async arrange(action,target){geometryPlan(this.get(target),action).forEach(v=>Object.assign(this.items.find(s=>s.id===v.id),v));}
  async capture(kind,options){if(kind==='shapeSize')this.get(options.shapes);return defaults().presets.find(p=>p.kind===kind).value;}
}
