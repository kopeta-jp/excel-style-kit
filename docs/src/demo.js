// Explicit opt-in simulation; never a fallback for failed Excel operations.
import {defaults,resizePlan,ptToMm} from './core.js?v=0.4.0';
export class DemoAdapter {
  constructor(){this.shape={id:'demo',name:'サンプル図形',left:100,top:100,width:110,height:36};}
  supportsNativeSelection(){return true;}
  async apply(p,options){const isShape=p.kind==='shapeSize'||(['auto','shapes'].includes(options.target)&&['color','font','fontSize'].includes(p.kind));if(p.kind==='shapeSize')Object.assign(this.shape,resizePlan([this.shape],p.value,options.resizeMode)[0]);return {target:isShape?'shapes':'cells',names:isShape?[this.shape.name]:[]};}
  async arrange(){throw new Error('複数図形の整列は未対応です。');}
  async capture(kind){if(kind==='shapeSize')return {width:ptToMm(this.shape.width),height:ptToMm(this.shape.height)};return defaults().presets.find(p=>p.kind===kind).value;}
}
