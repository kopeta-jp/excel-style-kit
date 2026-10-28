import {resizePlan, ptToMm, validatePreset} from './core.js?v=0.2.0';

const OUTER=['EdgeTop','EdgeBottom','EdgeLeft','EdgeRight'];
const ALL=[...OUTER,'InsideHorizontal','InsideVertical'];
export class ExcelAdapter {
  supportsNativeSelection() {
    return !!globalThis.Office?.context?.requirements?.isSetSupported('ExcelApi','1.19');
  }
  async run(fn) {
    if(!globalThis.Excel || !globalThis.Office?.context) throw new Error('Excel内からアドインを開いてください。');
    return Excel.run(fn);
  }
  async shapes(ctx) {
    const shape=await this.activeShape(ctx);
    if(!shape)throw new Error('Excel上で図形を1個選択してから実行してください。');
    return [shape];
  }
  async activeShape(ctx) {
    if(!this.supportsNativeSelection() || typeof ctx.workbook.getActiveShapeOrNullObject!=='function') throw new Error('Excel上の図形選択を使うにはExcelApi 1.19対応版が必要です。Excelを更新してください。セルの操作は引き続き利用できます。');
    const shape=ctx.workbook.getActiveShapeOrNullObject();
    // Do not request shape properties until null-object status has been resolved.
    shape.load('isNullObject');
    await ctx.sync();
    if(shape.isNullObject)return null;
    shape.load('id,name,type,left,top,width,height,lockAspectRatio');await ctx.sync();
    return shape;
  }
  async resolveShapes(ctx,kind,options={}) {
    if(kind==='shapeSize')return this.shapes(ctx,options.shapes);
    if(!['color','font','fontSize'].includes(kind))return null;
    if(options.target==='auto') {
      const shape=await this.activeShape(ctx);
      return shape?[shape]:null;
    }
    if(options.target==='shapes')return this.shapes(ctx,{mode:'native'});
    if(options.target==='listedShapes')throw new Error('旧版の一覧指定は廃止しました。Excel上で図形を選択してください。');
    return null;
  }
  async ranges(ctx) {
    const selected=ctx.workbook.getSelectedRanges();selected.areas.load('items');await ctx.sync();
    if(selected.areas.items.length>100) throw new Error('一度に選択できる範囲は100か所までです。');
    return selected.areas.items;
  }
  async apply(preset, options) {
    const p=validatePreset(preset),v=p.value;
    return this.run(async ctx=>{
      const shapes=await this.resolveShapes(ctx,p.kind,options),isShape=shapes!==null;
      if(isShape) {
        // Validate all before queueing writes; images/groups have no editable text here.
        if(['font','fontSize'].includes(p.kind)||(p.kind==='color'&&options.colorRole==='font')) {
          if(shapes.some(s=>!['GeometricShape','TextBox'].includes(s.type))) throw new Error('文字書式は通常の図形・テキストボックスのみ対応しています。画像・グループを対象から外してください。');
        }
        if(p.kind==='shapeSize') {
          const plan=resizePlan(shapes,v,options.resizeMode);
          // Direct width/height setters do not require changing lockAspectRatio.
          plan.forEach(update=>{const s=shapes.find(s=>s.id===update.id);s.width=update.width;s.height=update.height;});
          await ctx.sync();
          shapes.forEach(s=>s.load('width,height'));await ctx.sync();
          if(plan.some(p=>{const s=shapes.find(s=>s.id===p.id);return Math.abs(s.width-p.width)>0.75||Math.abs(s.height-p.height)>0.75;}))throw new Error('Excelが指定サイズを反映しませんでした。図形の保護や種類を確認してください。');
        } else shapes.forEach(s=>{
          if(p.kind==='font')s.textFrame.textRange.font.name=v.name;
          if(p.kind==='fontSize')s.textFrame.textRange.font.size=v.size;
          if(p.kind==='color') {
            if(options.colorRole==='fill')s.fill.setSolidColor(v.color);
            else if(options.colorRole==='font')s.textFrame.textRange.font.color=v.color;
            else {s.lineFormat.color=v.color;s.lineFormat.visible=true;}
          }
        });
      } else {
        const ranges=await this.ranges(ctx);
        if(p.kind==='numberFormat') {
          ranges.forEach(r=>r.load('rowCount,columnCount'));await ctx.sync();
          if(ranges.reduce((n,r)=>n+r.rowCount*r.columnCount,0)>100000) throw new Error('表示形式は一度に10万セルまでです。列全体ではなく必要な範囲を選択してください。');
        }
        ranges.forEach(r=>{
          if(p.kind==='font')r.format.font.name=v.name;
          if(p.kind==='fontSize')r.format.font.size=v.size;
          if(p.kind==='grid') {if(v.rowHeight!==undefined)r.format.rowHeight=v.rowHeight;if(v.columnWidth!==undefined)r.format.columnWidth=v.columnWidth;}
          if(p.kind==='color') {
            if(options.colorRole==='fill')r.format.fill.color=v.color;
            else if(options.colorRole==='font')r.format.font.color=v.color;
            else ALL.forEach(edge=>{r.format.borders.getItem(edge).color=v.color;});
          }
          if(p.kind==='border') {
            if(v.pattern==='none') [...ALL,'DiagonalDown','DiagonalUp'].forEach(edge=>r.format.borders.getItem(edge).style='None');
            else (v.pattern==='outer'?OUTER:v.pattern==='bottom'?['EdgeBottom']:ALL).forEach(edge=>{
              const border=r.format.borders.getItem(edge);border.style=v.style;border.weight=v.weight;border.color=v.color;
            });
          }
          if(p.kind==='numberFormat')r.numberFormat=Array.from({length:r.rowCount},()=>Array(r.columnCount).fill(v.code));
        });
      }
      await ctx.sync();
      return {target:isShape?'shapes':'cells',names:isShape?shapes.map(s=>s.name):[]};
    });
  }
  async arrange() {
    throw new Error('現在のパネル方式では複数選択した図形全体を取得できないため、整列は未対応です。');
  }
  async capture(kind,options) {
    return this.run(async ctx=>{
      const shapes=await this.resolveShapes(ctx,kind,options),isShape=shapes!==null;
      let source;
      if(isShape) {
        if(shapes.length!==1)throw new Error('登録元の図形を1個だけ指定してください。');
        source=shapes[0];
        if(kind==='shapeSize')return {width:+ptToMm(source.width).toFixed(2),height:+ptToMm(source.height).toFixed(2)};
        if((kind!=='color'||options.colorRole==='font')&&!['GeometricShape','TextBox'].includes(source.type)) throw new Error('通常の図形またはテキストボックスを選んでください。');
        if(kind==='color') {
          const object=options.colorRole==='fill'?source.fill:options.colorRole==='line'?source.lineFormat:source.textFrame.textRange.font;
          const prop=options.colorRole==='fill'?'foregroundColor':'color';object.load(prop);await ctx.sync();return {color:object[prop]};
        }
        source=source.textFrame.textRange;
      } else source=ctx.workbook.getSelectedRange().getCell(0,0);
      if(kind==='font'||kind==='fontSize') {
        const font=isShape?source.font:source.format.font;const prop=kind==='font'?'name':'size';font.load(prop);await ctx.sync();return {[prop]:font[prop]};
      }
      if(kind==='grid'){source.format.load('rowHeight,columnWidth');await ctx.sync();return {rowHeight:source.format.rowHeight,columnWidth:source.format.columnWidth};}
      if(kind==='numberFormat'){source.load('numberFormat');await ctx.sync();return {code:source.numberFormat[0][0]};}
      if(kind==='border') {
        const border=source.format.borders.getItem('EdgeBottom');border.load('style,weight,color');await ctx.sync();
        return {pattern:border.style==='None'?'none':'bottom',style:border.style==='None'?'Continuous':border.style,weight:border.weight||'Thin',color:border.color||'#000000'};
      }
      if(kind==='color') {
        const object=options.colorRole==='fill'?source.format.fill:options.colorRole==='line'?source.format.borders.getItem('EdgeBottom'):source.format.font;
        object.load('color');await ctx.sync();return {color:object.color};
      }
      throw new Error('この種類は読み取りできません。');
    });
  }
}
