// Local workbook mailbox. No network, executable strings, or macro injection.
export const NATIVE_CHANNEL='_ExcelStyleKitBridge';
export const NATIVE_PROTOCOL='ExcelStyleKitNative/v1';
const MODES=['both','width','height','fit'];
const ARRANGE=['left','center','right','top','middle','bottom','distributeH','distributeV','size','width','height','session'];
export class NativeBridge {
  constructor(run,wait=ms=>new Promise(resolve=>setTimeout(resolve,ms))) {this.run=run;this.wait=wait;this.busy=false;this.ready=false;this.session='';}
  async connect(){const result=await this.request('ping');this.session=result.session;if(!this.session)throw new Error('図形連携のバージョンが一致しません。');this.ready=true;return result;}
  async request(operation,args={}) {
    if(!['ping','resize','arrange','captureSize'].includes(operation))throw new Error('未対応の図形操作です。');
    if(operation==='resize'&&(!MODES.includes(args.mode)||![args.width,args.height].every(n=>Number.isFinite(n)&&n>0&&n<=6000)))throw new Error('図形サイズが不正です。');
    if(operation==='arrange'&&!ARRANGE.includes(args.mode))throw new Error('整列方法が不正です。');
    if(this.busy)throw new Error('図形操作の完了を待ってください。');
    this.busy=true;
    const token=`sk-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    let sheetId, channelId;
    try {
      await this.run(async ctx=>{
        const active=ctx.workbook.worksheets.getActiveWorksheet();active.load('id,name');
        let channel=ctx.workbook.worksheets.getItemOrNullObject(NATIVE_CHANNEL);channel.load('isNullObject');await ctx.sync();
        sheetId=active.id;
        if(channel.isNullObject) {channel=ctx.workbook.worksheets.add(NATIVE_CHANNEL);channel.visibility='Hidden';await ctx.sync();}
        else {const marker=channel.getRange('B1:B2');marker.load('values');await ctx.sync();if(marker.values[0][0]!==NATIVE_PROTOCOL)throw new Error('連携シートと同名の既存シートがあります。内容は変更していません。');if(['pending','processing'].includes(marker.values[1][0]))throw new Error('前の図形操作が残っています。完了を待つかExcelを再起動してください。');}
        channel.load('id');await ctx.sync();channelId=channel.id;
        const keys=['protocol','state','requestId','operation','arg1','arg2','mode','role','count','message','version','names','sheetName','width','height','session'];
        const vals=[NATIVE_PROTOCOL,'pending',token,operation,args.width||'',args.height||'',args.mode||'','','','','','',active.name,'','',this.session||''];
        channel.getRange('A1:B16').values=keys.map((key,i)=>[key,vals[i]]);await ctx.sync();
      });
      for(let i=0;i<30;i++) {
        await this.wait(300);
        const result=await this.run(async ctx=>{
          const channel=ctx.workbook.worksheets.getItem(channelId),active=ctx.workbook.worksheets.getActiveWorksheet();
          const data=channel.getRange('B1:B16');data.load('values');active.load('id');await ctx.sync();
          const v=data.values.map(row=>row[0]);
          if(v[2]!==token)throw new Error('別の図形操作を検出しました。結果を確認してください。');
          if(active.id!==sheetId){if(v[1]==='pending'){channel.getRange('B2').values=[['cancelled']];await ctx.sync();}throw new Error('シートが切り替わりました。図形を選択し直してください。');}
          if(v[1]==='error')throw new Error(`図形操作：${v[9]}`);
          if(v[1]==='done')return {target:'shapes',count:Number(v[8]),names:String(v[11]||'').split(' | ').filter(Boolean),width:Number(v[13]),height:Number(v[14]),version:String(v[10]),session:String(v[15]||'')};
          return null;
        });
        if(result)return result;
      }
      // Cancel an unclaimed command so installing the companion later cannot replay it.
      await this.run(async ctx=>{const channel=ctx.workbook.worksheets.getItem(channelId),state=channel.getRange('B2:B3');state.load('values');await ctx.sync();if(state.values[1][0]===token&&state.values[0][0]==='pending'){channel.getRange('B2').values=[['cancelled']];await ctx.sync();}});
      this.ready=false;
      throw new Error('図形連携が応答しません。ExcelStyleKitNative.xlamを開き、マクロが利用可能か確認してください。自動再実行はしません。');
    } finally {this.busy=false;}
  }
}
