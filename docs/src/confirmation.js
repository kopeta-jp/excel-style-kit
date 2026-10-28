// In-pane confirmation: Office webviews may suppress window.confirm().
export function askConfirmation(document,host,message) {
  return new Promise(resolve=>{
    const panel=document.createElement('div');panel.className='confirmation';
    panel.setAttribute('role','alertdialog');panel.setAttribute('aria-label',message);
    const text=document.createElement('p');text.textContent=message;
    const buttons=document.createElement('div');buttons.className='confirmation-actions';
    const cancel=document.createElement('button'),accept=document.createElement('button');
    cancel.type=accept.type='button';cancel.textContent='キャンセル';accept.textContent='実行する';accept.className='primary';
    const owner=host.closest('dialog');let settled=false;
    const finish=value=>{if(settled)return;settled=true;panel.remove();owner?.removeEventListener('close',onClose);resolve(value);};
    const onClose=()=>finish(false);
    owner?.addEventListener('close',onClose);
    cancel.addEventListener('click',()=>finish(false));accept.addEventListener('click',()=>finish(true));
    panel.addEventListener('keydown',event=>{if(event.key==='Escape'){event.preventDefault();event.stopPropagation();finish(false);}});
    buttons.append(cancel,accept);panel.append(text,buttons);host.append(panel);
    // Default focus is cancellation to avoid accidental destructive actions.
    cancel.focus();
  });
}
