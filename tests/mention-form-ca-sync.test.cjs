const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const assert=require('node:assert/strict');
const test=require('node:test');
const root=path.join(__dirname,'../メンション依頼フォーム_HTA_最新版');
const source=fs.readFileSync(process.env.MENTION_FORM_REVIEW_SOURCE||path.join(root,'mention-form.js'),'utf8');
const hta=fs.readFileSync(path.join(root,'メンション依頼フォーム.hta'),'utf8');

function form(legacy=false){
  const elements={};
  const document={getElementById:id=>elements[id]||null,activeElement:null};
  for(const match of hta.matchAll(/id="([^"]+)"/g)){
    const el={id:match[1],value:'',checked:false,disabled:false,style:{},className:'',innerText:'',listeners:{}};
    el.focus=()=>{document.activeElement=el;};
    if(legacy)el.attachEvent=(event,fn)=>{el.listeners[event]=fn;};
    else el.addEventListener=(event,fn)=>{el.listeners[event]=fn;};
    el.fire=event=>{if(el.listeners[event])el.listeners[event]();};
    let selectedIndex=0;
    Object.defineProperty(el,'selectedIndex',{get:()=>selectedIndex,set:v=>{selectedIndex=v;if(v===0)el.value='';}});
    elements[el.id]=el;
  }
  const ctx={document,screen:{availWidth:1920,availHeight:1080},alert:()=>{},confirm:()=>true};
  ctx.window={setTimeout:()=>1,clearTimeout:()=>{},setInterval:()=>1,clearInterval:()=>{}};
  vm.createContext(ctx);vm.runInContext(source,ctx);
  ctx.resizeApp=()=>{};ctx.bindSameCASync();
  const ids=no=>['org'+no,'ca'+no,'proxyOrg'+no,'proxyCA'+no+'_1','proxyCA'+no+'_2','proxyCA'+no+'_3'];
  const ca=no=>ids(no).map(id=>elements[id].value);
  const setCA=(no,name,event)=>ids(no).forEach((id,i)=>{elements[id].value=name+'-'+i;if(event)elements[id].fire(event);});
  const fill=(no,name)=>{
    setCA(no,name);
    for(const [field,value] of Object.entries({mailMemo:'メール'+no,processedAt:'2026/09/09 09:0'+no,dueDate:'2026/09/1'+no,type:'① 通常対応'}))elements[field+no].value=value;
    elements['short'+no].checked=no%2===1;elements['urgent'+no].checked=no%2===0;
  };
  const same=(no,on)=>{elements['sameCA'+no].checked=on;ctx.toggleSameCA(no);};
  const snapshot=no=>{const r=JSON.parse(JSON.stringify(ctx.collectRequest(no)));delete r.requestNo;return r;};
  fill(1,'A');fill(2,'B');fill(3,'C');ctx.visibleRequestCount=3;
  elements.requestBase.value='呉服';elements.requesterName.innerText='テスト依頼者';
  return {ctx,e:elements,ids,ca,setCA,fill,same,snapshot};
}

for(const link2 of [false,true])for(const link3 of [false,true]){
  test('依頼2を取り消しても繰り上げた内容が変わらない: sameCA2='+link2+', sameCA3='+link3,()=>{
    const f=form();if(link2)f.same(2,true);if(link3)f.same(3,true);
    const before=f.snapshot(3);
    f.ctx.cancelRequest(2);
    assert.deepEqual(f.snapshot(2),before);
    assert.equal(f.e.sameCA2.checked,false);assert.equal(f.e.sameCA3.checked,false);
    assert(f.ids(2).every(id=>!f.e[id].disabled));assert(f.ids(3).every(id=>!f.e[id].disabled));
    assert.equal(f.e.sameCALabel2.innerText,'依頼1と同一CA');
    assert.equal(f.ctx.visibleRequestCount,2);
    f.setCA(1,'A-変更','input');
    assert.deepEqual(f.snapshot(2),before);
    f.ctx.validateForm();assert.deepEqual(f.snapshot(2),before);
  });
}

test('依頼3の同一CAを先にONにしてから依頼2をONにすると6項目すべてが一致する',()=>{
  const f=form();const memo3=f.e.mailMemo3.value;f.same(3,true);f.same(2,true);
  assert.deepEqual(f.ca(2),f.ca(1));assert.deepEqual(f.ca(3),f.ca(1));
  assert.equal(f.e.mailMemo3.value,memo3);assert.equal(f.ctx.validateForm(),true);
});

test('依頼2の同一CAをOFFにしたときも依頼3へクリアが伝わり、その後の入力を反映する',()=>{
  const f=form();f.same(2,true);f.same(3,true);f.same(2,false);
  assert.deepEqual(f.ca(2),Array(6).fill(''));assert.deepEqual(f.ca(3),Array(6).fill(''));
  assert.equal(f.e.sameCA3.checked,true);assert.equal(f.ctx.validateForm(),false);
  f.setCA(2,'新CA','input');assert.deepEqual(f.ca(3),f.ca(2));
});

test('依頼1のクリアと再入力を依頼2・3へ反映する',()=>{
  const f=form();f.same(2,true);f.same(3,true);f.ctx.clearRequest(1);
  assert.deepEqual(f.ca(2),Array(6).fill(''));assert.deepEqual(f.ca(3),Array(6).fill(''));
  f.setCA(1,'再入力','change');assert.deepEqual(f.ca(2),f.ca(1));assert.deepEqual(f.ca(3),f.ca(1));
});

test('依頼2のクリアは依頼3へ伝わり、依頼1との連動は解除される',()=>{
  const f=form();f.same(2,true);f.same(3,true);f.ctx.clearRequest(2);
  assert.equal(f.e.sameCA2.checked,false);assert.equal(f.e.sameCA3.checked,true);
  assert.deepEqual(f.ca(3),Array(6).fill(''));
  f.setCA(1,'A-変更','input');assert.deepEqual(f.ca(2),Array(6).fill(''));assert.deepEqual(f.ca(3),Array(6).fill(''));
  f.setCA(2,'B-新規','input');assert.deepEqual(f.ca(3),f.ca(2));
});

test('取り消した依頼3を追加し直したとき同一CAのチェックと入力ロックが残らない',()=>{
  const f=form();f.same(3,true);f.ctx.cancelRequest(3);f.ctx.showRequest(3);
  assert.equal(f.e.sameCA3.checked,false);assert(f.ids(3).every(id=>!f.e[id].disabled));
  assert.deepEqual(f.ca(3),Array(6).fill(''));
});

test('依頼3がない状態で依頼2を取り消しても連動状態を初期化する',()=>{
  const f=form();f.ctx.visibleRequestCount=2;f.same(2,true);f.ctx.cancelRequest(2);f.ctx.showRequest(2);
  assert.equal(f.e.sameCA2.checked,false);assert(f.ids(2).every(id=>!f.e[id].disabled));
  assert.deepEqual(f.ca(2),Array(6).fill(''));
});

test('送信後の初期化で同一CA状態が残らず、次の依頼を入力できる',()=>{
  const f=form();f.same(2,true);f.same(3,true);f.ctx.resetAfterSend();
  for(const no of [2,3]){assert.equal(f.e['sameCA'+no].checked,false);assert(f.ids(no).every(id=>!f.e[id].disabled));}
  assert.equal(f.e.requestBase.value,'呉服');assert.equal(f.e.requesterName.innerText,'テスト依頼者');
  assert.equal(f.e.sendButton.disabled,false);
  f.fill(1,'次の依頼');f.ctx.showRequest(2);f.fill(2,'独立CA');
  assert.equal(f.ctx.validateForm(),true);assert.notDeepEqual(f.ca(2),f.ca(1));
});

test('送信前の確認でも同一CAの6項目を最新の値へそろえる',()=>{
  const f=form();f.same(2,true);f.same(3,true);f.setCA(1,'イベントなしの変更');
  assert.equal(f.ctx.validateForm(),true);assert.deepEqual(f.ca(2),f.ca(1));assert.deepEqual(f.ca(3),f.ca(1));
});

test('同一CAがOFFの依頼は他の依頼を編集しても変わらない',()=>{
  const f=form();const before2=f.snapshot(2),before3=f.snapshot(3);
  f.setCA(1,'A-変更','input');f.ctx.validateForm();
  assert.deepEqual(f.snapshot(2),before2);assert.deepEqual(f.snapshot(3),before3);
});

test('旧イベント方式でも依頼1から依頼3までCAの変更を反映する',()=>{
  const f=form(true);f.same(2,true);f.same(3,true);f.setCA(1,'旧イベント','onchange');
  assert.deepEqual(f.ca(2),f.ca(1));assert.deepEqual(f.ca(3),f.ca(1));
});

test('取消確認の「戻る」では値も同一CAの設定も変更しない',()=>{
  const f=form();f.same(2,true);f.same(3,true);const before=f.snapshot(2);
  f.ctx.requestCancel(2);f.ctx.closeCancelConfirm();
  assert.deepEqual(f.snapshot(2),before);assert.equal(f.e.sameCA2.checked,true);assert.equal(f.ctx.visibleRequestCount,3);
});

test('送信中でも次の別依頼を受付でき、送信ボタンのロックを追加しない',()=>{
  const f=form();const pending=[],launched=[];let nextId=0;
  f.ctx.getProductionCsvFolder=()=>'/share';
  f.ctx.createRequestId=()=>'ID-'+(++nextId);
  f.ctx.savePendingPackage=(base,who,id,time,requests)=>{pending.push({id,requests:JSON.parse(JSON.stringify(requests))});return '/pending/'+id+'.csv';};
  f.ctx.getPendingCount=()=>pending.length;f.ctx.removeCompletionReadyMarkerIfExists=()=>{};
  f.ctx.signalCompletionNotificationReady=()=>{};f.ctx.launchBackgroundWorker=p=>launched.push(p);
  f.ctx.visibleRequestCount=1;f.ctx.sendRequest();
  assert.equal(f.e.sendButton.disabled,false);
  f.ctx.confirmAcceptedAndMinimize();f.fill(1,'別の依頼');f.ctx.sendRequest();
  assert.equal(pending.length,2);assert.equal(launched.length,2);assert.equal(f.e.sendButton.disabled,false);
  assert.notEqual(pending[0].requests[0].caName,pending[1].requests[0].caName);
});
