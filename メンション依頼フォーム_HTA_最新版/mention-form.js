var CSV_SUBFOLDER_NAME = "書き込み用";
var PENDING_FOLDER_NAME = "MentionRequest_Pending";
var BACKGROUND_WORKER_NAME = "mention-request-worker.js";
var WORKER_MAX_RETRY_SECONDS = 300;
var activeCompletionReadyPath = "";
var completionReadySignaled = false;


/* =========================================================
   v28.12 選択肢マスタJS集約＋受付後最小化＋即時書き込み＋標準完了Popup
   ・Pending保存後すぐWorkerを起動
   ・共有CSVへの書き込みは待機なしで即時開始
   ・受付はメインHTA内モーダルで表示
   ・OK押下またはフォーム終了を合図に、完了時はWindows標準Popupを表示
   ・フォームを閉じてもWorkerと完了通知は継続
   ========================================================= */

function showAcceptedAndPrepareMinimize(){
    var button=$("systemModalOkButton");

    if(button){
        button.onclick=confirmAcceptedAndMinimize;
    }

    showSystemModal(
        "依頼を受け付けました",
        "送信処理を開始しました。\n\nOKを押すとフォームを最小化します。",
        "accepted",
        true
    );
}

function confirmAcceptedAndMinimize(){
    // Worker側の完了Popupを表示してよいタイミングを通知してから最小化する。
    signalCompletionNotificationReady();
    hideSystemModal();

    // 受付後に入力欄を初期化。拠点・依頼者は既存仕様どおり維持。
    resetAfterSend();
    setStatus("送信中です…");

    // HTML Help ActiveX Control の Minimize コマンドで
    // 現在のHTAウィンドウを直接最小化する。
    window.setTimeout(function(){
        var minimized=false;

        try{
            var ctrl=$("HHCtrlMinimizeWindowObject");
            if(ctrl){
                ctrl.Click();
                minimized=true;
            }
        }catch(minErr){}

        // ActiveX Controlが利用できない環境のみフォールバック。
        if(!minimized){
            try{
                var shell=new ActiveXObject("WScript.Shell");
                try{
                    shell.AppActivate(document.title);
                }catch(activateErr){}

                shell.SendKeys("% ");
                window.setTimeout(function(){
                    try{
                        shell.SendKeys("n");
                    }catch(keyErr2){}
                },180);
            }catch(keyErr1){}
        }
    },150);
}

var BASE_OPTIONS = [
    "",
    "首都圏",
    "札幌",
    "新潟",
    "呉服"
];

var TYPE_OPTIONS = [
    "",
    "① 通常対応",
    "② 日程確認",
    "③ 代理対応",
    "④ 再調整",
    "⑤ 候補者への確認が必要なため、内容を確認してから対応を進める",
    "⑥ 企業確認",
    "⑦ 面接日程の変更に伴い、候補者・企業双方への連絡が必要",
    "⑧ 至急確認",
    "⑨ 担当者へ確認後、回答内容に沿って処理を実施する",
    "⑩ 日程確定",
    "⑪ 複数候補日の調整が必要なため、関係者の予定を確認して対応する",
    "⑫ 情報更新",
    "⑬ 対応可否を確認し、必要に応じて担当部署へエスカレーションする",
    "⑭ 保留対応",
    "⑮ その他の個別対応"
];

var visibleRequestCount = 1;
var layoutTimer = null;
var revealTimer = null;
var hasPositionedWindow = false;
var lastAvailWidth = 0;
var lastAvailHeight = 0;
var screenWatchTimer = null;

function $(id){ return document.getElementById(id); }

function initApp(){
    // 受付モーダルのOKを押さずにフォームを閉じた場合も、
    // Workerへ「完了通知を表示してよい」ことを伝える。
    window.onunload=function(){
        signalCompletionNotificationReady();
    };

    populateBases();
    populateTypes();
    setRequesterName();

    try{
        getProductionCsvFolder(true);
        setStatus("");
    }catch(configErr){
        setStatus("書き込み用フォルダが見つかりません");
    }

    lastAvailWidth=screen.availWidth;
    lastAvailHeight=screen.availHeight;

    // 初回は画面を見せる前にサイズを確定
    resizeApp(true);

    // HTAを開いたまま解像度・表示倍率・接続モニターが変わった場合にも追従
    startScreenWatcher();

    bindSameCASync();

    /*
      v26:
      起動時にPendingがある場合、
      先に自動再送状態へ入れてからWatcherを開始する。
      これにより一瞬だけ「失敗＋再送ボタン」が表示されてから
      「自動再送中」へ切り替わるチラつきを防ぐ。
    */

    // 前回終了時などに残ったPendingを先に自動再送
    recoverPendingPackagesOnStartup();

    // その後にPending状態監視を開始
    startPendingWatcher();
}

function setRequesterName(){
    var displayName="";
    var userName="";
    var domainName="";
    var el=$("requesterName");

    try{
        var network=new ActiveXObject("WScript.Network");
        userName=String(network.UserName||"");
        domainName=String(network.UserDomain||"");

        // VBScriptで、VBA UserFormと同じWinNT FullNameを取得
        try{
            displayName=String(GetWindowsDisplayNameHTA()||"");
        }catch(nameErr){
            displayName="";
        }

        if(!displayName){
            displayName=userName;
        }

        if(el){
            el.innerText=displayName;
            el.title=domainName+"\\"+userName;
        }

    }catch(err){
        if(el){
            el.innerText=userName||"取得できませんでした";
        }
    }
}

function populateBases(){
    var i,sel,opt;
    sel=$("requestBase");
    if(!sel){ return; }

    while(sel.options.length>0){ sel.remove(0); }

    for(i=0;i<BASE_OPTIONS.length;i++){
        opt=document.createElement("option");
        opt.value=BASE_OPTIONS[i];
        opt.text=(BASE_OPTIONS[i]==="" ? "選択してください" : BASE_OPTIONS[i]);
        sel.add(opt);
    }
}

function populateTypes(){
    var i,j,sel,opt;
    for(i=1;i<=3;i++){
        sel=$("type"+i);
        while(sel.options.length>0){ sel.remove(0); }
        for(j=0;j<TYPE_OPTIONS.length;j++){
            opt=document.createElement("option");
            opt.value=TYPE_OPTIONS[j];
            opt.text=(TYPE_OPTIONS[j]==="" ? "選択してください" : TYPE_OPTIONS[j]);
            sel.add(opt);
        }
    }
}

function showRequest(no){
    // 2件目を開く前に依頼1、3件目を開く前に依頼2をチェック
    if(no===2){
        clearErrors();
        if(!validateRequestSection(1)){
            return;
        }
    }else if(no===3){
        clearErrors();
        if(!validateRequestSection(2)){
            return;
        }
    }

    // カード追加～ウィンドウサイズ確定まで一旦隠して、途中描画を見せない
    hideAppForLayout();

    if(no===2){
        $("request2").className=$("request2").className.replace(" hidden","");
        $("btnAdd2").style.display="none";
        visibleRequestCount=2;
    }else if(no===3){
        $("request3").className=$("request3").className.replace(" hidden","");
        $("btnAdd3").style.display="none";
        visibleRequestCount=3;
    }

    resizeApp(false);
}


function hideAppForLayout(){
    var app=$("app");
    if(app){
        app.style.visibility="hidden";
    }
}

function showAppAfterLayout(){
    if(revealTimer){
        window.clearTimeout(revealTimer);
        revealTimer=null;
    }

    revealTimer=window.setTimeout(function(){
        var app=$("app");
        if(app){
            app.style.visibility="visible";
        }
        revealTimer=null;
    },20);
}

function startScreenWatcher(){
    if(screenWatchTimer){
        window.clearInterval(screenWatchTimer);
    }

    screenWatchTimer=window.setInterval(function(){
        try{
            var w=screen.availWidth;
            var h=screen.availHeight;

            if(w!==lastAvailWidth || h!==lastAvailHeight){
                lastAvailWidth=w;
                lastAvailHeight=h;

                // 解像度やWindows表示倍率の変更後に自動再フィット
                resizeApp(true);
            }
        }catch(err){
        }
    },500);
}

function getViewportSize(){
    var w=0;
    var h=0;

    try{
        if(document.documentElement){
            w=document.documentElement.clientWidth||0;
            h=document.documentElement.clientHeight||0;
        }
    }catch(err){
    }

    if((!w || !h) && document.body){
        try{
            w=w||document.body.clientWidth||0;
            h=h||document.body.clientHeight||0;
        }catch(err2){
        }
    }

    return {width:w,height:h};
}

function getWindowOuterSize(fallbackW,fallbackH){
    var w=0;
    var h=0;

    try{
        w=window.outerWidth||0;
        h=window.outerHeight||0;
    }catch(err){
    }

    if(!w){w=fallbackW||0;}
    if(!h){h=fallbackH||0;}

    return {width:w,height:h};
}

function getCurrentChromeSize(){
    var viewport=getViewportSize();
    var outer=getWindowOuterSize(0,0);
    var chromeW=0;
    var chromeH=0;

    if(outer.width>0 && viewport.width>0){
        chromeW=outer.width-viewport.width;
    }

    if(outer.height>0 && viewport.height>0){
        chromeH=outer.height-viewport.height;
    }

    /*
      HTA / IE互換環境でouterWidth等が取れない場合だけ
      従来値をfallbackとして使用する。
    */
    if(chromeW<=0 || chromeW>200){
        chromeW=54;
    }

    if(chromeH<=0 || chromeH>200){
        chromeH=76;
    }

    return {width:chromeW,height:chromeH};
}

function centerCurrentWindow(outerW,outerH){
    try{
        var moveX=Math.max(0,Math.floor((screen.availWidth-outerW)/2));
        var moveY=Math.max(0,Math.floor((screen.availHeight-outerH)/2));
        window.moveTo(moveX,moveY);
        hasPositionedWindow=true;
    }catch(err){
    }
}

function resizeApp(centerOnFirst){
    hideAppForLayout();

    if(layoutTimer){
        window.clearTimeout(layoutTimer);
        layoutTimer=null;
    }

    layoutTimer=window.setTimeout(function(){
        try{
            var app=$("app");
            if(!app){return;}

            /*
              v19:
              フォームの自動縮小を完全に廃止。
              依頼1件でも3件でもカード・文字・入力欄サイズは同一。
              画面に収まらない高さだけ縦スクロールで対応する。
            */
            document.body.style.zoom="1";

            var chrome=getCurrentChromeSize();

            // app本体1680px + 左右余白14px×2
            var naturalClientW=app.offsetWidth+28;

            // 現在表示されているカード数に応じた自然高さ
            var naturalClientH=app.offsetHeight+30;

            // 24インチ 1920×1080 を想定し、画面端へ少し余白を残す
            var maxOuterW=screen.availWidth-40;
            var maxOuterH=screen.availHeight-34;

            var wantedW=Math.ceil(naturalClientW+chrome.width);
            var wantedH=Math.ceil(naturalClientH+chrome.height);

            // 横方向は縮小せず、画面に収まる範囲で固定幅
            if(wantedW>maxOuterW){
                wantedW=maxOuterW;
            }

            // 縦方向だけ画面サイズを上限にする
            if(wantedH>maxOuterH){
                wantedH=maxOuterH;
            }

            window.resizeTo(wantedW,wantedH);
            centerCurrentWindow(wantedW,wantedH);
            window.scrollTo(0,0);

        }catch(err){
        }finally{
            layoutTimer=null;
            showAppAfterLayout();
        }
    },15);
}

function val(id){
    var el=$(id);
    return el?String(el.value||""):"";
}

function checked(id){
    var el=$(id);
    return el?!!el.checked:false;
}

function collectRequest(no){
    return {
        requestNo:no,
        organization:val("org"+no),
        caName:val("ca"+no),
        proxyOrganization:val("proxyOrg"+no),
        proxyCA1:val("proxyCA"+no+"_1"),
        proxyCA2:val("proxyCA"+no+"_2"),
        proxyCA3:val("proxyCA"+no+"_3"),
        mailMemo:val("mailMemo"+no),
        processedAt:val("processedAt"+no),
        dueDate:val("dueDate"+no),
        shortTime:checked("short"+no),
        urgent:checked("urgent"+no),
        type:val("type"+no)
    };
}

function hasAnyInput(req){
    return !!(req.organization||req.caName||req.proxyOrganization||req.proxyCA1||
              req.proxyCA2||req.proxyCA3||req.mailMemo||req.processedAt||
              req.dueDate||req.shortTime||req.urgent||req.type);
}

function validateForm(){
    var i;

    clearErrors();

    // 表示中の依頼を、画面左→右の順番でチェック
    // 組織/CA名/代理CA組織/代理CA名 → メールメモ → 処理日時 → タイプ
    for(i=1;i<=visibleRequestCount;i++){
        if(!validateRequestSection(i)){
            return false;
        }
    }

    // 拠点は最後にチェック
    if(!trimValue("requestBase")){
        markError("requestBase");
        showValidationModal(0,"拠点を選択してください","requestBase");
        return false;
    }

    return true;
}

function validateRequestSection(i){
    if(!trimValue("org"+i)){
        markError("org"+i);
        showValidationModal(i,"依頼"+i+"：組織を入力してください","org"+i);
        return false;
    }

    if(!trimValue("ca"+i)){
        markError("ca"+i);
        showValidationModal(i,"依頼"+i+"：CA名を入力してください","ca"+i);
        return false;
    }

    var proxyOrg=trimValue("proxyOrg"+i);
    var proxyCA1=trimValue("proxyCA"+i+"_1");

    if(proxyOrg && !proxyCA1){
        markError("proxyCA"+i+"_1");
        showValidationModal(i,"依頼"+i+"：代理CA組織が入力されているため、代理CA名も入力してください","proxyCA"+i+"_1");
        return false;
    }

    if(proxyCA1 && !proxyOrg){
        markError("proxyOrg"+i);
        showValidationModal(i,"依頼"+i+"：代理CA名が入力されているため、代理CA組織も入力してください","proxyOrg"+i);
        return false;
    }

    if(!trimValue("mailMemo"+i)){
        markError("mailMemo"+i);
        showValidationModal(i,"依頼"+i+"：メールメモを入力してください","mailMemo"+i);
        return false;
    }

    if(!trimValue("processedAt"+i)){
        markError("processedAt"+i);
        showValidationModal(i,"依頼"+i+"：処理日時を入力してください","processedAt"+i);
        return false;
    }

    if(!trimValue("type"+i)){
        markError("type"+i);
        showValidationModal(i,"依頼"+i+"：タイプを選択してください","type"+i);
        return false;
    }

    return true;
}


function clearRequest(no){
    var ids=["org"+no,"ca"+no,"proxyOrg"+no,"proxyCA"+no+"_1","proxyCA"+no+"_2","proxyCA"+no+"_3","mailMemo"+no,"processedAt"+no,"dueDate"+no];
    var i,el;

    if(no===2 || no===3){
        el=$("sameCA"+no);
        if(el){ el.checked=false; }
        setSameCAFieldsLocked(no,false);
        updateSameCALabel(no,false);
    }

    for(i=0;i<ids.length;i++){
        el=$(ids[i]);
        if(el){ el.value=""; }
    }

    el=$("type"+no); if(el){ el.selectedIndex=0; }
    el=$("short"+no); if(el){ el.checked=false; }
    el=$("urgent"+no); if(el){ el.checked=false; }

    clearErrors();
    try{$("org"+no).focus();}catch(err){}
}

function getSameCASourceNo(no){
    return no===2 ? 1 : 2;
}

function getSameCAFieldPairs(no){
    var src=getSameCASourceNo(no);
    return [
        ["org"+src,"org"+no],
        ["ca"+src,"ca"+no],
        ["proxyOrg"+src,"proxyOrg"+no],
        ["proxyCA"+src+"_1","proxyCA"+no+"_1"],
        ["proxyCA"+src+"_2","proxyCA"+no+"_2"],
        ["proxyCA"+src+"_3","proxyCA"+no+"_3"]
    ];
}

function copySameCAValues(no){
    var pairs=getSameCAFieldPairs(no),i,a,b;
    for(i=0;i<pairs.length;i++){
        a=$(pairs[i][0]); b=$(pairs[i][1]);
        if(a&&b){ b.value=a.value; }
    }
}

function setSameCAFieldsLocked(no,locked){
    var pairs=getSameCAFieldPairs(no),i,b;
    for(i=0;i<pairs.length;i++){
        b=$(pairs[i][1]);
        if(b){ b.disabled=!!locked; }
    }
}

function updateSameCALabel(no,on){
    var src=getSameCASourceNo(no),label=$("sameCALabel"+no);
    if(label){
        label.innerText=on ? "依頼"+src+"のCAを反映中" : "依頼"+src+"と同一CA";
    }
}

function toggleSameCA(no){
    var box=$("sameCA"+no);
    if(!box){return;}

    if(box.checked){
        copySameCAValues(no);
        setSameCAFieldsLocked(no,true);
        updateSameCALabel(no,true);
    }else{
        // フリー入力へ切り替える際は、反映していたCA関連情報を全クリア
        setSameCAFieldsLocked(no,false);
        clearSameCAFields(no);
        updateSameCALabel(no,false);
    }
}

function clearSameCAFields(no){
    var ids=[
        "org"+no,
        "ca"+no,
        "proxyOrg"+no,
        "proxyCA"+no+"_1",
        "proxyCA"+no+"_2",
        "proxyCA"+no+"_3"
    ];
    var i,el;

    for(i=0;i<ids.length;i++){
        el=$(ids[i]);
        if(el){
            el.value="";
        }
    }

    clearErrors();
    try{
        $("org"+no).focus();
    }catch(err){
    }
}

function syncSameCATarget(no){
    var box=$("sameCA"+no);
    if(box&&box.checked){ copySameCAValues(no); }
}

function bindSameCASync(){
    var src1=["org1","ca1","proxyOrg1","proxyCA1_1","proxyCA1_2","proxyCA1_3"];
    var src2=["org2","ca2","proxyOrg2","proxyCA2_1","proxyCA2_2","proxyCA2_3"];
    var i;

    for(i=0;i<src1.length;i++){ bindSameCAEvent(src1[i],2); }
    for(i=0;i<src2.length;i++){ bindSameCAEvent(src2[i],3); }
}

function bindSameCAEvent(id,targetNo){
    var el=$(id);
    if(!el){return;}
    var fn=function(){
        syncSameCATarget(targetNo);
        if(targetNo===2){ syncSameCATarget(3); }
    };
    if(el.addEventListener){
        el.addEventListener("input",fn,false);
        el.addEventListener("change",fn,false);
    }else if(el.attachEvent){
        el.attachEvent("onkeyup",fn);
        el.attachEvent("onchange",fn);
    }
}


var modalFocusTarget="";

function showValidationModal(requestNo,message,focusId){
    var overlay=$("modalOverlay");
    var modal=$("validationModal");
    var title=$("modalTitle");
    var msg=$("modalMessage");

    if(!overlay || !modal || !title || !msg){
        alert(message);
        if(focusId){ focusField(focusId); }
        return;
    }

    modalFocusTarget=focusId||"";

    // 色クラスを初期化
    modal.className="validation-modal ";

    if(requestNo===1){
        modal.className+="modal-request1";
        title.innerText="依頼1｜入力確認";
    }else if(requestNo===2){
        modal.className+="modal-request2";
        title.innerText="依頼2｜入力確認";
    }else if(requestNo===3){
        modal.className+="modal-request3";
        title.innerText="依頼3｜入力確認";
    }else{
        modal.className+="modal-neutral";
        title.innerText="入力確認";
    }

    msg.innerText=message;
    overlay.className="modal-overlay";

    try{
        $("modalOkButton").focus();
    }catch(err){}
}

function closeValidationModal(){
    var overlay=$("modalOverlay");

    if(overlay){
        overlay.className="modal-overlay hidden";
    }

    if(modalFocusTarget){
        focusField(modalFocusTarget);
    }

    modalFocusTarget="";
}

function clearErrors(){
    var ids=[
        "requestBase",
        "org1","ca1","proxyOrg1","proxyCA1_1","mailMemo1","processedAt1","type1",
        "org2","ca2","proxyOrg2","proxyCA2_1","mailMemo2","processedAt2","type2",
        "org3","ca3","proxyOrg3","proxyCA3_1","mailMemo3","processedAt3","type3"
    ];

    var i,el;

    for(i=0;i<ids.length;i++){
        el=$(ids[i]);
        if(el){
            el.className=String(el.className||"")
                .replace(/\berror-field\b/g,"")
                .replace(/^\s+|\s+$/g,"");
        }
    }
}

function markError(id){
    var el=$(id);
    if(!el){ return; }

    var cls=String(el.className||"");
    if(cls.indexOf("error-field")<0){
        el.className=(cls+" error-field").replace(/^\s+|\s+$/g,"");
    }
}

function focusField(id){
    try{
        var el=$(id);
        if(el){ el.focus(); }
    }catch(err){}
}

function trimValue(id){
    return val(id).replace(/^\s+|\s+$/g,"");
}


function sendRequest(){
    if(!validateForm()){ return; }

    var requests=[],i,req;
    for(i=1;i<=visibleRequestCount;i++){
        req=collectRequest(i);
        if(i===1||hasAnyInput(req)){ requests.push(req); }
    }

    var requester=$("requesterName").innerText;
    var requestId=createRequestId(requester);
    var sentAt=formatDateTime(new Date());
    var baseName=val("requestBase");
    var lines=[];
    var pendingPath="";

    for(i=0;i<requests.length;i++){
        lines.push(makeCsvLine(requestId,sentAt,baseName,requester,requests[i]));
    }

    try{
        var csvFolder=getProductionCsvFolder(true);

        // 1. まずPendingへ保存。
        pendingPath=savePendingPackage(
            baseName,
            requester,
            requestId,
            sentAt,
            requests,
            lines,
            csvFolder
        );

        // 完了通知用の合図ファイル。Workerはこのファイルが作成されるまで
        // 完了HTAの表示を待つため、受付モーダルとの二重表示を防げる。
        activeCompletionReadyPath=pendingPath+".notify-ready";
        completionReadySignaled=false;
        removeCompletionReadyMarkerIfExists();

        beginPendingState("sending");

        // 2. v28.7: 新規送信も待機なし。
        //    Worker起動後、共有CSVへの書き込みを即時開始する。
        var notBeforeMs=0;

        // 3. Workerは先に別プロセス起動。
        //    フォームを閉じてもWorker自体は継続する。
        launchBackgroundWorker(pendingPath,csvFolder,notBeforeMs);

        // 4. 受付完了を表示。OKでフォーム最小化。
        showAcceptedAndPrepareMinimize();

    }catch(err){
        if(pendingPath && getPendingCount()>0){
            pendingRetryState="failed";
            pendingRetryStartedAt=0;
            updatePendingRetryUI();
        }else{
            pendingRetryState="";
            pendingRetryStartedAt=0;
            updatePendingRetryUI();

            alert(
                "送信受付に失敗しました。\n\n"+
                String(err.message||err.description||err)
            );
        }
    }
}


var pendingCancelRequestNo=0;

function requestCancel(no){
    if(no!==2 && no!==3){
        return;
    }

    var overlay=$("cancelModalOverlay");
    var modal=$("cancelConfirmModal");
    var title=$("cancelModalTitle");
    var message=$("cancelModalMessage");

    pendingCancelRequestNo=no;

    if(!overlay || !modal || !title || !message){
        // 念のためモーダルが使えない場合だけ標準confirmへフォールバック
        if(confirm("依頼"+no+"を取り消しますか？")){
            cancelRequest(no);
        }
        return;
    }

    modal.className="validation-modal cancel-confirm-modal ";

    if(no===2){
        modal.className+="modal-request2";
        title.innerText="依頼2｜取り消し確認";
        if(visibleRequestCount>=3){
            message.innerText="依頼2を取り消しますか？\n\n依頼3の内容は依頼2へ繰り上げます。";
        }else{
            message.innerText="依頼2を取り消しますか？\n\n入力した内容はクリアされます。";
        }
    }else{
        modal.className+="modal-request3";
        title.innerText="依頼3｜取り消し確認";
        message.innerText="依頼3を取り消しますか？\n\n入力した内容はクリアされます。";
    }

    overlay.className="modal-overlay";

    try{
        $("cancelBackButton").focus();
    }catch(err){}
}

function closeCancelConfirm(){
    var overlay=$("cancelModalOverlay");

    if(overlay){
        overlay.className="modal-overlay hidden";
    }

    pendingCancelRequestNo=0;
}

function confirmCancelRequest(){
    var no=pendingCancelRequestNo;

    closeCancelConfirm();

    if(no===2 || no===3){
        cancelRequest(no);
    }
}

function cancelRequest(no){
    if(no!==2 && no!==3){
        return;
    }

    hideAppForLayout();

    if(no===2){
        if(visibleRequestCount>=3){
            // 依頼3が存在する場合：
            // 依頼3の内容を依頼2へ繰り上げて、依頼3だけを閉じる
            copyRequestFields(3,2);
            clearRequestFields(3);

            addHiddenClass("request3");

            // 新しい依頼3を追加できるようにボタンを復活
            if($("btnAdd3")){
                $("btnAdd3").style.display="";
            }

            // 依頼2は表示したまま
            visibleRequestCount=2;

        }else{
            // 依頼3が存在しない場合：
            // 依頼2を普通にクリアして閉じる
            clearRequestFields(2);
            clearRequestFields(3);

            addHiddenClass("request2");
            addHiddenClass("request3");

            if($("btnAdd2")){
                $("btnAdd2").style.display="";
            }

            if($("btnAdd3")){
                $("btnAdd3").style.display="";
            }

            visibleRequestCount=1;
        }

    }else{
        // 依頼3のみ取消
        clearRequestFields(3);
        addHiddenClass("request3");

        if($("btnAdd3")){
            $("btnAdd3").style.display="";
        }

        visibleRequestCount=2;
    }

    clearErrors();
    resizeApp(false);
}

function copyRequestFields(fromNo,toNo){
    setValue("org"+toNo,val("org"+fromNo));
    setValue("ca"+toNo,val("ca"+fromNo));
    setValue("proxyOrg"+toNo,val("proxyOrg"+fromNo));
    setValue("proxyCA"+toNo+"_1",val("proxyCA"+fromNo+"_1"));
    setValue("proxyCA"+toNo+"_2",val("proxyCA"+fromNo+"_2"));
    setValue("proxyCA"+toNo+"_3",val("proxyCA"+fromNo+"_3"));
    setValue("mailMemo"+toNo,val("mailMemo"+fromNo));
    setValue("processedAt"+toNo,val("processedAt"+fromNo));
    setValue("dueDate"+toNo,val("dueDate"+fromNo));

    setChecked("short"+toNo,checked("short"+fromNo));
    setChecked("urgent"+toNo,checked("urgent"+fromNo));

    if($("type"+toNo) && $("type"+fromNo)){
        $("type"+toNo).value=$("type"+fromNo).value;
    }
}


function addHiddenClass(id){
    var el=$(id);
    if(!el){ return; }

    if(String(el.className).indexOf("hidden")<0){
        el.className+=" hidden";
    }
}

function clearRequestFields(i){
    setValue("org"+i,"");
    setValue("ca"+i,"");
    setValue("proxyOrg"+i,"");
    setValue("proxyCA"+i+"_1","");
    setValue("proxyCA"+i+"_2","");
    setValue("proxyCA"+i+"_3","");
    setValue("mailMemo"+i,"");
    setValue("processedAt"+i,"");
    setValue("dueDate"+i,"");

    setChecked("short"+i,false);
    setChecked("urgent"+i,false);

    if($("type"+i)){
        $("type"+i).selectedIndex=0;
    }
}

function resetAfterSend(){
    var i;

    for(i=1;i<=3;i++){
        clearRequestFields(i);
    }

    // 拠点・依頼者は維持
    addHiddenClass("request2");
    addHiddenClass("request3");

    if($("btnAdd2")){
        $("btnAdd2").style.display="";
    }

    if($("btnAdd3")){
        $("btnAdd3").style.display="";
    }

    visibleRequestCount=1;

    clearErrors();
    setStatus("");

    try{
        $("org1").focus();
    }catch(err){}

    hideAppForLayout();
    resizeApp(false);
}

function setValue(id,value){
    var el=$(id);
    if(el){
        el.value=value;
    }
}

function setChecked(id,value){
    var el=$(id);
    if(el){
        el.checked=value;
    }
}





function hideSystemModal(){
    var overlay=$("systemModalOverlay");

    if(overlay){
        overlay.className="system-modal-overlay hidden";
    }
}

function isSystemModalVisible(){
    var overlay=$("systemModalOverlay");
    if(!overlay){ return false; }
    return String(overlay.className||"").indexOf("hidden")<0;
}

function showSystemModal(title,message,kind,showButton){
    var overlay=$("systemModalOverlay");
    var modal=$("systemModal");
    var titleEl=$("systemModalTitle");
    var messageEl=$("systemModalMessage");
    var button=$("systemModalOkButton");

    if(!overlay || !modal || !titleEl || !messageEl){
        alert(message);
        return;
    }

    modal.className="system-modal ";

    if(kind==="accepted"){
        modal.className+="system-modal-accepted";
    }else{
        modal.className+="system-modal-neutral";
    }

    titleEl.innerText=title;
    messageEl.innerText=message;

    if(button){
        button.style.display=showButton?"inline-block":"none";
    }

    overlay.className="system-modal-overlay";
}

function getCurrentFolderPath(){
    var rawPath="";
    var host="";
    var path="";

    try{
        rawPath=String(window.location.pathname||"");
    }catch(err){
        rawPath="";
    }

    try{
        host=String(window.location.hostname||"");
    }catch(err2){
        host="";
    }

    try{
        path=decodeURIComponent(rawPath);
    }catch(err3){
        path=rawPath;
    }

    path=path.replace(/\//g,"\\");

    // UNC起動:
    // file://server/share/folder/file.hta
    // pathnameは \share\folder\file.hta となり、
    // server名はlocation.hostname側に入るため結合する。
    if(host && host.toLowerCase()!=="localhost"){
        while(path.charAt(0)==="\\"){
            path=path.substring(1);
        }
        path="\\\\"+host+"\\"+path;
    }else{
        // ローカル:
        // file:///C:/folder/file.hta → \C:\folder\file.hta
        if(/^\\[A-Za-z]:\\/.test(path)){
            path=path.substring(1);
        }
    }

    var fso=new ActiveXObject("Scripting.FileSystemObject");
    return fso.GetParentFolderName(path);
}


function getProductionCsvFolder(required){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var baseFolder=getCurrentFolderPath();
    var csvFolder=fso.BuildPath(baseFolder,CSV_SUBFOLDER_NAME);

    if(!fso.FolderExists(csvFolder)){
        if(required){
            throw new Error(
                "CSV保存先フォルダが見つかりません。\n\n"+
                "メンション依頼フォームと同じ場所に「"+CSV_SUBFOLDER_NAME+"」フォルダを作成してください。\n\n"+
                "確認先：\n"+csvFolder
            );
        }
        return "";
    }

    return csvFolder;
}

function quoteCommandArgument(value){
    var s=String(value==null?"":value);
    return '"'+s.replace(/"/g,'""')+'"';
}

function expandEnvironmentPath(pathValue){
    var shell=new ActiveXObject("WScript.Shell");
    return shell.ExpandEnvironmentStrings(String(pathValue||""));
}


var pendingWatchTimer=null;
var pendingRetryState="";
var pendingRetryStartedAt=0;
var pendingSuccessHideTimer=null;

function getPendingFilePaths(){
    var result=[];
    var fso,shell,pendingFolder,folder,files,enumerator,file;

    try{
        fso=new ActiveXObject("Scripting.FileSystemObject");
        shell=new ActiveXObject("WScript.Shell");
        pendingFolder=String(shell.SpecialFolders("Desktop"))+"\\"+PENDING_FOLDER_NAME;

        if(!fso.FolderExists(pendingFolder)){
            return result;
        }

        folder=fso.GetFolder(pendingFolder);
        files=folder.Files;
        enumerator=new Enumerator(files);

        for(;!enumerator.atEnd();enumerator.moveNext()){
            file=enumerator.item();
            var ext=String(fso.GetExtensionName(file.Name)).toLowerCase();
            if(ext==="csv"){
                result.push(String(file.Path));
            }
        }
    }catch(err){
    }

    return result;
}

function getPendingCount(){
    return getPendingFilePaths().length;
}

function setSendButtonVisible(visible){
    var send=$("sendButton");
    if(!send){return;}
    send.style.display=visible ? "inline-block" : "none";
}

function setRetryButtonVisible(visible){
    var button=$("retryPendingButton");
    if(!button){return;}
    button.style.display=visible ? "inline-block" : "none";
}

function beginPendingState(state){
    pendingRetryState=state||"";
    pendingRetryStartedAt=(new Date()).getTime();
    updatePendingRetryUI();
}

function updatePendingRetryUI(){
    var area=$("pendingRetryArea");
    var text=$("pendingRetryText");
    var count=getPendingCount();
    var now=(new Date()).getTime();
    var elapsed=pendingRetryStartedAt>0 ? (now-pendingRetryStartedAt) : 0;
    var maxMs=WORKER_MAX_RETRY_SECONDS*1000;

    if(!area || !text){
        return;
    }

    // 送信・再送していたPendingが消えた＝worker側で処理完了
    if(count<=0){
        var completedState=pendingRetryState;

        if(
            completedState==="sending" ||
            completedState==="auto" ||
            completedState==="manual"
        ){
            pendingRetryState="";
            pendingRetryStartedAt=0;
        }

        area.className="pending-retry-area hidden";
        text.className="pending-retry-text";
        text.innerText="";
        setRetryButtonVisible(false);
        setSendButtonVisible(true);
        return;
    }

    // 自動処理開始から300秒以上Pendingが残っている場合は失敗状態
    if(
        (pendingRetryState==="sending" ||
         pendingRetryState==="auto" ||
         pendingRetryState==="manual") &&
        pendingRetryStartedAt>0 &&
        elapsed>=maxMs
    ){
        pendingRetryState="failed";
        pendingRetryStartedAt=0;
    }

    area.className="pending-retry-area";

    if(pendingRetryState==="sending"){
        text.className="pending-retry-text";
        text.innerText="送信中です…";
        setRetryButtonVisible(false);
        setSendButtonVisible(true);

    }else if(pendingRetryState==="auto"){
        text.className="pending-retry-text";
        text.innerText="未送信データを自動再送中です…";
        setRetryButtonVisible(false);
        setSendButtonVisible(true);

    }else if(pendingRetryState==="manual"){
        text.className="pending-retry-text";
        text.innerText="再送中です…";
        setRetryButtonVisible(false);
        setSendButtonVisible(true);

    }else{
        // Pendingはあるが処理中ではない＝送信失敗
        pendingRetryState="failed";
        text.className="pending-retry-text failed";
        text.innerText="送信できませんでした│未送信 "+count+"件";
        setRetryButtonVisible(true);
        setSendButtonVisible(true);
    }
}

function startPendingWatcher(){
    if(pendingWatchTimer){
        window.clearInterval(pendingWatchTimer);
    }

    updatePendingRetryUI();

    pendingWatchTimer=window.setInterval(function(){
        updatePendingRetryUI();
    },1500);
}

function launchAllPendingWorkers(mode){
    var files=getPendingFilePaths();
    var csvFolder=getProductionCsvFolder(true);
    var i,started=0;

    if(files.length===0){
        updatePendingRetryUI();
        return 0;
    }

    beginPendingState(mode||"auto");

    for(i=0;i<files.length;i++){
        try{
            launchBackgroundWorker(files[i],csvFolder);
            started++;
        }catch(err){
            // 他のPendingの再送は継続
        }
    }

    return started;
}

function recoverPendingPackagesOnStartup(){
    try{
        launchAllPendingWorkers("auto");
    }catch(err){
        // CSV保存先が未設定などの場合はPendingを残して手動再送待ち
        pendingRetryState="failed";
        pendingRetryStartedAt=0;
        updatePendingRetryUI();
    }
}

function retryPendingPackages(){
    var count=getPendingCount();

    if(count<=0){
        updatePendingRetryUI();
        return;
    }

    try{
        var started=launchAllPendingWorkers("manual");

        if(started<=0){
            throw new Error("再送処理を開始できませんでした。");
        }

    }catch(err){
        pendingRetryState="failed";
        pendingRetryStartedAt=0;
        updatePendingRetryUI();

        alert(
            "未送信データの再送を開始できませんでした。\n\n"+
            String(err.message||err.description||err)
        );
    }
}

function launchBackgroundWorker(pendingPath,csvFolderFallback,notBeforeMs){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var shell=new ActiveXObject("WScript.Shell");
    var workerPath=getCurrentFolderPath()+"\\"+BACKGROUND_WORKER_NAME;

    pendingPath=String(pendingPath||"");

    if(!pendingPath || !fso.FileExists(pendingPath)){
        throw new Error("退避CSVの引き渡しに失敗しました。");
    }

    if(!fso.FileExists(workerPath)){
        throw new Error(BACKGROUND_WORKER_NAME+" が見つかりません。");
    }

    csvFolderFallback=String(csvFolderFallback||"");
    notBeforeMs=parseInt(notBeforeMs,10);
    if(isNaN(notBeforeMs) || notBeforeMs<0){
        notBeforeMs=0;
    }

    /*
      v28.11:
      Pure JScript Workerを別プロセスで非同期起動。
      受付OKまたはフォーム終了後、書き込み完了時に
      WorkerからWindows標準Popupを表示する。
      第3引数は互換性のため残すが、通常は0を渡して即時処理する。
    */
    var command=
        "wscript.exe //NoLogo //E:JScript "+
        quoteCommandArgument(workerPath)+" "+
        quoteCommandArgument(pendingPath)+" "+
        quoteCommandArgument(csvFolderFallback)+" "+
        quoteCommandArgument(String(notBeforeMs));

    var result=shell.Run(command,0,false);

    if(result!==0){
        throw new Error("バックグラウンド送信処理を起動できませんでした。");
    }
}

function getPendingFolderPath(){
    var shell=new ActiveXObject("WScript.Shell");
    var desktop=String(shell.SpecialFolders("Desktop"));
    return desktop+"\\"+PENDING_FOLDER_NAME;
}

function savePendingPackage(baseName,requester,requestId,sentAt,requests,lines,csvFolder){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var folderPath=getPendingFolderPath();

    ensureFolder(fso,folderPath);

    var fileName=sanitizeFileName(requestId)+".csv";
    var filePath=folderPath+"\\"+fileName;

    // 万一同名があれば上書きせず別名にする
    if(fso.FileExists(filePath)){
        fileName=sanitizeFileName(requestId)+"_"+String(new Date().getTime())+".csv";
        filePath=folderPath+"\\"+fileName;
    }

    /*
      v27:
      退避ファイルそのものをCSV化。
      日本語WindowsのANSI(CP932)で保存し、
      共有CSVへ反映する予定の行をそのまま保持する。
    */
    var file=fso.CreateTextFile(filePath,false,false);
    var i;

    file.WriteLine(makeHeaderLine());

    for(i=0;i<lines.length;i++){
        file.WriteLine(String(lines[i]||""));
    }

    file.Close();
    return filePath;
}

function ensureFolder(fso,folderPath){
    if(fso.FolderExists(folderPath)){ return; }

    var parent=fso.GetParentFolderName(folderPath);

    if(parent && !fso.FolderExists(parent)){
        ensureFolder(fso,parent);
    }

    fso.CreateFolder(folderPath);
}

function makeHeaderLine(){
    return "送信日時,RequestID,拠点,依頼者,依頼番号,組織,CA名,代理CA組織,代理CA1,代理CA2,代理CA3,メールメモ,処理日時,期日,時短,至急,タイプ";
}

function makeCsvLine(requestId,sentAt,baseName,requester,req){
    return csvJoin([
        sentAt,requestId,baseName,requester,String(req.requestNo),
        req.organization,req.caName,req.proxyOrganization,
        req.proxyCA1,req.proxyCA2,req.proxyCA3,
        req.mailMemo,req.processedAt,req.dueDate,
        req.shortTime?"●":"",
        req.urgent?"●":"",
        req.type
    ]);
}

function csvJoin(values){
    var out=[],i;
    for(i=0;i<values.length;i++){ out.push(csvEscape(values[i])); }
    return out.join(",");
}

function csvEscape(value){
    var s=String(value==null?"":value);
    s=s.replace(/"/g,'""');
    return '"'+s+'"';
}

function createRequestId(requester){
    var now=new Date();
    var userPart=String(requester||"USER")
        .replace(/[\\\/:*?"<>|\s]/g,"")
        .substring(0,12);

    var rand1=Math.floor(Math.random()*0x1000000).toString(16).toUpperCase();
    while(rand1.length<6){ rand1="0"+rand1; }

    var milli=String(now.getMilliseconds());
    while(milli.length<3){ milli="0"+milli; }

    return formatDate(now)+"-"+
           pad2(now.getHours())+pad2(now.getMinutes())+pad2(now.getSeconds())+milli+"-"+
           userPart+"-"+rand1;
}

function formatDate(d){
    return d.getFullYear()+pad2(d.getMonth()+1)+pad2(d.getDate());
}

function formatDateTime(d){
    return d.getFullYear()+"/"+pad2(d.getMonth()+1)+"/"+pad2(d.getDate())+" "+
           pad2(d.getHours())+":"+pad2(d.getMinutes())+":"+pad2(d.getSeconds());
}

function pad2(n){ return n<10?"0"+n:String(n); }

function sanitizeFileName(name){
    return String(name||"未選択").replace(/[\\\/:*?"<>|]/g,"_");
}

function removeCompletionReadyMarkerIfExists(){
    if(!activeCompletionReadyPath){return;}

    try{
        var fso=new ActiveXObject("Scripting.FileSystemObject");
        if(fso.FileExists(activeCompletionReadyPath)){
            fso.DeleteFile(activeCompletionReadyPath,true);
        }
    }catch(err){}
}

function signalCompletionNotificationReady(){
    if(!activeCompletionReadyPath || completionReadySignaled){
        return;
    }

    try{
        var fso=new ActiveXObject("Scripting.FileSystemObject");
        var file=fso.CreateTextFile(activeCompletionReadyPath,true,false);
        file.WriteLine("ready");
        file.Close();
        completionReadySignaled=true;
    }catch(err){
        // 完了通知用の合図に失敗しても、CSV送信処理自体には影響させない。
    }
}

function setStatus(message){
    var el=$("statusText");
    if(el){
        el.innerText=message||"";
    }
}
