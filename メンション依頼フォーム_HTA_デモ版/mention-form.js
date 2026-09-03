var PRODUCTION_CSV_FOLDER = "【本番CSV保存先のパスをここに設定】";
var PENDING_FOLDER_NAME = "MentionRequest_Pending";
var BACKGROUND_WORKER_NAME = "mention-request-worker.ps1";
var WORKER_RETRY_MIN_MS = 450;
var WORKER_RETRY_MAX_MS = 1250;
var WORKER_MAX_RETRY_SECONDS = 300;

var TYPE_OPTIONS = [
    "",
    "通常",
    "確認",
    "代理対応",
    "その他",
    "【長文サンプル】候補者・企業双方への確認が必要なため、処理前に担当者へエスカレーションする"
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
    populateTypes();
    setRequesterName();

    try{
        getProductionCsvFolder(true);
        setStatus("");
    }catch(configErr){
        setStatus("本番CSV保存先が未設定です");
    }

    lastAvailWidth=screen.availWidth;
    lastAvailHeight=screen.availHeight;

    // 初回は画面を見せる前にサイズを確定
    resizeApp(true);

    // HTAを開いたまま解像度・表示倍率・接続モニターが変わった場合にも追従
    startScreenWatcher();

    bindSameCASync();

    // Pendingの有無を監視し、再送UIを自動更新
    startPendingWatcher();

    // 前回終了時などに残ったPendingを自動再送
    recoverPendingPackagesOnStartup();
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

function populateTypes(){
    var i,j,sel,opt;
    for(i=1;i<=3;i++){
        sel=$("type"+i);
        while(sel.options.length>0){ sel.remove(0); }
        for(j=0;j<TYPE_OPTIONS.length;j++){
            opt=document.createElement("option");
            opt.value=TYPE_OPTIONS[j];
            opt.text=TYPE_OPTIONS[j];
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
            if(!app){ return; }

            // 一旦100%へ戻して、フォーム本来の寸法を測定
            document.body.style.zoom="1";

            var appStyle=null;
            var marginLeft=14;
            var marginRight=14;
            var marginTop=14;
            var marginBottom=16;

            /*
              getComputedStyleが利用できる環境では実値を使う。
              HTA/IE側で取得できない場合はCSS設定値をfallbackにする。
            */
            try{
                if(window.getComputedStyle){
                    appStyle=window.getComputedStyle(app,null);
                    if(appStyle){
                        marginLeft=parseInt(appStyle.marginLeft,10)||14;
                        marginRight=parseInt(appStyle.marginRight,10)||14;
                        marginTop=parseInt(appStyle.marginTop,10)||14;
                        marginBottom=parseInt(appStyle.marginBottom,10)||16;
                    }
                }else if(app.currentStyle){
                    appStyle=app.currentStyle;
                    marginLeft=parseInt(appStyle.marginLeft,10)||14;
                    marginRight=parseInt(appStyle.marginRight,10)||14;
                    marginTop=parseInt(appStyle.marginTop,10)||14;
                    marginBottom=parseInt(appStyle.marginBottom,10)||16;
                }
            }catch(styleErr){
            }

            var naturalW=app.offsetWidth+marginLeft+marginRight;
            var naturalH=app.offsetHeight+marginTop+marginBottom;

            var chrome=getCurrentChromeSize();

            // 画面端から少し余裕を残す
            var screenMarginW=28;
            var screenMarginH=34;

            var maxOuterW=screen.availWidth-screenMarginW;
            var maxOuterH=screen.availHeight-screenMarginH;

            var maxClientW=maxOuterW-chrome.width;
            var maxClientH=maxOuterH-chrome.height;

            var fitW=maxClientW/naturalW;
            var fitH=maxClientH/naturalH;
            var fitScale=Math.min(1,fitW,fitH);

            /*
              1920×1080級では約96%を上限の目安としつつ、
              実際に画面へ収まる倍率を最優先する。
            */
            var comfortScale=maxOuterW/1960;
            if(comfortScale>1){comfortScale=1;}

            var scale=Math.min(fitScale,comfortScale);

            if(scale<=0 || isNaN(scale)){
                scale=1;
            }

            document.body.style.zoom=String(scale);

            /*
              v13では左右14pxの余白も含めた「実際のフォーム全幅」で
              viewportサイズを決定する。
            */
            var targetClientW=Math.ceil(naturalW*scale);
            var targetClientH=Math.ceil(naturalH*scale);

            var wantedW=Math.ceil(targetClientW+chrome.width);
            var wantedH=Math.ceil(targetClientH+chrome.height);

            if(wantedW>maxOuterW){wantedW=maxOuterW;}
            if(wantedH>maxOuterH){wantedH=maxOuterH;}

            window.resizeTo(wantedW,wantedH);
            centerCurrentWindow(wantedW,wantedH);
            window.scrollTo(0,0);

            /*
              Windows表示倍率・テーマ・HTA枠幅による数pxの差を、
              resize後の実viewportを測って1回だけ補正する。
              固定のchrome幅だけに依存しないため、
              24インチ/125%等でも左右余白が揃いやすい。
            */
            window.setTimeout(function(){
                try{
                    var actualViewport=getViewportSize();
                    var actualOuter=getWindowOuterSize(wantedW,wantedH);

                    if(actualViewport.width<=0 || actualViewport.height<=0){
                        return;
                    }

                    var diffW=targetClientW-actualViewport.width;
                    var diffH=targetClientH-actualViewport.height;

                    // 極端な補正を避ける
                    if(diffW>80){diffW=80;}
                    if(diffW<-80){diffW=-80;}
                    if(diffH>80){diffH=80;}
                    if(diffH<-80){diffH=-80;}

                    if(Math.abs(diffW)>=2 || Math.abs(diffH)>=2){
                        var correctedW=Math.ceil(actualOuter.width+diffW);
                        var correctedH=Math.ceil(actualOuter.height+diffH);

                        if(correctedW>maxOuterW){correctedW=maxOuterW;}
                        if(correctedH>maxOuterH){correctedH=maxOuterH;}

                        window.resizeTo(correctedW,correctedH);
                        centerCurrentWindow(correctedW,correctedH);
                    }
                }catch(correctionErr){
                }
            },60);

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
        // 本番CSV格納先はプログラム内の固定パスを使用
        var csvFolder=getProductionCsvFolder(true);

        // ① まずPendingへ確実に保存
        pendingPath=savePendingPackage(
            baseName,
            requester,
            requestId,
            sentAt,
            requests,
            lines,
            csvFolder
        );

        // Pending保存後は送信ボタンを隠し、メッセージのみ表示
        beginPendingState("sending");

        // ② 別プロセスのworkerへ引き渡す
        launchBackgroundWorker(pendingPath,csvFolder);

        // ③ ここからHTAを閉じてもworkerは継続
        resetAfterSend();

        // 受付完了モーダルは表示しない。
        // footerの「送信中です…」表示とWindows通知で状態を伝える。

    }catch(err){
        if(pendingPath && getPendingCount()>0){
            // Pending保存後にworker起動などで失敗した場合は再送待ちへ
            pendingRetryState="failed";
            pendingRetryStartedAt=0;
            updatePendingRetryUI();
        }else{
            // Pending自体を保存できなかった場合は通常の送信ボタンへ戻す
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
    var raw=String(PRODUCTION_CSV_FOLDER||"").replace(/^\s+|\s+$/g,"");

    if(
        !raw ||
        raw.indexOf("本番CSV保存先のパスをここに設定")>=0
    ){
        if(required){
            throw new Error(
                "本番CSV保存先が未設定です。\n\n"+
                "管理者用のPRODUCTION_CSV_FOLDERへ共有フォルダのパスを設定してください。"
            );
        }
        return "";
    }

    return expandEnvironmentPath(raw);
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
            if(String(fso.GetExtensionName(file.Name)).toLowerCase()==="pending"){
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
        if(
            pendingRetryState==="sending" ||
            pendingRetryState==="auto" ||
            pendingRetryState==="manual"
        ){
            pendingRetryState="";
            pendingRetryStartedAt=0;
        }

        area.className="pending-retry-area hidden";
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
        text.innerText="送信中です… 完了後、Windows通知でお知らせします。";
        setRetryButtonVisible(false);
        setSendButtonVisible(true);

    }else if(pendingRetryState==="auto"){
        text.innerText="未送信データを自動再送中です…";
        setRetryButtonVisible(false);
        setSendButtonVisible(true);

    }else if(pendingRetryState==="manual"){
        text.innerText="再送中です… 完了後、Windows通知でお知らせします。";
        setRetryButtonVisible(false);
        setSendButtonVisible(true);

    }else{
        // Pendingはあるが処理中ではない＝送信失敗
        pendingRetryState="failed";
        text.innerText="⚠ 送信できませんでした｜未送信 "+count+"件";
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

function launchBackgroundWorker(pendingPath,csvFolderFallback){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var shell=new ActiveXObject("WScript.Shell");
    var workerPath=getCurrentFolderPath()+"\\"+BACKGROUND_WORKER_NAME;

    pendingPath=String(pendingPath||"");

    if(!pendingPath || !fso.FileExists(pendingPath)){
        throw new Error("Pendingファイルの引き渡しに失敗しました。");
    }

    if(!fso.FileExists(workerPath)){
        throw new Error(BACKGROUND_WORKER_NAME+" が見つかりません。");
    }

    csvFolderFallback=String(csvFolderFallback||"");

    var command=
        "powershell.exe "+
        "-NoProfile "+
        "-NonInteractive "+
        "-ExecutionPolicy Bypass "+
        "-WindowStyle Hidden "+
        "-File "+quoteCommandArgument(workerPath)+" "+
        "-PendingPath "+quoteCommandArgument(pendingPath)+" "+
        "-CsvFolder "+quoteCommandArgument(csvFolderFallback)+" "+
        "-RetryMinMilliseconds "+String(WORKER_RETRY_MIN_MS)+" "+
        "-RetryMaxMilliseconds "+String(WORKER_RETRY_MAX_MS)+" "+
        "-MaxRetrySeconds "+String(WORKER_MAX_RETRY_SECONDS);

    // 非同期起動。HTAを閉じてもworkerは継続。
    var result=shell.Run(command,0,false);

    if(result!==0){
        throw new Error("バックグラウンド送信処理を起動できませんでした。");
    }
}

function showAcceptedMessage(){
    try{
        showSystemModal(
            "送信完了",
            "メンション依頼を送信しました。\n\n受け付け完了後、右下のWindows通知でお知らせいたします。",
            "accepted",
            true
        );
    }catch(err){
        alert(
            "メンション依頼を送信しました。\n\n"+
            "受け付け完了後、右下のWindows通知でお知らせいたします。"
        );
    }
}

function getPendingFolderPath(){
    var shell=new ActiveXObject("WScript.Shell");
    var desktop=String(shell.SpecialFolders("Desktop"));
    return desktop+"\\"+PENDING_FOLDER_NAME;
}

function pendingDisplayValue(value){
    var s=String(value==null?"":value);

    // 万一改行が入っていても、人間向け表示が崩れにくいよう整形
    s=s.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
    s=s.replace(/\n/g," / ");

    return s;
}

function writePendingRequestReadable(file,req){
    file.WriteLine("【依頼 "+req.requestNo+"】");
    file.WriteLine("組織　　　　："+pendingDisplayValue(req.organization));
    file.WriteLine("CA名　　　　："+pendingDisplayValue(req.caName));
    file.WriteLine("代理CA組織　："+pendingDisplayValue(req.proxyOrganization));
    file.WriteLine("代理CA名1　 ："+pendingDisplayValue(req.proxyCA1));
    file.WriteLine("代理CA名2　 ："+pendingDisplayValue(req.proxyCA2));
    file.WriteLine("代理CA名3　 ："+pendingDisplayValue(req.proxyCA3));
    file.WriteLine("メールメモ　："+pendingDisplayValue(req.mailMemo));
    file.WriteLine("処理日時　　："+pendingDisplayValue(req.processedAt));
    file.WriteLine("期日　　　　："+pendingDisplayValue(req.dueDate));
    file.WriteLine("時短　　　　："+(req.shortTime?"●":""));
    file.WriteLine("至急　　　　："+(req.urgent?"●":""));
    file.WriteLine("タイプ　　　："+pendingDisplayValue(req.type));
    file.WriteLine("");
}

function savePendingPackage(baseName,requester,requestId,sentAt,requests,lines,csvFolder){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var folderPath=getPendingFolderPath();

    ensureFolder(fso,folderPath);

    var fileName=sanitizeFileName(requestId)+".pending";
    var filePath=folderPath+"\\"+fileName;

    // 万一同名があれば上書きせず別名にする
    if(fso.FileExists(filePath)){
        fileName=sanitizeFileName(requestId)+"_"+String(new Date().getTime())+".pending";
        filePath=folderPath+"\\"+fileName;
    }

    // Unicodeで保存。Windowsのメモ帳でそのまま日本語を確認できる。
    var file=fso.CreateTextFile(filePath,false,true);

    // =====================================================
    // 人が確認するための表示エリア
    // =====================================================
    file.WriteLine("メンション依頼フォーム｜未送信データ");
    file.WriteLine("============================================================");
    file.WriteLine("状態　　　　：未送信");
    file.WriteLine("保存日時　　："+pendingDisplayValue(sentAt));
    file.WriteLine("RequestID　 ："+pendingDisplayValue(requestId));
    file.WriteLine("拠点　　　　："+pendingDisplayValue(baseName));
    file.WriteLine("依頼者　　　："+pendingDisplayValue(requester));
    file.WriteLine("依頼件数　　："+String(requests.length));
    file.WriteLine("送信先　　　："+pendingDisplayValue(csvFolder));
    file.WriteLine("============================================================");
    file.WriteLine("");

    var i;
    for(i=0;i<requests.length;i++){
        writePendingRequestReadable(file,requests[i]);
    }

    file.WriteLine("============================================================");
    file.WriteLine("※ここから下はCSV再送用の機械データです。");
    file.WriteLine("※編集・削除しないでください。");
    file.WriteLine("----- MACHINE_DATA_BEGIN -----");
    file.WriteLine("FORMAT=MENTION_REQUEST_PENDING_V3");
    file.WriteLine("BASE="+encodeURIComponent(String(baseName||"")));
    file.WriteLine("REQUEST_ID="+encodeURIComponent(String(requestId||"")));
    file.WriteLine("CREATED="+encodeURIComponent(String(sentAt||"")));
    file.WriteLine("CSV_FOLDER="+encodeURIComponent(String(csvFolder||"")));

    for(i=0;i<lines.length;i++){
        // CSV行内に記号や改行があっても壊れないようURLエンコード
        file.WriteLine("LINE="+encodeURIComponent(String(lines[i]||"")));
    }

    file.WriteLine("----- MACHINE_DATA_END -----");

    file.Close();
    return filePath;
}

function readPendingPackage(filePath){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var file=fso.OpenTextFile(filePath,1,false,-1);
    var text=file.ReadAll();
    file.Close();

    var rows=String(text)
        .replace(/\r\n/g,"\n")
        .replace(/\r/g,"\n")
        .split("\n");

    var data={
        filePath:filePath,
        baseName:"",
        requestId:"",
        created:"",
        csvFolder:"",
        lines:[]
    };

    var format="";
    var i,row;

    // V1実験版も、今回のV2可読版も読めるようにする
    if(rows.length>0 && rows[0]==="MENTION_REQUEST_PENDING_V1"){
        format="MENTION_REQUEST_PENDING_V1";
    }

    for(i=0;i<rows.length;i++){
        row=rows[i];

        if(row.indexOf("FORMAT=")===0){
            format=row.substring(7);
        }else if(row.indexOf("BASE=")===0){
            data.baseName=decodeURIComponent(row.substring(5));
        }else if(row.indexOf("REQUEST_ID=")===0){
            data.requestId=decodeURIComponent(row.substring(11));
        }else if(row.indexOf("CREATED=")===0){
            data.created=decodeURIComponent(row.substring(8));
        }else if(row.indexOf("CSV_FOLDER=")===0){
            data.csvFolder=decodeURIComponent(row.substring(11));
        }else if(row.indexOf("LINE=")===0){
            data.lines.push(decodeURIComponent(row.substring(5)));
        }
    }

    if(format!=="MENTION_REQUEST_PENDING_V1" &&
       format!=="MENTION_REQUEST_PENDING_V2" &&
       format!=="MENTION_REQUEST_PENDING_V3"){
        throw new Error("一時ファイルの形式が不正です："+fso.GetFileName(filePath));
    }

    if(!data.baseName){
        throw new Error("拠点情報がありません："+fso.GetFileName(filePath));
    }

    if(data.lines.length===0){
        throw new Error("CSV書き込みデータがありません："+fso.GetFileName(filePath));
    }

    return data;
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
    return csvJoin([
        "送信日時","RequestID","拠点","依頼者","依頼番号",
        "組織","CA名","代理CA組織","代理CA1","代理CA2","代理CA3",
        "メールメモ","処理日時","期日","時短","至急","タイプ"
    ]);
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

function setStatus(message){
    var el=$("statusText");
    if(el){
        el.innerText=message||"";
    }
}
