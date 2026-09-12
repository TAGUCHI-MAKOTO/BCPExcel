// v28.23: CSV列順・依頼別RequestID・マルチモニター位置保持・同一CAオプション非連動
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
var lastAvailLeft = 0;
var lastAvailTop = 0;
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

    var initialWorkArea=getCurrentWorkArea();
    lastAvailLeft=initialWorkArea.left;
    lastAvailTop=initialWorkArea.top;
    lastAvailWidth=initialWorkArea.width;
    lastAvailHeight=initialWorkArea.height;

    // 初回は画面を見せる前にサイズを確定
    resizeApp(true);

    // HTAを開いたまま解像度・表示倍率・接続モニターが変わった場合にも追従
    startScreenWatcher();

    bindSameCASync();
    for(var no=1;no<=3;no++){ updateAttendanceUI(no); }

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
            var area=getCurrentWorkArea();

            if(area.left!==lastAvailLeft ||
               area.top!==lastAvailTop ||
               area.width!==lastAvailWidth ||
               area.height!==lastAvailHeight){
                lastAvailLeft=area.left;
                lastAvailTop=area.top;
                lastAvailWidth=area.width;
                lastAvailHeight=area.height;

                // 別モニターへ移動した場合や、解像度・表示倍率・タスクバー領域が変わった場合も
                // 現在いるモニターの作業領域内へ再フィットする。
                resizeApp(false);
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

function isFiniteScreenNumber(value){
    var n=Number(value);
    return !isNaN(n) && isFinite(n);
}

function getCurrentWindowPosition(){
    var left=0;
    var top=0;

    try{
        if(isFiniteScreenNumber(window.screenLeft)){
            left=Number(window.screenLeft);
        }else if(isFiniteScreenNumber(window.screenX)){
            left=Number(window.screenX);
        }
    }catch(errLeft){
    }

    try{
        if(isFiniteScreenNumber(window.screenTop)){
            top=Number(window.screenTop);
        }else if(isFiniteScreenNumber(window.screenY)){
            top=Number(window.screenY);
        }
    }catch(errTop){
    }

    return {left:left,top:top};
}

function getCurrentWorkArea(){
    var width=0;
    var height=0;
    var left=0;
    var top=0;

    try{
        width=isFiniteScreenNumber(screen.availWidth) ? Number(screen.availWidth) : Number(screen.width||0);
        height=isFiniteScreenNumber(screen.availHeight) ? Number(screen.availHeight) : Number(screen.height||0);

        /*
          v28.26:
          availLeft / availTop はWindowsの仮想スクリーン座標上で、
          現在のモニターの「タスクバー等を除いた作業領域」の左上を返す。
          IE/HTA環境で未提供の場合は screen.left / screen.top をfallbackにする。
        */
        if(isFiniteScreenNumber(screen.availLeft)){
            left=Number(screen.availLeft);
        }else if(isFiniteScreenNumber(screen.left)){
            left=Number(screen.left);
        }

        if(isFiniteScreenNumber(screen.availTop)){
            top=Number(screen.availTop);
        }else if(isFiniteScreenNumber(screen.top)){
            top=Number(screen.top);
        }
    }catch(err){
    }

    if(width<=0){ width=1920; }
    if(height<=0){ height=1080; }

    return {left:left,top:top,width:width,height:height};
}

function fitCurrentWindowIntoWorkArea(fallbackW,fallbackH,preferUpperCenter){
    try{
        var area=getCurrentWorkArea();
        var pos=getCurrentWindowPosition();
        var outer=getWindowOuterSize(fallbackW,fallbackH);
        var outerW=outer.width||fallbackW||0;
        var outerH=outer.height||fallbackH||0;
        var safeMargin=8;
        var minLeft=area.left+safeMargin;
        var minTop=area.top+safeMargin;
        var maxLeft=area.left+area.width-outerW-safeMargin;
        var maxTop=area.top+area.height-outerH-safeMargin;
        var moveX=pos.left;
        var moveY=pos.top;

        /*
          ウィンドウが作業領域より大きい場合は、まず作業領域内に収まるサイズへ補正。
          通常はresizeApp側ですでに上限設定済みだが、DPI差などによる実寸差の保険。
        */
        if(outerW>area.width-(safeMargin*2) || outerH>area.height-(safeMargin*2)){
            var correctedW=Math.min(outerW,Math.max(320,area.width-(safeMargin*2)));
            var correctedH=Math.min(outerH,Math.max(240,area.height-(safeMargin*2)));
            window.resizeTo(correctedW,correctedH);
            outerW=correctedW;
            outerH=correctedH;
            maxLeft=area.left+area.width-outerW-safeMargin;
            maxTop=area.top+area.height-outerH-safeMargin;
        }

        if(maxLeft<minLeft){
            minLeft=area.left;
            maxLeft=area.left;
        }
        if(maxTop<minTop){
            minTop=area.top;
            maxTop=area.top;
        }

        if(moveX<minLeft){ moveX=minLeft; }
        if(moveX>maxLeft){ moveX=maxLeft; }
        if(moveY<minTop){ moveY=minTop; }
        if(moveY>maxTop){ moveY=maxTop; }

        /*
          v28.27:
          2件目/3件目の展開などでフォームが縦に伸びたときは、
          単に下端へ収めるだけではなく「中央より少し上」を目安に持ち上げる。
          すでにそれより上に置かれている場合は、ユーザーの位置を尊重して動かさない。
        */
        if(preferUpperCenter){
            var freeHeight=Math.max(0,area.height-outerH-(safeMargin*2));
            var preferredTop=area.top+safeMargin+Math.round(freeHeight*0.38);

            if(preferredTop<minTop){ preferredTop=minTop; }
            if(preferredTop>maxTop){ preferredTop=maxTop; }
            if(moveY>preferredTop){ moveY=preferredTop; }
        }

        if(Math.abs(moveX-pos.left)>1 || Math.abs(moveY-pos.top)>1){
            window.moveTo(Math.round(moveX),Math.round(moveY));
        }
    }catch(err){
    }
}

function centerCurrentWindow(outerW,outerH){
    try{
        var area=getCurrentWorkArea();
        var moveX=area.left+Math.max(0,Math.floor((area.width-outerW)/2));
        var moveY=area.top+Math.max(0,Math.floor((area.height-outerH)/2));
        window.moveTo(moveX,moveY);
        hasPositionedWindow=true;
        fitCurrentWindowIntoWorkArea(outerW,outerH);
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

            // 現在いるモニターの作業領域（タスクバー等を除く）を基準に上限を決める。
            var workArea=getCurrentWorkArea();
            var maxOuterW=workArea.width-40;
            var maxOuterH=workArea.height-34;

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

            // v28.26:
            // 初回は「現在いるモニター」の中央へ配置。
            // 2件目/3件目の展開・取消・送信後リセット時は現在位置を尊重しつつ、
            // 下端/右端などが作業領域からはみ出す分だけ自動で戻す。
            if(centerOnFirst && !hasPositionedWindow){
                centerCurrentWindow(wantedW,wantedH);
            }else{
                fitCurrentWindowIntoWorkArea(wantedW,wantedH,true);
            }

            // resizeTo直後はHTAフレームの実寸反映が1テンポ遅れる場合があるため、
            // 非表示中にもう一度実寸で補正してから表示する。
            window.setTimeout(function(){
                fitCurrentWindowIntoWorkArea(wantedW,wantedH,!centerOnFirst || hasPositionedWindow);
            },10);

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
    var attendance=getAttendance(no);
    var noProxy=(attendance==="公休" || attendance==="時短（時間外）") && checked("noProxy"+no);
    var useProxy=(attendance==="公休" || attendance==="時短（時間外）") && !noProxy;
    return {
        requestNo:no,
        organization:val("org"+no),
        caName:val("ca"+no),
        proxyOrganization:"",
        proxyCA1:useProxy ? val("proxyCA"+no+"_1") : "",
        proxyCA2:useProxy ? val("proxyCA"+no+"_2") : "",
        proxyCA3:useProxy ? val("proxyCA"+no+"_3") : "",
        attendance:attendance,
        noProxy:noProxy,
        mailMemo:val("mailMemo"+no),
        processedAt:val("processedAt"+no),
        dueDate:val("dueDate"+no),
        shortTime:attendance==="時短（時間外）",
        urgent:attendance==="出社" && checked("urgent"+no),
        type:val("type"+no)
    };
}

function hasAnyInput(req){
    return !!(req.organization||req.caName||req.proxyOrganization||req.proxyCA1||
              req.proxyCA2||req.proxyCA3||req.mailMemo||req.processedAt||
              req.dueDate||req.attendance||req.noProxy||req.shortTime||req.urgent||req.type);
}

function validateForm(){
    var i;

    syncAllSameCA();
    clearErrors();

    // 表示中の依頼を、画面左→右の順番でチェック
    // 組織 → CA名 → 出社状況 → 代理CA名 → メールメモ → 処理日時 → タイプ
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

    var attendance=getAttendance(i);
    if(!attendance){
        markError("attendanceGroup"+i);
        showValidationModal(i,"依頼"+i+"：出社状況を選択してください","attendanceWork"+i);
        return false;
    }

    if(attendance!=="出社" && !checked("noProxy"+i) &&
       !trimValue("proxyCA"+i+"_1") && !trimValue("proxyCA"+i+"_2") && !trimValue("proxyCA"+i+"_3")){
        markError("proxyCA"+i+"_1");
        showValidationModal(i,"依頼"+i+"：代理CA名を入力してください（記載しない場合は「代理CA記載なし」を選択）","proxyCA"+i+"_1");
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
    clearRequestFields(no);
    syncAllSameCA();
    clearErrors();
    focusField("org"+no);
}

function getAttendance(no){
    if(checked("attendanceWork"+no)){ return "出社"; }
    if(checked("attendanceOff"+no)){ return "公休"; }
    if(checked("attendanceShort"+no)){ return "時短（時間外）"; }
    return "";
}

function setAttendance(no,value){
    setChecked("attendanceWork"+no,value==="出社");
    setChecked("attendanceOff"+no,value==="公休");
    setChecked("attendanceShort"+no,value==="時短（時間外）");
}

function updateAttendanceUI(no){
    var attendance=getAttendance(no);
    var working=attendance==="出社";
    var needsProxy=attendance==="公休" || attendance==="時短（時間外）";
    var locked=no>1 && checked("sameCA"+no);
    var inheritedNoProxy=false;
    var i,el;

    /*
      v28.28:
      同一CA中の「代理CA記載なし」はCA情報の一部として前依頼へ追従させる。
      ・前依頼ON  -> 自動ON＋表示＋編集不可
      ・前依頼OFF -> 自動OFF＋項目自体を非表示
      「至急」は従来どおり依頼ごとに独立。
    */
    if(locked && needsProxy){
        inheritedNoProxy=checked("noProxy"+getSameCASourceNo(no));
        setChecked("noProxy"+no,inheritedNoProxy);
    }

    // 非表示になる項目の選択・代理CA名は送信値に残さない。
    if(!working){ setChecked("urgent"+no,false); }
    if(!needsProxy){ setChecked("noProxy"+no,false); }
    var noProxy=needsProxy && checked("noProxy"+no);
    for(i=1;i<=3;i++){
        el=$("proxyCA"+no+"_"+i);
        if(el){
            if(!needsProxy || noProxy){ el.value=""; }
            el.disabled=locked || !needsProxy || noProxy;
        }
    }

    // オプション見出しは常時表示。状態に応じて選択項目だけを切り替える。
    el=$("urgentOption"+no); if(el){ el.style.display=working ? "block" : "none"; }
    el=$("noProxyOption"+no);
    if(el){
        if(locked && needsProxy){
            el.style.display=inheritedNoProxy ? "block" : "none";
        }else{
            el.style.display=needsProxy ? "block" : "none";
        }
    }

    var controls=["attendanceWork","attendanceOff","attendanceShort"];
    for(i=0;i<controls.length;i++){
        el=$(controls[i]+no);
        if(el){ el.disabled=locked; }
    }
    el=$("urgent"+no); if(el){ el.disabled=!working; }
    el=$("noProxy"+no); if(el){ el.disabled=locked || !needsProxy; }
}

function attendanceChanged(no){
    updateAttendanceUI(no);
    syncAllSameCA();
    clearErrors();
}

function getSameCASourceNo(no){
    return no===2 ? 1 : 2;
}

function getSameCAFieldPairs(no){
    var src=getSameCASourceNo(no);
    return [
        ["org"+src,"org"+no],
        ["ca"+src,"ca"+no],
        ["proxyCA"+src+"_1","proxyCA"+no+"_1"],
        ["proxyCA"+src+"_2","proxyCA"+no+"_2"],
        ["proxyCA"+src+"_3","proxyCA"+no+"_3"]
    ];
}

function copySameCAValues(no){
    var pairs=getSameCAFieldPairs(no),i,a,b;
    var src=getSameCASourceNo(no);
    for(i=0;i<pairs.length;i++){
        a=$(pairs[i][0]); b=$(pairs[i][1]);
        if(a&&b){ b.value=a.value; }
    }

    // v28.28: 同一CAでは出社状況まで引き継ぐ。
    // 「代理CA記載なし」はupdateAttendanceUI内で前依頼へ追従し、
    // 「至急」だけは依頼ごとに独立させる。
    setAttendance(no,getAttendance(src));
    updateAttendanceUI(no);
}

function setSameCAFieldsLocked(no,locked){
    var pairs=getSameCAFieldPairs(no),i,b;
    for(i=0;i<pairs.length;i++){
        b=$(pairs[i][1]);
        if(b){ b.disabled=!!locked; }
    }
    // 連動解除後も、出社状況に応じた代理CA欄の非活性は維持。
    updateAttendanceUI(no);
}

function updateInputNoteVisibility(no,on){
    var note=$("inputNote"+no);
    if(note){ note.style.display=on ? "none" : "block"; }
}

function updateSameCALabel(no,on){
    var src=getSameCASourceNo(no),label=$("sameCALabel"+no);
    if(label){
        label.innerText=on ? "依頼"+src+"のCAを反映中" : "依頼"+src+"と同一CA";
    }
    updateInputNoteVisibility(no,on);
}

function toggleSameCA(no){
    var box=$("sameCA"+no);
    if(!box){return;}
    if(box.checked){
        copySameCAValues(no);
    }else{
        clearSameCAFields(no);
    }
    setSameCAFieldsLocked(no,box.checked);
    updateSameCALabel(no,box.checked);
    syncAllSameCA();
}

function clearSameCAFields(no){
    var pairs=getSameCAFieldPairs(no),i;
    for(i=0;i<pairs.length;i++){ setValue(pairs[i][1],""); }
    setAttendance(no,"");
    setChecked("urgent"+no,false);
    setChecked("noProxy"+no,false);
    updateAttendanceUI(no);
    clearErrors();
    focusField("org"+no);
}

function syncSameCATarget(no){
    if(no<=visibleRequestCount && checked("sameCA"+no)){
        copySameCAValues(no);
    }
}

function syncAllSameCA(){
    // 必ず依頼1→2→3の順で連鎖を更新する。
    syncSameCATarget(2);
    syncSameCATarget(3);
}

function bindSameCASync(){
    var no,i;
    for(no=1;no<=2;no++){
        bindSameCAEvent("org"+no);
        bindSameCAEvent("ca"+no);
        for(i=1;i<=3;i++){ bindSameCAEvent("proxyCA"+no+"_"+i); }
    }
}

function bindSameCAEvent(id){
    var el=$(id);
    if(!el){return;}
    var fn=function(){ syncAllSameCA(); };
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
        "org1","ca1","attendanceGroup1","proxyCA1_1","mailMemo1","processedAt1","type1",
        "org2","ca2","attendanceGroup2","proxyCA2_1","mailMemo2","processedAt2","type2",
        "org3","ca3","attendanceGroup3","proxyCA3_1","mailMemo3","processedAt3","type3"
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


var sendConfirmationOpen=false;

function buildSendConfirmDetails(){
    var lines=[];
    var i;

    // 同一CA連動中でも、現在の各依頼カードの値を依頼ごとに表示する。
    syncAllSameCA();

    for(i=1;i<=visibleRequestCount;i++){
        var attendance=getAttendance(i);
        var line="依頼"+i+"　"+trimValue("ca"+i)+"CA："+attendance;

        // v28.29: 公休／時短（時間外）の場合は、代理CA情報も送信前確認へ表示する。
        if(attendance==="公休" || attendance==="時短（時間外）"){
            if(checked("noProxy"+i)){
                line+="　代理CA記載なし";
            }else{
                var proxyNames=[];
                var j,proxyName;
                for(j=1;j<=3;j++){
                    proxyName=trimValue("proxyCA"+i+"_"+j);
                    if(proxyName){ proxyNames.push(proxyName); }
                }
                line+="　代理CA名："+proxyNames.join("、");
            }
        }

        line+="　メールメモ："+trimValue("mailMemo"+i);
        lines.push(line);
    }

    return lines.join("\r\n");
}

function updateSendConfirmDetails(){
    var details=$("sendConfirmDetails");
    if(details){
        details.innerText=buildSendConfirmDetails();
    }
}

function centerSendConfirmModal(){
    var modal=$("sendConfirmModal");
    if(!modal){ return; }

    // 表示内容が1～3件で高さが変わっても、実寸を基準に画面中央へ配置する。
    modal.style.left="50%";
    modal.style.top="50%";
    modal.style.marginLeft=String(-Math.round(modal.offsetWidth/2))+"px";
    modal.style.marginTop=String(-Math.round(modal.offsetHeight/2))+"px";
}

function sendRequest(){
    if(sendConfirmationOpen || !validateForm()){ return; }

    // 送信前確認に、CA名・出社状況・必要時の代理CA情報・メールメモを表示。
    updateSendConfirmDetails();

    sendConfirmationOpen=true;
    $("sendConfirmOverlay").className="modal-overlay";
    centerSendConfirmModal();
    focusField("sendConfirmNo");
}

function closeSendConfirm(){
    sendConfirmationOpen=false;
    $("sendConfirmOverlay").className="modal-overlay hidden";
    focusField("sendButton");
}

function confirmSendRequest(){
    if(!sendConfirmationOpen){ return; }
    closeSendConfirm();
    submitConfirmedRequest();
}

function handleSendConfirmKey(event){
    event=event || window.event;
    var code=event.keyCode || event.which;
    if(code===27){
        closeSendConfirm();
    }else if(code===9){
        // 確認中にTabで背面の入力欄や送信ボタンへ移動しない。
        focusField(document.activeElement===$("sendConfirmNo") ? "sendConfirmYes" : "sendConfirmNo");
    }else{
        return;
    }
    if(event.preventDefault){ event.preventDefault(); }
    event.returnValue=false;
}

function submitConfirmedRequest(){
    if(!validateForm()){ return; }

    var requests=[],i,req;
    for(i=1;i<=visibleRequestCount;i++){
        req=collectRequest(i);
        if(i===1||hasAnyInput(req)){ requests.push(req); }
    }

    var requester=$("requesterName").innerText;
    // Pendingファイルを1送信単位でまとめるための内部ID。CSVには出力しない。
    var packageId="PKG-"+createRequestId(requester,0);
    var sentAt=formatDateTime(new Date());
    var baseName=val("requestBase");
    var lines=[];
    var pendingPath="";

    // v28.23: RequestIDは依頼ごとに個別発行する。
    for(i=0;i<requests.length;i++){
        var rowRequestId=createRequestId(requester,requests[i].requestNo);
        lines.push(makeCsvLine(rowRequestId,sentAt,baseName,requester,requests[i]));
    }

    try{
        var csvFolder=getProductionCsvFolder(true);

        // 1. まずPendingへ保存。
        pendingPath=savePendingPackage(
            baseName,
            requester,
            packageId,
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
            syncAllSameCA();
            // 元の依頼3が依頼1まで連動していた場合だけ、新しい依頼2も連動を維持。
            var keepSameCA=checked("sameCA3") && checked("sameCA2");
            setChecked("sameCA2",false);
            copyRequestFields(3,2);
            setChecked("sameCA2",keepSameCA);
            setSameCAFieldsLocked(2,keepSameCA);
            updateSameCALabel(2,keepSameCA);
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

    setAttendance(toNo,getAttendance(fromNo));
    setChecked("urgent"+toNo,checked("urgent"+fromNo));
    setChecked("noProxy"+toNo,checked("noProxy"+fromNo));
    updateAttendanceUI(toNo);

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

    setAttendance(i,"");
    setChecked("urgent"+i,false);
    setChecked("noProxy"+i,false);
    setChecked("sameCA"+i,false);
    if(i>1){
        setSameCAFieldsLocked(i,false);
        updateSameCALabel(i,false);
    }else{
        updateAttendanceUI(i);
    }

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

function savePendingPackage(baseName,requester,packageId,sentAt,requests,lines,csvFolder){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var folderPath=getPendingFolderPath();

    ensureFolder(fso,folderPath);

    var fileName=sanitizeFileName(packageId)+".csv";
    var filePath=folderPath+"\\"+fileName;

    // 万一同名があれば上書きせず別名にする
    if(fso.FileExists(filePath)){
        fileName=sanitizeFileName(packageId)+"_"+String(new Date().getTime())+".csv";
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
    return "送信日時,拠点,依頼者,依頼No.,組織,CA名,出社状況,代理CA1,代理CA2,代理CA3,オプション,メールメモ,処理日時,期日または面接日程,タイプ,RequestID";
}

function getCsvOptionValue(req){
    if(req.urgent){ return "至急"; }
    if(req.noProxy){ return "代理CA記載なし"; }
    return "";
}

function makeCsvLine(requestId,sentAt,baseName,requester,req){
    return csvJoin([
        sentAt,baseName,requester,String(req.requestNo),
        req.organization,req.caName,req.attendance,
        req.proxyCA1,req.proxyCA2,req.proxyCA3,
        getCsvOptionValue(req),
        req.mailMemo,req.processedAt,req.dueDate,
        req.type,requestId
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

function createRequestId(requester,requestNo){
    var now=new Date();
    var userPart=String(requester||"USER")
        .replace(/[\\\/:*?"<>|\s]/g,"")
        .substring(0,12);

    var rand1=Math.floor(Math.random()*0x1000000).toString(16).toUpperCase();
    while(rand1.length<6){ rand1="0"+rand1; }

    var milli=String(now.getMilliseconds());
    while(milli.length<3){ milli="0"+milli; }

    var requestPart=requestNo>0 ? "-R"+pad2(requestNo) : "";

    return formatDate(now)+"-"+
           pad2(now.getHours())+pad2(now.getMinutes())+pad2(now.getSeconds())+milli+"-"+
           userPart+requestPart+"-"+rand1;
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
