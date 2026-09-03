var CONFIG_FILE_NAME = "mention-form.config.ini";
var PENDING_FOLDER_NAME = "MentionRequest_Pending";
var BACKGROUND_WORKER_NAME = "mention-request-worker.ps1";
var WORKER_RETRY_MIN_MS = 450;
var WORKER_RETRY_MAX_MS = 1250;
var WORKER_WAIT_NOTICE_SECONDS = 60;

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
        getConfiguredCsvFolder(true);
        setStatus("");
    }catch(configErr){
        setStatus("共有CSV格納先が未設定です");
    }

    lastAvailWidth=screen.availWidth;
    lastAvailHeight=screen.availHeight;

    // 初回は画面を見せる前にサイズを確定
    resizeApp(true);

    // HTAを開いたまま解像度・表示倍率・接続モニターが変わった場合にも追従
    startScreenWatcher();

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

            // 100%に戻して、現在表示中の依頼カード数で自然サイズを測る
            document.body.style.zoom="1";

            var chromeW=54;
            var chromeH=76;

            var naturalW=app.offsetWidth;
            var naturalH=app.offsetHeight;

            // 利用可能画面の端に少しだけ余白を残す
            var outerMarginW=28;
            var outerMarginH=34;

            var maxW=screen.availWidth-outerMarginW;
            var maxH=screen.availHeight-outerMarginH;

            // 画面内へ確実に収める倍率
            var fitW=(maxW-chromeW)/naturalW;
            var fitH=(maxH-chromeH)/naturalH;
            var fitScale=Math.min(1,fitW,fitH);

            /*
              v26では「maxW / 2040」を使っていたため、
              1920×1080環境では約92%まで強制的に縮小されていた。

              v28では縮小を弱め、
              1920×1080級では約96%、
              それより広い画面では自然に100%へ近づける。
              ただし画面へ収まらない場合は fitScale を優先する。
            */
            var comfortScale=maxW/1960;
            if(comfortScale>1){ comfortScale=1; }

            var scale=Math.min(fitScale,comfortScale);

            if(scale<=0 || isNaN(scale)){
                scale=1;
            }

            document.body.style.zoom=String(scale);

            var wantedW=Math.ceil(naturalW*scale+chromeW);
            var wantedH=Math.ceil(naturalH*scale+chromeH);

            if(wantedW>maxW){ wantedW=maxW; }
            if(wantedH>maxH){ wantedH=maxH; }

            window.resizeTo(wantedW,wantedH);

            // 解像度変更時も見失わないよう中央へ再配置
            var moveX=Math.max(0,Math.floor((screen.availWidth-wantedW)/2));
            var moveY=Math.max(0,Math.floor((screen.availHeight-wantedH)/2));

            if(centerOnFirst || !hasPositionedWindow){
                window.moveTo(moveX,moveY);
                hasPositionedWindow=true;
            }else{
                window.moveTo(moveX,moveY);
            }

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
    // -----------------------------------------------------
    // 1～4：宛先情報
    // 「組織＋CA名」または「代理CA組織＋代理CA名」の
    // どちらか一式必須という従来ルールは維持する。
    // -----------------------------------------------------
    var normalOrg=trimValue("org"+i);
    var normalCA=trimValue("ca"+i);
    var proxyOrg=trimValue("proxyOrg"+i);
    var proxyCA1=trimValue("proxyCA"+i+"_1");

    var normalAny=!!(normalOrg||normalCA);
    var proxyAny=!!(proxyOrg||proxyCA1);

    var normalComplete=!!(normalOrg&&normalCA);
    var proxyComplete=!!(proxyOrg&&proxyCA1);

    // 何も入力されていない場合は、最初の項目「組織」から案内
    if(!normalAny&&!proxyAny){
        markError("org"+i);

        showValidationModal(
            i,
            "依頼"+i+"：組織を入力してください。\n\n代理対応の場合は「代理CA組織＋代理CA名」を入力してください。",
            "org"+i
        );
        return false;
    }

    // 組織側を入力し始めた場合：
    // 組織 → CA名 の順番で不足項目を案内
    if(normalAny&&!normalComplete&&!proxyComplete){
        if(!normalOrg){
            markError("org"+i);
            showValidationModal(
                i,
                "依頼"+i+"：組織を入力してください",
                "org"+i
            );
            return false;
        }

        if(!normalCA){
            markError("ca"+i);
            showValidationModal(
                i,
                "依頼"+i+"：CA名を入力してください",
                "ca"+i
            );
            return false;
        }
    }

    // 代理側を入力し始めた場合：
    // 代理CA組織 → 代理CA名 の順番で不足項目を案内
    if(proxyAny&&!proxyComplete&&!normalComplete){
        if(!proxyOrg){
            markError("proxyOrg"+i);
            showValidationModal(
                i,
                "依頼"+i+"：代理CA組織を入力してください",
                "proxyOrg"+i
            );
            return false;
        }

        if(!proxyCA1){
            markError("proxyCA"+i+"_1");
            showValidationModal(
                i,
                "依頼"+i+"：代理CA名を入力してください",
                "proxyCA"+i+"_1"
            );
            return false;
        }
    }

    // -----------------------------------------------------
    // 5：メールメモ
    // -----------------------------------------------------
    if(!trimValue("mailMemo"+i)){
        markError("mailMemo"+i);
        showValidationModal(
            i,
            "依頼"+i+"：メールメモを入力してください",
            "mailMemo"+i
        );
        return false;
    }

    // -----------------------------------------------------
    // 6：処理日時
    // -----------------------------------------------------
    if(!trimValue("processedAt"+i)){
        markError("processedAt"+i);
        showValidationModal(
            i,
            "依頼"+i+"：処理日時を入力してください",
            "processedAt"+i
        );
        return false;
    }

    // -----------------------------------------------------
    // 7：タイプ
    // -----------------------------------------------------
    if(!trimValue("type"+i)){
        markError("type"+i);
        showValidationModal(
            i,
            "依頼"+i+"：タイプを選択してください",
            "type"+i
        );
        return false;
    }

    return true;
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

    for(i=0;i<requests.length;i++){
        lines.push(makeCsvLine(requestId,sentAt,baseName,requester,requests[i]));
    }

    $("sendButton").disabled=true;

    try{
        // 本番CSV格納先を設定ファイルから取得
        var csvFolder=getConfiguredCsvFolder(true);

        // ① まずPendingへ確実に保存
        //    送信先もPendingへ固定しておくため、後から設定が変わっても誤送信しない。
        var pendingPath=savePendingPackage(
            baseName,
            requester,
            requestId,
            sentAt,
            requests,
            lines,
            csvFolder
        );

        // ② 別プロセスのworkerへ引き渡す
        launchBackgroundWorker(pendingPath,csvFolder);

        // ③ ここからHTAを閉じてもworkerは継続
        resetAfterSend();
        showAcceptedMessage();

    }catch(err){
        alert(
            "送信受付に失敗しました。\n\n"+
            String(err.message||err.description||err)
        );
    }

    $("sendButton").disabled=false;
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
            copyRequestFields(3,2);
            clearRequestFields(3);

            addHiddenClass("request3");

            if($("btnAdd3")){
                $("btnAdd3").style.display="";
            }

            visibleRequestCount=2;

        }else{
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
    var path="";

    try{
        path=decodeURIComponent(String(window.location.pathname||""));
    }catch(err){
        path=String(window.location.pathname||"");
    }

    path=path.replace(/\//g,"\\");

    if(/^\\[A-Za-z]:\\/.test(path)){
        path=path.substring(1);
    }

    var fso=new ActiveXObject("Scripting.FileSystemObject");
    return fso.GetParentFolderName(path);
}

function readProductionConfig(){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var configPath=getCurrentFolderPath()+"\\"+CONFIG_FILE_NAME;
    var result={};

    if(!fso.FileExists(configPath)){
        throw new Error(CONFIG_FILE_NAME+" が見つかりません。");
    }

    // configはUTF-16で保存（日本語UNCパス対応）
    var file=fso.OpenTextFile(configPath,1,false,-1);
    var text=file.ReadAll();
    file.Close();

    var rows=String(text)
        .replace(/\r\n/g,"\n")
        .replace(/\r/g,"\n")
        .split("\n");

    var i,row,pos,key,value;

    for(i=0;i<rows.length;i++){
        row=String(rows[i]||"").replace(/^\s+|\s+$/g,"");

        if(!row || row.charAt(0)===";" || row.charAt(0)==="#"){
            continue;
        }

        pos=row.indexOf("=");

        if(pos<=0){
            continue;
        }

        key=row.substring(0,pos).replace(/^\s+|\s+$/g,"").toUpperCase();
        value=row.substring(pos+1).replace(/^\s+|\s+$/g,"");

        result[key]=value;
    }

    return result;
}

function getConfiguredCsvFolder(required){
    var config=readProductionConfig();
    var raw=String(config.CSV_FOLDER||"").replace(/^\s+|\s+$/g,"");

    if(
        !raw ||
        raw==="【ここに共有フォルダのUNCパスを設定】" ||
        raw.indexOf("ここに共有フォルダのUNCパスを設定")>=0
    ){
        if(required){
            throw new Error(
                "共有CSV格納先が未設定です。\n\n"+
                CONFIG_FILE_NAME+" の CSV_FOLDER を設定してください。"
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

function recoverPendingPackagesOnStartup(){
    var fso,shell,pendingFolder,folder,files,enumerator,file;

    try{
        fso=new ActiveXObject("Scripting.FileSystemObject");
        shell=new ActiveXObject("WScript.Shell");
        pendingFolder=shell.SpecialFolders("Desktop")+"\\MentionRequest_Pending";

        if(!fso.FolderExists(pendingFolder)){
            return;
        }

        folder=fso.GetFolder(pendingFolder);
        files=folder.Files;
        enumerator=new Enumerator(files);

        for(;!enumerator.atEnd();enumerator.moveNext()){
            file=enumerator.item();

            if(String(fso.GetExtensionName(file.Name)).toLowerCase()==="pending"){
                try{
                    var recoveryFolder="";
                    try{
                        recoveryFolder=getConfiguredCsvFolder(false);
                    }catch(configErr){
                        recoveryFolder="";
                    }

                    launchBackgroundWorker(file.Path,recoveryFolder);
                }catch(err){
                }
            }
        }
    }catch(err){
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
        "-WaitNoticeSeconds "+String(WORKER_WAIT_NOTICE_SECONDS);

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

    if(fso.FileExists(filePath)){
        fileName=sanitizeFileName(requestId)+"_"+String(new Date().getTime())+".pending";
        filePath=folderPath+"\\"+fileName;
    }

    var file=fso.CreateTextFile(filePath,false,true);

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
