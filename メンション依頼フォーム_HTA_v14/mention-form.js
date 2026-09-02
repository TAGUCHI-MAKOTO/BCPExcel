var CSV_FOLDER = "%TEMP%\\MentionRequestCSV";
var RETRY_INTERVAL_MS = 80;
var RETRY_MAX = 40;

var TYPE_OPTIONS = [
    "",
    "通常",
    "確認",
    "代理対応",
    "その他",
    "【長文サンプル】候補者・企業双方への確認が必要なため、処理前に担当者へエスカレーションする"
];

var visibleRequestCount = 1;

function $(id){ return document.getElementById(id); }

function initApp(){
    populateTypes();
    setRequesterName();
    resizeApp();
    setStatus("試作版：CSVはTEMPフォルダへ保存");
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
    if(no===2){
        $("request2").className=$("request2").className.replace(" hidden","");
        $("btnAdd2").style.display="none";
        visibleRequestCount=2;
    }else if(no===3){
        $("request3").className=$("request3").className.replace(" hidden","");
        $("btnAdd3").style.display="none";
        visibleRequestCount=3;
    }
    resizeApp();
}

function resizeApp(){
    window.setTimeout(function(){
        try{
            var app=$("app");
            if(!app){ return; }
            document.body.style.zoom="1";
            var chromeW=70;
            var chromeH=82;
            var naturalW=app.offsetWidth;
            var naturalH=app.offsetHeight;
            var maxW=screen.availWidth-18;
            var maxH=screen.availHeight-18;
            var scaleW=(maxW-chromeW)/naturalW;
            var scaleH=(maxH-chromeH)/naturalH;
            var scale=Math.min(1,scaleW,scaleH);
            if(scale<0.78){ scale=0.78; }
            document.body.style.zoom=String(scale);
            var wantedW=Math.ceil(naturalW*scale+chromeW);
            var wantedH=Math.ceil(naturalH*scale+chromeH);
            if(wantedW>maxW){ wantedW=maxW; }
            if(wantedH>maxH){ wantedH=maxH; }
            window.resizeTo(wantedW,wantedH);
            var moveX=Math.max(0,Math.floor((screen.availWidth-wantedW)/2));
            var moveY=Math.max(0,Math.floor((screen.availHeight-wantedH)/2));
            window.moveTo(moveX,moveY);
            window.scrollTo(0,0);
        }catch(err){}
    },60);
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
    if(!val("requestBase")){
        markError("requestBase");
        alert("拠点を選択してください");
        focusField("requestBase");
        return false;
    }
    for(i=1;i<=visibleRequestCount;i++){
        if(!val("mailMemo"+i)){
            markError("mailMemo"+i);
            alert("依頼"+i+"：メールメモを入力してください");
            focusField("mailMemo"+i);
            return false;
        }
        if(!val("processedAt"+i)){
            markError("processedAt"+i);
            alert("依頼"+i+"：処理日時を入力してください");
            focusField("processedAt"+i);
            return false;
        }
        if(!val("type"+i)){
            markError("type"+i);
            alert("依頼"+i+"：タイプを選択してください");
            focusField("type"+i);
            return false;
        }
        var normalOrg = trimValue("org"+i);
        var normalCA = trimValue("ca"+i);
        var proxyOrg = trimValue("proxyOrg"+i);
        var proxyCA1 = trimValue("proxyCA"+i+"_1");
        var normalAny = !!(normalOrg || normalCA);
        var proxyAny = !!(proxyOrg || proxyCA1);
        var normalComplete = !!(normalOrg && normalCA);
        var proxyComplete = !!(proxyOrg && proxyCA1);
        if(!normalAny && !proxyAny){
            markError("org"+i);
            markError("ca"+i);
            markError("proxyOrg"+i);
            markError("proxyCA"+i+"_1");
            alert("依頼"+i+"：\n" + "「組織＋CA名」または「代理CA組織＋代理CA名」を入力してください");
            focusField("org"+i);
            return false;
        }
        if(normalAny && !normalComplete && !proxyComplete){
            if(!normalOrg){ markError("org"+i); }
            if(!normalCA){ markError("ca"+i); }
            alert("依頼"+i+"：「組織」と「CA名」はセットで入力してください");
            if(!normalOrg){ focusField("org"+i); }else{ focusField("ca"+i); }
            return false;
        }
        if(proxyAny && !proxyComplete && !normalComplete){
            if(!proxyOrg){ markError("proxyOrg"+i); }
            if(!proxyCA1){ markError("proxyCA"+i+"_1"); }
            alert("依頼"+i+"：「代理CA組織」と「代理CA名」はセットで入力してください");
            if(!proxyOrg){ focusField("proxyOrg"+i); }else{ focusField("proxyCA"+i+"_1"); }
            return false;
        }
    }
    return true;
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
    var lines=[];
    for(i=0;i<requests.length;i++){
        lines.push(makeCsvLine(requestId,sentAt,val("requestBase"),requester,requests[i]));
    }
    setStatus("送信中...");
    $("sendButton").disabled=true;
    writeCsvWithLock(
        val("requestBase"),
        lines,
        0,
        function(){
            setStatus("送信しました");
            $("sendButton").disabled=false;
            alert("メンション依頼が送信されました");
            resetAfterSend();
        },
        function(message){
            setStatus("送信失敗");
            $("sendButton").disabled=false;
            alert("CSVへの書き込みに失敗しました。\n\n"+message);
        }
    );
}

function resetAfterSend(){
    var i;
    for(i=1;i<=3;i++){
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
    if($("request2")){
        if($("request2").className.indexOf("hidden")<0){
            $("request2").className+=" hidden";
        }
    }
    if($("request3")){
        if($("request3").className.indexOf("hidden")<0){
            $("request3").className+=" hidden";
        }
    }
    if($("btnAdd2")){ $("btnAdd2").style.display=""; }
    if($("btnAdd3")){ $("btnAdd3").style.display=""; }
    visibleRequestCount=1;
    clearErrors();
    setStatus("");
    try{ $("org1").focus(); }catch(err){}
    resizeApp();
}

function setValue(id,value){
    var el=$(id);
    if(el){ el.value=value; }
}

function setChecked(id,value){
    var el=$(id);
    if(el){ el.checked=value; }
}

function writeCsvWithLock(baseName,lines,retryCount,onSuccess,onFailure){
    var fso,shell,folderPath,csvPath,lockPath,lockFile,csvFile;
    try{
        fso=new ActiveXObject("Scripting.FileSystemObject");
        shell=new ActiveXObject("WScript.Shell");
        folderPath=shell.ExpandEnvironmentStrings(CSV_FOLDER);
        ensureFolder(fso,folderPath);
        csvPath=folderPath+"\\"+sanitizeFileName(baseName)+"_"+formatDate(new Date())+".csv";
        lockPath=csvPath+".lock";
        removeStaleLock(fso,lockPath,15);
        lockFile=fso.CreateTextFile(lockPath,false,false);
        lockFile.WriteLine("locked");
        lockFile.Close();
        try{
            var isNew=!fso.FileExists(csvPath)||fso.GetFile(csvPath).Size===0;
            csvFile=fso.OpenTextFile(csvPath,8,true,0);
            if(isNew){ csvFile.WriteLine(makeHeaderLine()); }
            var i;
            for(i=0;i<lines.length;i++){ csvFile.WriteLine(lines[i]); }
            csvFile.Close();
            csvFile=null;
            if(fso.FileExists(lockPath)){ fso.DeleteFile(lockPath,true); }
            onSuccess();
            return;
        }catch(writeErr){
            try{ if(csvFile){ csvFile.Close(); } }catch(closeErr){}
            try{ if(fso.FileExists(lockPath)){ fso.DeleteFile(lockPath,true); } }catch(unlockErr){}
            throw writeErr;
        }
    }catch(err){
        if(retryCount<RETRY_MAX){
            window.setTimeout(function(){
                writeCsvWithLock(baseName,lines,retryCount+1,onSuccess,onFailure);
            },RETRY_INTERVAL_MS);
        }else{
            onFailure(String(err.message||err.description||err));
        }
    }
}

function ensureFolder(fso,folderPath){
    if(fso.FolderExists(folderPath)){ return; }
    var parent=fso.GetParentFolderName(folderPath);
    if(parent&&!fso.FolderExists(parent)){ ensureFolder(fso,parent); }
    fso.CreateFolder(folderPath);
}

function removeStaleLock(fso,lockPath,maxAgeSeconds){
    try{
        if(!fso.FileExists(lockPath)){ return; }
        var file=fso.GetFile(lockPath);
        var modified=new Date(file.DateLastModified);
        var age=(new Date().getTime()-modified.getTime())/1000;
        if(age>maxAgeSeconds){ fso.DeleteFile(lockPath,true); }
    }catch(err){}
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
    var rand=Math.floor(Math.random()*65536).toString(16).toUpperCase();
    while(rand.length<4){ rand="0"+rand; }
    return formatDate(now)+"-"+
           pad2(now.getHours())+pad2(now.getMinutes())+pad2(now.getSeconds())+"-"+
           userPart+"-"+rand;
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
    if(el){ el.innerText=message||""; }
}
