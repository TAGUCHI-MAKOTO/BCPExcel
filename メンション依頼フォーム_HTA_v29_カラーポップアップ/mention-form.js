/* v29 patch: v20 color base + v25/v28/v29 behavior */
document.write('<script type="text/javascript" src="mention-form-base.js"><\/script>');

var lastAvailWidth=0;
var lastAvailHeight=0;
var screenWatchTimer=null;

function fixRequest3ActionRow(){
    try{
        var type3=$("type3");
        if(!type3){ return; }
        var typeArea=type3.parentNode;
        var req3=$("request3");
        if(!typeArea || !req3){ return; }
        var divs=req3.getElementsByTagName("div");
        var action=null;
        var i;
        for(i=0;i<divs.length;i++){
            if((" "+divs[i].className+" ").indexOf(" action-row-cancel-only ")>=0){
                action=divs[i];
                break;
            }
        }
        if(action && action.parentNode===typeArea){
            typeArea.parentNode.insertBefore(action,typeArea.nextSibling);
        }
    }catch(err){}
}

function initApp(){
    populateTypes();
    setRequesterName();
    setStatus("試作版：CSVはTEMPフォルダへ保存");
    fixRequest3ActionRow();
    lastAvailWidth=screen.availWidth;
    lastAvailHeight=screen.availHeight;
    resizeApp(true);
    startScreenWatcher();
}

function startScreenWatcher(){
    if(screenWatchTimer){ window.clearInterval(screenWatchTimer); }
    screenWatchTimer=window.setInterval(function(){
        try{
            var w=screen.availWidth;
            var h=screen.availHeight;
            if(w!==lastAvailWidth || h!==lastAvailHeight){
                lastAvailWidth=w;
                lastAvailHeight=h;
                resizeApp(true);
            }
        }catch(err){}
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
            document.body.style.zoom="1";
            var chromeW=54;
            var chromeH=76;
            var naturalW=app.offsetWidth;
            var naturalH=app.offsetHeight;
            var outerMarginW=28;
            var outerMarginH=34;
            var maxW=screen.availWidth-outerMarginW;
            var maxH=screen.availHeight-outerMarginH;
            var fitW=(maxW-chromeW)/naturalW;
            var fitH=(maxH-chromeH)/naturalH;
            var fitScale=Math.min(1,fitW,fitH);
            var comfortScale=maxW/1960;
            if(comfortScale>1){ comfortScale=1; }
            var scale=Math.min(fitScale,comfortScale);
            if(scale<=0 || isNaN(scale)){ scale=1; }
            document.body.style.zoom=String(scale);
            var wantedW=Math.ceil(naturalW*scale+chromeW);
            var wantedH=Math.ceil(naturalH*scale+chromeH);
            if(wantedW>maxW){ wantedW=maxW; }
            if(wantedH>maxH){ wantedH=maxH; }
            window.resizeTo(wantedW,wantedH);
            var moveX=Math.max(0,Math.floor((screen.availWidth-wantedW)/2));
            var moveY=Math.max(0,Math.floor((screen.availHeight-wantedH)/2));
            window.moveTo(moveX,moveY);
            hasPositionedWindow=true;
            window.scrollTo(0,0);
        }catch(err){
        }finally{
            layoutTimer=null;
            showAppAfterLayout();
        }
    },15);
}

function validateForm(){
    var i;
    clearErrors();
    for(i=1;i<=visibleRequestCount;i++){
        if(!validateRequestSection(i)){ return false; }
    }
    if(!trimValue("requestBase")){
        markError("requestBase");
        showValidationModal(0,"拠点を選択してください","requestBase");
        return false;
    }
    return true;
}

function validateRequestSection(i){
    var normalOrg=trimValue("org"+i);
    var normalCA=trimValue("ca"+i);
    var proxyOrg=trimValue("proxyOrg"+i);
    var proxyCA1=trimValue("proxyCA"+i+"_1");
    var normalAny=!!(normalOrg||normalCA);
    var proxyAny=!!(proxyOrg||proxyCA1);
    var normalComplete=!!(normalOrg&&normalCA);
    var proxyComplete=!!(proxyOrg&&proxyCA1);

    if(!normalAny&&!proxyAny){
        markError("org"+i);
        showValidationModal(i,"依頼"+i+"：組織を入力してください。\n\n代理対応の場合は「代理CA組織＋代理CA名」を入力してください。","org"+i);
        return false;
    }

    if(normalAny&&!normalComplete&&!proxyComplete){
        if(!normalOrg){
            markError("org"+i);
            showValidationModal(i,"依頼"+i+"：組織を入力してください","org"+i);
            return false;
        }
        if(!normalCA){
            markError("ca"+i);
            showValidationModal(i,"依頼"+i+"：CA名を入力してください","ca"+i);
            return false;
        }
    }

    if(proxyAny&&!proxyComplete&&!normalComplete){
        if(!proxyOrg){
            markError("proxyOrg"+i);
            showValidationModal(i,"依頼"+i+"：代理CA組織を入力してください","proxyOrg"+i);
            return false;
        }
        if(!proxyCA1){
            markError("proxyCA"+i+"_1");
            showValidationModal(i,"依頼"+i+"：代理CA名を入力してください","proxyCA"+i+"_1");
            return false;
        }
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
