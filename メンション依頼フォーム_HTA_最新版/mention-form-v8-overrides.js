/* mention-form v8 repository compatibility overrides
   Approved v8 behavior layered on the existing production JS.
*/
(function(){
    var baseInitApp = initApp;
    var baseClearRequestFields = clearRequestFields;
    var baseCancelRequest = cancelRequest;

    getCurrentFolderPath = function(){
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

        if(host && host.toLowerCase()!=="localhost"){
            while(path.charAt(0)==="\\"){
                path=path.substring(1);
            }
            path="\\\\"+host+"\\"+path;
        }else{
            if(/^\\[A-Za-z]:\\/.test(path)){
                path=path.substring(1);
            }
        }

        var fso=new ActiveXObject("Scripting.FileSystemObject");
        return fso.GetParentFolderName(path);
    };

    validateRequestSection = function(i){
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
    };

    window.clearRequest = function(no){
        var ids=[
            "org"+no,"ca"+no,"proxyOrg"+no,
            "proxyCA"+no+"_1","proxyCA"+no+"_2","proxyCA"+no+"_3",
            "mailMemo"+no,"processedAt"+no,"dueDate"+no
        ];
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
    };

    window.getSameCASourceNo = function(no){
        return no===2 ? 1 : 2;
    };

    window.getSameCAFieldPairs = function(no){
        var src=getSameCASourceNo(no);
        return [
            ["org"+src,"org"+no],
            ["ca"+src,"ca"+no],
            ["proxyOrg"+src,"proxyOrg"+no],
            ["proxyCA"+src+"_1","proxyCA"+no+"_1"],
            ["proxyCA"+src+"_2","proxyCA"+no+"_2"],
            ["proxyCA"+src+"_3","proxyCA"+no+"_3"]
        ];
    };

    window.copySameCAValues = function(no){
        var pairs=getSameCAFieldPairs(no),i,a,b;
        for(i=0;i<pairs.length;i++){
            a=$(pairs[i][0]);
            b=$(pairs[i][1]);
            if(a&&b){ b.value=a.value; }
        }
    };

    window.setSameCAFieldsLocked = function(no,locked){
        var pairs=getSameCAFieldPairs(no),i,b;
        for(i=0;i<pairs.length;i++){
            b=$(pairs[i][1]);
            if(b){ b.disabled=!!locked; }
        }
    };

    window.updateSameCALabel = function(no,on){
        var src=getSameCASourceNo(no);
        var label=$("sameCALabel"+no);
        if(label){
            label.innerText=on ? "依頼"+src+"のCAを反映中" : "依頼"+src+"と同一CA";
        }
    };

    window.clearSameCAFields = function(no){
        var ids=[
            "org"+no,"ca"+no,"proxyOrg"+no,
            "proxyCA"+no+"_1","proxyCA"+no+"_2","proxyCA"+no+"_3"
        ];
        var i,el;

        for(i=0;i<ids.length;i++){
            el=$(ids[i]);
            if(el){ el.value=""; }
        }

        clearErrors();
        try{$("org"+no).focus();}catch(err){}
    };

    window.toggleSameCA = function(no){
        var box=$("sameCA"+no);
        if(!box){return;}

        if(box.checked){
            copySameCAValues(no);
            setSameCAFieldsLocked(no,true);
            updateSameCALabel(no,true);
        }else{
            setSameCAFieldsLocked(no,false);
            clearSameCAFields(no);
            updateSameCALabel(no,false);
        }
    };

    window.syncSameCATarget = function(no){
        var box=$("sameCA"+no);
        if(box&&box.checked){ copySameCAValues(no); }
    };

    window.bindSameCAEvent = function(id,targetNo){
        var el=$(id);
        if(!el){return;}

        var fn=function(){
            syncSameCATarget(targetNo);
            if(targetNo===2){
                syncSameCATarget(3);
            }
        };

        if(el.addEventListener){
            el.addEventListener("input",fn,false);
            el.addEventListener("change",fn,false);
        }else if(el.attachEvent){
            el.attachEvent("onkeyup",fn);
            el.attachEvent("onchange",fn);
        }
    };

    window.bindSameCASync = function(){
        var src1=["org1","ca1","proxyOrg1","proxyCA1_1","proxyCA1_2","proxyCA1_3"];
        var src2=["org2","ca2","proxyOrg2","proxyCA2_1","proxyCA2_2","proxyCA2_3"];
        var i;

        for(i=0;i<src1.length;i++){ bindSameCAEvent(src1[i],2); }
        for(i=0;i<src2.length;i++){ bindSameCAEvent(src2[i],3); }
    };

    clearRequestFields = function(i){
        if(i===2 || i===3){
            var box=$("sameCA"+i);
            if(box){ box.checked=false; }
            setSameCAFieldsLocked(i,false);
            updateSameCALabel(i,false);
        }
        baseClearRequestFields(i);
    };

    cancelRequest = function(no){
        if(no===2 || no===3){
            var box=$("sameCA"+no);
            if(box){ box.checked=false; }
            setSameCAFieldsLocked(no,false);
            updateSameCALabel(no,false);
        }
        baseCancelRequest(no);
    };

    showAcceptedMessage = function(){
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
    };

    initApp = function(){
        baseInitApp();
        bindSameCASync();
    };
})();
