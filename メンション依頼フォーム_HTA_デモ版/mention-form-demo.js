/* 仕様書用デモ：本番v13の画面を使い、実ファイル操作だけ無効化 */
var DEMO_PENDING_COUNT=1;
var DEMO_PENDING_STATE="idle";

function initApp(){
    populateTypes();
    setRequesterName();
    setStatus("");
    lastAvailWidth=screen.availWidth;
    lastAvailHeight=screen.availHeight;
    resizeApp(true);
    startScreenWatcher();
    bindSameCASync();
    updatePendingRetryUI();
}
function getPendingFilePaths(){return [];}
function getPendingCount(){return DEMO_PENDING_COUNT;}
function updatePendingRetryUI(){
    var area=$("pendingRetryArea"),text=$("pendingRetryText"),button=$("retryPendingButton");
    if(!area||!text||!button){return;}
    setSendButtonVisible(true);
    if(DEMO_PENDING_COUNT<=0){area.className="pending-retry-area hidden";text.innerText="";button.style.display="none";return;}
    area.className="pending-retry-area";
    if(DEMO_PENDING_STATE==="manual"){
        text.innerText="再送中です… 完了後、Windows通知でお知らせします。";
        button.style.display="none";
    }else{
        text.innerText="⚠ 送信できませんでした｜未送信 "+DEMO_PENDING_COUNT+"件";
        button.style.display="inline-block";
        button.disabled=false;
    }
}
function startPendingWatcher(){updatePendingRetryUI();}
function recoverPendingPackagesOnStartup(){updatePendingRetryUI();}
function launchAllPendingWorkers(mode){DEMO_PENDING_STATE="manual";updatePendingRetryUI();return DEMO_PENDING_COUNT;}
function launchBackgroundWorker(){return false;}
function retryPendingPackages(){
    DEMO_PENDING_STATE="manual";
    updatePendingRetryUI();
    window.setTimeout(function(){
        alert("【仕様書用デモ】\n\n再送ボタン押下時の表示確認です。実際のCSV書き込み・Pending再送は行いません。");
        DEMO_PENDING_STATE="idle";
        updatePendingRetryUI();
    },700);
}
function sendRequest(){
    if(!validateForm()){return;}
    alert("【仕様書用デモ】\n\n実際のCSV書き込み・Pending作成・Windows通知は行いません。");
}
