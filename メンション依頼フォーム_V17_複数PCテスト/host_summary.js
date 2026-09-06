(function () {
    var a = WScript.Arguments;
    if (a.length < 2) WScript.Quit(2);
    var sessionRoot = a.Item(0);
    var outputPath = a.Item(1);

    var fso = new ActiveXObject("Scripting.FileSystemObject");

    function readAnsi(path) {
        var ts = fso.OpenTextFile(path,1,false,0);
        var s = ts.ReadAll();
        ts.Close();
        return s.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
    }
    function config(path) {
        var o={}, lines=readAnsi(path).split("\n");
        for(var i=0;i<lines.length;i++){
            var pos=lines[i].indexOf("=");
            if(pos>0) o[lines[i].substr(0,pos)] = lines[i].substr(pos+1);
        }
        return o;
    }
    function csvParse(line) {
        var out=[], cur="", q=false, i=0;
        while(i<line.length){
            var ch=line.charAt(i);
            if(q){
                if(ch === '"'){
                    if(i+1<line.length && line.charAt(i+1)==='"'){ cur+='"'; i+=2; continue; }
                    q=false; i++; continue;
                }
                cur+=ch; i++; continue;
            } else {
                if(ch === '"'){ q=true; i++; continue; }
                if(ch === ','){ out.push(cur); cur=""; i++; continue; }
                cur+=ch; i++; continue;
            }
        }
        out.push(cur);
        return out;
    }
    function listFiles(folder, suffix) {
        var arr=[];
        if(!fso.FolderExists(folder)) return arr;
        var e=new Enumerator(fso.GetFolder(folder).Files);
        for(;!e.atEnd();e.moveNext()){
            var p=e.item().Path;
            if(!suffix || String(e.item().Name).toLowerCase().slice(-suffix.length)===suffix) arr.push(p);
        }
        return arr;
    }
    function val(o,k){ return o.hasOwnProperty(k) ? o[k] : ""; }

    var cfg=config(fso.BuildPath(sessionRoot,"config.txt"));
    var expectedParticipants=parseInt(cfg.EXPECTED_PARTICIPANTS,10);
    var packages=parseInt(cfg.PACKAGES_PER_PARTICIPANT,10);
    var rows=parseInt(cfg.ROWS_PER_PENDING,10);
    var scenario=cfg.SCENARIO;
    var envMode=cfg.ENVIRONMENT;
    var sessionId=cfg.SESSION_ID;

    var expectedKeys={}, expectedRequestIds={};
    var manifestFiles=listFiles(fso.BuildPath(sessionRoot,"manifests"),".txt");
    for(var m=0;m<manifestFiles.length;m++){
        var lines=readAnsi(manifestFiles[m]).split("\n");
        for(var i=0;i<lines.length;i++){
            if(!lines[i]) continue;
            expectedKeys[lines[i]]=true;
            var p=lines[i].lastIndexOf("|");
            if(p>0) expectedRequestIds[lines[i].substr(0,p)]=true;
        }
    }

    var agg={SUCCESS:0,DUPLICATE_SKIP:0,RECOVERED_PARTIAL:0,PENDING_DELETE_FAILED:0,
             RETRY_EXHAUSTED:0,FATAL:0,OTHER:0,RETRY_TOTAL:0,PENDING_REMAINING:0};
    var machines={}, participants={}, badReports=0;
    var reportFiles=listFiles(fso.BuildPath(sessionRoot,"reports"),".txt");

    for(var r=0;r<reportFiles.length;r++){
        var o=config(reportFiles[r]);
        var pid=val(o,"PARTICIPANT");
        var machine=val(o,"MACHINE");
        if(!pid || !machine){ badReports++; continue; }
        participants[pid]=true;
        machines[machine]=true;
        var keys=["SUCCESS","DUPLICATE_SKIP","RECOVERED_PARTIAL","PENDING_DELETE_FAILED",
                  "RETRY_EXHAUSTED","FATAL","OTHER","RETRY_TOTAL","PENDING_REMAINING"];
        for(var k=0;k<keys.length;k++){
            var n=parseInt(val(o,keys[k]) || "0",10);
            if(!isNaN(n)) agg[keys[k]]+=n;
        }
        var expectedRes=parseInt(val(o,"EXPECTED_RESULTS")||"0",10);
        var actualRes=parseInt(val(o,"RESULT_FILES")||"0",10);
        if(expectedRes!==actualRes) badReports++;
    }

    var csvFolder=fso.BuildPath(sessionRoot,"csv");
    var csvFiles=listFiles(csvFolder,".csv");
    var actualKeys={}, actualSessionRows=0, duplicateRows=0, badColumnRows=0;

    if(csvFiles.length===1){
        var lines=readAnsi(csvFiles[0]).split("\n");
        for(var c=1;c<lines.length;c++){
            if(!lines[c]) continue;
            var f=csvParse(lines[c]);
            if(f.length!==17){ badColumnRows++; continue; }
            var rid=f[1], rno=f[4];
            if(rid.indexOf("V17-"+sessionId+"-")===0){
                actualSessionRows++;
                var key=rid+"|"+rno;
                if(actualKeys.hasOwnProperty(key)) duplicateRows++;
                actualKeys[key]=(actualKeys[key]||0)+1;
            }
        }
    }

    var missing=0, unexpected=0;
    for(var ek in expectedKeys) if(expectedKeys.hasOwnProperty(ek) && !actualKeys.hasOwnProperty(ek)) missing++;
    for(var ak in actualKeys) if(actualKeys.hasOwnProperty(ak) && !expectedKeys.hasOwnProperty(ak)) unexpected++;

    var sharedPackages=0, uniquePackages=0;
    if(scenario==="UNIQUE") uniquePackages=packages;
    if(scenario==="DUPLICATE") sharedPackages=packages;
    if(scenario==="MIXED"){
        sharedPackages=Math.ceil(packages/2);
        uniquePackages=Math.floor(packages/2);
    }

    var expectedSuccess=(uniquePackages*expectedParticipants)+sharedPackages;
    var expectedDupSkip=sharedPackages*(expectedParticipants-1);
    var expectedFinalRows=((uniquePackages*expectedParticipants)+sharedPackages)*rows;

    var participantCount=0, machineCount=0, expectedKeyCount=0, actualKeyCount=0;
    for(var p1 in participants) if(participants.hasOwnProperty(p1)) participantCount++;
    for(var m1 in machines) if(machines.hasOwnProperty(m1)) machineCount++;
    for(var e1 in expectedKeys) if(expectedKeys.hasOwnProperty(e1)) expectedKeyCount++;
    for(var a1 in actualKeys) if(actualKeys.hasOwnProperty(a1)) actualKeyCount++;

    var lockExists=false;
    if(csvFiles.length===1) lockExists=fso.FileExists(csvFiles[0]+".lock");

    var dataPass=(
        reportFiles.length===expectedParticipants &&
        participantCount===expectedParticipants &&
        manifestFiles.length===expectedParticipants &&
        badReports===0 &&
        agg.SUCCESS===expectedSuccess &&
        agg.DUPLICATE_SKIP===expectedDupSkip &&
        agg.RECOVERED_PARTIAL===0 &&
        agg.PENDING_DELETE_FAILED===0 &&
        agg.RETRY_EXHAUSTED===0 &&
        agg.FATAL===0 &&
        agg.OTHER===0 &&
        agg.PENDING_REMAINING===0 &&
        csvFiles.length===1 &&
        expectedKeyCount===expectedFinalRows &&
        actualSessionRows===expectedFinalRows &&
        actualKeyCount===expectedFinalRows &&
        duplicateRows===0 &&
        badColumnRows===0 &&
        missing===0 &&
        unexpected===0 &&
        !lockExists
    );

    var result="CHECK";
    if(dataPass && envMode==="LOCAL_SIM") result="PASS_LOCAL_SIM";
    if(dataPass && envMode==="MULTI_PC" && machineCount>=2) result="PASS";
    if(dataPass && envMode==="MULTI_PC" && machineCount<2) result="CHECK_SINGLE_MACHINE_ONLY";

    var out=[];
    out.push("Mention Request Form v26 - Multi-PC Shared Folder Test v17");
    out.push("============================================================");
    out.push("Session ID             : "+sessionId);
    out.push("Environment            : "+envMode);
    out.push("Scenario               : "+scenario);
    out.push("Expected participants  : "+expectedParticipants);
    out.push("Participant reports    : "+reportFiles.length);
    out.push("Unique machine names   : "+machineCount);
    out.push("Packages / participant : "+packages);
    out.push("Rows / Pending         : "+rows);
    out.push("");
    out.push("SUCCESS                : "+agg.SUCCESS+" / expected "+expectedSuccess);
    out.push("DUPLICATE_SKIP         : "+agg.DUPLICATE_SKIP+" / expected "+expectedDupSkip);
    out.push("RECOVERED_PARTIAL      : "+agg.RECOVERED_PARTIAL);
    out.push("PENDING_DELETE_FAILED  : "+agg.PENDING_DELETE_FAILED);
    out.push("RETRY_EXHAUSTED        : "+agg.RETRY_EXHAUSTED);
    out.push("FATAL / OTHER          : "+(agg.FATAL+agg.OTHER));
    out.push("RETRY total            : "+agg.RETRY_TOTAL);
    out.push("Pending remaining      : "+agg.PENDING_REMAINING);
    out.push("");
    out.push("Expected CSV rows      : "+expectedFinalRows);
    out.push("Actual V17 CSV rows    : "+actualSessionRows);
    out.push("Expected unique keys   : "+expectedKeyCount);
    out.push("Actual unique keys     : "+actualKeyCount);
    out.push("Missing keys           : "+missing);
    out.push("Unexpected keys        : "+unexpected);
    out.push("CSV duplicate rows     : "+duplicateRows);
    out.push("Bad-column rows        : "+badColumnRows);
    out.push("Lock file remaining    : "+(lockExists ? "YES" : "NO"));
    out.push("");
    out.push("RESULT                  : "+result);
    out.push("============================================================");
    if(result==="PASS_LOCAL_SIM"){
        out.push("LOCAL_SIM validates the V17 flow on one PC.");
        out.push("It does NOT validate SMB locking between physical PCs.");
    }
    if(result==="PASS"){
        out.push("At least two different machine names participated.");
        out.push("This run validates shared-folder contention across multiple PCs.");
    }
    if(result==="CHECK_SINGLE_MACHINE_ONLY"){
        out.push("Data consistency passed, but only one machine name was detected.");
        out.push("Run again from two or more physical PCs for multi-PC validation.");
    }

    var ts=fso.CreateTextFile(outputPath,true,false);
    for(var z=0;z<out.length;z++) ts.WriteLine(out[z]);
    ts.Close();
    for(var y=0;y<out.length;y++) WScript.Echo(out[y]);
})();