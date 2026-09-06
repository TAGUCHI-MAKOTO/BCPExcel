(function () {
    var a = WScript.Arguments;
    if (a.length < 7) WScript.Quit(2);

    var resultDir = a.Item(0);
    var pendingDir = a.Item(1);
    var reportPath = a.Item(2);
    var participant = a.Item(3);
    var sessionId = a.Item(4);
    var packages = parseInt(a.Item(5),10);
    var started = a.Item(6);

    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var net = new ActiveXObject("WScript.Network");

    function p2(n){ return n < 10 ? "0" + n : "" + n; }
    function p3(n){
        if(n < 10) return "00" + n;
        if(n < 100) return "0" + n;
        return "" + n;
    }
    function isoLocal(d) {
        var tz = -d.getTimezoneOffset();
        var sign = tz >= 0 ? "+" : "-";
        var az = Math.abs(tz);
        return d.getFullYear() + "-" + p2(d.getMonth()+1) + "-" + p2(d.getDate()) + "T" +
               p2(d.getHours()) + ":" + p2(d.getMinutes()) + ":" + p2(d.getSeconds()) + "." +
               p3(d.getMilliseconds()) + sign + p2(Math.floor(az/60)) + ":" + p2(az%60);
    }

    function readUtf8(path) {
        var st = new ActiveXObject("ADODB.Stream");
        st.Type = 2;
        st.Charset = "utf-8";
        st.Open();
        st.LoadFromFile(path);
        var s = st.ReadText();
        st.Close();
        return s.replace(/^\uFEFF/,"").replace(/\r?\n/g,"").replace(/\r/g,"");
    }

    var counts = {
        SUCCESS:0, DUPLICATE_SKIP:0, RECOVERED_PARTIAL:0,
        PENDING_DELETE_FAILED:0, RETRY_EXHAUSTED:0, FATAL:0, OTHER:0
    };
    var retryTotal=0, durationTotal=0, resultFiles=0;

    if (fso.FolderExists(resultDir)) {
        var e = new Enumerator(fso.GetFolder(resultDir).Files);
        for (; !e.atEnd(); e.moveNext()) {
            var file = e.item();
            if (String(file.Name).toLowerCase().slice(-4) !== ".txt") continue;
            resultFiles++;
            var x = readUtf8(file.Path).split("|");
            var status = x[0] || "OTHER";
            if (typeof counts[status] !== "undefined") counts[status]++; else counts.OTHER++;
            if (x.length > 1 && !isNaN(parseInt(x[1],10))) retryTotal += parseInt(x[1],10);
            if (x.length > 2 && !isNaN(parseFloat(x[2]))) durationTotal += parseFloat(x[2]);
        }
    }

    var pendingRemaining=0;
    if (fso.FolderExists(pendingDir)) {
        var p = new Enumerator(fso.GetFolder(pendingDir).Files);
        for (; !p.atEnd(); p.moveNext()) {
            if (String(p.item().Name).toLowerCase().slice(-8) === ".pending") pendingRemaining++;
        }
    }

    var finished = isoLocal(new Date());
    var out = fso.CreateTextFile(reportPath, true, false);
    out.WriteLine("PARTICIPANT=" + participant);
    out.WriteLine("MACHINE=" + net.ComputerName);
    out.WriteLine("SESSION_ID=" + sessionId);
    out.WriteLine("EXPECTED_RESULTS=" + packages);
    out.WriteLine("RESULT_FILES=" + resultFiles);
    out.WriteLine("SUCCESS=" + counts.SUCCESS);
    out.WriteLine("DUPLICATE_SKIP=" + counts.DUPLICATE_SKIP);
    out.WriteLine("RECOVERED_PARTIAL=" + counts.RECOVERED_PARTIAL);
    out.WriteLine("PENDING_DELETE_FAILED=" + counts.PENDING_DELETE_FAILED);
    out.WriteLine("RETRY_EXHAUSTED=" + counts.RETRY_EXHAUSTED);
    out.WriteLine("FATAL=" + counts.FATAL);
    out.WriteLine("OTHER=" + counts.OTHER);
    out.WriteLine("RETRY_TOTAL=" + retryTotal);
    out.WriteLine("DURATION_TOTAL_MS=" + Math.round(durationTotal*100)/100);
    out.WriteLine("PENDING_REMAINING=" + pendingRemaining);
    out.WriteLine("STARTED=" + started);
    out.WriteLine("FINISHED=" + finished);
    out.Close();
})();