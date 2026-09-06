(function () {
    var a = WScript.Arguments;
    if (a.length < 9) {
        WScript.Echo("ERROR: missing args");
        WScript.Quit(2);
    }

    var pendingDir = a.Item(0);
    var csvFolder = a.Item(1);
    var participant = a.Item(2);
    var sessionId = a.Item(3);
    var scenario = a.Item(4);
    var packages = parseInt(a.Item(5),10);
    var rowsPer = parseInt(a.Item(6),10);
    var created = a.Item(7);
    var manifestPath = a.Item(8);

    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (!fso.FolderExists(pendingDir)) fso.CreateFolder(pendingDir);

    function csvQuote(v) {
        var s = String(v == null ? "" : v);
        return '"' + s.replace(/"/g, '""') + '"';
    }
    function enc(s){ return encodeURIComponent(String(s == null ? "" : s)); }

    var baseName = "V17TEST";
    var manifest = fso.CreateTextFile(manifestPath, true, false);

    for (var i=1; i<=packages; i++) {
        var shared = false;
        if (scenario === "DUPLICATE") shared = true;
        if (scenario === "MIXED" && (i % 2 === 1)) shared = true;

        var requestId;
        if (shared) {
            requestId = "V17-" + sessionId + "-SHARED-" + i;
        } else {
            requestId = "V17-" + sessionId + "-" + participant + "-" + i;
        }

        var machine = [];
        for (var r=1; r<=rowsPer; r++) {
            var requester = shared ? "v17_shared" : ("v17_" + participant);
            var org = shared ? ("V17_SHARED_ORG_" + i) : ("V17_" + participant + "_ORG_" + i);
            var ca = shared ? ("V17_SHARED_CA_" + i) : ("V17_" + participant + "_CA_" + i);
            var memo = shared ? ("V17_SHARED_" + i + "_" + r) : ("V17_" + participant + "_" + i + "_" + r);

            var fields = [
                created, requestId, baseName, requester, String(r),
                org, ca, "", "", "", "", memo, created, "", "", "", "V17TEST"
            ];

            var q = [];
            for (var x=0; x<fields.length; x++) q.push(csvQuote(fields[x]));
            machine.push(q.join(","));
            manifest.WriteLine(requestId + "|" + r);
        }

        var lines = [];
        lines.push("Mention Request Form - V17 Multi-PC Test Pending");
        lines.push("============================================================");
        lines.push("Status: Pending");
        lines.push("Created: " + created);
        lines.push("RequestID: " + requestId);
        lines.push("Participant: " + participant);
        lines.push("============================================================");
        lines.push("----- MACHINE_DATA_BEGIN -----");
        lines.push("FORMAT=MENTION_REQUEST_PENDING_V3");
        lines.push("BASE=" + enc(baseName));
        lines.push("REQUEST_ID=" + enc(requestId));
        lines.push("CREATED=" + enc(created));
        lines.push("CSV_FOLDER=" + enc(csvFolder));
        for (var m=0; m<machine.length; m++) lines.push("LINE=" + enc(machine[m]));
        lines.push("----- MACHINE_DATA_END -----");

        var path = fso.BuildPath(pendingDir, requestId + ".pending");
        var ts = fso.CreateTextFile(path, true, true); // UTF-16LE, like actual Pending test files
        for (var n=0; n<lines.length; n++) ts.WriteLine(lines[n]);
        ts.Close();
    }

    manifest.Close();
    WScript.Echo("CREATED=" + packages);
})();