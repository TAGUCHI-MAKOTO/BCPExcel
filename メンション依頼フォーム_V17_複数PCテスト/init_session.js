(function () {
    var a = WScript.Arguments;
    if (a.length < 7) {
        WScript.Echo("ERROR: missing args");
        WScript.Quit(2);
    }

    var sessionRoot = a.Item(0);
    var sessionId = a.Item(1);
    var envMode = a.Item(2);
    var expected = parseInt(a.Item(3),10);
    var scenario = a.Item(4);
    var packages = parseInt(a.Item(5),10);
    var rows = parseInt(a.Item(6),10);

    var fso = new ActiveXObject("Scripting.FileSystemObject");

    function ensure(path) {
        if (!fso.FolderExists(path)) fso.CreateFolder(path);
    }
    function p2(n){ return n < 10 ? "0" + n : "" + n; }
    function nowText() {
        var d = new Date();
        return d.getFullYear() + "/" + p2(d.getMonth()+1) + "/" + p2(d.getDate()) + " " +
               p2(d.getHours()) + ":" + p2(d.getMinutes()) + ":" + p2(d.getSeconds());
    }

    ensure(sessionRoot);
    ensure(fso.BuildPath(sessionRoot, "control"));
    ensure(fso.BuildPath(sessionRoot, "ready"));
    ensure(fso.BuildPath(sessionRoot, "done"));
    ensure(fso.BuildPath(sessionRoot, "reports"));
    ensure(fso.BuildPath(sessionRoot, "manifests"));
    ensure(fso.BuildPath(sessionRoot, "csv"));

    var created = nowText();
    var cfg = fso.CreateTextFile(fso.BuildPath(sessionRoot, "config.txt"), true, false);
    cfg.WriteLine("SESSION_ID=" + sessionId);
    cfg.WriteLine("ENVIRONMENT=" + envMode);
    cfg.WriteLine("EXPECTED_PARTICIPANTS=" + expected);
    cfg.WriteLine("SCENARIO=" + scenario);
    cfg.WriteLine("PACKAGES_PER_PARTICIPANT=" + packages);
    cfg.WriteLine("ROWS_PER_PENDING=" + rows);
    cfg.WriteLine("CREATED=" + created);
    cfg.WriteLine("BASE_NAME=V17TEST");
    cfg.Close();

    var csvPath = fso.BuildPath(fso.BuildPath(sessionRoot, "csv"),
                                "V17TEST_" + created.substr(0,10).replace(/\//g,"") + ".csv");

    var header = decodeURIComponent(
        "%E9%80%81%E4%BF%A1%E6%97%A5%E6%99%82%2CRequestID%2C%E6%8B%A0%E7%82%B9%2C%E4%BE%9D%E9%A0%BC%E8%80%85%2C%E4%BE%9D%E9%A0%BC%E7%95%AA%E5%8F%B7%2C%E7%B5%84%E7%B9%94%2CCA%E5%90%8D%2C%E4%BB%A3%E7%90%86CA%E7%B5%84%E7%B9%94%2C%E4%BB%A3%E7%90%86CA1%2C%E4%BB%A3%E7%90%86CA2%2C%E4%BB%A3%E7%90%86CA3%2C%E3%83%A1%E3%83%BC%E3%83%AB%E3%83%A1%E3%83%A2%2C%E5%87%A6%E7%90%86%E6%97%A5%E6%99%82%2C%E6%9C%9F%E6%97%A5%2C%E6%99%82%E7%9F%AD%2C%E8%87%B3%E6%80%A5%2C%E3%82%BF%E3%82%A4%E3%83%97"
    );

    var csv = fso.CreateTextFile(csvPath, true, false);
    csv.WriteLine(header);
    csv.Close();

    WScript.Echo("SESSION_ROOT=" + sessionRoot);
    WScript.Echo("CSV_PATH=" + csvPath);
    WScript.Quit(0);
})();