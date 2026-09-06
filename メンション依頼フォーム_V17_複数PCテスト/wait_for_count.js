(function () {
    var a = WScript.Arguments;
    if (a.length < 4) WScript.Quit(2);
    var folder = a.Item(0);
    var expected = parseInt(a.Item(1),10);
    var ext = a.Item(2).toLowerCase();
    var maxSec = parseInt(a.Item(3),10);
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var waited = 0, last = -1;

    function count() {
        var n = 0;
        if (!fso.FolderExists(folder)) return 0;
        var e = new Enumerator(fso.GetFolder(folder).Files);
        for (; !e.atEnd(); e.moveNext()) {
            var name = String(e.item().Name).toLowerCase();
            if (ext === "*" || name.substr(name.length-ext.length) === ext) n++;
        }
        return n;
    }

    while (waited < maxSec * 2) {
        var c = count();
        if (c !== last) {
            WScript.Echo("PROGRESS=" + c + "/" + expected);
            last = c;
        }
        if (c >= expected) WScript.Quit(0);
        WScript.Sleep(500);
        waited++;
    }
    WScript.Echo("TIMEOUT_COUNT=" + count());
    WScript.Quit(9);
})();