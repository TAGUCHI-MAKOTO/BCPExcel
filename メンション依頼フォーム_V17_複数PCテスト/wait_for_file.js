(function () {
    var a = WScript.Arguments;
    if (a.length < 2) WScript.Quit(2);
    var path = a.Item(0);
    var maxSec = parseInt(a.Item(1),10);
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var waited = 0;
    while (!fso.FileExists(path) && waited < maxSec * 10) {
        WScript.Sleep(100);
        waited++;
    }
    if (fso.FileExists(path)) WScript.Quit(0);
    WScript.Quit(9);
})();