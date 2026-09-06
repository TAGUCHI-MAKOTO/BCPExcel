(function () {
    var a = WScript.Arguments;
    if (a.length < 2) WScript.Quit(2);
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var p = a.Item(0);
    var text = a.Item(1);
    var ts = fso.CreateTextFile(p, true, false);
    ts.WriteLine(text);
    ts.Close();
})();