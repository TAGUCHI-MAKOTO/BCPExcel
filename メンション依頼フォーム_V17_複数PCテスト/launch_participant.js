(function () {
    var a = WScript.Arguments;
    if (a.length < 4) WScript.Quit(2);
    var shell = new ActiveXObject("WScript.Shell");
    var cmd = 'cmd.exe /d /c ""' + a.Item(0) + '" "' + a.Item(1) + '" "' +
              a.Item(2) + '" "' + a.Item(3) + '" AUTO"';
    shell.Run(cmd, 0, false);
})();