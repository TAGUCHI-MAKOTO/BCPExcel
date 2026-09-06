(function () {
    var a=WScript.Arguments;
    var folder=a.Item(0), ext=a.Item(1).toLowerCase();
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var n=0;
    if (fso.FolderExists(folder)) {
        var e=new Enumerator(fso.GetFolder(folder).Files);
        for(;!e.atEnd();e.moveNext()){
            var name=String(e.item().Name).toLowerCase();
            if(ext==="*" || name.substr(name.length-ext.length)===ext) n++;
        }
    }
    WScript.Echo(n);
})();