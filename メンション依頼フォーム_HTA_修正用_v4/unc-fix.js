// v4 UNC起動修正（修正確認用）
function getCurrentFolderPath(){
    var rawPath="";
    var host="";
    var path="";

    try{
        rawPath=String(window.location.pathname||"");
    }catch(err){
        rawPath="";
    }

    try{
        host=String(window.location.hostname||"");
    }catch(err2){
        host="";
    }

    try{
        path=decodeURIComponent(rawPath);
    }catch(err3){
        path=rawPath;
    }

    path=path.replace(/\//g,"\\");

    if(host && host.toLowerCase()!=="localhost"){
        while(path.charAt(0)==="\\"){
            path=path.substring(1);
        }
        path="\\\\"+host+"\\"+path;
    }else{
        if(/^\\[A-Za-z]:\\/.test(path)){
            path=path.substring(1);
        }
    }

    var fso=new ActiveXObject("Scripting.FileSystemObject");
    return fso.GetParentFolderName(path);
}

function readProductionConfig(){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var configPath=getCurrentFolderPath()+"\\"+CONFIG_FILE_NAME;
    var result={};

    if(!fso.FileExists(configPath)){
        throw new Error(
            CONFIG_FILE_NAME+" が見つかりません。\n\n"+
            "確認先：\n"+configPath+"\n\n"+
            "HTAと同じフォルダに "+CONFIG_FILE_NAME+" を置いてください。"
        );
    }

    var file=fso.OpenTextFile(configPath,1,false,-1);
    var text=file.ReadAll();
    file.Close();

    var rows=String(text)
        .replace(/\r\n/g,"\n")
        .replace(/\r/g,"\n")
        .split("\n");

    var i,row,pos,key,value;

    for(i=0;i<rows.length;i++){
        row=String(rows[i]||"").replace(/^\s+|\s+$/g,"");

        if(!row || row.charAt(0)===";" || row.charAt(0)==="#"){
            continue;
        }

        pos=row.indexOf("=");

        if(pos<=0){
            continue;
        }

        key=row.substring(0,pos).replace(/^\s+|\s+$/g,"").toUpperCase();
        value=row.substring(pos+1).replace(/^\s+|\s+$/g,"");

        result[key]=value;
    }

    return result;
}
