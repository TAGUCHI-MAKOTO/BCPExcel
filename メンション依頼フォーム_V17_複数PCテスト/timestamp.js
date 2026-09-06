(function () {
    function p2(n){ return n < 10 ? "0" + n : "" + n; }
    function p3(n){
        if(n < 10) return "00" + n;
        if(n < 100) return "0" + n;
        return "" + n;
    }
    var d = new Date();
    var tz = -d.getTimezoneOffset();
    var sign = tz >= 0 ? "+" : "-";
    var az = Math.abs(tz);
    var oh = Math.floor(az / 60);
    var om = az % 60;
    WScript.Echo(
        d.getFullYear() + "-" + p2(d.getMonth()+1) + "-" + p2(d.getDate()) + "T" +
        p2(d.getHours()) + ":" + p2(d.getMinutes()) + ":" + p2(d.getSeconds()) + "." +
        p3(d.getMilliseconds()) + sign + p2(oh) + ":" + p2(om)
    );
})();