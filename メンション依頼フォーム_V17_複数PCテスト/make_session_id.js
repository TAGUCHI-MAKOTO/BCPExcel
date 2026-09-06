(function () {
    function p2(n){ return n < 10 ? "0" + n : "" + n; }
    var d = new Date();
    var s = "V17_" + d.getFullYear() + p2(d.getMonth()+1) + p2(d.getDate()) + "_" +
            p2(d.getHours()) + p2(d.getMinutes()) + p2(d.getSeconds());
    WScript.Echo(s);
})();