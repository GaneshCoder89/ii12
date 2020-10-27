function showFrameListWeb() {
    var e = window.mydata;
    document.getElementById("idFrame").src = "rUp.html?data=" + e,
    document.getElementById("idFrame").style.visibility = "visible",
    document.getElementById("idFrame").width = "98%"
}
function showFrameAndroid() {
    document.getElementById("idFrame").width = "40%"
}
function onUpload(e) {
    var t = new FileReader;
    t.onload = function() {
        for (var e = t.result, n = new Uint8Array(e), d = new Array, i = 0; i != n.length; ++i)
            d[i] = String.fromCharCode(n[i]);
        var m = d.join("")
          , l = XLSX.read(m, {
            type: "binary"
        });
        l.SheetNames.forEach(function(e) {
            var t = l.Sheets[e]
              , n = {}
              , d = [];
            for (z in t)
                if ("!" !== z[0]) {
                    for (var i = 0, m = 0; m < z.length; m++)
                        if (!isNaN(z[m])) {
                            i = m;
                            break
                        }
                    var o = z.substring(0, i)
                      , a = parseInt(z.substring(i))
                      , r = t[z].v;
                    1 == a && r ? n[o] = r : (d[a] || (d[a] = {}),
                    d[a][n[o]] = r)
                }
            d.shift(),
            d.shift()
        });
        var o = l.SheetNames[0]
          , a = l.Sheets[o]
          , r = XLSX.utils.sheet_to_json(a, {
            blankRows: !1,
            defval: "",
            raw: !0
        });
        localStorage.aDxl = JSON.stringify(r),
        document.getElementById("idProgressBar").style.width = "100%",
        document.getElementById("idViews").style.visibility = "visible"
    }
    ,
    t.readAsArrayBuffer(document.getElementById("updS").files[0])
}
function showScan() {
    document.getElementById("idFrame").src = "sc.html",
    document.getElementById("idFrame").style.visibility = "visible";
   // document.getElementById("idFrame").width = "98%";

}

function showStep1() {
    document.getElementById("idFrame").src = "st1.html",
    document.getElementById("idFrame").style.visibility = "visible";
   // document.getElementById("idFrame").width = "98%"

}

function showStep2() {
    document.getElementById("idFrame").src = "st2.html",
    document.getElementById("idFrame").style.visibility = "visible"
   // document.getElementById("idFrame").width = "98%"

}

function showStep3() {
    document.getElementById("idFrame").src = "st3.html",
    document.getElementById("idFrame").style.visibility = "visible";
   // document.getElementById("idFrame").width = "98%"

}

function showStep4() {
    document.getElementById("idFrame").src = "st4.html",
    document.getElementById("idFrame").style.visibility = "visible";
  //  document.getElementById("idFrame").width = "98%"

}



function showFrameAnalytics() {
    document.getElementById("idFrame").src = "smUp.html",
    document.getElementById("idFrame").style.visibility = "visible",
    document.getElementById("idFrame").width = "98%"
}
function showFrameWeb() {
    document.getElementById("idFrame").width = "98%"
}
function showFrameKanban() {
    document.getElementById("idFrame").src = "page.html",
    document.getElementById("idFrame").style.visibility = "visible",
    document.getElementById("idFrame").width = "98%"
}
function showFormWeb() {
    document.getElementById("idFrame").src = "foUp.html",
    document.getElementById("idFrame").style.visibility = "visible",
    document.getElementById("idFrame").width = "98%"
}
function ShowERP() {
    document.getElementById("erp").style.display = "initial",
    document.getElementById("btnAPI").style.display = "initial",
    document.getElementById("btnUpd").style.display = "none",
    document.getElementById("prN").value = "Manage Production order",
    document.getElementById("prN").disabled = !0,
    document.getElementById("check2").checked = !1,
    document.getElementById("bp").value = "Production order"
}
function ShowAPI() {
    document.getElementById("api").style.visibility = "visible"
}
function ShowExcel() {
    document.getElementById("erp").style.display = "none",
    document.getElementById("btnAPI").style.display = "none",
    document.getElementById("btnUpd").style.display = "initial",
    document.getElementById("check1").checked = !1,
    document.getElementById("prN").value = "",
    document.getElementById("prN").disabled = !1
}
function onB() {
    window.location.href = "index.html"
}
