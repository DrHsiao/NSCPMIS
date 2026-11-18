
function menuOnclick(sysId) {
    let oDiv = document.getElementById(sysId);
    if (oDiv["style"]["display"] == "") {
        oDiv["style"]["display"] = "none";
        oDiv["style"]["visibility"] = "hidden";
    }
    else {
        oDiv["style"]["display"] = "";
        oDiv["style"]["visibility"] = "visible";
    }
}
