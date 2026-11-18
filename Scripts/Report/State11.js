$(document).ready(function () {
    $('.selectALL').on('change', function () {
        var checkboxes = $('.City-checkbox');
        checkboxes.prop('checked', $(this).prop('checked'));
    });
});

$(function () {
    $("#choose").change(function () {
        var selectedValue = $(this).val();
        $.ajax({
            url: stateUrl,
            type: 'post',
            cache: false,
            async: false,
            data: { value: selectedValue },
            success: function (result) {
                window.location.href = result;
            }
        });
    });
});

function uncheckCancel() {
    if (document.getElementsByName('Row')[0].checked) {
        document.getElementsByName('Field')[0].checked = false;
    }
}

function uncheckFinish() {
    if (document.getElementsByName('Field')[0].checked) {
        document.getElementsByName('Row')[0].checked = false;
    }
}

document.getElementById("downloadButton").addEventListener("click", function () {
    const downloadLink = document.getElementById("downloadLink");
    downloadMessage.style.display = "none";

    var currentDateTime = new Date();
    var formattedDateTime = currentDateTime.toISOString().replace(/[-:.T]/g, "").slice(0, 14);
    if ($('#one').is(':checked')) {
        var filename = "前瞻基礎建設計畫-縣市管河川及區域排水整體改善計畫" + formattedDateTime;
    }
    else {
        var filename = "都市總合治水計畫控管表" + formattedDateTime;
    }
    var plan = [];
    var checkboxes = document.querySelectorAll('input[name="plan"]:checked');
    for (var i = 0; i < checkboxes.length; i++) {
        plan.push(checkboxes[i].value);
    }
    var CITY = [];
    var checkboxes2 = document.querySelectorAll('input[name="CITY"]:checked');
    for (var i = 0; i < checkboxes2.length; i++) {
        CITY.push(checkboxes2[i].value);
    }
    var Year1 = document.getElementsByName("year1")[0].value;
    var Year2 = document.getElementsByName("year2")[0].value;
    var Month = document.getElementById("monthSelect").value;

    var downloadButton = document.getElementById("downloadButton"); // 使用 document.querySelector 也可以
    downloadButton.innerHTML = '<span class="loading-spinner"></span>下載中 　';

    var url = reportUrl +
        '?plan=' + encodeURIComponent(plan.join(',')) +
        '&CITY=' + encodeURIComponent(CITY.join(',')) +
        '&Year1=' + encodeURIComponent(Year1) +
        '&Year2=' + encodeURIComponent(Year2) +
        '&month=' + encodeURIComponent(Month) +
        '&type=' + encodeURIComponent('xlsx');

    fetch(url, {
        method: "GET",
        headers: {
            "Content-Type": "application/json"
        }
    })
        .then(response => response.blob())
        .then(blob => {
            const downloadLink = document.getElementById("downloadLink");
            const downloadMessage = document.getElementById("downloadMessage");

            // 更新下載連結的屬性和顯示訊息
            const blobUrl = URL.createObjectURL(blob);
            downloadLink.href = blobUrl;
            downloadLink.download = filename + ".xlsx";
            downloadLink.click();

            // 恢復按鈕文本
            downloadButton.innerHTML = 'EXCEL匯出';
        })
        .catch(error => {
            console.error("Error:", error);
            // 處理錯誤情況
            downloadButton.innerHTML = 'EXCEL匯出';
        });
});

document.getElementById("downloadButton2").addEventListener("click", function () {
    const downloadLink = document.getElementById("downloadLink");
    downloadMessage.style.display = "none";

    var currentDateTime = new Date();
    var formattedDateTime = currentDateTime.toISOString().replace(/[-:.T]/g, "").slice(0, 14);
    if ($('#one').is(':checked')) {
        var filename = "前瞻基礎建設計畫-縣市管河川及區域排水整體改善計畫" + formattedDateTime;
    }
    else {
        var filename = "都市總合治水計畫控管表" + formattedDateTime;
    }
    var plan = [];
    var checkboxes = document.querySelectorAll('input[name="plan"]:checked');
    for (var i = 0; i < checkboxes.length; i++) {
        plan.push(checkboxes[i].value);
    }
    var CITY = [];
    var checkboxes2 = document.querySelectorAll('input[name="CITY"]:checked');
    for (var i = 0; i < checkboxes2.length; i++) {
        CITY.push(checkboxes2[i].value);
    }
    var Year1 = document.getElementsByName("year1")[0].value;
    var Year2 = document.getElementsByName("year2")[0].value;
    var Month = document.getElementById("monthSelect").value;

    var downloadButton = document.getElementById("downloadButton2"); // 使用 document.querySelector 也可以
    downloadButton.innerHTML = '<span class="loading-spinner"></span>下載中 　';

    var url = reportUrl +
        '?plan=' + encodeURIComponent(plan.join(',')) +
        '&CITY=' + encodeURIComponent(CITY.join(',')) +
        '&Year1=' + encodeURIComponent(Year1) +
        '&Year2=' + encodeURIComponent(Year2) +
        '&month=' + encodeURIComponent(Month) +
        '&type=ODS';

    fetch(url, {
        method: "GET",
        headers: {
            "Content-Type": "application/json"
        }
    })
        .then(response => response.blob())
        .then(blob => {
            const downloadLink = document.getElementById("downloadLink");
            const downloadMessage = document.getElementById("downloadMessage");

            // 更新下載連結的屬性和顯示訊息
            const blobUrl = URL.createObjectURL(blob);
            downloadLink.href = blobUrl;
            downloadLink.download = filename + ".ods";
            downloadLink.click();

            // 恢復按鈕文本
            downloadButton.innerHTML = 'ODS匯出';
        })
        .catch(error => {
            console.error("Error:", error);
            // 處理錯誤情況
            downloadButton.innerHTML = 'ODS匯出';
        });
});