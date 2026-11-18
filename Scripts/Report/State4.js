function toggleCheckboxes() {
    var checkboxes = document.querySelectorAll('.city-checkbox');
    var selectAllCheckbox = document.getElementById('selectAll');

    for (var i = 0; i < checkboxes.length; i++) {
        checkboxes[i].checked = selectAllCheckbox.checked;
    }
}

function toggleCheckboxes2() {
    var checkboxes = document.querySelectorAll('.column-checkbox');
    var selectAllCheckbox = document.getElementById('selectAll2');

    for (var i = 0; i < checkboxes.length; i++) {
        checkboxes[i].checked = selectAllCheckbox.checked;
    }
}

// 在選擇框改變時觸發
document.getElementById('yearStart').addEventListener('change', function () {
    filterProjects();
});

document.getElementById('yearEnd').addEventListener('change', function () {
    filterProjects();
});

function filterProjects() {
    // 獲取選擇的年度
    var startYear = document.getElementById('yearStart').value;
    var endYear = document.getElementById('yearEnd').value;
    // 顯示或隱藏計劃項目
    var planItems = document.getElementsByClassName('plan-item');
    var seenValue = {}; // 用於追蹤已經看到的值

    // 將所有checkbox的選取狀態設置為空
    var checkboxes = document.querySelectorAll('#plan input[ type="checkbox"]');
    checkboxes.forEach(function (checkbox) {
        checkbox.checked = false;
    });

    for (var i = 0; i < planItems.length; i++) {
        var itemYear = planItems[i].getAttribute('data-id');
        var itemValue = planItems[i].querySelector('input').value;
        if (itemYear >= startYear && itemYear <= endYear) {
            if (seenValue[itemValue]) {
                planItems[i].style.display = 'none';
            } else {
                planItems[i].style.display = 'block';
                seenValue[itemValue] = true;
            }
        } else {
            planItems[i].style.display = 'none';
        }
    }
}

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

$(document).ready(function () {
    //若沒有選擇計畫名稱或縣市別時顯示提示訊息
    //顯示下載中圖示
    document.getElementById('myForm').addEventListener('submit', function (event) {
        var chk = true;
        var chkPlan = $('.Plan-checkbox:checked').length;
        if (chkPlan === 0) {
            alert('請選擇計畫名稱！');
            event.preventDefault();
            chk = false;
        }
        var chkCity = $('.city-checkbox:checked').length;
        if (chkCity === 0) {
            alert('請選擇縣市別！');
            event.preventDefault();
            chk = false;
        }
        var chkColumn = $('.column-checkbox:checked').length;
        if (chkColumn === 0) {
            alert('請選擇顯示欄位！');
            event.preventDefault();
            chk = false;
        }

        if (chk) {
            event.preventDefault();  //防止表單提交，使用 AJAX            

            //獲取表單資料
            var formData = new FormData(event.target);
            var clickedButton = event.submitter;
            formData.append('PDF', clickedButton.value);

            //設定檔案名稱和副檔名
            var filename = "A-4.跨年度細項執行情形表" + Date.now();
            var fileext = ".xlsx";
            var btnText = "EXCEL匯出";
            if (formData.get('PDF') == "1") {
                fileext = ".ods";
                btnText = "ODS匯出";
            }

            //按鈕顯示下載中圖示
            clickedButton.innerHTML = '<span class="loading-spinner"></span>下載中 　';

            //使用 AJAX 提交表單資料
            fetch(reportUrl, {
                method: 'POST',
                body: formData
            })
                .then(response => {
                    if (response.ok) {
                        return response.blob();
                    }
                    throw new Error('錯誤：未能取得檔案');
                })
                .then(blob => {
                    //建立超連結下載檔案
                    var link = document.createElement('a');
                    var url = URL.createObjectURL(blob); //創建 Blob 的 URL
                    link.href = url;
                    link.download = filename + fileext;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);

                    //按鈕恢復預設文字
                    clickedButton.innerHTML = btnText;
                })
                .catch(error => {
                    console.error('錯誤:', error);
                    //按鈕恢復預設文字
                    clickedButton.innerHTML = btnText;
                });
        }
    });
});
