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

function toggleCheckboxes3() {
    var checkboxes = document.querySelectorAll('.column-checkbox');
    var selectAllCheckbox = document.getElementById('selectAll3');

    for (var i = 0; i <= 9; i++) {
        checkboxes[i].checked = selectAllCheckbox.checked;
    }
}

function toggleCheckboxes4() {
    var checkboxes = document.querySelectorAll('.column-checkbox');
    var selectAllCheckbox = document.getElementById('selectAll4');

    for (var i = 10; i <= 16; i++) {
        checkboxes[i].checked = selectAllCheckbox.checked;
    }
}

function toggleCheckboxes5() {
    var checkboxes = document.querySelectorAll('.column-checkbox');
    var selectAllCheckbox = document.getElementById('selectAll5');
    var checkState = selectAllCheckbox.checked;
    for (var i = 17; i < checkboxes.length; i++) {
        if (checkboxes[i]) {
            checkboxes[i].checked = checkState;
        }
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

    var type = "";

    var radioInputs = document.querySelectorAll('input[name="water_type"]');
    radioInputs.forEach(radio => {
        if (radio.checked) {
            type = radio.value;
        }
    });

    // 將所有checkbox的選取狀態設置為空
    var checkboxes = document.querySelectorAll('#plan input[ type="checkbox"]');
    checkboxes.forEach(function (checkbox) {
        checkbox.checked = false;
    });

    for (var i = 0; i < planItems.length; i++) {
        var itemYear = planItems[i].getAttribute('data-id');
        var itemValue = planItems[i].querySelector('input').value;
        var itemType = planItems[i].getAttribute('data-Water_Type');
        if (itemYear >= startYear && itemYear <= endYear && itemType === type) {
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

var checkbox = document.getElementsByClassName("Type-checkbox");
for (var i = 0; i < checkbox.length; i++) {
    checkbox[i].addEventListener('change', function () {
        if (this.checked) {
            // 先重設 StartYear 下拉選單
            var startYearSelect = document.getElementById("yearStart");
            startYearSelect.selectedIndex = 0; // 將選取的索引設為0,表示選擇第一個選項(也就是"選擇")
            // 先重設 StartYear 下拉選單
            var endYearSelect = document.getElementById("yearEnd");
            endYearSelect.selectedIndex = 0; // 將選取的索引設為0,表示選擇第一個選項(也就是"選擇")

            ShowPlan2(this.value);
        }
        else {

        }
    });
}

function ShowPlan2(type) {
    var testElements = document.getElementsByClassName("plan-item");
    for (var i = 0; i < testElements.length; i++) {
        var element = testElements[i];
        var ID = element.getAttribute("data-PID");
        if (element.getAttribute("data-Water_Type") === type) {
            element.style.display = "block";
        } else {
            element.style.display = 'none';
        }
    }
}

$(document).ready(function () {
    $('#Year').change(function () { ChangeCustomer(); });

    ////若沒有選擇計畫名稱或縣市別時顯示提示訊息
    //$('.my-text button[type="submit"]').on('click', function (event) {
    //    var chkPlan = $('.Plan-checkbox:checked').length;
    //    if (chkPlan === 0) {
    //        alert('請選擇計畫名稱！');
    //        event.preventDefault();
    //    }
    //    var chkCity = $('.city-checkbox:checked').length;
    //    if (chkCity === 0) {
    //        alert('請選擇縣市別！');
    //        event.preventDefault();
    //    }
    //    var chkColumn = $('.column-checkbox:checked').length;
    //    if (chkColumn === 0) {
    //        alert('請選擇顯示欄位！');
    //        event.preventDefault();
    //    }
    //});
});

function ChangeCustomer() {
    var selectedValue = $('#Year option:selected').val();
    if ($.trim(selectedValue).length > 0) {
        if (currentUser != null) {
            GetOrders2(selectedValue);
        }
        else {
            GetOrders(selectedValue);
        }
    }
}

function GetOrders(Year) {
    $.ajax({
        url: orderUrl,
        data: { selectYear: Year, type: '污水使用' },
        type: 'post',
        cache: false,
        async: false,
        dataType: 'json',
        success: function (data1) {
            var data = JSON.parse(data1);
            if (data.length > 0) {
                $('#Value').empty();
                $('#Value').append($('<option value="" disabled selected></option>').val('').text('請選擇'));
                $.each(data, function (i, item) {
                    if (item.Text.split('-')[0].trim() === '') {
                        // 如果 A 为空值，跳过当前迴圈
                        return true; // 相当于 continue
                    }
                    $('#Value').append($('<option></option>').val(item.Value).text(item.Text));
                })
            }
            else {
                $('#Value').empty();
                $('#Value').append($('<option></option>').val('').text('------'));
            }
        }
    });
}

function GetOrders2(Year) {
    $.ajax({
        url: orderUrl,
        data: { selectYear: Year, type: '雨水使用' },
        type: 'post',
        cache: false,
        async: false,
        dataType: 'json',
        success: function (data1) {
            var data = JSON.parse(data1);
            if (data.length > 0) {
                $('#Value').empty();
                $('#Value').append($('<option value="" disabled selected></option>').val('').text('請選擇'));
                $.each(data, function (i, item) {
                    if (item.Text.split('-')[0].trim() === '') {
                        // 如果 A 为空值，跳过当前迴圈
                        return true; // 相当于 continue
                    }
                    $('#Value').append($('<option></option>').val(item.Value).text(item.Text));
                })
            }
            else {
                $('#Value').empty();
                $('#Value').append($('<option></option>').val('').text('------'));
            }
        }
    });
}

$(document).ready(function () {
    //若沒有選擇計畫名稱或縣市別時顯示提示訊息
    //顯示下載中圖示
    document.getElementById('dynamicForm').addEventListener('submit', function (event) {
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
            var filename = "E-1.工程細項執行表(縣市在列)" + Date.now();
            if (formData.get('Field') == "on") {
                filename = "E-1.工程細項執行表(縣市在欄)" + Date.now();
            }
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