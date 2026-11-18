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

var checkbox = document.getElementsByClassName("Year-checkbox");
for (var i = 0; i < checkbox.length; i++) {
    checkbox[i].addEventListener('change', function () {
        if (this.checked) {
            ShowPlanAndCheck(this.value);
        }
        else {
            CloseAndDeleteAndCancelCheck(this.value);
        }
    });
}

var planCheckboxes = document.getElementsByClassName("Town-checkbox");
for (var i = 0; i < planCheckboxes.length; i++) {
    planCheckboxes[i].addEventListener('change', function () {
        if (this.checked) {
            var checkboxes = document.getElementsByClassName("Year-checkbox");
            var ID = this.getAttribute('data-id').toString();
            for (var j = 0; j < checkboxes.length; j++) {
                if (checkboxes[j].getAttribute("data-id") === ID) {
                    checkboxes[j].checked = true;
                }
            }
        }
        else {
            var allChecked = false;

            var ID = this.getAttribute('data-id').toString();
            var choosePlanYear = document.querySelectorAll('.Town-checkbox[data-id="' + ID + '"]');
            choosePlanYear.forEach(function (checkbox) {
                if (checkbox.checked) {
                    allChecked = true;
                }
            });

            var checkboxes = document.getElementsByClassName("Year-checkbox");
            for (var j = 0; j < checkboxes.length; j++) {
                if (checkboxes[j].getAttribute("data-id") === ID) {
                    checkboxes[j].checked = allChecked;
                }
            }
        }

        // 點擊,取消點擊鄉鎮別時，設定對應的全選框勾選狀態
        var allChecked = $('.plan-item[data-id=' + ID + '] input[type="checkbox"].Town-checkbox').length === $('.plan-item[data-id=' + ID + '] input[type="checkbox"].Town-checkbox:checked').length;
        $('input[type="checkbox"][class="selectALL"][data-id=' + ID + ']').prop('checked', allChecked);
    });
}

function ShowPlan(test) {
    var testElements = document.getElementsByClassName("plan-item");
    for (var i = 0; i < testElements.length; i++) {
        var element = testElements[i];
        if (element.getAttribute("data-id") === test) {
            element.style.display = "block";
        }
        else {
            element.style.display = 'none';
        }
    }
}

//設定點擊縣市別勾選框時將鄉鎮別全部勾選
function ShowPlanAndCheck(city) {
    var testElements = document.getElementsByClassName("plan-item");
    for (var i = 0; i < testElements.length; i++) {
        var element = testElements[i];
        if (element.getAttribute("data-id") === city) {
            element.style.display = "block";
            var checkboxes = $('input[type="checkbox"]', element);
            checkboxes.each(function () {
                $(this).prop('checked', true);
            });
        } else {
            element.style.display = 'none';
        }
    }
}

//設定取消點擊縣市別勾選框時將鄉鎮別全部取消勾選
function CloseAndDeleteAndCancelCheck(city) {
    var testElements = document.getElementsByClassName("plan-item");
    for (var i = 0; i < testElements.length; i++) {
        var element = testElements[i];
        if (element.getAttribute("data-id") === city) {
            element.style.display = 'none';
            var checkboxes = $('input[type="checkbox"]', element);
            checkboxes.each(function () {
                $(this).prop('checked', false);
            });
        }
    }
}

$(document).ready(function () {
    $('.selectALL').on('change', function () {
        var city = $(this).data('id');
        var checkboxes = $('.' + city);
        checkboxes.prop('checked', $(this).prop('checked'));

        //點擊/取消點擊全選框時，設定對應的縣市別勾選狀態
        $('input[type="checkbox"][name="CITY"][data-id=' + city + ']').prop('checked', $(this).prop('checked'));
    });

    //若沒有選擇計畫分類或縣市鄉鎮別時顯示提示訊息
    $('.my-text button[type="submit"]').on('click', function (event) {
        var chkPlan = $('#Plan input[type="checkbox"]:checked').length;
        if (chkPlan === 0) {
            alert('請選擇計畫分類！');
            event.preventDefault();
        }

        var chkCity = $('#Year input[type="checkbox"]:checked').length;
        if (chkCity === 0) {
            alert('請選擇縣市鄉鎮別！');
            event.preventDefault();
        }
    });
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
                        return true;
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
                        return true;
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