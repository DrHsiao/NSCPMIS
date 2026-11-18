$(document).ready(function () {
    $('#Year').change(function () { ChangeCustomer(); });

    //若沒有選擇縣市別時顯示提示訊息
    $('.my-text button[type="submit"]').on('click', function (event) {
        var chkCityList = $('input[type="checkbox"][name="CityList"]:checked').length;
        if (chkCityList === 0) {
            alert('請選擇縣市別！');
            event.preventDefault();
        }
    });
});

function ChangeCustomer() {
    var selectedValue = $('#Year option:selected').val();
    if ($.trim(selectedValue).length > 0) {
        GetOrders(selectedValue);
    }
}

function GetOrders(Year) {
    $.ajax({
        url: orderUrl,
        data: { selectYear: Year },
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

$(function () {
    $("#Value").change(function () {
        var selectedValue = $(this).val();
        $.ajax({
            url: pjYearUrl,
            type: 'post',
            cache: false,
            async: false,
            data: { selectedPID: selectedValue },
            dataType: 'json',
            success: function (data1) {
                var data = JSON.parse(data1);
                if (data.length > 0) {
                    $('#PJYear').empty();
                    $('#PJYear').append($('<div></div>').val('').text('計畫年度'));
                    $.each(data, function (i, item) {
                        $('#PJYear').append($('<label></label>').append($('<input>', {
                            type: 'checkbox',
                            name: 'PJYear',
                            value: item.SYR
                        })).append(item.SYR));
                    });
                }
            }
        });
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

$(function () {
    $('#City').multipleSelect({
        placeholder: "縣市(可複選)",
        width: '18%',
        selectAll: true,
        multiple: true,
        multipleWidth: 150,
        onClick: function (v) {
            console.log("item checked");
            console.log(v.value);
            console.log(v);
        }
    });
});

function checkAll1(area) {
    var checkboxes = document.getElementsByClassName('city-checkbox');
    for (var i = 0; i < checkboxes.length; i++) {
        var checkbox = checkboxes[i];
        if (checkbox.id == area) {
            checkbox.checked = document.getElementById('Site1').checked;
        }
    }
}

function checkAll2(area) {
    var checkboxes = document.getElementsByClassName('city-checkbox');
    for (var i = 0; i < checkboxes.length; i++) {
        var checkbox = checkboxes[i];
        if (checkbox.id == area) {
            checkbox.checked = document.getElementById('Site2').checked;
        }
    }
}

function checkAll3(area) {
    var checkboxes = document.getElementsByClassName('city-checkbox');
    for (var i = 0; i < checkboxes.length; i++) {
        var checkbox = checkboxes[i];
        if (checkbox.id == area) {
            checkbox.checked = document.getElementById('Site3').checked;
        }
    }
}