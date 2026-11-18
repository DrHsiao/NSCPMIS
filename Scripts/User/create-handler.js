function toggleCheckboxes() {
    var checkboxes = document.querySelectorAll('.city-checkbox');
    var selectAllCheckbox = document.getElementById('selectAll');

    for (var i = 0; i < checkboxes.length; i++) {
        checkboxes[i].checked = selectAllCheckbox.checked;
    }
}

function Hello(id) {
    var City = id + 'i';
    $('.city2-checkbox').closest('td').hide();
    $('#' + City).attr('colspan', '2');
    $('#' + City).show();

    //City2Element.attr('colspan', '2');
    //City2Element.show();
}

$(function () {
    var selectedCity; // 儲存被選中的城市

    $('.city-checkbox').change(function () {
        var checkboxChecked = $(this).is(':checked');
        selectedCity = $(this).attr('name');

        if (checkboxChecked) {
            selectedCity = $(this).attr('name');
            // 隱藏所有城市的新增內容
            $('.city2-checkbox').closest('td').hide();

            $.ajax({
                url: cityUrl,
                type: 'post',
                cache: false,
                async: false,
                data: { City: selectedCity },
                success: function (result1) {
                    var result = JSON.parse(result1);
                    if (result.length > 0) {
                        var City2 = selectedCity + 'i'; // 建立 City2 的 ID
                        var City2Element = $('#' + City2);

                        if (City2Element.length === 0) {
                            $('tr#Second').append('<td colspan="2" id="' + City2 + '"></td>');
                            City2Element = $('#' + City2);

                            //City2Element.append($('<label></label>').append($('<input>', {
                            //    type: 'checkbox',
                            //    name: 'SelectedboxTOWN',
                            //    value: selectedCity,
                            //    checked: true,
                            //    class: 'city2-checkbox'

                            //})).append('&nbsp;').append(selectedCity).append('&nbsp;').css('display', 'none'));



                            $.each(result, function (i, item) {
                                if (i !== 0 && i % 7 === 0) {
                                    City2Element.append('<br />');
                                }

                                City2Element.append($('<label></label>').append($('<input>', {
                                    type: 'checkbox',
                                    name: 'SelectedboxTOWN',
                                    value: selectedCity + item.TOWN,
                                    checked: true,
                                    class: 'city2-checkbox'
                                })).append('&nbsp;').append(item.TOWN).append('&nbsp;'));
                            });

                            City2Element.append($('<label></label>').css('display', 'none').append($('<input>', {
                                type: 'checkbox',
                                name: 'SelectedboxTOWN',
                                value: selectedCity,
                                checked: true,
                                class: 'city2-checkbox'
                            })).append('&nbsp;').append('').append('&nbsp;'));   /////////////////這是新增縣市鄉鎮的空值，如果不需要一定要拿掉
                        } else {
                            City2Element.attr('colspan', '2');
                            City2Element.show();
                        }
                    }
                }
            });
        } else {
            // 取消選擇時，隱藏相關的新增內容
            $('#' + selectedCity + 'i').remove();
        }
    });
});

//DeptCreate用
var selectElement = document.getElementById("dropdownlist");
if (selectElement) {
    selectElement.addEventListener("change", function () {
        var currentURL = window.location.href;
        var url = new URL(currentURL);
        url.searchParams.set("Dept", this.value);
        // 跳轉到新的 URL
        window.location.href = url.href;
    });
}
