//鄉鎮對應首頁縣市
function GetTown(city) {
    //alert(city)
    $(".TOWN").html("");
    $.ajax({
        url: '@Url.Content("~/")ASYNC/GetTowns',
        method: 'post',
        data: { "city": city },
        success: function (d) {
            var p = jQuery.parseJSON(d);
            $(".TOWN").prepend("<option value=''>鄉鎮</option>");
            for (i = 0; i < p.length; i++) {
                $(".TOWN").append("<option value='" + p[i].TOWN + "'>" + p[i].TOWN + "</option>");
            }
        }
    });
}

function cp() {
    $("#pe").append($("#cp").html());
}

function cpFund() {
    $("#fund_table_target").append($("#fund_table").html());
}

function cpFile() {
    $("#file_table_target").append($("#file_table").html());
}

function getRowIndex(button, tbName) {
    const row = button.closest("tr"); // 找到按鈕所在的行
    const table = document.getElementById(tbName);
    const rowIndex = Array.from(table.querySelector("tbody").rows).indexOf(row); // 計算相對索引
    //console.log("Row Index:", rowIndex);
    return rowIndex;
}

function moveDown(button, tbName) {
    const rowIndex = getRowIndex(button, tbName);
    const table = document.getElementById(tbName);
    const tbody = table.querySelector("tbody");
    const rows = tbody.querySelectorAll("tr");
    if (rowIndex > 0 && rowIndex < rows.length) {
        const currentRow = rows[rowIndex];
        const previousRow = rows[rowIndex - 1];
        tbody.insertBefore(currentRow, previousRow);
    }
}

function moveUp(button, tbName) {
    const rowIndex = getRowIndex(button, tbName);
    const table = document.getElementById(tbName);
    const tbody = table.querySelector("tbody");
    const rows = tbody.querySelectorAll("tr");
    if (rowIndex >= 0 && rowIndex < rows.length - 1) {
        const currentRow = rows[rowIndex];
        const nextRow = rows[rowIndex + 1];
        tbody.insertBefore(nextRow, currentRow);
    }
}

function duplicateRow(button, tbName) {
    const rowIndex = getRowIndex(button, tbName);
    const table = document.getElementById(tbName);
    const tbody = table.querySelector("tbody");
    const rows = tbody.querySelectorAll("tr");

    if (rowIndex >= 0 && rowIndex < rows.length) {
        const rowToCopy = rows[rowIndex]; // 取得要複製的行
        const newRow = rowToCopy.cloneNode(true); // 深度複製行及其內容
        tbody.appendChild(newRow); // 將複製的行新增到 tbody 的最後
    } else {
        console.error("Invalid row index");
    }
}

function duplicateTOWNRow(button, tbName) {
    const rowIndex = getRowIndex(button, tbName);
    const table = document.getElementById(tbName);
    const tbody = table.querySelector("tbody");
    const rows = tbody.querySelectorAll("tr");

    if (rowIndex >= 0 && rowIndex < rows.length) {
        const rowToCopy = rows[rowIndex]; // 取得要複製的行
        const newRow = rowToCopy.cloneNode(true); // 深度複製行及其內容
        tbody.appendChild(newRow); // 將複製的行新增到 tbody 的最後
    } else {
        console.error("Invalid row index");
    }

    var lis = document.getElementsByClassName('TOWN');//目標
    var lisLen = lis.length;
    for (var i = 0; i < lisLen; i++) {
        //var存在变量提升, lis[i].index = i 为事先存储元素的索引值。若不写，打印的元素索引值会一直是元素的总长度。
        lis[i].index = i;//想省略这句，在for循环中，可用es6中的let替换var
        lis[i].onclick = function () {
            console.log(this.index);//元素索引值
        }
    }
}

function del(button, tbName) {
    const rowIndex = getRowIndex(button, tbName);
    const table = document.getElementById(tbName);
    const tbody = table.querySelector("tbody");
    const rows = tbody.querySelectorAll("tr");

    if (rowIndex >= 0 && rowIndex < rows.length) {
        const rowToRemove = rows[rowIndex];
        removeRow(rowToRemove); // 使用移除行的函式
    } else {
        console.error("Invalid row index");
    }


}

function removeRow(row) {
    if (row && row.parentNode) {
        row.parentNode.removeChild(row); // 從父節點中移除該行
    } else {
        console.error("Invalid row or parentNode not found");
    }
}

function CpTo3() {
    //$("#table3_main_body").html("");

    //$("#tab3").html($("#tab2")());
    /*var towns = $(".TOWN");

    var i = 0;

    towns.each(function () {
        var town = $(this).val();
        var po = $(".POS_SYB").eq(i);
        if (town.length > 0) {
            $("#table3_main_body").append(GenHtml(i, town, po.val()));
            i++;
        }



    })*/
}