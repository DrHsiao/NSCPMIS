(function () {

    // === API URL 統一管理 ===
    const apiEl = document.getElementById('apiUrls');
    const api = {
        check: apiEl.dataset.check,
        modify: apiEl.dataset.modify,
        delete: apiEl.dataset.delete
    };

    // === 事件綁定（支援動態 DOM） ===
    $(document).on('click', '.btn-update', function () {
        const row = getRow($(this));
        const id = getId(row);
        checkItem(1, id, row);
    });

    $(document).on('click', '.btn-delete', function () {
        if (!confirm('確定要刪除嗎？')) return;

        const row = getRow($(this));
        const id = getId(row);
        checkItem(2, id, row);
    });

    $(document).on('click', '.btn-rupdate', function () {
        const row = getRow($(this));
        const id = getId(row);
        checkItem(3, id, row);
    });

    // === 共用方法 ===
    function getRow($el) {
        return $el.closest('tr');
    }

    function getId(row) {
        return row.find('.item-id').val();
    }

    // === 檢查是否被使用 ===
    function checkItem(type, id, row) {
        $.post(api.check, { type, id })
            .done(res => {
                if (!res.success) {
                    alert(res.message);
                    return;
                }

                if (type === 3) { // 雨水科
                    var oname = row.find('.item-oname').val();
                    var name = row.find('.item-name').val();
                    if (oname === name) {
                        // 名稱未修改，直接執行動作
                        executeAction(type, id, row);
                        return;
                    }
                }

                if (res.isUnused) {
                    executeAction(type, id, row);
                } else {
                    const html = buildMessage(type, res.usedIn);
                    showDialog(type, html, id, row);
                }
            })
            .fail(() => alert('檢查失敗'));
    }

    // === 根據類型執行動作 ===
    function executeAction(type, id, row) {
        if (type === 1) {
            doModify(id, row);
        } else if (type === 2) {
            doDelete(id, row);
        } else if (type === 3) {
            doModify(id, row);
        }
    }

    // === 組提示訊息 ===
    function buildMessage(type, usedIn) {
        let html = '';

        if (type === 2) {
            html += '刪除後下列資料將無法對應原始資料，需重新綁定，確定要刪除嗎？';
        } else {
            html += '修改後下列資料將無法對應原始資料，需重新綁定，確定要修改嗎？';
        }

        html += '<br><hr>';
        if (usedIn.apms?.length) {
            html += `※先期作業提案： ${usedIn.apms.join(', ')}<br>`;
        }
        if (type === 2) {
            if (usedIn.pjcs?.length) {
                html += `※規劃報告/實施計畫： ${usedIn.pjcs.join(', ')}<br>`;
            }
            if (usedIn.tdms?.length) {
                html += `※工程登記冊： ${usedIn.tdms.join(', ')}<br>`;
            }
        }

        return html;
    }

    // === 修改 ===
    function doModify(id, row) {
        const data = {
            __RequestVerificationToken: $('input[name="__RequestVerificationToken"]').val(),
            id,
            name: row.find('.item-name').val(),
            tp: row.find('.item-tp').val(),
            sup: row.find('.item-sup').val()
        };
        
        if (data.name === "") {
            alert('請輸入下水道系統名稱');
            return;
        }

        $.post(api.modify, data)
            .done(res => {
                if (!res.success) {
                    alert('修改失敗：' + res.message);
                    return;
                }
                alert('修改成功');
            })
            .fail(() => alert('發生不明錯誤'));
    }

    // === 刪除 ===
    function doDelete(id, row) {
        $.post(api.delete, { id, __RequestVerificationToken: $('input[name="__RequestVerificationToken"]').val() })
            .done(res => {
                if (!res.success) {
                    alert('刪除失敗：' + res.message);
                    return;
                }

                alert('刪除成功');
                row.remove();
            })
            .fail(() => alert('發生不明錯誤'));
    }

    // === Dialog ===
    function showDialog(type, messageHtml, id, row) {
        const dlg = document.getElementById('webdialog');

        const titleMap = {
            1: '修改確認',
            2: '刪除確認',
            3: '修改確認'
        };

        $('#dialogTitle').text(titleMap[type] || '確認');
        $('#modalContent').html(messageHtml);

        $('#cancelBtn').off('click').on('click', () => dlg.close());

        $('#confirmBtn').off('click').on('click', () => {
            executeAction(type, id, row);
            dlg.close();
        });

        dlg.showModal();
    }

})();