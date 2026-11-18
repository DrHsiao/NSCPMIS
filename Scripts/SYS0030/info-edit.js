(function () {
    function safeParse(json) {
        try {
            return json ? JSON.parse(json) : [];
        } catch {
            return [];
        }
    }

    function getVal(selector) {
        return $.trim($(selector).val() || '');
    }
    
    // === 資料設定 ===
    const el = document.getElementById('apiData');

    const api = {
        deleteUrl: el.dataset.deleteurl,
        deleteFileUrl: el.dataset.deletefileurl,
        mode: el.dataset.mode,
        apms: safeParse(el.dataset.apms),
        tdms: safeParse(el.dataset.tdms)
    };

    //  === 刪除按鈕 ===
    $(document).on('click', '.btn-delete', function () {
        handleDelete();
    });

    // === 刪除 ===
    function handleDelete() {

        if (!confirm('確定要刪除嗎？')) return;

        if (!checkRelatedData(api.tdms, '刪除', '工程登記冊')) return;

        const id = $('#PJ_ID').val();
        if (!id) return;

        $.ajax({
            type: 'POST',
            url: api.deleteUrl,
            data: { id, check: false },
            dataType: 'json',
            success: function (res) {
                alert(res.message);
                if (res.success) {
                    window.parent.location.reload();
                }
            },
            error: function (ex) {
                alert('刪除失敗，請稍後重試');
            }
        });
    }

    // === 檔案刪除 ===
    window.DeleteFile = function (fileId) {

        if (!confirm('確定要刪除此檔案嗎？')) return;

        $.ajax({
            type: 'POST',
            url: api.deleteFileUrl,
            data: { fileId },
            dataType: 'json',
            success: function (res) {
                alert(res.message);
                if (res.success) {
                    location.reload();
                }
            },
            error: function () {
                alert('刪除失敗，請稍後重試');
            }
        });
    };


    // === 關閉 ===
    window.CloseDialog = function () {
        window.parent.location.reload();
    };

    // === 修改確認 ===
    window.handleSubmit = function () {

        if (api.mode === 'copy') {
            return ValidateSave();
        }

        if (!ValidateSave()) return false;

        const oldName = $('.item-oname').val();
        const newName = $('#PJ_NAME').val();

        if (oldName === newName) {
            return true;
        }

        return checkRelatedData(api.apms, '修改', '先期作業提案');
    };
    
    function checkRelatedData(list, actionName, dataName) {

        if (!Array.isArray(list) || list.length === 0) {
            return true;
        }

        const message = list.join(",");

        return confirm(
            `下列${dataName}資料與本實施計畫有關聯，${actionName}後需重新綁定，確定要${actionName}嗎？\n\n${message}`
        );
    }

    // === 驗證資料 ===
    window.ValidateSave = function () {

        const rules = [
            // 1. 檢查造冊編碼 (檢查前三碼是否為空，視貴單位規則可自行調整檢查欄位)
            // 使用 $.trim() 去除空白，確保使用者不是只按空白鍵
            {
                valid: () =>
                    getVal("#SER_FN_1") &&
                    getVal("#SER_FN_2") &&
                    getVal("#SER_FN_3") &&
                    getVal("#SER_FN_4") &&
                    getVal("#SER_FN_5") &&
                    getVal("#SER_FN_6") &&
                    getVal("#SER_FN_7"),
                message: "[造冊編碼] 為必填欄位，請填寫完整。"
            },
            // 2. 檢查縣市 (下拉選單)
            {
                valid: () => getVal("#SER_CITY"),
                message: "[縣市] 為必填欄位，請選擇。"
            },
            // 3. 檢查雨水下水道系統 (下拉選單)
            // 注意：如果您的 ViewBag.systems 預設選項值不是空字串，請依照實際狀況調整
            {
                valid: () => getVal("#SYS_ID"),
                message: "[雨水下水道系統] 為必填欄位，請選擇。"
            },
            // 4. 檢查文件類別 (下拉選單)
            {
                valid: () => getVal("#PJ_ID_TYPE"),
                message: "[文件類別] 為必填欄位，請選擇。"
            }
        ];

        let errors = [];

        rules.forEach(r => {
            if (!r.valid()) {
                errors.push(r.message);
            }
        });

        if (errors.length > 0) {
            alert("資料驗證失敗，請檢查以下項目：\n\n" + errors.join("\n"));
            return false;
        }

        return true;
    };

})();