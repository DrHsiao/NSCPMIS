const NewsModule = (function () {
    // DOM 元素快取
    const $type = $('#TYPE');
    const $newsFiles = $('#newsFiles');
    const $body = $('#BODY');

    // ===== 私有函式 =====

    // 根據選擇的訊息類別顯示或隱藏相關檔案上傳區塊
    function toggleFileUpload() {
        $type.val() === '02' ? $newsFiles.show() : $newsFiles.hide();
    }

    // 綁定訊息類別變更事件
    function bindTypeChangeHandler() {
        $type.off('change').on('change', toggleFileUpload);
        toggleFileUpload();
    }

    // 初始化 Summernote
    function initSummernote() {
        $body.summernote({
            toolbar: [
                ['style', ['bold', 'italic', 'underline', 'strikethrough']],
                //['fontsize', ['fontsize']], // 字體大小選擇器(目前會被AntiXSS檔掉，先關閉)
                ['color', ['forecolor']], // 只有文字顏色
                //['color', ['forecolor', 'backcolor']],  // 同時有文字顏色和背景色(目前會被AntiXSS檔掉，先關閉)
                ['para', ['ul', 'ol', 'paragraph']],
                ['link']
            ],
            height: 300,
            placeholder: '請輸入您的內容...'
        });
    }

    // 禁用輸入欄位（排除 .always-enabled）
    function disableInputs() {
        $('form :input').not('.always-enabled').prop('disabled', true);
        $body.summernote('disable');
    }

    // 更新送出按鈕文字
    function updateSubmitButtonText() {
        const id = parseInt($('#NEWSID').val(), 10) || 0;
        const submitBtn = $('input[type="submit"]');
        if (submitBtn.length) {
            submitBtn.val(id > 0 ? '更新公告' : '發佈公告');
        }
    }

    // 綁定刪除檔案按鈕的點擊事件
    function bindDeleteButtonClick() {
        $(document).on('click', '.btn-delete', function (e) {
            e.preventDefault();

            const fid = $(this).data('id');
            if (!fid) return alert('刪除失敗！');
            if (!confirm('確定要刪除嗎？刪除後需重新上傳。')) return;

            const $delInput = $('#DEL_FG_' + fid);
            if ($delInput.length) $delInput.val('*');

            $.ajax({
                url: delNewsFileUrl,
                method: 'POST',
                data: {
                    id: $('#NEWSID').val(),
                    fid: fid
                },
                success: function (result) {
                    if (result.success) {
                        $("#row-" + fid).remove();
                        alert('刪除成功！');
                    } else {
                        alert('刪除失敗：' + (result.message || '未知錯誤'));
                    }
                },
                error: function (xhr, status, error) {
                    console.error('AJAX 錯誤：', error);
                    alert('刪除時發生錯誤，請稍後再試。');
                }
            });
        });
    }

    // 綁定儲存按鈕的點擊事件
    function validateFormSubmit() {
        $('form').on('submit', function (e) {
            if (!confirm('已確實填寫內容，確定送出？')) {
                e.preventDefault();
                return;
            }

            let isValid = true;
            const bodyError = $('[data-valmsg-for="BODY"]');
            const fileError = $('[name="fileError"]');

            // 清空錯誤訊息
            bodyError.text('');
            fileError.text('');

            const content = $body.summernote('isEmpty') ? '' : $body.summernote('code');
            const textContent = $('<div>').html(content).text().trim();
            const bodyMaxLength = 1000;

            if (textContent.length === 0) {
                bodyError.text('內容欄位是必要的。');
                isValid = false;
            } else if (textContent.length > bodyMaxLength) {
                bodyError.text('內容最多只能輸入 ' + bodyMaxLength + ' 個字，目前已輸入 ' + textContent.length + ' 個字。');
                isValid = false;
            }

            if (!validateAllFiles()) {
                fileError.text('上傳的檔案總大小不可超過 ' + DEFAULT_TOTAL_FILE_SIZE_MB + 'MB');
                isValid = false;
            }

            if (!isValid) {
                e.preventDefault();
            } else {
                $body.val(content);

                // 鎖住按鈕
                const submitBtn = $('input[type="submit"]');
                if (submitBtn.length) {
                    submitBtn.prop("disabled", true).val("儲存中...");
                }
            }
        });
    }

    // 當 StartDate 的值改變時，更新 EndDate 的最小日期
    function bindStartDateSync() {
        const startDateInput = document.getElementById('STARTDATE');
        const endDateInput = document.getElementById('ENDDATE');

        if (!startDateInput || !endDateInput) return;

        startDateInput.addEventListener('change', () => {
            const start = startDateInput.value;
            const end = endDateInput.value;

            endDateInput.min = start;
            if (end && end < start) endDateInput.value = '';
        });
    }
    
    // 新增檔案上傳列
    function addFileUploadRow() {
        const newRow = $("#file_table").html();
        $("#file_table_target").append(newRow);
    }

    // 驗證單一檔案
    function validateFile(input) {
        const errorSpan = input.nextElementSibling;
        errorSpan.textContent = '';
        input.setCustomValidity('');

        const file = input.files[0];
        if (!file) return;

        const result = validateSingleFile(file, {
            allowedTypes: allowedExtensions,
            allowedMimes: [
                'image/jpeg',
                'image/png',
                'application/pdf',
                'application/msword',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'application/vnd.ms-excel',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'application/x-zip-compressed'
            ],
            maxSizeMB: 20,
            maxTotalMB: 20
        });

        if (!result.valid) {
            errorSpan.textContent = result.reason;
            input.setCustomValidity(result.reason);
        }
    }

    // 初始化
    function init() {
        bindTypeChangeHandler();
        initSummernote();

        if (isViewMode) {
            disableInputs();
            return;
        }

        updateSubmitButtonText();
        bindDeleteButtonClick();
        validateFormSubmit();
        bindStartDateSync();
    }
    
    return {
        init,
        addFileUploadRow,
        validateFile
    };
})();

$(document).ready(function () {
    NewsModule.init();
});