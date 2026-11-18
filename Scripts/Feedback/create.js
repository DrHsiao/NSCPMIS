let index = 0; // 檔案上傳列的索引

const FileUploadManager = {
    init: function () {
        this.cacheElements();
        this.bindEvents();
        this.addFileUploadRow(); // 預設一筆
    },

    cacheElements: function () {
        this.$tbody = $('#uploadTableBody');
        this.$btnAdd = $('#btn-add');
        this.$form = $('form');
    },

    bindEvents: function () {
        // 新增檔案列按鈕
        this.$btnAdd.on('click', () => {
            this.addFileUploadRow();
        });

        // 委派監聽刪除按鈕
        $(document).on('click', '.btn-delete', (e) => {
            this.removeFileUploadRow($(e.currentTarget));
        });

        // 委派監聽檔案欄位變更事件驗證
        $(document).on('change', '.file-input', (e) => {
            this.validateFile(e.currentTarget);
        });
    },

    addFileUploadRow: function () {
        const newRow = '<tr>' +
            '<td>' +
            '<input type="text" name="FileUploads[' + index + '].FILE_DESC" id="FileUploads[' + index + '].FILE_DESC" ' +
            'class="form-control" maxlength="100" autocomplete="off" placeholder="檔案說明" />' +
            '</td>' +
            '<td>' +
            '<input type="file" name="FileUploads[' + index + '].File" ' +
            'class="form-control file-input" accept="' + acceptFileString + '" />' +
            '<span name="fileError" class="text-danger"></span>' +
            '</td>' +
            '<td>' +
            '<button type="button" class="btn btn-sm btn-danger btn-delete">移除</button>' +
            '</td>' +
            '</tr>';
        this.$tbody.append(newRow);
        index++;
    },

    removeFileUploadRow: function ($btn) {
        if (confirm('確定要移除這個檔案上傳區塊嗎？')) {
            $btn.closest('tr').remove();
        }
    },

    validateFile: function (input) {
        const file = input.files[0];
        const $input = $(input);
        const $errorSpan = $input.next('span[name="fileError"]');

        $errorSpan.text('');
        input.setCustomValidity('');

        if (!file) return;

        const result = validateSingleFile(file, {
            allowedTypes: allowedExtensions,
            allowedMimes: ['image/jpeg', 'image/png']
        });

        if (!result.valid) {
            $errorSpan.text(result.reason);
            input.setCustomValidity(result.reason);
        }
    }
};

const FormValidator = {
    init() {
        this.cacheElements();
        this.bindEvents();
    },

    cacheElements: function () {
        this.$form = $('form');
    },

    bindEvents: function () {
        // 表單提交驗證
        this.$form.on('submit', (e) => {
            this.handleFormSubmit(e);
        });
    },

    handleFormSubmit: function (e) {
        if (!confirm('已確實填寫內容，確定送出？')) {
            e.preventDefault();
            return;
        }

        this.$form.find('[name="fileError"]').text('');

        if (!validateAllFiles()) {
            alert('上傳的檔案總大小不可超過 ' + DEFAULT_TOTAL_FILE_SIZE_MB + 'MB');
            e.preventDefault();
        }

        // 鎖住按鈕
        const submitBtn = $('input[type="submit"]');
        if (submitBtn.length) {
            submitBtn.prop("disabled", true).val("儲存中...");
        }
    }
};

$(function () {
    FileUploadManager.init();
    FormValidator.init();
});