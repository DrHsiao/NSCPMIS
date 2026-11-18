const ReplyEditor = {
    init() {
        this.bindEvents();
    },

    bindEvents() {
        $(document).on('click', '.btn-edit', this.handleEditClick);
        $(document).on('click', '.btn-cancel', this.handleCancelClick);
        $(document).on('click', '.btn-confirm', this.handleConfirmClick);
        $(document).on('click', '.btn-delete', this.handleDeleteClick);
    },

    // 回覆功能：綁定「編輯」按鈕
    handleEditClick(e) {
        const id = $(e.currentTarget).data('id');
        ReplyEditor.enterEditMode(id);
    },

    // 回覆功能：綁定「取消編輯」按鈕
    handleCancelClick(e) {
        const id = $(e.currentTarget).data('id');
        ReplyEditor.cancelEdit(id);
    },

    // 回覆功能：綁定「確認修改」按鈕
    handleConfirmClick(e) {
        if (!confirm("確定修改這則回覆嗎？")) { return; }

        const $btn = $(e.currentTarget);
        const id = $btn.data('id');
        const content = $('#textarea-' + id).val();

        if (content.trim() === "") {
            alert("請輸入回覆內容");
            return;
        }

        $btn.prop('disabled', true);

        $.ajax({
            url: editUrl,
            type: 'POST',
            data: { id, content },
            success: function () {
                alert("修改成功");
                $('#content-' + id).text(content).show();
                $('#textarea-' + id).hide();
                ReplyEditor.toggleButtons(id, false);
            },
            error: function (xhr, status, error) {
                console.error("AJAX 錯誤：", error);
                alert("修改失敗");
            },
            complete: function (xhr, status) {
                $btn.prop('disabled', false);
            }
        });
    },

    // 回覆功能：綁定「刪除」按鈕
    handleDeleteClick(e) {
        if (!confirm("確定刪除這則回覆嗎？")) { return; }

        const $btn = $(e.currentTarget);
        const id = $btn.data('id');
        $btn.prop('disabled', true);

        $.ajax({
            url: delUrl,
            type: 'POST',
            data: { id },
            success: function () {
                alert("刪除成功");
                $('.reply-box[data-id="' + id + '"]').remove();
            },
            error: function (xhr, status, error) {
                console.error("AJAX 錯誤：", error);
                alert("刪除失敗");
            },
            complete: function (xhr, status) {
                $btn.prop('disabled', false);
            }
        });
    },

    enterEditMode: function (id) {
        $('#content-' + id).hide();
        $('#textarea-' + id).show();
        this.toggleButtons(id, true);
    },

    cancelEdit: function (id) {
        $('#textarea-' + id).hide();
        $('#content-' + id).show();
        this.toggleButtons(id, false);
    },

    // 回覆功能：切換編輯模式
    toggleButtons: function (id, editing) {
        $('.btn-edit[data-id="' + id + '"]').toggle(!editing);
        $('.btn-delete[data-id="' + id + '"]').toggle(!editing);
        $('.btn-confirm[data-id="' + id + '"]').toggle(editing);
        $('.btn-cancel[data-id="' + id + '"]').toggle(editing);
    }
};

const FormValidator = {
    init() {
        this.bindEvents();
    },

    bindEvents: function () {
        // 表單提交驗證
        $('form').on('submit', function (e) {
            if (!confirm('已確實填寫內容，確定送出？')) {
                e.preventDefault();
                return;
            }

            // 鎖住按鈕
            const submitBtn = $('input[type="submit"]');
            if (submitBtn.length) {
                submitBtn.prop("disabled", true).val("儲存中...");
            }
        });
    }
};

$(function () {
    if (canEdit) {
        ReplyEditor.init();
        FormValidator.init();
    }
});