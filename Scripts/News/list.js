$(function () {
    $(document).on('click', '.btn-delete', function () {
        const $btn = $(this);
        const id = $btn.data("id");
        const status = $btn.data("status");

        if (!id) {
            alert("刪除失敗！");
            return;
        }

        const confirmMessage = (status === "T")
            ? "此最新消息正在公告中，確定要刪除嗎？"
            : "確定刪除？";

        if (!confirm(confirmMessage)) return;

        $btn.prop('disabled', true); // 避免重複點擊

        $.ajax({
            url: delNewsUrl,
            method: 'POST',
            data: { id },
            success(result) {
                if (result.success) {
                    $("#row-" + id).remove();
                    alert("刪除成功！");
                } else {
                    alert("刪除失敗：" + (result.message || "未知錯誤"));
                    $btn.prop('disabled', false); // 失敗時要還原按鈕狀態
                }
            },
            error(xhr, status, error) {
                console.error("AJAX 錯誤：", error);
                alert("刪除時發生錯誤，請稍後再試。");
                $btn.prop('disabled', false); // 失敗時要還原按鈕狀態
            }
        });
    });
});