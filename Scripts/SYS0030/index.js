// === 資料設定 ===
const indexEl = document.getElementById('indexData');

const indexApi = {
    deleteUrl: indexEl.dataset.deleteurl
};

// === 刪除 ===
function Delete(id) {

    if (!confirm('確定要刪除嗎？')) return;
    
    if (!id) return;

    deleteRequest(id, true);
}
function deleteRequest(id, check) {

    $.ajax({
        type: 'POST',
        url: indexApi.deleteUrl,
        data: { id, check },
        dataType: 'json',
        success: function (res) {

            // 有關聯資料時需要二次確認
            if (res.isUnused === false && check === true) {

                if (checkRelatedData(res.usedIn.tdms, '刪除', '工程登記冊')) {
                    return deleteRequest(id, false);
                }

                return;
            }

            alert(res.message);

            if (res.success) {
                window.parent.location.reload();
            }
        },
        error: function () {
            alert('刪除失敗，請稍後重試');
        }
    });
}

function checkRelatedData(list, actionName, dataName) {

    if (!Array.isArray(list) || list.length === 0) {
        return true;
    }

    const message = list.join(",");

    return confirm(
        `下列${dataName}資料與本實施計畫有關聯，${actionName}後需重新綁定，確定要${actionName}嗎？\n\n${message}`
    );
}