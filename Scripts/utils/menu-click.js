// 儲存系統操作紀錄(使用者點擊功能頁面紀錄)
var MenuClick = (function ($) {

    var lastClickTime = {};
    var CLICK_INTERVAL = 5000; // 5 秒限制

    function init(options) {
        options = options || {};
        var postUrl = options.postUrl; // 後端 POST 路徑

        // 事件委派監聽 menu 點擊
        $(document).on('click', 'ul.dropdown-menu ul li a', function (e) {
            e.preventDefault();

            var $link = $(this);
            var url = $link.attr('href');
            var menuText = $link.text().trim();
            var now = Date.now();

            // 如果同一個 menu 5 秒內已經點過，就不再 POST
            if (lastClickTime[menuText] && now - lastClickTime[menuText] < CLICK_INTERVAL) {
                window.location.href = url;
                return;
            }

            lastClickTime[menuText] = now;
            
            $.post(postUrl, { path_name: menuText })
                .always(function () {
                    // 完成後跳轉頁面
                    window.location.href = url;
                });
        });
    }

    return {
        init: init
    };

})(jQuery);