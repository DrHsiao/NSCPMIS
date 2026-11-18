// 動態取得nav區塊的高度，防止頁面的文字被nav區塊檔住
var NavAdjust = (function ($) {

    var FIXED_MARGIN_TOP = 40; // page-header 的固定 margin-top

    function adjustNextContent() {
        var $nav = $('nav');
        var $next = $('.page-header').parent().parent();// $nav.next();
        if ($nav.length === 0 || $next.length === 0) return;

        var navbarHeight = $nav.outerHeight(true);
        $next.css('margin-top', navbarHeight - FIXED_MARGIN_TOP);
    }

    function debounce(func, wait) {
        var timeout;
        return function () {
            clearTimeout(timeout);
            timeout = setTimeout(func, wait);
        };
    }

    function init() {
        // 頁面載入時計算一次
        $(document).ready(adjustNextContent);

        // 視窗 resize 時計算一次(debounce)
        $(window).resize(debounce(adjustNextContent, 100));
    }

    return {
        init: init
    };

})(jQuery);