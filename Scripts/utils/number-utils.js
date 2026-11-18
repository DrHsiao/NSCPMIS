// 畫面輸入即時格式化
document.querySelectorAll('.fund-input').forEach(input => {
    input.addEventListener('input', function (e) {
        let value = this.value.replace(/,/g, '');
        if (!isNaN(value) && value.length > 0) {
            this.value = Number(value).toLocaleString('en-US');
        } else {
            this.value = '';
        }
    });
});

// form submit 前清掉千分位逗號
const form = document.querySelector('form');
form.addEventListener('submit', function () {
    document.querySelectorAll('.fund-input').forEach(input => {
        input.value = input.value.replace(/,/g, '');
    });
});