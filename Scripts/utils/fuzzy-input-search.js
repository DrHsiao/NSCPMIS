var fuzzySrchDivClass = '.fuzzy-srch-wrapper';
var fuzzySrchTxtClass = '.fuzzy-srch-txt';
var fuzzySrchDdlClass = '.fuzzy-srch-ddl';

const fuzzySrchDivs = document.querySelectorAll(fuzzySrchDivClass);
const fuzzySrchBoxes = document.querySelectorAll(fuzzySrchTxtClass);
const fuzzyDropdowns = document.querySelectorAll(fuzzySrchDdlClass);

// 點擊搜尋框時顯示對應的下拉選單資料
function txtFuzzyFilter(input) {
    const wrapper = input.closest(fuzzySrchDivClass);  // 取得當前的區塊
    const dropdown = wrapper.querySelector(fuzzySrchDdlClass);  // 取得對應的下拉選單
    const items = dropdown.getElementsByTagName('li');
    const filter = input.value.toLowerCase();

    // ==========================================
    // 💡 [升級新增]：當使用者自己手動打字修改時，清空隱藏的代碼
    // 代表他目前處於「未選擇清單項目」的狀態，避免送出錯誤的舊代碼
    // ==========================================
    const hiddenInput = wrapper.querySelector('input[type="hidden"]');
    if (hiddenInput) {
        hiddenInput.value = '';
    }

    let hasResults = false;
    dropdown.style.display = 'block'; // 顯示下拉選單

    for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const textValue = item.textContent || item.innerText;

        if (textValue.toLowerCase().indexOf(filter) > -1) {
            item.style.display = ''; // 顯示符合的選項
            hasResults = true;
        } else {
            item.style.display = 'none'; // 隱藏不符合的選項
        }
    }

    if (!hasResults) {
        dropdown.style.display = 'none'; // 如果沒有匹配結果則隱藏下拉選單
    }
}

// 點擊下拉選單某一選項時，將資料填入搜尋框內
function filterSelectItem(item) {
    const wrapper = item.closest(fuzzySrchDivClass);
    const input = wrapper.querySelector(fuzzySrchTxtClass); // 取得對應的搜尋框

    // 1. 將名稱 (Text) 填入顯示的搜尋框內
    input.value = item.textContent || item.innerText;

    // ==========================================
    // 💡 [升級新增]：尋找是否有隱藏的 input，如果有，就把 data-val 塞進去
    // ==========================================
    const hiddenInput = wrapper.querySelector('input[type="hidden"]');
    if (hiddenInput) {
        // 抓取 <li> 上的 data-val 屬性，如果沒有則給空字串
        hiddenInput.value = item.getAttribute('data-val') || '';
    }

    // 3. 隱藏下拉選單
    wrapper.querySelector(fuzzySrchDdlClass).style.display = 'none';
}

// 綁定點擊下拉選項事件
fuzzyDropdowns.forEach(dropdown => {
    const liItems = dropdown.querySelectorAll('li');
    liItems.forEach(li => {
        // 使用 mousedown 體驗會比 click 更好，能防止 input 失去焦點時的衝突
        li.addEventListener('mousedown', function (event) {
            filterSelectItem(li);
        });
    });
});


// 點擊搜尋框時，隱藏其他模糊搜尋區塊裡的下拉選單
fuzzySrchBoxes.forEach(function (input) {
    input.addEventListener('click', function (event) {
        // 點擊當前搜尋框時，隱藏其他搜尋框的下拉選單
        fuzzySrchDivs.forEach(function (wrapper) {
            const dropdown = wrapper.querySelector(fuzzySrchDdlClass);
            if (wrapper !== input.closest(fuzzySrchDivClass)) {
                dropdown.style.display = 'none'; // 隱藏其他搜尋框的下拉選單
            }
        });

        // 讓當前下拉選單保持顯示
        const currentDropdown = input.closest(fuzzySrchDivClass).querySelector(fuzzySrchDdlClass);
        currentDropdown.style.display = 'block';

        // 防止點擊輸入框時，觸發其他隱藏事件
        event.stopPropagation();
    });
});

// 點擊頁面任意地方時，隱藏模糊搜尋區塊的下拉選單
document.addEventListener('click', function (event) {
    fuzzySrchDivs.forEach(function (wrapper) {
        const dropdown = wrapper.querySelector(fuzzySrchDdlClass);

        // 如果點擊的地方不是搜尋框或下拉選單，則隱藏下拉選單
        if (!wrapper.contains(event.target)) {
            dropdown.style.display = 'none';
        }
    });
});

// 防止在搜尋框或下拉選單內點擊時觸發隱藏操作
fuzzySrchDivs.forEach(function (wrapper) {
    wrapper.addEventListener('click', function (event) {
        event.stopPropagation(); // 阻止事件傳播
    });
});

// 按下Tab鍵時，隱藏所有下拉選單
document.addEventListener('keydown', (event) => {
    if (event.key === 'Tab') {
        fuzzyDropdowns.forEach(dropdown => dropdown.style.display = 'none');
    }
});