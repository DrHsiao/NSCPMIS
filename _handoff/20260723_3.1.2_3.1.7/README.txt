3.1.2 / 3.1.7 功能測試修改 — 交付快照 (2026-07-23，未 commit)
===============================================================

【內容】此資料夾為 untracked（不進 git、不受 VS 還原影響）的備份快照。

變更清單（原始位置 → 影響功能）：
  1. classes\john\util.cs                    3.1.7 新增 GetCitiesByDept（依部門查 SECCITY 縣市）
  2. Controllers\NSPMIS\SYS0030Controller.cs 3.1.7 縣市依部門/系統別連動端點/系統別管理每頁20/新增文件必填後端防呆
  3. Views\SYS0030\SYS0030Index.cshtml       3.1.7 主頁縣市→系統別連動
  4. Views\SYS0030\SYSInfo.cshtml            3.1.7 新增文件三欄必填(縣市/污水下水道系統/文件類別)+連動
  5. Views\SYS0030\SYSManager.cshtml         3.1.7 移除殘留 GetTowns onchange
  6. Views\TDE0020\TDE0020IndexTab1.cshtml   3.1.2 每列「一次儲存」按鈕(前後各一，觸發整列存檔，排除引用前月)
  6b.Controllers\NSPMIS\TDE0020Controller.cs 3.1.2 工程名稱查詢修復(2026-07-24)：
        (1)td_name 由前綴比對 td_name+"%" 改為包含比對 "%"+td_name+"%"
        (2)TAB1 效能修復：原本對查詢全部結果逐列各發一次 GenIncSql(N+1)，
           名稱查常見字命中數千筆→逾時查無結果；改為先去重分頁、只對當前頁列 enrich。
        (已編進 WCals.dll)

編譯產出：
  bin\WCals.dll / WCals.pdb  （已含上述 util.cs + SYS0030Controller.cs 的變更，Debug 編譯）

【部署到對方測試站的最小檔案集】（覆蓋既有站台即可）：
  - bin\WCals.dll
  - Views\SYS0030\SYS0030Index.cshtml
  - Views\SYS0030\SYSInfo.cshtml
  - Views\SYS0030\SYSManager.cshtml
  - Views\TDE0020\TDE0020IndexTab1.cshtml
  （util.cs 與 SYS0030Controller.cs 的變更已編進 WCals.dll，不需另傳原始碼）

【若工作區被 VS 還原】：先關閉 VS，執行 RESTORE.ps1 即可復原。

-----------------------------------------------------------------
【追加 3.1.4 工程登記冊 TDE0040 (src314\)】(2026-07-23)
帳號 020000 / 020000。填報 6 子彈窗頁籤補成完整版 + 主查詢縣市依部門。
變更 7 檔：
  8. Controllers\NSPMIS\TDE0040Controller.cs  主查詢縣市依部門(SECCITY)
  9. Views\TDE0042\PAGE2.cshtml   頁籤→年度經費/保留款
 10. Views\TDE0043\PAGE3.cshtml   頁籤→預算資料/契約資料/結算資料/決算資料
 11. Views\TDE0045\PAGE5.cshtml   頁籤→工作控制點/辦理情形及施工進度/地方預算執行數填報
 12. Views\TDE0046\PAGE6.cshtml   頁籤→停復工記錄/估驗進度/驗收紀錄/用戶排水設備承裝技工/施工廠商人員/接管戶數/化糞池填拆除統計/監造廠商人員
 13. Views\TDE0048\PAGE8.cshtml   頁籤→年度經費雨水箱(管)涵工程/抽水站工程/雨水調節池工程
 14. Views\TDE0049\PAGE9.cshtml   頁籤→管線工程/污水處理廠工程/抽水站工程
 15. Views\TDE0040\TDE0040Index.cshtml  補上 closePage6_1/3/4/5/6/7/8 函式
     (修真bug：施工資料子頁關閉鈕呼叫未定義函式→停復工等無法關閉)
 (TDE0040Controller.cs 已編進 WCals.dll)
 部署最小集：bin\WCals.dll + 上述 6 個 PAGE*.cshtml + TDE0040Index.cshtml
 已互動測試通過(2026-07-23)：施工廠商人員/監造廠商人員 證照號碼+姓名 DB 寫入讀取正常
   (新增→讀回→刪除還原乾淨)；停復工/子頁關閉函式已定義可正常關閉。

-----------------------------------------------------------------
【追加 3.1.2 rev(1) 鄉鎮別與縣市別連動 (2026-07-25)】帳號 020000 / 020000。
問題：縣市別選宜蘭縣，鄉鎮別下拉卻列出全台鄉鎮（截圖出現台中市「大甲區」）。
根因：連動 JS 綁在 #city 的 onchange，但縣市政府帳號的縣市下拉只有一項且已預選，
      onchange 永遠不會觸發；而 Index 初始載入給的是 util.GetTowns(null)＝全台鄉鎮。
變更 2 檔：
 16. Controllers\NSPMIS\TDE0020Controller.cs  Index 初始鄉鎮改 util.GetTowns(loginCity)
     （中央帳號 loginCity 為 null 時維持全部，由 onchange 過濾，行為不變）
 17. Views\TDE0020\TDE0020Index.cshtml        resetForm() 重置後改依縣市重載鄉鎮（原本清空）
 (TDE0020Controller.cs 已編進 WCals.dll)
 部署最小集：bin\WCals.dll + Views\TDE0020\TDE0020Index.cshtml
 已實測(2026-07-25)：初始鄉鎮=宜蘭縣 13 鄉鎮、無大甲區；與 /ASYNC/GetTowns 回傳清單一致；
   重置後仍為 13 項；查詢與「一次儲存」鈕回歸正常；無 console 錯誤。

【驗證帳號】3.1.2/3.1.7 → B150200；3.1.4 與 3.1.2 rev(1) → 020000（密碼同帳號）。
