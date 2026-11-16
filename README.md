# 個人健康檢驗紀錄 Personal Health Check Record

把健保署「健康存摺」下載的 XML 檔（或本工具產生的統計 Excel）匯入，自動整理就醫、用藥與檢驗統計的單頁網頁工具。

**線上使用：** https://james-jj2.github.io/personal-health-check-record/ （需先在 Settings → Pages 開啟）

**離線使用：** 按右上角 Code → Download ZIP，解壓縮後用瀏覽器開啟 `健康存摺統計工具.html`。

- `index.html`、`健康存摺統計工具.html`：統計工具（兩個檔案內容相同，前者給線上網址用）。用瀏覽器開啟，把檔案拖進去即可。電腦、筆電、手機都能用。
- `health_bank_to_excel.py`：把健康存摺 XML 轉成含公式的 Excel（`python health_bank_to_excel.py 健康存摺.xml 輸出.xlsx`，需安裝 openpyxl）。

## 功能

- 總覽：就醫次數、部分負擔、健保點數，年度／院所／主診斷統計
- 檢驗趨勢：每個項目一列，可勾選、可展開；勾選後列出曲線、統計表、逐次數值對照表
- 快速選取：甲狀腺、血糖、血脂、肝腎功能、血壓體重
- 門診紀錄、用藥、牙科、疫苗與影像報告
- 匯出 Excel、深色／淺色切換

## 隱私

資料只在瀏覽器裡處理，不會上傳。**請不要把自己的健康存摺 XML 或 Excel 檔放進這個 repo**（`.gitignore` 已排除 `.xml`、`.xlsx`）。頁面一開始顯示的是虛構的示範資料。

本工具只整理資料，不提供醫療判讀。
