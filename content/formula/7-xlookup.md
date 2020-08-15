---
title: "XLOOKUP"
metaTitle: "XLOOKUP"
metaDescription: "Excel XLOOKUP 公式教學"
---

### XLOOKUP 特性

> =XLOOKUP(要找的值, 從哪裡找, 傳回什麼, [錯誤說明], [相符類型], [搜尋模式])

- 與VLOOKUP差異最大的地方在於參數「從哪裡找(資料來源)」：VLOOKUP是尋找一個表格(table_array)並據此指定傳回第幾欄，最大限制莫過於因此第一欄必須包含要查閱的值；
- XLOOKUP明確指定「從哪裡找(lookup_array)」及「傳回什麼(return_array)」，在此兩個參數都彈性的情況下，根據實際傳入的引數，不論是**垂直查詢(VLOOKUP)**、**水平查詢(HLOOKUP)**或是**要查閱的值在傳回值的右邊(原VLOOKUP無法查詢，須使用INDEX+MATCH)**，使用XLOOKUP可囊括上列所有情況。
- 【從哪裡找】及【傳回什麼】列或欄的維度要一致，否則公式會有錯誤。
- XLOOKUP傳回**參照**(cell range reference)，非值。
- 第4個選擇性參數，錯誤說明，提供了在回傳#N/A(找不到)時可以顯示的說明，使用查詢公式再也不用外包IFERROR提示錯誤。
- 第5及第6個選擇性的參數為更進階的操作，例如從後面開始查詢、模糊比對等。
- 此公式目前僅在OFFICE 365支援，建議不要使用於需要提供給客戶的底稿，除非確定客戶開啟XLOOKUP公式不會出現亂碼。

### 基本範例

- 要找的值：1122
- 從哪裡找：A5:A7
- 傳回什麼：B5:B7

<div style="margin-left:30px;">
<iframe width="500" height="220" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='5.1%20XLOOKUP'!A1%3AG9&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>

### 更多範例

- 含錯誤訊息(第4個參數)
- 由後方查詢(第5個參數)
- 資料來源查閱索引在右邊
- 水平查詢

<div style="margin-left:30px;">
<iframe width="700" height="467" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='5.2%20XLOOKUP'!A1%3AH20&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>

### 相對位址使用

- 相較VLOOKUP，使用XLOOKUP時，須更熟悉相對位址的使用，例如下方使用XLOOKUP公式做出的匯率表，僅須編寫[C4]儲存格的公式(黃底處)，查閱處鎖欄不鎖列，傳回處鎖列不鎖欄，剩餘公式都是用拉的。

<div style="margin-left:30px;">
<iframe width="404" height="338" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='2018.04%E5%8C%AF%E7%8E%87%E8%A1%A8(XLOOKUP)'!A1%3AF14&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>

### 水平 + 垂直查詢 (XLOOKUP/XLOOKUP)

> =XLOOKUP(要找的列值, 從哪些列找, XLOOKUP(要找的欄值, 從哪些欄找, 完整資料來源)

- 此公式用到XLOOKUP一重要特性：XLOOKUP傳回參照(cell range reference)，而非值。
- 內層XLOOKUP的第三個引數是資料來源的完整表格，而非單一列或欄，如此會傳回一個動態陣列讓外層XLOOKUP使用，作為其回傳的資料範圍。
- 內外層XLOOKUP的第一個引數分別固定列及固定欄。
- INDEX/MATCH/MATCH可達成同樣目的。

<div style="margin-left:30px;">
<iframe width="650" height="300" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='5.3%20XLOOKUP%20%E6%B0%B4%E5%B9%B3%E5%9E%82%E7%9B%B4%E6%9F%A5%E8%A9%A2'!A1%3AJ12&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>