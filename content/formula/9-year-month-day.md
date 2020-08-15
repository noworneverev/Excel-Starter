---
title: "YEAR/MONTH/DAY"
metaTitle: "YEAR/MONTH/DAY"
metaDescription: "Excel 時間公式教學"
---

### 時間公式

> =YEAR(日期), =MONTH(日期), =DAY(日期)

1. 快捷鍵：插入目前日期，按 Ctrl+; (分號)；插入目前時間，按 Ctrl+Shift+; (分號)。
2. 日期可比較先後，也可相加減，原理係因日期格式本質為一個數值，「1」為「1900/1/1」，「2」為「1900/1/2」以此類推，所以下方公式中2018/3/31 > 2017/1/1傳回TRUE (43190 > 42736)，兩日期相差454天(43190 – 42736)
3. 若要取得某日期的年月日，要用YEAR、MONTH及DAY函數或資料剖析功能，非LEFT或RIGHT等截取字元函數，如上所述日期本質為一數值，所以以LEFT函數取2018/3/31左4個字元，會取到4319而非所需的2018年。

<div style="margin-left:30px;">
<iframe width="644" height="294" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='7.%20YEAR%20MONTH%20DAY'!A1%3AF12&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>
