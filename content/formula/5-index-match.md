---
title: "INDEX/MATCH"
metaTitle: "INDEX/MATCH"
metaDescription: "Excel INDEX MATCH 公式教學"
---

import {Link} from "gatsby";

### 單條件查詢(INDEX/MATCH)

> =INDEX(資料來源, 第幾列, 第幾欄)  
> =MATCH(要找的值, 從哪裡找, 符合類型)

1. INDEX常與MATCH搭配使用，MATCH 函數會搜尋儲存格範圍中的指定項目，並傳回該項目於該範圍中的相對位置。 例如，若範圍 A1:A3 中含有值 5、25 及 38，則公式 =MATCH(25, A1:A3, 0) 會傳回數字 2，因為 25 是範圍中的第二個項目。

<div style="margin-left:30px;">

<iframe width="261" height="156" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='4.1%20INDEX%20MATCH'!A1%3AD6&wdInConfigurator=True"></iframe>

</div>

2. 實際用法：用MATCH先找到你的資料是第幾列，再從INDEX公式指定要回傳的欄位，範例檔的4月匯率表為使用INDEX & MATCH完成可參考。
> =INDEX(資料來源, MATCH(~~~), 第幾欄)

<div style="margin-left:30px;">

<iframe width="581" height="240" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='4.1%20INDEX%20MATCH'!A9%3AI18&wdInConfigurator=True"></iframe>

</div>

### 水平+垂直查詢(INDEX/MATCH/MATCH)

> =INDEX(資料來源, MATCH(要找的列值, 從哪些列找, 0), MATCH(要找的欄值, 從哪些欄找, 0))

1. 使用INDEX搭配兩個MATCH，第一個MATCH尋找列，第二個尋找欄，即可同時進行垂直及水平查詢。

<div style="margin-left:30px;">
<iframe width="700" height="299" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='4.1.1%20INDEX%20MATCH%20MATCH'!A1%3AK12&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>

2. 注意兩個MATCH公式第一個引數的寫法 **$A2** 及 **C$1**，分別在列及欄用相對位址，這樣公式才可以直接用拉的填滿。
3. 使用XLOOKUP/XLOOKUP可達成同樣目的。

### 多條件查詢(INDEX/MATCH/陣列)(單一結果)

1. 利用在SUMPRODUCT公式提到的陣列判斷式，進行多條件查詢：先利用MATCH找出資料在第幾列：

>  {=MATCH(1, (F3=A2:A32)\*(G3=D2:D32)\*(C2:C32>9000), 0)}

2. MATCH要尋找的值設定為「1」，因為後面陣列判斷會傳回1(TRUE)或0(FALSE)，傳回1代表所有條件符合(1\*1\*1=1)，此時MATCH傳回該筆資料在第幾列，最後再利用INDEX定位第幾欄得到結果。(注意：公式打完後務必按**Ctrl+ Shift+ Enter**，告訴Excel這是**陣列公式**)

> =INDEX(A2:D32, MATCH(1, (F3=A2:A32)\*(G3=D2:D32)\*(C2:C32>9000),0), 3)

<div style="margin-left:30px;">
<iframe width="700" height="740" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='4.2%20INDEX%20MATCH%20%E9%99%A3%E5%88%97'!A1%3AN33&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>

