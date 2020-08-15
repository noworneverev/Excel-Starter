---
title: "SUMPRODUCT"
metaTitle: "SUMPRODUCT"
metaDescription: "Excel SUMPRODUCT 使用教學"
---

import {Link} from "gatsby";

### 基本使用(對應的陣列相乘後相加(乘積和))

> =SUMPRODUCT(陣列1, [陣列2], [陣列3], …)

1. 如下範例，SUMPRODUCT公式傳回所有對應元素乘積的總和，注意陣列的維度要一致。

<div style="margin-left:30px;">

<iframe width="389" height="174" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='2.%20SUMPRODUCT_1'!A1%3AF8&wdDownloadButton=True&wdInConfigurator=True"></iframe>

</div>

### 實際應用(多條件判斷加總運算)

1. 以下面的範例說明，將左方的資料，以**科目編號**及**分類**作為條件整理出其對應之**餘額**：

> **科目編號為1122**且**分類為瓶胚**的餘額 = SUMPRODUCT(($A$2:$A$31=$F4)\*($D$2:$D$31=G$3), $C$2:$C$31)

- 公式中陣列1為($A$2:$A$31=$F4)\*($D$2:$D$31=G$3)，陣列中的判斷式會傳回TRUE或FALSE的陣列(A2=F4成立 → 傳回TRUE，A3=F4成立 → 傳回TRUE ...直到 A31=F4不成立→傳回FALSE，如此會有一個布林陣列)，
- A2:A31=F4與D2:D31=G3這兩個陣列相乘的意思是：若A2=F4傳回TRUE，接著若D2=G3也傳回TRUE時，
- TRUE為1，FALSE為0，TRUE*TRUE = 1，也就是此時陣列1的運算結果為1，再乘上陣列2(C2:C31)的餘額數字，可得到目前第一個算出來的數字8,777。
- 若($A$2:$A$31=F4)*($D$2:$D$31=$G$3)，這兩個判斷式中任一個跑出FALSE，即傳回0，例如A13不等於F4，陣列1運算出來是0，乘上後面的餘額後得到就是0，所以不會計入乘積和中。
- 接著公式以此類推將所有合乎條件者(A2:A31範圍中編號是1122，且D2:D31範圍中分類是瓶胚者)相加，得到**科目編號為1122**且**分類為瓶胚**的總和。
- 只需要寫有上色此儲存格的公式，其餘用拉的，注意$F4及G$3的用法，$F4只鎖住F欄，G$3只鎖住第3列。
- 以<Link to="/pivot-table">樞紐分析表</Link>可快速得到相同結果。

<div style="margin-left:30px;">

<iframe width="700" height="717" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='2.%20SUMPRODUCT_2'!A1%3AV33&wdDownloadButton=True&wdInConfigurator=True"></iframe>

</div>

2. 簡而言之，=SUMPRODUCT((條件1)\*(條件2)\*(條件3), 要相加的數值)，可將符合條件的欄位相加。
3. 公式可改寫成=SUMPRODUCT(($A$2:$A$31=$F4)\*1, ($D$2:$D$31=G$3)\*1, $C$2:$C$31)，將兩個條件放在陣列1與2中，再分別「\*1」，乘以1為必要步驟，將條件判斷完的TRUE或FALSE轉成1/0陣列。

<div style="margin-left:30px;">

<iframe width="450" height="135" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='2.%20SUMPRODUCT_2'!F10%3AK14&wdDownloadButton=True&wdInConfigurator=True"></iframe>

</div>