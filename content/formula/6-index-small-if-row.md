---
title: "INDEX/SMALL/IF/ROW"
metaTitle: "INDEX/SMALL/IF/ROW"
metaDescription: "Excel 多條件查詢 多個結果 公式教常"
---

### 多條件查詢(INDEX/SMALL/IF/ROW)(多個結果)

1. 若在多條件查詢下有可能有多筆資料符合，以上列公式組合可以得到結果，下方公式為例：

> =INDEX($A:$E, SMALL(IF(($G$3=$B$2:$B$40)\*($H$3=$E$2:$E$40)\*($D$2:$D$40>5000),ROW($2:$40)),ROW(1:1)),4) 

- =IF(($G$3=$B$2:$B$40)\*($H$3=$E$2:$E$40)\*($D$2:$D$40>5000), ROW($2:$40))
  - 先利用IF結合陣列判斷式判斷3個條件，若3個條件符合，IF公式為TRUE，傳回ROW($2:$40)
  - ROW函數會傳回第幾列，例如=ROW(A1)=1，因為A1在第一列；=ROW(2:2)為第2列
  - 公式設定ROW($2:$40)是因為前面的條件設定中，陣列一致都是列2~列40，全部加$符號固定住(因為之後公式要往下拉，原始資料位置不能動)
  - IF公式在這裡的作用是儲存查詢目標在第幾列(陣列)。

- =SMALL(IF(～), ROW(1:1))
  - =SMALL(資料, 第幾小)，SMALL公式會在一資料範圍中傳回指定第幾小的數值。
  - 資料的部分已經是用上面IF公式儲存好符合條件的陣列，所以當IF傳回陣列{3, 12, 16, 20}，用ROW(1:1)=1，來找出最小的那筆，也就是第一個符合所有條件的目標。
  - ROW(1:1)不用加絕對符號($)固定(因為要用來下拉公式)，往下拉變ROW(2:2)，找出第2筆的查詢目標，以此類推。

- SMALL + IF + ROW的組合決定了資料在第幾列，再用INDEX設定資料來源與要查詢第幾欄，傳回所查詢的資料，最後將公式下拉得所有符合條件的結果。

<div style="margin-left:30px;">

<iframe width="700" height="900" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='4.3%20INDEX%20SMALL%20IF%20%E9%99%A3%E5%88%97'!A1%3AT41&wdDownloadButton=True&wdInConfigurator=True"></iframe>

</div>

- 陣列公式在實際做底稿時不太可能用到，此範例僅演示Excel用公式可以做到的操作，看不懂跳過沒關係。


