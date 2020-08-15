---
title: "SUMIF"
metaTitle: "SUMIF"
metaDescription: "Excel SUMIF 公式教學"
---

### 基本SUMIF 

> =(SUMIF準則範圍, 加總準則, [加總範圍])

1. 若沒有加總儲存格的範圍，加總範圍默認為準則範圍，簡單範例如用"α"這個符號來標示你要加總的地方(A2:A5)，將"α"打在如[B2]、[B3]儲存格：=SUMIF(B2:B5, "α", A2:A5) = 300

<div style="margin-left:30px;">

<iframe width="605" height="102" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='1.%20SUMIF%E3%80%81SUMFIFS'!A2%3AF5&wdDownloadButton=True&wdInConfigurator=True"></iframe>

</div>

2. 在打公式時按**F4**將該範圍變成**絕對位置**，如此公式才不會在複製到別的位置時跑掉(一般Excel的公式都是**相對位置**，所以例如複製[A11]=SUM(A1:A10)到[B11]時公式會自動變成= SUM(B1:B10))。

<div style="margin-left:30px;">

<iframe width="605" height="176" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='1.%20SUMIF%E3%80%81SUMFIFS'!A7%3AF14&wdDownloadButton=True&wdInConfigurator=True"></iframe>

</div>

### 模糊條件求和

1. 利用萬用字元「**&ast;**」配合SUMIF，可用以判斷準則範圍內含**某特定字元**，例如要加總下表的活存金額時，加總準則輸入"&ast;活存&ast;"，星號代表未知的字元，所以XX活存或活存XX都會被判斷是要加的。

<div style="margin-left:30px;">

<iframe width="605" height="300" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='1.%20SUMIF%E3%80%81SUMFIFS'!A19%3AE33&wdDownloadButton=True&wdInConfigurator=True"></iframe>

</div>