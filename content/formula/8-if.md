---
title: "IF"
metaTitle: "IF"
metaDescription: "EXCEL IF 公式教學"
---

### 條件判斷

> =IF(logical_test, if True, if False)

1. 常搭配AND()公式使用，此公式用作邏輯測試判斷TRUE or FALSE：
  - =AND(1=1,2=2,3=3) 傳回TRUE
  - =AND(1=1,2=2,3=2) 傳回FALSE

<div style="margin-left:30px;">
<iframe width="538" height="135" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='6.2%20IF'!A1%3AH5&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>

2. IF可以寫成巢狀來處理更多的情況(第一個條件不符時，進行下一個條件判斷)：
  - IF(logic_test, if True, if(logic_test, if True, if(logic_test, if True, False)))

<div style="margin-left:30px;">
<iframe width="700" height="215" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='6.1%20IF'!A1%3AX8&wdDownloadButton=True&wdInConfigurator=True"></iframe>
</div>

3. 因為Excel公式都要寫成一行，較難閱讀，如果真的遇到要寫巢狀公式時，可以拿紙筆出來實際寫較有感覺。  
IF…THEN  
  xxx  
ELSE IF  
  xxx  
ELSE  
  xxx  
END IF