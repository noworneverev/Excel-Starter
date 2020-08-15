---
title: "Excel 巨集"
metaTitle: "Excel 巨集"
metaDescription: "Excel 巨集"
---

### 錄製巨集
1. 右鍵上方功能列 >【自訂功能區】> 勾選【開發人員】
2. 【開發人員】>【錄製巨集】
3. 如果要讓巨集在每次開啟活頁簿時都可以使用，一般會將巨集儲存在**個人巨集活頁簿**(personal.xlsb)，錄製完儲存後，之後當開啟Excel時，個人巨集活頁簿也會跟著啟動，但它會是以隱藏活頁簿的形式隱藏起來，單純讓使用者執行儲存在其中的巨集。  
![](https://imgur.com/ATaMFGQ.png)&nbsp;&nbsp;
![](https://imgur.com/wmfiEJ0.png)

### 巨集範例
1. 在按下錄製巨集後，Excel會記錄之後的每一步驟，例如下圖新增一張工作表後，新增標題後選取[A2]
![](https://imgur.com/3fx8aIk.gif)  
2. 自動產生的巨集程式如下：
```
Sub 巨集1()
'
' 巨集1 巨集
'
' 快速鍵: Ctrl+Shift+A
'
    Sheets.Add After:=ActiveSheet       ' <-在目前工作表後新增工作表
    Range("A1").Select                  ' <-選取[A1]
    ActiveCell.FormulaR1C1 = "我是標題"　' <-在目前的儲存格輸入"我是標題"
    Range("A2").Select
End Sub
```
3. VBA較一般程式語言容易學習，因為可以透過錄製巨集的方式知道Excel是透過怎樣的程式碼處理怎樣的程序。改寫上述程式碼，將After改Before，功能變成在目前工作表「前」新增工作表。

### 使用巨集
1. 錄製設定的快捷鍵
2. 插入按鈕(表單控制項)  
![](https://imgur.com/O4uO7Mm.png)
3. 自訂上方功能列的按鈕觸發巨集  
![](https://imgur.com/p0w8eOo.png)
4. 開啟VBE(ALT + F11)執行(F5)巨集。  

