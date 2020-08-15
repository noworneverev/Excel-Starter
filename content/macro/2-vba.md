---
title: "VBA 編寫自定義快捷鍵"
metaTitle: "VBA 編寫自定義快捷鍵"
metaDescription: "VBA 編寫自定義快捷鍵"
---

import {Link} from "gatsby";

### 開啟shortcuts.xlam

1. 安裝<Link to="/custom-shortcut">Excel 自訂快捷鍵</Link>
2. 開發人員 → Visual Basic (Alt + F11)
3. 模組裡已有22個簡單的巨集，以填滿黃色(Fill_Yellow)做說明：
```
Sub Fill_Yellow()
  Selection.Interior.Color = RGB(255, 255, 0)
End Sub
```
4. 點選Microsoft Excel物件中的ThisWorkbook，右邊下拉式選單選Open會跑出：
```
 Private Sub Workbook_Open()

 End Sub
 ```
5. 指定快速鍵如以下寫法：<br />("+y"的加號是Shift的意思，後方引數是函式的名稱，即步驟3的Fill_Yellow(子程序名稱))
```
Private Sub Workbook_Open()
  Application.OnKey "+y", "Fill_Yellow"
End Sub
```
6. 上面幾行程式的白話文：在活頁簿開啟時(Workbook_Open)載入快捷鍵「Shift + Y」("+y")，連結的功能為填滿黃色(Fill_Yellow)。
7. 可結合Shift、Ctrl與Alt自訂快速鍵，範例請詳shortcuts.xlam其他快速鍵(使用Ctrl組合快速鍵要小心不要覆蓋掉原本Excel的功能，除非真的不會用到)
8. 呼叫Excel原有功能寫法如下：<br />(此為shortcuts.xlam裡的其中一個功能-增加小數，是Excel的原生功能，VBA語法執行該Excel的控制項名字(DecimalsIncrease)即可啟動該功能，Excel所有控制項原名詳[連結](https://www.microsoft.com/en-us/download/details.aspx?id=50745)。)  
```
Private Sub DecimalsIncrease()
  Application.CommandBars.ExecuteMso ("DecimalsIncrease")
End Sub
```
9. 其餘快速鍵更詳細的說明可參考[MSDN文件](https://docs.microsoft.com/zh-tw/office/vba/api/Excel.Application.OnKey)。