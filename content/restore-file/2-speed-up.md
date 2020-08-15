---
title: "檔案整理加速"
metaTitle: "Excel 檔案整理加速"
metaDescription: "Excel 檔案整理加速"
---

### 清除無效名稱

- 【資料】>【編輯連結】，中斷無效連結<br />
<img src="https://imgur.com/gdmgEcN.png"  style="margin:5px" alt="" width="180" height="120" />
<img src="https://imgur.com/StHncYA.png"  style="margin:5px" alt="" width="450" height="280" />
- 【公式】>【名稱管理員】，將名稱管理員中不需要的名稱刪除<br />
<img src="https://imgur.com/AwWtiDc.png"  style="margin:5px" alt="" width="180" height="120" />
<img src="https://imgur.com/z7V8FML.png"  style="margin:5px" alt="" width="450" height="200"  />
- 若工作簿開啟時間緩慢，很有可能是無效名稱過多，可參考以下連結文章以VBA批次刪除：  
[[心得] Excel 清除無效連結](https://www.ptt.cc/bbs/Accounting/M.1585674918.A.1AA.html)
或是開啟VBE(ALT + F11)新增模組執行此段程式碼：

```
Sub RemoveInvalidNames()
  Dim i As Integer
  Dim name As name
  Dim workbookNames As Names
  Set workbookNames = ActiveWorkbook.Names

  i = 0
  For Each name In workbookNames
    If InStr(name.Value, "#REF") > 0 Then
      i = i + 1
      ActiveWorkbook.Names(name.name).Delete
    End If
  Next
  MsgBox "清理完成，共清除" + Str(i) + "個無效名稱！"
End Sub
```

### Inquire COM增益集
- 右鍵上方功能列 >【自訂功能區】> 勾選【開發人員】>【COM 增益集】> 勾選【Inquire】
<img src="https://imgur.com/rcXLroW.png"  style="margin:5px" alt="" width="270" height="120" />
<img src="https://imgur.com/gvbt8dI.png"  style="margin:5px" alt="" width="400" height="200" />
<img src="https://imgur.com/Ll3Y5T4.png"  style="margin:5px" alt="" />
- 在 【Inquire】 索引標籤上，使用【清除多餘的儲存格格式設定】 ![](https://imgur.com/VvyJePO.png)
- 此功能還可用在清除多餘空白列

### 另存成xlsb
- 二進位活頁簿(.xlsb)可大幅度壓縮檔案及提升開啟、儲存速度
- 將檔案另存為xlsb時，是以二進位的方式儲存而非XML的形式(xlsx)，越大的檔案壓縮的效果會越明顯
- 二進位活頁簿的優缺點可參考：[https://analystcave.com/excel-working-with-large-excel-files-the-xlsb-format/](https://analystcave.com/excel-working-with-large-excel-files-the-xlsb-format/)  <br />
<img src="https://imgur.com/x6U0thM.png"   alt="" width="328" height="135" />

