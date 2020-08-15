---
title: "PDF2Excel"
metaTitle: "PDF轉Excel"
metaDescription: "PDF轉Excel"
---

### PDF2Excel
此程式以Python開發，使用轉換精準度最高的函式庫，再搭配另一支VBA小程式，可快速將PDF檔案裡的表格轉成Excel格式輸出。

### 下載
PDF2Excel.exe：[https://tinyurl.com/yaxku6bj](https://tinyurl.com/yaxku6bj)  
搭配使用的VBA：[https://tinyurl.com/yd7t5m9b](https://tinyurl.com/yd7t5m9b)

### Demo
短片演示：![https://i.imgur.com/WCbBVIe.gif](https://i.imgur.com/WCbBVIe.gif)  

影片演示：
<iframe width="560" height="315" src="https://www.youtube.com/embed/0vEI2oiTanM" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>

### 使用限制
此程式適用於無須文字辨識(OCR)的PDF，僅測試於Windows作業系統。

### 使用方法
- 點開PDF2Excel.exe，無須安裝，因將Python環境包裝在單一執行檔中，開啟程式可能會需要點時間。
- 選取分割PDF表格的方式，垂直及水平各分三種切割方式：
  - lines
  - lines_strict
  - text
- 如果PDF的表格用「線」分欄的話，則垂直分割選「lines」；列只有文字沒有線的話，水平分割則選「text」可達到最佳效果，公開資訊觀測站上財報用這種方式應可截取到最完整的表格。
- 選擇PDF檔案存放的資料夾，**批次轉換資料夾內所有PDF檔案**成Excel檔案，只轉換PDF檔案裡的**表格**，文字區塊一律跳過。
- Excel工作表命名原則以PDF頁碼當作工作表名稱，例如轉換第一頁的表格，輸出的Excel工作表名稱為"Sheet1"；若一頁裡偵測到多個表格，例如第三頁有兩個表格，輸出"Sheet3_1"、"Sheet3_2"，以此類推。
- 輸出Excel檔案後，使用VBA增益集([Text2Column.xlam](https://tinyurl.com/yd7t5m9b))，將字串轉成Excel可運算的儲存格格式。
- 轉換僅在本地端執行，無須擔心資料外洩，若有疑慮請詳下方原始碼。

### 給開發者
以Python寫成，關鍵的轉換只有十幾行程式，若已有Python環境可參考[PDFPlumber Github文件](https://github.com/jsvine/pdfplumber)自行客製參數，若熟pandas可以再更進一步依照提取出的資料另做處理。  
pip install pdfplumber  
pip install pandas  

### 開放原始碼
[https://github.com/noworneverev/PDF2Excel](https://github.com/noworneverev/PDF2Excel)