---
title: "檔案毀損修復"
metaTitle: "Excel 檔案毀損修復"
metaDescription: "Excel 檔案毀損修復"
---

### 匯入Google Sheets或Web Excel

利用[Google Sheets](https://docs.google.com/spreadsheets/u/0/)或是[Web Excel](https://www.office.com/launch/excel)將毀損的Excel檔案匯入或是利用雲端硬碟的方式預覽，底層引擎Google使用Java，Microsoft則是C++，共通點是網頁版必然使用JavaScript渲染，兩者處理最後介面的方式與桌面版Excel不同，有非常大的機率可以修復(看到資料)。(註：部分事務所電腦會限制上傳(匯入) 。)

- Google Sheets
  <img src="https://imgur.com/Omo0OWI.png"  alt="google" />  

- Web Excel
  <img src="https://imgur.com/lPJ3WPa.png"   alt="web excel" />

### Excel內建修復

開啟舊檔，選擇「開啟並修復」

<img src="https://imgur.com/XJIu26I.png"  alt="" width="600" height="400"/>

### 由毀損檔案匯入資料
- 【資料】>【取得資料】>【從檔案】>【從活頁簿】，從跳出的視窗選擇要匯入的資料
  <img src="https://imgur.com/xPKlYDm.png"  style="margin:5px" alt="" width="500" height="450" />
  <img src="https://imgur.com/B4hxYkT.png" style="margin:5px"  alt="" width="500" height="380"/>
- 若無法匯入，但可以由上面步驟知道工作表名稱時，可嘗試直接用公式匯入資料，例如活頁簿名稱為「明細帳.xlsx」，該檔案中有一個「工作表1」，在[A1]輸入「=[明細帳.xlsx]工作表1!A1」，再將公式拉到其他儲存格。
  
### 取消隱藏
- 若雙擊檔案都沒有反應，有可能因活頁簿被設定為隱藏的狀態，將隱藏解除即可，【檢視】>【取消隱藏視窗】。
  <img src="https://imgur.com/IpgIXNN.png"   alt="" />