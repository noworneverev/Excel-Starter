---
title: "Excel 樞紐分析表"
metaTitle: "Excel 樞紐分析表"
metaDescription: "This is the meta description for this page"
---

import {Link} from "gatsby";

### 用途

- 計算、建立摘要和分析資料(插入 > 樞紐分析表)
- 主要用於大量不同列同一欄有同樣值的情況，例如帳齡分析表(同一客戶)、傳票整理(同一分錄、會科)

### 背後邏輯

- 利用分組(Group By)和轉置(Transpose)將資料彙整並計算(Sum, Count, Avg, etc.)

### 相關快捷鍵

- ALT + F5：更新目前選擇的樞紐分析表
- ALT + F1：產生目前選擇樞紐分析表的圖表(Pivot Chart)

### 範例

- 使用公式<Link to="/formula/3-sumproduct">SUMPRODUCT</Link>整理的資料，用樞紐分析表可得同樣結果

<img src="https://imgur.com/1BGk0of.png" style="margin:5px"  alt="pivot1" width="400" height="140" />
<img src="https://imgur.com/af02YHz.png" style="margin:5px"  alt="pivot3"  width="290" height="180" />
<img src="https://imgur.com/h8Kpzab.png" style="margin:5px"  alt="pivot2" />



### 交叉分析篩選器(Silcers)

- 按一下樞紐分析表中的任一位置，選取【樞紐分析表分析】> 插入【交叉分析篩選器】

<iframe width="677" height="528" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='2.%20SUMPRODUCT_2'!M1%3AV25&wdDownloadButton=True&wdInConfigurator=True"></iframe>
