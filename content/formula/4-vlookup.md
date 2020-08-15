---
title: "VLOOKUP"
metaTitle: "VLOOKUP"
metaDescription: "EXCEL VLOOKUP 公式教學"
---

import {Link} from "gatsby";

### 找值

> =VLOOKUP(你要找的值, 在哪裡找, 回傳第幾欄, 是否完全符合)

1. 為了避免因為資料來源沒該筆查找資料而跑出#N/A的狀況，常將公式改寫加入「IFERROR」進行錯誤處理，將#N/A變空白""或0。
2. 在哪裡找(資料來源)不是設整欄的話，要記得將範圍位址固定住(按F4)。
3. 如果傳回資料是空白，沒有資料時，該儲存格格式沒有另外設定的話會顯示0，如下方[C7]，如果想要跟資料來源一樣是空白的話，在公式後面加上「& ""」，Excel會因此判斷格式為文字，即可顯示空白，如下方[C8]。

<div style="margin-left:30px;">

<iframe width="546" height="282" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=56FC2EC950646865%21118&authkey=%21APgHYfXy-vTH64Y&em=2&wdAllowInteractivity=False&AllowTyping=True&Item='3.%20VLOOKUP'!A1%3AH14&wdDownloadButton=True&wdInConfigurator=True"></iframe>

</div>

4. VLOOKUP先天限制是它只能尋找表格最左欄的值，此時Excel還有<Link to="/formula/5-index-match">INDEX</Link>與<Link to="/formula/5-index-match">MATCH</Link>的組合可以來取代VLOOKUP，查找大量資料時速度也是遠勝VLOOKUP；另微軟於2020正式推出<Link to="/formula/7-xlookup">**XLOOKUP**</Link>，不論是使用的彈性或效能上都優於VLOOKUP。

