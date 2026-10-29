# Receipt OCR that updates the stock sheet by itself

Google Apps Script for my family's electronics shop in Mongolia (two shops and a warehouse). The staff photograph a
receipt or a slip with a phone; Gemini reads the product names, quantities and dates; my mother marks each line
correct or wrong on a daily review sheet; the checked numbers are written into the stock workbook (Google Sheets),
and every written cell keeps a link back to its receipt line.

**My mother's daily checking: about an hour on a normal day → about ten minutes.** (In busy seasons it was two hours.)

- Portfolio, with videos of the live system: https://khongor-portfolio.vercel.app/
- 日本語の説明は下にあります。

## How it works

```
① Phone photo ─ into a Drive folder per shop and kind (deliveries / sales / transfers / stocktaking)
② Every hour (and at once after a check), processNewReceipts finds new photos
③ Gemini API reads the slip → names matched against a master list of ~580 products
④ A review spreadsheet per place, kind and month, one tab per day: my mother picks Зөв / Буруу (correct / wrong)
⑤ The checked numbers go into the stock workbook; each cell's note links to its receipt line
⑥ A progress calendar is fed to a web page (Cloudflare Pages, not in this repository)
```

Colours, the same in the review sheets and the stock workbook: yellow = the AI was confident, pink = needs a check,
green = checked by my mother.

## What is in here

| Path | What |
|---|---|
| `ocr/` | the receipt OCR: Gemini call and prompt (`ocr-gemini.gs`), review sheets (`ocr-sheet.gs`, `ocr-review.gs`), writing into the workbook (`ocr-post.gs`), new products (`ocr-new-product.gs`), reference repair (`ocr-relink.gs`), the hourly run (`ocr-main.gs`), the progress calendar feed (`ocr-calendar.gs`), reports and checks |
| `*.gs` (root) | the stock workbook: yearly files and rollover, half-month sheets, formulas, protection, price and name sync, transfers between shops, the product list |
| `profile.gs`, `config.gs` | the one place for settings; `DRY_RUN` is the single switch for every script |
| `tests/` | about 1,600 checks, run with Node.js against a fake Spreadsheet / Drive / Gemini (`fake_book.js`) |
| `build.sh` | joins the files into one bundle to paste into the Apps Script editor |

Things I had to solve:

1. **New products**: when my mother approves a product the workbook does not have, the AI picks its category (the
   price counts too, so an expensive item never lands in "small goods") and a row is added to every later month.
2. **References that repair themselves**: added rows shift saved cell positions, so each month's row order is kept
   as a hash and a change rechecks that month. Broken references went from 133 to 3.
3. **The six-minute limit**: every run watches its deadline, stops cleanly and resumes a minute later.
4. **A fast feed**: the calendar's data is stored compressed in Script Properties instead of opening eight files per
   request: 14 s → 1.2 s.

## Running the tests

```bash
for t in tests/*_test.js; do node "$t" . || break; done
```

## Using it

1. `sh build.sh ocr` (receipt OCR, its own Apps Script project) or `sh build.sh` (the stock workbook) writes
   `neg-file/*.gs`; paste it into the Apps Script editor.
2. Replace the `YOUR_..._ID` placeholders (`ocr/ocr-config.gs`, `year-books.gs`, `product-list.gs`,
   `tosgon-names.gs`, `ocr/ocr-post.gs`, `ocr/ocr-reread.gs`) with your own spreadsheet and folder IDs.
3. Project Settings → Script properties: `GEMINI_KEY` (the Gemini API key), `OCR_REVIEWERS` (JSON of the reviewers'
   Google accounts), `OCR_FEED_TOKEN` (only for the calendar feed). No key or address is kept in the code.
4. Run with `DRY_RUN = true` first: nothing changes, the log says what would happen. Then `DRY_RUN = false` and
   `installOcrTrigger` (OCR) or `installAutomation` (workbook).

The code and its comments are written for this shop: the sheet names are Mongolian, and the shops are Соронзон,
My home and Тосгон (a warehouse workbook, whose own scripts are not in this repository).

## About this copy

This is a cleaned copy of a private repository, published with my family's consent: spreadsheet and folder IDs are
placeholders, account addresses are made up, and no business data (stock sheets, prices, receipts) is included.
All rights reserved; please ask before reusing the code.

---

## 日本語

家族が経営する家電店（モンゴル、2店舗＋倉庫）のための Google Apps Script です。スタッフがレシートや伝票をスマートフォンで撮影すると、
Gemini が品名・数量・日付を読み取り、約580品目の商品マスタと照合します。母が日ごとの確認シートで「正しい／誤り」を選ぶと、
在庫表（Google スプレッドシート）へ自動で記入され、セルのメモから元のレシートの行へ戻れます。

**普段は1時間かかっていた母の確認作業が、約10分に。**（忙しい時期は毎日2時間）

- 新商品の自動追加（価格も判断材料にしてカテゴリーを選ぶ）
- 参照のずれの自動修復（行の並びをハッシュで記録、ずれた参照 133件 → 3件）
- 6分の実行制限への対応（締め切りで区切り、1分後に自動で再開）
- 進捗サイトの応答 14秒 → 1.2秒（データを圧縮して Script Properties に保存）
- 約1,600件の自動テスト（Spreadsheet・Drive・Gemini はフェイクで検証）、試運転モード（`DRY_RUN`）

ポートフォリオ（実際の画面の動画つき）: https://khongor-portfolio.vercel.app/

このリポジトリは非公開リポジトリから家族の了承を得て公開したものです。スプレッドシートやフォルダのIDは置き換え、
アカウントのアドレスは架空のものにし、在庫表・価格・レシートなどの業務データは含めていません。
