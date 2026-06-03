# 20260601 New

## 知識是永遠的勇者｜葉佳山教學網站

本專案整合葉佳山教學網站與微型餐飲創業教學範例。網站主題包含：

- 教學場域：弘光科技大學餐旅系創業系列、旅館管理系列
- 爬山虎系列：登山日誌與網誌
- 餐旅顧問系列：餐旅顧問經歷
- 創業系列範例：La Neige 法式雪冰室

### 教學範例網站

本機預覽：

```bash
python3 -m http.server 4173 --directory docs
```

開啟：

```text
http://localhost:4173
```

GitHub Pages 公開網址：

```text
https://davidivy.github.io/20260601-new/
```

主要頁面：

- 首頁：`/`
- 創業系列：`/entrepreneurship.html`
- 創業計劃書網頁版：`/business-plan.html`
- 旅館管理系列：`/hotel-management.html`
- 爬山虎系列：`/climbing.html`
- 餐旅顧問系列：`/consulting.html`
- La Neige 法式雪冰室範例：`/la-neige/`

### 主要交付物

- `docs/index.html`：葉佳山教學網站首頁
- `docs/entrepreneurship.html`：弘光餐旅系創業系列教學頁
- `docs/business-plan.html`：La Neige 法式雪冰室創業計劃書網頁版
- `docs/la-neige/`：La Neige 法式雪冰室互動範例
- `ice_shop_design/法式雪冰室創業計劃書.md`：創業計劃書 Markdown
- `La_Neige_法式雪冰室創業計劃書.docx`：創業計劃書 Word 檔
- `La_Neige_法式雪冰室創業教學簡報.pptx`：教學簡報
- `ice_shop_design/website/`：網站原始工作區

### 重新產生文件與簡報

```bash
python3 scripts/build_ice_shop_business_docx.py
node scripts/build_ice_shop_teaching_deck.mjs
```
