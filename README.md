# 20260601 New

## La Neige 法式雪冰室教學案例

本專案整合一個微型餐飲創業教學範例：以「台灣冰品 x 法式生活感 x 父子退休創業」為主題，完成品牌故事、創業計劃書、簡報與互動式網站原型。

### 教學範例網站

本機預覽：

```bash
python3 -m http.server 4173 --directory docs
```

開啟：

```text
http://localhost:4173
```

GitHub Pages 設定：

1. 將本 repo 推上 GitHub。
2. 進入 GitHub Repository `Settings`。
3. 選擇 `Pages`。
4. Source 選 `Deploy from a branch`。
5. Branch 選 `main`，Folder 選 `/docs`。
6. 儲存後即可取得公開教學網址。

預期網址格式：

```text
https://davidivy.github.io/20260601-new/
```

### 主要交付物

- `docs/index.html`：可上 GitHub Pages 的互動網站範例
- `ice_shop_design/法式雪冰室創業計劃書.md`：創業計劃書 Markdown
- `La_Neige_法式雪冰室創業計劃書.docx`：創業計劃書 Word 檔
- `La_Neige_法式雪冰室創業教學簡報.pptx`：教學簡報
- `ice_shop_design/website/`：網站原始工作區

### 重新產生文件與簡報

```bash
python3 scripts/build_ice_shop_business_docx.py
node scripts/build_ice_shop_teaching_deck.mjs
```
