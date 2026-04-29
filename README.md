# 試験工房の開き方

一番かんたんな方法は、同じフォルダにある `start-site.bat` をダブルクリックすることです。

## 方法1: ダブルクリックで開く

1. `start-site.bat` をダブルクリックします。
2. 少し待つとブラウザで `http://127.0.0.1:4173/` が開きます。
3. 使い終わったら、`School Test Site Server` という黒い画面を閉じます。

## 方法2: index.html を直接開く

1. `index.html` をダブルクリックします。
2. ブラウザで開ければそのまま使えます。

もし `index.html` を直接開いてうまく表示されないときは、`start-site.bat` のほうを使ってください。

## 方法3: PowerShell で開く

このフォルダで PowerShell を開いて、次を実行します。

```powershell
python -m http.server 4173
```

そのあとブラウザで次を開きます。

[http://127.0.0.1:4173/](http://127.0.0.1:4173/)

## ファイルの説明

- `index.html`: サイト本体
- `styles.css`: 見た目
- `script.js`: テスト作成の動き
- `start-site.bat`: ワンクリック起動用
