# 文書送付書・受領書自動でつくる君

法律事務所向けの文書送付書・受領書自動生成ツールです。
相手方から送られてきたFAX/PDFの受領書部分にOCRで文字位置を検出し、受領日・記名・押印を自動で書き込みます。

## 機能

### 受領書自動生成
- PDFをドラッグ&ドロップするだけで受領書を自動生成
- 複数ファイルの一括処理に対応
- 様々な受領書書式に自動対応:
  - 送付書+受領書の複合ページ（1ページ目が上下分割）
  - 独立型の受領書ページ
  - 「行」パターン（二重打消し線+「先生」追記）
  - 「殿」「宛」パターン（「行」処理をスキップ）
- 受領書ページのみを抽出して1ページPDFとして出力

### 文書送付書自動生成
- 文書送付書のテンプレートからPDFを生成

## 必要環境

- **Windows 10/11**
- **Node.js** v18以上
- **Tesseract OCR**（日本語言語パック `jpn` 含む）
  - [ダウンロード](https://github.com/UB-Mannheim/tesseract/wiki)
  - インストール時に「Japanese」言語パックを選択
- **Ghostscript** 10.x
  - [ダウンロード](https://ghostscript.com/releases/gsdnld.html)
- **游明朝フォント**（`C:\Windows\Fonts\yumin.ttf`）
  - Windows標準搭載

## セットアップ

```bash
# 1. リポジトリをクローン
git clone https://github.com/YOUR_USERNAME/文書送付書・受領書自動でつくる君.git
cd 文書送付書・受領書自動でつくる君

# 2. 依存パッケージをインストール
npm install

# 3. 設定ファイルを作成
cp config.sample.json config.json
# config.json を編集して事務所情報を入力

# 4. 印鑑画像を配置（任意）
mkdir seal
# seal/stamp.png に印鑑画像（PNG）を配置
# 印鑑画像がない場合は「㊞」で代替されます

# 5. サーバー起動
node web/server.js
```

ブラウザで `http://localhost:3000` を開いてください。

## 設定ファイル（config.json）

```json
{
  "officeName": "○○法律事務所",
  "lawyerNames": ["山田", "山田太郎"],
  "faxNumbers": ["03-1234-5678"],
  "port": 3000
}
```

| キー | 説明 |
|------|------|
| `officeName` | 事務所名 |
| `lawyerNames` | 弁護士名のリスト（署名に使用） |
| `faxNumbers` | FAX番号のリスト |
| `port` | サーバーのポート番号 |

## 使い方

1. ブラウザで `http://localhost:3000` を開く
2. 「受領書つくる君」モードを選択
3. PDFファイルをドラッグ&ドロップ（複数可）
4. 受領日・署名者を確認して「生成」ボタンを押す
5. 生成されたPDFがダウンロードされます

## ディレクトリ構成

```
├── config.json          # 事務所設定（要作成）
├── config.sample.json   # 設定ファイルのテンプレート
├── package.json
├── seal/                # 印鑑画像フォルダ
│   └── stamp.png        # 印鑑画像（任意）
├── system/
│   ├── receipt.js       # 受領書生成ロジック
│   └── generate.js      # 文書送付書生成ロジック
├── web/
│   ├── server.js        # HTTPサーバー
│   └── public/
│       ├── index.html   # フロントエンド
│       ├── app.js       # フロントエンドJS
│       └── style.css    # スタイル
└── output/              # 生成されたPDFの出力先
```

## ライセンス

MIT License
