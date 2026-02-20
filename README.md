# tangenhairetsu-auto

Google スプレッドシート上で、Google Drive 配下の PDF（年間指導計画など）から
「活動時期・単元名・配当時数」を抽出し、4月〜3月の単元配列表を自動生成する Apps Script です。

Gemini API を使った PDF 解析を基本とし、失敗時は OCR（Drive Advanced Service）経由でテキスト抽出して再解析します。

## 主な機能

- Drive フォルダ配下（サブフォルダ含む）の PDF を再帰的に収集
- PDF から単元情報を JSON 形式で抽出（Gemini API）
- 解析失敗時のフォールバック（OCR → Gemini 再解析）
- スプレッドシートに以下を自動出力
  - `単元計画一覧`（全教科まとめ）
  - `計画_<教科名>`（教科別シート）
  - `抽出データ`（抽出結果の明細）
- `処理キャッシュ` シートで PDF 更新日時ベースの再処理抑制

## ファイル構成

- `コード.js`: メインロジック
- `appsscript.json`: マニフェスト（スコープ・Advanced Service 設定）
- `.clasp.json`: clasp 連携設定

## 前提条件

- Google アカウント
- Google スプレッドシートを利用可能
- Gemini API キーを取得済み（取得先: https://aistudio.google.com/api-keys）
- `Drive API`（Advanced Google Services）を有効化済み

## セットアップ

### 1. スクリプトを配置

`clasp` を使う場合:

```bash
npm i -g @google/clasp
clasp login
clasp pull   # 既存プロジェクトを取得する場合
clasp push   # ローカル変更を反映する場合
```

### 2. Script Properties を設定

Apps Script エディタで以下を設定します。

- キー: `GEMINI_API_KEY`
- 値: あなたの Gemini API キー

### 3. Advanced Service を有効化

Apps Script エディタのサービス設定で `Drive API` を有効にします。

このプロジェクトの `appsscript.json` では `Drive v2` を利用する前提です。

## 使い方

参照スプレッドシート（コピーして利用）:  
https://docs.google.com/spreadsheets/d/1hiUPlonFgph2jL6dh1D8mDEOCyEijx5CSppf4j3Shhs/copy

1. 対象スプレッドシートをコピーして開く
2. 任意のシート `A1` に、対象 Drive フォルダの URL またはフォルダ ID を入力
3. メニュー `単元配列表` → `Drive PDFから生成` を実行
4. 出力シートを確認

キャッシュを消して全件再処理したい場合:

- メニュー `単元配列表` → `処理キャッシュをクリア`

## 出力仕様

### `単元計画一覧`

- 列: `教科`, `PDFファイル`, `4月` ... `3月`
- セル内には `単元名（時数）` を改行区切りで出力

### `計画_<教科名>`

- 教科別に `PDFファイル`, `4月` ... `3月` を出力

### `抽出データ`

- 列: `sourceFile`, `subject`, `period`, `months`, `unit`, `hours`
- Gemini の抽出結果を明細として保持

### `処理キャッシュ`

- 内部利用シート（通常は非表示）
- `fileId`, `updatedMs` などを保存し、未更新 PDF の再解析を回避

## 注意事項

- PDF のレイアウト品質によって抽出精度は変動します。
- Gemini API 呼び出しには課金・レート制限があります。
- OCR フォールバック時は一時的な Google ドキュメントを作成し、処理後にゴミ箱へ移動します。
- 教科名推定はファイル名ベースの簡易判定（`算数` / `社会`）です。必要に応じてロジックを拡張してください。

## エラー時の確認ポイント

- `GEMINI_API_KEY` が Script Properties に設定されているか
- `A1` に有効なフォルダ URL / ID が入っているか
- 指定フォルダ配下に PDF が存在するか
- Drive API（Advanced Service）が有効化されているか
- 実行ユーザーに対象フォルダ・ファイルへのアクセス権があるか
