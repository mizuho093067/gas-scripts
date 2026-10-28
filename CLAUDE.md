# CLAUDE.md

このファイルは、Claude Code がこのリポジトリで作業する際のガイドです。

## プロジェクト概要

- **プロジェクト名**: gas-scripts
- **内容**: Google Apps Script（GAS）のスクリプト集
- **技術スタック**: Google Apps Script（JavaScript / V8 ランタイム）

## GitHubリポジトリ

https://github.com/mizuho093067/gas-scripts

## コミュニケーション

- 返答は**必ず日本語**で行うこと。
- コード内のコメントも日本語で記述する。

## Git運用ルール

- **コードを変更するたびに、必ずコミットして GitHub にプッシュすること。**
  - 1つの作業（機能追加・修正など）が終わったら、まとめて放置せずにその都度プッシュする。
  - 手順:
    ```
    git add <変更したファイル>
    git commit -m "変更内容の説明"
    git push origin main
    ```
- ブランチは `main` を使用する。
- コミットメッセージは**日本語**で、何を変更したかが分かるように簡潔に書く。
  - 例: `スプレッドシートの集計スクリプトを追加`、`メール送信時の宛先エラーを修正`
- プッシュに失敗した場合は、エラー内容を報告し、強制プッシュ（`--force`）は行わないこと。
- 不要なファイル（テスト用の一時ファイルなど）はコミットしない。
- API キー・パスワード・スプレッドシート ID などの機密情報はコミットしない。

## ファイル構成

```
gas-scripts/
├── CLAUDE.md               # このファイル
└── summary_dashboard.gs    # 売上データの月次集計・サマリー書き込み・棒グラフ作成（毎朝9時に自動実行）
```

- スクリプトを追加したら、上の構成図も更新する。

## 開発・実行方法

- スクリプトは Google Apps Script エディタ（https://script.google.com）に貼り付けて実行する。
- 外部ライブラリは原則使わず、GAS 標準のサービス（`SpreadsheetApp`、`GmailApp` など）で実装する。

## コーディング規約

- JavaScript は `const` / `let` を使用し、`var` は使わない。
- 変数名・関数名はキャメルケース（例: `sheetName`, `sendReport()`）。
- 各関数の先頭に、何をする関数かを日本語のコメントで書く。
- スプレッドシート ID などの設定値はコードに直接書かず、スクリプトプロパティ（`PropertiesService`）から読み込む。

```js
// 設定値の読み込み例
const sheetId = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
```

## 注意事項

- GAS には実行時間の上限（1回あたり約6分）があるため、大量データはまとめて読み書きする（`getValues()` / `setValues()`）。
- メール送信やデータ削除など取り消しにくい処理は、実行前にユーザーに確認する。
