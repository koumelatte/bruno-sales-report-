# 売上レポート自動集計ツール（Bruno × GAS 連携）

> API テストツール「Bruno」と Google Apps Script を連携させた、売上データの自動集計・レポート生成ツール

[![GAS](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=flat-square&logo=google&logoColor=white)](https://www.google.com/script/start/)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)

---

## 📌 概要

売上データの取得から集計・レポート出力までを **GAS と REST API** で自動化するツールです。Bruno を使ったAPIリクエスト管理と、GAS による自動集計・出力を組み合わせています。

---

## ✨ 主な機能

- **REST API 経由でデータ取得** — Bruno で定義したリクエストでデータを取得
- **GAS による自動集計** — 取得データをスプレッドシートに自動整形・集計
- **レポート自動出力** — 日次・週次の集計結果をシート出力
- **手作業ゼロ** — トリガー設定で全自動実行

---

## 🛠️ 使用技術

| 技術 | 用途 |
|---|---|
| Google Apps Script | データ処理・集計・出力 |
| Bruno | API リクエスト管理・テスト |
| Google スプレッドシート | レポート出力先 |
| REST API | 外部データ取得 |

---

## 📄 ライセンス

MIT License

---

<div align="right">

**Kei Assist** — 作業を仕組みに変える

</div>
