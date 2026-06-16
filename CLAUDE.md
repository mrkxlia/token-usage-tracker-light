# CLAUDE.md

このリポジトリで作業する際のガイド。実装の不変条件（壊すと回帰になる前提）と
設計意図を、ソースに散在するコメントから集約したもの。

## 概要

AIコーディングエージェント（Claude Code / Codex CLI / Cline）のローカルログを解析し、
リポジトリ・セッション・モデル・ツール・日次別のトークン消費とコストを集計して、
**自己完結 HTML レポート（`report.html`）** を生成する軽量ツール。
元の重量級ツール（Python + SQLite + Streamlit）の代替で、毎回スクリプトを実行して
最新の `report.html` を再生成する運用。

## 構成

- `collect.ps1` — データ収集〜集計〜HTML生成までを行う唯一の実行スクリプト。
- `pricing.json` — 同梱の既定単価表（1M トークンあたり USD）。
- `tests/collect.Tests.ps1` — Pester v5 テスト。`tests/fixtures/` の固定ログに対する期待値を検証。
- `.github/workflows/test.yml` — CI（push/PR で Pester 実行）。

## アーキテクチャ（`collect.ps1` の処理フロー）

```
パーサ（ツール別）          → Read-ClaudeLogs / Read-CodexLogs / Read-ClineLogs
  └ 共通レコード             → New-UsageRecord（source/model/input/output/cache_* など13フィールド）
dedup                       → Merge-UsageRecords（複合キー source|message_id、衝突時は output 最大を採用）
集計                        → New-AggRow（グループ内合算）→ Get-AxisAggregate（軸別グループ化）→ New-Summary
レポート組立                 → New-Dataset（5軸 + summary）→ Build-Report（include / exclude の2データセット）
HTML 生成                    → New-ReportHtml（CSS/JS/HTML テンプレート + JSON データアイランド埋め込み）
エントリポイント             → Invoke-Main（既定パス解決・出力書込み・サマリー表示）
```

- ドットソース（`. ./collect.ps1`）時は関数定義のみ。直接実行時のみ `Invoke-Main` が走る
  （`$MyInvocation.InvocationName -ne '.'` で分岐）。テストはこの仕組みで関数を直接呼ぶ。

## 不変条件（変更時は要注意・テストが依存）

1. **単一ファイル配布** — 収集ロジックは `collect.ps1` 1 本（＋同梱 `pricing.json`）。
   外部テンプレートや追加モジュールに分割しない方針。
2. **自己完結 HTML** — 生成 `report.html` は `file://` で動作。外部リクエスト・Webサーバ不要。
   CSS/JS はインライン、データは `<script type="application/json" id="data">` のデータアイランド。
3. **BOM 無し UTF-8 出力** — `file://` での文字化け回避（`Invoke-Main` の WriteAllText）。
4. **データアイランドのエスケープ** — 埋め込み JSON は `</` → `<\/`、`<!--` → `<\!--` に無害化
   （`<script>` 破壊防止。`JSON.parse` 時に復元される）。`charset=utf-8` を宣言。
5. **コスト規約**
   - Codex の `output_tokens` は `reasoning_output_tokens` を内包（OpenAI Responses API 仕様）。
     コストは `output_tokens` のみで計算し、reasoning は表示用に保持して**二重計上しない**。
   - 単価未登録のモデルはコストを **「未割当トークン」** として別計上（**0 円化しない**）。
   - Claude の cache_creation は `usage.cache_creation.{ephemeral_5m,ephemeral_1h}_input_tokens`
     があればそれを、無ければ `cache_creation_input_tokens` 全額を 5m 扱い。
   - Cline の `cacheWrites` は TTL 不明のため 5m 単価で計上。
6. **`SkippedLineCount` のセマンティクス** — パース不能な **1 行 / 1 レコード候補** の数のみを数える。
   ファイル全体のパース失敗（例: Cline の `ui_messages.json`）は「行」ではないため加算しない。
7. **PowerShell 5.1 / 7 両対応** — `Get-Prop` 等で null 差異を吸収。`Set-StrictMode -Version Latest`。

## 単価（`pricing.json`）

1M トークンあたり USD。読み込み優先順（最初に存在するもの）:

1. `-PricingPath <file>`
2. `$HOME/.tokentracker/pricing.json`（利用者の上書き用）
3. スクリプト同梱の `pricing.json`

- `aliases`: デプロイ名 → 正規モデルID。
- モデル解決: alias 適用 → 単価表に無ければ末尾の日付サフィックス `-YYYYMMDD` を除去して再検索。
  解決結果（未登録の `$null` 含む）は `Pricing.cache` にメモ化される。
- 単価を更新したら **`meta.effective_date` も更新**する（レポート上部に「単価の有効日」として表示）。

## 開発・テスト

```powershell
Import-Module Pester -MinimumVersion 5.0.0
Invoke-Pester ./tests
```

- `tests/fixtures/` は**固定・以後不変**。期待値を変える場合はテスト側の期待値も連動して更新する。
- レポートを目視確認するには `pwsh ./collect.ps1` を実行し、生成 `report.html` を `file://` で開く。

## 将来検討（現状スコープ外）

実装の負債というより設計トレードオフのため見送っている項目:

- 埋め込み HTML/CSS/JS の外部テンプレート分割（「単一ファイル配布」とトレードオフ）。
- `Read-*Logs` / `Get-AxisAggregate` / `New-AggRow` の共通基盤化・汎用化。
- 入力パス検証・`-Verbose` でのスキップ理由ログ。
- 埋め込み JS のモジュール化・フォーマッタ共通化。
- レコード単位でのコスト事前計算（集計の単一パス化）— 個人ログ規模では効果小。
