# token-usage-tracker-light

AIコーディングエージェント（**Claude Code** / **Codex CLI** / **Cline**）が残すローカルログを
解析し、リポジトリ・セッション・モデル・ツール別のトークン消費とコストを集計して、
**自己完結 HTML レポート（`report.html`）** を生成する軽量ツールです。

- データ収集は **PowerShell スクリプト 1 本**（`collect.ps1`）のみ。
- 生成された `report.html` は **ダブルクリック / ドラッグ&ドロップ** でブラウザ表示できます
  （`file://`、**Web サーバ不要・外部リクエストなし・オフライン動作**）。
- 集計は PowerShell 側で完了し、結果を JSON として HTML に埋め込み、少量のインライン JS が
  タブ・列ソート・コストバー・サブエージェント表示トグルを描画します。

元の重量級ツール（Python + SQLite + Streamlit）の代替として、毎回スクリプトを実行して
最新の `report.html` を再生成する運用を想定しています。

## 使い方

```powershell
# 既定パスを走査して report.html を生成
pwsh ./collect.ps1

# パスや出力先・単価を指定
pwsh ./collect.ps1 -ClaudeRoot D:\logs -OutFile usage.html -PricingPath .\pricing.json

# Windows PowerShell 5.1 の場合
powershell -ExecutionPolicy Bypass -File collect.ps1
```

生成された `report.html` をダブルクリックすれば、そのまま閲覧できます。

## 引数

| 引数 | 既定値 | 説明 |
|------|--------|------|
| `-ClaudeRoot` | `$HOME/.claude/projects` | Claude Code のログルート |
| `-CodexRoot` | `$HOME/.codex/sessions` | Codex CLI のログルート |
| `-ClineRoot` | `$env:APPDATA/Code/User/globalStorage/saoudrizwan.claude-dev/tasks` | Cline のタスクルート |
| `-PricingPath` | （後述の探索順） | 単価ファイル |
| `-OutFile` | `report.html` | 出力先 HTML |
| `-Tz` | `+09:00` | 日次バケットのタイムゾーン（表示用。集計内部は UTC） |
| `-IncludeSubagents` | `$true` | 初期表示でサブエージェントを含めるか（HTML 側でトグル切替可） |

## 対応ログ形式（裏取り済み）

- **Claude Code**: `~/.claude/projects/**/*.jsonl`。`type=="assistant"` かつ `message.usage`/`message.id`
  を持つ行を採用。`usage.cache_creation.{ephemeral_5m,ephemeral_1h}_input_tokens` があればそれを、
  無ければ `cache_creation_input_tokens` 全額を 5m として計上。パスに `subagents` を含むものは
  サブエージェント扱い。同一 `message.id` の重複行は **dedup**（output 最大の行を採用）。
- **Codex CLI**: `~/.codex/sessions/YYYY/MM/DD/rollout-*.jsonl`。`session_meta`→セッション/cwd、
  `turn_context`→model（直近値）、`event_msg` の `token_count` の `last_token_usage` を 1 呼び出しの
  使用量として採用。`input = input_tokens - cached_input_tokens`、`cache_read = cached_input_tokens`。
  **`output_tokens` は `reasoning_output_tokens` を内包**するため、コストは `output_tokens` のみで
  計算し reasoning は二重計上しません（表示用に保持）。
- **Cline**: `<globalStorage>/saoudrizwan.claude-dev/tasks/<id>/`。`ui_messages.json` の
  `say=="api_req_started"` の `text`(JSON文字列) から `tokensIn/tokensOut/cacheReads/cacheWrites` を抽出。
  `cacheWrites` は 5m キャッシュ単価で計上。モデルは `task_metadata.json` →
  `api_conversation_history.json` の正規表現抽出の順で解決し、取得できなければ「未割当」。

## 単価設定 `pricing.json`

1M トークンあたりの USD 単価。読み込み優先順（最初に存在するもの）:

1. `-PricingPath <file>`
2. `$HOME/.tokentracker/pricing.json`（利用者の上書き用）
3. スクリプト同梱の `pricing.json`（既定）

スキーマ:

```json
{
  "meta": { "effective_date": "2026-06-15" },
  "aliases": { "my-opus": "claude-opus-4-8" },
  "models": {
    "claude-opus-4-8": { "input": 5.0, "output": 25.0, "cache_write_1h": 10.0, "cache_write_5m": 6.25, "cache_read": 0.5 }
  }
}
```

- `aliases`: デプロイ名 → 正規モデルID の対応。
- モデル解決: `aliases` 適用 → 単価表に無ければ末尾の日付サフィックス `-YYYYMMDD` を除去して再検索。
- 単価未登録のモデルはコストを **「未割当トークン」** として別計上します（0 円化しません）。

### 単価の更新

知りたいのは定価ではなく実課金レート（例: Azure Foundry）であることが多いため、`-PricingPath` か
`$HOME/.tokentracker/pricing.json` で上書きしてください。値を更新したら `meta.effective_date` も
更新します。この日付はレポート上部に「単価の有効日」として表示され、古い単価での誤集計に
気づけるようにしています。

## テスト

[Pester](https://pester.dev) v5 を使用します。

```powershell
Import-Module Pester -MinimumVersion 5.0.0
Invoke-Pester ./tests
```

`tests/fixtures/` の固定サンプルログに対する期待値（トークン・コスト・dedup・未割当）を
アサートします。GitHub Actions（`.github/workflows/test.yml`）でも自動実行します。

## 制限事項

- 主対象は **Windows（PowerShell 5.1 / 7）**。Cline の `globalStorage` は OS 依存のため、
  Windows 以外では `-ClineRoot` で明示指定してください。
- 永続化（DB）や Web ダッシュボードは持ちません。毎回スクリプトを実行して再生成します。
- `server_tool_use`（Web 検索/取得などの件数課金）はコスト計算対象外です。
