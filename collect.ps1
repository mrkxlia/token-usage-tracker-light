#requires -Version 5.1
<#
.SYNOPSIS
    AIエージェント（Claude Code / Codex CLI / Cline）のローカルログを解析し、
    トークン/コストを集計した自己完結 HTML レポート（report.html）を生成する。

.DESCRIPTION
    データ収集～集計～HTML生成までを1スクリプトで完結。生成HTMLは file:// で
    ダブルクリック/ドラッグ&ドロップ表示でき、外部リソース・Webサーバ不要。
    インラインJSがデータアイランド(JSON)を読んでタブ/ソート/バー/サブエージェント
    トグルを描画する。

    重要な単価仕様:
      - Codex の output_tokens は reasoning_output_tokens を内包する（OpenAI Responses
        API 仕様。ccusage も同様）。したがってコストは output_tokens のみで計算し、
        reasoning は表示用に保持するだけで課金へ加算しない。
      - Claude の cache_creation は usage.cache_creation.{ephemeral_5m,ephemeral_1h}_
        input_tokens があればそれを、無ければ cache_creation_input_tokens 全額を 5m 扱い。

.EXAMPLE
    pwsh ./collect.ps1
    pwsh ./collect.ps1 -ClaudeRoot D:\logs -OutFile usage.html -PricingPath .\pricing.json
#>
[CmdletBinding()]
param(
    [string]$ClaudeRoot,
    [string]$CodexRoot,
    [string]$ClineRoot,
    [string]$PricingPath,
    [string]$OutFile = 'report.html',
    [string]$Tz = '+09:00',
    [bool]$IncludeSubagents = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 壊れ行スキップ数（Invoke-Main でリセット、サマリーに表示）
$script:SkippedLineCount = 0

# ---------------------------------------------------------------------------
# 小ユーティリティ
# ---------------------------------------------------------------------------

# PSCustomObject/Hashtable から安全にプロパティを取り出す（5.1/7 の null 差異対策）
function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [hashtable]) {
        if ($Object.ContainsKey($Name)) {
            $v = $Object[$Name]
            if ($null -eq $v) { return $Default }
            return $v
        }
        return $Default
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

function ConvertTo-Iso8601Utc {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    try {
        $dto = [datetimeoffset]::Parse(
            $Value, [cultureinfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal)
        return $dto.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [cultureinfo]::InvariantCulture)
    } catch {
        return $Value
    }
}

function Get-DayBucket {
    param([string]$Iso, [string]$Tz)
    if ([string]::IsNullOrWhiteSpace($Iso)) { return '(no date)' }
    try {
        $dto = [datetimeoffset]::Parse(
            $Iso, [cultureinfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal)
        $sign = 1; $t = $Tz
        if ($t.StartsWith('-')) { $sign = -1; $t = $t.Substring(1) }
        elseif ($t.StartsWith('+')) { $t = $t.Substring(1) }
        $span = [timespan]::Parse($t)
        if ($sign -lt 0) { $span = $span.Negate() }
        return $dto.ToOffset($span).ToString('yyyy-MM-dd')
    } catch {
        return '(no date)'
    }
}

function ConvertFrom-JsonLine {
    param([string]$Line)
    return ($Line | ConvertFrom-Json -Depth 64)
}

function New-UsageRecord {
    [pscustomobject]@{
        source            = ''
        message_id        = ''
        session_id        = ''
        model             = ''
        timestamp         = ''
        input             = 0
        output            = 0
        reasoning         = 0
        cache_creation_1h = 0
        cache_creation_5m = 0
        cache_read        = 0
        repo_path         = ''
        git_branch        = ''
        is_subagent       = $false
    }
}

# ---------------------------------------------------------------------------
# 単価
# ---------------------------------------------------------------------------

function Import-PricingTable {
    param([string]$PricingPath)

    $candidates = @()
    if ($PricingPath) { $candidates += $PricingPath }
    $candidates += (Join-Path $HOME '.tokentracker/pricing.json')
    $candidates += (Join-Path $PSScriptRoot 'pricing.json')

    $path = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $path) { throw "pricing.json が見つかりません（探索: $($candidates -join ', ')）" }

    $raw = Get-Content -Path $path -Raw -Encoding UTF8
    $obj = $raw | ConvertFrom-Json -Depth 64

    $aliases = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    $aProp = $obj.PSObject.Properties['aliases']
    if ($aProp -and $aProp.Value) {
        foreach ($p in $aProp.Value.PSObject.Properties) { $aliases[$p.Name] = $p.Value }
    }

    $models = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $obj.models.PSObject.Properties) { $models[$p.Name] = $p.Value }

    [pscustomobject]@{
        effective_date = (Get-Prop $obj.meta 'effective_date' '')
        source         = (Get-Prop $obj.meta 'source' '')
        aliases        = $aliases
        models         = $models
        path           = $path
    }
}

function Resolve-ModelPricing {
    param([string]$Model, $Pricing)
    if ([string]::IsNullOrWhiteSpace($Model)) { return $null }
    $m = $Model
    if ($Pricing.aliases.ContainsKey($m)) { $m = [string]$Pricing.aliases[$m] }
    if ($Pricing.models.ContainsKey($m)) { return $Pricing.models[$m] }
    $stripped = [regex]::Replace($m, '-\d{8}$', '')
    if ($stripped -ne $m -and $Pricing.models.ContainsKey($stripped)) { return $Pricing.models[$stripped] }
    return $null
}

function Get-UsageCost {
    param($Record, $Pricing)
    $price = Resolve-ModelPricing -Model $Record.model -Pricing $Pricing
    if ($null -eq $price) { return $null }   # 未登録 → 未割当（0円化しない）
    # reasoning は output_tokens に内包のため加算しない
    $cost = (
        [double]$Record.input             * [double](Get-Prop $price 'input' 0) +
        [double]$Record.output            * [double](Get-Prop $price 'output' 0) +
        [double]$Record.cache_creation_1h * [double](Get-Prop $price 'cache_write_1h' 0) +
        [double]$Record.cache_creation_5m * [double](Get-Prop $price 'cache_write_5m' 0) +
        [double]$Record.cache_read        * [double](Get-Prop $price 'cache_read' 0)
    ) / 1e6
    return $cost
}

# ---------------------------------------------------------------------------
# パーサ
# ---------------------------------------------------------------------------

function Read-ClaudeLogs {
    param([string]$Root)
    if (-not $Root -or -not (Test-Path $Root)) { return }
    foreach ($file in Get-ChildItem -Path $Root -Recurse -File -Filter '*.jsonl' -ErrorAction SilentlyContinue) {
        $isSub = ($file.FullName -match '[\\/]subagents[\\/]')
        foreach ($line in [System.IO.File]::ReadAllLines($file.FullName)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try { $obj = ConvertFrom-JsonLine $line } catch { $script:SkippedLineCount++; continue }
            if ((Get-Prop $obj 'type') -ne 'assistant') { continue }
            $msg = Get-Prop $obj 'message'
            $usage = Get-Prop $msg 'usage'
            $id = Get-Prop $msg 'id' ''
            if ($null -eq $usage -or [string]::IsNullOrEmpty($id)) { continue }

            $rec = New-UsageRecord
            $rec.source     = 'claude'
            $rec.message_id = $id
            $rec.session_id = Get-Prop $obj 'sessionId' ''
            $rec.model      = Get-Prop $msg 'model' ''
            $rec.timestamp  = ConvertTo-Iso8601Utc (Get-Prop $obj 'timestamp' '')
            $rec.input      = [int](Get-Prop $usage 'input_tokens' 0)
            $rec.output     = [int](Get-Prop $usage 'output_tokens' 0)
            $rec.cache_read = [int](Get-Prop $usage 'cache_read_input_tokens' 0)
            $cc = Get-Prop $usage 'cache_creation'
            if ($null -ne $cc) {
                $rec.cache_creation_1h = [int](Get-Prop $cc 'ephemeral_1h_input_tokens' 0)
                $rec.cache_creation_5m = [int](Get-Prop $cc 'ephemeral_5m_input_tokens' 0)
            } else {
                $rec.cache_creation_5m = [int](Get-Prop $usage 'cache_creation_input_tokens' 0)
            }
            $rec.repo_path   = Get-Prop $obj 'cwd' ''
            $rec.git_branch  = Get-Prop $obj 'gitBranch' ''
            $rec.is_subagent = $isSub
            Write-Output $rec
        }
    }
}

function Read-CodexLogs {
    param([string]$Root)
    if (-not $Root -or -not (Test-Path $Root)) { return }
    foreach ($file in Get-ChildItem -Path $Root -Recurse -File -Filter 'rollout-*.jsonl' -ErrorAction SilentlyContinue) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $sessionId = ''; $cwd = ''; $model = ''; $tokenIndex = 0
        foreach ($line in [System.IO.File]::ReadAllLines($file.FullName)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try { $obj = ConvertFrom-JsonLine $line } catch { $script:SkippedLineCount++; continue }
            $type = Get-Prop $obj 'type'
            $payload = Get-Prop $obj 'payload'
            switch ($type) {
                'session_meta' {
                    $sessionId = Get-Prop $payload 'id' $sessionId
                    $cwd       = Get-Prop $payload 'cwd' $cwd
                }
                { $_ -eq 'turn_context' -or $_ -eq 'turnContext' } {
                    $model = Get-Prop $payload 'model' $model
                }
                'event_msg' {
                    if ((Get-Prop $payload 'type') -ne 'token_count') { break }
                    $info = Get-Prop $payload 'info'
                    $u = Get-Prop $info 'last_token_usage'
                    if ($null -eq $u) { $u = Get-Prop $info 'total_token_usage' }
                    if ($null -eq $u) { break }

                    $inTok  = [int](Get-Prop $u 'input_tokens' 0)
                    $cached = [int](Get-Prop $u 'cached_input_tokens' 0)
                    $rec = New-UsageRecord
                    $rec.source     = 'codex'
                    $rec.message_id = "$base#$tokenIndex"
                    $rec.session_id = $sessionId
                    $rec.model      = $model
                    $rec.timestamp  = ConvertTo-Iso8601Utc (Get-Prop $obj 'timestamp' '')
                    $rec.input      = [math]::Max(0, $inTok - $cached)   # cached は input の内訳→控除
                    $rec.cache_read = $cached
                    $rec.output     = [int](Get-Prop $u 'output_tokens' 0)
                    $rec.reasoning  = [int](Get-Prop $u 'reasoning_output_tokens' 0)
                    $rec.repo_path  = $cwd
                    $tokenIndex++
                    Write-Output $rec
                }
            }
        }
    }
}

function Read-ClineLogs {
    param([string]$Root)
    if (-not $Root -or -not (Test-Path $Root)) { return }
    foreach ($taskDir in Get-ChildItem -Path $Root -Directory -ErrorAction SilentlyContinue) {
        $taskId  = $taskDir.Name
        $uiPath  = Join-Path $taskDir.FullName 'ui_messages.json'
        if (-not (Test-Path $uiPath)) { continue }
        $metaPath = Join-Path $taskDir.FullName 'task_metadata.json'
        $histPath = Join-Path $taskDir.FullName 'api_conversation_history.json'

        $meta = $null
        if (Test-Path $metaPath) {
            try { $meta = (Get-Content $metaPath -Raw -Encoding UTF8) | ConvertFrom-Json -Depth 64 } catch { $meta = $null }
        }
        $repo = Get-Prop $meta 'cwdOnTaskInitialization' ''
        if ([string]::IsNullOrEmpty($repo)) { $repo = Get-Prop $meta 'shadowGitConfigWorkTree' '' }

        $model = Get-Prop $meta 'model' ''
        if ([string]::IsNullOrEmpty($model)) { $model = Get-Prop $meta 'apiModelId' '' }
        if ([string]::IsNullOrEmpty($model)) { $model = Get-Prop $meta 'modelId' '' }
        if ([string]::IsNullOrEmpty($model) -and (Test-Path $histPath)) {
            $hist = Get-Content $histPath -Raw -Encoding UTF8
            $m = [regex]::Match($hist, '\b(claude-[\w.-]+|gpt-[\w.-]+|o\d[\w.-]*)\b')
            if ($m.Success) { $model = $m.Value }
        }

        try { $ui = (Get-Content $uiPath -Raw -Encoding UTF8) | ConvertFrom-Json -Depth 64 } catch { $script:SkippedLineCount++; continue }
        $idx = 0
        foreach ($e in @($ui)) {
            if ((Get-Prop $e 'say') -ne 'api_req_started') { continue }
            $text = Get-Prop $e 'text' ''
            try { $t = $text | ConvertFrom-Json -Depth 64 } catch { $script:SkippedLineCount++; continue }

            $rec = New-UsageRecord
            $rec.source            = 'cline'
            $rec.message_id        = "$taskId#$idx"
            $rec.session_id        = $taskId
            $rec.model             = $model
            $rec.input             = [int](Get-Prop $t 'tokensIn' 0)
            $rec.output            = [int](Get-Prop $t 'tokensOut' 0)
            $rec.cache_read        = [int](Get-Prop $t 'cacheReads' 0)
            $rec.cache_creation_5m = [int](Get-Prop $t 'cacheWrites' 0)   # TTL不明→5m仮定
            $rec.repo_path         = $repo
            $ts = Get-Prop $e 'ts' 0
            if ($ts) {
                $rec.timestamp = ([datetimeoffset]::FromUnixTimeMilliseconds([long]$ts)).UtcDateTime.ToString(
                    'yyyy-MM-ddTHH:mm:ssZ', [cultureinfo]::InvariantCulture)
            }
            $idx++
            Write-Output $rec
        }
    }
}

# ---------------------------------------------------------------------------
# dedup / 集計
# ---------------------------------------------------------------------------

function Merge-UsageRecords {
    param([object[]]$Records)
    $map = [ordered]@{}
    foreach ($r in @($Records)) {
        if ($null -eq $r) { continue }
        $key = "$($r.source)|$($r.message_id)"     # 複合キー（ツール間衝突防止）
        if (-not $map.Contains($key)) {
            $map[$key] = $r
        } elseif ([int]$r.output -gt [int]$map[$key].output) {
            $map[$key] = $r                          # 衝突時は output 最大の行をまるごと採用
        }
    }
    return @($map.Values)
}

function New-AggRow {
    param([string]$Key, [object[]]$Items, $Pricing)
    $inSum = 0; $outSum = 0; $crSum = 0; $ccSum = 0; $cost = 0.0; $hasCost = $false
    foreach ($r in $Items) {
        $inSum += [int]$r.input; $outSum += [int]$r.output; $crSum += [int]$r.cache_read
        $ccSum += ([int]$r.cache_creation_1h + [int]$r.cache_creation_5m)
        $c = Get-UsageCost -Record $r -Pricing $Pricing
        if ($null -ne $c) { $cost += $c; $hasCost = $true }
    }
    $n = @($Items).Count
    [pscustomobject]@{
        key            = $Key
        input          = $inSum
        output         = $outSum
        cache_read     = $crSum
        cache_creation = $ccSum
        tokens         = ($inSum + $outSum + $crSum + $ccSum)
        events         = $n
        cost           = $(if ($hasCost) { [math]::Round($cost, 6) } else { $null })
        out_in         = $(if ($inSum -gt 0) { [math]::Round($outSum / $inSum, 3) } else { $null })
        cache_rate     = $(if (($inSum + $crSum) -gt 0) { [math]::Round($crSum / ($inSum + $crSum), 3) } else { $null })
        cost_per_event = $(if ($hasCost -and $n -gt 0) { [math]::Round($cost / $n, 6) } else { $null })
        unassigned     = (-not $hasCost)
    }
}

function Get-AxisAggregate {
    param([object[]]$Records, [string]$Kind, $Pricing, [string]$Tz)
    $groups = [ordered]@{}
    foreach ($r in $Records) {
        switch ($Kind) {
            'model'   { $key = $(if ($r.model)     { $r.model }     else { '(unassigned)' }) }
            'session' { $key = $(if ($r.session_id) { $r.session_id } else { '(none)' }) }
            'repo'    { $key = $(if ($r.repo_path)  { $r.repo_path }  else { '(none)' }) }
            'tool'    { $key = $r.source }
            'daily'   { $key = Get-DayBucket $r.timestamp $Tz }
            default   { $key = '(none)' }
        }
        if (-not $groups.Contains($key)) { $groups[$key] = New-Object System.Collections.ArrayList }
        [void]$groups[$key].Add($r)
    }
    $rows = foreach ($k in $groups.Keys) { New-AggRow -Key $k -Items $groups[$k].ToArray() -Pricing $Pricing }
    # 既定ソート: コスト降順、未割当はトークン降順で後段
    return @($rows | Sort-Object `
        @{ Expression = { if ($_.unassigned) { 0 } else { 1 } }; Descending = $true }, `
        @{ Expression = { if ($_.unassigned) { $_.tokens } else { $_.cost } }; Descending = $true })
}

function New-Summary {
    param([object[]]$Records, $Pricing)
    $ti = 0; $to = 0; $tc = 0; $tcost = 0.0; $unTok = 0; $savings = 0.0
    $unModels = [ordered]@{}
    foreach ($r in $Records) {
        $tok = [int]$r.input + [int]$r.output + [int]$r.cache_read + [int]$r.cache_creation_1h + [int]$r.cache_creation_5m
        $c = Get-UsageCost -Record $r -Pricing $Pricing
        if ($null -eq $c) {
            $unTok += $tok
            $mk = $(if ($r.model) { $r.model } else { '(empty)' })
            $unModels[$mk] = $true
        } else {
            $tcost += $c
            $price = Resolve-ModelPricing -Model $r.model -Pricing $Pricing
            if ($price) {
                $sv = [int]$r.cache_read * ([double](Get-Prop $price 'input' 0) - [double](Get-Prop $price 'cache_read' 0)) / 1e6
                if ($sv -gt 0) { $savings += $sv }
            }
        }
        $ti += [int]$r.input; $to += [int]$r.output; $tc += [int]$r.cache_read
    }
    [pscustomobject]@{
        total_input       = $ti
        total_output      = $to
        total_cache       = $tc
        total_cost        = [math]::Round($tcost, 6)
        unassigned_tokens = $unTok
        event_count       = @($Records).Count
        estimated_savings = [math]::Round($savings, 6)
        unassigned_models = @($unModels.Keys)
        skipped_lines     = $script:SkippedLineCount
    }
}

function New-Dataset {
    param([object[]]$Records, $Pricing, [string]$Tz)
    [pscustomobject]@{
        summary = New-Summary -Records $Records -Pricing $Pricing
        axes    = [pscustomobject]@{
            model   = Get-AxisAggregate -Records $Records -Kind 'model'   -Pricing $Pricing -Tz $Tz
            session = Get-AxisAggregate -Records $Records -Kind 'session' -Pricing $Pricing -Tz $Tz
            repo    = Get-AxisAggregate -Records $Records -Kind 'repo'    -Pricing $Pricing -Tz $Tz
            daily   = Get-AxisAggregate -Records $Records -Kind 'daily'   -Pricing $Pricing -Tz $Tz
            tool    = Get-AxisAggregate -Records $Records -Kind 'tool'    -Pricing $Pricing -Tz $Tz
        }
    }
}

function Build-Report {
    param([object[]]$Records, $Pricing, [hashtable]$Roots = @{}, [string]$Tz = '+09:00', [bool]$IncludeSubagents = $true)
    $all     = @($Records)
    $noSub   = @($all | Where-Object { -not $_.is_subagent })
    $include = New-Dataset -Records $all   -Pricing $Pricing -Tz $Tz
    $exclude = New-Dataset -Records $noSub -Pricing $Pricing -Tz $Tz

    $rootList = foreach ($k in $Roots.Keys) { [pscustomobject]@{ label = $k; path = [string]$Roots[$k] } }

    [pscustomobject]@{
        meta = [pscustomobject]@{
            generated_at   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [cultureinfo]::InvariantCulture)
            effective_date = $Pricing.effective_date
            tz             = $Tz
            record_count   = $all.Count
            default_view   = $(if ($IncludeSubagents) { 'include' } else { 'exclude' })
            roots          = @($rootList)
        }
        summary  = $include.summary       # トップレベル（include基準）
        datasets = [pscustomobject]@{ include = $include; exclude = $exclude }
    }
}

# ---------------------------------------------------------------------------
# HTML 生成
# ---------------------------------------------------------------------------

function New-ReportHtml {
    param($Report)

    $json = $Report | ConvertTo-Json -Depth 12 -Compress
    # </script> と <!-- を無害化（データアイランド破壊防止）。JSON.parse 時に '\/' は '/' に戻る
    $json = $json.Replace('</', '<\/').Replace('<!--', '<\!--')

    $css = @'
:root{--bg:#0f1115;--card:#181b22;--fg:#e6e8eb;--muted:#9aa3ad;--line:#2a2f3a;--bar:#3b82f6;--bar2:#10b981;}
*{box-sizing:border-box}body{margin:0;font:14px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--fg)}
header{padding:16px 20px;border-bottom:1px solid var(--line)}
h1{font-size:18px;margin:0 0 4px}.meta{color:var(--muted);font-size:12px}
.wrap{padding:16px 20px;max-width:1200px;margin:0 auto}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px}
.card .k{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.04em}
.card .v{font-size:20px;font-weight:600;margin-top:4px;word-break:break-all}
.controls{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}
.tabs{display:flex;flex-wrap:wrap;gap:6px}
.tab,.toggle{background:var(--card);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:6px 12px;cursor:pointer;font:inherit}
.tab.active{background:var(--bar);border-color:var(--bar)}
.toggle.active{background:var(--bar2);border-color:var(--bar2)}
table{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:10px;overflow:hidden}
th,td{padding:8px 10px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap}
th:first-child,td:first-child{text-align:left;white-space:normal;word-break:break-all;max-width:380px}
thead th{position:sticky;top:0;background:#10131a;cursor:pointer;user-select:none;font-size:12px;color:var(--muted)}
thead th:hover{color:var(--fg)}
tbody tr:nth-child(odd){background:rgba(255,255,255,.02)}
tbody tr:hover{background:rgba(59,130,246,.12)}
tr.total{font-weight:700;background:rgba(255,255,255,.05)}
.barcell{position:relative;min-width:120px}
.bar{position:absolute;left:0;top:50%;transform:translateY(-50%);height:60%;background:rgba(59,130,246,.35);border-radius:4px;z-index:0}
.barcell span{position:relative;z-index:1}
.un{color:#f59e0b}.muted{color:var(--muted)}
details{margin-top:14px;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:10px}
summary{cursor:pointer;color:var(--muted)}
code{background:#10131a;padding:1px 5px;border-radius:4px}
'@

    $js = @'
const REPORT = JSON.parse(document.getElementById("data").textContent);
const AXES = [["model","モデル別"],["session","セッション別"],["repo","リポジトリ別"],["daily","日次"],["tool","ツール別"]];
const COLS = [
  ["key","名前","str"],["events","件数","num"],["input","input","num"],["output","output","num"],
  ["cache_read","cache_read","num"],["cache_creation","cache_write","num"],["cost","コスト$","cost"],
  ["out_in","out/in","num3"],["cache_rate","cache率","num3"],["cost_per_event","$/件","cost"]
];
let view = REPORT.meta.default_view || "include";
let tab = "model";
let sort = {col:"cost", dir:-1};

function esc(s){return String(s==null?"":s);}
function fmtNum(n){return (n==null)?"-":Number(n).toLocaleString();}
function fmtCost(n){return (n==null)?'<span class="un">未割当</span>':("$"+Number(n).toFixed(6));}
function fmt3(n){return (n==null)?"-":Number(n).toFixed(3);}

function renderSummary(){
  const s = REPORT.datasets[view].summary;
  const cards = [
    ["総コスト (USD)","$"+Number(s.total_cost).toFixed(4)],
    ["input トークン",fmtNum(s.total_input)],
    ["output トークン",fmtNum(s.total_output)],
    ["cache_read トークン",fmtNum(s.total_cache)],
    ["イベント数",fmtNum(s.event_count)],
    ["推定節約額 (USD)","$"+Number(s.estimated_savings).toFixed(4)],
    ["未割当トークン",fmtNum(s.unassigned_tokens)],
    ["スキップ壊れ行",fmtNum(s.skipped_lines)],
  ];
  let h = cards.map(c=>`<div class="card"><div class="k">${esc(c[0])}</div><div class="v">${c[1]}</div></div>`).join("");
  const um = (s.unassigned_models||[]).filter(m=>m&&m!=="(empty)");
  const umTxt = (s.unassigned_models&&s.unassigned_models.length)?s.unassigned_models.map(esc).join(", "):"なし";
  h += `<div class="card"><div class="k">未登録モデル</div><div class="v" style="font-size:13px">${umTxt}</div></div>`;
  document.getElementById("cards").innerHTML = h;
}

function rowsForTab(){
  let rows = (REPORT.datasets[view].axes[tab]||[]).slice();
  const c = sort.col, d = sort.dir;
  rows.sort((a,b)=>{
    let x=a[c], y=b[c];
    if(x==null&&y==null)return 0; if(x==null)return 1; if(y==null)return -1;
    if(typeof x==="string")return d*x.localeCompare(y);
    return d*(x-y);
  });
  return rows;
}

function renderTable(){
  const rows = rowsForTab();
  const maxCost = Math.max(1e-12,...rows.map(r=>r.cost||0));
  const tot = {events:0,input:0,output:0,cache_read:0,cache_creation:0,cost:0};
  rows.forEach(r=>{tot.events+=r.events;tot.input+=r.input;tot.output+=r.output;tot.cache_read+=r.cache_read;tot.cache_creation+=r.cache_creation;tot.cost+=(r.cost||0);});
  let thead = "<tr>"+COLS.map(c=>`<th data-c="${c[0]}">${esc(c[1])}${sort.col===c[0]?(sort.dir<0?" ▼":" ▲"):""}</th>`).join("")+"</tr>";
  let body = rows.map(r=>{
    const w = r.cost!=null?Math.round(100*r.cost/maxCost):0;
    return "<tr>"+
      `<td>${esc(r.key)}${r.unassigned?' <span class="un">●</span>':''}</td>`+
      `<td>${fmtNum(r.events)}</td><td>${fmtNum(r.input)}</td><td>${fmtNum(r.output)}</td>`+
      `<td>${fmtNum(r.cache_read)}</td><td>${fmtNum(r.cache_creation)}</td>`+
      `<td class="barcell"><div class="bar" style="width:${w}%"></div><span>${fmtCost(r.cost)}</span></td>`+
      `<td>${fmt3(r.out_in)}</td><td>${fmt3(r.cache_rate)}</td><td>${fmtCost(r.cost_per_event)}</td></tr>`;
  }).join("");
  const totOutIn = tot.input>0?(tot.output/tot.input):null;
  const totRate = (tot.input+tot.cache_read)>0?(tot.cache_read/(tot.input+tot.cache_read)):null;
  body += `<tr class="total"><td>合計</td><td>${fmtNum(tot.events)}</td><td>${fmtNum(tot.input)}</td><td>${fmtNum(tot.output)}</td><td>${fmtNum(tot.cache_read)}</td><td>${fmtNum(tot.cache_creation)}</td><td>${fmtCost(Number(tot.cost.toFixed(6)))}</td><td>${fmt3(totOutIn)}</td><td>${fmt3(totRate)}</td><td>-</td></tr>`;
  document.getElementById("table").innerHTML = `<table><thead>${thead}</thead><tbody>${body}</tbody></table>`;
  document.querySelectorAll("thead th").forEach(th=>th.onclick=()=>{
    const c=th.getAttribute("data-c");
    if(sort.col===c){sort.dir*=-1;}else{sort.col=c;sort.dir=(c==="key")?1:-1;}
    renderTable();
  });
}

function renderTabs(){
  document.getElementById("tabs").innerHTML = AXES.map(a=>`<button class="tab ${tab===a[0]?'active':''}" data-t="${a[0]}">${esc(a[1])}</button>`).join("");
  document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{tab=b.getAttribute("data-t");renderTabs();renderTable();});
}

function renderToggle(){
  const b=document.getElementById("subToggle");
  b.className="toggle "+(view==="include"?"active":"");
  b.textContent = view==="include"?"サブエージェント: 含む":"サブエージェント: 除く";
  b.onclick=()=>{view=(view==="include")?"exclude":"include";renderToggle();renderSummary();renderTable();};
}

renderToggle();renderSummary();renderTabs();renderTable();
'@

    $genAt = $Report.meta.generated_at
    $eff   = $Report.meta.effective_date
    $tz    = $Report.meta.tz
    $cnt   = $Report.meta.record_count
    $roots = ($Report.meta.roots | ForEach-Object { "$($_.label)=$($_.path)" }) -join ' / '
    $rootsHtml = [System.Web.HttpUtility]::HtmlEncode($roots)

    $html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Token Usage Report</title>
<style>$css</style>
</head>
<body>
<header>
  <h1>AIエージェント トークン消費レポート</h1>
  <div class="meta">生成: $genAt ・ 対象件数: $cnt ・ 単価 effective_date: $eff ・ TZ: $tz</div>
  <div class="meta">対象ルート: $rootsHtml</div>
</header>
<div class="wrap">
  <div id="cards" class="cards"></div>
  <div class="controls">
    <div id="tabs" class="tabs"></div>
    <button id="subToggle" class="toggle"></button>
  </div>
  <div id="table"></div>
  <details><summary>このレポートについて</summary>
    <p class="muted">PowerShell の <code>collect.ps1</code> がローカルログを解析・集計し、結果を JSON として
    埋め込んだ自己完結 HTML です。Webサーバ不要・外部リクエストなしで動作します。
    Codex の reasoning トークンは output に内包されるためコストへ二重加算していません。
    Cline の cacheWrites は 5m キャッシュ単価で計上しています。未登録モデルのコストは
    「未割当」として別計上しています。</p>
  </details>
</div>
<script type="application/json" id="data">$json</script>
<script>$js</script>
</body>
</html>
"@
    return $html
}

# HtmlEncode を使えるように（5.1/7 双方）
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

function Invoke-Main {
    param(
        [string]$ClaudeRoot,
        [string]$CodexRoot,
        [string]$ClineRoot,
        [string]$PricingPath,
        [string]$OutFile = 'report.html',
        [string]$Tz = '+09:00',
        [bool]$IncludeSubagents = $true
    )
    $script:SkippedLineCount = 0

    if (-not $ClaudeRoot) { $ClaudeRoot = Join-Path $HOME '.claude/projects' }
    if (-not $CodexRoot)  { $CodexRoot  = Join-Path $HOME '.codex/sessions' }
    if (-not $ClineRoot -and $env:APPDATA) {
        $ClineRoot = Join-Path $env:APPDATA 'Code/User/globalStorage/saoudrizwan.claude-dev/tasks'
    }

    $pricing = Import-PricingTable -PricingPath $PricingPath

    $records = @()
    $records += @(Read-ClaudeLogs -Root $ClaudeRoot)
    $records += @(Read-CodexLogs  -Root $CodexRoot)
    $records += @(Read-ClineLogs  -Root $ClineRoot)

    $merged = @(Merge-UsageRecords -Records $records)

    $roots = @{}
    if ($ClaudeRoot) { $roots['Claude'] = $ClaudeRoot }
    if ($CodexRoot)  { $roots['Codex']  = $CodexRoot }
    if ($ClineRoot)  { $roots['Cline']  = $ClineRoot }

    $report = Build-Report -Records $merged -Pricing $pricing -Roots $roots -Tz $Tz -IncludeSubagents $IncludeSubagents
    $html   = New-ReportHtml -Report $report

    # BOM無し UTF-8 で出力（file:// での文字化け回避）
    $full = $OutFile
    if (-not [System.IO.Path]::IsPathRooted($full)) { $full = Join-Path (Get-Location).Path $full }
    [System.IO.File]::WriteAllText($full, $html, (New-Object System.Text.UTF8Encoding($false)))

    Write-Host "生成しました: $full"
    Write-Host ("  レコード: {0} 件 / 総コスト: `${1:N6} / 未割当トークン: {2} / スキップ行: {3}" -f `
        $report.meta.record_count, $report.summary.total_cost, $report.summary.unassigned_tokens, $report.summary.skipped_lines)
    return $full
}

# dot-source 時（テスト）は関数定義のみ。直接実行時のみメインを走らせる
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-Main -ClaudeRoot $ClaudeRoot -CodexRoot $CodexRoot -ClineRoot $ClineRoot `
        -PricingPath $PricingPath -OutFile $OutFile -Tz $Tz -IncludeSubagents $IncludeSubagents | Out-Null
}
