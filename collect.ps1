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

# 単価は「100万トークンあたりUSD」で定義されるため、コスト算出時にこの値で除算する
$script:TokensPerMillion = 1e6

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
    return ($Line | ConvertFrom-Json)
}

# 1行JSONを安全にパースする。失敗時は壊れ行カウンタを進めて $null を返す
# （呼び出し側は戻り値の null 判定だけで continue できる）。
function ConvertFrom-JsonLineSafe {
    param([string]$Line)
    try { return (ConvertFrom-JsonLine $Line) }
    catch { $script:SkippedLineCount++; return $null }
}

# 渡された候補のうち最初の非空文字列を返す（複数キーのフォールバック解決用）。
function Get-FirstNonEmpty {
    param([string[]]$Values)
    foreach ($v in $Values) {
        if (-not [string]::IsNullOrEmpty($v)) { return $v }
    }
    return ''
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
    $obj = $raw | ConvertFrom-Json

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
    ) / $script:TokensPerMillion
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
            $obj = ConvertFrom-JsonLineSafe $line
            if ($null -eq $obj) { continue }
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
            $obj = ConvertFrom-JsonLineSafe $line
            if ($null -eq $obj) { continue }
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
            try { $meta = (Get-Content $metaPath -Raw -Encoding UTF8) | ConvertFrom-Json } catch { $meta = $null }
        }
        $repo = Get-FirstNonEmpty @(
            (Get-Prop $meta 'cwdOnTaskInitialization' ''),
            (Get-Prop $meta 'shadowGitConfigWorkTree' ''))

        $model = Get-FirstNonEmpty @(
            (Get-Prop $meta 'model' ''),
            (Get-Prop $meta 'apiModelId' ''),
            (Get-Prop $meta 'modelId' ''))
        if ([string]::IsNullOrEmpty($model) -and (Test-Path $histPath)) {
            $hist = Get-Content $histPath -Raw -Encoding UTF8
            $m = [regex]::Match($hist, '\b(claude-[\w.-]+|gpt-[\w.-]+|o\d[\w.-]*)\b')
            if ($m.Success) { $model = $m.Value }
        }

        $ui = ConvertFrom-JsonLineSafe (Get-Content $uiPath -Raw -Encoding UTF8)
        if ($null -eq $ui) { continue }
        $idx = 0
        foreach ($e in @($ui)) {
            if ((Get-Prop $e 'say') -ne 'api_req_started') { continue }
            $text = Get-Prop $e 'text' ''
            $t = ConvertFrom-JsonLineSafe $text
            if ($null -eq $t) { continue }

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
                $sv = [int]$r.cache_read * ([double](Get-Prop $price 'input' 0) - [double](Get-Prop $price 'cache_read' 0)) / $script:TokensPerMillion
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
:root{
--bg:#0f1115;--card:#181b22;--fg:#e6e8eb;--muted:#9aa3ad;--line:#2a2f3a;
--accent:#3b82f6;--accent2:#10b981;--un:#f59e0b;
--th-bg:#10131a;--code-bg:#10131a;--grid:#2a2f3a;
--bar-fill:rgba(59,130,246,.30);--row-hover:rgba(59,130,246,.12);
--zebra:rgba(255,255,255,.025);--total-bg:rgba(255,255,255,.06);
--chip-bg:rgba(255,255,255,.06);--shadow:0 1px 2px rgba(0,0,0,.40);
}
:root[data-theme="light"]{
--bg:#f5f7fa;--card:#ffffff;--fg:#1b1f24;--muted:#5b6571;--line:#e3e8ef;
--accent:#2563eb;--accent2:#059669;--un:#b45309;
--th-bg:#eef2f7;--code-bg:#eef2f7;--grid:#e3e8ef;
--bar-fill:rgba(37,99,235,.16);--row-hover:rgba(37,99,235,.08);
--zebra:rgba(15,23,42,.025);--total-bg:rgba(15,23,42,.05);
--chip-bg:rgba(15,23,42,.05);--shadow:0 1px 3px rgba(15,23,42,.10);
}
*{box-sizing:border-box}
body{margin:0;font:14px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--fg);transition:background .2s ease,color .2s ease}
header{padding:16px 20px;border-bottom:1px solid var(--line);display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap}
.head-main{min-width:0}
h1{font-size:18px;margin:0 0 8px;letter-spacing:.01em}
.meta{display:flex;flex-wrap:wrap;gap:6px}
.chip{background:var(--chip-bg);color:var(--muted);font-size:11px;border-radius:999px;padding:3px 10px;white-space:nowrap}
.chip b{color:var(--fg);font-weight:600;margin-left:2px}
.icon-btn{flex:none;width:38px;height:38px;display:inline-flex;align-items:center;justify-content:center;background:var(--card);color:var(--fg);border:1px solid var(--line);border-radius:10px;cursor:pointer;box-shadow:var(--shadow)}
.icon-btn:hover{border-color:var(--accent)}
.icon-btn svg{width:18px;height:18px}
.wrap{padding:20px;max-width:1200px;margin:0 auto}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:20px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;box-shadow:var(--shadow)}
.card .k{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.05em}
.card .v{font-size:22px;font-weight:650;margin-top:6px;word-break:break-all;font-variant-numeric:tabular-nums}
.card.primary{grid-column:span 2;border-color:var(--accent)}
.card.primary .k{color:var(--accent)}
.card.primary .v{font-size:30px}
.card .v.accent-un{color:var(--un)}
.charts{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px}
.chartbox{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;box-shadow:var(--shadow);min-width:0}
.chartbox h2{font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin:0 0 10px;font-weight:600}
.chartbox svg{display:block;width:100%;height:auto}
.chart-empty{color:var(--muted);font-size:12px;padding:24px 0;text-align:center}
.controls{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}
.tabs{display:flex;flex-wrap:wrap;gap:6px}
.tab,.toggle{background:var(--card);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:7px 13px;cursor:pointer;font:inherit;box-shadow:var(--shadow)}
.tab.active{background:var(--accent);border-color:var(--accent);color:#fff}
.toggle.active{background:var(--accent2);border-color:var(--accent2);color:#fff}
.search{margin-left:auto}
.search input{background:var(--card);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:7px 11px;font:inherit;min-width:200px}
.search input::placeholder{color:var(--muted)}
.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);border:0}
.tablewrap{max-height:72vh;overflow:auto;border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow)}
table{width:100%;border-collapse:separate;border-spacing:0;background:var(--card)}
th,td{padding:9px 11px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap;font-variant-numeric:tabular-nums}
th:first-child,td:first-child{text-align:left;white-space:normal;word-break:break-all;max-width:380px}
thead th{position:sticky;top:0;background:var(--th-bg);cursor:pointer;user-select:none;font-size:12px;color:var(--muted);z-index:2}
thead th:hover{color:var(--fg)}
tbody tr:nth-child(odd){background:var(--zebra)}
tbody tr:hover{background:var(--row-hover)}
tr.total{font-weight:700;background:var(--total-bg)}
tbody tr:last-child td{border-bottom:0}
.barcell{position:relative;min-width:120px}
.bar{position:absolute;left:0;top:50%;transform:translateY(-50%);height:62%;background:var(--bar-fill);border-radius:4px;z-index:0}
.barcell span{position:relative;z-index:1}
.un{color:var(--un)}.muted{color:var(--muted)}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
details{margin-top:16px;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px;box-shadow:var(--shadow)}
summary{cursor:pointer;color:var(--muted)}
code{background:var(--code-bg);padding:1px 5px;border-radius:4px}
@media (max-width:640px){
.charts{grid-template-columns:1fr}
.card.primary{grid-column:span 1}
.search{margin-left:0;width:100%}
.search input{width:100%;min-width:0}
.wrap{padding:14px}
}
'@

    $js = @'
const REPORT = JSON.parse(document.getElementById("data").textContent);
const AXES = [["model","モデル別"],["session","セッション別"],["repo","リポジトリ別"],["daily","日次"],["tool","ツール別"]];
const COLS = [
  ["key","名前","str"],["events","件数","num"],["input","input","num"],["output","output","num"],
  ["cache_read","cache_read","num"],["cache_creation","cache_write","num"],["cost","コスト$","cost"],
  ["out_in","out/in","num3"],["cache_rate","cache率","num3"],["cost_per_event","$/件","cost"]
];
const SUN='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M19.1 4.9l-1.4 1.4M6.3 17.7l-1.4 1.4"/></svg>';
const MOON='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>';

let view = REPORT.meta.default_view || "include";
let tab = "model";
let sort = {col:"cost", dir:-1};
let filter = "";

const ENT = {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"};
function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(c){return ENT[c];});}
function fmtNum(n){return (n==null)?"-":Number(n).toLocaleString();}
function fmtCompact(n){
  if(n==null)return "-";
  const v=Number(n);
  try{ return new Intl.NumberFormat(undefined,{notation:"compact",maximumFractionDigits:1}).format(v); }
  catch(e){ return v.toLocaleString(); }
}
function fmtCost(n){return (n==null)?'<span class="un">未割当</span>':("$"+Number(n).toFixed(6));}
function fmt3(n){return (n==null)?"-":Number(n).toFixed(3);}
function truncate(s,n){s=String(s==null?"":s);return s.length>n?s.slice(0,n-1)+"…":s;}

/* ---- theme ---- */
function readTheme(){
  try{ const t=localStorage.getItem("ttl-theme"); if(t==="light"||t==="dark")return t; }catch(e){}
  return (window.matchMedia&&window.matchMedia("(prefers-color-scheme: light)").matches)?"light":"dark";
}
let theme = readTheme();
function syncThemeBtn(){
  const b=document.getElementById("themeToggle");
  if(!b)return;
  b.innerHTML=(theme==="light")?MOON:SUN;
  b.setAttribute("aria-pressed",String(theme==="light"));
  b.setAttribute("aria-label",(theme==="light")?"ダークテーマに切り替え":"ライトテーマに切り替え");
}
function applyTheme(t){
  theme=t;
  document.documentElement.setAttribute("data-theme",t);
  try{ localStorage.setItem("ttl-theme",t); }catch(e){}
  syncThemeBtn();
  /* SVGはfill="var(--…)"でCSS変数に追従するため再描画不要 */
}
function initTheme(){
  document.documentElement.setAttribute("data-theme",theme);
  syncThemeBtn();
  const b=document.getElementById("themeToggle");
  if(b)b.onclick=()=>applyTheme(theme==="light"?"dark":"light");
}

/* ---- summary ---- */
function card(o){
  const t = (o.full!=null)?` title="${esc(o.full)}"`:"";
  return `<div class="card ${o.cls||''}"><div class="k">${esc(o.k)}</div><div class="v ${o.vcls||''}"${t}>${o.v}</div></div>`;
}
function renderSummary(){
  const s = REPORT.datasets[view].summary;
  const um = (s.unassigned_models||[]).filter(m=>m&&m!=="(empty)");
  const umTxt = um.length ? um.map(esc).join(", ") : "なし";
  const cards = [
    {k:"総コスト (USD)", v:"$"+Number(s.total_cost).toFixed(4), cls:"primary"},
    {k:"input トークン", v:fmtCompact(s.total_input), full:fmtNum(s.total_input)},
    {k:"output トークン", v:fmtCompact(s.total_output), full:fmtNum(s.total_output)},
    {k:"cache_read トークン", v:fmtCompact(s.total_cache), full:fmtNum(s.total_cache)},
    {k:"イベント数", v:fmtCompact(s.event_count), full:fmtNum(s.event_count)},
    {k:"推定節約額 (USD)", v:"$"+Number(s.estimated_savings).toFixed(4)},
    {k:"未割当トークン", v:fmtCompact(s.unassigned_tokens), full:fmtNum(s.unassigned_tokens), vcls:(s.unassigned_tokens>0?"accent-un":"")},
    {k:"スキップ壊れ行", v:fmtNum(s.skipped_lines)},
    {k:"未登録モデル", v:`<span style="font-size:13px">${esc(umTxt)}</span>`}
  ];
  document.getElementById("cards").innerHTML = cards.map(card).join("");
}

/* ---- rows ---- */
function sortRows(rows){
  const c=sort.col, d=sort.dir;
  return rows.sort((a,b)=>{
    let x=a[c], y=b[c];
    if(x==null&&y==null)return 0; if(x==null)return 1; if(y==null)return -1;
    if(typeof x==="string")return d*x.localeCompare(y);
    return d*(x-y);
  });
}
function rowsForTab(){
  let rows = (REPORT.datasets[view].axes[tab]||[]).slice();
  if(filter){ const f=filter.toLowerCase(); rows = rows.filter(r=>String(r.key==null?"":r.key).toLowerCase().indexOf(f)>=0); }
  return sortRows(rows);
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
  if(!rows.length){ body = `<tr><td colspan="${COLS.length}" class="muted" style="text-align:center;padding:18px">該当データなし</td></tr>`; }
  const totOutIn = tot.input>0?(tot.output/tot.input):null;
  const totRate = (tot.input+tot.cache_read)>0?(tot.cache_read/(tot.input+tot.cache_read)):null;
  body += `<tr class="total"><td>合計</td><td>${fmtNum(tot.events)}</td><td>${fmtNum(tot.input)}</td><td>${fmtNum(tot.output)}</td><td>${fmtNum(tot.cache_read)}</td><td>${fmtNum(tot.cache_creation)}</td><td>${fmtCost(Number(tot.cost.toFixed(6)))}</td><td>${fmt3(totOutIn)}</td><td>${fmt3(totRate)}</td><td>-</td></tr>`;
  document.getElementById("table").innerHTML = `<table><thead>${thead}</thead><tbody>${body}</tbody></table>`;
  document.querySelectorAll("thead th").forEach(th=>th.onclick=()=>{
    const c=th.getAttribute("data-c");
    if(sort.col===c){sort.dir*=-1;}else{sort.col=c;sort.dir=(c==="key")?1:-1;}
    renderTable();
  });
  renderCharts();
}

/* ---- charts (inline SVG, no external deps) ---- */
function costChart(rows){
  const el=document.getElementById("chartCost");
  if(!el)return;
  if(!rows.length){ el.innerHTML='<div class="chart-empty">該当データなし</div>'; return; }
  const W=560, rowH=30, padT=6, padB=6, labelW=170, valW=82;
  const H=padT+padB+rows.length*rowH;
  const max=Math.max(1e-12,...rows.map(r=>r.cost||0));
  const barMax=W-labelW-valW;
  let svg=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="表示中の軸のコスト上位内訳">`;
  rows.forEach((r,i)=>{
    const y=padT+i*rowH, cy=y+rowH/2, c=r.cost||0;
    const w=Math.max(0,Math.round(barMax*c/max));
    const label=esc(truncate(r.key,24)), val=esc("$"+c.toFixed(4));
    svg+=`<text x="0" y="${cy}" dominant-baseline="middle" font-size="12" fill="var(--fg)">${label}</text>`;
    svg+=`<rect x="${labelW}" y="${y+5}" width="${barMax}" height="${rowH-10}" rx="4" fill="var(--zebra)"/>`;
    svg+=`<rect x="${labelW}" y="${y+5}" width="${w}" height="${rowH-10}" rx="4" fill="var(--accent)"><title>${label}: ${val}</title></rect>`;
    svg+=`<text x="${W}" y="${cy}" dominant-baseline="middle" text-anchor="end" font-size="12" fill="var(--muted)">${val}</text>`;
  });
  el.innerHTML=svg+`</svg>`;
}
function dailyChart(){
  const el=document.getElementById("chartDaily");
  if(!el)return;
  let daily=(REPORT.datasets[view].axes.daily||[]).slice().filter(r=>/^\d{4}-\d{2}-\d{2}$/.test(String(r.key==null?"":r.key)));
  daily.sort((a,b)=>String(a.key).localeCompare(String(b.key)));
  if(!daily.length){ el.innerHTML='<div class="chart-empty">日次データなし</div>'; return; }
  const W=560, H=200, padL=8, padR=8, padT=10, padB=34;
  const n=daily.length, max=Math.max(1e-12,...daily.map(r=>r.cost||0));
  const plotW=W-padL-padR, plotH=H-padT-padB, slot=plotW/n;
  const bw=Math.max(2,Math.min(40,slot*0.6)), step=Math.max(1,Math.ceil(n/6));
  let svg=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="日次コスト推移">`;
  svg+=`<line x1="${padL}" y1="${padT+plotH}" x2="${W-padR}" y2="${padT+plotH}" stroke="var(--grid)" stroke-width="1"/>`;
  daily.forEach((r,i)=>{
    const c=r.cost||0, h=Math.round(plotH*c/max);
    const x=padL+slot*i+(slot-bw)/2, y=padT+plotH-h, val=esc("$"+c.toFixed(4));
    svg+=`<rect x="${x.toFixed(1)}" y="${y}" width="${bw.toFixed(1)}" height="${h}" rx="2" fill="var(--accent2)"><title>${esc(r.key)}: ${val}</title></rect>`;
    if(n<=8 || i===0 || i===n-1 || i%step===0){
      const lx=padL+slot*i+slot/2;
      svg+=`<text x="${lx.toFixed(1)}" y="${H-padB+16}" text-anchor="middle" font-size="10" fill="var(--muted)">${esc(String(r.key).slice(5))}</text>`;
    }
  });
  el.innerHTML=svg+`</svg>`;
}
function renderCharts(){
  let rows = rowsForTab().map(r=>({key:r.key, cost:r.cost||0}));
  rows.sort((a,b)=>b.cost-a.cost);
  const TOPN=10;
  if(rows.length>TOPN){
    const top=rows.slice(0,TOPN);
    const rest=rows.slice(TOPN).reduce((s,r)=>s+(r.cost||0),0);
    top.push({key:"その他 ("+(rows.length-TOPN)+")", cost:rest});
    rows=top;
  }
  costChart(rows);
  dailyChart();
}

/* ---- controls ---- */
function syncSearch(){
  const wrap=document.getElementById("searchWrap");
  if(wrap)wrap.style.display=(tab==="daily"||tab==="tool")?"none":"";
}
function renderTabs(){
  document.getElementById("tabs").innerHTML = AXES.map(a=>`<button class="tab ${tab===a[0]?'active':''}" role="tab" aria-selected="${tab===a[0]}" data-t="${a[0]}">${esc(a[1])}</button>`).join("");
  document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{tab=b.getAttribute("data-t");renderTabs();syncSearch();renderTable();});
}
function renderToggle(){
  const b=document.getElementById("subToggle");
  b.className="toggle "+(view==="include"?"active":"");
  b.textContent = view==="include"?"サブエージェント: 含む":"サブエージェント: 除く";
  b.setAttribute("aria-pressed",String(view==="include"));
  b.onclick=()=>{view=(view==="include")?"exclude":"include";renderToggle();renderSummary();renderTable();};
}
function initSearch(){
  const inp=document.getElementById("tableSearch");
  if(inp)inp.addEventListener("input",()=>{filter=inp.value.trim();renderTable();});
}

initTheme();renderToggle();renderSummary();renderTabs();syncSearch();initSearch();renderTable();
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
<script>
(function(){try{var t=localStorage.getItem("ttl-theme");if(t!=="light"&&t!=="dark"){t=(window.matchMedia&&window.matchMedia("(prefers-color-scheme: light)").matches)?"light":"dark";}document.documentElement.setAttribute("data-theme",t);}catch(e){}})();
</script>
<style>$css</style>
</head>
<body>
<header>
  <div class="head-main">
    <h1>AIエージェント トークン消費レポート</h1>
    <div class="meta">
      <span class="chip">生成<b>$genAt</b></span>
      <span class="chip">対象件数<b>$cnt</b></span>
      <span class="chip">単価日<b>$eff</b></span>
      <span class="chip">TZ<b>$tz</b></span>
      <span class="chip">ルート<b>$rootsHtml</b></span>
    </div>
  </div>
  <button id="themeToggle" class="icon-btn" type="button" aria-label="テーマ切り替え"></button>
</header>
<div class="wrap">
  <div id="cards" class="cards"></div>
  <div class="charts">
    <div class="chartbox"><h2>コスト内訳（表示中の軸・上位）</h2><div id="chartCost"></div></div>
    <div class="chartbox"><h2>日次コスト推移</h2><div id="chartDaily"></div></div>
  </div>
  <div class="controls">
    <div id="tabs" class="tabs" role="tablist"></div>
    <button id="subToggle" class="toggle" type="button"></button>
    <div id="searchWrap" class="search">
      <label for="tableSearch" class="sr-only">名前で絞り込み</label>
      <input id="tableSearch" type="search" placeholder="名前で絞り込み…" autocomplete="off">
    </div>
  </div>
  <div class="tablewrap"><div id="table"></div></div>
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
