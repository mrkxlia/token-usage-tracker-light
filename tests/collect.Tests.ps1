# Pester v5 tests for collect.ps1
#
# 期待値はすべて tests/fixtures の固定ログから手計算したもの（出どころを各 It にコメント）。
# fixtures は「固定・以後不変」。値を変える場合はここの期待値も連動して更新すること。
#
# 単価（pricing.json 既定値, 1Mトークンあたり USD）:
#   claude-sonnet-4-6: in 3.0  out 15.0 cw1h 6.0  cw5m 3.75 cr 0.3
#   claude-opus-4-8:   in 5.0  out 25.0 cw1h 10.0 cw5m 6.25 cr 0.5
#   claude-haiku-4-5:  in 1.0  out 5.0  cw1h 2.0  cw5m 1.25 cr 0.1
#   gpt-5:             in 1.25 out 10.0 cw* 0.0          cr 0.125
#   gpt-5-mini:        in 0.25 out 2.0  cw* 0.0          cr 0.025

BeforeAll {
    # dot-source は必ず BeforeAll 内（Pester v5 の Discovery/Run 2フェーズ対策）
    $script:CollectPath = Join-Path $PSScriptRoot '..' 'collect.ps1'
    . $script:CollectPath
    $script:Fx        = Join-Path $PSScriptRoot 'fixtures'
    $script:ClaudeFx  = Join-Path $script:Fx 'claude_projects'
    $script:CodexFx   = Join-Path $script:Fx 'codex_sessions'
    $script:ClineFx   = Join-Path $script:Fx 'cline_tasks'
    $script:Pricing   = Import-PricingTable -PricingPath (Join-Path $PSScriptRoot '..' 'pricing.json')

    function Approx($a, $b, $tol = 1e-9) { return ([math]::Abs([double]$a - [double]$b) -lt $tol) }
}

Describe 'Import-PricingTable / Resolve-ModelPricing' {
    It 'loads bundled pricing with expected rates' {
        $script:Pricing.effective_date | Should -Be '2026-06-15'
        (Resolve-ModelPricing -Model 'claude-sonnet-4-6' -Pricing $script:Pricing).input  | Should -Be 3.0
        (Resolve-ModelPricing -Model 'claude-sonnet-4-6' -Pricing $script:Pricing).output | Should -Be 15.0
    }
    It 'resolves a dated model suffix by stripping -\d{8}$' {
        (Resolve-ModelPricing -Model 'claude-sonnet-4-6-20251001' -Pricing $script:Pricing).input | Should -Be 3.0
    }
    It 'returns $null for unknown models' {
        Resolve-ModelPricing -Model 'totally-unknown-model' -Pricing $script:Pricing | Should -BeNullOrEmpty
        Resolve-ModelPricing -Model '' -Pricing $script:Pricing | Should -BeNullOrEmpty
    }
    It 'applies aliases' {
        $p = Import-PricingTable -PricingPath (Join-Path $PSScriptRoot '..' 'pricing.json')
        $p.aliases['my-opus'] = 'claude-opus-4-8'
        (Resolve-ModelPricing -Model 'my-opus' -Pricing $p).output | Should -Be 25.0
    }
}

Describe 'Read-ClaudeLogs' {
    BeforeAll { $script:claude = @(Read-ClaudeLogs -Root $script:ClaudeFx) }

    It 'skips non-usage and broken lines, keeps usage rows incl. subagent (5xA + C + D + B = 8 raw)' {
        $script:claude.Count | Should -Be 8
    }
    It 'maps nested cache_creation to 5m/1h (msg_A)' {
        $a = $script:claude | Where-Object message_id -eq 'msg_A' | Select-Object -First 1
        $a.input             | Should -Be 100
        $a.output            | Should -Be 50
        $a.cache_read        | Should -Be 1000
        $a.cache_creation_1h | Should -Be 200
        $a.cache_creation_5m | Should -Be 300
        $a.model             | Should -Be 'claude-sonnet-4-6'
        $a.repo_path         | Should -Be '/home/u/repoA'
        $a.is_subagent       | Should -BeFalse
    }
    It 'falls back to 5m when nested cache_creation is absent (msg_C)' {
        $c = $script:claude | Where-Object message_id -eq 'msg_C' | Select-Object -First 1
        $c.cache_creation_5m | Should -Be 400
        $c.cache_creation_1h | Should -Be 0
    }
    It 'flags subagent rows (msg_B)' {
        $b = $script:claude | Where-Object message_id -eq 'msg_B' | Select-Object -First 1
        $b.is_subagent | Should -BeTrue
        $b.model       | Should -Be 'claude-opus-4-8'
    }
}

Describe 'Read-CodexLogs' {
    BeforeAll { $script:codex = @(Read-CodexLogs -Root $script:CodexFx) }

    It 'indexes only token_count events (2 records, intervening event_msg ignored)' {
        $script:codex.Count | Should -Be 2
        ($script:codex.message_id) | Should -Contain 'rollout-2026-06-15T01-00-00-abc#0'
        ($script:codex.message_id) | Should -Contain 'rollout-2026-06-15T01-00-00-abc#1'
    }
    It 'normalizes input=input-cached, cache_read=cached, reasoning kept separate (record 0)' {
        $r = $script:codex | Where-Object message_id -like '*#0'
        $r.input      | Should -Be 800   # 1000 - 200
        $r.cache_read | Should -Be 200
        $r.output     | Should -Be 300   # reasoning is included within output_tokens
        $r.reasoning  | Should -Be 50
        $r.model      | Should -Be 'gpt-5'
        $r.session_id | Should -Be 'sess-codex-1'
    }
    It 'uses the latest turn_context model (record 1 = gpt-5-mini)' {
        $r = $script:codex | Where-Object message_id -like '*#1'
        $r.model  | Should -Be 'gpt-5-mini'
        $r.input  | Should -Be 500
    }
}

Describe 'Read-ClineLogs' {
    BeforeAll { $script:cline = @(Read-ClineLogs -Root $script:ClineFx) }

    It 'parses api_req_started rows from each task (2 + 1 = 3 records)' {
        $script:cline.Count | Should -Be 3
    }
    It 'maps tokens and cacheWrites -> cache_creation_5m (task-123 row 0)' {
        $r = $script:cline | Where-Object message_id -eq 'task-123#0'
        $r.input             | Should -Be 50
        $r.output            | Should -Be 20
        $r.cache_read        | Should -Be 10
        $r.cache_creation_5m | Should -Be 5
        $r.model             | Should -Be 'claude-haiku-4-5'
        $r.repo_path         | Should -Be '/home/u/repoC'
        $r.timestamp         | Should -Match 'Z$'   # ISO8601 UTC
    }
    It 'leaves model empty (unassigned) when none is found (task-456)' {
        $r = $script:cline | Where-Object message_id -eq 'task-456#0'
        $r.model | Should -BeNullOrEmpty
        $r.input | Should -Be 70
    }
}

Describe 'Merge-UsageRecords (dedup)' {
    It 'folds 5 duplicate msg_A into 1 (composite key source+message_id), Claude total 4' {
        $merged = @(Merge-UsageRecords -Records (Read-ClaudeLogs -Root $script:ClaudeFx))
        $merged.Count | Should -Be 4
        ($merged | Where-Object message_id -eq 'msg_A').Count | Should -Be 1
    }
    It 'does not collide identical indices across sources' {
        $all = @(Read-CodexLogs -Root $script:CodexFx) + @(Read-ClineLogs -Root $script:ClineFx)
        $merged = @(Merge-UsageRecords -Records $all)
        $merged.Count | Should -Be ($all.Count)
    }
}

Describe 'Get-UsageCost' {
    It 'computes Claude msg_A cost (reasoning n/a)' {
        # (100*3 + 50*15 + 200*6 + 300*3.75 + 1000*0.3)/1e6 = 3675/1e6
        $a = (Read-ClaudeLogs -Root $script:ClaudeFx) | Where-Object message_id -eq 'msg_A' | Select-Object -First 1
        Approx (Get-UsageCost -Record $a -Pricing $script:Pricing) 0.003675 | Should -BeTrue
    }
    It 'computes Codex cost WITHOUT double-counting reasoning (record 0)' {
        # (800*1.25 + 300*10 + 200*0.125)/1e6 = 4025/1e6  ; reasoning(50) NOT added
        $r = (Read-CodexLogs -Root $script:CodexFx) | Where-Object message_id -like '*#0'
        Approx (Get-UsageCost -Record $r -Pricing $script:Pricing) 0.004025 | Should -BeTrue
    }
    It 'computes Cline cost (task-123 row 0)' {
        # (50*1 + 20*5 + 5*1.25 + 10*0.1)/1e6 = 157.25/1e6
        $r = (Read-ClineLogs -Root $script:ClineFx) | Where-Object message_id -eq 'task-123#0'
        Approx (Get-UsageCost -Record $r -Pricing $script:Pricing) 0.00015725 | Should -BeTrue
    }
    It 'returns $null for unassigned (unknown/empty model)' {
        $r = (Read-ClineLogs -Root $script:ClineFx) | Where-Object message_id -eq 'task-456#0'
        Get-UsageCost -Record $r -Pricing $script:Pricing | Should -BeNullOrEmpty
    }
}

Describe 'New-ReportHtml' {
    BeforeAll {
        $recs = @(Merge-UsageRecords -Records (
            @(Read-ClaudeLogs -Root $script:ClaudeFx) +
            @(Read-CodexLogs  -Root $script:CodexFx) +
            @(Read-ClineLogs  -Root $script:ClineFx)))
        $script:report = Build-Report -Records $recs -Pricing $script:Pricing -Roots @{ claude = $script:ClaudeFx }
        $script:html   = New-ReportHtml -Report $script:report
    }
    It 'declares utf-8 charset' {
        $script:html | Should -Match '<meta charset="utf-8">'
    }
    It 'escapes the closing script tag from embedded data' {
        # repo_path contains a closing-script-tag sequence which must be neutralized to the
        # escaped form. Needles are built by concatenation so the literal tag never appears on
        # a source line (Pester source introspection mishandles a literal closing tag here).
        $close   = '<' + '/script>'
        $needleRaw = 'repo' + $close + 'x'
        $needleEsc = 'repo<\' + '/script>x'
        $script:html.Contains($needleEsc) | Should -BeTrue   # escaped form present
        $script:html.Contains($needleRaw) | Should -BeFalse  # raw closing tag absent
    }
    It 'surfaces unassigned tokens count' {
        $script:report.summary.unassigned_tokens | Should -BeGreaterThan 0
    }
}

Describe 'End-to-end Invoke-Main' {
    It 'writes a BOM-less UTF-8 HTML file' {
        $out = Join-Path ([System.IO.Path]::GetTempPath()) ('report-{0}.html' -f ([guid]::NewGuid()))
        Invoke-Main -ClaudeRoot $script:ClaudeFx -CodexRoot $script:CodexFx -ClineRoot $script:ClineFx `
                    -PricingPath (Join-Path $PSScriptRoot '..' 'pricing.json') -OutFile $out
        Test-Path $out | Should -BeTrue
        $bytes = [System.IO.File]::ReadAllBytes($out)
        # BOM無し: 先頭3バイトが EF BB BF でないこと
        ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) | Should -BeFalse
        Remove-Item $out -ErrorAction SilentlyContinue
    }
}
