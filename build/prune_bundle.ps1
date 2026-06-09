# Strip a built TSMIS Consolidator bundle down to runtime-only files, then GUARD
# against shipping anything a corporate DLP scanner would flag.
#
# Why this exists
#   In the sibling TSMIS Reports Exporter, bundled third-party documentation
#   shipped examples containing test credit-card numbers (e.g. 4111111111111111);
#   Microsoft 365 / SharePoint DLP detects "Credit Card Number" and BLOCKS the
#   file -- so a released zip becomes partly inaccessible. Rather than chase one
#   file, this script strips ALL prose documentation from the bundle and
#   sanitizes dist-info METADATA to its headers, then GUARDS by failing the
#   build if any non-license doc -- or any credit-card / private-key / AWS-key /
#   US-SSN content -- remains. License and notice files are kept (legally
#   required for OSS redistribution; never flagged). So a future dependency bump
#   can't silently reintroduce blocked content.
#
# Usage
#   Run automatically by build.ps1 on a fresh build. Also reusable on any
#   already-built / extracted release folder (the "TSMIS Consolidator" dir that
#   contains the .exe and _internal\):
#     powershell -ExecutionPolicy Bypass -File build\prune_bundle.ps1 -Target "C:\path\to\TSMIS Consolidator"
#   -GuardOnly verifies a bundle without pruning (e.g. to audit a release).

param(
    [Parameter(Mandatory = $true)][string]$Target,
    [switch]$GuardOnly,     # verify only -- do not delete anything
    [switch]$Quiet
)
$ErrorActionPreference = "Stop"

function Log($m) { if (-not $Quiet) { Write-Host $m } }

if (-not (Test-Path $Target)) { throw "Target not found: $Target" }
$Target = (Resolve-Path $Target).Path

# --- credit-card detection (mirrors DLP: IIN prefix + length + Luhn) ------
function Test-Luhn([string]$n) {
    $sum = 0; $alt = $false
    for ($i = $n.Length - 1; $i -ge 0; $i--) {
        $d = [int][string]$n[$i]
        if ($alt) { $d *= 2; if ($d -gt 9) { $d -= 9 } }
        $sum += $d; $alt = -not $alt
    }
    return ($sum % 10) -eq 0
}
function Test-CreditCard([string]$n) {
    # n is digits only. Require a real card brand prefix + canonical length so
    # random 16-digit hashes in bundled files are not false-positives.
    $len = $n.Length
    $ok = $false
    if ($n -match '^4' -and ($len -eq 13 -or $len -eq 16 -or $len -eq 19)) { $ok = $true }            # Visa
    elseif ($n -match '^3[47]' -and $len -eq 15) { $ok = $true }                                       # Amex
    elseif ($n -match '^5[1-5]' -and $len -eq 16) { $ok = $true }                                       # Mastercard
    elseif ($n -match '^2[2-7]' -and $len -eq 16) {                                                     # Mastercard 2-series
        $p = [int]$n.Substring(0, 4); if ($p -ge 2221 -and $p -le 2720) { $ok = $true } }
    elseif ($n -match '^(6011|65|64[4-9])' -and ($len -eq 16 -or $len -eq 19)) { $ok = $true }          # Discover
    elseif ($n -match '^(30[0-5]|36|38)' -and $len -eq 14) { $ok = $true }                              # Diners
    elseif ($n -match '^35' -and $len -ge 16 -and $len -le 19) { $ok = $true }                          # JCB
    if (-not $ok) { return $false }
    return (Test-Luhn $n)
}

# --- 1. prune non-runtime files ------------------------------------------
$before = (Get-ChildItem $Target -Recurse -File -ErrorAction Ignore | Measure-Object Length -Sum).Sum
$internal = Join-Path $Target "_internal"

if (-not $GuardOnly) {
    $removed = 0

    # Safety net for the PyInstaller `excludes` (Pillow / pypdfium2): if a hook
    # re-bundled them, drop the package dir + its dist-info. (openpyxl imports
    # Pillow eagerly, so it loads when present -- but the app's used code paths,
    # text/table extraction + plain workbooks, don't need it and tolerate its
    # absence; excluding it is verified safe by the frozen self-test,
    # build/full_smoke.py.)
    foreach ($name in @("PIL", "Pillow", "pypdfium2", "pypdfium2_raw")) {
        Get-ChildItem $internal -Directory -ErrorAction Ignore |
            Where-Object { $_.Name -eq $name -or $_.Name -like "$name-*" } |
            ForEach-Object { Remove-Item $_.FullName -Recurse -Force; $removed++; Log "  - removed _internal\$($_.Name)" }
    }

    # Generic dead weight in bundled Python packages: test suites and type stubs
    # (never imported at runtime).
    Get-ChildItem $internal -Recurse -Directory -ErrorAction Ignore |
        Where-Object { $_.Name -in @("tests", "test") } |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force; $removed++ }
    Get-ChildItem $internal -Recurse -File -Filter *.pyi -ErrorAction Ignore |
        ForEach-Object { Remove-Item $_.FullName -Force; $removed++ }

    # --- Documentation removal (the DLP surface) -----------------------------
    # Third-party docs are the proven data-loss-prevention risk. Strip ALL prose
    # docs bundle-wide -- *.md / *.markdown / *.rst and stray README/CHANGELOG/
    # HISTORY/AUTHORS/CONTRIBUTING/NEWS text -- EXCEPT license/notice files,
    # which OSS licenses legally require us to redistribute (and which never
    # carry DLP-flagged content). Our own clean "Start Here.txt" is at the app
    # root, not under _internal, so it is untouched.
    $licenseLike = '(?i)^(license|licence|copying|notice|copyright|third.?party)'
    Get-ChildItem $internal -Recurse -File -Include *.md, *.markdown, *.rst -ErrorAction Ignore |
        Where-Object { $_.BaseName -notmatch $licenseLike } |
        ForEach-Object { Remove-Item $_.FullName -Force; $removed++ }
    Get-ChildItem $internal -Recurse -File -ErrorAction Ignore |
        Where-Object { $_.Name -match '(?i)^(readme|changelog|changes|history|authors|contributing|news)(\.|$)' `
                       -and $_.BaseName -notmatch $licenseLike } |
        ForEach-Object { Remove-Item $_.FullName -Force; $removed++ }

    # Sanitize dist-info/egg-info METADATA: each embeds the package's FULL README
    # (pdfplumber's is 600+ lines) as the long-description body. Keep only the
    # RFC822 headers (Name/Version/... so importlib.metadata.version still works)
    # and drop everything after the first blank line -- removing the doc body
    # without breaking runtime version lookups. RECORD hashes aren't re-verified
    # at runtime, so a now-stale METADATA hash is harmless (we already delete
    # other listed files, e.g. tests/).
    Get-ChildItem $internal -Recurse -File -Filter "METADATA" -ErrorAction Ignore |
        Where-Object { $_.Directory.Name -like "*.dist-info" -or $_.Directory.Name -like "*.egg-info" } |
        ForEach-Object {
            $lines = [System.IO.File]::ReadAllLines($_.FullName)
            $end = [Array]::IndexOf($lines, "")     # first blank line ends the headers
            if ($end -gt 0 -and $end -lt $lines.Length - 1) {
                [System.IO.File]::WriteAllLines($_.FullName, $lines[0..($end - 1)])
                $removed++
            }
        }

    $after = (Get-ChildItem $Target -Recurse -File -ErrorAction Ignore | Measure-Object Length -Sum).Sum
    Log ("==> Pruned {0} item(s), reclaimed {1:N1} MB" -f $removed, (($before - $after) / 1MB))
}

# --- 2. GUARD -------------------------------------------------------------
# Fail the build if the bundle still carries documentation, or any high-confidence
# sensitive data type a corporate DLP scanner flags. This is the backstop that
# stops a future dependency bump from silently reintroducing the problem.

# (a) No prose docs (markdown / rst) may remain anywhere -- except license/notice
#     files, which we keep for legal redistribution.
$leftoverDocs = Get-ChildItem $Target -Recurse -File -Include *.md, *.markdown, *.rst -ErrorAction Ignore |
    Where-Object { $_.BaseName -notmatch '(?i)^(license|licence|copying|notice|copyright|third.?party)' }
if ($leftoverDocs) {
    throw "GUARD FAILED: documentation still bundled:`n  " + (($leftoverDocs.FullName) -join "`n  ")
}

# (b) No high-confidence sensitive data in any text file across the bundle.
#     Patterns mirror the common DLP "sensitive information types", each chosen
#     to avoid false positives that would wrongly block a release:
#       credit cards : brand IIN prefix + canonical length + Luhn
#       private keys : PEM "BEGIN ... PRIVATE KEY" blocks
#       AWS keys     : AKIA + 16 base32
#       US SSN       : dashed, with the invalid area/group/serial ranges excluded
$textExt = @('.md', '.markdown', '.rst', '.txt', '.text', '.json', '.yml', '.yaml',
    '.js', '.mjs', '.cjs', '.ts', '.html', '.htm', '.css', '.xml', '.csv', '.cfg',
    '.ini', '.toml', '.py', '.pem', '.crt', '.cer', '.key')
$textName = @('METADATA')           # extension-less files worth scanning
$rxCard = [regex]'(?<![0-9A-Za-z])[0-9](?:[ -]?[0-9]){12,18}(?![0-9A-Za-z])'
$rxKey  = [regex]'-----BEGIN (?:[A-Z0-9 ]+ )?PRIVATE KEY-----'
$rxAws  = [regex]'\bAKIA[0-9A-Z]{16}\b'
$rxSsn  = [regex]'(?<![0-9-])(?!000|666|9[0-9]{2})[0-9]{3}-(?!00)[0-9]{2}-(?!0000)[0-9]{4}(?![0-9-])'

$hits = New-Object System.Collections.Generic.List[string]
Get-ChildItem $Target -Recurse -File -ErrorAction Ignore | ForEach-Object {
    $f = $_
    if (-not (($textExt -contains $f.Extension.ToLower()) -or ($textName -contains $f.Name))) { return }
    try { $text = [System.IO.File]::ReadAllText($f.FullName) } catch { return }
    $rel = $f.FullName.Substring($Target.Length + 1)
    foreach ($m in $rxCard.Matches($text)) {
        if (Test-CreditCard ($m.Value -replace '[^0-9]', '')) { $hits.Add("  [credit-card] $rel"); break }
    }
    if ($rxKey.IsMatch($text)) { $hits.Add("  [private-key] $rel") }
    if ($rxAws.IsMatch($text)) { $hits.Add("  [aws-key]     $rel") }
    if ($rxSsn.IsMatch($text)) { $hits.Add("  [us-ssn]      $rel") }
}
if ($hits.Count -gt 0) {
    throw ("GUARD FAILED: sensitive data found in the bundle:`n" + (($hits | Select-Object -Unique) -join "`n"))
}

Log "==> Guard passed: no docs and no credit-card / private-key / AWS-key / SSN content in the bundle."
