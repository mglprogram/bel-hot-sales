# bel-hot-sales installer (Windows / PowerShell)
#
# Copies the skill into ~/.agents/skills/bel-hot-sales and checks dependencies.
# Re-runnable: safe to run again to update.
#
#   powershell -ExecutionPolicy Bypass -File install.ps1
#
[CmdletBinding()]
param(
    [string]$SkillsRoot = (Join-Path $env:USERPROFILE ".agents\skills"),
    [string]$SourceDir  = "",
    [switch]$Force,
    [switch]$SkipDeps
)

$ErrorActionPreference = "Stop"
$SkillName = "bel-hot-sales"

# Resolve the source directory: prefer an explicit -SourceDir, else the directory
# this script actually lives in. $PSScriptRoot is reliable under -File but empty
# when the script is piped/dot-sourced, so fall back to the invocation path and
# finally to the current directory.
if (-not $SourceDir) {
    if ($PSScriptRoot) {
        $SourceDir = $PSScriptRoot
    } elseif ($MyInvocation.MyCommand.Path) {
        $SourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    } elseif ($PSCommandPath) {
        $SourceDir = Split-Path -Parent $PSCommandPath
    } else {
        $SourceDir = (Get-Location).Path
    }
}
$SourceDir = (Resolve-Path -LiteralPath $SourceDir).Path

$Target = Join-Path $SkillsRoot $SkillName

function Info($m) { Write-Host "  $m" }
function Ok($m)   { Write-Host "  [OK] $m"   -ForegroundColor Green }
function Warn($m) { Write-Host "  [!]  $m"   -ForegroundColor Yellow }
function Fail($m) { Write-Host "  [X]  $m"   -ForegroundColor Red }

Write-Host ""
Write-Host "bel-hot-sales installer" -ForegroundColor Cyan
Write-Host "=======================" -ForegroundColor Cyan

# --- 1. sanity-check the source -------------------------------------------
Write-Host ""
Write-Host "1. Checking source"
$srcSkillMd = Join-Path $SourceDir "SKILL.md"
if (-not (Test-Path $srcSkillMd)) {
    Fail "SKILL.md not found in '$SourceDir'."
    Info "Run this script from inside the cloned repo, or pass -SourceDir <path>."
    exit 1
}
$scriptSrc = Join-Path $SourceDir "scripts\bel_pipeline.py"
if (-not (Test-Path $scriptSrc)) {
    Warn "scripts/bel_pipeline.py not found. The skill will still load, but its"
    Warn "bundled script will be missing and the agent must reimplement the logic."
} else {
    Ok "found SKILL.md and scripts/bel_pipeline.py"
}

# --- 2. copy ---------------------------------------------------------------
Write-Host ""
Write-Host "2. Installing to $Target"

if (Test-Path $Target) {
    if (-not $Force) {
        Warn "Target already exists. Overwriting (use -Force to skip this prompt)."
    }
    Remove-Item $Target -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $Target | Out-Null

foreach ($item in @("SKILL.md", "README.md", "LICENSE")) {
    $p = Join-Path $SourceDir $item
    if (Test-Path $p) { Copy-Item $p $Target -Force }
}
foreach ($dir in @("scripts", "evals")) {
    $p = Join-Path $SourceDir $dir
    if (Test-Path $p) {
        Copy-Item $p $Target -Recurse -Force
        Get-ChildItem (Join-Path $Target $dir) -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Ok "copied files"

# --- 3. verify the layout that skill discovery actually requires -----------
Write-Host ""
Write-Host "3. Verifying layout"
$check = Join-Path $Target "SKILL.md"
if (-not (Test-Path $check)) {
    Fail "SKILL.md is not at the root of $Target -- the skill will NOT load."
    exit 1
}
$fm = Get-Content $check -TotalCount 5 -Encoding UTF8
if (-not ($fm -join "`n" -match "^---")) {
    Warn "SKILL.md does not start with '---' YAML frontmatter; it may not be parsed."
} else {
    Ok "SKILL.md present with frontmatter at the skill root"
}
$nameLine = ($fm | Select-String -Pattern "^name:\s*(.+)$" | Select-Object -First 1)
if ($nameLine) {
    $declared = $nameLine.Matches[0].Groups[1].Value.Trim()
    if ($declared -ne $SkillName) {
        Warn "frontmatter name is '$declared' but the folder is '$SkillName'."
        Warn "Keep them equal to avoid confusion."
    } else {
        Ok "frontmatter name matches folder name ('$SkillName')"
    }
}

# --- 4. dependencies -------------------------------------------------------
Write-Host ""
Write-Host "4. Checking dependencies"
if ($SkipDeps) {
    Info "skipped (-SkipDeps)"
} else {
    $py = $null
    foreach ($cand in @("python", "py", "python3")) {
        $cmd = Get-Command $cand -ErrorAction SilentlyContinue
        if ($cmd) { $py = $cmd.Source; break }
    }
    if (-not $py) {
        Fail "Python not found on PATH."
        Info "The skill is installed and will still load, but the bundled script"
        Info "cannot run. Either install Python 3, or let the agent follow SKILL.md"
        Info "and implement the logic itself."
    } else {
        Ok "python: $py"
        $has = & $py -c "import openpyxl, sys; sys.stdout.write(openpyxl.__version__)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $has) {
            Ok "openpyxl $has"
        } else {
            Warn "openpyxl is not installed."
            $ans = Read-Host "  Install it now with 'pip install openpyxl'? [y/N]"
            if ($ans -match "^(y|yes)$") {
                & $py -m pip install --quiet openpyxl
                if ($LASTEXITCODE -eq 0) { Ok "openpyxl installed" }
                else { Fail "pip install failed -- install it manually." }
            } else {
                Info "skipped. Run: $py -m pip install openpyxl"
            }
        }
    }
}

# --- 5. done ---------------------------------------------------------------
Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Info "Installed to: $Target"
Info ""
Info "Test it by asking your agent something like:"
Info "  '这个十部半托管库存表给你，筛选出两周都有销量的产品，按 V3 格式输出'"
Info ""
Info "Note: a skill is only discovered when the agent starts a NEW session."
