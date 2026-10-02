# Regenerates demo/index.html from preview/index.html. Run from the repo
# root after changing the Test build, then commit demo/index.html:
#   pwsh demo/build-demo.ps1
# The demo is the REAL admin app code running against demo/demo-mock.js's
# invented in-memory data (no Microsoft sign-in, no SharePoint, nothing
# saved). The only edits made to the copy are the four below.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$t = [System.IO.File]::ReadAllText("$root\preview\index.html")

# 1. Swap the Microsoft sign-in library for the fake one.
$msal = '<script src="https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/lib/msal-browser.min.js"></script>'
if(-not $t.Contains($msal)){ throw 'MSAL script tag not found - update build-demo.ps1' }
$t = $t.Replace($msal, '<script src="demo-mock.js"></script>')

# 2. Replace the PREVIEW BUILD banner text with a DEMO banner (no link to production).
$t = [regex]::Replace($t, '🧪 <strong>PREVIEW BUILD</strong>.*?</a>', '🎬 <strong>DEMO</strong> — sample data only. Every name here is invented; nothing is saved or sent, and a reload resets it.')

# 3. Disable the link to the real public intake form (it would post to the real Worker).
$t = $t.Replace('href="../intake/" target="_blank" rel="noopener"', 'href="#" onclick="alert(''Disabled in the demo.''); return false;"')

# 4. Title.
$t = $t.Replace('<title>FPHM Scheduler</title>', '<title>FPHM Scheduler (Demo)</title>')

[System.IO.File]::WriteAllText("$root\demo\index.html", $t, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Wrote demo/index.html"
