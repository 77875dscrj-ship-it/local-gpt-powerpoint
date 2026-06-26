$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifest = Join-Path $root "manifest.xml"
$pptx = Join-Path $root "Local-GPT-PowerPoint.pptx"
$template = "C:\Users\saman\Documents\Codex\2026-06-25\gpt-for-powerpoint\work\chatgpt-office-addin\templates\PowerPointPresentationWithTaskPane.pptx"
$id = "6d0ef0de-25e9-4ca1-9f66-a40fc58f6f6b"

& (Join-Path $root "scripts\ensure-cert.ps1")

$devKey = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
New-Item -Path $devKey -Force | Out-Null
New-ItemProperty -Path $devKey -Name $id -Value $manifest -PropertyType String -Force | Out-Null
New-ItemProperty -Path $devKey -Name "RefreshAddins" -Value 1 -PropertyType DWord -Force | Out-Null

$python = "C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if (!(Test-Path $python)) { $python = "python" }
& $python (Join-Path $root "scripts\make-pptx.py") $template $pptx $id
if ($LASTEXITCODE -ne 0) {
  throw "PPTX 생성에 실패했습니다."
}

Start-Process -FilePath (Join-Path $root "start-local-gpt.cmd") -WindowStyle Hidden
Start-Sleep -Seconds 2
Start-Process $pptx

Write-Host "Local GPT registered and opened:"
Write-Host $pptx
