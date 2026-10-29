$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
$env:HTTPS_PROXY="http://127.0.0.1:7897"
$env:HTTP_PROXY="http://127.0.0.1:7897"
Set-Location "C:\Users\Administrator\Desktop\qc-report-tool"
Remove-Item -LiteralPath ".\_scan.ps1",".\_scan.py",".\_scan.txt" -ErrorAction SilentlyContinue
git add -A
git status --short
Write-Output "--- commit ---"
git commit -m "chore: remove sensitive info for public release (names, company paths, printer)"
Write-Output "--- push ---"
git push origin main
Write-Output "--- make public ---"
gh repo edit Guo-dou/qc-report-tool --visibility public --accept-visibility-change-consequences 2>&1
Write-Output "--- final info ---"
gh repo view Guo-dou/qc-report-tool --json name,visibility,url
