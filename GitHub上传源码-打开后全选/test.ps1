$ErrorActionPreference = 'Stop'
$repoDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $repoDir 'build.ps1') | Out-Host
$runDir = Join-Path (Join-Path ([IO.Path]::GetTempPath()) 'JdHelperTests') ([Guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $runDir -Force | Out-Null
$reportPath = Join-Path $runDir 'result.json'
$exePath = Join-Path (Join-Path $repoDir 'bin') '京东查单工具.exe'
$testProcess = Start-Process -FilePath $exePath -ArgumentList @('--self-test','--report',('"' + $reportPath + '"')) -WindowStyle Hidden -Wait -PassThru
if ($testProcess.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $reportPath)) { throw ('Self-test failed. See local report: ' + $reportPath) }
$report = Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($report.Error -or $report.Passed -ne 49) { throw ('Unexpected test result. See local report: ' + $reportPath) }
Write-Output ('PASS: ' + $report.Passed + ' tests. No WeChat operations or customer data were used.')
Write-Output ('Synthetic fixtures and report: ' + $runDir)
