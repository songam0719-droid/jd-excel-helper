$ErrorActionPreference = 'Stop'
$repoDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = Join-Path $repoDir 'src'
$buildDir = Join-Path $repoDir 'bin'
$frameworkDir = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
if (-not (Test-Path -LiteralPath $frameworkDir)) { $frameworkDir = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319' }
$compiler = Join-Path $frameworkDir 'csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { throw 'A .NET Framework C# compiler is required on Windows.' }
New-Item -ItemType Directory -Path $buildDir -Force | Out-Null
$exePath = Join-Path $buildDir '京东查单工具.exe'
$refs = @('System.Core.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Xml.Linq.dll','System.Web.Extensions.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll')
$compilerArgs = @('/nologo','/optimize+','/target:winexe','/platform:anycpu',('/out:' + $exePath))
foreach ($ref in $refs) { $compilerArgs += '/r:' + (Join-Path $frameworkDir $ref) }
foreach ($ref in @('UIAutomationClient.dll','UIAutomationTypes.dll','WindowsBase.dll')) { $compilerArgs += '/r:' + (Join-Path $frameworkDir ('WPF\' + $ref)) }
foreach ($name in @('Core.cs','FailureMarks.cs','Automation.cs','UiRules.cs','Program.cs','SelfTests.cs','LayoutTests.cs')) { $compilerArgs += Join-Path $sourceDir $name }
& $compiler @compilerArgs
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
Write-Output $exePath
