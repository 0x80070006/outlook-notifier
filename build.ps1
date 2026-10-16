# Builds dist\OutlookNotifier.msi using only what ships with Windows (.NET Framework csc, msi.dll, makecab).
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$obj = Join-Path $root 'obj'
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force $obj, $dist | Out-Null

$fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$csc = Join-Path $fw 'csc.exe'
$wpf = Join-Path $fw 'WPF'

Write-Host 'Icon...'
& (Join-Path $root 'build\make-icon.ps1') -Path (Join-Path $obj 'app.ico')

Write-Host 'Application...'
$sources = Get-ChildItem (Join-Path $root 'src') -Filter *.cs | ForEach-Object { $_.FullName }
& $csc /nologo /target:winexe /optimize+ /codepage:65001 /platform:anycpu `
    "/out:$obj\OutlookNotifier.exe" "/win32icon:$obj\app.ico" "/win32manifest:$root\src\app.manifest" `
    "/resource:$obj\app.ico,OutlookNotifier.app.ico" `
    "/resource:$root\src\ui\Popup.xaml,OutlookNotifier.Popup.xaml" `
    "/resource:$root\src\ui\Settings.xaml,OutlookNotifier.Settings.xaml" "/lib:$wpf" `
    /r:PresentationFramework.dll /r:PresentationCore.dll /r:WindowsBase.dll /r:System.Xaml.dll `
    /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Core.dll `
    $sources
if ($LASTEXITCODE) { throw 'Application build failed' }

Write-Host 'Installer...'
& $csc /nologo /target:exe /optimize+ /codepage:65001 "/out:$obj\MsiBuilder.exe" (Join-Path $root 'installer\MsiBuilder.cs')
if ($LASTEXITCODE) { throw 'MsiBuilder build failed' }
& "$obj\MsiBuilder.exe" "$obj\OutlookNotifier.exe" "$obj\app.ico" "$dist\OutlookNotifier.msi" $obj
if ($LASTEXITCODE) { throw 'MSI generation failed' }

Copy-Item "$obj\OutlookNotifier.exe" $dist -Force
