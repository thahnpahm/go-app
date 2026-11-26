# Tao tan du cua mot app gia, chay go-app.ps1, kiem tra xoa dung va giu dung.
$ErrorActionPreference = 'Stop'
$script = Join-Path $PSScriptRoot 'go-app.ps1'
$hkcu = [Microsoft.Win32.Registry]::CurrentUser
$cv = 'Software\Microsoft\Windows\CurrentVersion'
$app = 'ZzGoAppTest'
$exe = "$env:LOCALAPPDATA\$app\app.exe"
$shared = 'Software\Classes\cplfile'   # kieu file co san trong HKLM, gia lap bi app ghi de icon
if ($hkcu.OpenSubKey($shared)) { throw "HKCU\$shared da ton tai, khong chay test de tranh dung du lieu that." }

function Set-Reg($path, $name, $value) { $k = $hkcu.CreateSubKey($path); $k.SetValue($name, $value); $k.Close() }
function Assert($cond, $msg) { if (-not $cond) { throw "FAIL: $msg" }; Write-Host "ok - $msg" }
function Invoke-GoApp([string[]]$more) {
    powershell.exe -NoProfile -File $script @more -SkipViewCheck -NoRestart | Out-Host
    return $LASTEXITCODE
}

try {
    Set-Reg "Software\$app" '' 'x'
    Set-Reg "$cv\Run" $app "`"$exe`""
    Set-Reg "$cv\Uninstall\$app" 'UninstallString' "`"$env:LOCALAPPDATA\$app\uninst.exe`""
    Set-Reg 'Software\Classes\ZZT.Doc\shell\open\command' '' "`"$exe`" %1"
    Set-Reg 'Software\Classes\.zzgotest' '' 'ZZT.Doc'
    Set-Reg 'Software\Classes\.zzgotest\OpenWithProgids' 'ZZT.Doc' ''
    Set-Reg 'Software\Classes\.zzgotest\OpenWithProgids' 'Keep.Prog' ''
    Set-Reg "$shared\DefaultIcon" '' "$exe,1"
    New-Item -ItemType Directory -Force "$env:LOCALAPPDATA\$app" | Out-Null

    foreach ($bad in 'win', 'soft', '...', 'abcd|') {
        Assert ((Invoke-GoApp '-Pattern', "`"$bad`"", '-DryRun') -eq 1) "tu choi mau qua rong: $bad"
    }
    Assert ((Invoke-GoApp '-Pattern', $app, '-DryRun') -eq 0) 'dry run chay duoc'
    Assert ([bool]$hkcu.OpenSubKey("Software\$app")) 'dry run khong xoa gi'

    Assert ((Invoke-GoApp '-Pattern', $app, '-Yes') -eq 0) 'xoa khong loi'
    Assert (-not $hkcu.OpenSubKey("Software\$app")) 'xoa khoa Software'
    Assert (-not $hkcu.OpenSubKey("$cv\Uninstall\$app")) 'xoa muc go cai dat'
    Assert ($null -eq $hkcu.OpenSubKey("$cv\Run").GetValue($app)) 'xoa muc tu chay'
    Assert (-not $hkcu.OpenSubKey('Software\Classes\ZZT.Doc')) 'xoa kieu file cua app'
    $ext = $hkcu.OpenSubKey('Software\Classes\.zzgotest')
    Assert ($null -eq $ext.GetValue('')) 'duoi file khong con tro toi app'
    $names = $ext.OpenSubKey('OpenWithProgids').GetValueNames()
    Assert ($names -contains 'Keep.Prog' -and $names -notcontains 'ZZT.Doc') 'giu muc cua app khac'
    Assert (-not $hkcu.OpenSubKey($shared)) 'bo ghi de icon len kieu file dung chung'
    Assert (-not (Test-Path "$env:LOCALAPPDATA\$app")) 'chuyen thu muc vao Thung rac'
    Write-Host 'TAT CA TEST DAT'
} finally {
    $hkcu.DeleteSubKeyTree("Software\$app", $false)
    $hkcu.DeleteSubKeyTree("$cv\Uninstall\$app", $false)
    $hkcu.OpenSubKey("$cv\Run", $true).DeleteValue($app, $false)
    foreach ($k in 'Software\Classes\ZZT.Doc', 'Software\Classes\.zzgotest', $shared) { $hkcu.DeleteSubKeyTree($k, $false) }
    if (Test-Path "$env:LOCALAPPDATA\$app") { [IO.Directory]::Delete("$env:LOCALAPPDATA\$app", $true) }
}
