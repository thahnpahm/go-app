# Don tan du cua mot ung dung da go do dang: registry HKCU + thu muc AppData. Xem README.md.
param(
    [string]$Pattern,          # regex ten/duong dan app, vd: terabox|dubox   hoac   \\Programs\\riff\\
    [switch]$DryRun,           # chi liet ke
    [switch]$Yes,              # khong hoi (dung cho test)
    [switch]$NoRestart,        # khong khoi dong lai Explorer
    [switch]$SkipViewCheck     # bo qua kiem tra registry ao (dung cho test)
)
$ErrorActionPreference = 'Stop'
if (-not $Pattern) { $Pattern = Read-Host 'Ten app can don (regex, vd: terabox|dubox)' }

# Mau khop bat ky chuoi nao duoi day la qua rong -> tu choi.
$protected = 'Microsoft Office', 'Windows', 'Google Chrome', 'Word.Document.12', 'Explorer', 'System32', 'Software', 'CLSID',
    'C:\Windows\', 'C:\Program Files\', "$env:LOCALAPPDATA\Programs\", "$env:APPDATA\Microsoft\"
if ($Pattern.Length -lt 3 -or @($protected -match $Pattern).Count) {
    Write-Host "Tu choi: mau '$Pattern' qua ngan hoac qua rong."; exit 1
}

$hkcu = [Microsoft.Win32.Registry]::CurrentUser
$hklm = [Microsoft.Win32.Registry]::LocalMachine
$cv = 'Software\Microsoft\Windows\CurrentVersion'

# Mot so ung dung (vd Claude Desktop) cho terminal thay ban AO cua HKCU: xoa o do khong co tac dung.
function Test-RealView {
    $id = [guid]::NewGuid().ToString(); $key = "Software\go-app-probe-$id"
    $k = $hkcu.CreateSubKey($key); $k.SetValue('', $id); $k.Close()
    try {
        $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $r = Invoke-CimMethod -Namespace root\default -ClassName StdRegProv -MethodName GetStringValue `
            -Arguments @{ hDefKey = [uint32]2147483651; sSubKeyName = "$sid\$key"; sValueName = '' }
        return $r.sValue -eq $id
    } finally { $hkcu.DeleteSubKey($key, $false) }
}
if (-not $SkipViewCheck -and -not (Test-RealView)) {
    Write-Host 'CANH BAO: terminal nay chi thay registry AO. Hay bam dup go-app.cmd tu File Explorer.'
    if (-not $DryRun) { exit 2 }
}

# 1. Ung dung con cai that thi phai go bang trinh go cua no truoc.
$uninstallKeys = @(); $alive = @(); $machine = @()
foreach ($src in @(
        @($hkcu, "$cv\Uninstall"),
        @($hklm, "$cv\Uninstall"),
        @($hklm, 'Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'))) {
    $root = $src[0].OpenSubKey($src[1]); if (-not $root) { continue }
    foreach ($n in $root.GetSubKeyNames()) {
        $k = $root.OpenSubKey($n); if (-not $k) { continue }
        $text = (@($n) + @('DisplayName', 'Publisher', 'UninstallString', 'InstallLocation' | ForEach-Object { [string]$k.GetValue($_) })) -join "`n"
        if ($text -notmatch $Pattern) { continue }
        $cmd = [string]$k.GetValue('UninstallString')
        if ($cmd -match '^\s*"([^"]+)"' -or $cmd -match '^\s*(.+?\.exe)') {
            if (Test-Path -LiteralPath $Matches[1]) { $alive += "$($k.GetValue('DisplayName')): $cmd" }
        }
        if ($src[0] -eq $hkcu) { $uninstallKeys += "$($src[1])\$n" } else { $machine += "HKLM\$($src[1])\$n" }
    }
}
if ($alive) {
    Write-Host 'Ung dung van con trinh go cai dat. Hay go trong Settings > Apps truoc:'
    $alive | ForEach-Object { Write-Host "  $_" }
    if (-not $DryRun) { exit 3 }
}

# 2. HKCU\Software\Classes
$classes = $hkcu.OpenSubKey('Software\Classes', $true)
function Find-Hit($key, [string]$path) {
    if ($path -and $path.Split('\')[-1] -match $Pattern) { return [pscustomobject]@{ Path = $path; Name = $null } }
    foreach ($n in $key.GetValueNames()) {
        if ("$n`n$(@($key.GetValue($n)) -join "`n")" -match $Pattern) { [pscustomobject]@{ Path = $path; Name = $n } }
    }
    foreach ($s in $key.GetSubKeyNames()) {
        try { $c = $key.OpenSubKey($s) } catch { continue }
        if ($c) { Find-Hit $c $(if ($path) { "$path\$s" } else { $s }); $c.Close() }
    }
}
function Test-MachineClass([string]$rel) { [bool]$hklm.OpenSubKey("Software\Classes\$rel") }

$containers = 'CLSID', 'TypeLib', 'Interface', 'AppID', 'WOW6432Node', 'Local Settings', 'Applications',
    'SystemFileAssociations', '*', 'Directory', 'Folder', 'Drive', 'AllFilesystemObjects'
Write-Host 'Dang do registry, cho khoang 1 phut...'
$ownedKeys = @{}      # khoa chi thuoc ve app -> xoa ca khoa
$values = @()         # gia tri ghi de len khoa dung chung -> chi xoa gia tri
foreach ($h in @(Find-Hit $classes '' | Where-Object { $_.Path })) {
    $seg = $h.Path.Split('\')
    $own = $null
    if ($seg[0] -in 'CLSID', 'TypeLib' -and $seg.Count -ge 2) { $own = "$($seg[0])\$($seg[1])" }
    elseif ($seg[0] -eq 'WOW6432Node' -and $seg.Count -ge 3 -and $seg[1] -in 'CLSID', 'TypeLib') { $own = "$($seg[0])\$($seg[1])\$($seg[2])" }
    elseif (-not $seg[0].StartsWith('.') -and $seg[0] -notin $containers) { $own = $seg[0] }
    if ($own -and -not (Test-MachineClass $own)) { $ownedKeys[$own] = $true }
    elseif ($null -ne $h.Name) { $values += $h }
    elseif (-not (Test-MachineClass $h.Path)) { $ownedKeys[$h.Path] = $true }
}

# 3. Duoi file con tro toi kieu file cua app (vd .docx -> WPS.Docx.6)
$progIds = @($ownedKeys.Keys | Where-Object { $_ -notmatch '\\' })
$subKeyRefs = @()
foreach ($ext in @($classes.GetSubKeyNames() | Where-Object { $_.StartsWith('.') })) {
    $k = $classes.OpenSubKey($ext)
    if ([string]$k.GetValue('') -in $progIds) { $values += [pscustomobject]@{ Path = $ext; Name = '' } }
    $o = $k.OpenSubKey('OpenWithProgids')
    if ($o) { $o.GetValueNames() | Where-Object { $_ -in $progIds } | ForEach-Object { $values += [pscustomobject]@{ Path = "$ext\OpenWithProgids"; Name = $_ } } }
    $k.GetSubKeyNames() | Where-Object { $_ -in $progIds } | ForEach-Object { $subKeyRefs += "$ext\$_" }
}
$fileExtRefs = @()
$fe = $hkcu.OpenSubKey("$cv\Explorer\FileExts")
if ($fe) {
    foreach ($ext in $fe.GetSubKeyNames()) {
        $o = $fe.OpenSubKey("$ext\OpenWithProgids")
        if ($o) { $o.GetValueNames() | Where-Object { $_ -in $progIds -or $_ -match $Pattern } |
            ForEach-Object { $fileExtRefs += [pscustomobject]@{ Path = "$cv\Explorer\FileExts\$ext\OpenWithProgids"; Name = $_ } } }
    }
}

# 4. Khoa khac, muc tu chay, thu muc
$keys = @($uninstallKeys)
foreach ($parent in 'Software', 'Software\Google\Chrome\NativeMessagingHosts', 'Software\Microsoft\Edge\NativeMessagingHosts',
    'Software\Mozilla\NativeMessagingHosts', "$cv\Start\TileProperties") {
    $p = $hkcu.OpenSubKey($parent); if (-not $p) { continue }
    foreach ($n in $p.GetSubKeyNames()) {
        $default = if ($parent -match 'NativeMessagingHosts') { [string]$p.OpenSubKey($n).GetValue('') }
        if ("$n`n$default" -match $Pattern) { $keys += "$parent\$n" }
    }
}
$runValues = @()
$run = $hkcu.OpenSubKey("$cv\Run")
if ($run) { $runValues = @($run.GetValueNames() | Where-Object { "$_`n$($run.GetValue($_))" -match $Pattern }) }
$folders = @(Get-ChildItem -LiteralPath $env:APPDATA, $env:LOCALAPPDATA, "$env:LOCALAPPDATA\Programs" -Directory -Force -ErrorAction SilentlyContinue |
    Where-Object { "$($_.FullName)\" -match $(if ($Pattern.Contains('\\')) { $Pattern } else { "\\[^\\]*(?:$Pattern)[^\\]*\\$" }) } |
    ForEach-Object FullName)

$total = $ownedKeys.Count + $values.Count + $subKeyRefs.Count + $fileExtRefs.Count + $keys.Count + $runValues.Count + $folders.Count
if ($machine) { Write-Host 'Muc trong HKLM (can quyen admin, script khong dong vao):'; $machine | ForEach-Object { Write-Host "  $_" } }
if (-not $total) { Write-Host 'Khong thay gi de don.'; exit 0 }
Write-Host "Tim thay $total muc:"
$ownedKeys.Keys | Sort-Object | ForEach-Object { Write-Host "  khoa app     Classes\$_" }
$values + $fileExtRefs | ForEach-Object { Write-Host "  gia tri      $($_.Path) [$($_.Name)]" }
$subKeyRefs | ForEach-Object { Write-Host "  khoa con     Classes\$_" }
$keys | ForEach-Object { Write-Host "  khoa         HKCU\$_" }
$runValues | ForEach-Object { Write-Host "  tu chay      $_" }
$folders | ForEach-Object { Write-Host "  thu muc      $_  (vao Thung rac)" }
if ($DryRun) { Write-Host 'DryRun: chua xoa gi.'; exit 0 }
if (-not $Yes -and (Read-Host 'Doc ky danh sach. Go y roi Enter de xoa') -ne 'y') { Write-Host 'Da huy, chua xoa gi.'; exit 0 }

# 5. Sao luu (loi thi dung)
$bk = Join-Path $PSScriptRoot ('backups\' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
New-Item -ItemType Directory -Force $bk | Out-Null
$exports = @('Software\Classes', "$cv\Uninstall", "$cv\Run", "$cv\Explorer\FileExts") + $keys
$i = 0
foreach ($e in $exports) {
    if (-not $hkcu.OpenSubKey($e)) { continue }
    $i++
    $ErrorActionPreference = 'Continue'; reg.exe export "HKCU\$e" (Join-Path $bk "$i.reg") /y | Out-Null; $code = $LASTEXITCODE; $ErrorActionPreference = 'Stop'
    if ($code -ne 0) { Write-Host "Sao luu that bai: $e. Dung lai, chua xoa gi."; exit 5 }
}
Write-Host "Da sao luu vao $bk"

# 6. Xoa
$script:failed = 0
function Invoke-Safe([scriptblock]$action, [string]$what) {
    try { & $action } catch { $script:failed++; Write-Host "Khong xoa duoc: $what - $($_.Exception.Message)" }
}
foreach ($v in $values) {
    Invoke-Safe {
        $k = $classes.OpenSubKey($v.Path, $true); if (-not $k) { return }
        $k.DeleteValue($v.Name, $false); $k.Close()
        $p = $v.Path
        while ($p) {   # don khoa rong, vi khoa rong trong HKCU van che khoa cua HKLM
            $k = $classes.OpenSubKey($p); if (-not $k) { break }
            $empty = -not $k.ValueCount -and -not $k.SubKeyCount; $k.Close()
            if (-not $empty) { break }
            $classes.DeleteSubKey($p)
            $n = $p.LastIndexOf('\'); $p = if ($n -gt 0) { $p.Substring(0, $n) } else { '' }
        }
    } $v.Path
}
foreach ($k in @($ownedKeys.Keys) + $subKeyRefs) { Invoke-Safe { $classes.DeleteSubKeyTree($k, $false) } $k }
foreach ($v in $fileExtRefs) { Invoke-Safe { $k = $hkcu.OpenSubKey($v.Path, $true); if ($k) { $k.DeleteValue($v.Name, $false) } } $v.Path }
foreach ($k in $keys) { Invoke-Safe { $hkcu.DeleteSubKeyTree($k, $false) } $k }
if ($runValues) { $run = $hkcu.OpenSubKey("$cv\Run", $true); foreach ($v in $runValues) { Invoke-Safe { $run.DeleteValue($v, $false) } $v } }
Add-Type -AssemblyName Microsoft.VisualBasic
foreach ($f in $folders) { Invoke-Safe { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($f, 'OnlyErrorDialogs', 'SendToRecycleBin') } $f }

if (-not $NoRestart) {
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache_*.db" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 4
    if (-not (Get-Process explorer -ErrorAction SilentlyContinue)) { Start-Process explorer.exe }
}
Write-Host "Xong. Loi: $script:failed."
if ($script:failed) { exit 4 } else { exit 0 }
