# 一鍵復原：把備份的 6 個檔 + 編譯 DLL 蓋回工作區原始位置。
# 用途：若 tracked 檔被 VS 還原回 git HEAD，執行此腳本即可恢復 3.1.2 / 3.1.7 的修改。
# 用法（PowerShell，於 repo 根目錄或任意位置）：  powershell -ExecutionPolicy Bypass -File _handoff\20260723_3.1.2_3.1.7\RESTORE.ps1
# 注意：復原前請先「關閉 Visual Studio」，否則可能立刻又被還原。

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)   # repo 根目錄
$bak  = $PSScriptRoot

$map = @{
  'src\util.cs'                 = 'classes\john\util.cs'
  'src\TDE0020Controller.cs'    = 'Controllers\NSPMIS\TDE0020Controller.cs'
  'src\SYS0030Controller.cs'    = 'Controllers\NSPMIS\SYS0030Controller.cs'
  'src\SYS0030Index.cshtml'     = 'Views\SYS0030\SYS0030Index.cshtml'
  'src\SYSInfo.cshtml'          = 'Views\SYS0030\SYSInfo.cshtml'
  'src\SYSManager.cshtml'       = 'Views\SYS0030\SYSManager.cshtml'
  'src\TDE0020IndexTab1.cshtml' = 'Views\TDE0020\TDE0020IndexTab1.cshtml'
  'src\TDE0020Index.cshtml'     = 'Views\TDE0020\TDE0020Index.cshtml'
  'src314\TDE0040Controller.cs' = 'Controllers\NSPMIS\TDE0040Controller.cs'
  'src314\TDE0040Index.cshtml'  = 'Views\TDE0040\TDE0040Index.cshtml'
  'src314\PAGE2.cshtml'         = 'Views\TDE0042\PAGE2.cshtml'
  'src314\PAGE3.cshtml'         = 'Views\TDE0043\PAGE3.cshtml'
  'src314\PAGE5.cshtml'         = 'Views\TDE0045\PAGE5.cshtml'
  'src314\PAGE6.cshtml'         = 'Views\TDE0046\PAGE6.cshtml'
  'src314\PAGE8.cshtml'         = 'Views\TDE0048\PAGE8.cshtml'
  'src314\PAGE9.cshtml'         = 'Views\TDE0049\PAGE9.cshtml'
  'bin\WCals.dll'               = 'bin\WCals.dll'
  'bin\WCals.pdb'               = 'bin\WCals.pdb'
}

foreach ($k in $map.Keys) {
  $srcFile = Join-Path $bak $k
  $dstFile = Join-Path $root $map[$k]
  if (Test-Path $srcFile) {
    Copy-Item -Force $srcFile $dstFile
    Write-Host "restored -> $($map[$k])"
  }
}
Write-Host "`n完成。若剛還原完，記得不要開 VS；如需重編譯請自行 MSBuild。"
