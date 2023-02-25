# Script設置フォルダパス
$Current  = Split-Path $myInvocation.MyCommand.path 
# 実行プログラムパス
$Powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

# ---- 条件設定 ----
# $cscript = "C:\Windows\System32\cscript.exe"
$ini = "settings.ini"
$scriptFile = "oracledb_to_csv.wsf"

# スクリプトを呼び出す
Start-Process -FilePath $scriptFile $ini -Wait

exit 0