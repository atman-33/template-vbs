# Script�ݒu�t�H���_�p�X
$Current  = Split-Path $myInvocation.MyCommand.path 
# ���s�v���O�����p�X
$Powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

# ---- �����ݒ� ----
# $cscript = "C:\Windows\System32\cscript.exe"
$ini = "settings.ini"
$scriptFile = "oracledb_to_csv.wsf"

# �X�N���v�g���Ăяo��
Start-Process -FilePath $scriptFile $ini -Wait

exit 0