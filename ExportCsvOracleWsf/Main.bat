@echo off

rem if文、for文の中で変数を使う場合は!を使用可能
@setlocal enabledelayedexpansion

rem ---- 条件設定 ----
set script=OracleToCsv.wsf
set ini=Config.ini

cd %~dp0
Cscript %script% %ini%

pause