Option Explicit

Msgbox GetCurrentDirectory()

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : カレントフォルダを取得する。
' note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetCurrentDirectory()

    Dim objWshShell     ' WshShell オブジェクト

    Set objWshShell = WScript.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        WScript.Echo "エラー: " & Err.Description
        wscript.quit(1)
    End If
    getCurrentDirectory = objWshShell.CurrentDirectory

End Function