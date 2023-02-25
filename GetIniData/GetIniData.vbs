Option Explicit

Msgbox GetIniData(GetCurrentDirectory() & "\Config.ini", "test1", "data1")

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
    GetCurrentDirectory = objWshShell.CurrentDirectory

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : iniファイル、セクション名、パラメータ名からデータを取得する。
' note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetIniData(strIniFileName, strSection, strKey)

    Dim objFSO, objIniFile, objSectionDic, strReadLine, objKeyDic, arrReadLine
    Dim strTempSection, objTempdic

    ' ファイル入出力の定数
    Const conForReading = 1, conForWriting = 2, conForAppending = 8
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' ファイルのOPEN
    Set objIniFile = objFSO.OpenTextFile(strIniFileName, conForReading, False)
    If Err.Number <> 0 Then
        ' エラーメッセージを出力
        wscript.echo "INIファイル名: " & strIniFileName
        wscript.quit(1)
    End If

    ' 格納先Dictionaryオブジェクトの作成
    Set objSectionDic = CreateObject("Scripting.Dictionary")

    ' ファイルのリードREAD
    strReadLine = objIniFile.ReadLine
    Do While objIniFile.AtEndofStream = False
        ' ステートメント開始行を検索
        If (strReadLine <> " ") And (StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0) Then
            ' セクション名を取得
            strTempSection = Mid(strReadLine, 2, (Len(strReadLine) - 2))
            ' キー用Dictionaryオブジェクト作成
            Set objKeyDic = CreateObject("Scripting.Dictionary")
            ' ファイルの最終行になるまでLoop
            Do While objIniFile.AtEndofStream = False
                strReadLine = objIniFile.ReadLine
                If (strReadLine <> "") And (StrComp(";", Left(strReadLine, 1)) <> 0) Then
                    ' 次のステートメント開始行が出現したら、Loop終了
                    If StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0 Then
                        Exit Do
                    End If
                    ' １セクション内の定義をDictionaryオブジェクトに格納する
                    arrReadLine = Split(strReadLine, "=", 2, vbTextCompare)
                    objKeyDic.Add UCase(arrReadLine(0)), arrReadLine(1)
                End If
            Loop
            ' オブジェクトに格納する
            objSectionDic.Add UCase(strTempSection), objKeyDic
        Else
            strReadLine = objIniFile.ReadLine
        End If
    Loop
    ' ファイルのCLOSE
    objIniFile.Close

    ' ini配列から指定したセクション、パラメータに対応するデータを取得
    strSection = UCase(strSection)  ' 大文字化
    strKey = UCase(strKey)          ' 大文字化

    If objSectionDic.Exists(strSection) Then
        Set objTempdic = objSectionDic.Item(strSection)
        If objTempdic.Exists(strKey) Then
            getIniData = objSectionDic(strSection)(strKey)
            Exit Function
        End If
    End If
    getIniData = ""

End Function
