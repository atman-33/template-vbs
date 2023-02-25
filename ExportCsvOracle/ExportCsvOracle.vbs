Option Explicit

Const DEBUG_MODE = 1                ' 1:デバッグモード, 0:通常モード
Const INI_FILE = "Config.ini"     ' iniファイル名

Dim SQL_FOLDER_PATH, CSV_FOLDER_PATH
Dim SDB_PROVIDER, SDB_DATA_SOURCE, SDB_USER, SDB_PASS

Dim ini
ini = GetCurrentDirectory() & "\" & INI_FILE

' #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
' 0. iniファイルの読み込み　※実行VBSファイルと同じカレントフォルダに保存
SDB_PROVIDER = GetIniData(ini, "source_db", "provider")
SDB_DATA_SOURCE = GetIniData(ini, "source_db", "data_source")
SDB_USER = GetIniData(ini, "source_db", "user_id")
SDB_PASS = GetIniData(ini, "source_db", "password")

SQL_FOLDER_PATH = GetCurrentDirectory() & "\" & GetIniData(ini, "path", "sql_folder")
CSV_FOLDER_PATH = GetCurrentDirectory() & "\" & GetIniData(ini, "path", "csv_folder")
' #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####

Dim objAdoCon       ' ADO接続
Dim strSQLFiles     ' 実行するSQLのファイル群
Dim strSQLFile      ' 実行するSQLのファイル
Dim strSQL          ' 実行するSQL
Dim objAdoRS        ' ADOレコードセット
Dim csvText         ' SQLでSELECTしたCSVテキスト内容

' 1. DB接続
OpenDBOracle objAdoCon, SDB_PROVIDER, SDB_DATA_SOURCE, SDB_USER, SDB_PASS

' 2. SQLファイル群の読込
strSQLFiles = GetFileNames(SQL_FOLDER_PATH)

' 3. CSV生成
For Each strSQLFile In strSQLFiles                      ' 各SQLファイルを繰り返し処理
    WScript.Echo strSQLFile
    strSQL = GetFileText(strSQLFile)                    ' SQL文の取得
    ' Msgbox strSQL
    Set objAdoRS = ExcuteSQLgetRS(objAdoCon, strSQL)    ' SQL文を実行してレコードセットを取得
    csvText = GetCSVTextFromRS(objAdoRS)                ' レコードセットをCSV形式のテキストに変換

    WriteFile CSV_FOLDER_PATH & "\" & GetBaseName(strSQLFile), csvText, "csv"   ' CSVファイル生成
Next 

' 4. DB切断
CloseDB objAdoCon
Set objAdoCon = Nothing

If DEBUG_MODE = 1 Then
    WScript.Echo "処理が完了しました。"
End If

' Msgbox GetIniData(GetCurrentDirectory() & "\Config.ini", "test1", "data1")

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : DB接続（オラクル）
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub OpenDBOracle(ByRef objAdoCon, provider, dataSource, user, pass)

    If DEBUG_MODE = 1 Then
        WScript.Echo "DBに接続します。"
    End If

    Dim constr

    Set objAdoCon = WScript.CreateObject("ADODB.Connection")
    
    constr = "Provider=" & provider & ";Data Source=" & dataSource _
                & ";User ID=" & user & ";Password=" & pass

    If DEBUG_MODE = 1 Then
        WScript.Echo constr
    End If

    objAdoCon.ConnectionString = constr
    objAdoCon.Open

    If DEBUG_MODE = 1 Then
        WScript.Echo "DBに接続しました。"
    End If

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : トランザクション開始
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub BeginTrans(ByRef objAdoCon)

    If DEBUG_MODE = 1 Then
        WScript.Echo "トランザクションを開始します。"
    End If
    objAdoCon.BeginTrans

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : コミット
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub CommitTrans(ByRef objAdoCon)

    If DEBUG_MODE = 1 Then
        WScript.Echo "コミットします。"
    End If
    objAdoCon.CommitTrans

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : ロールバック
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub RollbackTrans(ByRef objAdoCon)

    If DEBUG_MODE = 1 Then
        WScript.Echo "ロールバックします。"
    End If
    objAdoCon.RollbackTrans

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : DB切断
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub CloseDB(ByRef objAdoCon)

    If DEBUG_MODE = 1 Then
        WScript.Echo "DBを切断します。"
    End If
    objAdoCon.Close

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : SQL SELECT文を実行してレコードセットを取得
' note  : 戻り値 -> レコードセット
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function ExcuteSQLgetRS(objAdoCon, strSQL)

    Dim objAdoRS  ' レコードセット

    ' Msgbox "ExcuteSQLgetRS -> SQL: " & strSQL    

    Set objAdoRS = objAdoCon.Execute(strSQL)

    ' Msgbox objAdoRS(0).value
    ' Msgbox "EOF:" & objAdoRS.EOF
    
    Set ExcuteSQLgetRS = objAdoRS   ' Object のため Set を忘れないこと

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : レコードセットをCSVに変換
' note  : 戻り値 -> CSV形式のテキスト
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetCSVTextFromRS(ByRef objAdoRS)

    Dim csvText
    Dim i

    csvText = ""
    Do While objAdoRS.EOF <> True

        'フィールドの数ループ
        For i = 0 to objAdoRS.fields.count -1
            If i <> objAdoRS.fields.count -1 then
                csvText = csvText & objAdoRS(i).value & ", "
            Else
                csvText = csvText & objAdoRS(i).value
            End If
        Next

        'テキスト改行
       csvText = csvText & vbCrLf 

       objAdoRS.MoveNext
    Loop

    objAdoRS.Close
    Set objAdoRS = Nothing

    If DEBUG_MODE = 1 Then
        WScript.Echo csvText
    End If
    GetCSVTextFromRS = csvText

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : 指定したファイルにテキストデータを書き込む。
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub WriteFile(fileBaseName, data, extension)

    Dim objFSO      ' FileSystemObject
    Dim objFile     ' ファイル書き込み用
    Dim fileName

    fileName = fileBaseName & "." & extension
    'Msgbox "witeFile.fileName: " & fileName 
    'Msgbox "witeFile.data: " & data 

    On Error Resume Next
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile(fileName, 2, True)
        If Err.Number = 0 Then
            objFile.Write(data)
            objFile.Close
        Else
            WScript.Echo "ファイルオープンエラー: " & Err.Description
        End If
    Else
        WScript.Echo "エラー: " & Err.Description
    End If

    On Error Goto 0

    Set objFile = Nothing
    Set objFSO = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : フォルダ内の各ファイル名を取得して配列で戻す。
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetFileNames(folderName)

    Dim objFileSys, objFolder, tmpFile, objFile

    ' ファイルシステムを扱うオブジェクトを作成
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    ' 引数 crrDir のフォルダのオブジェクトを取得
    Set objFolder = objFileSys.GetFolder(folderName)

    ' ファイルが無い場合
    IF objFolder.Files.Count = 0 then
        GetFileNames = -1
        Exit Function
    End IF

    ' FolderオブジェクトのFilesプロパティからFileオブジェクトを取得
    tmpFile = ""
    For Each objFile In objFolder.Files

        ' 取得したファイルのファイル名格納
        IF tmpFile = "" then
            tmpFile = folderName & "\" & objFile.Name
        Else
            tmpFile = tmpFile & "|" & folderName & "\" & objFile.Name
        End IF
    Next

    GetFileNames = split(tmpFile, "|")

    Set objFolder = Nothing
    Set objFileSys = Nothing

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : ファイル内のテキストを全取得
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetFileText(fileName)

    On Error Resume Next

    Dim objFSO      ' FileSystemObject
    Dim objFile     ' ファイル読み込み用

    GetFileText = ""    
    ' Msgbox "GetFileText -> fileName: " & fileName 
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile(fileName)
        If Err.Number = 0 Then
            GetFileText = objFile.ReadAll
            WScript.Echo getSQL
            objFile.Close
        Else
            WScript.Echo "ファイルオープンエラー: " & Err.Description
        End If
    Else
        WScript.Echo "エラー: " & Err.Description
    End If

    Set objFile = Nothing
    Set objFSO = Nothing

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : 拡張子無しのファイル名を取得する。
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetBaseName(fileName)

    Dim objFileSys
 
    'ファイルシステムを扱うオブジェクトを作成
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
    '拡張子無しのファイル名を取得
    GetBaseName = objFileSys.GetBaseName(fileName)

End Function

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
            GetIniData = objSectionDic(strSection)(strKey)
            Exit Function
        End If
    End If
    GetIniData = ""

End Function
