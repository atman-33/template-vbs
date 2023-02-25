Option Explicit

' #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
' フォルダ・ファイルの名称取得関連

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


' #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
' フォルダ・ファイルの削除関連

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' brief : 指定したファイルを削除する。
' note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub DeleteFile (ByVal strFile)

	Dim objFso
	Set objFso = CreateObject("Scripting.FileSystemObject")

	' フォルダを削除
	objFso.DeleteFile strFile,True		' 注意：戻り値が無い場合は引数を（）で括らないこと

	Set objFso = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' brief : 指定したフォルダを削除する。
' note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub DeleteFolder (ByVal strFolder)

	Dim objFso
	Set objFso = CreateObject("Scripting.FileSystemObject")

	' フォルダを削除
	objFso.DeleteFolder strFolder,True	' 注意：戻り値が無い場合は引数を（）で括らないこと

	Set objFso = Nothing

End Sub


' #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
' ファイル操作関連

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
