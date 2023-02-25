Option Explicit

' #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
' �t�H���_�E�t�@�C���̖��̎擾�֘A

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : �J�����g�t�H���_���擾����B
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetCurrentDirectory()

    Dim objWshShell     ' WshShell �I�u�W�F�N�g

    Set objWshShell = WScript.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        WScript.Echo "�G���[: " & Err.Description
        wscript.quit(1)
    End If
    GetCurrentDirectory = objWshShell.CurrentDirectory

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : �t�H���_���̊e�t�@�C�������擾���Ĕz��Ŗ߂��B
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetFileNames(folderName)

    Dim objFileSys, objFolder, tmpFile, objFile

    ' �t�@�C���V�X�e���������I�u�W�F�N�g���쐬
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    ' ���� crrDir �̃t�H���_�̃I�u�W�F�N�g���擾
    Set objFolder = objFileSys.GetFolder(folderName)

    ' �t�@�C���������ꍇ
    IF objFolder.Files.Count = 0 then
        GetFileNames = -1
        Exit Function
    End IF

    ' Folder�I�u�W�F�N�g��Files�v���p�e�B����File�I�u�W�F�N�g���擾
    tmpFile = ""
    For Each objFile In objFolder.Files

        ' �擾�����t�@�C���̃t�@�C�����i�[
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
' breif : �g���q�����̃t�@�C�������擾����B
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetBaseName(fileName)

    Dim objFileSys
 
    '�t�@�C���V�X�e���������I�u�W�F�N�g���쐬
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
    '�g���q�����̃t�@�C�������擾
    GetBaseName = objFileSys.GetBaseName(fileName)

End Function


' #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
' �t�H���_�E�t�@�C���̍폜�֘A

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' brief : �w�肵���t�@�C�����폜����B
' note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub DeleteFile (ByVal strFile)

	Dim objFso
	Set objFso = CreateObject("Scripting.FileSystemObject")

	' �t�H���_���폜
	objFso.DeleteFile strFile,True		' ���ӁF�߂�l�������ꍇ�͈������i�j�Ŋ���Ȃ�����

	Set objFso = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' brief : �w�肵���t�H���_���폜����B
' note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub DeleteFolder (ByVal strFolder)

	Dim objFso
	Set objFso = CreateObject("Scripting.FileSystemObject")

	' �t�H���_���폜
	objFso.DeleteFolder strFolder,True	' ���ӁF�߂�l�������ꍇ�͈������i�j�Ŋ���Ȃ�����

	Set objFso = Nothing

End Sub


' #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
' �t�@�C������֘A

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : �t�@�C�����̃e�L�X�g��S�擾
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetFileText(fileName)

    On Error Resume Next

    Dim objFSO      ' FileSystemObject
    Dim objFile     ' �t�@�C���ǂݍ��ݗp

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
            WScript.Echo "�t�@�C���I�[�v���G���[: " & Err.Description
        End If
    Else
        WScript.Echo "�G���[: " & Err.Description
    End If

    Set objFile = Nothing
    Set objFSO = Nothing

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : �w�肵���t�@�C���Ƀe�L�X�g�f�[�^���������ށB
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub WriteFile(fileBaseName, data, extension)

    Dim objFSO      ' FileSystemObject
    Dim objFile     ' �t�@�C���������ݗp
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
            WScript.Echo "�t�@�C���I�[�v���G���[: " & Err.Description
        End If
    Else
        WScript.Echo "�G���[: " & Err.Description
    End If

    On Error Goto 0

    Set objFile = Nothing
    Set objFSO = Nothing

End Sub
