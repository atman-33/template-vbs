Option Explicit

Const DEBUG_MODE = 1                ' 1:�f�o�b�O���[�h, 0:�ʏ탂�[�h
Const INI_FILE = "Config.ini"     ' ini�t�@�C����

Dim SQL_FOLDER_PATH, CSV_FOLDER_PATH
Dim SDB_PROVIDER, SDB_DATA_SOURCE, SDB_USER, SDB_PASS

Dim ini
ini = GetCurrentDirectory() & "\" & INI_FILE

' #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
' 0. ini�t�@�C���̓ǂݍ��݁@�����sVBS�t�@�C���Ɠ����J�����g�t�H���_�ɕۑ�
SDB_PROVIDER = GetIniData(ini, "source_db", "provider")
SDB_DATA_SOURCE = GetIniData(ini, "source_db", "data_source")
SDB_USER = GetIniData(ini, "source_db", "user_id")
SDB_PASS = GetIniData(ini, "source_db", "password")

SQL_FOLDER_PATH = GetCurrentDirectory() & "\" & GetIniData(ini, "path", "sql_folder")
CSV_FOLDER_PATH = GetCurrentDirectory() & "\" & GetIniData(ini, "path", "csv_folder")
' #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####

Dim objAdoCon       ' ADO�ڑ�
Dim strSQLFiles     ' ���s����SQL�̃t�@�C���Q
Dim strSQLFile      ' ���s����SQL�̃t�@�C��
Dim strSQL          ' ���s����SQL
Dim objAdoRS        ' ADO���R�[�h�Z�b�g
Dim csvText         ' SQL��SELECT����CSV�e�L�X�g���e

' 1. DB�ڑ�
OpenDBOracle objAdoCon, SDB_PROVIDER, SDB_DATA_SOURCE, SDB_USER, SDB_PASS

' 2. SQL�t�@�C���Q�̓Ǎ�
strSQLFiles = GetFileNames(SQL_FOLDER_PATH)

' 3. CSV����
For Each strSQLFile In strSQLFiles                      ' �eSQL�t�@�C�����J��Ԃ�����
    WScript.Echo strSQLFile
    strSQL = GetFileText(strSQLFile)                    ' SQL���̎擾
    ' Msgbox strSQL
    Set objAdoRS = ExcuteSQLgetRS(objAdoCon, strSQL)    ' SQL�������s���ă��R�[�h�Z�b�g���擾
    csvText = GetCSVTextFromRS(objAdoRS)                ' ���R�[�h�Z�b�g��CSV�`���̃e�L�X�g�ɕϊ�

    WriteFile CSV_FOLDER_PATH & "\" & GetBaseName(strSQLFile), csvText, "csv"   ' CSV�t�@�C������
Next 

' 4. DB�ؒf
CloseDB objAdoCon
Set objAdoCon = Nothing

If DEBUG_MODE = 1 Then
    WScript.Echo "�������������܂����B"
End If

' Msgbox GetIniData(GetCurrentDirectory() & "\Config.ini", "test1", "data1")

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : DB�ڑ��i�I���N���j
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub OpenDBOracle(ByRef objAdoCon, provider, dataSource, user, pass)

    If DEBUG_MODE = 1 Then
        WScript.Echo "DB�ɐڑ����܂��B"
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
        WScript.Echo "DB�ɐڑ����܂����B"
    End If

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : �g�����U�N�V�����J�n
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub BeginTrans(ByRef objAdoCon)

    If DEBUG_MODE = 1 Then
        WScript.Echo "�g�����U�N�V�������J�n���܂��B"
    End If
    objAdoCon.BeginTrans

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : �R�~�b�g
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub CommitTrans(ByRef objAdoCon)

    If DEBUG_MODE = 1 Then
        WScript.Echo "�R�~�b�g���܂��B"
    End If
    objAdoCon.CommitTrans

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : ���[���o�b�N
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub RollbackTrans(ByRef objAdoCon)

    If DEBUG_MODE = 1 Then
        WScript.Echo "���[���o�b�N���܂��B"
    End If
    objAdoCon.RollbackTrans

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : DB�ؒf
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub CloseDB(ByRef objAdoCon)

    If DEBUG_MODE = 1 Then
        WScript.Echo "DB��ؒf���܂��B"
    End If
    objAdoCon.Close

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : SQL SELECT�������s���ă��R�[�h�Z�b�g���擾
' note  : �߂�l -> ���R�[�h�Z�b�g
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function ExcuteSQLgetRS(objAdoCon, strSQL)

    Dim objAdoRS  ' ���R�[�h�Z�b�g

    ' Msgbox "ExcuteSQLgetRS -> SQL: " & strSQL    

    Set objAdoRS = objAdoCon.Execute(strSQL)

    ' Msgbox objAdoRS(0).value
    ' Msgbox "EOF:" & objAdoRS.EOF
    
    Set ExcuteSQLgetRS = objAdoRS   ' Object �̂��� Set ��Y��Ȃ�����

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : ���R�[�h�Z�b�g��CSV�ɕϊ�
' note  : �߂�l -> CSV�`���̃e�L�X�g
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetCSVTextFromRS(ByRef objAdoRS)

    Dim csvText
    Dim i

    csvText = ""
    Do While objAdoRS.EOF <> True

        '�t�B�[���h�̐����[�v
        For i = 0 to objAdoRS.fields.count -1
            If i <> objAdoRS.fields.count -1 then
                csvText = csvText & objAdoRS(i).value & ", "
            Else
                csvText = csvText & objAdoRS(i).value
            End If
        Next

        '�e�L�X�g���s
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
' breif : ini�t�@�C���A�Z�N�V�������A�p�����[�^������f�[�^���擾����B
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Function GetIniData(strIniFileName, strSection, strKey)

    Dim objFSO, objIniFile, objSectionDic, strReadLine, objKeyDic, arrReadLine
    Dim strTempSection, objTempdic

    ' �t�@�C�����o�͂̒萔
    Const conForReading = 1, conForWriting = 2, conForAppending = 8
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' �t�@�C����OPEN
    Set objIniFile = objFSO.OpenTextFile(strIniFileName, conForReading, False)
    If Err.Number <> 0 Then
        ' �G���[���b�Z�[�W���o��
        wscript.echo "INI�t�@�C����: " & strIniFileName
        wscript.quit(1)
    End If
    
    ' �i�[��Dictionary�I�u�W�F�N�g�̍쐬
    Set objSectionDic = CreateObject("Scripting.Dictionary")
 
    ' �t�@�C���̃��[�hREAD
    strReadLine = objIniFile.ReadLine
    Do While objIniFile.AtEndofStream = False
        ' �X�e�[�g�����g�J�n�s������
        If (strReadLine <> " ") And (StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0) Then
            ' �Z�N�V���������擾
            strTempSection = Mid(strReadLine, 2, (Len(strReadLine) - 2))
            ' �L�[�pDictionary�I�u�W�F�N�g�쐬
            Set objKeyDic = CreateObject("Scripting.Dictionary")
            ' �t�@�C���̍ŏI�s�ɂȂ�܂�Loop
            Do While objIniFile.AtEndofStream = False
                strReadLine = objIniFile.ReadLine
                If (strReadLine <> "") And (StrComp(";", Left(strReadLine, 1)) <> 0) Then
                    ' ���̃X�e�[�g�����g�J�n�s���o��������ALoop�I��
                    If StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0 Then
                        Exit Do
                    End If
                    ' �P�Z�N�V�������̒�`��Dictionary�I�u�W�F�N�g�Ɋi�[����
                    arrReadLine = Split(strReadLine, "=", 2, vbTextCompare)
                    objKeyDic.Add UCase(arrReadLine(0)), arrReadLine(1)
                End If
            Loop
            ' �I�u�W�F�N�g�Ɋi�[����
            objSectionDic.Add UCase(strTempSection), objKeyDic
        Else
            strReadLine = objIniFile.ReadLine
        End If
    Loop
    ' �t�@�C����CLOSE
    objIniFile.Close

    ' ini�z�񂩂�w�肵���Z�N�V�����A�p�����[�^�ɑΉ�����f�[�^���擾
    strSection = UCase(strSection)  ' �啶����
    strKey = UCase(strKey)          ' �啶����

    If objSectionDic.Exists(strSection) Then
        Set objTempdic = objSectionDic.Item(strSection)
        If objTempdic.Exists(strKey) Then
            GetIniData = objSectionDic(strSection)(strKey)
            Exit Function
        End If
    End If
    GetIniData = ""

End Function
