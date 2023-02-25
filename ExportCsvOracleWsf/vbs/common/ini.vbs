Option Explicit

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