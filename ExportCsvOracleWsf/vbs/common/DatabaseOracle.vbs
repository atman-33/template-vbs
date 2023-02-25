Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : DB�ڑ��i�I���N���j�Ɋւ���N���X
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Class DatabaseOracle

    Private objAdoCon   ' ADO�ڑ�
    Private objAdoRS    ' ���R�[�h�Z�b�g

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : �R���X�g���N�^
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Private Sub Class_Initialize

        Set objAdoCon = WScript.CreateObject("ADODB.Connection")

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : �f�B�X�g���N�^
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- 
    Private Sub Class_Terminate

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : DB�ڑ��i�I���N���j
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub OpenDBOracle(ByVal provider,ByVal dataSource,ByVal user,ByVal pass)

        WScript.Echo "DB�ɐڑ����܂��B"

        Dim constr
    
        constr = "Provider=" & provider & ";Data Source=" & dataSource _
                    & ";User ID=" & user & ";Password=" & pass

        WScript.Echo "OpenDBOracle �ڑ��q: " & constr

        objAdoCon.ConnectionString = constr
        objAdoCon.Open

        WScript.Echo "DB�ɐڑ����܂����B"

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : �g�����U�N�V�����J�n
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub BeginTrans()

        WScript.Echo "�g�����U�N�V�������J�n���܂��B"

        objAdoCon.BeginTrans

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : �R�~�b�g
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub CommitTrans()

        WScript.Echo "�R�~�b�g���܂��B"

        objAdoCon.CommitTrans

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : ���[���o�b�N
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub RollbackTrans()

        WScript.Echo "���[���o�b�N���܂��B"

        objAdoCon.RollbackTrans

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : DB�ؒf
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub CloseDB()

        WScript.Echo "DB��ؒf���܂��B"

        objAdoCon.Close

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : SQL SELECT�������s���ă��R�[�h�Z�b�g���擾
    ' note  : �߂�l -> ���R�[�h�Z�b�g
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub ExcuteSQLgetRS(ByVal strSQL)

        ' Msgbox "ExcuteSQLgetRS -> SQL: " & strSQL    

        Set objAdoRS = objAdoCon.Execute(strSQL)    ' SQL���s
    
    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : ���R�[�h�Z�b�g��CSV�ɕϊ�
    ' note  : �߂�l -> CSV�`���̃e�L�X�g
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Function GetCSVTextFromRS()

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

        WScript.Echo "GetCSVTextFromRS CSV�ϊ���: " & csvText

        GetCSVTextFromRS = csvText

    End Function
End Class
