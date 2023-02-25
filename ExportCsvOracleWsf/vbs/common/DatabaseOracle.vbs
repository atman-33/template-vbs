Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' breif : DB接続（オラクル）に関するクラス
' note  : 
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Class DatabaseOracle

    Private objAdoCon   ' ADO接続
    Private objAdoRS    ' レコードセット

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : コンストラクタ
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Private Sub Class_Initialize

        Set objAdoCon = WScript.CreateObject("ADODB.Connection")

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : ディストラクタ
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- 
    Private Sub Class_Terminate

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : DB接続（オラクル）
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub OpenDBOracle(ByVal provider,ByVal dataSource,ByVal user,ByVal pass)

        WScript.Echo "DBに接続します。"

        Dim constr
    
        constr = "Provider=" & provider & ";Data Source=" & dataSource _
                    & ";User ID=" & user & ";Password=" & pass

        WScript.Echo "OpenDBOracle 接続子: " & constr

        objAdoCon.ConnectionString = constr
        objAdoCon.Open

        WScript.Echo "DBに接続しました。"

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : トランザクション開始
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub BeginTrans()

        WScript.Echo "トランザクションを開始します。"

        objAdoCon.BeginTrans

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : コミット
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub CommitTrans()

        WScript.Echo "コミットします。"

        objAdoCon.CommitTrans

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : ロールバック
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub RollbackTrans()

        WScript.Echo "ロールバックします。"

        objAdoCon.RollbackTrans

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : DB切断
    ' note  : 
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub CloseDB()

        WScript.Echo "DBを切断します。"

        objAdoCon.Close

    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : SQL SELECT文を実行してレコードセットを取得
    ' note  : 戻り値 -> レコードセット
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Sub ExcuteSQLgetRS(ByVal strSQL)

        ' Msgbox "ExcuteSQLgetRS -> SQL: " & strSQL    

        Set objAdoRS = objAdoCon.Execute(strSQL)    ' SQL実行
    
    End Sub

    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    ' breif : レコードセットをCSVに変換
    ' note  : 戻り値 -> CSV形式のテキスト
    ' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
    Function GetCSVTextFromRS()

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

        WScript.Echo "GetCSVTextFromRS CSV変換後: " & csvText

        GetCSVTextFromRS = csvText

    End Function
End Class
