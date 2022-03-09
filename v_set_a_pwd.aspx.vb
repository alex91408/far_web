
'
' 970905 ching2  add  set0
'



Imports System.Data.OleDb               'OleDb         
Imports System.Configuration            'WEB.CONFIG    
Imports System.Security.Cryptography    'MD5            
Imports Microsoft.VisualBasic           'Year Now Hour 
Imports System
Imports System.IO
Imports System.Data.SqlClient           ' 970904 ching2 add for  Dim dr As SqlDataReader



Partial Class v_set5_pwd
    Inherits System.Web.UI.Page


    '----------------------------------------------------------------------------------------------------------
    ' funCheck_acc()
    '
    ' 到 SKH.mdf 之 acc table 查 acc 是否存在
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    Public Function funCheck_acc(ByVal acc1 As String) As Boolean



        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand

        Dim Acc As String


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300
        'ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()


        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand

        cnFU1.ConnectionString = Session("ConnectionString")
        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")
        cmdSQL1.CommandText = "SELECT SAL01 FROM SAL WHERE SAL01 = '" & acc1 & "'  and  SAL_CHECK_TYPE <> '新增'  "
        cnFU1.Open()

        Try
            Acc = ""
            Acc = CType(cmdSQL1.ExecuteScalar, String)
            If Acc = "" Then
                cnFU1.Close()
                Return False          '找不到帳號
            Else
                cnFU1.Close()
                Return True
            End If
        Finally
        End Try

        cnFU1.Close()
        Return False              'sql 逾時 

    End Function



    '----------------------------------------------------------------------------------------------------------
    ' funCheck_pwd()
    '
    ' 到 SKH.mdf 之 acc table 查 pwd 是否存在
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    Public Function funCheck_pwd(ByVal acc1 As String, ByVal pwd1 As String) As Boolean

        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand

        Dim Acc As String


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300
        'ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()



        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand

        cnFU1.ConnectionString = Session("ConnectionString")
        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")
        cmdSQL1.CommandText = "SELECT SAL01  FROM SAL WHERE SAL01 = '" & acc1 & "'" & " and SAL09 = '" & pwd1 & "'  and  SAL_CHECK_TYPE <> '新增'  "
        cnFU1.Open()

        Try
            Acc = ""
            Acc = CType(cmdSQL1.ExecuteScalar, String)
            If Acc = "" Then
                cmdSQL1.CommandText = "UPDATE SAL set SAL06=SAL06+1,SAL07=getdate() WHERE SAL01 = '" & acc1 & "'"
                cmdSQL1.ExecuteScalar()
                cnFU1.Close()
                Return False          '密碼不對
            Else
                cnFU1.Close()
                Return True
            End If
        Finally
        End Try

        cnFU1.Close()
        Return False              'sql 交易逾時

    End Function

    '----------------------------------------------------------------------------------------------------------
    ' funCheck_pwd()
    '
    ' 到 SKH.mdf 之 acc table 查 pwd 是否存在
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    Public Function funCheck_pwd_rule(ByVal acc1 As String, ByVal pwd1 As String) As Boolean

        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300
        'ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()



        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand
        Dim dr As OleDbDataReader

        cnFU1.ConnectionString = Session("ConnectionString")
        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")
        cmdSQL1.CommandText = "SELECT SAL11,SAL12,SAL13  FROM SAL WHERE SAL01 = '" & acc1 & "'  and  SAL_CHECK_TYPE <> '新增'  "
        cnFU1.Open()

        dr = cmdSQL1.ExecuteReader

        dr.Read()
        Dim bpwd1 As String = dr("SAL11").ToString
        Dim bpwd2 As String = dr("SAL12").ToString
        Dim bpwd3 As String = dr("SAL13").ToString

        If pwd1 <> bpwd1 And pwd1 <> bpwd2 And pwd1 <> bpwd3 Then
            cnFU1.Close()
            Return True
        End If

        cnFU1.Close()
        Return False              'sql 交易逾時

    End Function



    '----------------------------------------------------------------------------------------------------------
    ' funCheck_md5()
    '
    ' 使用MD5檢查密碼name 是否正確
    ' Imports System.Security.Cryptography    'MD5
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    Public Function funCheck_md5(ByRef pwd1 As String) As String

        Dim bb() As Byte
        Dim i As Integer

        bb = System.Text.Encoding.ASCII.GetBytes(pwd1)
        bb = MD5CryptoServiceProvider.Create.ComputeHash(bb)

        pwd1 = ""

        For i = 0 To 16 - 1            ' Convert byte array to a text string.
            pwd1 = pwd1 & bb(i).ToString("x2")
        Next i

        Return ("OK")

    End Function




    '----------------------------------------------------------------------------------------------------------
    ' funWrite_web_Log()
    '
    ' 寫 c:\winonesys\run\web_log\   log 字串
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    Public Function funWrite_web_Log(ByVal sIn As String) As String

        '建立檔案操作物件
        Dim MyFile As System.IO.File
        Dim FileStream As System.IO.Stream
        Dim MyStreamWriter As System.IO.StreamWriter

        '建立檔案
        FileStream = File.Open("c:\winonesys\run\web_log\" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & ".txt", IO.FileMode.Append)
        MyStreamWriter = New System.IO.StreamWriter(FileStream, System.Text.Encoding.Default)

        '利用streamwriter寫入檔案
        Call MyStreamWriter.Write(sIn)

        '關閉物件
        MyStreamWriter.Close()
        FileStream.Close()
        MyFile = Nothing

        Return ("ok")

    End Function

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Label1.Text = "使用人員：" & Session("acc")

        'Put user code to initialize the page here

        '      lblMDB.Text = Session("v_menu_1")

        '-------------------------------------------------
        ' 970905 ching2 add set0
        Dim set0(150) As String
        Dim iSet As Integer

        set0 = Session("set0")

        For iSet = 0 To 150
            If set0(iSet) = "" Then
                Response.Redirect("v_login.aspx")     '避免直接輸入網址跳入
                'Exit For
            End If

            If set0(iSet) = "v0b" Then
                Exit For
            End If

        Next iSet
        '-------------------------------------------------

        If IsPostBack = False Then
            Me.txtAcc.Text = Session("acc")
        End If


    End Sub



    Private Sub butMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butMenu.Click

        '設定 session 表示使用者已經 login
        'Session("Login") = True

        '轉到主選單
        Response.Redirect("v_log0_in.aspx")
    End Sub



    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click

        txtPwd.Text() = ""
        txtPwd2.Text() = ""

    End Sub



    Private Sub cmdRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRun.Click


        Dim sSQL As String


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300
        'ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()

        Dim Conn As New Data.OleDb.OleDbConnection(Session("ConnectionString"))



        Dim objCmd As New OleDbCommand
        Dim strPwd As String
        Dim strPwd2 As String
        Dim errPwd() As String = {"12345678", "23456789", "34567890", "45678901", "56789012", "67890123", "78901234", "89012345", "90123456", "01234567", "87654321", "98765432", "09876543", "10987654", "21098765", "32109876", "43210987", "54321098", "65432109", "76543210", "11111111", "22222222", "33333333", "44444444", "55555555", "66666666", "77777777", "88888888", "99999999", "00000000"}
        Dim ii As Integer

        lblStat.Text = ""
        lblStat.Visible = False


        If IsNumeric(Me.txtPwd2.Text) = False Then

            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"
            script += "alert('密碼只能為0-9 ');"
            'script += "window.location.href='v_set_a_pwd.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Exit Sub

        End If


        If Me.txtPwd2.Text.Length <> 8 Then

            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"
            script += "alert('密碼不足8碼 ');"
            'script += "window.location.href='v_set_a_pwd.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Exit Sub

        End If

        For ii = 0 To errPwd.Length - 1
            If Me.txtPwd2.Text = errPwd(ii) Then

                '顯示訊息並且轉回首頁
                Dim script As String = ""

                script += "<script>"
                script += "alert('密碼不能為升降冪或相同數字 ');"
                'script += "window.location.href='v_set_a_pwd.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)

                Exit Sub

            End If

        Next



        If Me.txtPwd2.Text <> Me.txtPwd3.Text Then
            ' Response.Write("新密碼和再次確認密碼不同")

            ' MsgBox("1111", MsgBoxStyle.OkOnly)

            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"
            script += "alert('New password & comfirm password not the same. ');"
            'script += "window.location.href='v_set_a_pwd.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Exit Sub

        End If

        strPwd = Me.txtAcc.Text + Me.txtPwd.Text
        funCheck_md5(strPwd)


        '----------------------------------------------------------------------------------------------------
        '判斷是否有(pwd)
        If Me.txtPwd.Text = "" Or funCheck_pwd(Me.txtAcc.Text, strPwd) = False Then
            funWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "Password change error !! " & vbCrLf)
            'Me.lblStat.Text = "The Account Password error !!"
            'lblStat.Visible = True
            Session("Login") = False

            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"
            script += "alert('The Account Password error !!');"
            script += "window.location.href='v_set_a_pwd.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Exit Sub

        Else    ' 登入成功

            strPwd2 = Me.txtAcc.Text + Me.txtPwd2.Text
            funCheck_md5(strPwd2)

            If funCheck_pwd_rule(Me.txtAcc.Text, strPwd2) = False Then

                '顯示訊息並且轉回首頁
                Dim script2 As String = ""

                script2 += "<script>"
                script2 += "alert('不可與前三次密碼相同');"
                script2 += "window.location.href='v_set_a_pwd.aspx';"
                script2 += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script2)

                Exit Sub
            End If

            funWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "Account change password is ok !! " & vbCrLf)


            Try

                Dim pwd_point As String
                Dim next_pwd_point As Integer
                sSQL = "select SAL14 from SAL where SAL01 ='" & Me.txtAcc.Text & "'  and  SAL_CHECK_TYPE <> '新增'  "
                objCmd.CommandText = sSQL
                objCmd.Connection = Conn
                objCmd.Connection.Open()
                pwd_point = objCmd.ExecuteScalar.ToString()
                If pwd_point = "" Then
                    pwd_point = "1"
                End If
                next_pwd_point = (pwd_point + 1) Mod 3
                sSQL = "update SAL set SAL09 = '" & strPwd2 & "', SAL10 = getdate() ,SAL14 = " + next_pwd_point.ToString()
                If pwd_point = 1 Then
                    sSQL += " ,SAL11='" + strPwd + "'"
                ElseIf pwd_point = 2 Then
                    sSQL += " ,SAL12='" + strPwd + "'"
                ElseIf pwd_point = 3 Then
                    sSQL += " ,SAL13='" + strPwd + "'"
                End If

                sSQL += " where SAL01 ='" & Me.txtAcc.Text & "' "

                objCmd.CommandText = sSQL

                objCmd.ExecuteNonQuery()
                'Response.Write("OK")
            Catch ex As Exception
                'Response.Write(ex.Message)
                lblStat.Text = "password change fail - " & ex.Message
                lblStat.Visible = True
                objCmd.Connection.Close()
                Session("Login") = False
                Exit Sub

            End Try

            'lblStat.Text = "Password change ok,Pleas use new password !"
            'lblStat.Visible = True
            Session("Login") = False


            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"
            script += "alert('Password change ok,Pleas use new password !');"
            script += "window.location.href='V_Login.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

        End If



        objCmd.Connection.Close()

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "0.登出"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Session("acc") = ""

        Response.Redirect("V_Login.aspx")
    End Sub
End Class
