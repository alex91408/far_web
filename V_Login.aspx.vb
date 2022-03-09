
Imports System.Data.OleDb               'OleDb         
Imports System.Configuration            'WEB.CONFIG    
Imports System.Security.Cryptography    'MD5            
Imports Microsoft.VisualBasic           'Year Now Hour 
Imports System
Imports System.IO

Imports System.Data.SqlClient           ' 970904 ching2 add for  Dim dr As SqlDataReader

Partial Class V_Login
    Inherits System.Web.UI.Page

    '副程式

    '----------------------------------------------------------------------------------------------------------
    ' funCheck_acc()
    '
    ' 到 SKH.mdf之 acc table 查 acc 是否存在
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


    Public Function funCheck_pwdCount(ByVal acc1 As String) As Boolean



        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand

        Dim ErrCount As Integer


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300

        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand

        cnFU1.ConnectionString = Session("ConnectionString")

        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")
        cmdSQL1.CommandText = "SELECT SAL06 FROM SAL WHERE SAL01 = '" & acc1 & "'  and  SAL_CHECK_TYPE <> '新增'  "
        cnFU1.Open()

        Try            
            ErrCount = CType(cmdSQL1.ExecuteScalar, Integer)
            If ErrCount >= Session("ErrCount") Then
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


    '副程式
    '----------------------------------------------------------------------------------------------------------
    ' funCheck_acc()
    '
    ' 到 SKH.mdf 之 acc table 查 acc 是否存在
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    Public Function funCheck_pwd_date(ByVal acc1 As String) As Boolean


        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand

        Dim pass_date As Date


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------        

        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand

        cnFU1.ConnectionString = Session("ConnectionString")

        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")
        cmdSQL1.CommandText = "SELECT SAL10 FROM SAL WHERE SAL01 = '" & acc1 & "'  and  SAL_CHECK_TYPE <> '新增'  "
        cnFU1.Open()

        Try

            Dim pass_d As String
            pass_d = cmdSQL1.ExecuteScalar.ToString()
            If pass_d = "" Then

                '顯示訊息並且轉回首頁
                Dim script As String = ""

                script += "<script>"
                script += "alert('第一次登入請變更密碼 ');"
                script += "window.location.href='v_set_a_pwd.aspx';"      ' 轉到密碼變更網頁

                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)
                cnFU1.Close()
                Return True          '第一次登入要改密碼

            End If
            pass_date = CDate(pass_d)

            Dim web_pass_date As Integer

            web_pass_date = System.Configuration.ConfigurationManager.AppSettings("pass_date")

            If web_pass_date < DateDiff(DateInterval.Day, pass_date, Now) Then


                '顯示訊息並且轉回首頁
                Dim script As String = ""

                script += "<script>"
                script += "alert('The password expired, set password again. ');"
                script += "window.location.href='v_set_a_pwd.aspx';"      ' 轉到密碼變更網頁

                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)
                cnFU1.Close()
                Return True          '帳號超過期限要改密碼
            Else
                cnFU1.Close()
                Return False
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
                cmdSQL1.CommandText = "UPDATE SAL set SAL06 = SAL06 + 1, SAL07 = getdate() WHERE SAL01 = '" & acc1 & "'"
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

        Return 1

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

        Return 1

    End Function

    '副程式 970905  ching2    v1.3   add 使用者權限


    '----------------------------------------------------------------------------------------------------------
    ' funGet_Set()
    '
    ' 取 acc table 之使用者權限 放至 SESSIOSN( SET0 ) 中
    '
    '----------------------------------------------------------------------------------------------------------


    Public Function funGet_Set(ByVal acc1 As String) As Boolean


        Dim set0(150) As String
        Dim sSQL As String
        Dim i As Integer


        Dim cn As New SqlConnection
        Dim cmd As New SqlCommand
        Dim dr As SqlDataReader

        cn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString1").ToString()
        cmd.Connection = cn
        cmd.CommandTimeout = Session("CommandTimeout")

        '- 呼叫預存程序 
        Dim tSQL As String = ""    
        Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))
            
            cmd1.Parameters("@ao_id").Value = acc1
            cmd1.Parameters("@action").Value = "登入"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Try

            cn.Open()

            'UPDATE 錯誤的次數
            sSQL = "UPDATE SAL SET SAL06 = 0 WHERE SAL01='" & acc1 & "'"
            cmd.CommandText = sSQL
            cmd.ExecuteScalar()


            sSQL = "select SAL03, SAL04, SAL05,GRP01,t3.BRA05,t1.* from GRP t1,SAL t2,BRA t3"
            sSQL = sSQL & " WHERE t2.SAL01 = '" & acc1 & "' and t2.SAL04 = t1.GRP01 and t2.SAL03 =t3.BRA01"

            'sSQL = "select SAL03, SAL04, SAL05 from SAL WHERE SAL01 = '" & acc1 & "'"

            cmd.CommandText = sSQL

            dr = cmd.ExecuteReader

            dr.Read()

            Session("SAL03") = dr("SAL03")
            Session("SAL04") = dr("SAL04")
            Session("SAL05") = dr("SAL05")
            Session("BRA05") = dr("BRA05")
            
            If dr("v01") = "1" Then
                set0(i) = "v01"
                i = i + 1
            End If

            If dr("v01a") = "1" Then
                set0(i) = "v01a"
                i = i + 1
            End If

            If dr("v02") = "1" Then
                set0(i) = "v02"
                i = i + 1
            End If

            If dr("v02a") = "1" Then
                set0(i) = "v02a"
                i = i + 1
            End If

            If dr("v03") = "1" Then
                set0(i) = "v03"
                i = i + 1
            End If

            If dr("v03a") = "1" Then
                set0(i) = "v03a"
                i = i + 1
            End If

            If dr("v04") = "1" Then
                set0(i) = "v04"
                i = i + 1
            End If

            If dr("v04a") = "1" Then
                set0(i) = "v04a"
                i = i + 1
            End If

            If dr("v05") = "1" Then
                set0(i) = "v05"
                i = i + 1
            End If

            If dr("v05a") = "1" Then
                set0(i) = "v05a"
                i = i + 1
            End If

            If dr("v05b") = "1" Then
                set0(i) = "v05b"
                i = i + 1
            End If

            If dr("v11") = "1" Then
                set0(i) = "v11"
                i = i + 1
            End If

            If dr("v11a") = "1" Then
                set0(i) = "v11a"
                i = i + 1
            End If

            If dr("v11b") = "1" Then
                set0(i) = "v11b"
                i = i + 1
            End If

            If dr("v11c") = "1" Then
                set0(i) = "v11c"
                i = i + 1
            End If

            If dr("v11d") = "1" Then
                set0(i) = "v11d"
                i = i + 1
            End If

            If dr("v11e") = "1" Then
                set0(i) = "v11e"
                i = i + 1
            End If


            If dr("v12") = "1" Then
                set0(i) = "v12"
                i = i + 1
            End If

            If dr("v12a") = "1" Then
                set0(i) = "v12a"
                i = i + 1
            End If

            If dr("v12b") = "1" Then
                set0(i) = "v12b"
                i = i + 1
            End If

            If dr("v12c") = "1" Then
                set0(i) = "v12c"
                i = i + 1
            End If

            If dr("v12d") = "1" Then
                set0(i) = "v12d"
                i = i + 1
            End If

            If dr("v12e") = "1" Then
                set0(i) = "v12e"
                i = i + 1
            End If

            If dr("v12f") = "1" Then
                set0(i) = "v12f"
                i = i + 1
            End If

            If dr("v12g") = "1" Then
                set0(i) = "v12g"
                i = i + 1
            End If

            If dr("v12h") = "1" Then
                set0(i) = "v12h"
                i = i + 1
            End If

            If dr("v12i") = "1" Then
                set0(i) = "v12i"
                i = i + 1
            End If

            If dr("v13") = "1" Then
                set0(i) = "v13"
                i = i + 1
            End If

            If dr("v13a") = "1" Then
                set0(i) = "v13a"
                i = i + 1
            End If

            If dr("v13b") = "1" Then
                set0(i) = "v13b"
                i = i + 1
            End If

            If dr("v13c") = "1" Then
                set0(i) = "v13c"
                i = i + 1
            End If

            If dr("v13d") = "1" Then
                set0(i) = "v13d"
                i = i + 1
            End If

            If dr("v13e") = "1" Then
                set0(i) = "v13e"
                i = i + 1
            End If


            If dr("v14") = "1" Then
                set0(i) = "v14"
                i = i + 1
            End If

            If dr("v14a") = "1" Then
                set0(i) = "v14a"
                i = i + 1
            End If

            If dr("v0b") = "1" Then
                set0(i) = "v0b"
                i = i + 1
            End If

            If dr("vemail") = "1" Then
                set0(i) = "vemail"
                i = i + 1
            End If

            Session("set0") = set0

        Catch ex As SqlException
            'SQL 錯誤
            Response.Write("資料存取錯誤" & ex.Message)

        Catch ex As Exception
            ' 一舨錯誤
            Response.Write("其他錯誤")

        Finally

            If cn.State = Data.ConnectionState.Open Then
                cn.Close()
            End If

        End Try

        cn.Close()

        Return False              'sql 逾時 

    End Function

    Protected Sub butLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles butLogin.Click

        Dim strPwd As String           '密碼字串

        '登入狀態清為空白
        Me.lblStat.Text = ""

    '1030322
    If Len(Me.txtAcc.Text) = 6 And Mid(Me.txtAcc.Text, 1, 1) = "0" Then

      Me.txtAcc.Text = Mid(Me.txtAcc.Text, 2, 5)

    End If


    '----------------------------------------------------------------------------------------------------
    '判斷是否有 acc   
    If Me.txtAcc.Text = "" Or funCheck_acc(Me.txtAcc.Text) = False Then
      funWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "Account not exist !! " & vbCrLf)


      '- 呼叫預存程序 
      Dim tSQL As String = ""
      Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        cmd1.CommandType = Data.CommandType.StoredProcedure

        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

        cmd1.Parameters("@ao_id").Value = Session("acc")
        cmd1.Parameters("@action").Value = "登入-帳號不存在(" & Me.txtAcc.Text & ")"

        cn1.Open()

        Dim aaa As Integer = cmd1.ExecuteScalar

        cmd1.Dispose()
      End Using



      Me.lblStat.Text = "Account not exist !!"
      '顯示訊息並且轉回首頁
      Dim script As String = ""

      script += "<script>"
      script += "alert('Account not exist !!');"
      script += "window.location.href='V_Login.aspx';"
      script += "</script>"
      '輸出JavaScript
      ClientScript.RegisterStartupScript(Me.GetType, "", script)

      Exit Sub
    End If

    '----------------------------------------------------------------------------------------------------
    '判斷密碼錯誤是否超過三次   
    If funCheck_pwdCount(Me.txtAcc.Text) = False Then
      funWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "密碼錯誤已超過限制的次數 !! " & vbCrLf)


      '- 呼叫預存程序 
      Dim tSQL As String = ""
      Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        cmd1.CommandType = Data.CommandType.StoredProcedure

        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

        cmd1.Parameters("@ao_id").Value = Session("acc")
        cmd1.Parameters("@action").Value = "登入-密碼錯誤已超過限制的次數(" & Me.txtAcc.Text & ")"

        cn1.Open()

        Dim aaa As Integer = cmd1.ExecuteScalar

        cmd1.Dispose()
      End Using




      Me.lblStat.Text = "密碼錯誤已超過限制的次數 !!"
      '顯示訊息並且轉回首頁
      Dim script As String = ""

      script += "<script>"
      script += "alert('密碼錯誤已超過限制的次數');"
      script += "window.location.href='V_Login.aspx';"
      script += "</script>"
      '輸出JavaScript
      ClientScript.RegisterStartupScript(Me.GetType, "", script)

      Exit Sub
    End If

    '----------------------------------------------------------------------------------------------------
    If funCheck_pwd_date(Me.txtAcc.Text) = True Then    ' 要更改密碼
      Session("acc") = Me.txtAcc.Text
      funGet_Set(Me.txtAcc.Text)


      '- 呼叫預存程序 
      Dim tSQL As String = ""
      Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        cmd1.CommandType = Data.CommandType.StoredProcedure

        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

        cmd1.Parameters("@ao_id").Value = Session("acc")
        cmd1.Parameters("@action").Value = "首次登入-21.變更密碼"

        cn1.Open()

        Dim aaa As Integer = cmd1.ExecuteScalar

        cmd1.Dispose()
      End Using

      Response.Redirect("v_set_a_pwd.aspx")           ' 轉到密碼變更網頁
      Exit Sub
    End If

    strPwd = Me.txtAcc.Text + Me.txtPwd.Text
    funCheck_md5(strPwd)

    '----------------------------------------------------------------------------------------------------
    '判斷是否有(pwd)
    If Me.txtPwd.Text = "" Or funCheck_pwd(Me.txtAcc.Text, strPwd) = False Then
      funWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "Password error !! " & vbCrLf)

      '- 呼叫預存程序 
      Dim tSQL As String = ""
      Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        cmd1.CommandType = Data.CommandType.StoredProcedure

        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

        cmd1.Parameters("@ao_id").Value = Session("acc")
        cmd1.Parameters("@action").Value = "登入-密碼錯誤(" & Me.txtAcc.Text & ")"

        cn1.Open()

        Dim aaa As Integer = cmd1.ExecuteScalar

        cmd1.Dispose()
      End Using



      Me.lblStat.Text = "Password error !! "

      '顯示訊息並且轉回首頁
      Dim script As String = ""

      script += "<script>"
      script += "alert('Password error !!');"
      script += "window.location.href='V_Login.aspx';"
      script += "</script>"
      '輸出JavaScript
      ClientScript.RegisterStartupScript(Me.GetType, "", script)


      Exit Sub
    Else    ' 登入成功


      funGet_Set(Me.txtAcc.Text)  '--> Session(set0)

      funWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "Account is OK. !! " & vbCrLf)

      '970905 use   funGet_Set(Me.txtAcc.Text) control

      Session("Login") = True
      Session("acc") = Me.txtAcc.Text
      Response.Redirect("v_log0_in.aspx")
    End If


  End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Put user code to initialize the page here
        '在此加入要初始化頁面的使用者程式碼



        If Dir("e:\rec_download\" & Session("acc") & "_rec.mp3") = Session("acc") & "_rec.mp3" Then
            Kill("e:\rec_download\" & Session("acc") & "_rec.mp3")
        End If


        If Not Me.IsPostBack Then

            Session("Login") = False
            Session("acc") = ""

            '----------------------------------------------------------------------------------------------------
            '從web.config 取出字串
            '----------------------------------------------------------------------------------------------------

            Session("ConnectionString") = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()
            Session("CommandTimeout") = System.Configuration.ConfigurationManager.AppSettings("CommandTimeout").ToString()
            Session("v_menu_1") = System.Configuration.ConfigurationManager.AppSettings("v_menu_1").ToString()
            Session("ErrCount") = System.Configuration.ConfigurationManager.AppSettings("ErrCount").ToString()

            Session.Timeout = Session("CommandTimeout")

            '登入狀態清為空白
            Me.lblStat.Text = ""

            lblMDB.Text = Session("v_menu_1")
        End If


    End Sub


   
End Class
