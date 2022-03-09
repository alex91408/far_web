
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Data.OleDb
Imports System.Configuration


Imports System.Security.Cryptography    'MD5


Imports System.Data.SqlClient           ' 970904 ching2 add for  Dim dr As SqlDataReader


Partial Class v_set_d_acc
    Inherits System.Web.UI.Page


    '副程式

    '----------------------------------------------------------------------------------------------------------
    ' funCheck_acc()
    '
    ' 到 SKH.mdf之 acc table 查 acc 是否存在
    '
    '
    '----------------------------------------------------------------------------------------------------------


    Public Function funCheck_GRP(ByVal GRP01 As String) As Boolean



        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand

        Dim GRP As String

        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand

        cnFU1.ConnectionString = Session("ConnectionString")

        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")
        cmdSQL1.CommandText = "SELECT GRP01 FROM GRP WHERE GRP01 = '" & GRP01 & "'"
        cnFU1.Open()

        Try
            GRP = ""
            GRP = CType(cmdSQL1.ExecuteScalar, String)
            If GRP = "" Then
                cnFU1.Close()
                Return False
            Else
                cnFU1.Close()
                Return True
            End If
        Finally

        End Try

        cnFU1.Close()
        Return False              'sql 逾時 

    End Function



    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        'Put user code to initialize the page here

        Label1.Text = "使用人員：" & Session("acc")




        '0--------------------------------------------------
        If Session("Login") = False Then
            Response.Redirect("v_login.aspx")
        End If

        lblMDB.Text = Session("v_menu_1")

        addGRP.Enabled = False
        '' GridView1.Columns(0).Visible = False
        '' GridView1.Columns(1).Visible = False
        SaveAsExecl.Visible = False


        'If Not Me.IsPostBack Then


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

            If set0(iSet) = "v13" Then
                Exit For
            End If

        Next iSet
        For iSet = 0 To 150
            If set0(iSet) = "v13a" Then
                SaveAsExecl.Visible = True
            End If
            If set0(iSet) = "v13b" Then
                addGRP.Enabled = True
            End If
            If set0(iSet) = "v13c" Then
                '       GridView1.Columns(0).Visible = True
            End If
            If set0(iSet) = "v13d" Then
                '        GridView1.Columns(1).Visible = True
            End If

        Next iSet
        '-------------------------------------------------
        'SqlDataSource1.ConnectionString = Session("ConnectionString")

        lblStat.Text = ""
        lblStat.Visible = False

        Dim sSQL As String = ""

        'ching2 980527 remark sSQL = "select * from GRP where GRP01 <>'" + Session("SAL04") + "' order by GRP01"
        sSQL = "select * from GRP  order by GRP01"
        SqlDataSource1.SelectCommand = sSQL

        'End If

    End Sub



    Private Sub addUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles addGRP.Click

        Dim sSQL As String

        Dim Conn As New OleDb.OleDbConnection(Session("ConnectionString"))

        Dim objCmd As New OleDbCommand
       
        '971217 ching2 add for user =""
        '顯示訊息並且轉回首頁

        If Trim(Me.txtGRP01.Text) = "" Or funCheck_GRP(Me.txtGRP01.Text) = True Then

            '- 呼叫預存程序 
            Dim tSQL2 As String = ""
            Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-群組代碼不可為空白或群組代碼已存在"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

            Dim script1 As String = ""

            script1 += "<script>"
            script1 += "alert('群組代碼不可為空白或群組代碼已存在,請重輸');"
            script1 += "window.location.href='v_set_d_acc.aspx';"
            script1 += "</script>"

            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script1)
            Exit Sub

        End If

        lblStat.Text = ""
        lblStat.Visible = False


        Try
            Dim NDate As String = ""

            NDate = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4) & ":" & Second(Now)

            'sSQL = "insert into GRP ( GRP01, GRP02) values  ( '" & txtGRP01.Text & "', '" & txtGRP02.Text & "' )"

            sSQL = "insert into GRP ( GRP_CHECK_TYPE, GRP_CHECK, GRP01, GRP01_, GRP02, GRP02_) values  (  '新增',  'Y', '" & txtGRP01.Text & "', '" & txtGRP01.Text & "', '-',  '" & txtGRP02.Text & "' )"

            objCmd.CommandText = sSQL
            objCmd.Connection = Conn
            objCmd.Connection.Open()
            objCmd.ExecuteNonQuery()
            'Response.Write("OK")
        Catch ex As Exception
            'Response.Write(ex.Message)

            '- 呼叫預存程序 
            Dim tSQL1 As String = ""
            Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-新增群組代碼失敗(" & Trim(txtGRP01.Text) & ")"

                cn1.Open()
                Dim aaa As Integer = cmd1.ExecuteScalar
                cmd1.Dispose()
            End Using

            lblStat.Text = "新增群組代碼失敗" & ex.Message
            lblStat.Visible = True
            objCmd.Connection.Close()
            Exit Sub

        End Try


        objCmd.Connection.Close()


        '顯示訊息並且轉回首頁
        Dim script As String = ""

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "13.群組管理-新增群組代碼成功(" & Trim(txtGRP01.Text) & ")"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        script += "<script>"
        script += "alert('新增群組代碼成功');"
        script += "window.location.href='v_set_d_acc.aspx';"
        script += "</script>"

        '輸出JavaScript
        ClientScript.RegisterStartupScript(Me.GetType, "", script)


    End Sub


    Protected Sub butMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles butMenu.Click


        '設定 session 表示使用者已經 login
        'Session("Login") = True
        '- 呼叫預存程序 
        Dim tSQL As String = ""
        Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "13.群組管理->回主選單"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        '轉到主選單
        Response.Redirect("v_log0_in.aspx")

    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        '處理'GridView' 的控制項 'GridView' 必須置於有 runat=server 的表單標記之中
        'Dim a = "1"
    End Sub

    Private Sub SaveAsExecl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsExecl.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "13.群組管理-save excel"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        GridView1.Columns(0).Visible = False
        GridView1.Columns(1).Visible = False
        Dim ii As Integer

        For ii = 4 To 38
            '980527 ching2 remark 
            'GridView1.Columns(ii).Visible = False
        Next



        Dim style As String = "<style> .text { mso-number-format:\@; } </script> "
        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)


        Response.Clear()
        Response.Charset = "big5"  ' 在2003EXCEL和2000EXCEL, 中文都不會變亂碼
        Response.ContentType = "Content-Language;content=zh-tw"   ' 新加的
        Response.ContentType = "application/vnd.ms-excel"
        Response.AppendHeader("Content-Disposition", "attachment; filename=Report_13.xls")

        Response.Write("<meta http-equiv=Content-Type content=text/html")
        Response.ContentEncoding = System.Text.UTF8Encoding.UTF8

        Response.Buffer = True

        GridView1.AllowPaging = False
        GridView1.AllowSorting = False
        GridView1.DataBind()



        title.RenderControl(hw)



        Dim j As Integer

        Dim o1 As String


        o1 = ""

        'For j = 0 To GridView1.Rows.Count
        For j = 5 To 32 - 1

            'For i = 2 To 34 - 1

            If CType(GridView1.Rows(j).Cells(4).FindControl("v01"), CheckBox).Checked = True Then
                txtT.Text = "T"
                txtT.RenderControl(hw)
            Else
                txtT.Text = "F"
                txtT.RenderControl(hw)
            End If


            If CType(GridView1.Rows(j).Cells(5).FindControl("v01a"), CheckBox).Checked = True Then
                txtT.Text = "T"
                txtT.RenderControl(hw)
            Else
                txtT.Text = "F"
                txtT.RenderControl(hw)
            End If


            'Next i
        Next j



        GridView1.RenderControl(hw)

        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()
    End Sub

    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender


        For Each row As GridViewRow In Me.GridView1.Rows
            ' Session("ac_12") = row.Cells(0).Text
            'CType(row.FindControl("linkbutton3"), LinkButton).Attributes.Add("onclick", "if (confirm('您確定要刪除 " + Session("ac_12") + " 嗎?')==false)  return false;")
        Next









    End Sub

 
    'Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
    Protected Sub CustomersGridView_RowCommand(ByVal source As Object, ByVal e As GridViewCommandEventArgs)



        If e.CommandName = "Update1" Then


            ' 移至   Protected Sub GridView1_RowDeleted( 
  

        End If




        If e.CommandName = "Edit" Then

            Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-修改"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If

        If e.CommandName = "Cancel" Then
            Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-取消修改"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If



        If e.CommandName = "Delete" Then

            Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-刪除(" & e.CommandArgument & ")"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using



            Dim index As Integer = Convert.ToInt32(e.CommandArgument)

        
            Dim SSQL As String = ""

            'SSQL = "delete GRP  WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"
            SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '刪除' , GRP_CHECK = 'Y' , GRP02_ = '', v01_ = '', v01a_ = '', v02_ = '', v02a_ ='', v03_ ='', v03a_ ='', v04_ ='', v04a_ ='', v05_ ='', v05a_ ='', v05b_ ='', v11_ ='', v11a_ ='', v11b_ ='', v11c_ ='', v11d_ ='', v11e_ ='', v12_ ='', v12a_ ='', v12b_ ='', v12c_ ='', v12d_ ='', v12e_ ='', v12f_ ='', v12g_ ='', v12h_ ='', v12i_ ='', v13_ ='', v13a_ ='', v13b_ ='', v13c_ ='', v13d_ ='', v13f_ ='', v14_ ='', v14a_ =''  WHERE (GRP01 = '" & Trim(index) & "')"
            SqlDataSource1.UpdateCommand = SSQL
            SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
            SqlDataSource1.Update()


        End If






        If e.CommandName = "CHECK" Then

            Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))
                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-覆核(" & Session("acc") & ")"
                cn1.Open()
                Dim aaa As Integer = cmd1.ExecuteScalar
                cmd1.Dispose()
            End Using




            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            Dim strIdx As String
            '顯示訊息並且轉回首頁
            Dim script As String = ""

            ' Retrieve the row that contains the button clicked 
            ' by the user from the Rows collection.
            Dim row As GridViewRow = GridView1.Rows(index)
            strIdx = row.Cells(5).Text
            Dim bra_check_type As String
            bra_check_type = Trim(row.Cells(1).Text)



            If bra_check_type = "修改" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '-' , GRP_CHECK = 'N' , GRP02_ = '', v01_ = '', v01a_ = '', v02_ = '', v02a_ ='', v03_ ='', v03a_ ='', v04_ ='', v04a_ ='', v05_ ='', v05a_ ='', v05b_ ='', v11_ ='', v11a_ ='', v11b_ ='', v11c_ ='', v11d_ ='', v11e_ ='', v12_ ='', v12a_ ='', v12b_ ='', v12c_ ='', v12d_ ='', v12e_ ='', v12f_ ='', v12g_ ='', v12h_ ='', v12i_ ='', v13_ ='', v13a_ ='', v13b_ ='', v13c_ ='', v13d_ ='', v13f_ ='', v14_ ='', v14a_ ='' "



                    Dim ckb As CheckBox
                    ckb = CType(row.Cells(9).FindControl("v01"), CheckBox)

                    If (ckb.Checked = True) Then

                    Else

                    End If


                    SSQL = SSQL & ", GRP02  = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v01    = '" & row.Cells(15).Text.ToString & "' "
                    SSQL = SSQL & ", v01a   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v02    = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v02a   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v03    = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v03a   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v04    = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v04a   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v05    = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v05a   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v05b   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v11    = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v11a   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v11b   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v11c   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v11d   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v11e   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12    = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12a   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12b   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12c   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12d   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12e   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12f   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12g   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12h   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v12i   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v13    = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v13a   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v13b   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v13c   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v13d   = '" & Trim(row.Cells(2).Text) & "' "
                    'SSQL = SSQL & ", v13e   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v13f   = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v14    = '" & Trim(row.Cells(2).Text) & "' "
                    SSQL = SSQL & ", v14a   = '" & Trim(row.Cells(2).Text) & "' "


                    SSQL = SSQL & " WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"


                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "13.群組管理-覆核-修改 失敗(" & Trim(row.Cells(5).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-修改 失敗,請在試一次" + Trim(row.Cells(5).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""
                Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-覆核-修改 成功(" & Trim(row.Cells(5).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-修改 成功 " + Trim(row.Cells(5).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


            If bra_check_type = "刪除" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "delete GRP  WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"

                    'SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '-' , GRP_CHECK = 'N' , GRP02_ = '', v01_ = '', v01a_ = '', v02_ = '', v02a_ ='', v03_ ='', v03a_ ='', v04_ ='', v04a_ ='', v05_ ='', v05a_ ='', v05b_ ='', v11_ ='', v11a_ ='', v11b_ ='', v11c_ ='', v11d_ ='', v11e_ ='', v12_ ='', v12a_ ='', v12b_ ='', v12c_ ='', v12d_ ='', v12e_ ='', v12f_ ='', v12g_ ='', v12h_ ='', v12i_ ='', v13_ ='', v13a_ ='', v13b_ ='', v13c_ ='', v13d_ ='', v13f_ ='', v14_ ='', v14a_ ='' "
                    'SSQL = SSQL & " WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"


                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.DeleteCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Delete()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "13.群組管理-覆核-刪除 失敗(" & Trim(row.Cells(5).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-刪除 失敗,請在試一次" + Trim(row.Cells(5).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""
                Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-覆核-刪除 成功(" & Trim(row.Cells(5).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-刪除 成功 " + Trim(row.Cells(5).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


            If bra_check_type = "新增" Then

                Try

                    Dim SSQL As String = ""

                    'SSQL = "UPDATE BRA SET BRA01_ = '' , BRA02 = '" & Trim(row.Cells(5).Text) & "' , BRA02_ = '' , BRA03 = '" & Trim(row.Cells(5).Text) & "' , BRA03_ = '' , BRA04 = '" & Trim(row.Cells(7).Text) & "' , BRA04_ = '' , BRA05 = '" & Trim(row.Cells(9).Text) & "', BRA05_ = ''  , BRA_CHECK_TYPE = '-' , BRA_CHECK = 'N' WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

                    SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '-' , GRP_CHECK = 'N' , GRP02_ = '', v01_ = '', v01a_ = '', v02_ = '', v02a_ ='', v03_ ='', v03a_ ='', v04_ ='', v04a_ ='', v05_ ='', v05a_ ='', v05b_ ='', v11_ ='', v11a_ ='', v11b_ ='', v11c_ ='', v11d_ ='', v11e_ ='', v12_ ='', v12a_ ='', v12b_ ='', v12c_ ='', v12d_ ='', v12e_ ='', v12f_ ='', v12g_ ='', v12h_ ='', v12i_ ='', v13_ ='', v13a_ ='', v13b_ ='', v13c_ ='', v13d_ ='', v13f_ ='', v14_ ='', v14a_ ='' "
                    SSQL = SSQL & " WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"


                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "13.群組管理-覆核-新增 失敗(" & Trim(row.Cells(5).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-新增 失敗,請在試一次" + Trim(row.Cells(5).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""
                Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-覆核-新增 成功(" & Trim(row.Cells(5).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-新增 成功 " + Trim(row.Cells(5).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


        End If





    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound




        If e.Row.RowType = DataControlRowType.DataRow Then

            e.Row.Cells(0).Attributes.Add("class", "text")
            e.Row.Cells(1).Attributes.Add("class", "text")
            e.Row.Cells(2).Attributes.Add("class", "text")
            e.Row.Cells(3).Attributes.Add("class", "text")
            e.Row.Cells(4).Attributes.Add("class", "text")
            e.Row.Cells(5).Attributes.Add("class", "text")
            e.Row.Cells(6).Attributes.Add("class", "text")
            e.Row.Cells(7).Attributes.Add("class", "text")
            e.Row.Cells(8).Attributes.Add("class", "text")
            e.Row.Cells(9).Attributes.Add("class", "text")
            e.Row.Cells(10).Attributes.Add("class", "text")
            e.Row.Cells(11).Attributes.Add("class", "text")

        End If








        'If e.Row.RowIndex <> -1 And Session("SaveAsExecl_flag") = "0" Then

        '    'CType(e.Row.Cells(0).Controls(0), Button).Attributes.Add("OnClick", "return confirm('Are you sure resend fax?');") '= strJava

        '    'CType(e.Row.Cells(0).Controls(0), Button).Attributes.Add("OnClick", "if ( !confirm('Are you sure resend fax?')) {return false};")
        '    'CType(e.Row.Cells(1).Controls(0), Button).Attributes.Add("OnClick", "if ( !confirm('Are you sure resend mail?')) {return false};")
        'End If

        'If e.Row.RowType = DataControlRowType.DataRow Then            
        '    e.Row.Cells(2).Attributes.Add("class", "text")
        '    e.Row.Cells(3).Attributes.Add("class", "text")

        'End If

        'If e.Row.RowIndex = GridView1.EditIndex And e.Row.RowIndex <> -1 Then
        '    'CType(e.Row.Cells(33).Controls(0), TextBox).Width = Unit.Point(100) 'Faxtel
        '    'CType(e.Row.Cells(34).Controls(0), TextBox).Width = Unit.Point(20) 'Period
        '    'CType(e.Row.Cells(35).Controls(0), TextBox).Width = Unit.Point(20) 'Fax_times
        '    'CType(e.Row.Cells(36).Controls(0), TextBox).Width = Unit.Point(20) 'maxtimes
        '    'CType(e.Row.Cells(37).Controls(0), TextBox).Width = Unit.Point(50) 'Faxerrcode
        '    'CType(e.Row.Cells(38).Controls(0), TextBox).Width = Unit.Point(20) 'EmailFlag

        '    ''- 呼叫預存程序 
        '    'Dim tSQL As String = ""
        '    'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '    '    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        '    '    cmd1.CommandType = Data.CommandType.StoredProcedure

        '    '    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
        '    '    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

        '    '    cmd1.Parameters("@ao_id").Value = Session("acc")
        '    '    cmd1.Parameters("@action").Value = "13.群組管理-編輯群組代碼成功(" & e.Row.Cells(2).Text & ")"

        '    '    cn1.Open()

        '    '    Dim aaa As Integer = cmd1.ExecuteScalar

        '    '    cmd1.Dispose()
        '    'End Using
        'Else
        '    Dim oButton As LinkButton
        '    If e.Row.RowType = DataControlRowType.DataRow Then  '找到該 Row 第一個 Cell 中的刪除鈕，即 CommandField 中的刪除鈕 
        '        'Button = FindCommandlink(e.Row.Cells(3), "Delete")
        '        'oButton.OnClientClick = "if (confirm('您確定要刪除 " + e.Row.Cells(5).Text + " 嗎?')==false)  {return false;}"

        '    End If
        'End If



    End Sub

    Function FindCommandlink(ByVal Control As Control, ByVal CommandName As String) As LinkButton
        Dim oChildCtrl As Control
        For Each oChildCtrl In Control.Controls
            If TypeOf (oChildCtrl) Is LinkButton Then
                If String.Compare(CType(oChildCtrl, LinkButton).CommandName, CommandName, True) = 0 Then Return oChildCtrl
            End If

        Next
        Return Nothing
    End Function

    Protected Sub GridView1_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView1.RowDeleted

        ''- 呼叫預存程序 
        ''Dim tSQL As String = ""
        ''Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        ''    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        ''    cmd1.CommandType = Data.CommandType.StoredProcedure

        ''    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
        ''    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

        ''    cmd1.Parameters("@ao_id").Value = Session("acc")
        ''    cmd1.Parameters("@action").Value = "13.群組管理-刪除群組代碼成功(" & e.Values(0).ToString & ")"

        ''    cn1.Open()

        ''    Dim aaa As Integer = cmd1.ExecuteScalar

        ''    cmd1.Dispose()
        ''End Using

    End Sub

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing

    End Sub

    Protected Sub GridView1_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs) Handles GridView1.RowUpdated


        
        Dim strIdx As String
        '顯示訊息並且轉回首頁

        strIdx = Trim(e.Keys.Item(0).ToString)




        Dim SSQL As String = ""

        'SSQL = "delete GRP  WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"

        'SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '刪除' , GRP_CHECK = 'Y' , GRP02_ = '', v01_ = '', v01a_ = '', v02_ = '', v02a_ ='', v03_ ='', v03a_ ='', v04_ ='', v04a_ ='', v05_ ='', v05a_ ='', v05b_ ='', v11_ ='', v11a_ ='', v11b_ ='', v11c_ ='', v11d_ ='', v11e_ ='', v12_ ='', v12a_ ='', v12b_ ='', v12c_ ='', v12d_ ='', v12e_ ='', v12f_ ='', v12g_ ='', v12h_ ='', v12i_ ='', v13_ ='', v13a_ ='', v13b_ ='', v13c_ ='', v13d_ ='', v13f_ ='', v14_ ='', v14a_ =''  WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"

        'Dim SSQL As String = ""

        'SSQL = "UPDATE BRA SET BRA01_ = '' , BRA02 = '" & Trim(e.NewValues(1).Text) & "' , BRA02_ = '' , BRA03 = '" & Trim(row.Cells(5).Text) & "' , BRA03_ = '' , BRA04 = '" & Trim(row.Cells(7).Text) & "' , BRA04_ = '' , BRA05 = '" & Trim(row.Cells(9).Text) & "', BRA05_ = ''  , BRA_CHECK_TYPE = '-' , BRA_CHECK = 'N' WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

        SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '修改' , GRP_CHECK = 'N' "

        SSQL = SSQL & ", GRP02_  = '" & Trim(e.NewValues(0).ToString) & "' "
        SSQL = SSQL & ", v01_    = '" & Trim(e.NewValues(1).ToString) & "' "
        SSQL = SSQL & ", v01a_   = '" & Trim(e.NewValues(2).ToString) & "' "
        SSQL = SSQL & ", v02_    = '" & Trim(e.NewValues(3).ToString) & "' "
        SSQL = SSQL & ", v02a_   = '" & Trim(e.NewValues(4).ToString) & "' "
        SSQL = SSQL & ", v03_    = '" & Trim(e.NewValues(5).ToString) & "' "
        SSQL = SSQL & ", v03a_   = '" & Trim(e.NewValues(6).ToString) & "' "
        SSQL = SSQL & ", v04_    = '" & Trim(e.NewValues(7).ToString) & "' "
        SSQL = SSQL & ", v04a_   = '" & Trim(e.NewValues(8).ToString) & "' "
        SSQL = SSQL & ", v05_    = '" & Trim(e.NewValues(9).ToString) & "' "
        SSQL = SSQL & ", v05a_   = '" & Trim(e.NewValues(10).ToString) & "' "
        SSQL = SSQL & ", v05b_   = '" & Trim(e.NewValues(11).ToString) & "' "
        SSQL = SSQL & ", v11_    = '" & Trim(e.NewValues(12).ToString) & "' "
        SSQL = SSQL & ", v11a_   = '" & Trim(e.NewValues(13).ToString) & "' "
        SSQL = SSQL & ", v11b_   = '" & Trim(e.NewValues(14).ToString) & "' "
        SSQL = SSQL & ", v11c_   = '" & Trim(e.NewValues(15).ToString) & "' "
        SSQL = SSQL & ", v11d_   = '" & Trim(e.NewValues(16).ToString) & "' "
        SSQL = SSQL & ", v11e_   = '" & Trim(e.NewValues(17).ToString) & "' "
        SSQL = SSQL & ", v12_    = '" & Trim(e.NewValues(18).ToString) & "' "
        SSQL = SSQL & ", v12a_   = '" & Trim(e.NewValues(19).ToString) & "' "
        SSQL = SSQL & ", v12b_   = '" & Trim(e.NewValues(20).ToString) & "' "
        SSQL = SSQL & ", v12c_   = '" & Trim(e.NewValues(21).ToString) & "' "
        SSQL = SSQL & ", v12d_   = '" & Trim(e.NewValues(22).ToString) & "' "
        SSQL = SSQL & ", v12e_   = '" & Trim(e.NewValues(23).ToString) & "' "
        SSQL = SSQL & ", v12f_   = '" & Trim(e.NewValues(24).ToString) & "' "
        SSQL = SSQL & ", v12g_   = '" & Trim(e.NewValues(25).ToString) & "' "
        SSQL = SSQL & ", v12h_   = '" & Trim(e.NewValues(26).ToString) & "' "
        SSQL = SSQL & ", v12i_   = '" & Trim(e.NewValues(27).ToString) & "' "
        SSQL = SSQL & ", v13_    = '" & Trim(e.NewValues(28).ToString) & "' "
        SSQL = SSQL & ", v13a_   = '" & Trim(e.NewValues(29).ToString) & "' "
        SSQL = SSQL & ", v13b_   = '" & Trim(e.NewValues(30).ToString) & "' "
        SSQL = SSQL & ", v13c_   = '" & Trim(e.NewValues(31).ToString) & "' "
        SSQL = SSQL & ", v13d_   = '" & Trim(e.NewValues(32).ToString) & "' "
        'SSQL = SSQL & ", v13e_   = '" & Trim(e.NewValues(33).ToString) & "' "
        SSQL = SSQL & ", v13f_   = '" & Trim(e.NewValues(33).ToString) & "' "
        SSQL = SSQL & ", v14_    = '" & Trim(e.NewValues(34).ToString) & "' "
        SSQL = SSQL & ", v14a_   = '" & Trim(e.NewValues(35).ToString) & "' "
 

        SSQL = SSQL & " WHERE (GRP01 = '" & Trim(e.Keys.Item(0).ToString) & "')"





        SqlDataSource1.UpdateCommand = SSQL
        SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
        SqlDataSource1.Update()






        '- 呼叫預存程序 
        Dim tSQL As String = ""
        Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            'cmd1.Parameters("@action").Value = "13.群組管理-編輯群組代碼成功"
            'cmd1.Parameters("@action").Value = "13.群組管理-編輯群組代碼成功(" & Trim(e.OldValues.Item(0).ToString) & "/" & Trim(e.OldValues.Item(1).ToString) & "/" & Trim(e.OldValues.Item(2).ToString) & "/" & Trim(e.OldValues.Item(3).ToString) & "/" & Trim(e.OldValues.Item(4).ToString) & "/" & Trim(e.OldValues.Item(5).ToString) & "->" & Trim(e.NewValues.Item(0).ToString) & "/" & Trim(e.NewValues.Item(1).ToString) & "/" & Trim(e.NewValues.Item(2).ToString) & "/" & Trim(e.NewValues.Item(3).ToString) & "/" & Trim(e.NewValues.Item(4).ToString) & "/" & Trim(e.NewValues.Item(5).ToString) & "/" & ")"

            Dim i As Integer
            Dim o1 As String
            Dim n1 As String


            o1 = ""
            n1 = ""
            For i = 5 To 34 - 1
                If e.OldValues.Item(i) = True Then
                    o1 = o1 + "T/"
                Else
                    o1 = o1 + "F/"
                End If

                If e.NewValues.Item(i) = True Then
                    n1 = n1 + "T/"
                Else
                    n1 = n1 + "F/"
                End If

            Next i

            cmd1.Parameters("@action").Value = "13.群組管理-編輯群組代碼成功(" & o1 & "->" & n1 & ")"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using




    End Sub

  

    'Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender

    '    For Each row As GridViewRow In Me.GridView1.Rows
    '        Session("ac_12") = row.Cells(2).Text
    '        CType(row.FindControl("linkbutton3"), LinkButton).Attributes.Add("onclick", "if (confirm('您確定要刪除 " + Session("ac_12") + " 嗎?')==false)  return false;")
    '    Next
    'End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub

    Protected Sub SqlDataSource1_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles SqlDataSource1.Selecting

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
