<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_set_a_pwd.aspx.vb" Inherits="v_set5_pwd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
  <style type="text/css">
    .style3
    {
      text-align: center;
    }
    .style2
    {
      text-align: right;
    }
    .style5
    {
      color: #FF0000;
    }
  </style>
    
</head>




<script type="text/javascript">

function choiceUser(fieldClientID,isMutilSelect){
window.open('UserSelector.aspx?fieldClientID='+fieldClientID.id+'&mutilSelect='+isMutilSelect,'','width=220,height=220','');
}
function choiceDate(fieldClientID){
window.open('dtPicker.aspx?fieldClientID='+fieldClientID.id,'','width=220,height=220','');
}

</script>





<body bgcolor="#99ccff" style="text-align: center; background-color: #ccffff;" topmargin="3">
    <form id="form1" runat="server">
      <div class="style3">
    <asp:Image ID="Image2" runat="server" Height="86px" Width="760px" 
          ImageUrl="~/image/FAR_TOP2.JPG" style="text-align: center" />
      </div>
    <table style="width: 100%; height: 100%" border="1">
        <tr>
     <td width="97%" border="1" height="4%" class="style2">
            <asp:Label ID="Label1" runat="server" Text="Label" BackColor="#FFC080" 
            BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" 
            height="23px"></asp:Label>
            <asp:button id="Button1" runat="server" Width="50px" height="25px" Text="登出" 
            ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px"></asp:button>
          </td> </tr>
        <tr>
     <td width="97%" align="left" border="1" height="4%" 
            style="background-color: gainsboro">
          <center><b> 
              <asp:Label ID="title" runat="server" Text="【21.變更密碼】"></asp:Label></b>&nbsp;</center></td> </tr>
       <tr><td style="text-align: center; height: 16%;"> 
              <br />
              <FONT face="新細明體">
                      <br />
                        <span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA">使用者帳號</span>： &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
                        <asp:textbox id="txtAcc" Width="120px" runat="server" Height="22px" Enabled="False" BackColor="#C0C0FF"></asp:textbox>
              <span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA">
              <br />
              <br />
           ： &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;<asp:textbox id="txtPwd" Width="120px" runat="server" Height="22px" TextMode="Password" BackColor="#C0C0FF" MaxLength="8"></asp:textbox>
              <br />
              <br />
              <span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA">新密碼</span>： &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                        &nbsp; &nbsp; &nbsp;<asp:textbox id="txtPwd2" Width="120px" runat="server" Height="22px" TextMode="Password" BackColor="#C0C0FF" MaxLength="8"></asp:textbox>
              <br />
              <br />
              <span>確認新密碼<span style="font-family: 新細明體">： &nbsp; &nbsp;
            &nbsp;&nbsp; </span>&nbsp;<asp:TextBox ID="txtPwd4" runat="server" Height="22px" 
                TextMode="Password" Width="120px" BackColor="#C0C0FF" MaxLength="8"></asp:TextBox>
              <br />
              <br />
              確認新密碼font-family: 新細明體">： &nbsp; &nbsp;
            &nbsp;&nbsp; </span>&nbsp;<asp:TextBox ID="txtPwd3" runat="server" Height="22px" TextMode="Password" Width="120px" BackColor="#C0C0FF" MaxLength="8"></asp:TextBox>
              <br />
              <br />
&nbsp;
                        
                  <asp:button id="cmdRun" Width="56px" runat="server" Text="確定" ForeColor="#0000C0" BorderColor="#00AEEF"
                          BorderStyle="Solid" BackColor="#E0E0E0"></asp:button>&nbsp;
                        
                        <asp:button id="cmdClear" Width="112px" runat="server" Text="重新設定" ForeColor="#0000C0" BorderColor="#00AEEF"
                          BorderStyle="Solid" BackColor="#E0E0E0"></asp:button>&nbsp;&nbsp;&nbsp;
                  <asp:Button id="butMenu" Width="95px" runat="server" Text="回主選單" ForeColor="#0000C0" BorderColor="#00AEEF"
                    BorderStyle="Solid" BackColor="#E0E0E0"></asp:Button>
              <br />
              <br />
              <span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                              mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                              mso-bidi-language: AR-SA"><span class="style5">註：密碼變更規則<br />
              </span></span></span></FONT>
              <br />
              <asp:TextBox ID="TextBox1" runat="server" Enabled="False" Height="104px" 
                TextMode="MultiLine" Width="417px">1.不可升冪或降冪，12345678／87654321
2.不可連續，11111111／22222222
3.不可和舊密碼相同
4.密碼長度為8位，由0~9數字組合而成</asp:TextBox>
              <br />
              
        </tr>
        <tr><td style="text-align: center; height: 16%;"> 
                        
                        <asp:Label id="lblStat" runat="server" Height="22px" Width="380px" Visible="False" ForeColor="Red">lblStat</asp:Label>
              
        </tr>
        </table>
      <br />
    
  
    </form>
</body>
</html>
