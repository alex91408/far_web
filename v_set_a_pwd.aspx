<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_set_a_pwd.aspx.vb" Inherits="v_set5_pwd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
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
      <table style="width: 760px">
        <tr>
          <td colspan="3">
            <asp:Image ID="Image1" runat="server" Height="86px" ImageUrl="~/image/FAR_TOP2.JPG"
              Width="760px" /></td>
        </tr>
        <tr><td colspan="3" style ="text-align: right ">  
            <asp:Label ID="Label1" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
            <asp:button id="Button1" runat="server" Width="50px" height="25px" Text="�n�X" ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px"></asp:button>
        </td></tr>
        <tr>
          <td colspan="3" style="background-color: #ccffff">
            <br />
                  <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB"></asp:Label><br />
              �i<strong>21.�ܧ�K�X</strong>�j<br />
            <br />
              <table style="WIDTH: 560px; HEIGHT: 108px; background-color: black;" height="108" cellSpacing="1" cellPadding="1"
                width="560" bgColor="#0080c0" border="0">
                <tr bgColor="#ffffff">
                  <td bgColor="#0080c0" height="22" style="width: 550px; background-color: #ccffff;">
                    <p class="style5" style="MARGIN: 0px 4px; LINE-HEIGHT: 150%" align="center"><font face="�s�ө���" size="2"><b style="background-color: #ccffff"></b></font>&nbsp;</p>
                  </td>
                </tr>

                <tr bgColor="#ffffff" style="font-family: Times New Roman">
                  <td vAlign="middle" align="left" style="height: 204px; width: 550px; background-color: #ccffff;" bgcolor="#99ccff">

                    <blockquote><FONT face="�s�ө���">
                      <br />
                        <span style="font-size: 12pt; font-family: �ө���; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA">�ϥΪ̱b��</span>�G &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
                        <asp:textbox id="txtAcc" Width="120px" runat="server" Height="22px" Enabled="False" BackColor="#C0C0FF"></asp:textbox>
<br />
<br />
                        <span style="font-size: 12pt; font-family: �ө���; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA">�±K�X</span>�G &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;<asp:textbox id="txtPwd" Width="120px" runat="server" Height="22px" TextMode="Password" BackColor="#C0C0FF" MaxLength="8"></asp:textbox><br />
<br />
                        <span style="font-size: 12pt; font-family: �ө���; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA">�s�K�X</span>�G &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                        &nbsp; &nbsp; &nbsp;<asp:textbox id="txtPwd2" Width="120px" runat="server" Height="22px" TextMode="Password" BackColor="#C0C0FF" MaxLength="8"></asp:textbox></FONT></blockquote>
<p>
<font face="�s�ө���"><span>&nbsp; &nbsp;&nbsp;
    &nbsp;&nbsp; &nbsp; <span>�T�{�s�K�X</span><span style="font-family: �s�ө���">�G &nbsp; &nbsp;
            &nbsp;&nbsp; </span>&nbsp;<asp:TextBox ID="txtPwd3" runat="server" Height="22px" TextMode="Password" Width="120px" BackColor="#C0C0FF" MaxLength="8"></asp:TextBox></span><br />
    <br />
    
</font>
&nbsp; &nbsp; &nbsp;&nbsp;
                        
                  <asp:button id="cmdRun" Width="56px" runat="server" Text="�T�w" ForeColor="#0000C0" BorderColor="#00AEEF"
                          BorderStyle="Solid" BackColor="#E0E0E0"></asp:button>&nbsp;
                        
                        <asp:button id="cmdClear" Width="112px" runat="server" Text="���s�]�w" ForeColor="#0000C0" BorderColor="#00AEEF"
                          BorderStyle="Solid" BackColor="#E0E0E0"></asp:button>&nbsp;&nbsp;&nbsp;
                  <asp:Button id="butMenu" Width="95px" runat="server" Text="�^�D���" ForeColor="#0000C0" BorderColor="#00AEEF"
                    BorderStyle="Solid" BackColor="#E0E0E0"></asp:Button></p>
                      <p>
                          <span style="font-size: 12pt; font-family: �ө���; mso-bidi-font-family: 'Times New Roman';
                              mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                              mso-bidi-language: AR-SA">&nbsp; &nbsp;&nbsp; ���G�K�X�ܧ�W�h</span></p>
                      <p>
                          <span lang="EN-US" style="font-size: 12pt; font-family: �ө���; mso-bidi-font-family: 'Times New Roman';
                              mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                              mso-bidi-language: AR-SA">&nbsp; &nbsp; &nbsp; 1.</span><span style="font-size: 12pt; font-family: �ө���;
                                  mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US;
                                  mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">���i�ɾ��έ����A<span lang="EN-US">12345678</span>��<span
                                      lang="EN-US">87654321</span></span></p>
                      <p>
                          <span style="font-size: 12pt; font-family: �ө���; mso-bidi-font-family: 'Times New Roman';
                              mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                              mso-bidi-language: AR-SA"><span lang="EN-US">&nbsp; &nbsp; &nbsp; 2.<span style="font-size: 12pt;
                                  font-family: �ө���; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt;
                                  mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">���i�s��A<span
                                      lang="EN-US">11111111</span>��<span lang="EN-US">22222222</span></span></span></span></p>
                      <p>
                          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 3.<span style="font-size: 12pt; font-family: �ө���;
                              mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US;
                              mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">���i�M�±K�X�ۦP</span></p>
                      <p>
                          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 4.<span style="font-size: 12pt; font-family: �ө���;
                              mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US;
                              mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">�K�X���׬�<span lang="EN-US">8</span>��A��<span
                                  lang="EN-US">0~9</span>�Ʀr�զX�Ӧ�</span><span style="font-size: 12pt; font-family: �ө���;
                                      mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US;
                                      mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA"></span><span style="font-size: 12pt;
                                          font-family: �ө���; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt;
                                          mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA"><span
                                              lang="EN-US"><span style="font-size: 12pt; font-family: �ө���; mso-bidi-font-family: 'Times New Roman';
                                                  mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                                                  mso-bidi-language: AR-SA"><span lang="EN-US"></span></span></span></span></p>
                      <p>
  <br />
  &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
                        
                        <asp:Label id="lblStat" runat="server" Height="22px" Width="380px" Visible="False" ForeColor="Red">lblStat</asp:Label><BR>
                      </p>
                  </td>
                  
                </tr>
              </table>
                  &nbsp;<br />
            <br />
            <br />
          </td>
        </tr>
        <tr>
          <td colspan="3" style="height: 21px">
          </td>
        </tr>
      </table>
      <br />
      <br />
      <br />
      <br />
      <br />
      <br />
    
  
    </form>
</body>
</html>
