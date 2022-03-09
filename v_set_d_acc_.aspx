<%@ Page Language="VB" AutoEventWireup="true" EnableEventValidation = "false" CodeFile="v_set_d_acc.aspx.vb" Inherits="v_set_d_acc" %>

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





<body style="text-align: center; background-color: #ccffff;" bgcolor="#99ccff" leftmargin="0" rightmargin="0" topmargin="0">
    <form id="form1" runat="server">
    <div>
      <table style="width: 3000px; height: 132px; left: 0px; top: 0px; text-align: center;">
        <tr>
          <td colspan="3">
              <asp:Image ID="Image2" runat="server" Height="86px" ImageUrl="~/image/FAR_TOP5.JPG"
                  Width="1592px" /></td>
        </tr>
        <tr>
          <td colspan="3" align="center" style="height: 46px; text-align: center">
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              <asp:Label ID="Label1" runat="server" BackColor="#FFC080" BorderColor="Black"
                    BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" Height="23px" Text="Label"></asp:Label>
                <asp:Button ID="Button1" runat="server" BackColor="#E0E0E0" BorderColor="black" BorderStyle="Solid"
                    BorderWidth="1px" Font-Size="Medium" ForeColor="Black" Height="25px" Text="登出"
                    Width="50px" /><br />
              &nbsp;<asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB"></asp:Label>&nbsp;<br />
              <asp:Label ID="title" runat="server" Font-Bold="True" Text="【13.群組管理】"></asp:Label></td>
        </tr>
        <tr style="font-size: 12pt">
          <td align="center" colspan="3">
            <p style="text-align: center">
                &nbsp;&nbsp;&nbsp; &nbsp;
                </p>
              <p>
                  群組代碼
                  <asp:TextBox ID="txtGRP01" runat="server" BackColor="#C0C0FF" BorderColor="White" MaxLength="2"
                      Style="z-index: 108; left: 255px; top: 570px" Width="82px"></asp:TextBox>
                  群組名稱
                  <asp:TextBox ID="txtGRP02" runat="server" BackColor="#C0C0FF" BorderColor="White"
                      MaxLength="40" Style="z-index: 108; left: 255px; top: 570px" Width="82px"></asp:TextBox>&nbsp; 
              <asp:Button ID="addGRP" runat="server" BorderColor="#00AEEF" BorderStyle="Solid"
                CausesValidation="False" ForeColor="#0000C0" Height="24" Style="z-index: 107; left: 192px;
                top: 569px" Text="新增" Width="51px" /></p>
            <p>
              &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<asp:Label ID="lblStat" runat="server" ForeColor="Blue" Height="22px" Style="z-index: 106;
                left: 116px; top: 603px" Visible="False" Width="534px">lblStat</asp:Label></p>
              <p>
                  <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                  
                    ProviderName="<%$ ConnectionStrings:FARConnectionString1.ProviderName %>"
                    SelectCommand="SELECT * FROM [GRP]" 
                    DeleteCommand="UPDATE GRP SET GRP_CHECK_TYPE = '刪除' , GRP_CHECK = 'Y' where (GRP01 ='99A') "
                    OldValuesParameterFormatString="original_{0}"                  
                   
                    UpdateCommand="UPDATE GRP SET GRP_CHECK_TYPE = '修改' , GRP_CHECK = 'Y' , GRP02_ = @GRP02, v01_ = @v01, v01a_ = @v01a, v02_ = @v02, v02a_ =@v02a, v03_ =@v03, v03a_ =@v03a, v04_ =@v04, v04a_ =@v04a, v05_ =@v05, v05a_ =@v05a, v05b_ =@v05b, v11_ =@v11, v11a_ =@v11a, v11b_ =@v11b, v11c_ =@v11c, v11d_ =@v11d, v11e_ = @v11e, v12_ =@v12, v12a_ =@v12a, v12b_ =@v12b, v12c_ =@v12c, v12d_ =@v12d, v12e_ =@v12e, v12f_ =@v12f, v12g_ =@v12g, v12h_ =@v12h, v12i_ =@v12i, v13_ =@v13, v13a_ =@v13a, v13b_ =@v13b, v13c_ =@v13c, v13d_ =@v13d,  v14_ =@v14, v14a_ =@v14a where GRP01 =  '97' " 
                   
                    InsertCommand="INSERT INTO [GRP] ([GRP01]) VALUES (@GRP01) " 
                    
                    ConnectionString="<%$ ConnectionStrings:FARConnectionString1 %>" >
               
               
                      <DeleteParameters>
                          <asp:Parameter Name="GRP01_" />
                       </DeleteParameters>
                      
                      <UpdateParameters>
                          <asp:Parameter Name="GRP02_" />
                          <asp:Parameter Name="v01_" />
                          <asp:Parameter Name="v01a_" />
                          <asp:Parameter Name="v02" />
                          <asp:Parameter Name="v02a" />
                          <asp:Parameter Name="v03_" />
                          <asp:Parameter Name="v03a_" />
                          <asp:Parameter Name="v04_" />
                          <asp:Parameter Name="v04a_" />
                          <asp:Parameter Name="v05_" />
                          <asp:Parameter Name="v05a_" />
                          <asp:Parameter Name="v05b_" />
                          <asp:Parameter Name="v11_" />
                          <asp:Parameter Name="v11a_" />
                          <asp:Parameter Name="v11b_" />
                          <asp:Parameter Name="v11c_" />
                          <asp:Parameter Name="v11d_" />
                          <asp:Parameter Name="v11e_" />
                          <asp:Parameter Name="v12_" />
                          <asp:Parameter Name="v12a_" />
                           <asp:Parameter Name="v12b_" />
                          <asp:Parameter Name="v12c_" />
                          <asp:Parameter Name="v12d_" />
                          <asp:Parameter Name="v12e_" />
                          <asp:Parameter Name="v12f_" />
                          <asp:Parameter Name="v12g_" />
                          <asp:Parameter Name="v12h_" />
                          <asp:Parameter Name="v12i_" />
                          <asp:Parameter Name="v13_" />
                          <asp:Parameter Name="v13a_" />
                          <asp:Parameter Name="v13b_" />
                          <asp:Parameter Name="v13c_" />
                          <asp:Parameter Name="v13d_" />
                         
                          <asp:Parameter Name="v14_" />
                          <asp:Parameter Name="v14a_" />
                          <asp:Parameter Name="GRP01" />
             
                     </UpdateParameters>
                      
                  </asp:SqlDataSource>
                  &nbsp;</p>
              <p>
                  <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
                      AutoGenerateColumns="False" DataKeyNames="GRP01" DataSourceID="SqlDataSource1" PageSize="50" BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" GridLines="Vertical">
                      <Columns>
                          <asp:CommandField ShowDeleteButton="True" ShowEditButton="True" />
                        <asp:ButtonField Text="覆核" CommandName="CHECK" >
                          <ItemStyle Width="40px" />
                        </asp:ButtonField>
    
    
                        <asp:BoundField DataField="GRP_CHECK_TYPE" HeaderText="覆核_型別" ReadOnly="True"  SortExpression="GRP_CHECK_TYPE" >
                             <ItemStyle Width="40px" ForeColor="Red" />
                        </asp:BoundField>
                        
                        <asp:BoundField DataField="GRP_CHECK" HeaderText="覆核" ReadOnly="True" SortExpression="GRP_CHECK" >
                            <ItemStyle ForeColor="Red" />
                        </asp:BoundField>
 
 
                        <asp:TemplateField InsertVisible="False" ShowHeader="False">
                        
                            <EditItemTemplate>
                                <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" CommandName="Update"
                                    Text="更新"  Width="40px" >
                                </asp:LinkButton>
                                
                                <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" CommandName="Cancel"
                                    Text="取消"  Width="40px">
                                </asp:LinkButton>
                                
                            </EditItemTemplate>
                            
                            <ItemTemplate>
                                <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" CommandName="Edit"
                                    Text="編輯" Width="40px" > 
                                </asp:LinkButton>
                                    
                            </ItemTemplate>
                            
                        </asp:TemplateField>
                        

                      <asp:TemplateField InsertVisible="False" ShowHeader="False">
                            <ItemTemplate>
                                <asp:LinkButton ID="LinkButton3" runat="server" CausesValidation="False" CommandName="Delete"
                                    Text="刪除" CommandArgument='<%#Eval("GRP01")%>' Width="40px" >  
                                </asp:LinkButton>
                                </ItemTemplate>
                        </asp:TemplateField>
                   
         
 
                    
                    
                     
                          <asp:BoundField DataField="GRP01" HeaderText="群組代碼" SortExpression="GRP01"  ReadOnly="True" />

                           
                          <asp:BoundField DataField="GRP01_" HeaderText="群組代碼_覆核" SortExpression="GRP01_" ReadOnly="True" />
                    
                          <asp:BoundField DataField="GRP02" HeaderText="群組名稱" SortExpression="GRP02" >
                               <ItemStyle Width="150px" />
                               <HeaderStyle Wrap="True" />
                          </asp:BoundField>
 
                           <asp:BoundField DataField="GRP02_" HeaderText="群組名稱_覆核" SortExpression="GRP02_" ReadOnly="True">
                              <ItemStyle Width="150px" Wrap="True" />
                              <HeaderStyle Wrap="True" /> 
                          </asp:BoundField>
                   
                   
                          <asp:CheckBoxField DataField="v01" HeaderText="1.電話錄音線路使用統計表" SortExpression="v01" />
                          <asp:CheckBoxField DataField="v01_" HeaderText="1.電話錄音線路使用統計表_覆核" SortExpression="v01_"  ReadOnly="True"/>
 
                          <asp:CheckBoxField DataField="v01a" HeaderText="1.儲存" SortExpression="v01a" />
                          <asp:CheckBoxField DataField="v01a_" HeaderText="1.儲存_覆核" SortExpression="v01a"  ReadOnly="True"/>
 
                          <asp:CheckBoxField DataField="v02" HeaderText="2.電話錄音進線時段統計表" SortExpression="v02" />
                          <asp:CheckBoxField DataField="v02_" HeaderText="2.電話錄音進線時段統計表_覆核" SortExpression="v02_"  ReadOnly="True"/>
 
                          <asp:CheckBoxField DataField="v02a" HeaderText="2.儲存" SortExpression="v02a" />
                          <asp:CheckBoxField DataField="v02a_" HeaderText="2.儲存_覆核" SortExpression="v02a_"  ReadOnly="True"/>
 
                          <asp:CheckBoxField DataField="v03" HeaderText="3.各項服務統計表" SortExpression="v03" />
                          <asp:CheckBoxField DataField="v03_" HeaderText="3.各項服務統計表_覆核" SortExpression="v03"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v03a" HeaderText="3.儲存" SortExpression="v03a" />
                          <asp:CheckBoxField DataField="v03a_" HeaderText="3.儲存_覆核" SortExpression="v03a"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v04" HeaderText="4. 各分行電話錄音統計表" SortExpression="v04" />
                          <asp:CheckBoxField DataField="v04_" HeaderText="4. 各分行電話錄音統計表_覆核" SortExpression="v04"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v04a" HeaderText="4.儲存" SortExpression="v04a" />
                          <asp:CheckBoxField DataField="v04a_" HeaderText="4.儲存_覆核" SortExpression="v04a"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v05" HeaderText="5.各項服務明細查詢" SortExpression="v05" />
                          <asp:CheckBoxField DataField="v05_" HeaderText="5.各項服務明細查詢_覆核" SortExpression="v05"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v05a" HeaderText="5.儲存" SortExpression="v05a" />
                          <asp:CheckBoxField DataField="v05a_" HeaderText="5.儲存_覆核" SortExpression="v05a"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v05b" HeaderText="5.下載" SortExpression="v05b" />
                          <asp:CheckBoxField DataField="v05b_" HeaderText="5.下載_覆核" SortExpression="v05b"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11" HeaderText="11.分行資料維護" SortExpression="v11" />
                          <asp:CheckBoxField DataField="v11_" HeaderText="11.分行資料維護_覆核" SortExpression="v11"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11a" HeaderText="11.儲存" SortExpression="v11a" />
                          <asp:CheckBoxField DataField="v11a_" HeaderText="11.儲存_覆核" SortExpression="v11a"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11b" HeaderText="11.新增" SortExpression="v11b" />
                          <asp:CheckBoxField DataField="v11b_" HeaderText="11.新增_覆核" SortExpression="v11b"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11c" HeaderText="11.刪除" SortExpression="v11c" />
                          <asp:CheckBoxField DataField="v11c_" HeaderText="11.刪除_覆核" SortExpression="v11c"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11d" HeaderText="11.編輯" SortExpression="v11d" />
                          <asp:CheckBoxField DataField="v11d_" HeaderText="11.編輯_覆核" SortExpression="v11d"  ReadOnly="True"/>
                          
                          <asp:CheckBoxField DataField="v11e" HeaderText="11.覆核" SortExpression="v11e" />
                          <asp:CheckBoxField DataField="v11e_" HeaderText="11.覆核_覆核" SortExpression="v11e"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12" HeaderText="12.系統使用者資料維護" SortExpression="v12" />
                          <asp:CheckBoxField DataField="v12_" HeaderText="12.系統使用者資料維護_覆核" SortExpression="v12"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12a" HeaderText="12.儲存" SortExpression="v12a" />
                          <asp:CheckBoxField DataField="v12a_" HeaderText="12.儲存_覆核" SortExpression="v12a"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12b" HeaderText="12.查詢" SortExpression="v12b" />
                          <asp:CheckBoxField DataField="v12b_" HeaderText="12.查詢_覆核" SortExpression="v12b"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12c" HeaderText="12.新增" SortExpression="v12c" />
                          <asp:CheckBoxField DataField="v12c_" HeaderText="12.新增_覆核" SortExpression="v12c"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12d" HeaderText="12.密碼重置" SortExpression="v12d" />
                          <asp:CheckBoxField DataField="v12d_" HeaderText="12.密碼重置_覆核" SortExpression="v12d"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12e" HeaderText="12.編輯" SortExpression="v12e" />
                          <asp:CheckBoxField DataField="v12e_" HeaderText="12.編輯_覆核" SortExpression="v12e"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12f" HeaderText="12.刪除" SortExpression="v12f" />
                          <asp:CheckBoxField DataField="v12f_" HeaderText="12.刪除_覆核" SortExpression="v12f"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12g" HeaderText="12.功能群組" SortExpression="v12g" />
                          <asp:CheckBoxField DataField="v12g_" HeaderText="12.功能群組_覆核" SortExpression="v12g"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12h" HeaderText="12.權限等級" SortExpression="v12h" />
                          <asp:CheckBoxField DataField="v12h_" HeaderText="12.權限等級_覆核" SortExpression="v12h"  ReadOnly="True"/>
                          
                          <asp:CheckBoxField DataField="v12i" HeaderText="12.覆核" SortExpression="v12i" />
                          <asp:CheckBoxField DataField="v12i_" HeaderText="12.覆核_覆核" SortExpression="v12i"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13" HeaderText="13.群組維護" SortExpression="v13" />
                          <asp:CheckBoxField DataField="v13_" HeaderText="13.群組維護_覆核" SortExpression="v13"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13a" HeaderText="13.儲存" SortExpression="v13a" />
                          <asp:CheckBoxField DataField="v13a_" HeaderText="13.儲存_覆核" SortExpression="v13a"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13b" HeaderText="13.新增" SortExpression="v13b" />
                          <asp:CheckBoxField DataField="v13b_" HeaderText="13.新增_覆核" SortExpression="v13b"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13c" HeaderText="13.刪除" SortExpression="v13c" />
                          <asp:CheckBoxField DataField="v13c_" HeaderText="13.刪除_覆核" SortExpression="v13c"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13d" HeaderText="13.編輯" SortExpression="v13d" />
                          <asp:CheckBoxField DataField="v13d_" HeaderText="13.編輯_覆核" SortExpression="v13d"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13e" HeaderText="v13e" SortExpression="v13e" Visible="False" />
                          <asp:CheckBoxField DataField="v13e_" HeaderText="v13e_覆核" SortExpression="v13e" Visible="False"  ReadOnly="True"/>
                          
                          <asp:CheckBoxField DataField="v13f" HeaderText="v13.覆核" SortExpression="v13f" />
                          <asp:CheckBoxField DataField="v13f_" HeaderText="v13.覆核_覆核" SortExpression="v13f" ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v14" HeaderText="14.使用者&amp;權限異動明細查詢" SortExpression="v14" />
                          <asp:CheckBoxField DataField="v14_" HeaderText="14.使用者&amp;權限異動明細查詢_覆核" SortExpression="v14"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v14a" HeaderText="14.儲存" SortExpression="v14a" />
                          <asp:CheckBoxField DataField="v14a_" HeaderText="14.儲存_覆核" SortExpression="v14a"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v0b" HeaderText="v0b" SortExpression="v0b" Visible="False" />

                          <asp:CheckBoxField DataField="vemail" HeaderText="email" SortExpression="vemail" Visible="False" />

                    </Columns>
                      <HeaderStyle VerticalAlign="Top" BackColor="#000084" Font-Bold="True" ForeColor="White" />
                      <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
                      <RowStyle BackColor="#EEEEEE" ForeColor="Black" />
                      <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
                      <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                      <AlternatingRowStyle BackColor="#DCDCDC" />
                  </asp:GridView>
                  &nbsp;</p>
              <p>
                  <asp:Button ID="SaveAsExecl" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="儲存Excel" Width="95px" Height="25px" />
                <asp:Button ID="butMenu" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                BorderStyle="Solid" ForeColor="#0000C0" Text="回主選單" Width="95px" Height="25px" />
              </p>
              <p>
                  &nbsp; &nbsp; &nbsp;&nbsp;</p>
              <p>
                  &nbsp;</p>
              <p>
                  &nbsp;&nbsp;</p>
            <p>
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</p>
          </td>
        </tr>
      </table>
    
    </div>
    <asp:TextBox ID="txtT" runat="server" Visible="False"></asp:TextBox>
    </form>
</body>
</html>
