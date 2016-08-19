<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="MonthEnd_process.aspx.vb" Inherits="MonthEnd_process" title=" Month End Process" %>

<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
<script type="text/javascript" language="javascript">
     
	 function GetDate(name)
    {
        if (name =="txtfrm")
        {var str = '<%= txtfrm.ClientID %>'; }
        else if (name=="txtto")
        { var str = '<%= txtto.ClientID %>'; } 
              
        window.open("PopUp.aspx?Id=" + str.toString(), "Calendar", "width=240,height=230 ,top=400 , left=600 "  );
        return false;
        
    }
</script>
    <table style="width: 836px">
        <tr>
            <td colspan="4" style="height: 8px; background-color: lightgrey">
                <h4 style="font-size: medium">
                    &nbsp;Month End Process</h4>
            </td>
        </tr>
    </table>
    <table id="Table1" border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
            <td style="width: 112px; height: 19px">
            </td>
            <td align="left" style="width: 288px; height: 19px">
            </td>
            <td align="left" style="width: 370px; height: 19px">
            </td>
            <td align="center" style="width: 129px; height: 19px">
            </td>
        </tr>
        <tr>
            <td align="right" style="width: 112px; height: 30px">
                </td>
            <td align="center" style="height: 30px" colspan="2">
                <asp:RadioButtonList ID="chkview" runat="server" AutoPostBack="True" RepeatColumns="2"
                    Width="474px">
                    <asp:ListItem>Check Invoice</asp:ListItem>
                    <asp:ListItem>Check Product Code</asp:ListItem>
                </asp:RadioButtonList></td>
            <td align="center" style="width: 129px; height: 30px">
            </td>
        </tr>
    </table>
    &nbsp;<asp:MultiView ID="MultiView1" runat="server">
        <asp:View ID="ViewMonthEnd" runat="server">
            <table id="Table2" border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
            <td style="height: 29px" colspan="4">
             <div>
                 <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="Medium"
                     ForeColor="DodgerBlue" Text="CHECK INVOICE PRICE"></asp:Label>&nbsp;</div>
            </td>
                </tr>
                <tr>
                    <td align="right" style="width: 112px; height: 28px">
                        Month</td>
                    <td align="left" style="width: 288px; height: 28px">
                :
                <asp:TextBox ID="txtmonth" runat="server" MaxLength="6" Width="105px"></asp:TextBox>
                <asp:Label ID="Label10" runat="server" Font-Size="Smaller" ForeColor="Red" Text="*YYYYMM"></asp:Label></td>
                    <td align="left" style="width: 370px; height: 28px">
                    </td>
                    <td align="center" style="width: 129px; height: 28px">
                    </td>
                </tr>
                <tr>
                    <td style="width: 112px; height: 28px">
                    </td>
                    <td align="left" style="width: 288px; height: 28px">
                        &nbsp;
                &nbsp;
            <asp:Label ID="lblmsg" runat="server" ForeColor="Red"></asp:Label></td>
                    <td align="left" style="width: 370px; height: 28px">
                    </td>
                    <td align="center" style="width: 129px; height: 28px">
                    </td>
                </tr>
                <tr>
                    <td style="width: 112px; height: 29px">
                    </td>
                    <td align="left" style="width: 288px; height: 29px">
                &nbsp;
                <asp:Button ID="btnEnter" runat="server" Text="Enter" Width="56px" />
                <asp:Button ID="btncancel" runat="server" Text="Cancel" /></td>
                    <td align="left" style="width: 370px; height: 29px">
                    </td>
                    <td align="center" style="width: 129px; height: 29px">
                    </td>
                </tr>
            </table>
        </asp:View>
        <asp:View ID="ViewPRDCD" runat="server">
            <table id="Table3" border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                    <td colspan="4" style="height: 28px">
                        <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="Medium"
                            ForeColor="DodgerBlue" Text="CHECK PRODUCT CODE"></asp:Label></td>
                </tr>
                <tr>
                    <td align="right" style="width: 95px; height: 29px">
                    </td>
                    <td align="left" style="height: 29px">
                    </td>
                    <td align="left" style="height: 29px">
                    </td>
                    <td align="center" style="width: 129px; height: 29px">
                    </td>
                </tr>
                <tr>
                    <td align="right" style="width: 95px; height: 29px">
                        &nbsp;Invoice Date</td>
                    <td align="right" style="height: 29px">
                        From</td>
                    <td align="left" style="height: 29px">
                        :
                        <asp:TextBox ID="txtfrm" runat="server" MaxLength="8" Width="105px"></asp:TextBox>&nbsp;<asp:ImageButton
                            ID="ImageButton4" runat="server" Height="24px" ImageAlign="AbsBottom" ImageUrl="~/Images/calendar_icon2.jpg"
                            OnClientClick="return GetDate('txtfrm');" Width="24px" />
                        <asp:Label ID="Label3" runat="server" Font-Size="Smaller" ForeColor="Red" Text="*YYYYMMDD"></asp:Label>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtfrm"
                            ErrorMessage="**Request" Font-Bold="True"></asp:RequiredFieldValidator></td>
                    <td align="center" style="width: 129px; height: 29px">
                    </td>
                </tr>
                <tr>
                    <td style="width: 95px; height: 29px">
                    </td>
                    <td align="right" style="height: 29px">
                        To</td>
                    <td align="left" style="height: 29px">
                        :
                        <asp:TextBox ID="txtto" runat="server" MaxLength="8" Width="105px"></asp:TextBox>&nbsp;<asp:ImageButton
                            ID="ImageButton1" runat="server" Height="24px" ImageAlign="AbsBottom" ImageUrl="~/Images/calendar_icon2.jpg"
                            OnClientClick="return GetDate('txtto');" Width="24px" />
                        <asp:Label ID="Label4" runat="server" Font-Size="Smaller" ForeColor="Red" Text="*YYYYMMDD"></asp:Label>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtto"
                            ErrorMessage="**Request" Font-Bold="True"></asp:RequiredFieldValidator></td>
                    <td align="center" style="width: 129px; height: 29px">
                    </td>
                </tr>
                <tr>
                    <td style="width: 95px; height: 29px">
                    </td>
                    <td align="right" style="height: 29px">
                    </td>
                    <td align="left" style="height: 29px">
                        &nbsp;
                        <asp:Button ID="btnPRDCD" runat="server" Text="Enter" Width="56px" />&nbsp;<asp:Button
                            ID="btnPRDCDcancel" runat="server" Text="Cancel" /></td>
                    <td align="center" style="width: 129px; height: 29px">
                    </td>
                </tr>
                <tr>
                    <td style="width: 95px; height: 29px">
                    </td>
                    <td align="right" style="height: 29px">
                    </td>
                    <td align="left" style="height: 29px">
                        &nbsp;<asp:Label ID="lblerr" runat="server" ForeColor="Red"></asp:Label></td>
                    <td align="center" style="width: 129px; height: 29px">
                    </td>
                </tr>
                <tr>
                    <td style="width: 95px; height: 29px">
                    </td>
                    <td align="center" colspan="3" style="height: 29px">
                        <asp:GridView ID="gdvPRDCD" runat="server" AutoGenerateColumns="False" BackColor="Gray"
                            CellPadding="4" CellSpacing="1" ForeColor="#333333" GridLines="None">
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <Columns>
                                <asp:BoundField DataField="fac" HeaderText="Factory">
                                    <ItemStyle HorizontalAlign="Center" />
                                    <HeaderStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="tlprod" HeaderText="Part No." />
                                <asp:BoundField DataField="TLDATE" HeaderText="Invoice Date" />
                                <asp:BoundField DataField="pprdcd" HeaderText="Product Code Sale">
                                    <ItemStyle HorizontalAlign="Center" Width="100px" />
                                    <HeaderStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="tlpdcd" HeaderText="Product Code Invoice">
                                    <ItemStyle HorizontalAlign="Center" Width="100px" />
                                    <HeaderStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="cnme" HeaderText="Maker">
                                    <ItemStyle Font-Size="Small" Width="300px" />
                                </asp:BoundField>
                            </Columns>
                            <RowStyle BackColor="#EFF3FB" />
                            <EditRowStyle BackColor="#2461BF" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <AlternatingRowStyle BackColor="White" />
                        </asp:GridView>
                    </td>
                </tr>
            </table>
        </asp:View>
    </asp:MultiView>
</asp:Content>

