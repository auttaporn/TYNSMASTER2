<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="SALEDO_SUMMARYREPORT02.aspx.vb" Inherits="SALEDO_SUMMARYREPORT02" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
    <asp:MultiView ID="MultiView1" runat="server">
        <asp:View ID="ViewCondition" runat="server">
            <table style="width: 273px; height: 60px">
                <tr>
                    <td colspan="4" style="height: 47px">
                        <h3 style="width: 549px; height: 12px">
                            Print Sales &nbsp;Summary by Customer (Report No.2)</h3>
                    </td>
                </tr>
                <tr>
                    <td style="width: 1px; color: #ffffff; height: 26px; background-color: #003399">
                        Bill to:</td>
                    <td colspan="3" style="height: 20px; background-color: #bee1ff">
                        <asp:DropDownList ID="ddlBillto" runat="server" Width="121px">
                            <asp:ListItem Value="0">ALL</asp:ListItem>
                            <asp:ListItem Value="1">YIC</asp:ListItem>
                            <asp:ListItem Value="2">CUSTOMER</asp:ListItem>
                            <asp:ListItem Value="2">TYNS</asp:ListItem>
                        </asp:DropDownList></td>
                </tr>
                <tr>
                    <td style="width: 1px; color: #ffffff; height: 26px; background-color: #003399">
                        Company :</td>
                    <td colspan="3" style="height: 20px; background-color: #bee1ff">
                        <asp:DropDownList ID="ddlCompany" runat="server" Width="337px">
                        </asp:DropDownList></td>
                </tr>
                <tr>
                    <td style="width: 1px; height: 26px; color: #ffffff; background-color: #003399;">
                        FromDate:</td>
                    <td style="height: 20px; background-color: #bee1ff; width: 135px;">
                        <asp:TextBox ID="txtFrom" runat="server" Width="123px"></asp:TextBox>&nbsp;
                    </td>
                    <td style="width: 1px; color: #000000; height: 26px; background-color: #bee1ff">
                        To&nbsp;</td>
                    <td style="width: 3px; height: 20px; background-color: #bee1ff">
                        <asp:TextBox ID="txtTo" runat="server" Width="130px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td style="width: 1px; height: 25px; color: #ffffff; background-color: #003399;">
                    </td>
                    <td style="height: 25px; background-color: #bee1ff; width: 135px;">
                        <asp:Label ID="Label1" runat="server" Font-Size="8pt" ForeColor="#FF0033" Height="16px"
                            Width="94px">Ex. 01012009</asp:Label>
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</td>
                    <td style="width: 1px; color: #000000; height: 25px; background-color: #bee1ff">
                    </td>
                    <td style="width: 3px; height: 25px; background-color: #bee1ff">
                        <asp:Label
                            ID="Label2" runat="server" Font-Size="8pt" ForeColor="#FF0033" Height="16px"
                            Width="94px">Ex. 31012009</asp:Label></td>
                </tr>
                <tr>
                    <td style="width: 1px; height: 20px">
                        </td>
                    <td style="height: 20px; width: 135px;">
                        <asp:Button ID="btnsubmit" runat="server" Text="Submit" /></td>
                    <td style="height: 20px">
                    </td>
                    <td style="width: 3px; height: 20px">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1px; height: 20px">
                    </td>
                    <td colspan="3" style="height: 20px">
                        <asp:Label ID="lblmessage" runat="server" Font-Size="Small" ForeColor="Red" Width="440px"></asp:Label></td>
                </tr>
            </table>
        </asp:View>
        &nbsp; &nbsp;
        <asp:View ID="ViewOpen" runat="server">
            <h3 style="width: 556px; height: 12px">
                Print Sales &nbsp;Summary by Customer (Report No.2)</h3>
            &nbsp;<table style="width: 108px">
                <tr>
                    <td style="height: 190px">
            <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="Filename"
                Width="557px" ForeColor="#333333" GridLines="None">
                <PagerSettings FirstPageText="" LastPageText="" NextPageText="" PreviousPageText="" />
                <FooterStyle BackColor="#507CD1" ForeColor="White" Font-Bold="True" />
                <Columns>
                    <asp:CommandField ButtonType="Button" SelectImageUrl="~/images/download.jpg" ShowSelectButton="True">
                        <ItemStyle Height="10px" HorizontalAlign="Center" VerticalAlign="Middle" Width="2px" />
                    </asp:CommandField>
                    <asp:BoundField DataField="Filename" HeaderText="File">
                        <ItemStyle Font-Size="Small" Width="350px" />
                    </asp:BoundField>
                </Columns>
                <RowStyle BackColor="#EFF3FB" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" />
                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="White" />
                <AlternatingRowStyle BorderColor="Black" BorderStyle="Solid" BackColor="White" />
                <EditRowStyle BackColor="#2461BF" />
            </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td style="height: 26px; text-align: right">
                        &nbsp;<asp:Button ID="Button1" runat="server" PostBackUrl="~/TAPPgmGroupAdminDomestic/MainMenuAdmDomestic.aspx"
                            Text="Back" /></td>
                </tr>
            </table>
            <br />
            <asp:Label ID="lblMsg" runat="server" Font-Size="Small" ForeColor="Red" Width="557px"></asp:Label><br />
            <br />
        </asp:View>
        &nbsp;&nbsp;
        <br />
        <br />
    </asp:MultiView>
</asp:Content>

