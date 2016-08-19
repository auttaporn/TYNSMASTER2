<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="SALEDO_DETAILREPORT01.aspx.vb" Inherits="SALEDO_DETAILREPORT01" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
    <asp:MultiView ID="MultiView1" runat="server">
        <asp:View ID="ViewCondition" runat="server">
    <table style="width: 237px; height: 60px">
        <tr>
            <td colspan="3" style="height: 47px">
                <h3 style="width: 549px; height: 12px">
                    Print&nbsp; Sales&nbsp; Detail Report&nbsp; (Report No.1)</h3>
            </td>
        </tr>
        <tr>
            <td style="width: 31px; height: 20px">
            </td>
            <td style="width: 2px; color: #ffffff; height: 20px; background-color: #003399">
                Bill to :</td>
            <td style="height: 20px; background-color: #bee1ff">
                <asp:DropDownList ID="ddlBillto" runat="server" Width="105px">
                    <asp:ListItem Value="0">ALL</asp:ListItem>
                    <asp:ListItem Value="1">YIC</asp:ListItem>
                    <asp:ListItem Value="2">CUSTOMER</asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="width: 31px; height: 20px">
            </td>
            <td style="width: 2px; color: #ffffff; height: 20px; background-color: #003399">
                Company:</td>
            <td style="height: 20px; background-color: #bee1ff">
                <asp:DropDownList ID="ddlCompany" runat="server" Width="330px">
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="width: 31px">
            </td>
            <td style="width: 2px; color: #ffffff; background-color: #003399;">
                CarMaker:</td>
            <td style="height: 20px; background-color: #bee1ff">
                <asp:TextBox ID="txtCarMaker" runat="server" Width="87px"></asp:TextBox></td>
        </tr>
        <tr>
            <td style="width: 31px">
            </td>
            <td style="width: 2px; color: #ffffff; background-color: #003399;">
                From:</td>
            <td style="height: 20px; background-color: #bee1ff">
                <asp:TextBox ID="txtFrom" runat="server" Width="150px" MaxLength="8"></asp:TextBox>
                To&nbsp;
                <asp:TextBox ID="txtTo" runat="server" Width="150px" MaxLength="8"></asp:TextBox></td>
        </tr>
        <tr>
            <td style="width: 31px; height: 20px">
            </td>
            <td style="width: 2px; color: #ffffff; background-color: #003399;">
            </td>
            <td style="height: 20px; background-color: #bee1ff;">
                <asp:Label ID="Label1" runat="server" Font-Size="8pt" ForeColor="#FF0033" Height="16px"
                    Width="94px">Ex. 01012009</asp:Label>
                &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<asp:Label
                    ID="Label2" runat="server" Font-Size="8pt" ForeColor="#FF0033" Height="16px"
                    Width="94px">Ex. 31012009</asp:Label></td>
        </tr>
        <tr>
            <td style="width: 31px; height: 20px">
            </td>
            <td style="width: 2px; color: #ffffff; background-color: #003399;">
                F/CType:</td>
            <td style="height: 20px; background-color: #bee1ff;">
                <asp:DropDownList ID="ddlFCType" runat="server" AutoPostBack="True">
                    <asp:ListItem Value="1">Fix</asp:ListItem>
                    <asp:ListItem Value="2">Temporary</asp:ListItem>
                    <asp:ListItem Selected="True">All</asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="width: 31px; height: 20px">
            </td>
            <td style="width: 2px; color: #ffffff; background-color: #003399;">
                PriceType:</td>
            <td style="height: 20px; background-color: #bee1ff;">
                <asp:DropDownList ID="ddlSalePrice" runat="server" AutoPostBack="True">
                    <asp:ListItem Value="1">Fix</asp:ListItem>
                    <asp:ListItem Value="2">Temporary</asp:ListItem>
                    <asp:ListItem Selected="True">All</asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="width: 31px; height: 20px">
            </td>
            <td style="width: 2px; height: 20px">
                </td>
            <td style="height: 20px">
                <asp:Button ID="btnsubmit" runat="server" Text="Submit" /></td>
        </tr>
        <tr>
            <td style="width: 31px; height: 20px">
            </td>
            <td style="width: 2px; height: 20px">
            </td>
            <td style="height: 20px">
                <asp:Label ID="lblmessage" runat="server" Font-Size="Small" ForeColor="Red" Width="440px"></asp:Label></td>
        </tr>
    </table>
        </asp:View>
        &nbsp;
        <asp:View ID="ViewOpen" runat="server">
            <h3 style="width: 556px; height: 12px">
                Print&nbsp; Sales&nbsp; Detail Report&nbsp; (Report No.1)</h3>
            <br />
            &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<table style="width: 159px; height: 77px">
                <tr>
                    <td>
            <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="Filename"
                Width="557px" ForeColor="#333333" GridLines="None">
                <PagerSettings FirstPageText="" LastPageText="" NextPageText="" PreviousPageText="" />
                <FooterStyle BackColor="#507CD1" ForeColor="White" Font-Bold="True" />
                <Columns>
                    <asp:CommandField ButtonType="Button" SelectImageUrl="~/images/download.jpg" ShowSelectButton="True">
                        <ItemStyle Height="10px" Width="2px" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:CommandField>
                    <asp:BoundField DataField="Filename" HeaderText="File">
                        <ItemStyle Width="350px" Font-Size="Small" />
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
                    <td style="text-align: right">
                        <asp:Button ID="Button1" runat="server" Text="Back" PostBackUrl="~/TAPPgmGroupAdminDomestic/MainMenuAdmDomestic.aspx" /></td>
                </tr>
            </table>
            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</asp:View>
    </asp:MultiView>
</asp:Content>

