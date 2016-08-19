<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="SALEDO_BYCARLINEREPORT06.aspx.vb" Inherits="SALEDO_BYCARLINEREPORT06" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
    <table style="width: 100px">
        <tr>
            <td colspan="4" style="height: 47px">
                <h3 style="width: 549px; height: 12px">
                    Print Sales Detail by Carline Report (Report No.6)</h3>
            </td>
        </tr>
        <tr>
            <td style="color: #ffffff; height: 26px; background-color: #003399">
                Bill to:</td>
            <td colspan="3" style="height: 26px; background-color: #bee1ff">
                <asp:DropDownList ID="ddlBillto" runat="server" Width="96px">
                    <asp:ListItem Value="0">ALL</asp:ListItem>
                    <asp:ListItem Value="1">YIC</asp:ListItem>
                    <asp:ListItem Value="2">CUSTOMER</asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="color: #ffffff; height: 26px; background-color: #003399">
                Company :</td>
            <td colspan="3" style="height: 26px; background-color: #bee1ff">
                <asp:DropDownList ID="ddlCompany" runat="server" Width="294px">
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="color: #ffffff; height: 26px; background-color: #003399">
                Maker Group:</td>
            <td style="height: 26px; background-color: #bee1ff">
                <asp:TextBox ID="txtCarMaker" runat="server" Width="75px"></asp:TextBox>
                <asp:Label ID="Label3" runat="server" Font-Size="8pt" ForeColor="#FF0033" Height="16px"
                    Width="66px">Ex.M01</asp:Label></td>
            <td style="width: 8px; color: #ffffff; height: 26px; background-color: #bee1ff">
            </td>
            <td style="width: 175px; height: 26px; background-color: #bee1ff">
            </td>
        </tr>
        <tr>
            <td style="color: #ffffff; background-color: #003399">
                Form:</td>
            <td style="background-color: #bee1ff">
                <asp:TextBox ID="txtFrMonth" runat="server" Width="99px"></asp:TextBox></td>
            <td style="width: 8px; color: black; background-color: #bee1ff">
                To</td>
            <td style="width: 175px; background-color: #bee1ff">
                <asp:TextBox ID="txtToMonth" runat="server" Width="98px"></asp:TextBox></td>
        </tr>
        <tr>
            <td style="color: #ffffff; background-color: #003399">
            </td>
            <td style="background-color: #bee1ff">
                <asp:Label ID="Label1" runat="server" Font-Size="8pt" ForeColor="#FF0033" Height="16px"
                    Width="94px">Ex. 01012009</asp:Label></td>
            <td style="width: 8px; color: #ffffff; background-color: #bee1ff">
            </td>
            <td style="width: 175px; background-color: #bee1ff">
                <asp:Label ID="Label2" runat="server" Font-Size="8pt" ForeColor="#FF0033" Height="16px"
                    Width="94px">Ex. 31012009</asp:Label></td>
        </tr>
        <tr>
            <td style="height: 26px">
            </td>
            <td style="width: 152px; height: 26px">
                <asp:Button ID="btnPrint" runat="server" Text="Print" Width="54px" /></td>
            <td style="width: 8px; height: 26px">
            </td>
            <td style="width: 175px; height: 26px">
            </td>
        </tr>
        <tr>
            <td style="height: 18px">
            </td>
            <td colspan="3" style="height: 18px">
                <asp:Label ID="lblmessage" runat="server" Font-Size="Small" ForeColor="Red" Width="377px"></asp:Label></td>
        </tr>
    </table>
</asp:Content>

