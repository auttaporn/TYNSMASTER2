<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="StatementMonthly.aspx.vb" Inherits="Default2" title="Statement" %>
<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
    &nbsp;<table style="width: 165px">
        <tr>
            <td colspan="4" style="height: 41px; background-color: transparent;">
                <h3 style="width: 549px; height: 27px; background-color: palegreen; font-size: medium;">
                    KSO Statement Report Monthly&nbsp;</h3>
            </td>
        </tr>
        <tr>
            <td style="width: 111px; color: #ffffff; height: 26px; background-color: #003399; font-size: medium;">
                Month :</td>
            <td style="width: 71px; height: 37px; background-color: #bee1ff">
                <asp:DropDownList ID="ddlMonth" runat="server">
                    <asp:ListItem Value="01">January</asp:ListItem>
                    <asp:ListItem Value="02">Febuary</asp:ListItem>
                    <asp:ListItem Value="03">March</asp:ListItem>
                    <asp:ListItem Value="04">Apirl</asp:ListItem>
                    <asp:ListItem Value="05">May</asp:ListItem>
                    <asp:ListItem Value="06">June</asp:ListItem>
                    <asp:ListItem Value="07">July</asp:ListItem>
                    <asp:ListItem Value="08">August</asp:ListItem>
                    <asp:ListItem Value="09">September</asp:ListItem>
                    <asp:ListItem Value="10">October</asp:ListItem>
                    <asp:ListItem Value="11">November</asp:ListItem>
                    <asp:ListItem Value="12">December</asp:ListItem>
                </asp:DropDownList></td>
            <td align="right" style="width: 124px; color: #ffffff; height: 37px; background-color: #003399; font-size: medium;">
                Year :</td>
            <td style="height: 37px; background-color: #bee1ff">
                <asp:DropDownList ID="ddlYear" runat="server">
                    <asp:ListItem>2015</asp:ListItem>
                    <asp:ListItem>2016</asp:ListItem>
                    <asp:ListItem>2017</asp:ListItem>
                    <asp:ListItem>2018</asp:ListItem>
                    <asp:ListItem>2019</asp:ListItem>
                    <asp:ListItem>2020</asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="width: 111px; color: #ffffff; height: 26px; background-color: #003399; font-size: medium;">
                </td>
            <td colspan="3" style="font-size: medium; color: #ff0033; height: 26px; background-color: #bee1ff">
    <asp:Button ID="btnPrint" runat="server" Text="Print" Width="72px" /></td>
        </tr>
        <tr>
            <td style="width: 111px; height: 32px;">
            </td>
            <td style="width: 71px; height: 32px;">
            </td>
            <td style="width: 124px; height: 32px;">
            </td>
            <td style="width: 6px; height: 32px;">
            </td>
        </tr>
        <tr>
            <td style="width: 111px">
            </td>
            <td style="width: 71px">
    </td>
            <td style="width: 124px">
            </td>
            <td style="width: 6px">
            </td>
        </tr>
        <tr>
            <td style="width: 111px">
            </td>
            <td colspan="3">
                <asp:Label ID="lblmsg" runat="server" ForeColor="Red" Width="411px"></asp:Label></td>
        </tr>
    </table>
</asp:Content>

