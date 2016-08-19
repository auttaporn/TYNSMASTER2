<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="SALEDO_PLANSALERESULTREPORT05.aspx.vb" Inherits="SALEDO_PLANSALERESULTREPORT05" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
    <table style="width: 165px">
        <tr>
            <td colspan="4" style="height: 47px">
                <h3 style="width: 549px; height: 12px">
                    Print Plan &amp; Sale Result&nbsp; Report (Report No.5)</h3>
            </td>
        </tr>
        <tr>
            <td style="color: #ffffff; background-color: #003399">
                Bill to :</td>
            <td colspan="3" style="background-color: #bee1ff">
                <asp:DropDownList ID="ddlBillto" runat="server" Width="105px">
                    <asp:ListItem Value="0">ALL</asp:ListItem>
                    <asp:ListItem Value="1">YIC</asp:ListItem>
                    <asp:ListItem Value="2">CUSTOMER</asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="color: #ffffff; background-color: #003399">
                Company:</td>
            <td colspan="3" style="background-color: #bee1ff">
                <asp:DropDownList ID="ddlCompany" runat="server" Width="282px">
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="color: #ffffff; background-color: #003399">
                From Year :</td>
            <td style="background-color: #bee1ff; width: 189px;">
                <asp:TextBox ID="txtYear" runat="server" Width="86px"></asp:TextBox></td>
            <td style="color: #ffffff; background-color: #bee1ff">
            </td>
            <td style="background-color: #bee1ff">
            </td>
        </tr>
        <tr>
            <td style="color: #ffffff; background-color: #003399">
                From Month :</td>
            <td style="background-color: #bee1ff; width: 189px;">
                <asp:DropDownList ID="ddlFrom" runat="server" Width="96px">
                    <asp:ListItem Value="1">January</asp:ListItem>
                    <asp:ListItem Value="7">July</asp:ListItem>
                </asp:DropDownList></td>
            <td style="color: #000000; background-color: #bee1ff">
                To Month :</td>
            <td style="background-color: #bee1ff">
                <asp:DropDownList ID="ddlTo" runat="server" Width="96px">
                    <asp:ListItem Value="1">January</asp:ListItem>
                    <asp:ListItem Value="2">February</asp:ListItem>
                     <asp:ListItem Value="3">March</asp:ListItem>
                    <asp:ListItem Value="4">April</asp:ListItem>
                     <asp:ListItem Value="5">May</asp:ListItem>
                    <asp:ListItem Value="6">June</asp:ListItem>
                     <asp:ListItem Value="7">July</asp:ListItem>
                    <asp:ListItem Value="8">August</asp:ListItem>
                     <asp:ListItem Value="9">September</asp:ListItem>
                    <asp:ListItem Value="10">October</asp:ListItem>
                    <asp:ListItem Value="11">November</asp:ListItem>
                    <asp:ListItem Value="12">December</asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="color: #ffffff; background-color: #003399">
                Group by</td>
            <td style="background-color: #bee1ff; width: 189px;">
                <asp:DropDownList ID="ddlGroupby" runat="server" Width="148px">
                    <asp:ListItem Value="MKP">MAKER</asp:ListItem>
                    <asp:ListItem Value="PDTP">PRODUCT TYPE</asp:ListItem>
                    <asp:ListItem Value="PDGP">PRODUCT GROUP</asp:ListItem>
                </asp:DropDownList></td>
            <td style="color: #ffffff; background-color: #bee1ff">
            </td>
            <td style="background-color: #bee1ff">
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td style="width: 189px">
                <asp:Button ID="btnSubmit" runat="server" Text="Submit" /></td>
            <td style="width: 72px">
            </td>
            <td style="width: 6px">
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td colspan="3">
                <asp:Label ID="lblmessage" runat="server" ForeColor="Red" Width="411px"></asp:Label></td>
        </tr>
    </table>
</asp:Content>

