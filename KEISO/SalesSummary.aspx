<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="SalesSummary.aspx.vb" Inherits="TYL_TYLINV" title="TYL Invoice Detial Listing Report" %>
<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
    <script>
        $(function () {    //$( "#datepicker" ).datepicker();  
            $("#<%=txtfromdate.ClientID%>").datepicker({ dateFormat: 'yymmdd' });
        $("#<%=txttodate.ClientID%>").datepicker({ dateFormat: 'yymmdd' });
    });
    </script>
            <table style="width: 165px">
        <tr>
            <td colspan="4" style="height: 41px; background-color: transparent;">
                <h3 style="width: 549px; height: 27px; background-color: gainsboro; font-size: medium;">
                    TYE Invoice Detial Listing Report
                </h3>
            </td>
        </tr>
        <tr>
            <td style="width: 111px; color: #ffffff; height: 26px; background-color: #003399; font-size: medium;">
                From Month :</td>
            <td style="width: 71px; height: 37px; background-color: #bee1ff">
                <asp:TextBox ID="txtfromdate" runat="server" Width="98px"></asp:TextBox></td>
            <td align="right" style="width: 124px; color: #ffffff; height: 37px; background-color: #003399; font-size: medium;">
                To Month :</td>
            <td style="height: 37px; background-color: #bee1ff">
                <asp:TextBox ID="txttodate" runat="server" Width="107px"></asp:TextBox></td>
        </tr>
        <tr>
            <td style="width: 111px; color: #ffffff; height: 26px; background-color: #003399; font-size: medium;">
                Format Date :</td>
            <td colspan="3" style="font-size: medium; color: #ff0033; height: 26px; background-color: #bee1ff">
                *YYYYMMDD (Ex.20100131)</td>
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
                <asp:Button ID="btnSubmit" runat="server" Text="Submit" /></td>
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
