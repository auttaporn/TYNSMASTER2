<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="uploadSalesOrder.aspx.vb" Inherits="Default2" title="Upload Invoice" %>
<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
<div style="text-align:center;">
<table >
<tr>
<td class="col1" style="height: 35px">Upload Invoice :</td>
<td class="col2" style="height: 35px"><asp:FileUpload ID="FileUpload1" runat="server" /></td></tr>
<tr>
<td class="col1" style="height: 33px"></td>
<td class="col2" style="height: 33px"><asp:Button ID="btnUpload" runat="server" Text="Uploads" /></td></tr>
</table></div>
    
</asp:Content>

