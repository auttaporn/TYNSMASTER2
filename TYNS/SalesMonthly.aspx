<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="SalesMonthly.aspx.vb" Inherits="SalesMonthly" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
<script>  
    $( function() {    //$( "#datepicker" ).datepicker();  
   $("#<%=datestart.ClientID %>").datepicker({ dateFormat: 'yy-mm' });
    $("#<%=dateto.ClientID %>").datepicker({ dateFormat: 'yy-mm' });
    } );  
    </script>
    <h1>Sales Report By Month</h1>
<table>
    <tr>
        
        <td><p>Date Start: <asp:TextBox id="datestart" runat="server" ></asp:TextBox></p></td>
        
        <td><p>To:<asp:TextBox id="dateto" runat="server" ></asp:TextBox></p></td>
    </tr>
    <tr>
        <td></td>
        <td>
            <asp:Button ID="btnGenerate" runat="server" Text="Print" Width="64px" /></td>
    </tr>
</table>
</asp:Content>

