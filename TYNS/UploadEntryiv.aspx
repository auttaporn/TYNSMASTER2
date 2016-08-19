<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="UploadEntryiv.aspx.vb" Inherits="UploadEntryiv" title="Upload Entry" %>

<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
<script type="text/javascript" language="javascript">
function checkFileExtension(elem) {
        var filePath = elem.value;


        if(filePath.indexOf('.') == -1)
            return false;
        

        var validExtensions = new Array();
        var ext = filePath.substring(filePath.lastIndexOf('.') + 1).toLowerCase();
    

        
        validExtensions[0] = 'xls';
        validExtensions[1] = 'xlsx';
    

        for(var i = 0; i < validExtensions.length; i++) {
            if(ext == validExtensions[i])
                return true;
        }


        alert('The file extension ' + ext.toUpperCase() + ' is not allowed!');
        return false;
    }
</script>

    <br /><h1>
    Import Export Invoice Entry for TYNS</h1><br />
    <br />
    <fieldset style="border: #4682B4 2px solid; padding: 5px 5px 5px 5px; width: 90%;">
              <legend>Import Files</legend>          
    <table border="1px">
    <tr>
    <td>
    Selected Files :
    </td>
    <td>
    <asp:FileUpload ID="FileUpload1" runat="server" Width="344px" />
    </td>
    </tr>
    <tr>
    <td>
    </td>
    <td><asp:Button ID="btnUpload" runat="server" Text="Upload" />
    </td>
    </tr>
    </table>
    
    </fieldset>
    <br />
    <br />
    <fieldset style="border: #4682B4 2px solid; padding: 5px 5px 5px 5px; width: 90%;">
    <legend>Export Excel File</legend>
    <table border="1px">
    <tr>
    <td>
        Invoice Date From :</td>
    <td><asp:TextBox ID="txtDateS" runat="server"></asp:TextBox></td>
    <td>
        Invoice Date To : </td>
    <td><asp:TextBox ID="txtDateE" runat="server"></asp:TextBox></td>
    </tr>
    <tr>
    <td></td>
    <td><asp:TextBox ID="txtLotNo" runat="server" Visible="False"></asp:TextBox></td>
    <td><asp:Button ID="btnSearch" Text="Export" runat="server" /></td>
    <td></td>
    </tr>
    </table>
    </fieldset>
    <br />
    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#DEDFDE" BorderWidth="1px" CellPadding="4" ForeColor="Black" GridLines="Vertical" Width="100%" BorderStyle="None">
    <Columns>
  
    <asp:BoundField DataField="EEORID" HeaderText="Order ID" >
        <HeaderStyle Width="12px" />
    </asp:BoundField>
    <asp:BoundField DataField="EEPROD" HeaderText="PRODUCT" >
        <HeaderStyle Width="20px" />
    </asp:BoundField>
    <asp:BoundField DataField="EEQTY" HeaderText="Q'TY" >
        <HeaderStyle Width="12px" />
    </asp:BoundField>
    <asp:BoundField DataField="EEEXIV" HeaderText="EXP_ENTRY" >
        <HeaderStyle Width="15px" />
    </asp:BoundField>
    <asp:BoundField DataField="EECSNO" HeaderText="CUSTOM_NO">
        <HeaderStyle Width="15px" />
    </asp:BoundField>
    <asp:BoundField DataField="EECPDT" HeaderText="COMPLETE_DATE" >
        <HeaderStyle Width="15px" />
    </asp:BoundField>
    <asp:BoundField DataField="EELOTN" HeaderText="LOT" >
    </asp:BoundField>
    <asp:TemplateField HeaderText="Export" HeaderStyle-Width="12px">
        <ItemTemplate>
            <asp:LinkButton ID="lnkView"  runat="server"
                                    OnClick="lnkView_Click">
                                    <asp:Image ID="imgExport" runat="server" AlternateText="Export" CssClass="CommandIcon"
                                        ImageUrl="~/Images/excel.png"></asp:Image>
                                </asp:LinkButton>
        </ItemTemplate>
        <HeaderStyle Width="12px" />
    </asp:TemplateField>
    </Columns>
        <FooterStyle BackColor="#CCCC99" />
        <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
        <SelectedRowStyle BackColor="#CE5D5A" ForeColor="White" Font-Bold="True" />
        <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="White" />
        <RowStyle BackColor="#F7F7DE" />
    </asp:GridView>
    
    
    <br />
    
</asp:Content>

