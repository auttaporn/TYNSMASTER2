<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="INVOICEprice.aspx.vb" Inherits="TAPPRG_INVOICEprice" title="Check Invoice with No. Price" %>
<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
<table style="width: 100%">
        <tr>
            <td colspan="4" style="height: 42px">
                <h3 style="width: 100%; height: 40px; background-color: gainsboro; vertical-align: middle;">
                    Invoice Price</h3>
            </td>
        </tr>
    <tr>
        <td align="right" style="font-weight: bold; width: 188px; color: #ffffff; font-family: Arial;
            height: 26px; background-color: #009999">
            Report :</td>
        <td style="width: 187px; height: 26px; background-color: #99cccc">
            <asp:DropDownList ID="ddlreport" runat="server" Width="119px">
                <asp:ListItem Value="non">No TAP Price</asp:ListItem>
                <asp:ListItem Value="price">Invoice Price</asp:ListItem>
                <asp:ListItem Value="FC">No TAP FC</asp:ListItem>
            </asp:DropDownList></td>
    </tr>
        <tr>
        <td style="width: 188px; color: #ffffff; height: 26px; background-color: #003399; font-family: Arial;" align="right">
                Company :</td>
            <td style="width: 187px; height: 26px; background-color: #bee1ff">
                <asp:DropDownList ID="ddlCompany" runat="server">
                    <asp:ListItem>TAP-B</asp:ListItem>
                    <asp:ListItem>TAP-C</asp:ListItem>
                    <asp:ListItem>TAP-P</asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td style="width: 188px; color: #ffffff; height: 26px; background-color: #003399; font-family: Arial;" align="right">
                From Year :</td>
            <td style="width: 187px; height: 26px; background-color: #bee1ff">
                <asp:TextBox ID="txtFrYear" runat="server" Width="75px" MaxLength="4"></asp:TextBox>
                </td>
           
        </tr>
        <tr>
            <td style="width: 188px; color: #ffffff; background-color: #003399; font-family: Arial; height: 26px;" align="right">
                From Month:</td>
            <td style="width: 187px; background-color: #bee1ff">
                <asp:DropDownList ID="ddlMonth" runat="server">
                    <asp:ListItem Value="01">January</asp:ListItem>
                    <asp:ListItem Value="02">February</asp:ListItem>
                    <asp:ListItem Value="03">March</asp:ListItem>
                    <asp:ListItem Value="04">April</asp:ListItem>
                    <asp:ListItem Value="05">May</asp:ListItem>
                    <asp:ListItem Value="06">June</asp:ListItem>
                    <asp:ListItem Value="07">July</asp:ListItem>
                    <asp:ListItem Value="08">August</asp:ListItem>
                    <asp:ListItem Value="09">September</asp:ListItem>
                    <asp:ListItem Value="10">October</asp:ListItem>
                    <asp:ListItem Value="11">November</asp:ListItem>
                    <asp:ListItem Value="12">December</asp:ListItem>
                </asp:DropDownList></td>
           
        </tr>
        <tr>
            <td style="color: #ffffff; height: 27px; width: 188px;">
            </td>
            <td style="width: 188px;  height: 27px;" >
                <asp:CheckBox ID="cbinvoiceno" runat="server" Text="Inovice No." Font-Bold="True" Font-Size="Small" ForeColor="Black"  /></td>
           
        </tr>
        <tr>
            <td style="height: 26px; width: 188px;">
            </td>
            <td style="width: 187px; height: 26px">
                <asp:Button ID="btnPrint" runat="server" Text="Print" Width="54px" OnClick="btnPrint_Click" /></td>
            
            
        </tr>
        <tr>
            <td style="height: 18px; width: 188px;">
            </td>
            <td colspan="3" style="height: 18px">
                <asp:Label ID="lblmessage" runat="server" Font-Size="Small" ForeColor="Red" Width="377px" Height="25px"></asp:Label></td>
        </tr>
    <tr>
        <td style="width: 188px; height: 55px;">
        </td>
        <td colspan="3" style="height: 55px">
            <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Size="Small" Font-Underline="True"
                ForeColor="#FF0000" Text="หมายเหตุ *"></asp:Label>
            <br />
            <asp:Label ID="Label2" runat="server" Font-Size="Small" ForeColor="Red" Text='ในกรณีที่ติด Invoice ให้สังเกตุที่ Product Code ถ้าเป็น 410,419,420,425,430,440,460,490 ถึงจะ Upload เป็น "S" นอกนั้นให้เป็น "N"'
                Width="462px"></asp:Label></td>
    </tr>
    </table>
    <table style="width: 731px; height: 67px">
        <tr>
            <td style="width: 13px">
            </td>
            <td>
            </td>
            <td valign="top">
                <asp:Label ID="lblerr" runat="server" ForeColor="Red" Font-Bold="True" Font-Size="Medium" Width="314px"></asp:Label></td>
        </tr>
        <tr>
            <td style="width: 13px">
            </td>
            <td>
                <asp:GridView ID="grdshow" runat="server" CellPadding="4" CellSpacing="1" ForeColor="#333333"
                    GridLines="None" Font-Size="Small">
                    <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                    <RowStyle BackColor="#EFF3FB" />
                    <EditRowStyle BackColor="#2461BF" />
                    <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                    <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" />
                    <Columns>
                        <asp:CommandField  ButtonType="Button" ShowSelectButton="True" />
                    </Columns>
                </asp:GridView>
            </td>
            <td valign="top">
                <asp:GridView ID="grddetail" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None" AutoGenerateColumns="False" CellSpacing="1">
                    <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
                    <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
                    <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
                    <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText="ADD" ShowSelectButton="True" />
                        <asp:BoundField DataField="ntypat" HeaderText="Part No." />
                        <asp:BoundField DataField="PDCD" HeaderText="Product" />
                        <asp:BoundField DataField="nttfdt" HeaderText="StartDate" />
                        <asp:BoundField DataField="ntaspr" HeaderText="FC" />
                    </Columns>
                </asp:GridView>
            </td>
        </tr>
    </table>
    <br />
</asp:Content>

