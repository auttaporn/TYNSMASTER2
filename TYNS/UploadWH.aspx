<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="UploadWH.aspx.vb" Inherits="UploadWH" title="Upload Factory Cost For W/H " %>


<asp:Content ID="Content1" ContentPlaceHolderID="Content" Runat="Server">
    <span style="font-size: 13pt; color: #0000cc"><span style="font-size: 12pt"><strong>
	Upload Factory Cost For W/H 
        <br />
        <br />

			
		 
				<table  border="0" cellspacing="3" cellpadding="3" align="center" style="width: 82%" >

	
	
		<tr>	<td align="Right"   class=Head2 style="width: 190px; color: #ffffff; background-color: #0066cc;"> Select File to Upload:  </td>
			<td class=Head3 bgcolor="lightcyan"><input id="uploadedFile" type="file" runat="server">&nbsp;
                <asp:LinkButton ID="LinkButton1" runat="server">File Format</asp:LinkButton></td></tr>
		<tr>	<td align="Right"   class=Head2 style="width: 190px; color: #ffffff; background-color: #0066cc;"> User Code:  </td>
			<td class=Head3 bgcolor="lightcyan"><input id="txtUserCode" Maxlength="9" type ="Text"  Text = "          " runat="server" disabled="disabled"></td></tr>

		<tr>	<td class=Head2  align="Right" style="width: 190px; color: #ffffff; background-color: #0066cc;">Upload now.  :</td>
			<td class=Head3 bgcolor="lightcyan"><input type=button id="upload"  value=" Upload" runat="server"></td></tr>
		<tr>	
			<td class=Head2 align="Right" style="width: 190px; color: #ffffff; background-color: #0066cc;">Message :</td>
			<td class=Head3 bgcolor="lightcyan"><asp:Label id="message" runat="server"/>&nbsp;&nbsp</td></tr>
                    <tr>
                        <td align="right" bgcolor="#ffffff" class="Head2" style="width: 190px">
                            </td>
                        <td bgcolor="#ffffff" class="Head3">
                        </td>
                    </tr>
                    <tr>
                        <td align="right" bgcolor="#ffffff" class="Head2" style="width: 190px">
                            <asp:Label ID="lblmsg" runat="server" Font-Bold="True" ForeColor="Blue"></asp:Label></td>
                        <td bgcolor="#ffffff" class="Head3">
                        </td>
                    </tr>
                    <tr>
                        <td align="center" bgcolor="#ffffff" class="Head2" colspan="2">
                            <br />
                            <asp:GridView ID="grdshow" runat="server" AutoGenerateColumns="False" BackColor="Black"
                                BorderColor="Gray" CellPadding="4" CellSpacing="1" ForeColor="#333333" GridLines="None"
                                Height="2px" HorizontalAlign="Center" >
                                <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                <Columns>
                                    <asp:TemplateField HeaderText="Filename">
                                     <ItemTemplate>
                                            <a href='<%# DataBinder.Eval(Container.DataItem,"link") %>' target="_blank">
                                                <%#DataBinder.Eval(Container.DataItem, "strfilename")%></a>
                                        </ItemTemplate>
                                        </asp:TemplateField>
                                    <asp:BoundField DataField="strFileName" HeaderText="File Name">
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="strFileType" HeaderText="File Type">
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="strFileSize" HeaderText="File Size">
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="strFileDtMod" HeaderText="Last Modified">
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <RowStyle BackColor="#EFF3FB" />
                                <EditRowStyle BackColor="#2461BF" />
                                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" />
                            </asp:GridView>
            </td>
                    </tr>
	</table>
    </strong></span></span>
				
</asp:Content>

