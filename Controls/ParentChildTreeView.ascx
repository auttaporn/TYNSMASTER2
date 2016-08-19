<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ParentChildTreeView.ascx.cs" Inherits="VTS.Web.UI.ParentChildTreeView" %>

<%--<asp:ScriptManager ID="ScriptManager1" runat="server">
</asp:ScriptManager>
--%><%--<asp:UpdatePanel ID="UpdatePanel1" runat="server">--%>
   <%-- <ContentTemplate>--%>
        <asp:TreeView ID="treeView" runat="server" NodeIndent="20" ExpandDepth="0"
            NodeStyle-HorizontalPadding="2" ShowLines="true"                                                  
            ExpandImageUrl="~/Images/plus.png" CollapseImageUrl="~/Images/minus.png" 
            onselectednodechanged="treeView_SelectedNodeChanged" 
           >        
        </asp:TreeView>
   <%-- </ContentTemplate>--%>
<%--</asp:UpdatePanel>
--%>
