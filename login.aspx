<%@ Page Language="VB" AutoEventWireup="false" CodeFile="login.aspx.vb" Inherits="login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Web Site Name</title>
    <link rel="stylesheet" href="App_Themes/Theme1/style2.css" />
</head>
<body>
<form id="form2" runat="server">
    <div class="BodyContent">

    <div class="BorderBorder" style="left: 0px; top: 0px"><div class="BorderBL"><div></div></div>
    <div class="BorderBR"><div></div></div><div class="BorderTL"></div>
    <div class="BorderTR"><div></div></div><div class="BorderT"></div><div class="BorderR"><div></div>
    </div><div class="BorderB"><div></div></div><div class="BorderL"></div><div class="BorderC"></div><div class="Border">

        <div class="Header" style="left: 0px; top: 0px"><br />
            <h1>
    &nbsp; &nbsp; &nbsp;
</h1>
            <h2>
                &nbsp;</h2>
          </div>
        </div>
        <div class="Columns"><div class="Column1">
        <div class="Block">

            <span class="BlockHeader"><span>User Profile</span></span>
            <div class="BlockContentBorder">
            <table>
            <tr>
            <td><asp:Label ID="lb1" Text="Usermame :" runat="server"></asp:Label> </td>
            </tr>
            <tr>
            <td><asp:TextBox ID="txtusername" runat="server"></asp:TextBox></td>
            </tr>
            <tr>
            <td><asp:Label ID="lb2" Text="Password :" runat="server"></asp:Label></td>
            </tr>
            <tr>
            <td><asp:TextBox ID="txtpassword" runat="server" TextMode="Password"></asp:TextBox></td>
            </tr>
            <tr>
            <td style="text-align:right"><asp:Button ID="btnLogin" runat="server" Text="Login" /></td>
            </tr>
            </table>
                    &nbsp;</div>
                   

        </div>
        <div class="Block">

           

        </div>

        </div><div class="MainColumn"><div class="Article">

<asp:Image ID="Image1" ImageUrl="~/images/Classifieds_Image27102553152145.jpg" runat="server" Height="384px" Width="686px" />
        </div></div></div>
        <div class="Footer">
    &nbsp;
        
                            <br />
                            Version 1.0 © Copyright 2015
                            <br />
                            Thai Yazaki Corporation Limited. All Rights Reserved</div>                

    </div>
    </div>
    &nbsp;
    </form>
</body>

</html>
