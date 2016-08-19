
Partial Class popup
    Inherits System.Web.UI.Page

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        Dim strId As String = Request.QueryString("Id")
        Dim scr As String = "window.opener.document.getElementById('" + strId + "').value = '" + formatD(Calendar1.SelectedDate) + "';window.close()"
        Page.ClientScript.RegisterStartupScript(Page.GetType(), "da", scr, True)

    End Sub
    Function formatD(ByVal s As String)
        Dim strS
        If Len(Trim(s)) < 8 Then
            formatD = " "
            Exit Function
        End If
        formatD = " "
      
        '30/12/2005  ---> 20051230
        s = Trim(s)
        strS = s.Split("/")
        If Len(strS(1)) = 1 Then strS(1) = "0" + strS(1)
        If Len(strS(0)) = 1 Then strS(0) = "0" + strS(0)
        Return strS(2) + strS(0) + strS(1)
    End Function
End Class
