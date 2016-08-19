Imports System.Data
Imports System.Data.OleDb
Partial Class MasterPage
    Inherits System.Web.UI.MasterPage

    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("userid") Is Nothing Then
            Response.Redirect("~/login.aspx")
        End If
        lbuser.Text = Session("userid")
        If Page.IsPostBack = False Then
            AddHandler TreeView1.buttonClick, AddressOf TreeView1_buttonClick
            LoadMenu()
        End If
        AddHandler TreeView1.buttonClick, AddressOf TreeView1_buttonClick
        LoadMenu()
    End Sub

    Protected Sub Label2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.Click
        Session.RemoveAll()
        Response.Redirect("~/login.aspx")
    End Sub

    Protected Sub TreeView1_buttonClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Tree As New TreeView
        Tree = CType(sender, TreeView)
        Dim str As String
        str = Tree.SelectedNode.Target
        Response.Redirect("~/" & str)
    End Sub
    Private Sub LoadMenu()
        If Page.IsPostBack = False Then
            Dim str As String
            str = "Server=TYLO26SYS116;Database=test;User Id=sa; Password=1234;"
            str = "Provider=IBMDA400 ;Data Source=10.200.1.5;User ID=TY#00006;Password=L333"
            Dim Conn As New OleDbConnection(str)
            str = "select * from TYNSSALE_D.TYNSMN"
            'Dim Conn As New SqlConnection(str)
            'str = "select * from TYNSMN"

            Dim ds As New DataSet

            Dim cmd As New OleDbCommand(str, Conn)
            Dim da As New OleDbDataAdapter(cmd)
            '   Dim cmd As New SqlCommand(str, Conn)
            'Dim da As New SqlDataAdapter(cmd)
            da.Fill(ds, "tb1")
            Dim i As Integer
            i = 0
            For Each r As DataRow In ds.Tables(0).Rows
                If r(1).ToString() = "0" Then
                    ds.Tables(0).Rows(i)(1) = DBNull.Value
                End If
                i = i + 1
            Next
            TreeView1.DataSource = ds.Tables(0)
            TreeView1.KeyMember = "NSFMID"
            TreeView1.DisplayMember = "NSFNAM"
            TreeView1.ValueMember = "NSFMID"
            TreeView1.ParentMember = "NSFPID"
            TreeView1.NavigateMember = "NSURL"
        End If


    End Sub
End Class

