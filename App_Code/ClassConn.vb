Imports Microsoft.VisualBasic

Public Class Classconn
    Public Shared strConnSql As String = "Provider=SQLOLEDB;Password=db2005;User ID=dbconnect;Persist Security Info=True;Initial Catalog=YGSS112003;Data Source=10.200.1.36;Connect Timeout=10000;"
    Public Shared strConnSql400 As String = "Provider=SQLOLEDB;Password=db2005;User ID=dbconnect;Persist Security Info=True;Initial Catalog=AS400;Data Source=10.200.1.36;Connect Timeout=10000;"
    Public Shared strSQL22_32 As String = "Provider=SQLOLEDB;Password=sa;User ID=sa;Persist Security Info=True;Initial Catalog=TYL-FGVS;Data Source=10.200.22.32;Connect Timeout=10000;"
    Public Shared strCon400 As String = "Provider=IBMDA400;Data Source=10.200.1.5;User ID=TY#00006;Password=L333;"
    Public Shared LIBRARY_FILE As String = "KAMONRAT.WTOORD"
    Public Shared crtDate As Decimal = Date.Now.ToString("yyyyMMdd")
    Public Shared CrtTime As Decimal = Date.Now.ToString("HHmm")
    'Public Shared LIBRARY_TIOD As String = "KAMONRAT.TIOD"
    'Public Shared LIBRARY_TITD As String = "KAMONRAT.TITD"
    'Public Shared FILE As String = ""
    Public Shared Function OpenSqlConn() As OleDb.OleDbConnection
        Dim oConn As New OleDb.OleDbConnection
        If oConn.State = ConnectionState.Open Then
            oConn.Close()
        End If
        oConn.ConnectionString = strConnSql
        oConn.Open()
        Return oConn
    End Function

    Public Shared Function CloseConn(ByRef oConn As OleDb.OleDbConnection) As Boolean
        If oConn.State = ConnectionState.Open Then
            Try
                oConn.Close()
            Catch ex As Exception
                Return False
            End Try
        End If
        Return True
    End Function
End Class
