Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb
Public Class Classcon
    Public Shared strConnSql As String = "Provider=SQLOLEDB;Password=db2005;User ID=dbconnect;Persist Security Info=True;Initial Catalog=YGSS112003;Data Source=10.200.1.36;Connect Timeout=10000;"
    'Public Shared strConnSql As String = "Provider=SQLOLEDB;Password=db2005;User ID=dbconnect;Persist Security Info=True;Initial Catalog=YGSS;Data Source=10.200.1.36;Connect Timeout=10000;"

    Public Shared strConnSql400 As String = "Provider=SQLOLEDB;Password=db2005;User ID=dbconnect;Persist Security Info=True;Initial Catalog=AS400;Data Source=10.200.1.36;Connect Timeout=10000;"
    Public Shared strCon400 As String = " Provider=IBMDA400;Data Source=10.200.1.5;User ID=TY#00006;Password=L333;"
    'Public Shared strCon400 As String = "Provider=IBMDA400;Data Source=10.200.1.5;User ID=TY015282;Password=J18062529;"
    Public Shared LIBTBD As String =  "OCSTYLBD"
    Public Shared LIBTBE As String =  "OCSTYLBE"
    Public Shared LIBTC As String =  "OCSTYL1"
    Public Shared LIBTP As String = "OCSTYL2"
    Public Shared LIBTEST As String = "#PUI"

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

