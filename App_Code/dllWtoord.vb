Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb

Public Class dllWtoord
    Private mWF01 As Decimal
    Private mWF02 As String
    Private mWF03 As String
    Private mWF04 As String
    Private mWF05 As String
    Private mWF06 As String
    Private mWF07 As String
    Private mWF08 As Decimal
    Private mWF09 As String
    Private mWF10 As String
    Private mWF11 As String
    Private mWF12 As String
    Private mWF13 As String
    Private mWF14 As String
    Private mWF15 As String
    Private mWF16 As String
    Private mWF17 As String
    Private mWF18 As String
    Private mWF19 As String
    Private mWF20 As String
    Private mWF21 As String
    Private mWF22 As String
    Private mWF23 As String
    Private mWF24 As String
    Private mWF25 As String
    Private mWF26 As String
    Private mWF27 As String
    Private mWF28 As String
    Private mWF29 As String
    Private mWF30 As String
    Private mWF31 As String
    Private mWF32 As String
    Private mWF33 As String
    Private mWF34 As String
    Private mWF35 As String
    Private mWF36 As String
    Private mWF37 As String
    Private mWF38 As Decimal
    Private mWF39 As String
    Private mWF40 As String
    Private mWF41 As String
    Private mWF42 As String
    Private mWF43 As Decimal
    Private mWF44 As Decimal
    Private mWF45 As Decimal
    Private mWF46 As String
    Private mWF47 As String
    Private mWF48 As String
    Private mWF49 As String
    Private mWF50 As String
    Private mWF51 As String
    Private mWF52 As String
    Private mWF53 As String
    Private mWF54 As String
    Private mWF55 As String
    Private mWF56 As String
    Private mWF57 As String
    Private mWF58 As String
    Private mWF59 As String

    Private mAccStatus As Boolean

    Public Property WF01() As Decimal
        Get
            Return mWF01
        End Get
        Set(ByVal value As Decimal)
            mWF01 = value
        End Set
    End Property
    Public Property WF02() As String
        Get
            Return mWF02
        End Get
        Set(ByVal value As String)
            mWF02 = value
        End Set
    End Property
    Public Property WF03() As String
        Get
            Return mWF03
        End Get
        Set(ByVal value As String)
            mWF03 = value
        End Set
    End Property
    Public Property WF04() As String
        Get
            Return mWF04
        End Get
        Set(ByVal value As String)
            mWF04 = value
        End Set
    End Property
    Public Property WF05() As String
        Get
            Return mWF05
        End Get
        Set(ByVal value As String)
            mWF05 = value
        End Set
    End Property
    Public Property WF06() As String
        Get
            Return mWF06
        End Get
        Set(ByVal value As String)
            mWF06 = value
        End Set
    End Property
    Public Property WF07() As String
        Get
            Return mWF07
        End Get
        Set(ByVal value As String)
            mWF07 = value
        End Set
    End Property
    Public Property WF08() As Decimal
        Get
            Return mWF08
        End Get
        Set(ByVal value As Decimal)
            mWF08 = value
        End Set
    End Property
    Public Property WF09() As String
        Get
            Return mWF09
        End Get
        Set(ByVal value As String)
            mWF09 = value
        End Set
    End Property
    Public Property WF10() As String
        Get
            Return mWF10
        End Get
        Set(ByVal value As String)
            mWF10 = value
        End Set
    End Property
    Public Property WF11() As String
        Get
            Return mWF11
        End Get
        Set(ByVal value As String)
            mWF11 = value
        End Set
    End Property
    Public Property WF12() As String
        Get
            Return mWF12
        End Get
        Set(ByVal value As String)
            mWF12 = value
        End Set
    End Property
    Public Property WF13() As String
        Get
            Return mWF13
        End Get
        Set(ByVal value As String)
            mWF13 = value
        End Set
    End Property
    Public Property WF14() As String
        Get
            Return mWF14
        End Get
        Set(ByVal value As String)
            mWF14 = value
        End Set
    End Property
    Public Property WF15() As String
        Get
            Return mWF15
        End Get
        Set(ByVal value As String)
            mWF15 = value
        End Set
    End Property
    Public Property WF16() As String
        Get
            Return mWF16
        End Get
        Set(ByVal value As String)
            mWF16 = value
        End Set
    End Property
    Public Property WF17() As String
        Get
            Return mWF17
        End Get
        Set(ByVal value As String)
            mWF17 = value
        End Set
    End Property
    Public Property WF18() As String
        Get
            Return mWF18
        End Get
        Set(ByVal value As String)
            mWF18 = value
        End Set
    End Property
    Public Property WF19() As String
        Get
            Return mWF19
        End Get
        Set(ByVal value As String)
            mWF19 = value
        End Set
    End Property
    Public Property WF20() As String
        Get
            Return mWF20
        End Get
        Set(ByVal value As String)
            mWF20 = value
        End Set
    End Property
    Public Property WF21() As String
        Get
            Return mWF21
        End Get
        Set(ByVal value As String)
            mWF21 = value
        End Set
    End Property
    Public Property WF22() As String
        Get
            Return mWF22
        End Get
        Set(ByVal value As String)
            mWF22 = value
        End Set
    End Property
    Public Property WF23() As String
        Get
            Return mWF23
        End Get
        Set(ByVal value As String)
            mWF23 = value
        End Set
    End Property
    Public Property WF24() As String
        Get
            Return mWF24
        End Get
        Set(ByVal value As String)
            mWF24 = value
        End Set
    End Property
    Public Property WF25() As String
        Get
            Return mWF25
        End Get
        Set(ByVal value As String)
            mWF25 = value
        End Set
    End Property
    Public Property WF26() As String
        Get
            Return mWF26
        End Get
        Set(ByVal value As String)
            mWF26 = value
        End Set
    End Property
    Public Property WF27() As String
        Get
            Return mWF27
        End Get
        Set(ByVal value As String)
            mWF27 = value
        End Set
    End Property
    Public Property WF28() As String
        Get
            Return mWF28
        End Get
        Set(ByVal value As String)
            mWF28 = value
        End Set
    End Property
    Public Property WF29() As String
        Get
            Return mWF29
        End Get
        Set(ByVal value As String)
            mWF29 = value
        End Set
    End Property
    Public Property WF30() As String
        Get
            Return mWF30
        End Get
        Set(ByVal value As String)
            mWF30 = value
        End Set
    End Property
    Public Property WF31() As String
        Get
            Return mWF31
        End Get
        Set(ByVal value As String)
            mWF31 = value
        End Set
    End Property
    Public Property WF32() As String
        Get
            Return mWF32
        End Get
        Set(ByVal value As String)
            mWF32 = value
        End Set
    End Property
    Public Property WF33() As String
        Get
            Return mWF33
        End Get
        Set(ByVal value As String)
            mWF33 = value
        End Set
    End Property
    Public Property WF34() As String
        Get
            Return mWF34
        End Get
        Set(ByVal value As String)
            mWF34 = value
        End Set
    End Property
    Public Property WF35() As String
        Get
            Return mWF35
        End Get
        Set(ByVal value As String)
            mWF35 = value
        End Set
    End Property
    Public Property WF36() As String
        Get
            Return mWF36
        End Get
        Set(ByVal value As String)
            mWF36 = value
        End Set
    End Property
    Public Property WF37() As String
        Get
            Return mWF37
        End Get
        Set(ByVal value As String)
            mWF37 = value
        End Set
    End Property
    Public Property WF38() As Decimal
        Get
            Return mWF38
        End Get
        Set(ByVal value As Decimal)
            mWF38 = value
        End Set
    End Property
    Public Property WF39() As String
        Get
            Return mWF39
        End Get
        Set(ByVal value As String)
            mWF39 = value
        End Set
    End Property
    Public Property WF40() As String
        Get
            Return mWF40
        End Get
        Set(ByVal value As String)
            value = Left(value, value.LastIndexOf("-"))
            'value = (Replace(Trim(value), "-", ""))
            value = Replace(Trim(value), " ", "")
            mWF40 = value
        End Set
    End Property
    Public Property WF41() As String
        Get
            Return mWF41
        End Get
        Set(ByVal value As String)
            mWF41 = value
        End Set
    End Property
    Public Property WF42() As String
        Get
            Return mWF42
        End Get
        Set(ByVal value As String)
            mWF42 = value
        End Set
    End Property
    Public Property WF43() As Decimal
        Get
            Return mWF43
        End Get
        Set(ByVal value As Decimal)
            mWF43 = value
        End Set
    End Property
    Public Property WF44() As Decimal
        Get
            Return mWF44
        End Get
        Set(ByVal value As Decimal)
            mWF44 = value
        End Set
    End Property
    Public Property WF45() As Decimal
        Get
            Return mWF45
        End Get
        Set(ByVal value As Decimal)
            mWF45 = value
        End Set
    End Property
    Public Property WF46() As String
        Get
            Return mWF46
        End Get
        Set(ByVal value As String)
            mWF46 = value
        End Set
    End Property
    Public Property WF47() As String
        Get
            Return mWF47
        End Get
        Set(ByVal value As String)
            mWF47 = value
        End Set
    End Property
    Public Property WF48() As String
        Get
            Return mWF48
        End Get
        Set(ByVal value As String)
            mWF48 = value
        End Set
    End Property
    Public Property WF49() As String
        Get
            Return mWF49
        End Get
        Set(ByVal value As String)
            mWF49 = value
        End Set
    End Property
    Public Property WF50() As String
        Get
            Return mWF50
        End Get
        Set(ByVal value As String)
            mWF50 = value
        End Set
    End Property
    Public Property WF51() As String
        Get
            Return mWF51
        End Get
        Set(ByVal value As String)
            mWF51 = value
        End Set
    End Property
    Public Property WF52() As String
        Get
            Return mWF52
        End Get
        Set(ByVal value As String)
            mWF52 = value
        End Set
    End Property
    Public Property WF53() As String
        Get
            Return mWF53
        End Get
        Set(ByVal value As String)
            mWF53 = value
        End Set
    End Property
    Public Property WF54() As String
        Get
            Return mWF54
        End Get
        Set(ByVal value As String)
            mWF54 = value
        End Set
    End Property
    Public Property WF55() As String
        Get
            Return mWF55
        End Get
        Set(ByVal value As String)
            mWF55 = value
        End Set
    End Property
    Public Property WF56() As String
        Get
            Return mWF56
        End Get
        Set(ByVal value As String)
            mWF56 = value
        End Set
    End Property
    Public Property WF57() As String
        Get
            Return mWF57
        End Get
        Set(ByVal value As String)
            mWF57 = value
        End Set
    End Property
    Public Property WF58() As String
        Get
            Return mWF58
        End Get
        Set(ByVal value As String)
            mWF58 = value
        End Set
    End Property
    Public Property WF59() As String
        Get
            Return mWF59
        End Get
        Set(ByVal value As String)
            mWF59 = value
        End Set
    End Property

    Public ReadOnly Property AccStatus() As Boolean
        Get
            Return mAccStatus
        End Get
    End Property


    Public Sub insert()
        Dim strsql As String
        strsql = "INSERT INTO " & Classconn.LIBRARY_FILE & " "
        strsql += "(WF01,WF02,WF03,WF04,WF05,WF06,WF07,WF08,WF09,WF10,"
        strsql += "WF11,WF12,WF13,WF14,WF15,WF16,WF17,WF18,WF19,WF20,"
        strsql += "WF21,WF22,WF23,WF24,WF25,WF26,WF27,WF28,WF29,WF30,"
        strsql += "WF31,WF32,WF33,WF34,WF35,WF36,WF37,WF38,WF39,WF40,"
        strsql += "WF41,WF42,WF43,WF44,WF45,WF46,WF47,WF48,WF49,WF50,"
        strsql += "WF51,WF52,WF53,WF54,WF55,WF56,WF57,WF58,WF59)"
        strsql += "VALUES (?,?,?,?,?,?,?,?,?,?,"
        strsql += "?,?,?,?,?,?,?,?,?,?,"
        strsql += "?,?,?,?,?,?,?,?,?,?,"
        strsql += "?,?,?,?,?,?,?,?,?,?,"
        strsql += "?,?,?,?,?,?,?,?,?,?,"
        strsql += "?,?,?,?,?,?,?,?,?"

        Dim conn As New OleDbConnection(Classconn.strCon400)
        Dim insertDA As New OleDbDataAdapter
        conn.Open()
        Dim insertCMD As New OleDbCommand(strsql, conn)

        insertDA.InsertCommand = insertCMD

        insertCMD.Parameters.Add("WF01", OleDbType.Decimal, 6).Value = Me.mWF01
        insertCMD.Parameters.Add("WF02", OleDbType.VarChar, 4).Value = Me.mWF02
        insertCMD.Parameters.Add("WF03", OleDbType.VarChar, 1).Value = Me.mWF03
        insertCMD.Parameters.Add("WF04", OleDbType.VarChar, 40).Value = Me.mWF04
        insertCMD.Parameters.Add("WF05", OleDbType.VarChar, 1).Value = Me.mWF05
        insertCMD.Parameters.Add("WF06", OleDbType.VarChar, 20).Value = Me.mWF06
        insertCMD.Parameters.Add("WF07", OleDbType.VarChar, 2).Value = Me.mWF07
        insertCMD.Parameters.Add("WF08", OleDbType.Decimal, 1).Value = Me.mWF08
        insertCMD.Parameters.Add("WF09", OleDbType.VarChar, 15).Value = Me.mWF09
        insertCMD.Parameters.Add("WF10", OleDbType.VarChar, 12).Value = Me.mWF10
        insertCMD.Parameters.Add("WF11", OleDbType.VarChar, 10).Value = Me.mWF11
        insertCMD.Parameters.Add("WF12", OleDbType.VarChar, 5).Value = Me.mWF12
        insertCMD.Parameters.Add("WF13", OleDbType.VarChar, 10).Value = Me.mWF13
        insertCMD.Parameters.Add("WF14", OleDbType.VarChar, 5).Value = Me.mWF14
        insertCMD.Parameters.Add("WF15", OleDbType.VarChar, 10).Value = Me.mWF15
        insertCMD.Parameters.Add("WF16", OleDbType.VarChar, 2).Value = Me.mWF16
        insertCMD.Parameters.Add("WF17", OleDbType.VarChar, 10).Value = Me.mWF17
        insertCMD.Parameters.Add("WF18", OleDbType.VarChar, 2).Value = Me.mWF18
        insertCMD.Parameters.Add("WF19", OleDbType.VarChar, 13).Value = Me.mWF19
        insertCMD.Parameters.Add("WF20", OleDbType.VarChar, 10).Value = Me.mWF20
        insertCMD.Parameters.Add("WF21", OleDbType.VarChar, 10).Value = Me.mWF21
        insertCMD.Parameters.Add("WF22", OleDbType.VarChar, 5).Value = Me.mWF22
        insertCMD.Parameters.Add("WF23", OleDbType.VarChar, 10).Value = Me.mWF23
        insertCMD.Parameters.Add("WF24", OleDbType.VarChar, 5).Value = Me.mWF24
        insertCMD.Parameters.Add("WF25", OleDbType.VarChar, 13).Value = Me.mWF25
        insertCMD.Parameters.Add("WF26", OleDbType.VarChar, 10).Value = Me.mWF26
        insertCMD.Parameters.Add("WF27", OleDbType.VarChar, 10).Value = Me.mWF27
        insertCMD.Parameters.Add("WF28", OleDbType.VarChar, 5).Value = Me.mWF28
        insertCMD.Parameters.Add("WF29", OleDbType.VarChar, 10).Value = Me.mWF29
        insertCMD.Parameters.Add("WF30", OleDbType.VarChar, 5).Value = Me.mWF30
        insertCMD.Parameters.Add("WF31", OleDbType.VarChar, 13).Value = Me.mWF31
        insertCMD.Parameters.Add("WF32", OleDbType.VarChar, 10).Value = Me.mWF32
        insertCMD.Parameters.Add("WF33", OleDbType.VarChar, 10).Value = Me.mWF33
        insertCMD.Parameters.Add("WF34", OleDbType.VarChar, 5).Value = Me.mWF34
        insertCMD.Parameters.Add("WF35", OleDbType.VarChar, 10).Value = Me.mWF35
        insertCMD.Parameters.Add("WF36", OleDbType.VarChar, 5).Value = Me.mWF36
        insertCMD.Parameters.Add("WF37", OleDbType.VarChar, 1).Value = Me.mWF37
        insertCMD.Parameters.Add("WF38", OleDbType.Decimal, 3).Value = Me.mWF38
        insertCMD.Parameters.Add("WF39", OleDbType.VarChar, 15).Value = Me.mWF39
        insertCMD.Parameters.Add("WF40", OleDbType.VarChar, 40).Value = Me.mWF40
        insertCMD.Parameters.Add("WF41", OleDbType.VarChar, 6).Value = Me.mWF41
        insertCMD.Parameters.Add("WF42", OleDbType.VarChar, 10).Value = Me.mWF42
        insertCMD.Parameters.Add("WF43", OleDbType.Decimal, 6).Value = Me.mWF43
        insertCMD.Parameters.Add("WF44", OleDbType.Decimal, 7).Value = Me.mWF44
        insertCMD.Parameters.Add("WF45", OleDbType.Decimal, 7).Value = Me.mWF45
        insertCMD.Parameters.Add("WF46", OleDbType.VarChar, 20).Value = Me.mWF46
        insertCMD.Parameters.Add("WF47", OleDbType.VarChar, 5).Value = Me.mWF47
        insertCMD.Parameters.Add("WF48", OleDbType.VarChar, 10).Value = Me.mWF48
        insertCMD.Parameters.Add("WF49", OleDbType.VarChar, 5).Value = Me.mWF49
        insertCMD.Parameters.Add("WF50", OleDbType.VarChar, 10).Value = Me.mWF50
        insertCMD.Parameters.Add("WF51", OleDbType.VarChar, 5).Value = Me.mWF51
        insertCMD.Parameters.Add("WF52", OleDbType.VarChar, 10).Value = Me.mWF52
        insertCMD.Parameters.Add("WF53", OleDbType.VarChar, 5).Value = Me.mWF53
        insertCMD.Parameters.Add("WF54", OleDbType.VarChar, 60).Value = Me.mWF54
        insertCMD.Parameters.Add("WF55", OleDbType.VarChar, 10).Value = Me.mWF55
        insertCMD.Parameters.Add("WF56", OleDbType.VarChar, 5).Value = Me.mWF56
        insertCMD.Parameters.Add("WF57", OleDbType.VarChar, 4).Value = Me.mWF57
        insertCMD.Parameters.Add("WF58", OleDbType.VarChar, 1).Value = Me.mWF58
        insertCMD.Parameters.Add("WF59", OleDbType.VarChar, 3).Value = Me.mWF59
        Try
            insertCMD.ExecuteNonQuery()
            Me.mAccStatus = True
        Catch ex As System.Exception
            Me.mAccStatus = False
            insertDA.InsertCommand.Connection.Close()
            conn.Close()
        End Try

        insertDA.InsertCommand.Connection.Close()
        conn.Close()
    End Sub
    Public Sub delete()
        Dim strsql As String
        strsql = "delete from " & Classconn.LIBRARY_FILE & " "
        'strsql += "Where YWBTCD = ? and YWPLNR = ? "

        Dim conn As New OleDbConnection(Classconn.strCon400)
        Dim delDA As New OleDbDataAdapter
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.Open()

        Dim delCMD As New OleDbCommand(strsql, conn)
        delDA.InsertCommand = delCMD
        Try
            delCMD.ExecuteNonQuery()
            Me.mAccStatus = True
        Catch ex As System.Exception
            Me.mAccStatus = False
            conn.Close()
        End Try


        delCMD.Dispose()
        conn.Close()


    End Sub

    Public Sub selectdata()

        Dim dt As New DataTable
        Dim strsql As String
        strsql = "select * from " & Classconn.LIBRARY_FILE
        Dim da As New OleDbDataAdapter(strsql, Classconn.strCon400)
        da.Fill(dt)

        If dt.Rows.Count < 0 Then
            Me.mAccStatus = False
        Else
            Me.mAccStatus = True
        End If

    End Sub

End Class
