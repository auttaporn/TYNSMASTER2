Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb

Public Class ClassMaster
    Private Conn As New ClassConn
#Region "PROPERTY"
    Dim _YYEAR As Integer
    Dim _YMONTH As Integer
    Dim _YOH As Integer
    Dim _YROYAL As Integer
    Dim _YPROFIT As Integer
    Dim _YMONTHCU As Decimal
    Dim _YEXRATE As Decimal
    Dim _YICCUCD As String = "999999"  '**** pui update 18/11/2011 add new customer (JOHOKU) ****

    Public Property YICCUCD()
        Get
            YICCUCD = _YICCUCD
        End Get
        Set(ByVal value)
            _YICCUCD = value
        End Set
    End Property

    Public Property YYEAR()
        Get
            YYEAR = _YYEAR
        End Get
        Set(ByVal value)
            _YYEAR = value
        End Set
    End Property
    Public Property YMONTH()
        Get
            YMONTH = _YMONTH
        End Get
        Set(ByVal value)
            _YMONTH = value
        End Set
    End Property
    Public Property YOH()
        Get
            YOH = _YOH
        End Get
        Set(ByVal value)
            _YOH = value
        End Set
    End Property
    Public Property YROYAL()
        Get
            YROYAL = _YROYAL
        End Get
        Set(ByVal value)
            _YROYAL = value
        End Set
    End Property
    Public Property YPROFIT()
        Get
            YPROFIT = _YPROFIT
        End Get
        Set(ByVal value)
            _YPROFIT = value
        End Set
    End Property
    Public Property YMONTHCU()
        Get
            YMONTHCU = _YMONTHCU
        End Get
        Set(ByVal value)
            _YMONTHCU = value
        End Set
    End Property
    Public Property YEXRATE()
        Get
            YEXRATE = _YEXRATE
        End Get
        Set(ByVal value)
            _YEXRATE = value
        End Set
    End Property
#End Region
    Public Shared Function CheckDate(ByRef strDate As String, ByRef oMSG As String, ByVal iFormat As String, ByVal iCase As String)
        Dim iDate As String = ""
        Dim iMonth As String = ""
        Dim iYear As String = ""
        If strDate.Length = 8 Then
            Select Case iFormat
                Case "ddmmyyyy"
                    iDate = Left(strDate, 2)
                    iMonth = Mid(strDate, 3, 2)
                    iYear = Right(strDate, 4)
                Case "yyyymmdd"
                    iDate = Right(strDate, 2)
                    iMonth = Mid(strDate, 5, 2)
                    iYear = Left(strDate, 4)
            End Select
            If iDate > 31 Or iDate < 1 Then
                oMSG = "DATE INVALID, DATE BETWEEN 1-31"
            End If
            If iMonth > 12 Or iMonth < 1 Then
                oMSG = "MONTH INVALID, MONTH BETWEEN 1-12"
            End If
            Select Case iCase
                Case "YTHA1"
                    'Format 01122009 --> 01122552
                    If iYear < 2500 Then
                        iYear += 543
                    End If
                    strDate = iDate & iMonth & iYear
                Case "YENG1"
                    'Format 01122552 --> 01122009
                    If iYear > 2500 Then
                        iYear -= 543
                    End If
                    strDate = iDate & iMonth & iYear
                Case "YTHA2"
                    'Format 20091201 --> 25521201
                    If iYear < 2500 Then
                        iYear += 543
                    End If
                    strDate = iYear & iMonth & iDate
                Case "YENG2"
                    'Format 25521201 --> 20091201
                    If iYear > 2500 Then
                        iYear -= 543
                    End If
                    strDate = iYear & iMonth & iDate
            End Select
        End If
        Return strDate
    End Function

    Public Shared Function GetMaster(ByVal Condition As String) As DataTable

        If Left(Condition.ToUpper.Trim, 5) = "WHERE" Then '**** pui update 18/11/2011 add new customer (JOHOKU) ****
            Condition += " and yiccucd='999999'"
        Else
            Condition += "where  yiccucd='999999'"
        End If
 
        Dim strMaster As String = ""
        ' strMaster = "SELECT DISTINCT * FROM YIC_SPMASTER " & Condition '**** pui update 18/11/2011 add new customer (JOHOKU) ****
        strMaster = "SELECT DISTINCT * FROM YIC_MARKUP " & Condition
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function

    Public Shared Function GetMRKMaster(ByVal Condition As String) As DataTable '**** pui update 18/11/2011 add new customer (JOHOKU) ****

        Dim strMaster As String = ""
        ' strMaster = "SELECT DISTINCT * FROM YIC_SPMASTER " & Condition '**** pui update 18/11/2011 add new customer (JOHOKU) ****
        strMaster = "SELECT DISTINCT * FROM YIC_MARKUP " & Condition
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function
    Public Shared Function GetMonth(ByVal Condition As String) As DataTable
        ' Public Shared Function GetMonth(ByVal Condition As String) As DataTable
        Dim strMaster As String = ""
        'strMaster = "SELECT DISTINCT YIMONTH FROM YIC_SPMASTER " & Condition  '**** pui update 18/11/2011 add new customer (JOHOKU) ****
        If Left(Condition.ToUpper.Trim, 5) = "WHERE" Then
            Condition += " and yiccucd='999999'"
        Else
            Condition += "where  yiccucd='999999'"
        End If
        strMaster = "SELECT DISTINCT YIMONTH FROM YIC_MARKUP " & Condition
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function

    Public Shared Function GetMRKMonth(ByVal Condition As String) As DataTable
        ' Public Shared Function GetMonth(ByVal Condition As String) As DataTable
        Dim strMaster As String = ""
        '  strMaster = "SELECT DISTINCT YIMONTH FROM YIC_SPMASTER " & Condition '**** pui update 18/11/2011 add new customer (JOHOKU) ****
        strMaster = "SELECT DISTINCT YIMONTH FROM YIC_MARKUP " & Condition
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function
    Public Shared Function GetSTRDATE(ByVal Condition As String) As DataTable
        Dim strMaster As String = ""
        strMaster = "SELECT DISTINCT YICSTDT FROM YIC_PRICEADJ " & Condition
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function
    Public Shared Function GetYear(ByVal Condition As String) As DataTable
        ' Public Shared Function GetYear(ByVal Condition As String) As DataTable 
        Dim strMaster As String = ""
        '  strMaster = "SELECT DISTINCT YIYEAR FROM YIC_SPMASTER " & Condition '**** pui update 18/11/2011 add new customer (JOHOKU) ****
        If Left(Condition.ToUpper.Trim, 5) = "WHERE" Then
            Condition += " and yiccucd='999999'"
        Else
            Condition += "where  yiccucd='999999'"
        End If
        strMaster = "SELECT DISTINCT YIYEAR FROM YIC_MARKUP " & Condition
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function
    Public Shared Function GetMRKYear(ByVal Condition As String) As DataTable
        ' Public Shared Function GetYear(ByVal Condition As String) As DataTable 
        Dim strMaster As String = ""
        strMaster = "SELECT DISTINCT YIYEAR FROM YIC_MARKUP " & Condition
        'strMaster = "SELECT DISTINCT YIYEAR FROM YIC_SPMASTER " & Condition '**** pui update 18/11/2011 add new customer (JOHOKU) ****
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function

    Public Shared Function GetSALEDATE(ByVal Condition As String) As DataTable
        Dim strMaster As String = ""
        strMaster = "SELECT DISTINCT YICSTDTS FROM YIC_PRICEADJ " & Condition
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function

    Public Shared Function getpermis(ByVal menutype As String, ByVal userid As String)
        Dim strsql As String
        strsql = " select *  from websalepermis where sysname='YICAPP' and menuID='" & menutype & "' and uid='" & userid & "'" '   YICREPORT / YICMAIN
        Dim DA As New OleDbDataAdapter(strsql, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT.Rows.Count
    End Function
    Public Shared Function GetCUSTOMER() As DataTable
        Dim strMaster As String = ""
        strMaster = " select * from openquery(as400,'select DISTINCT TTCUST,TNAME from ygss.tcust')"
        Dim DA As New OleDbDataAdapter(strMaster, ClassConn.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function
    Public Function UpdateMRKMaster() As String
        ' Public Function UpdateMaster() As String '**** pui update 18/11/2011 add new customer (JOHOKU) ****
        Dim oCmd As New OleDb.OleDbCommand
        '  oCmd.CommandText = "UPDATE YIC_SPMASTER SET YOH=?, YROYAL=?, YPROFIT=?, YMONTHCU=? WHERE YYEAR=? AND YMONTH=?"
        '  oCmd.Connection = Conn.OpenSqlConn
        'oCmd.Parameters.AddWithValue("@_YYEAR", _YYEAR)
        'oCmd.Parameters.AddWithValue("@_YMONTH", _YMONTH)
        oCmd.Parameters.AddWithValue("@_YOH", _YOH)
        oCmd.Parameters.AddWithValue("@_YROYAL", _YROYAL)
        oCmd.Parameters.AddWithValue("@_YPROFIT", _YPROFIT)
        oCmd.Parameters.AddWithValue("@_YMONTHCU", _YMONTHCU)
        oCmd.Parameters.AddWithValue("@_YEXRATE", _YEXRATE)
        oCmd.Parameters.Add("@YYEAR", SqlDbType.Int, 4).Value = _YYEAR
        oCmd.Parameters.Add("@YMONTH", SqlDbType.Int, 2).Value = _YMONTH

        oCmd.Parameters.Add("@YICCUCD", SqlDbType.Char, 6).Value = _YICCUCD

        ''oCmd.CommandText = "UPDATE YIC_SPMASTER SET YOH=?, YROYAL=?, YPROFIT=?, YMONTHCU=?, YEXRATE=? " & _ '**** pui update 18/11/2011 add new customer (JOHOKU) ****
        oCmd.CommandText = "UPDATE YIC_MARKUP SET YOH=?, YROYAL=?, YPROFIT=?, YMONTHCU=?, YEXRATE=? " & _
                           "WHERE YYEAR=? AND YMONTH=? AND YICPERD=1 AND YICCUCD=?"
        oCmd.Connection = ClassConn.OpenSqlConn
        Try
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Return "Error : " & ex.Source & ":: " & ex.Message
        End Try
        Return "Update Complete..."
    End Function



    'Public Shared Function GetYear(ByVal Condition As String) As DataTable
    '    Dim strMaster As String = ""
    '    'If oConn.State = ConnectionState.Closed Then
    '    '    oConn.Open()
    '    'End If

    '    strMaster = "SELECT DISTINCT YIYEAR FROM YIC_SPMASTER " & Condition
    '    Dim DA As New OleDbDataAdapter(strMaster, Conn.OpenSqlConn)
    '    Dim DT As New DataTable
    '    DA.Fill(DT)
    '    Return DT
    'End Function
    'Public Shared Function GetMonth(ByVal Condition As String) As DataTable
    '    Dim strMaster As String = ""
    '    'If oConn.State = ConnectionState.Closed Then
    '    '    oConn.Open()
    '    'End If

    '    strMaster = "SELECT DISTINCT YIMONTH FROM YIC_SPMASTER " & Condition
    '    Dim DA As New OleDbDataAdapter(strMaster, Conn.OpenSqlConn)
    '    Dim DT As New DataTable
    '    DA.Fill(DT)
    '    Return DT
    'End Function
    'Public Shared Function GetMaster(ByVal Condition As String) As DataTable
    '    Dim strMaster As String = ""
    '    'If oConn.State = ConnectionState.Closed Then
    '    '    oConn.Open()
    '    'End If

    '    strMaster = "SELECT DISTINCT * FROM YIC_SPMASTER " & Condition
    '    Dim DA As New OleDbDataAdapter(strMaster, Conn.strConnSql)
    '    Dim DT As New DataTable
    '    DA.Fill(DT)
    '    Return DT
    'End Function
End Class
