Imports Microsoft.VisualBasic
Imports System.Data
Imports system.Data.OleDb
Imports system.Data.SqlClient
Public Class test



    Private mPeYear As Decimal
    Private mPeriod As Integer
    Private mPeMonth As String
    Private mfactory As String
    Private mCarline As String
    Private mCustomer As String

    Private c3 As Integer

    Private mDtYear As DataTable
    Private mDtCustomer As DataTable
    Private mDtCarline As DataTable


    Public Property PeYear() As Decimal
        Get
            Return mPeYear
        End Get
        Set(ByVal value As Decimal)
            mPeYear = value
        End Set
    End Property

    Public Property period() As Integer
        Get
            Return mPeriod
        End Get
        Set(ByVal value As Integer)
            mPeriod = value
            If value = 2 Then
                mPeYear = mPeYear + 1
                mPeMonth = " And pmont in ( 1,2,3,4,5,6 ) "
            ElseIf value = 1 Then
                mPeMonth = " And pmont in ( 7,8,9,10,11,12 ) "
            End If
        End Set
    End Property

    Public Property Factory() As String
        Get
            Return mfactory
        End Get
        Set(ByVal value As String)
            mfactory = value
            
        End Set
    End Property

    Public Property Customer() As String
        Get
            Return mCustomer
        End Get
        Set(ByVal value As String)
            mCustomer = value
        End Set
    End Property

    Public Property Carline() As String
        Get
            Return mCarline
        End Get
        Set(ByVal value As String)
            mCarline = value
        End Set
    End Property

    Public ReadOnly Property dtyear() As DataTable
        Get
            Return mDtYear
        End Get
    End Property

    Public ReadOnly Property PeMonth() As String
        Get
            Return mPeMonth
        End Get
    End Property

    Public ReadOnly Property DtCustomer() As DataTable
        Get
            Return mDtCustomer
        End Get
    End Property

    Public ReadOnly Property DtCarline() As DataTable
        Get
            Return mDtCarline
        End Get
    End Property




    Public Sub getYear()
        Dim dt As New DataTable
        Dim strsql As String
        strsql = "select distinct pyear from sp_shortplan"
        Dim da As New OleDbDataAdapter(strsql, ClassConn.strconnsql)
        da.Fill(dt)
        mDtYear = dt
    End Sub


    Public Sub getCustomer()
        Dim strsql As String
        strsql = "select distinct CKEY , left(CNAME,26) + '...' as CNAME "
        strsql += "from SP_SHORTPLAN INNER JOIN v_YTAPCUS ON rtrim(SP_SHORTPLAN.PCUCD) = rtrim(v_YTAPCUS.CKEY) "
        Dim da As New OleDbDataAdapter(strsql, Classconn.strConnSql)
        Dim dt As New DataTable
        da.Fill(dt)
        da.Dispose()
        mDtCustomer = dt
    End Sub
    Public Sub getCustomer(ByVal vPyear As Integer, ByVal vPeriod As Integer)
        Me.PeYear = vPyear
        Me.period = vPeriod

        Dim year As Decimal = Me.PeYear
        Dim month As String = Me.PeMonth


        Dim strsql As String
        strsql = "select distinct CKEY , left(CNAME,26) + '...' as CNAME "
        strsql += "from SP_SHORTPLAN INNER JOIN v_YTAPCUS ON rtrim(SP_SHORTPLAN.PCUCD) = rtrim(v_YTAPCUS.CKEY) "
        strsql += "where pyear = " & year & " " & month


        Dim da As New OleDbDataAdapter(strsql, Classconn.strConnSql)
        Dim dt As New DataTable
        da.Fill(dt)
        da.Dispose()
        mDtCustomer = dt
    End Sub
    Public Sub getCustomer(ByVal vPyear As Integer, ByVal vPeriod As Integer, ByVal FAC As String)
        Me.PeYear = vPyear
        Me.period = vPeriod

        Dim year As Decimal = Me.PeYear
        Dim strMonth As String = Me.PeMonth


        Dim strsql As String
        strsql = "select distinct CKEY , left(CNAME,26) + '...' as CNAME "
        strsql += "from SP_SHORTPLAN INNER JOIN v_YTAPCUS ON rtrim(SP_SHORTPLAN.PCUCD) = rtrim(v_YTAPCUS.CKEY) "
        strsql += "where pyear = " & year & " " & strMonth
        strsql += " AND PFACD = '" & FAC & "'"


        Dim da As New OleDbDataAdapter(strsql, Classconn.strConnSql)
        Dim dt As New DataTable
        da.Fill(dt)
        da.Dispose()
        mDtCustomer = dt
    End Sub




    Public Sub getCarline(ByVal vPyear As Integer, ByVal vPeriod As Integer, ByVal FAC As String, ByVal vCustomer As String, ByVal vSpStatus As String, ByVal vProdGp As String)
        Dim dt As New DataTable

        Me.PeYear = vPyear
        Me.period = vPeriod

        Dim year As Decimal = Me.PeYear
        Dim strMonth As String = Me.PeMonth

        Dim strsql As String

        strsql = "select DISTINCT left(sp_shortplan.pcrln,4) as carlineCD,V_YTAPCALN.CNAME AS carlineNM "
        strsql += "from SP_SHORTPLAN left join V_YTAPCALN on V_YTAPCALN.CKEY = sp_shortplan.pcrln "
        strsql += "where rtrim(pcucd) = '" & vCustomer & "' "
        strsql += "AND PPDGP = '" & vProdGp & "' "  '01 = wh, 02 = mt
        strsql += "AND STATUS in " & vSpStatus & " "
        strsql += "And pyear = " & year & " " & strMonth

        strsql += "order by left(sp_shortplan.pcrln,4) ,V_YTAPCALN.CNAME "


        Dim da As New OleDbDataAdapter(strsql, Classconn.strConnSql)
        da.Fill(dt)

        Me.mDtCarline = dt
    End Sub










End Class
