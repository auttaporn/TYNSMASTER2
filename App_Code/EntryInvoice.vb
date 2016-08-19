Imports Microsoft.VisualBasic

Public Class EntryInvoice

    Private _EEORID As Decimal
    Public Property EEORID() As Decimal
        Get
            Return _EEORID
        End Get
        Set(ByVal value As Decimal)
            _EEORID = value
        End Set
    End Property

    Private _EEPROD As String
    Public Property EEPROD() As String
        Get
            Return _EEPROD
        End Get
        Set(ByVal value As String)
            _EEPROD = value
        End Set
    End Property

    Private _EEINVN As String
    Public Property EEINVN() As String
        Get
            Return _EEINVN
        End Get
        Set(ByVal value As String)
            _EEINVN = value
        End Set
    End Property

    Private _EEEXIV As String
    Public Property EEEXIV() As String
        Get
            Return _EEEXIV
        End Get
        Set(ByVal value As String)
            _EEEXIV = value
        End Set
    End Property

    Private _EECSNO As String
    Public Property EECSNO() As String
        Get
            Return _EECSNO
        End Get
        Set(ByVal value As String)
            _EECSNO = value
        End Set
    End Property

    Private _EEDLDT As Decimal
    Public Property EEDLDT() As Decimal
        Get
            Return _EEDLDT
        End Get
        Set(ByVal value As Decimal)
            _EEDLDT = value
        End Set
    End Property

    Private _EEQTY As Decimal
    Public Property EEQTY() As Decimal
        Get
            Return _EEQTY
        End Get
        Set(ByVal value As Decimal)
            _EEQTY = value
        End Set
    End Property

    Private _EECPDT As Decimal
    Public Property EECPDT() As Decimal
        Get
            Return _EECPDT
        End Get
        Set(ByVal value As Decimal)
            _EECPDT = value
        End Set
    End Property


End Class
