Imports Microsoft.VisualBasic

Public Class clConvert
    Dim mConvDate As String
    Public Property convDate() As String ' ddmmyyyy to yymmdd
        Get
            Return mConvDate
        End Get
        Set(ByVal value As String)
            If value <> "" Then
                mConvDate = Right(Trim(value), 2) & Mid(Trim(value), 3, 2) & Left(Trim(value), 2)
            Else
                mConvDate = 0
            End If


        End Set
    End Property


End Class
