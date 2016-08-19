Imports Microsoft.VisualBasic

Public Class User
    Private mUserName As String
    Private mUserID As String
    Private mUserPWD As String


    Public Property Username() As String
        Get
            Return mUserName
        End Get
        Set(ByVal value As String)
            mUserName = value
        End Set
    End Property

    Public Sub getProfile(ByVal vUserName As String, ByVal vUserID As String)




    End Sub


End Class
