Imports Microsoft.VisualBasic

Public Class functionCHAR
    Public Shared Function checkasscii(ByVal strRecord As String)

        Dim cASC As Integer = Asc(Right(strRecord, 1).ToString)
        If cASC = 13 Then
            strRecord = Trim(Replace(strRecord, (Right(strRecord, 1).ToString), ""))
        Else
            strRecord = strRecord.Trim
        End If
        Return strRecord
    End Function
End Class
