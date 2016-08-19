Imports Microsoft.VisualBasic

Public Class GetFieldFromText
    Public Shared Sub GetValuetoArray(ByVal parmString As String, ByRef aDetail As Array)
        Dim LenString As Integer = Len(Trim(parmString))
        Dim CharStr As String = ""
        Dim ii, Pos1, Pos2 As Integer
        Dim NoChar As Integer = 0
        Dim idx As Integer = 0
        Dim NoChr34 As Integer = 0
        For ii = 1 To LenString
            CharStr = Mid(parmString, ii, 1)
            NoChar += 1
            Select Case CharStr
                Case Is <> Chr(34)
                    Pos2 += 1
                Case Chr(34)
                    If NoChr34 = 0 Then
                        Pos1 = ii + 1
                    End If
                    NoChr34 += 1
            End Select
            If (CharStr = ",") And (NoChr34 = 0) Then
                Pos2 = 0
                NoChar = 0
            End If
            If (CharStr = ",") And (NoChar = 3) Then
                aDetail(idx) = ""
                Call Initial_Value(idx, NoChar, NoChr34, Pos2)
            Else
                If (NoChr34 > 1) And (Pos2 > 0) Then
                    aDetail(idx) = Mid(parmString, Pos1, Pos2)
                    Call Initial_Value(idx, NoChar, NoChr34, Pos2)
                End If
            End If
        Next
    End Sub
    Public Shared Sub Initial_Value(ByRef idx As Integer, ByRef NoChar As Integer, _
                                    ByRef NoChr34 As Integer, ByRef Pos2 As Integer)
        idx += 1
        NoChr34 = 0
        Pos2 = 0
        NoChar = 0
    End Sub
End Class
