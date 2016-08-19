Imports Microsoft.VisualBasic

Public Class ConvertNumbertoText
    Public Shared Function SpellNumber(ByVal MyNum As Double) As String
        Dim Currency As String = ""
        Dim MyNumber As String = ""
        Dim Satang As String = ""
        Dim Temp As String = ""
        Dim DecimalPlace, Count As Decimal
        Dim Place(9) As String

        Place(2) = " Thousand "
        Place(3) = " Million "
        Place(4) = " Billion "
        Place(5) = " Trillion "
        ' String representation of amount.
        MyNumber = Trim(Str(MyNum))
        ' Position of decimal place 0 if none.
        DecimalPlace = InStr(MyNumber, ".")
        ' Convert SATANG and set MyNumber to dollar amount.
        If DecimalPlace > 0 Then

            Satang = GetTens(Left((Mid(MyNumber, DecimalPlace + 1) & "00"), 2))
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If
        Count = 1
        Do While MyNumber <> ""
            Temp = GetHundreds(Right(MyNumber, 3))
            If Temp <> "" Then
                Currency = Temp & Place(Count) & Currency
            End If
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
                MyNumber = ""
            End If
            Count = Count + 1
        Loop
        Select Case Currency
            Case ""
                Currency = "No BAHT"
            Case "One"
                Currency = "One BAHT"
            Case Else
                Currency = Currency & " BAHT"
        End Select
        Select Case Satang
            Case ""
                Satang = " and No SATANG"
            Case "One"
                Satang = " and One SATANG"
            Case Else
                Satang = " and " & Satang & " SATANG"
        End Select
        Return (Currency & Satang)
    End Function
    '*******************************************

    ' Converts a number from 100-999 into text *

    '*******************************************
    Public Shared Function GetHundreds(ByVal MyNumber As String) As String
        Dim Result As String = ""
        If Val(MyNumber) > 0 Then
            MyNumber = Right("000" & MyNumber, 3)
            ' Convert the hundreds place.
            If Mid(MyNumber, 1, 1) <> "0" Then
                Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
            End If
            ' Convert the tens and ones place.
            If Mid(MyNumber, 2, 1) <> "0" Then
                Result = Result & GetTens(Mid(MyNumber, 2))
            Else
                Result = Result & GetDigit(Mid(MyNumber, 3))
            End If
        End If
        Return Result
    End Function

    '*********************************************

    ' Converts a number from 10 to 99 into text. *
    Public Shared Function GetTens(ByVal TensText As String) As String
        Dim Result As String = ""
        Result = ""           ' Null out the temporary function value.
        If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
            Select Case Val(TensText)
                Case 10 : Result = "Ten"
                Case 11 : Result = "Eleven"
                Case 12 : Result = "Twelve"
                Case 13 : Result = "Thirteen"
                Case 14 : Result = "Fourteen"
                Case 15 : Result = "Fifteen"
                Case 16 : Result = "Sixteen"
                Case 17 : Result = "Seventeen"
                Case 18 : Result = "Eighteen"
                Case 19 : Result = "Nineteen"
                Case Else
            End Select
        Else                                 ' If value between 20-99...
            Select Case Val(Left(TensText, 1))
                Case 2 : Result = "Twenty "
                Case 3 : Result = "Thirty "
                Case 4 : Result = "Forty "
                Case 5 : Result = "Fifty "
                Case 6 : Result = "Sixty "
                Case 7 : Result = "Seventy "
                Case 8 : Result = "Eighty "
                Case 9 : Result = "Ninety "
            End Select
            Result = Result & GetDigit((Right(TensText, 1)))  ' Retrieve ones place.
        End If
        Return (Result)
    End Function

    '*******************************************

    ' Converts a number from 1 to 9 into text. *

    '*******************************************
    Public Shared Function GetDigit(ByVal Digit As String) As String
        Select Case Val(Digit)
            Case 1 : Digit = "One"
            Case 2 : Digit = "Two"
            Case 3 : Digit = "Three"
            Case 4 : Digit = "Four"
            Case 5 : Digit = "Five"
            Case 6 : Digit = "Six"
            Case 7 : Digit = "Seven"
            Case 8 : Digit = "Eight"
            Case 9 : Digit = "Nine"
            Case Else : Digit = ""
        End Select
        Return Digit
    End Function
End Class
