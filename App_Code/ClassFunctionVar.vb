Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports Classconn
Public Class ClassFunctionVar

    Public Shared ReadOnly Property CRTDT() As Decimal
        Get
            Return Val(Date.Now.ToString("yyyyMMdd"))
        End Get
    End Property
    Public Shared ReadOnly Property CRTTM() As Decimal
        Get
            Return Val(Date.Now.ToString("HHmm"))
        End Get
    End Property
    Public Shared STRDate As String = Val(Date.Now.ToString("yyyyMMdd"))
    Public Shared STRDate2 As String = Val(Date.Now.ToString("ddMMyyyy"))
    Public Shared STRTime As String = Val(Date.Now.ToString("HHmm"))


    Public Shared Function GetMakerCodeAS400(ByVal Library_VDMA As String, ByVal PartNo As String) As String
        Dim MAFMCD As String
        Dim strcmd As String
        strcmd = " Select MAFMCD From  " & Library_VDMA & _
            " Where MAPROD='" & PartNo & "'"
        Dim clConn As New OleDbConnection(Classconn.strCon400)
        If clConn.State = Data.ConnectionState.Closed Then
            clConn.Open()
        End If

        Dim cmdAS400 As New OleDbCommand(strcmd, clConn)
        cmdAS400.CommandText = strcmd
        MAFMCD = cmdAS400.ExecuteScalar()
        cmdAS400.Dispose()
        clConn.Close()
        Return MAFMCD
    End Function
    Public Shared Function GetModelCodeAS400(ByVal Library_VYMC As String, ByVal MAFMCD As String) As String
        Dim strCmd, ModelName As String
        strCmd = " Select LEFT(MCCDAT,15) From  " & Library_VYMC & _
            " Where MCPKEY = 'VYYOB" & MAFMCD & "'"
        Dim clConn As New OleDbConnection(Classconn.strCon400)
        If clConn.State = Data.ConnectionState.Closed Then
            clConn.Open()
        End If

        Dim cmdAS400 As New OleDbCommand(strCmd, clConn)
        ModelName = cmdAS400.ExecuteScalar()
        cmdAS400.Dispose()
        clConn.Close()
        Return ModelName
    End Function
    Public Shared Function GetSuffixAS400(ByVal Library_TICH As String, ByVal PartNo As String, ByVal OrderDate As Decimal) As String
        Dim strCmd As String
        Dim PartSuffix As String = ""
        Dim CntRec As Decimal
        strCmd = "SELECT COUNT(*)   FROM   " & Library_TICH
        strCmd += " Where CHPART ='" & PartNo & "' AND CHSTRD <= " & OrderDate
        Dim clConn As New OleDbConnection(Classconn.strCon400)
        If clConn.State = Data.ConnectionState.Closed Then
            clConn.Open()
        End If

        Dim cmdAS400 As New OleDbCommand(strCmd, clConn)
        CntRec = cmdAS400.ExecuteScalar()
        PartSuffix = ""
        If CntRec > 0 Then
            strCmd = " Select MAX(CHMKSF),MAX(CHYZSF) From  " & Library_TICH
            strCmd += " Where CHCHG ='" & PartNo & "' AND CHSTRD <= " & OrderDate
            cmdAS400.CommandText = strCmd
            PartSuffix = cmdAS400.ExecuteScalar()
        End If
        cmdAS400.Dispose()
        clConn.Close()
        Return (PartSuffix)
    End Function
    Public Shared Function GetCarlineFromYPRD(ByVal Library_YPRD As String, ByVal ProductNo As String) As String
        Dim strCmd, CarlineCode As String
        CarlineCode = ""
        strCmd = " Select PCARLN From  " & Library_YPRD & _
            " Where PPROD = '" & ProductNo & "'"
        Dim clConn As New OleDbConnection(Classconn.strCon400)
        If clConn.State = Data.ConnectionState.Closed Then
            clConn.Open()
        End If
        Dim cmdAS400 As New OleDbCommand(strCmd, clConn)
        CarlineCode = cmdAS400.ExecuteScalar()
        cmdAS400.Dispose()
        clConn.Close()
        Return CarlineCode
    End Function
    Public Shared Function GetDeliveryTime(ByVal Library_MRTime As String, ByVal ProductNo As String, ByVal ETA_TAP As String, _
                                           ByVal MakerCode As String, ByVal CarlineCode As String, ByVal ETD_TAP As String) As String
        Dim strCmd, DeliveryTime As String
        CarlineCode = ""
        strCmd = " Select MPRD From  " & Library_MRTime
        strCmd += " Where MMAKER = '" & MakerCode & "'"
        If CarlineCode <> "" Then
            strCmd += " And MCALN = '" & CarlineCode & "'"
        End If
        If ProductNo <> "" Then
            strCmd += " And MPROD = '" & CarlineCode & "'"
        End If
        If ETA_TAP <> "" Then
            strCmd += " And META >= '" & Trim(ETA_TAP) & "'"
        End If
        If ETD_TAP <> "" Then
            strCmd += " And METD <= '" & Trim(ETD_TAP) & "'"
        End If
        Dim clConn As New OleDbConnection(Classconn.strCon400)
        If clConn.State = Data.ConnectionState.Closed Then
            clConn.Open()
        End If
        Dim cmdAS400 As New OleDbCommand(strCmd, clConn)
        DeliveryTime = cmdAS400.ExecuteScalar()
        cmdAS400.Dispose()
        clConn.Close()
        Return DeliveryTime
    End Function
    Public Shared Function GetSavePath(ByVal Maker As String) As String
        Dim SavePath As String
        Select Case Maker
            Case "M01"
                SavePath = ("~/Uploads/TMT/")
            Case "M02"
                SavePath = ("~/Uploads/HINO/")
            Case "M03"
                SavePath = ("~/Uploads/ISUZU/")
            Case "M04"
                SavePath = ("~/Uploads/NISSAN/")
            Case "M07"
                SavePath = ("~/Uploads/HONDA/")
            Case "M08"
                SavePath = ("~/Uploads/AAT/")
            Case "M11"
                SavePath = ("~/Uploads/MMth/")
            Case Else
                SavePath = ("~/Uploads/OTHERS/")
        End Select
        Return (SavePath)
    End Function
    Public Shared Function GetPriceTAP(ByVal Library As String, ByVal PriceType As String, _
                                       ByVal NormalSpare As String, _
                                       ByVal PartNo As String, _
                                       ByVal Customer As String) As Data.DataTable
        PartNo = Replace(PartNo, "-", "")
        PartNo = Replace(PartNo, " ", "")
        Dim strSql As String = ""
        Dim dt As New Data.DataTable
        Dim clConn As New OleDbConnection(Classconn.strCon400)
        If clConn.State = Data.ConnectionState.Closed Then
            clConn.Open()
        End If
        strSql = "SELECT * FROM " & Library & ".TVDMC "
        strSql += " WHERE MCPROD='" & PartNo & "' AND MCNMSP='" & NormalSpare & "' "
        strSql += " AND MCDOEX='" & PriceType & "' AND MCLDFL <> '1' AND MCCUST= " & Customer
        strSql += " ORDER BY MCSTDT DESC "
        Dim oAdap As New Data.OleDb.OleDbDataAdapter(strSql, clConn)
        oAdap.Fill(dt)
        clConn.Close()
        Return dt
    End Function
    Public Shared Function GetDataUser(ByVal pUSERID As String, ByVal pPGMID As String) As DataTable
        Dim strSql As String = ""
        Dim ConnSql As New OleDbConnection(Classconn.strConnSql)
        Dim dt As New DataTable
        If ConnSql.State = ConnectionState.Closed Then
            ConnSql.Open()
        End If
        strSql = "SELECT DISTINCT UID, PLANT, DEFCOM, ULEVEL, MNAME FROM WEBSALEPERMIS A "
        strSql += "INNER JOIN WEBSALEMENU B ON A.MENUID = B.MID "
        strSql += "INNER JOIN USERS C ON A.UID = C.USERID "
        strSql += "WHERE UID='" & pUSERID & "' "
        strSql += "AND menuid='" & pPGMID & "' "
        Dim da As New OleDbDataAdapter(strSql, ConnSql)
        da.Fill(dt)
        ConnSql.Close()
        Return dt
    End Function
    Public Shared Function GetCCodeFromCustoms(ByVal Library As String, ByVal pPKey As String) As String
        Dim strCmd, CustomsCode As String
        CustomsCode = ""
        strCmd = " Select CCData From  " & Library & _
            " Where trim(CCDKEY) = '" & Trim(pPKey) & "'"
        Dim clConn As New OleDbConnection(Classconn.strCon400)
        If clConn.State = Data.ConnectionState.Closed Then
            clConn.Open()
        End If
        Dim cmdAS400 As New OleDbCommand(strCmd, clConn)
        CustomsCode = cmdAS400.ExecuteScalar()
        cmdAS400.Dispose()
        clConn.Close()
        Return CustomsCode
    End Function
    Public Shared Function GetBarcode(ByVal Code As String, ByRef Result As String)
        Select Case Code
            Case "0"
                Result = "0"
            Case "1"
                Result = "1"
            Case "2"
                Result = "2"
            Case "3"
                Result = "3"
            Case "4"
                Result = "4"
            Case "5"
                Result = "5"
            Case "6"
                Result = "6"
            Case "7"
                Result = "7"
            Case "8"
                Result = "8"
            Case "9"
                Result = "9"
            Case "A"
                Result = "10"
            Case "B"
                Result = "11"
            Case "C"
                Result = "12"
            Case "D"
                Result = "13"
            Case "E"
                Result = "14"
            Case "F"
                Result = "15"
            Case "G"
                Result = "16"
            Case "H"
                Result = "17"
            Case "I"
                Result = "18"
            Case "J"
                Result = "19"
            Case "K"
                Result = "20"
            Case "L"
                Result = "21"
            Case "M"
                Result = "22"
            Case "N"
                Result = "23"
            Case "O"
                Result = "24"
            Case "P"
                Result = "25"
            Case "Q"
                Result = "26"
            Case "R"
                Result = "27"
            Case "S"
                Result = "28"
            Case "T"
                Result = "29"
            Case "U"
                Result = "30"
            Case "V"
                Result = "31"
            Case "W"
                Result = "32"
            Case "X"
                Result = "33"
            Case "Y"
                Result = "34"
            Case "Z"
                Result = "35"
            Case "-"
                Result = "36"
            Case "."
                Result = "37"
            Case " "
                Result = "38"
            Case "$"
                Result = "39"
            Case "/"
                Result = "40"
            Case "+"
                Result = "41"
            Case "%"
                Result = "42"
        End Select
        Return Result
    End Function

    Public Shared Function ResultBarcode(ByVal Code As String, ByRef Result As String)
        Select Case Code
            Case "0"
                Result = "0"
            Case "1"
                Result = "1"
            Case "2"
                Result = "2"
            Case "3"
                Result = "3"
            Case "4"
                Result = "4"
            Case "5"
                Result = "5"
            Case "6"
                Result = "6"
            Case "7"
                Result = "7"
            Case "8"
                Result = "8"
            Case "9"
                Result = "9"
            Case "10"
                Result = "A"
            Case "11"
                Result = "B"
            Case "12"
                Result = "C"
            Case "13"
                Result = "D"
            Case "14"
                Result = "E"
            Case "15"
                Result = "F"
            Case "16"
                Result = "G"
            Case "17"
                Result = "H"
            Case "18"
                Result = "I"
            Case "19"
                Result = "J"
            Case "20"
                Result = "K"
            Case "21"
                Result = "L"
            Case "22"
                Result = "M"
            Case "23"
                Result = "N"
            Case "24"
                Result = "O"
            Case "25"
                Result = "P"
            Case "26"
                Result = "Q"
            Case "27"
                Result = "R"
            Case "28"
                Result = "S"
            Case "29"
                Result = "T"
            Case "30"
                Result = "U"
            Case "31"
                Result = "V"
            Case "32"
                Result = "W"
            Case "33"
                Result = "X"
            Case "34"
                Result = "Y"
            Case "35"
                Result = "Z"
            Case "36"
                Result = "-"
            Case "37"
                Result = "."
            Case "38"
                Result = " "
            Case "39"
                Result = "$"
            Case "40"
                Result = "/"
            Case "41"
                Result = "+"
            Case "42"
                Result = "%"
        End Select
        Return Result
    End Function
    Public Shared Function checkasscii(ByVal strRecord As String)   '****** check ASSCII CODE in text file ********

        Dim cASC As Integer = Asc(Right(strRecord, 1).ToString)
        If cASC = 13 Then
            strRecord = Trim(Replace(strRecord, (Right(strRecord, 1).ToString), ""))
        Else
            strRecord = strRecord.Trim
        End If
        Return strRecord
    End Function
End Class
