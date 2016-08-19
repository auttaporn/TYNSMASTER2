Imports Microsoft.VisualBasic
'Imports System.Data

Public Class Forecast
    Private strWSConn As String = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=dbconnect;Password=db2005;Initial Catalog=YGSS112003;Data Source=thyzws1036;"
    Private strASConn As String = "Provider=IBMDA400;Data Source=10.200.1.5;User ID=TY#00006;Password=L333;"
    Private strPMONTH As String
    Private FactoryCode As String
    Private MakerCode As String
    Private ForecastType As String
    Private USERID As String
    Private STRUNAME As String

    Private Function Clear400(ByVal strMONTH As String, ByVal strMaker As String)
        Dim oComm As New OleDb.OleDbCommand
        Dim strSQL As String
        Dim int As Integer
        strSQL = "DELETE FROM ISUZU.MMOW WHERE FCRENO ='" & strMaker & "' AND FCDREC ='" & strMONTH & "'"
        oComm.Connection = OpenConn(strASConn)
        oComm.CommandText = strSQL
        int = oComm.ExecuteNonQuery()
    End Function
    Function TranTo400(ByVal strMONTH As String, ByVal strMaker As String)
        Dim strSQL As String
        Dim oReader As OleDb.OleDbDataReader
        Dim oComm As New OleDb.OleDbCommand
        Dim oAdapt As New OleDb.OleDbDataAdapter
        Dim oDset As New DataSet
        Dim objCB As OleDb.OleDbCommandBuilder
        Dim oConnWS As New OleDb.OleDbConnection
        Dim oConnAS As New OleDb.OleDbConnection
        Dim oNewRow As DataRow
        'Dim strLib As String = "#WINIJ"
        Dim strLib As String = "ISUZU"
        oConnWS = OpenConn(strWSConn)
        oConnAS = OpenConn(strASConn)
        strSQL = "Select PMONTH,FACCODE,MAKER,PARTNO,PRROCESS,SHOPCODE,RECAREA,LASTLOT,SNEPPACK,SNEPORDER,SUM(GMON1)as GMON1,SUM(GMON2) as GMON2,SUM(GMON3) as GMON3 " & _
                 ",SUM(N4)as N4,SUM(N5)as N5,CRTDTE,UPUSER FROM MMO_HONDA WHERE MAKER ='" & strMaker & "' AND PMONTH ='" & strMONTH & "'" & _
                 " GROUP BY PMONTH,FACCODE,MAKER,PARTNO,PRROCESS,SHOPCODE,RECAREA,LASTLOT,SNEPPACK,SNEPORDER,CRTDTE,UPUSER"
        oComm.Connection = OpenConn(strWSConn)
        oComm.CommandText = strSQL
        oReader = oComm.ExecuteReader
        Clear400(strMONTH, strMaker)
        strSQL = "Select * FROM " & strLib & ".MMOW WHERE FCRENO ='" & strMaker & "' AND FCCONS ='" & strMONTH & "'"
        oAdapt = New OleDb.OleDbDataAdapter(strSQL, oConnAS)
        objCB = New OleDb.OleDbCommandBuilder(oAdapt)
        oAdapt.Fill(oDset, "MMOW")
        Try
            Do While oReader.Read
                oNewRow = oDset.Tables("MMOW").NewRow()
                oNewRow("FCDREC") = IIf(oReader("PMONTH") Is System.DBNull.Value, 0, oReader("PMONTH"))
                oNewRow("FCFACT") = IIf(oReader("FACCODE") Is System.DBNull.Value, "", oReader("FACCODE"))
                oNewRow("FCRENO") = IIf(oReader("MAKER") Is System.DBNull.Value, "", oReader("MAKER"))
                'oNewRow("FC") = oReader("ITEM")
                oNewRow("FCPART") = IIf(oReader("PARTNO") Is System.DBNull.Value, "", oReader("PARTNO"))
                oNewRow("FCPROC") = IIf(oReader("PRROCESS") Is System.DBNull.Value, "", oReader("PRROCESS"))
                oNewRow("FCPLNR") = IIf(oReader("SHOPCODE") Is System.DBNull.Value, "", oReader("SHOPCODE"))
                oNewRow("FCPREC") = IIf(oReader("RECAREA") Is System.DBNull.Value, "", oReader("RECAREA"))
                oNewRow("FCRDOC") = IIf(oReader("LASTLOT") Is System.DBNull.Value, "", oReader("LASTLOT"))
                oNewRow("FCCONS") = IIf(oReader("SNEPPACK") Is System.DBNull.Value, "", oReader("SNEPPACK"))
                oNewRow("FCCHG") = IIf(oReader("SNEPORDER") Is System.DBNull.Value, "", oReader("SNEPORDER"))
                oNewRow("FCW6TQ") = IIf(oReader("GMON1") Is System.DBNull.Value, 0, oReader("GMON1"))
                oNewRow("FCN2TQ") = IIf(oReader("GMON2") Is System.DBNull.Value, 0, oReader("GMON2"))
                oNewRow("FCN3TQ") = IIf(oReader("GMON3") Is System.DBNull.Value, 0, oReader("GMON3"))
                'oNewRow("FCW6TQ") = oReader("MN3N1")
                'oNewRow("FCN2TQ") = oReader("MN3N2")
                'oNewRow("FCN3TQ") = oReader("MN3N3")
                oNewRow("FCW66D") = IIf(oReader("PMONTH") Is System.DBNull.Value, 0, oReader("PMONTH"))
                oNewRow("FCN4TQ") = IIf(oReader("N4") Is System.DBNull.Value, 0, oReader("N4"))
                oNewRow("FCW5TQ") = IIf(oReader("N5") Is System.DBNull.Value, 0, oReader("N5"))
                oNewRow("FCCRTD") = IIf(oReader("CRTDTE") Is System.DBNull.Value, 0, oReader("CRTDTE"))
                '                oNewRow("FCNPAT") = IIf(oReader("UPUSER") Is System.DBNull.Value, 0, oReader("UPUSER"))
                oDset.Tables("MMOW").Rows.Add(oNewRow)
            Loop
            oAdapt.Update(oDset, "MMOW")
        Catch ex As Exception
            Return (ex.Source & " " & ex.Message)
        End Try
    End Function
    Private Function CONtoDec(ByVal SOURSE) As String
        Dim NULL = System.DBNull.Value
        If SOURSE Is NULL Then
            Return 0
        Else
            Return IIf(IsNumeric(SOURSE), SOURSE, 0)
        End If
    End Function
    Private Function ClearF()
        Dim oComm As New OleDb.OleDbCommand
        Dim strSQL As String
        strSQL = "DELETE FROM MMO_HONDA WHERE MAKER ='" & MakerCode & "' AND PMONTH ='" & strPMONTH & "'"
        oComm.Connection = OpenConn(strWSConn)
        oComm.CommandText = strSQL
        oComm.ExecuteNonQuery()
    End Function
    Private Function TransNDS(ByVal oDataset As DataSet)
        Dim strSQL As String
        Dim i As Integer
        Dim oAdapt As New OleDb.OleDbDataAdapter
        Dim oDset As New DataSet
        Dim oConn As New OleDb.OleDbConnection
        Dim objCB As OleDb.OleDbCommandBuilder
        Dim oSRCRow As DataRow
        Dim oNewRow As DataRow

        oConn = OpenConn(strWSConn)
        strSQL = "Select * FROM MMO_HONDA WHERE MAKER ='" & MakerCode & "' AND PMONTH ='" & strPMONTH & "'"
        oAdapt = New OleDb.OleDbDataAdapter(strSQL, oConn)
        objCB = New OleDb.OleDbCommandBuilder(oAdapt)
        oAdapt.Fill(oDset, "MMO_HONDA")
        i = 1
        Try
            For Each oSRCRow In oDataset.Tables("EXCEL").Rows
                If Not oSRCRow.Item(2) Is System.DBNull.Value Then
                    If oSRCRow.Item(2).ToString <> "" Then
                        oNewRow = oDset.Tables("MMO_HONDA").NewRow()
                        oNewRow("PMONTH") = strPMONTH
                        oNewRow("FACCODE") = FactoryCode
                        oNewRow("MAKER") = MakerCode
                        oNewRow("ITEM") = i
                        oNewRow("PARTNO") = Left(oSRCRow.Item(2), 13)
                        oNewRow("PARTNAME") = Trim(oSRCRow.Item(3))
                        oNewRow("PRROCESS") = ""
                        oNewRow("SHOPCODE") = ""
                        oNewRow("RECAREA") = ""
                        oNewRow("LASTLOT") = ""
                        oNewRow("SNEPPACK") = 0
                        oNewRow("SNEPORDER") = 0
                        oNewRow("DN101") = 0 : oNewRow("DN102") = 0 : oNewRow("DN103") = 0 : oNewRow("DN104") = 0
                        oNewRow("DN105") = 0 : oNewRow("DN106") = 0 : oNewRow("DN107") = 0 : oNewRow("DN108") = 0
                        oNewRow("DN109") = 0 : oNewRow("DN110") = 0 : oNewRow("DN111") = 0 : oNewRow("DN112") = 0
                        oNewRow("DN113") = 0 : oNewRow("DN114") = 0 : oNewRow("DN115") = 0 : oNewRow("DN116") = 0
                        oNewRow("DN117") = 0 : oNewRow("DN118") = 0 : oNewRow("DN119") = 0 : oNewRow("DN120") = 0
                        oNewRow("DN121") = 0 : oNewRow("DN122") = 0 : oNewRow("DN123") = 0 : oNewRow("DN124") = 0
                        oNewRow("DN128") = 0 : oNewRow("DN129") = 0 : oNewRow("DN130") = 0 : oNewRow("DN131") = 0
                        oNewRow("DN201") = 0 : oNewRow("DN202") = 0 : oNewRow("DN203") = 0 : oNewRow("DN204") = 0
                        oNewRow("DN205") = 0 : oNewRow("DN206") = 0 : oNewRow("DN207") = 0 : oNewRow("DN208") = 0
                        oNewRow("DN209") = 0 : oNewRow("DN210") = 0 : oNewRow("DN211") = 0 : oNewRow("DN212") = 0
                        oNewRow("DN213") = 0 : oNewRow("DN214") = 0 : oNewRow("DN215") = 0 : oNewRow("DN216") = 0
                        oNewRow("DN217") = 0 : oNewRow("DN218") = 0 : oNewRow("DN219") = 0 : oNewRow("DN220") = 0
                        oNewRow("DN222") = 0 : oNewRow("DN223") = 0 : oNewRow("DN224") = 0 : oNewRow("DN225") = 0
                        oNewRow("DN227") = 0 : oNewRow("DN228") = 0 : oNewRow("DN229") = 0 : oNewRow("DN230") = 0
                        oNewRow("DN231") = 0 : oNewRow("DN301") = 0 : oNewRow("DN302") = 0
                        oNewRow("DN303") = 0 : oNewRow("DN304") = 0 : oNewRow("DN305") = 0 : oNewRow("DN306") = 0
                        oNewRow("DN307") = 0 : oNewRow("DN308") = 0 : oNewRow("DN309") = 0 : oNewRow("DN310") = 0
                        oNewRow("DN311") = 0 : oNewRow("DN312") = 0 : oNewRow("DN313") = 0 : oNewRow("DN314") = 0
                        oNewRow("DN315") = 0 : oNewRow("DN316") = 0 : oNewRow("DN317") = 0 : oNewRow("DN318") = 0
                        oNewRow("DN319") = 0 : oNewRow("DN320") = 0 : oNewRow("DN321") = 0 : oNewRow("DN322") = 0
                        oNewRow("DN323") = 0 : oNewRow("DN324") = 0 : oNewRow("DN325") = 0 : oNewRow("DN326") = 0
                        oNewRow("DN327") = 0 : oNewRow("DN328") = 0 : oNewRow("DN329") = 0 : oNewRow("DN330") = 0
                        oNewRow("DN402") = 0 : oNewRow("DN403") = 0 : oNewRow("DN404") = 0 : oNewRow("WN101") = 0
                        oNewRow("WN102") = 0 : oNewRow("WN103") = 0
                        oNewRow("WN104") = 0 : oNewRow("WN105") = 0 : oNewRow("WN201") = 0 : oNewRow("WN202") = 0
                        oNewRow("WN301") = 0 : oNewRow("WN302") = 0 : oNewRow("WN303") = 0 : oNewRow("WN304") = 0
                        oNewRow("WN305") = 0
                        oNewRow("GMON1") = IIf(IsNumeric(Trim(oSRCRow.Item(4))), Trim(oSRCRow.Item(4)), 0)
                        oNewRow("GMON2") = IIf(IsNumeric(Trim(oSRCRow.Item(5))), Trim(oSRCRow.Item(5)), 0)
                        oNewRow("GMON3") = IIf(IsNumeric(Trim(oSRCRow.Item(6))), Trim(oSRCRow.Item(6)), 0)
                        oNewRow("MN3N1") = 0
                        oNewRow("MN3N2") = 0
                        oNewRow("MN3N3") = 0
                        oNewRow("N4") = IIf(IsNumeric(Trim(oSRCRow.Item(7))), Trim(oSRCRow.Item(7)), 0)
                        oNewRow("N5") = IIf(IsNumeric(Trim(oSRCRow.Item(8))), Trim(oSRCRow.Item(8)), 0)
                        oNewRow("CRTDTE") = Year(System.DateTime.Now) & Right("0" & Month(System.DateTime.Now), 2) & Right("0" & Day(System.DateTime.Now), 2)
                        oNewRow("UPUSER") = USERID
                        oDset.Tables("MMO_HONDA").Rows.Add(oNewRow)
                        oAdapt.Update(oDset, "MMO_HONDA")
                        i += 1
                    End If
                End If
            Next

        Catch ex As Exception
            Return ex.Message
        End Try
        oAdapt.Dispose()
        oDset.Dispose()
        oConn.Dispose()
        objCB.Dispose()
    End Function
    Function TansData(ByVal strPath As String, ByVal strFile As String) As DataSet
        Dim strSQL As String
        Dim strType As String
        Dim oAdapt As New OleDb.OleDbDataAdapter
        Dim oDset As New DataSet
        Dim oConn As New OleDb.OleDbConnection
        Dim strXSLConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & ";Extended Properties=""Excel 8.0;HDR=1;IMEX=1"""
        oConn = OpenConn(strXSLConn)
        strType = Left(strFile, 3)
        strSQL = "Select * From [M07$A2:I]"
        ClearF()
        oAdapt = New OleDb.OleDbDataAdapter(strSQL, oConn)
        oAdapt.Fill(oDset, "EXCEL")
        TransNDS(oDset)
        TranTo400(strPMONTH, MakerCode)
        Return oDset
    End Function
    Function OpenConn(ByVal strConString As String) As OleDb.OleDbConnection
        Dim oConn As New OleDb.OleDbConnection
        oConn.ConnectionString = strConString
        oConn.Open()
        Return oConn
    End Function
    Function CheckUser(ByVal UID As String, ByVal upass As String) As Boolean
        Dim oConn As New OleDb.OleDbConnection
        Dim oReader As OleDb.OleDbDataReader
        Dim oComm As New OleDb.OleDbCommand
        Dim strSQL As String
        Dim oForecast As New Forecast
        oConn = oForecast.OpenConn(strWSConn)
        strSQL = "SELECT * FROM V_USERS WHERE USERID = '" & UID & "' AND PASSWORD= '" & upass & "' AND MENUID LIKE 'MTAPDM07%'"
        oComm.Connection = oConn
        oComm.CommandText = strSQL
        oReader = oComm.ExecuteReader
        If oReader.Read Then
            STRUNAME = oReader("EMPNAME")
            USERID = oReader("UserID")
            Return True
        Else
            Return False
        End If
    End Function
#Region "Property"
    Property ConnString(ByVal SV As String)

        Get
            Select Case SV
                Case "AS"
                    ConnString = strASConn
                Case "WS"
                    ConnString = strWSConn
            End Select
        End Get

        Set(ByVal Value)
            Select Case SV
                Case "AS"
                    strASConn = Value
                Case "WS"
                    strWSConn = Value
            End Select
        End Set
    End Property
    Property PMONTH()
        Get
            PMONTH = strPMONTH
        End Get
        Set(ByVal Value)
            strPMONTH = Value
        End Set
    End Property
    Property Factory()
        Get
            Factory = FactoryCode
        End Get
        Set(ByVal Value)
            FactoryCode = Value
        End Set
    End Property
    Property Maker()
        Get
            Maker = MakerCode
        End Get
        Set(ByVal Value)
            MakerCode = Value
        End Set
    End Property
    Property Forecast()
        Get
            Forecast = ForecastType
        End Get
        Set(ByVal Value)
            ForecastType = Value
        End Set
    End Property
    Property UID()
        Get
            UID = USERID
        End Get
        Set(ByVal Value)
            USERID = Value
        End Set
    End Property
    ReadOnly Property Uname()
        Get
            Uname = STRUNAME
        End Get
    End Property
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region
End Class
