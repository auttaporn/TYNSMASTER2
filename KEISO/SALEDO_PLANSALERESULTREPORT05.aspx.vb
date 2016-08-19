Imports System.Data
Imports System.Data.OleDb
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Partial Class SALEDO_PLANSALERESULTREPORT05
    Inherits System.Web.UI.Page
    Private Conn As New OleDbConnection(Classconn.strConnSql)
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("R5-REPORT.xlt")
    Private ReportName As String = "DOMESTIC-REPORT5"
    Private crtDate As String = Val(Date.Now.ToString("ddMMyyyy"))
    Private crtTime As String = Val(Date.Now.ToString("HHmm"))
    Private oTable As New DataTable
    Private oRow As DataRow
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = ExcelEngine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private Filename, FrMonth As String
    Private curMKG, curPDTP, curPTNM, curPDNM, curPDCD, curKGP, curGPN As String
    Private intCurrow, intNext, intMonthNo, intRun, intPos, ampos(12) As Integer
    Private aspa(12), aspp(12), asra(12), asrp(12) As Double
    Private aptpa(12), aptpp(12), aptra(12), aptrp(12) As Double
    Private apgpa(12), apgpp(12), apgra(12), apgrp(12) As Double
    Private dsumspa, dsumspp, dsumsra, dsumsrp As Double

    Protected Sub SetPlant()
        Dim oTable As New DataTable
        Dim oRow As DataRow
        oTable = ClassFunctionVar.GetDataUser(Session("USERID"), "MTAPDADR05")
        If oTable.Rows.Count = 0 Then
            ddlCompany.Text = "No Permission"
            Exit Sub
        End If
        oRow = oTable.Rows(0)
        ddlCompany.DataTextField = "cname"
        ddlCompany.DataValueField = "hicocd"
        ddlCompany.DataSource = getCompany(Trim(oRow("ULEVEL")), Trim(oRow("DEFCOM")))
        ddlCompany.DataBind()
    End Sub
    Protected Function getCompany(ByVal iCase As String, ByVal ShortName As String) As DataTable
        Dim strsql As String
        strsql = "select distinct a.hicocd,b.companycode, "
        strsql += "b.shortname + ' : ' + b.longname  as cname from yssh a "
        strsql += "inner join company b on a.hicocd=b.companycode "

        Dim strdate As String = ClassFunctionVar.STRDate
        CommonFunction.CheckDate(strdate, "", "yyyymmdd", "YENG2")
        Select Case iCase
            Case "1"
                strsql += "where and b.shortname = '" & ShortName & "' "
            Case "5"
                ' If CInt(strdate) <= 20100403 Then
                strsql += "where hicocd in ('A','E')"
                ' End If
        End Select
        strsql += "order by hicocd asc "

        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Dim da As New OleDbDataAdapter(strsql, conn)
        Dim dt As New DataTable
        da.Fill(dt)
        Return dt
    End Function
    Protected Function GetData() As DataTable
        Dim strSql As String = ""
        Dim Condition1 As String = ""
        Dim Condition2 As String = ""
        If Conn.State = ConnectionState.Open Then
            Conn.Close()
        End If
        Conn.Open()
        If (ddlCompany.SelectedValue <> "ALL") And (ddlCompany.SelectedValue <> "") Then
            Condition1 = " AND HICOCD = '" & ddlCompany.SelectedValue & "' "
            Condition2 = " AND PCOCD = '" & ddlCompany.SelectedValue & "' "
        End If
        Condition1 += " AND  HIBLDT <= '" & txtYear.Text & Right("00" & ddlTo.SelectedValue, 2) & "31" & "' "
        Condition1 += " AND  HIBLDT >= '" & txtYear.Text & Right("00" & ddlFrom.SelectedValue, 2) & "01" & "' "
        If ddlBillto.SelectedValue <> "0" Then
            Condition1 += " AND  HIRPFG = '" & ddlBillto.SelectedValue & "' "
        End If
        Condition2 += " AND  PMONTH >= '" & ddlFrom.SelectedValue & "'  AND pmonth <= '" & ddlTo.SelectedValue & "' "

        Dim Sql1 As String = "ROUND(ISNULL(SUM(ROUND(CASE WHEN YSSH.HIDTTP = 'I' THEN YSSD.DIFTAM * YSSH.HIFIXR ELSE 0 END " & _
                             "- CASE WHEN YSSH.HIDTTP = 'C' THEN YSSD.DIFTAM * YSSH.HIFIXR ELSE 0 END " & _
                             "+ CASE WHEN YSSH.HIDTTP = 'D' THEN YSSD.DIFTAM * YSSH.HIFIXR ELSE 0 END, 2)), 0), 2) AS SALEAMOUNT, " & _
                             "isnull(SUM(ROUND(CASE WHEN YSSH.HIDTTP = 'I' THEN YSSD.DIFCAM ELSE 0 END " & _
                             "- CASE WHEN YSSH.HIDTTP = 'C' THEN YSSD.DIFCAM ELSE 0 END " & _
                             "+ CASE WHEN YSSH.HIDTTP = 'D' THEN YSSD.DIFCAM ELSE 0 END " & _
                             "+ CASE WHEN YSSH.HIDTTP = 'A' THEN YSSD.DIFCAM ELSE 0 END, 2)), 0) AS SALECOST " & _
                             "FROM  YSSH "
        Dim Sql2 As String = "WHERE (YSSH.HICAFG <> 'Y') AND (YSSH.HIIVFG = 'D') AND " & _
                             "LEFT(HIBLDT,4) = '" & txtYear.Text & "' " & Condition1
        Dim Sql3 As String = "WHERE (YSSP2.PTYPE = 'DO') AND (YSSP2.PYEAR = '" & txtYear.Text & "') AND (YSSP2.PAMT > 0) " & Condition2

        Select Case ddlGroupby.SelectedValue
            Case "MKP"
                strSql = "SELECT 'SR'AS TYPE,LEFT(HIBLDT,4) AS ATYEAR, " & _
                         "CAST(SUBSTRING(CAST(HIBLDT AS CHAR(8)),5,2)AS INT) AS ATMONTH, MGP.CNAME, HIMKGP, "
                strSql += Sql1
                strSql += "INNER JOIN YSSD ON YSSH.HIDONO = YSSD.DIDONO  AND YSSH.HICOCD = YSSD.DICOCD  "
                strSql += "INNER JOIN v_YTAPMGP MGP ON YSSH.HIMKGP = MGP.CKEY  "
                strSql += Sql2
                strSql += " GROUP BY LEFT(HIBLDT,4),CAST(SUBSTRING(CAST(HIBLDT AS CHAR(8)),5,2)AS INT),HIMKGP,MGP.CNAME "
                strSql += "UNION "
                strSql += "SELECT 'SP' AS TYPE, YSSP2.PYEAR AS ATYEAR, YSSP2.PMONTH AS ATMONTH, MGP.CNAME, LEFT(YSSP2.PCUST, 3) AS HIMKGP, "
                strSql += "SUM(YSSP2.PAMT) AS SALEAMOUNT, SUM(YSSP2.PPRF) AS SALECOST "
                strSql += "FROM YSSP2 "
                strSql += "INNER JOIN v_YTAPMGP MGP ON LEFT(YSSP2.PCUST, 3) = MGP.CKEY "
                strSql += Sql3
                strSql += "GROUP BY LEFT(YSSP2.PCUST, 3), YSSP2.PYEAR, YSSP2.PMONTH, MGP.CNAME "
                strSql += "ORDER BY HIMKGP, ATMONTH "
            Case "PDTP"
                strSql = "SELECT 'SR'AS TYPE,LEFT(HIBLDT,4) AS ATYEAR, " & _
                          "CAST(SUBSTRING(CAST(HIBLDT AS CHAR(8)),5,2)AS INT) AS ATMONTH, MGP.CNAME, HIMKGP,  "
                strSql += "LEFT(DIPDCD,3) as DIPDCD, isnull(PRD.CNAME,'') AS PDNM, DIPDTP,  isnull(PRD.CFLAG2,'') as PTNAME,  isnull(PRD.PDGPN,'') as PDGPN,  "
                strSql += Sql1
                strSql += "INNER JOIN YSSD ON YSSH.HIDONO = YSSD.DIDONO AND YSSH.HICOCD= YSSD.DICOCD "
                strSql += "LEFT OUTER JOIN v_YTAPMGP MGP ON YSSH.HIMKGP = MGP.CKEY "
                strSql += "LEFT OUTER JOIN v_YTAPPRD PRD ON LEFT(YSSD.DIPDCD,3) = PRD.CKEY "
                strSql += Sql2
                strSql += "GROUP BY LEFT(HIBLDT,4),CAST(SUBSTRING(CAST(HIBLDT AS CHAR(8)),5,2)AS INT), PRD.CNAME,MGP.CNAME, HIMKGP, "
                strSql += "LEFT(DIPDCD,3), DIPDTP, PRD.CFLAG2, PRD.PDGPN "
                strSql += "UNION "
                strSql += "SELECT 'SP' AS TYPE, YSSP2.PYEAR AS ATYEAR, YSSP2.PMONTH AS ATMONTH, MGP.CNAME, LEFT(YSSP2.PCUST, 3) AS HIMKGP, "
                strSql += "PPRDCD AS DIPDTP, isnull(PRD.CNAME,'') AS PDNM, PPRDTP, isnull(PRD.CFLAG2,'') as PTNAME,  isnull(PRD.PDGPN,'') as PDGPN, "
                strSql += "SUM(YSSP2.PAMT) AS SALEAMOUNT, SUM(YSSP2.PPRF) AS SALECOST "
                strSql += "FROM YSSP2 "
                strSql += "INNER JOIN v_YTAPMGP MGP ON LEFT(YSSP2.PCUST, 3) = MGP.CKEY "
                strSql += "LEFT OUTER JOIN v_YTAPPRD PRD ON PPRDCD = PRD.CKEY "
                strSql += Sql3
                strSql += "GROUP BY LEFT(YSSP2.PCUST, 3), YSSP2.PYEAR, YSSP2.PMONTH, PPRDCD, PPRDTP, "
                strSql += "PRD.CNAME, MGP.CNAME, PRD.CFLAG2, PRD.PDGPN "
                strSql += "ORDER BY HIMKGP,  DIPDTP, DIPDCD, ATMONTH, TYPE DESC "
            Case Else
                strSql = "SELECT 'SR'AS TYPE, LEFT(HIBLDT,4) AS ATYEAR, CAST(SUBSTRING(CAST(HIBLDT AS CHAR(8)),5,2)AS INT) AS ATMONTH, "
                strSql += "LEFT(DIPDCD,3) AS DIPDCD, isnull(PRD.CNAME,'') AS PDNM, DIPDTP,  isnull(PRD.CFLAG2,'') AS PTNAME,  isnull(PRD.PDGPN,'') as PDGPN, "
                strSql += Sql1
                strSql += "INNER JOIN YSSD ON YSSH.HIDONO = YSSD.DIDONO AND HICOCD = DICOCD "
                strSql += "LEFT OUTER JOIN v_YTAPPRD prd ON left( YSSD.dipdcd,3) = PRD.CKEY "
                strSql += Sql2
                strSql += "GROUP BY LEFT(HIBLDT,4), CAST(SUBSTRING(CAST(HIBLDT AS CHAR(8)),5,2)AS INT), LEFT(DIPDCD,3),"
                strSql += "PRD.CNAME, DIPDTP,  PRD.CFLAG2,  PRD.PDGPN "
                strSql += "UNION "
                strSql += "SELECT 'SP' AS TYPE, YSSP2.PYEAR AS ATYEAR, YSSP2.PMONTH AS ATMONTH, PPRDCD  AS DIPDCD, isnull(PRD.CNAME,'') AS PDNM, "
                strSql += "PPRDTP, isnull(PRD.CFLAG2,'') AS PTNAME, isnull(PRD.PDGPN,'') as PDGPN, SUM(YSSP2.PAMT) AS SALEAMOUNT, SUM(YSSP2.PPRF) AS SALECOST  "
                strSql += "FROM YSSP2 "
                strSql += "LEFT OUTER JOIN v_YTAPPRD PRD ON PPRDCD = PRD.CKEY "
                strSql += Sql3
                strSql += "GROUP BY  YSSP2.PYEAR, YSSP2.PMONTH, PPRDCD, PPRDTP, PRD.CNAME, PRD.CFLAG2, PRD.PDGPN "
                strSql += "ORDER BY  DIPDTP, DIPDCD, ATMONTH, TYPE DESC  "
        End Select
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(strSql, Conn)
        da.SelectCommand.CommandTimeout = 240
        da.Fill(ds, "PLANSALE")
        Return ds.Tables("PLANSALE")
    End Function

    Protected Sub WriteHMonth()
        Call SetPageProperties()
        With sheet2
            Workbook.Worksheets(0).Range("A1:S10").CopyTo(.Range("A1"))
            If ddlGroupby.SelectedValue = "MKP" Then
                .Range(4, 1).Value = "Unit "
                .Range(4, 3).Value = "1000"
				If ddlCompany.SelectedValue ="E" Then
				 .Range(4,9).Value = "YIC Asia Pacific Corporation Ltd."
				Else
				 .Range(4,9).Value = "Thai Arrow Products Co.,Ltd."
				End if
            ElseIf ddlGroupby.SelectedValue = "PDTP" Then
                .Range(4, 1).Value = "MARKER "
                '.Range(4, 3).Value = Trim(oRow("HIMKGP")) & "   :   " & Trim(oRow("CNAME")
                ''PORR 06/08/2014
                .Range(4, 3).Value = Trim(oRow("HIMKGP")) & "   :   " & IIf(Trim(oRow("CNAME").ToString) Is System.DBNull.Value, "", Trim(oRow("CNAME").ToString))
                .Range(5, 1).Value = "Unit "
                .Range(5, 3).Value = "1000 BATH"
                .Range(6, 1).Value = "PRODUCT CODE"
            Else
                .Range(4, 1).Value = "Unit "
                .Range(4, 3).Value = "1000 BAHT"
                .Range(6, 1).Value = "PRODUCT CODE"
            End If
            .Range("A2").Value = "DOMESTIC"
            .Range("A3").Value = ddlFrom.SelectedItem.ToString & " - " & ddlTo.SelectedItem.ToString & "    " & txtYear.Text
            .Range("O4").Value = "Printdate	: " & Now()
            .Range("O5").Value = "Print by		: " & Session("userId") & "  -  " & Session("username")
            .Range("S5").Value = intMonthNo
            If FrMonth = "0" Then
                .Range("C6").Value = "JANUARY"
                .Range("E6").Value = "FEBRUARY"
                .Range("G6").Value = "MARCH"
                .Range("I6").Value = "APRIL"
                .Range("K6").Value = "MAY"
                .Range("M6").Value = "JUNE"
            Else
                .Range("C6").Value = "JULY"
                .Range("E6").Value = "AUGUST"
                .Range("G6").Value = "SEPTEMBER"
                .Range("I6").Value = "OCTOBER"
                .Range("K6").Value = "NOVEMBER"
                .Range("M6").Value = "DECEMBER"
            End If
        End With
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        If txtYear.Text = "" Then
            lblmessage.Text = "Pleas input year!"
            txtYear.Focus()
            Exit Sub
        End If
        Call WriteData()
    End Sub

    Protected Sub WriteData()
        oTable = GetData()
        If oTable.Rows.Count <= 0 Then
            lblmessage.Text = "No Data"
            Exit Sub
        End If
        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create()
        intMonthNo = 6 'ddlTo.SelectedValue - ddlFrom.SelectedValue + 1
        If ddlFrom.SelectedValue < 7 Then
            FrMonth = 0
        ElseIf ddlFrom.SelectedValue > 6 Then
            FrMonth = 6
        End If
        oRow = oTable.Rows(0)
        Call WriteHMonth()
        Call WriteDetail()
        sheet.Remove()
        Filename = ReportName & "(" & ddlcompany.selectedvalue & "-" & ddlGroupby.SelectedValue & ")-" & crtDate & "-" & crtTime & ".xls"
        CommonFunction.SaveTmpFile(CurTempPath, Filename)
        Workbook.SaveAs(CurTempPath & Filename)
        Workbook.SaveAs(Filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
    End Sub

    Protected Sub WriteDetail()
        intCurrow = 8
        intNext = 11
        ampos(ddlFrom.SelectedValue) = 3
        'Position Column / Month
        For i As Integer = (ddlFrom.SelectedValue + 1) To ddlTo.SelectedValue
            ampos(i) = ampos(i - 1) + 2
        Next
        With sheet2
            If ddlGroupby.SelectedValue = "MKP" Then
                curMKG = Trim(oRow("HIMKGP"))
                .Name = "PLAN AND RESULT REPORT"
                .Range(intCurrow, 1).Value = Trim(oRow("HIMKGP"))
                .Range(intCurrow + 1, 1).Value = Trim(oRow("CNAME"))
                Call WriteMKP()
            ElseIf ddlGroupby.SelectedValue = "PDTP" Then
                If oRow("DIPDTP") IsNot DBNull.Value Then
                    curPDTP = Trim(oRow("DIPDTP"))
                End If
                If oRow("PTNAME") IsNot DBNull.Value Then
                    curPTNM = Trim(oRow("PTNAME"))
                End If
                If oRow("PDNM") IsNot DBNull.Value Then
                    curPDNM = Trim(oRow("PDNM"))
                End If
                If oRow("DIPDCD") IsNot DBNull.Value Then
                    curPDCD = Trim(oRow("DIPDCD"))
                End If
                If oRow("HIMKGP") IsNot DBNull.Value Then
                    curKGP = Trim(oRow("HIMKGP"))
                End If
                If oRow("PDGPN") IsNot DBNull.Value Then
                    curGPN = Trim(oRow("PDGPN"))
                End If
                .Name = Trim(oRow("HIMKGP"))
                .Range(intCurrow, 1).Value = "'" & Trim(oRow("DIPDCD"))
                .Range(intCurrow + 1, 1).Value = oRow("PDNM")
                Call WritePDTP(Workbook, sheet2)
            Else
                If oRow("DIPDTP") IsNot DBNull.Value Then
                    curPDTP = Trim(oRow("DIPDTP"))
                End If
                If oRow("PTNAME") IsNot DBNull.Value Then
                    curPTNM = Trim(oRow("PTNAME"))
                End If
                If oRow("PDNM") IsNot DBNull.Value Then
                    curPDNM = Trim(oRow("PDNM"))
                End If
                If oRow("DIPDCD") IsNot DBNull.Value Then
                    curPDCD = Trim(oRow("DIPDCD"))
                End If
                If oRow("PDGPN") IsNot DBNull.Value Then
                    curGPN = Trim(oRow("PDGPN"))
                End If
                If oRow("DIPDCD") IsNot DBNull.Value Then
                    .Range(intCurrow, 1).Value = "'" & Trim(oRow("DIPDCD"))
                End If
                If oRow("PDNM") IsNot DBNull.Value Then
                    .Range(intCurrow + 1, 1).Value = oRow("PDNM")
                End If
                Call WritePDGP(Workbook, sheet2)
                End If
        End With
    End Sub

    Protected Sub InputSumDetail(ByRef sheet2 As IWorksheet, ByVal iCase As String)
        With sheet2
            Select Case iCase
                Case "AVGDetail"
                    .Range(intCurrow, 17).Value = dsumspa / intMonthNo ' ampos(8)
                    .Range(intCurrow + 1, 17).Value = dsumspp / intMonthNo
                    .Range(intCurrow, 17 + 1).Value = dsumsra / intMonthNo
                    .Range(intCurrow + 1, 17 + 1).Value = dsumsrp / intMonthNo
                Case "AVGGrand"
                    .Range("Q" & intCurrow).Formula = "=O" & intCurrow & "/S5"
                    .Range("Q" & intCurrow + 1).Formula = "=O" & intCurrow + 1 & "/S5"
                    .Range("R" & intCurrow).Formula = "=P" & intCurrow & "/S5"
                    .Range("R" & intCurrow + 1).Formula = "=P" & intCurrow + 1 & "/S5"
            End Select
        End With
    End Sub
    Protected Sub inputSumTotal(ByRef sheet2 As IWorksheet, ByVal iCase As String)
        With sheet2
            intCurrow = intNext
            Workbook.Worksheets(0).Range("A8:S10").CopyTo(.Range("A" & intCurrow))
            .Range("A" & intNext & ":S" & intNext + 2).CellStyle.ColorIndex = 35
            .Range(intCurrow + 1, 1).Value = "TOTAL"
            Select Case iCase
                Case "1"
                    For i As Integer = ddlFrom.SelectedValue To ddlTo.SelectedValue
                        .Range(intCurrow, ampos(i)).Value = aptpa(i)
                        If aptpp(i) <> 0 And aptpp(i) <> 0 Then
                            .Range(intCurrow + 1, ampos(i)).Value = aptpp(i)
                            .Range(intCurrow + 2, ampos(i)).Value = aptpp(i) * 100 / aptpa(i)
                        End If
                        .Range(intCurrow, ampos(i) + 1).Value = aptra(i)
                        If aptra(i) <> 0 And aptra(i) - asrp(i) <> 0 Then
                            .Range(intCurrow + 1, ampos(i) + 1).Value = aptra(i) - aptrp(i)
                            .Range(intCurrow + 2, ampos(i) + 1).Value = (aptra(i) - aptrp(i)) * 100 / aptra(i)
                        End If
                        aptpp(i) = 0
                        aptpa(i) = 0
                        aptrp(i) = 0
                        aptra(i) = 0
                    Next
                Case "2"
                    .Range(intCurrow, 1).Value = curGPN
                    For i As Integer = ddlFrom.SelectedValue To ddlTo.SelectedValue
                        .Range(intCurrow, ampos(i)).Value = apgpa(i)
                        If apgpp(i) <> 0 And apgpp(i) <> 0 Then
                            .Range(intCurrow + 1, ampos(i)).Value = apgpp(i)
                            .Range(intCurrow + 2, ampos(i)).Value = apgpp(i) * 100 / apgpa(i)
                        End If
                        .Range(intCurrow, ampos(i) + 1).Value = apgra(i)
                        If apgra(i) <> 0 And apgra(i) - apgrp(i) <> 0 Then
                            .Range(intCurrow + 1, ampos(i) + 1).Value = apgra(i) - apgrp(i)
                            .Range(intCurrow + 2, ampos(i) + 1).Value = (apgra(i) - apgrp(i)) * 100 / apgra(i)
                        End If
                        apgpa(i) = 0
                        apgpp(i) = 0
                        apgra(i) = 0
                        apgrp(i) = 0
                    Next
                Case "3"
                    .Range(intCurrow, 1).Value = "GRAND"
                    For i As Integer = ddlFrom.SelectedValue To ddlTo.SelectedValue
                        .Range(intCurrow, ampos(i)).Value = aspa(i)
                        If aspp(i) <> 0 And aspp(i) <> 0 Then
                            .Range(intCurrow + 1, ampos(i)).Value = aspp(i)
                            .Range(intCurrow + 2, ampos(i)).Value = aspp(i) * 100 / aspa(i)
                        End If
                        .Range(intCurrow, ampos(i) + 1).Value = asra(i)
                        If asra(i) <> 0 And asra(i) - asrp(i) <> 0 Then
                            .Range(intCurrow + 1, ampos(i) + 1).Value = asra(i) - asrp(i)
                            .Range(intCurrow + 2, ampos(i) + 1).Value = (asra(i) - asrp(i)) * 100 / asra(i)
                        End If
                        aspp(i) = 0
                        aspa(i) = 0
                        asra(i) = 0
                        asrp(i) = 0
                    Next
            End Select
        End With
    End Sub

    Protected Sub WriteMKP()
        With sheet2
            For Each oRow In oTable.Rows
                intPos = oRow("ATMONTH")
                If curMKG <> Trim(oRow("HIMKGP")) Then
                    Call InputSumDetail(sheet2, "AVGDetail")
                    dsumspa = 0
                    dsumspp = 0
                    dsumsra = 0
                    dsumsrp = 0
                    intCurrow = intNext
                    Workbook.Worksheets(0).Range("A8:S10").CopyTo(.Range("A" & intCurrow))
                    .Range(intCurrow, 1).Value = Trim(oRow("HIMKGP"))
                    .Range(intCurrow + 1, 1).Value = Trim(oRow("CNAME"))
                    curMKG = Trim(oRow("HIMKGP"))
                    intNext = intNext + 3
                End If

                If oRow("type") = "SP" Then
                    aspa(intPos) = aspa(intPos) + CDbl(oRow("saleamount"))
                    aspp(intPos) = aspp(intPos) + CDbl(oRow("salecost"))
                    dsumspa = dsumspa + CDbl(oRow("saleamount"))
                    dsumspp = dsumspp + CDbl(oRow("salecost"))
                    .Range(intCurrow, ampos(intPos)).Value = oRow("saleamount")
                    .Range(intCurrow + 1, ampos(intPos)).Value = CDbl(oRow("salecost"))
                    .Range(intCurrow + 2, ampos(intPos)).Value = CDbl(oRow("salecost")) * 100 / CDbl(oRow("saleamount"))
                Else
                    .Range(intCurrow, ampos(intPos) + 1).Value = CDbl(oRow("saleamount")) / 1000
                    .Range(intCurrow + 1, ampos(intPos) + 1).Value = (CDbl(oRow("saleamount")) / 1000) - CDbl(oRow("salecost")) / 1000
                    .Range(intCurrow + 2, ampos(intPos) + 1).Value = ((CDbl(oRow("saleamount")) / 1000) - (CDbl(oRow("salecost"))) / 1000) * 100 / (CDbl(oRow("saleamount")) / 1000)
                    asra(intPos) = asra(intPos) + (CDbl(oRow("saleamount")) / 1000)
                    asrp(intPos) = asrp(intPos) + (CDbl(oRow("salecost")) / 1000)
                    dsumsra = dsumsra + (CDbl(oRow("saleamount")) / 1000)
                    dsumsrp = dsumsrp + .Range(intCurrow + 1, ampos(intPos) + 1).Value
                End If
            Next
            Call InputSumDetail(sheet2, "AVGDetail")
            Call inputSumTotal(sheet2, "3")
            Call InputSumDetail(sheet2, "AVGGrand")
            dsumspa = 0
            dsumspp = 0
            dsumsra = 0
            dsumsrp = 0
        End With
    End Sub
    Protected Sub WritePDTP(ByRef workbook As IWorkbook, ByRef sheet2 As IWorksheet)
        For Each oRow In oTable.Rows
            intPos = oRow("ATMONTH")
            If curPDTP <> Trim(oRow("DIPDTP")) Or curKGP <> Trim(oRow("HIMKGP")) Then
                Call InputSumDetail(sheet2, "AVGDetail")
                dsumspa = 0
                dsumspp = 0
                dsumsra = 0
                dsumsrp = 0
                Call inputSumTotal(sheet2, "1")
                Call InputSumDetail(sheet2, "AVGGrand")
                intNext = intNext + 3
                sheet2.Range(intCurrow, 1).Value = curPTNM
                sheet2.Range(intCurrow + 1, 1).Value = "TOTAL"
                curPDTP = Trim(oRow("DIPDTP"))
                curPTNM = IIf(oRow("PTname") Is DBNull.Value, "", oRow("PTname"))
                curPDNM = IIf(oRow("pdnm") Is DBNull.Value, "", oRow("pdnm"))
                intCurrow = intNext
            End If

            If curGPN <> Trim(oRow("pdgpn")) Or curKGP <> Trim(oRow("himkgp")) Then
                workbook.Worksheets(0).Range("A8:S10").CopyTo(sheet2.Range("A" & intCurrow))
                Call inputSumTotal(sheet2, "2")
                Call InputSumDetail(sheet2, "AVGGrand")
                intNext = intNext + 3
                intCurrow = intNext
                curGPN = Trim(oRow("pdgpn"))
            End If

            If curPDCD <> Trim(oRow("DIPDCD")) And curKGP = Trim(oRow("himkgp")) Then
                Call InputSumDetail(sheet2, "AVGDetail")
                dsumspa = 0
                dsumspp = 0
                dsumsra = 0
                dsumsrp = 0
                intCurrow = intNext
                workbook.Worksheets(0).Range("A8:S10").CopyTo(sheet2.Range("A" & intCurrow))
                sheet2.Range(intCurrow, 1).Value = "'" & Trim(oRow("dipdcd"))
                sheet2.Range(intCurrow + 1, 1).Value = Trim(oRow("pdnm"))
                curPDCD = Trim(oRow("DIPDCD"))
                intNext = intNext + 3
            End If
            If curKGP <> Trim(oRow("himkgp")) Then
                Call inputSumTotal(sheet2, "3")
                Call InputSumDetail(sheet2, "AVGGrand")
                intCurrow = 8
                intNext = 11
                For i As Integer = 0 To 12
                    aptpp(i) = 0.0
                    aptpa(i) = 0.0
                    aptrp(i) = 0.0
                    aptra(i) = 0.0
                    apgpa(i) = 0.0
                    apgpp(i) = 0.0
                    apgra(i) = 0.0
                    apgrp(i) = 0.0
                Next
                sheet2 = workbook.Worksheets.Create(Trim(oRow("himkgp")))
                Call WriteHMonth()
                sheet2.Range(intCurrow, 1).Value = "'" & Trim(oRow("dipdcd"))
                sheet2.Range(intCurrow + 1, 1).Value = oRow("pdnm")
                curKGP = Trim(oRow("himkgp"))
                curPDCD = Trim(oRow("DIPDCD"))
                curPDTP = Trim(oRow("DIPDTP"))
            End If
            With sheet2
                If oRow("type") = "SP" Then
                    apgpa(intPos) = apgpa(intPos) + CDbl(oRow("saleamount"))
                    apgpp(intPos) = apgpp(intPos) + CDbl(oRow("salecost"))

                    aspa(intPos) = aspa(intPos) + CDbl(oRow("saleamount"))
                    aspp(intPos) = aspp(intPos) + CDbl(oRow("salecost"))

                    aptpa(intPos) = aptpa(intPos) + CDbl(oRow("saleamount"))
                    aptpp(intPos) = aptpp(intPos) + CDbl(oRow("salecost"))

                    dsumspa = dsumspa + CDbl(oRow("saleamount"))
                    dsumspp = dsumspp + CDbl(oRow("salecost"))

                    .Range(intCurrow, ampos(intPos)).Value = oRow("saleamount")
                    .Range(intCurrow + 1, ampos(intPos)).Value = CDbl(oRow("salecost"))
                    .Range(intCurrow + 2, ampos(intPos)).Value = CDbl(oRow("salecost")) * 100 / CDbl(oRow("saleamount"))
                Else
                    .Range(intCurrow, ampos(intPos) + 1).Value = CDbl(oRow("saleamount")) / 1000
                    .Range(intCurrow + 1, ampos(intPos) + 1).Value = (CDbl(oRow("saleamount")) / 1000) - CDbl(oRow("salecost")) / 1000
                    .Range(intCurrow + 2, ampos(intPos) + 1).Formula = "= " & .Range(intCurrow + 1, ampos(intPos) + 1).Value & " *100/" & .Range(intCurrow, ampos(intPos) + 1).Value

                    apgra(intPos) = apgra(intPos) + (CDbl(oRow("saleamount")) / 1000)
                    apgrp(intPos) = apgrp(intPos) + (CDbl(oRow("salecost")) / 1000)

                    asra(intPos) = asra(intPos) + (CDbl(oRow("saleamount") / 1000))
                    asrp(intPos) = asrp(intPos) + (CDbl(oRow("salecost")) / 1000)

                    aptra(intPos) = aptra(intPos) + (CDbl(oRow("saleamount")) / 1000)
                    aptrp(intPos) = aptrp(intPos) + (CDbl(oRow("salecost")) / 1000)

                    dsumsra = dsumsra + (CDbl(oRow("saleamount")) / 1000)
                    dsumsrp = dsumsrp + .Range(intCurrow + 1, ampos(intPos) + 1).Value
                End If
            End With
        Next

        Call InputSumDetail(sheet2, "AVGDetail")
        Call inputSumTotal(sheet2, "1")
        Call InputSumDetail(sheet2, "AVGGrand")
        intNext = intNext + 3
        sheet2.Range(intCurrow, 1).Value = curPTNM
        sheet2.Range(intCurrow + 1, 1).Value = "TOTAL"
        intCurrow = intNext

        workbook.Worksheets(0).Range("A8:S10").CopyTo(sheet2.Range("A" & intCurrow))
        Call inputSumTotal(sheet2, "2")
        Call InputSumDetail(sheet2, "AVGGrand")
        intNext = intNext + 3
        intCurrow = intNext
        curGPN = Trim(oRow("pdgpn"))

        Call InputSumDetail(sheet2, "AVGDetail")
        Call inputSumTotal(sheet2, "3")
        Call InputSumDetail(sheet2, "AVGGrand")
        dsumspa = 0
        dsumspp = 0
        dsumsra = 0
        dsumsrp = 0
    End Sub

    Protected Sub WritePDGP(ByRef workbook As IWorkbook, ByRef sheet2 As IWorksheet)
        For Each oRow In oTable.Rows
            intPos = oRow("ATMONTH")

            If curPDTP <> Trim(oRow("DIPDTP")) Then
                Call InputSumDetail(sheet2, "AVGDetail")
                dsumspa = 0
                dsumspp = 0
                dsumsra = 0
                dsumsrp = 0
                Call inputSumTotal(sheet2, "1")
                Call InputSumDetail(sheet2, "AVGGrand")
                intNext = intNext + 3
                sheet2.Range(intCurrow, 1).Value = curPTNM
                sheet2.Range(intCurrow + 1, 1).Value = "TOTAL"
                curPDTP = Trim(oRow("DIPDTP"))
                curPTNM = Trim(oRow("PTname"))
                curPDNM = Trim(oRow("pdnm"))
                intCurrow = intNext
            End If
            If oRow("PDGPN") IsNot DBNull.Value Then
                If curGPN <> Trim(oRow("pdgpn")) Then
                    workbook.Worksheets(0).Range("A8:S10").CopyTo(sheet2.Range("A" & intCurrow))
                    Call inputSumTotal(sheet2, "2")
                    Call InputSumDetail(sheet2, "AVGGrand")
                    intNext = intNext + 3
                    intCurrow = intNext
                    curGPN = Trim(oRow("pdgpn"))
                End If
            End If

            If curPDCD <> Trim(oRow("DIPDCD")) Then
                Call InputSumDetail(sheet2, "AVGDetail")
                dsumspa = 0
                dsumspp = 0
                dsumsra = 0
                dsumsrp = 0
                intCurrow = intNext
                workbook.Worksheets(0).Range("A8:S10").CopyTo(sheet2.Range("A" & intCurrow))
                If oRow("dipdcd") IsNot DBNull.Value Then
                    sheet2.Range(intCurrow, 1).Value = "'" & Trim(oRow("dipdcd"))
                End If
                If oRow("PDNM") IsNot DBNull.Value Then
                    sheet2.Range(intCurrow + 1, 1).Value = Trim(oRow("pdnm"))
                End If
                curPDCD = Trim(oRow("DIPDCD"))
                intNext = intNext + 3
            End If

            With sheet2
                If oRow("type") = "SP" Then
                    apgpa(intPos) = apgpa(intPos) + CDbl(oRow("saleamount"))
                    apgpp(intPos) = apgpp(intPos) + CDbl(oRow("salecost"))

                    aspa(intPos) = aspa(intPos) + CDbl(oRow("saleamount"))
                    aspp(intPos) = aspp(intPos) + CDbl(oRow("salecost"))

                    aptpa(intPos) = aptpa(intPos) + CDbl(oRow("saleamount"))
                    aptpp(intPos) = aptpp(intPos) + CDbl(oRow("salecost"))

                    dsumspa = dsumspa + CDbl(oRow("saleamount"))
                    dsumspp = dsumspp + CDbl(oRow("salecost"))

                    .Range(intCurrow, ampos(intPos)).Value = oRow("saleamount")
                    .Range(intCurrow + 1, ampos(intPos)).Value = CDbl(oRow("salecost"))
                    .Range(intCurrow + 2, ampos(intPos)).Value = CDbl(oRow("salecost")) * 100 / CDbl(oRow("saleamount"))
                Else
                    .Range(intCurrow, ampos(intPos) + 1).Value = CDbl(oRow("saleamount")) / 1000
                    .Range(intCurrow + 1, ampos(intPos) + 1).Value = (CDbl(oRow("saleamount")) / 1000) - CDbl(oRow("salecost")) / 1000
                    .Range(intCurrow + 2, ampos(intPos) + 1).Formula = "= " & .Range(intCurrow + 1, ampos(intPos) + 1).Value & " *100/" & .Range(intCurrow, ampos(intPos) + 1).Value

                    apgra(intPos) = apgra(intPos) + (CDbl(oRow("saleamount")) / 1000)
                    apgrp(intPos) = apgrp(intPos) + (CDbl(oRow("salecost")) / 1000)

                    asra(intPos) = asra(intPos) + (CDbl(oRow("saleamount") / 1000))
                    asrp(intPos) = asrp(intPos) + (CDbl(oRow("salecost")) / 1000)

                    aptra(intPos) = aptra(intPos) + (CDbl(oRow("saleamount")) / 1000)
                    aptrp(intPos) = aptrp(intPos) + (CDbl(oRow("salecost")) / 1000)

                    dsumsra = dsumsra + (CDbl(oRow("saleamount")) / 1000)
                    dsumsrp = dsumsrp + .Range(intCurrow + 1, ampos(intPos) + 1).Value
                End If
            End With
        Next

        Call InputSumDetail(sheet2, "AVGDetail")
        intCurrow = intNext

        Workbook.Worksheets(0).Range("A8:S10").CopyTo(sheet2.Range("A" & intCurrow))
        Call inputSumTotal(sheet2, "2")
        Call InputSumDetail(sheet2, "AVGGrand")
        intNext = intNext + 3
        intCurrow = intNext
        curGPN = Trim(oRow("pdgpn"))

        Call InputSumDetail(sheet2, "AVGDetail")
        Call inputSumTotal(sheet2, "3")
        Call InputSumDetail(sheet2, "AVGGrand")
        dsumspa = 0
        dsumspp = 0
        dsumsra = 0
        dsumsrp = 0
    End Sub

    Protected Sub SetPageProperties()
        With sheet2
            .SetColumnWidth(1, 12.71)                       'COLUMN A  -- MAKER 
            .SetColumnWidth(2, 2)                           'COLUMN B  -- SA, PR, %
            For i As Integer = 3 To 18
                '.SetColumnWidth(i, 11.5)                    'COLUMN C  -- PLAN & RESULT 
                .SetColumnWidth(i, 9.86)
            Next
            '.SetColumnWidth(18, 8.86)                       'COLUMN R  -- %
            .SetColumnWidth(18, 8.86)
            .SetRowHeight(1, 25.5)
            .SetRowHeight(2, 21.5)
            .SetRowHeight(3, 21.5)
            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Landscape
            .PageSetup.LeftMargin = 0.25
            .PageSetup.RightMargin = 0.25
            .PageSetup.TopMargin = 0.25
            .PageSetup.BottomMargin = 0.25
            .PageSetup.Zoom = 60
            .PageSetup.PrintTitleRows = "$1:$7"
            .Range("A8").FreezePanes()
        End With
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Call SetPlant()
        End If
    End Sub
End Class
