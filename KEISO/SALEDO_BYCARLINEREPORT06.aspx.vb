Imports System.Data
Imports System.Data.OleDb
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine

Partial Class SALEDO_BYCARLINEREPORT06
    Inherits System.Web.UI.Page
    Private Conn As New OleDbConnection(Classconn.strConnSql)
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("R6-REPORT.xlt")
    Private ReportName As String = "EXPORT-REPORT6"
    Private crtDate As String = Val(Date.Now.ToString("ddMMyyyy"))
    Private crtTime As String = Val(Date.Now.ToString("HHmm"))
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private oTable As New DataTable
    Private oRow As DataRow
    Private Filename As String = ""
    Private strKD, strCL4, strCL5 As String
    Private curPType, curHISTCd, curCL4, curCl5 As String
    Private curMKsf, curPROD, currCus, curdipdgp, curpgname As String
    Private intPNo, intPGno, intPdcdNo, intCurrRow As Integer

    Protected Sub SetPlant()
        Dim oTable As New DataTable
        Dim oRow As DataRow
        oTable = ClassFunctionVar.GetDataUser(Session("USERID"), "MTAPDADR06")
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
                If CInt(strdate) <= 20100403 Then
                    strsql += "where hicocd in ('A','E') "
                End If
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
        Dim Condition As String = ""
        If Conn.State = ConnectionState.Open Then
            Conn.Close()
        End If
        Conn.Open()

        If txtCarMaker.Text <> "" Then
            Condition = "(HIMKGP = '" & Trim(txtCarMaker.Text) & "') AND "
        End If
        If (ddlCompany.SelectedValue <> "ALL") And (ddlCompany.SelectedValue <> "") Then
            Condition = "(HICOCD = '" & ddlCompany.SelectedValue & "') AND "
        End If
        Condition += "(HIBLDT >= " & Right(txtFrMonth.Text, 4) & Mid(txtFrMonth.Text, 3, 2) & Left(txtFrMonth.Text, 2) & " ) "
        Condition += "AND (HIBLDT<= " & Right(txtToMonth.Text, 4) & Mid(txtToMonth.Text, 3, 2) & Left(txtToMonth.Text, 2) & ") "
        Condition += "AND (HICAFG <>'Y') AND (HIIVFG = 'D') AND DISQTY <> 0 AND CTL1.CTYPE = 'MGP' "
        If ddlBillto.SelectedValue <> "0" Then
            Condition += "AND HIRPFG ='" & ddlBillto.SelectedValue & "' "
        End If
        strSql = "SELECT HI = CASE WHEN LEFT(HISTCD,2)='7' THEN 'SPARE PART' WHEN LEFT(HISTCD,2)='8' THEN 'KD' ELSE 'CBU' End, "
        strSql += "HIMKGP AS Cus, HIDTTP, CL5 = CASE WHEN LEFT(HISTCD,2)<>'7'   THEN LEFT(DICALN,5) ELSE ' ' END, LTRIM(CTL.CNAME) AS CNAME, DIPDGP,  "
        strSql += "PG.SHORTNAME AS PGNAME,HIREAS, DIPDCD, DIPROD, DISPUN, DIMKSF, DIFCUN, DIPDTP, "
        strSql += "(ISNULL(SUM((CASE WHEN H.HIDTTP = 'I' THEN DIFTAM*HIFIXR  ELSE 0 END) "
        strSql += "+(CASE WHEN H.HIDTTP = 'D' THEN DIFTAM*HIFIXR ELSE 0 END) "
        strSql += "-(CASE WHEN H.HIDTTP = 'C' THEN DIFTAM*HIFIXR ELSE 0 END)),0)) AS DISPAM, "
        strSql += "(ISNULL(SUM ((CASE WHEN H.HIDTTP = 'I' THEN D.DIFCAM ELSE 0 END)"
        strSql += "- (CASE WHEN H.HIDTTP = 'C' THEN D.DIFCAM ELSE 0 END)"
        strSql += "+ (CASE WHEN H.HIDTTP = 'D' THEN D.DIFCAM ELSE 0 END)),0)) AS FC,"
        strSql += "SUM(DISQTY)AS DISQTY, SUM(CASE WHEN  SUBSTRING(DICALN,6,1) <> '1'   "
        strSql += "OR NOT HIREAS IN('3','2','') THEN 0 ELSE DISQTY END)AS CQTY, "
        strSql += "SUM(ROUND(DISQTY*D.DIMANH,2)) AS DIMANHT, DICOPP , SUM(ROUND(DISQTY*D.DICOPP,2)) AS UCOPP, "
        strSql += "SUM(DISPAM) - SUM(DIFCAM) AS PROFIT, DIMANH, ISNULL(CTL1.CNAME,'-') AS CUSNAME "
        strSql += "FROM  YSSH H "
        strSql += "INNER JOIN YSSD D ON HIDONO = DIDONO AND HICOCD = DICOCD "
        strSql += "LEFT OUTER JOIN  YSCTL CTL ON LEFT(D.DICALN,4) = CKEY   AND CTL.CKEY <> '' "
        strSql += "INNER JOIN PRODUCTGROUP PG ON PRODUCTGROUPCODE = DIPDGP "
        strSql += "LEFT OUTER JOIN  YSCTL  CTL1 ON HIMKGP = CTL1.CKEY   AND CTL1.CKEY <> '' "
        strSql += "WHERE " & Condition
        strSql += "GROUP BY  LEFT(HISTCD,2), HIMKGP, DIPDCD, DIMKSF, DIPROD, DIMANH, "
        strSql += "DISPUN, DIFCUN, DIPDTP, DIPDGP, CTL.CNAME, PG.SHORTNAME, "
        strSql += "HIDTTP, HIREAS, DIFCUN, DIMKSF ,CTL1.CNAME, LEFT(DICALN,5), DICOPP,DICALN  "
        strSql += "ORDER BY HIMKGP DESC, HI DESC, LEFT(DICALN,5), DIPDGP, DIPDCD, DIPROD, DIMKSF, CTL.CNAME, PG.SHORTNAME "
'response.write (strsql)
'response.END
        Dim DS As New DataSet
        Dim DA As New OleDbDataAdapter(strSql, Conn)
        DA.SelectCommand.CommandTimeout = 1000
        DA.Fill(DS, "SalebyCARL")
        Return DS.Tables("SalebyCARL")
    End Function

    Protected Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If txtFrMonth.Text = "" Or txtToMonth.Text = "" Then
            lblmessage.Text = "Please input date"
            Exit Sub
        End If
         oTable = GetData()
       If oTable.Rows.Count > 1 Then
            oRow = oTable.Rows(0)
            appExcel.DefaultFilePath = Server.MapPath(".")
            Workbook = appExcel.Workbooks.Open(TemplateFile)
            sheet = Workbook.Worksheets(0)
            sheet2 = Workbook.Worksheets.Create(oRow("Cus"))
            Call WriteHMonth()
            Call WriteDetail()
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            sheet.Remove()
            Filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
            CommonFunction.SaveTmpFile(CurTempPath, Filename)
            Workbook.SaveAs(Filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
        Else
            lblmessage.Text = "No Data"
            Exit Sub
        End If
    End Sub

    Protected Sub WriteHMonth()
        Call SetPageProperties()
        With sheet2
            Workbook.Worksheets(0).Range("A1:Q9").CopyTo(.Range("A1"))
            .Range(2, 2).Value = Left(txtFrMonth.Text, 2) & "/" & Mid(txtFrMonth.Text, 3, 2) & "/" & Right(txtFrMonth.Text, 4) & _
                                 " - " & Left(txtToMonth.Text, 2) & "/" & Mid(txtToMonth.Text, 3, 2) & "/" & Right(txtToMonth.Text, 4)
            If oRow("CUSNAME") IsNot DBNull.Value Then
                .Range(5, 3).Value = oRow("CUSNAME")
            End If

			If ddlCompany.SelectedValue = "E" Then
                .Range(4, 4).Value = "YIC Asia Pacific Corporation Ltd."
            Else
                .Range(4, 4).Value = "Thai Arrow Products Co.,Ltd."
            End If
		.Range(3, 2).Value = "DOMESTIC"
            .Range(4, 12).Value = "Printdate  :  " & Now()
            .Range(5, 12).Value = "Print by     :  " & Session("userId") & " - " & Session("username")
        End With
    End Sub
    Protected Sub DefaultValue()
        intPNo = 0
        intPGno = 0
        intPdcdNo = 0
        intCurrRow = 9

        curMKsf = ""
        currCus = oRow("cus")
        curCL4 = Left(oRow("cl5"), 4)
        curCl5 = oRow("cl5")
        curPROD = oRow("diprod")
        curdipdgp = oRow("dipdgp")
        curpgname = oRow("pgname")
        curPType = oRow("dipdcd")
        curHISTCd = oRow("hi")

        sheet2.Range(4, 3).Value = "::   " & oRow("Cus")
        sheet2.Range(5, 3).Value = "::   " & oRow("cusname")
        sheet2.Range(6, 2).Value = oRow("hi")

        strKD = "8"
        strCL4 = "8"
        strCL5 = "8"
    End Sub

    Protected Sub WriteDetail()
        Call DefaultValue()
        For Each oRow In oTable.Rows
            If Trim(currCus) <> Trim(oRow("Cus")) Then
                Call COPYLINE("  TOTAL ", curPType, "10")
                Call SumPg(intCurrRow, intPdcdNo)
                intCurrRow += 1
                intPNo = intPNo + 1
                intPGno = intPGno + 1

                Call COPYLINE("  TOTAL ", curpgname, "10")
                Call SumPg(intCurrRow, intPGno)
                intCurrRow += 1
                intPNo = intPNo + 1

                sheet.Range("A10:Q10").CopyTo(sheet2.Range("A" & intCurrRow))
                strCL5 = strCL5 & "|" & intCurrRow
                sheet2.Range(intCurrRow, 2).Value = LTrim(curCl5 & "   TOTAL ")
                Call SumCL5(intCurrRow, intPNo)
                intCurrRow += 1

                sheet.Range("A10:Q10").CopyTo(sheet2.Range("A" & intCurrRow))
                strCL4 = strCL4 & "|" & intCurrRow
                sheet2.Range(intCurrRow, 2).Value = LTrim(curCL4 & "  TOTAL ")
                intCurrRow += 1

                Call COPYLINE("  TOTAL ", LTrim(curHISTCd), "10")
                strKD = strKD & "|" & intCurrRow
                intCurrRow += 1

                sheet.Range("A10:Q10").CopyTo(sheet2.Range("A" & intCurrRow))
                Call COPYLINE(currCus, "GRAND  TOTAL  :  ", "10")
                Call PrintTotal(intCurrRow)

                sheet2 = Workbook.Worksheets.Create(oRow("Cus"))
                Call WriteHMonth()
                Call DefaultValue()
            End If

            If curPType <> oRow("dipdcd") Or curdipdgp <> orow("dipdgp") Or curCl5 <> orow("cl5") Or  curCL4 <> Left(orow("cl5"), 4) Or curHISTCd <> orow("hi") Then 'Or curdipdgp <> rs("dipdgp") Or curCl5 <> rs("cl5") Or curCL4 <> Left(rs("cl5").value, 4) Or curHISTCd <> rs("hi").value 
                Call COPYLINE("  TOTAL  ", curPType, "10")
                Call SumPg(intCurrRow, intPdcdNo)
                intCurrRow += 1
                intPNo = intPNo + 1
                intPGno = intPGno + 1
                intPdcdNo = 0
            End If

            If curdipdgp <> oRow("dipdgp") Or curCl5 <> oRow("cl5") Or curCL4 <> Left(oRow("cl5"), 4) Or curHISTCd <> oRow("hi") Then 'Or curCl5 <> rs("cl5") Or curCL4 <> Left(rs("cl5").value, 4) Or curHISTCd <> rs("hi").value
                Call COPYLINE("  TOTAL ", curpgname, "10")
                Call SumPg(intCurrRow, intPGno)
                curdipdgp = oRow("dipdgp")
                curpgname = oRow("pgname")
                intCurrRow += 1
                intPNo = intPNo + 1
                intPGno = 0
            End If

            If curCl5 <> oRow("cl5") Or curCL4 <> Left(oRow("cl5"), 4) Or curHISTCd <> oRow("hi") Then 'Or curCL4 <> Left(rs("cl5").value, 4) Or curHISTCd <> rs("hi").value 
                Call COPYLINE("  TOTAL ", LTrim(curCl5), "10")
                strCL5 = strCL5 & "|" & intCurrRow
                Call SumCL5(intCurrRow, intPNo)
                intCurrRow += 1
                curPType = ""
                intPNo = 0
            End If

            If curCL4 <> Left(oRow("cl5"), 4) Or curHISTCd <> oRow("hi") Then 'Or curHISTCd <> rs("hi").value
                Call COPYLINE("  TOTAL ", LTrim(curCL4), "10")
                strCL4 = strCL4 & "|" & intCurrRow
                curCL4 = Left(oRow("cl5"), 4)
                intCurrRow += 1
            End If

            If curHISTCd <> oRow("hi") Or Trim(currCus) <> Trim(oRow("Cus")) Then ' Or Trim(currCus) <> Trim(rs("Cus").value)
                Call COPYLINE("  TOTAL ", LTrim(curHISTCd), "10")
                strKD = strKD & "|" & intCurrRow
                intCurrRow += 3
                sheet.Range("A6:Q8").CopyTo(sheet2.Range("A" & intCurrRow))
                sheet2.Range(intCurrRow, 2).Value = oRow("hi")
                intCurrRow += 3
                curHISTCd = oRow("hi")
            End If

            sheet.Range("A9:Q9").CopyTo(sheet2.Range("A" & intCurrRow))

            If curCl5 <> oRow("cl5") Then
                If oRow("cname") IsNot DBNull.Value And oRow("cl5") IsNot DBNull.Value Then
                    sheet2.Range(intCurrRow, 2).Value = Trim(oRow("cname")) & " - " & Trim(oRow("cl5"))
                ElseIf oRow("cname") IsNot DBNull.Value Then
                    sheet2.Range(intCurrRow, 2).Value = Trim(oRow("cname"))
                ElseIf oRow("cl5") IsNot DBNull.Value Then
                    sheet2.Range(intCurrRow, 2).Value = Trim(oRow("cl5"))
                End If
                curCl5 = oRow("Cl5")
            End If

            If curPType <> oRow("dipdcd") Or intPdcdNo = 0 Then
                sheet2.Range(intCurrRow, 3).Text = oRow("dipdcd")
                If oRow("cname") IsNot DBNull.Value And oRow("cl5") IsNot DBNull.Value Then
                    sheet2.Range(intCurrRow, 2).Value = Trim(oRow("cname")) & " - " & Trim(oRow("cl5"))
                ElseIf oRow("cname") IsNot DBNull.Value Then
                    sheet2.Range(intCurrRow, 2).Value = Trim(oRow("cname"))
                ElseIf oRow("cl5") IsNot DBNull.Value Then
                    sheet2.Range(intCurrRow, 2).Value = Trim(oRow("cl5"))
                End If
                curPType = oRow("dipdcd")
            End If
            With sheet2
                .Range(intCurrRow, 4).Text = Trim(oRow("diprod"))
                .Range(intCurrRow, 5).Text = oRow("diMKSF")
                curMKsf = oRow("diMKsf")
                curPROD = oRow("diprod")
				if curHISTCd <>"SPARE PART" then 
					.Range(intCurrRow, 6).Value = CInt(oRow("Cqty"))
				else 
				.Range(intCurrRow, 6).Value = 0
				end if
                .Range(intCurrRow, 7).Value = CInt(oRow("disqty"))
                .Range(intCurrRow, 8).Value = oRow("dispun")
                .Range(intCurrRow, 9).Value = oRow("dispam")
                .Range(intCurrRow, 10).Value = oRow("difcun")
                .Range(intCurrRow, 11).Value = oRow("fc")
                .Range(intCurrRow, 12).Value = .Range(intCurrRow, 9).Value - .Range(intCurrRow, 11).Value
                If .Range(intCurrRow, 9).Value <> 0 Then
                    .Range(intCurrRow, 13).Value = (CDbl(.Range(intCurrRow, 12).Value) * 100) / CDbl(.Range(intCurrRow, 9).Value)
                Else
                    .Range(intCurrRow, 13).Value = 0
                End If
                .Range(intCurrRow, 14).Value = oRow("dimanh")
                .Range(intCurrRow, 15).Value = oRow("dimanhT")
                .Range(intCurrRow, 15).Value = .Range(intCurrRow, 14).Value * .Range(intCurrRow, 7).Value
                .Range(intCurrRow, 16).Value = oRow("diCOPP")
                .Range(intCurrRow, 17).Value = oRow("UCOPP")
            End With
            intCurrRow += 1
            intPNo = intPNo + 1
            intPGno = intPGno + 1
            intPdcdNo = intPdcdNo + 1
        Next
        Call COPYLINE("  TOTAL ", curPType, "10")
        Call SumPg(intCurrRow, intPdcdNo)
        intCurrRow += 1
        intPNo = intPNo + 1
        intPGno = intPGno + 1

        Call COPYLINE("  TOTAL ", curpgname, "10")
        Call SumPg(intCurrRow, intPGno)
        intCurrRow += 1
        intPNo = intPNo + 1

        sheet.Range("A10:Q10").CopyTo(sheet2.Range("A" & intCurrRow))
        strCL5 = strCL5 & "|" & intCurrRow
        sheet2.Range(intCurrRow, 2).Value = LTrim(curCl5 & "   TOTAL ")
        Call SumCL5(intCurrRow, intPNo)
        intCurrRow += 1

        sheet.Range("A10:Q10").CopyTo(sheet2.Range("A" & intCurrRow))
        strCL4 = strCL4 & "|" & intCurrRow
        sheet2.Range(intCurrRow, 2).Value = LTrim(curCL4 & "  TOTAL ")
        intCurrRow += 1

        Call COPYLINE("  TOTAL ", LTrim(curHISTCd), "10")
        strKD = strKD & "|" & intCurrRow
        intCurrRow += 1

        sheet.Range("A10:Q10").CopyTo(sheet2.Range("A" & intCurrRow))
        Call COPYLINE(currCus, "GRAND  TOTAL  :  ", "10")
        Call PrintTotal(intCurrRow)
    End Sub
    Protected Sub COPYLINE(ByVal strValue As String, ByVal strTotal As String, ByVal intRow As Integer)
        sheet.Range("A" & intRow & ":Q" & intRow).CopyTo(sheet2.Range("A" & intCurrrow))
        sheet2.Range(intCurrRow, 2).Value = strTotal & strValue
    End Sub

    Protected Sub SumPg(ByVal intRow, ByVal CL5No)
        With sheet2
            .Range(intRow, 6).FormulaR1C1 = "=FIXeD( SUM(R[-" & CL5No & "]C:R[-1]C),0)"
            .Range(intRow, 7).FormulaR1C1 = "=FIXeD( SUM(R[-" & CL5No & "]C:R[-1]C),0)"
            .Range(intRow, 9).FormulaR1C1 = "=FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 11).FormulaR1C1 = "=FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 12).FormulaR1C1 = "=FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 15).FormulaR1C1 = "=FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 16).FormulaR1C1 = "=FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 17).FormulaR1C1 = "=FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
        End With
    End Sub

    Protected Sub SumCL5(ByVal intRow, ByVal CL5No)
        With sheet2
            .Range(intRow, 6).FormulaR1C1 = "= FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 7).FormulaR1C1 = "= FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 9).FormulaR1C1 = "= FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 11).FormulaR1C1 = "= FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 12).FormulaR1C1 = "= FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 15).FormulaR1C1 = "= FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 16).FormulaR1C1 = "= FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
            .Range(intRow, 17).FormulaR1C1 = "= FIXeD(SUM(R[-" & CL5No & "]C:R[-1]C))"
        End With
    End Sub

    Protected Sub PrintTotal(ByVal intTR)
        Dim cl4 As Integer = 1
        Dim kd As Integer = 1
        Dim aCL4() As String = Split(strCL4, "|")
        Dim aKD() As String = Split(strKD, "|")
        Dim aCL5() As String = Split(strCL5, "|")
        With sheet2
            For i As Integer = 1 To UBound(aCL5)
                If CInt(aCL5(i)) >= CInt(aCL4(cl4)) Then
                    cl4 = cl4 + 1
                End If

                If CInt(aCL5(i)) >= CInt(aKD(kd)) Then
                    kd = kd + 1
                End If

                For iCol As Integer = 6 To 17
                    If Not iCol = 8 And Not iCol = 10 And Not iCol = 13 And Not iCol = 14 Then
                        .Range(CInt(aCL4(cl4)), iCol).Formula += "+ " & Mid(.Range(CInt(aCL5(i)), iCol).Formula, 2, .Range(CInt(aCL5(i)), iCol).Formula.Length)
                        .Range(CInt(aKD(kd)), iCol).Formula += "+ " & Mid(.Range(CInt(aCL5(i)), iCol).Formula, 2, .Range(CInt(aCL5(i)), iCol).Formula.Length)
                        .Range(intTR, iCol).Formula += "+ " & Mid(.Range(CInt(aCL5(i)), iCol).Formula, 2, .Range(CInt(aCL5(i)), iCol).Formula.Length)
                    End If
                Next
            Next
        End With
    End Sub

    Protected Sub SetPageProperties()
        With sheet2
            .SetColumnWidth(1, 0.5)                     'COLUMN A  -- 
            .SetColumnWidth(2, 18.29)                   'COLUMN B  -- CARLINE CODE
            .SetColumnWidth(3, 9)                       'COLUMN C  -- PRODUCT CODE
            .SetColumnWidth(4, 15.86)                   'COLUMN D  -- PART NO
            .SetColumnWidth(5, 5.86)                    'COLUMN E  -- SUFFIX 
            .SetColumnWidth(6, 7.86)                    'COLUMN F  -- PQTY
            .SetColumnWidth(7, 9.29)                    'COLUMN G  -- QTY 
            .SetColumnWidth(8, 10.57)                   'COLUMN H  -- UNIT PRICE
            .SetColumnWidth(9, 15.43)                   'COLUMN I  -- AMONT
            .SetColumnWidth(10, 9.86)                   'COLUMN J  -- UNIT COST
            .SetColumnWidth(11, 14.86)                  'COLUMN K  -- COST AMOUNT
            .SetColumnWidth(12, 12.71)                  'COLUMN L  -- PROFIT AMOUNT
            .SetColumnWidth(13, 7.6)                    'COLUMN M  -- PROFIT %
            .SetColumnWidth(14, 8.43)                   'COLUMN N  -- UNIT
            .SetColumnWidth(15, 10.5)                   'COLUMN O  -- HOURS AMOUNT
            .SetColumnWidth(16, 11)                     'COLUMN P  -- UNIT
            .SetColumnWidth(17, 14.29)                   'COLUMN Q  -- AMOUNT
            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Landscape
            .PageSetup.LeftMargin = 0.5
            .PageSetup.RightMargin = 0.5
            .PageSetup.TopMargin = 0.5
            .PageSetup.BottomMargin = 0.5
            .PageSetup.Zoom = 70
            .PageSetup.PrintTitleRows = "$1:$8"
            .Range("A9").FreezePanes()
        End With
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Call SetPlant()
        End If
    End Sub
End Class
