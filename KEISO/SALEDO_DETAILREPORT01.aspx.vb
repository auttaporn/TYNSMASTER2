Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports Syncfusion.XlsIO
Imports System.Windows.Forms
Imports System.Web.UI
Partial Class SALEDO_DETAILREPORT01
    Inherits System.Web.UI.Page
    Private conn As New OleDbConnection(Classconn.strConnSql)
    Private DefaultPath As String = ("~/tmp/")
    Private Rptpath As String = Server.MapPath("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("R1-REPORT.xlt")
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private ReportName As String = "DOMESTIC-REPORT1"
    Private dt As New DataTable
    Private idx As Integer = 0
    Private drow As DataRow
    Private oTable As New DataTable
    Private oRow As DataRow
    Private intFile As Integer
    Private intCurrRow As Integer
    Private intSheet, intNext, intCurr As Integer
    Private curBLDate, curINV, curCus, hibtcd, intSTC As String
    Private iFile As Integer = 1
    Private GrandCol7 As Double = 0.0
    Private GrandCol9 As Double = 0.0
    Private GrandCol11 As Double = 0.0
    Private SubCol7, SubCol8, SubCol9, SubCol10, SubCol11 As Double
    Private FileName As String
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet

    Protected Sub SetPlant()
        Dim oTable As New DataTable
        Dim oRow As DataRow
        oTable = ClassFunctionVar.GetDataUser(Session("USERID"), "MTAPDADR01")
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
    Protected Function Getdata() As DataTable
        Dim condition, Sql As String
        condition = ""
        If (ddlCompany.SelectedValue <> "ALL") And (ddlCompany.SelectedValue <> "") Then
            condition += " AND HICOCD = '" & Trim(ddlCompany.SelectedValue) & "' "
        End If
        If txtCarMaker.Text <> "" Then
            condition += " AND HIBTCD = '" & Trim(txtCarMaker.Text) & "' "
        End If

        If txtFrom.Text <> "" Then
            condition += " AND HIBLDT >= '" & Right(txtFrom.Text, 4) & Mid(txtFrom.Text, 3, 2) & Left(txtFrom.Text, 2) & "' "
        End If

        If txtTo.Text <> "" And Left(txtTo.Text, 2) <= "31" Then
            condition += " AND HIBLDT <= '" & Right(txtTo.Text, 4) & Mid(txtTo.Text, 3, 2) & Left(txtTo.Text, 2) & "' "
        End If

        If ddlFCType.SelectedValue <> "All" Then
            condition += " AND DIFCTP= '" & ddlFCType.SelectedValue & "' "
        End If

        If ddlSalePrice.SelectedValue <> "All" Then
            condition += " AND DISPTP= '" & ddlSalePrice.SelectedValue & "' "
        End If

        If ddlBillto.SelectedValue <> "0" Then
            condition += " AND HIRPFG= '" & ddlBillto.SelectedValue & "' "
        End If

        Sql = "SELECT HIBLDT, HIDONO, DIPROD, DIMKSF, DIFACD,(YSSD.DISPFT * YSSH.HIFIXR) AS DISPUN,"
        Sql = Sql & "DIFCUN, DIPDCD, LEFT(DICALN,5)AS CARLINE, HIREFE, HIBTCD, CNAME, HIFIXR, "
        Sql = Sql & "ROUND(ISNULL(SUM(ROUND(CASE WHEN YSSH.HIDTTP = 'I' THEN YSSD.DISQTY ELSE 0 END- "
        Sql = Sql & "CASE WHEN YSSH.HIDTTP = 'C' THEN YSSD.DISQTY ELSE 0 END + "
        Sql = Sql & "CASE WHEN YSSH.HIDTTP = 'D' THEN YSSD.DISQTY ELSE 0 END,2)),0),2) AS QTYAM, "
        Sql = Sql & "ROUND(ISNULL(SUM(ROUND(CASE WHEN YSSH.HIDTTP = 'I' THEN (YSSD.DIFTAM * YSSH.HIFIXR) ELSE 0 END- "
        Sql = Sql & "CASE WHEN YSSH.HIDTTP = 'C' THEN (YSSD.DIFTAM * YSSH.HIFIXR) ELSE 0 END +CASE "
        Sql = Sql & "WHEN YSSH.HIDTTP = 'D' THEN (YSSD.DIFTAM * YSSH.HIFIXR) ELSE 0 END,2)),0),2) AS SaleAm, "
        Sql = Sql & "ROUND(isnull(SUM (ROUND(CASE WHEN YSSH.HIDTTP = 'I' THEN YSSD.DIFCAM ELSE 0 END- "
        Sql = Sql & "CASE WHEN YSSH.HIDTTP = 'C' THEN YSSD.DIFCAM ELSE 0 END+CASE WHEN YSSH.HIDTTP = 'D' THEN "
        Sql = Sql & "YSSD.DIFCAM ELSE 0 END+CASE WHEN YSSH.HIDTTP = 'A' THEN YSSD.DIFCAM ELSE 0 END,2)),0 ),2) AS FCAM, "
        Sql = Sql & "DIFCTP , DISPTP,DIMANH, DICOPP "
        Sql = Sql & "FROM YSSH "
        Sql = Sql & "INNER JOIN YSSD ON HIDONO = DIDONO "
        Sql = Sql & "INNER JOIN YSCTL ON HIBTCD = CKEY AND CTYPE = 'CUS' "
        Sql = Sql & "WHERE  HIIVFG = 'D'  " & condition & " "
        Sql = Sql & "GROUP BY HIBLDT, HIDONO, DIPROD, DIFACD, HIFIXR, DIFCUN, DIPDCD, "
        Sql = Sql & "LEFT(DICALN,5), HIREFE, HIBTCD, DIMKSF, CNAME, DISPFT, "
        Sql = Sql & "DIFCTP, DISPTP, DIMANH, DICOPP "
        Sql = Sql & "ORDER BY HIBTCD, HIBLDT, HIDONO"
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Dim da As New OleDbDataAdapter(Sql, conn)
        Dim ds As New DataSet
        da.SelectCommand.CommandTimeout = 1000
        da.Fill(ds, "SaleDetail")
        Return ds.Tables("SaleDetail")
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If Session("userid") = "" Then
        '    Response.Redirect("~/")
        'End If

        If Not Page.IsPostBack Then
            txtFrom.Text = ClassFunctionVar.STRDate2
            txtTo.Text = ClassFunctionVar.STRDate2
            Call SetPlant()

            MultiView1.SetActiveView(ViewCondition)
            ViewCondition.DataBind()
        End If
    End Sub

    Protected Sub btnsubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmit.Click
        If (Left(txtTo.Text, 2) - Left(txtFrom.Text, 2)) > 31 Then
            lblmessage.Text = "Message : Date can not over  31 day"
            Exit Sub
        End If
        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)    'default first sheet    
        oTable = Getdata()

        If oTable.Rows.Count < 1 Then
            lblmessage.Text = "Message: No Data"
        Else
            oRow = oTable.Rows(0)
            Call PrintDetail(oTable)
        End If
    End Sub

    Protected Sub PrintH()
        With sheet2
            .Range(intCurrRow, 2).Value = "'" & Trim(oRow("HIBLDT"))
            .Range(intCurrRow, 3).Value = Trim(oRow("DIFACD"))
            .Range(2, 3).Value = Left(txtFrom.Text, 2) & "/" & Mid(txtFrom.Text, 3, 2) & "/" & Right(txtFrom.Text, 4) & _
                                 "  -  " & Left(txtTo.Text, 2) & "/" & Mid(txtTo.Text, 3, 2) & "/" & Right(txtTo.Text, 4)
            .Range(3, 3).Value = "DOMESTIC"
            .Range(4, 3).Value = Trim(oRow("HIBTCD"))
            .Range(5, 3).Value = Trim(oRow("CNAME"))
            .Range(5, 7).Value = "Fix Rate : " & Trim(oRow("HIFIXR"))
            .Range(4, 20).Value = "Printdate  : " & Now()
            .Range(5, 20).Value = "Print by     : " & Session("userId") & "  -  " & Session("username")
        End With
    End Sub
    Protected Sub PrintDetail(ByVal oTable As DataTable)
        intCurrRow = 8
        intNext = 8
        intFile = 1
        Dim sumhibtcd As Integer = 0
        Dim sumcurINV As Integer = 0
        curINV = Trim(oRow("HIDONO"))
        curBLDate = Trim(oRow("HIBLDT"))
        hibtcd = Trim(oRow("HIBTCD"))

        sheet2 = appExcel.Worksheets.Create(Trim(oRow("HIBTCD")))
        Call SetPageProperties(sheet2)
        sheet.Range("A1:V8").CopyTo(sheet2.Range("A1"))
        Call PrintH()
        For Each oRow In oTable.Rows
            If hibtcd <> Trim(oRow("HIBTCD")) Or (curINV <> Trim(oRow("HIDONO"))) Then  '<> Customer Code , Invoice No
                If curINV <> Trim(oRow("HIDONO")) Then   'New Invoice No / New Line
                    Call CheckInvNo(sumcurINV)
                End If
                If hibtcd <> Trim(oRow("HIBTCD")) Then   '1 Customer Code / 1 Sheet
                    Call CheckCustomerCode(sumhibtcd, hibtcd) ', rowStyle
                End If
            End If
            Call DataDetail(sumcurINV, sumhibtcd)
        Next
        Call CheckInvNo(sumcurINV)
        Call GrandTotal2(sumhibtcd)
        Call SaveFile(iFile)
        MultiView1.SetActiveView(ViewOpen)
        ViewOpen.DataBind()
        GridView2.DataSource = dt
        GridView2.DataBind()
        If GridView2.Rows.Count = 0 Then
            lblmessage.Text = "Message : No Data"
            MultiView1.SetActiveView(ViewCondition)
        End If
    End Sub

    Protected Sub CheckInvNo(ByRef sumcurINV As Integer)
        If sumcurINV > 0 Then
            Call SubTotal(sheet, sheet2, intCurrRow)
            intSTC = intSTC & intCurrRow & "|"
            intNext = intNext + 1
            intCurrRow = intNext
            sheet.Range("A8:V8").CopyTo(sheet2.Range("A" & intNext))
            sumcurINV = 0
        End If
    End Sub
    Protected Sub CheckCustomerCode(ByRef sumhibtcd As Integer, ByRef hibtcd As String)
        If sumhibtcd > 0 Then
            intNext = intNext + 1
            intCurrRow = intNext
            Call GrandTotal(intCurrRow)
            sumhibtcd = 0
        End If
        Call SaveFile(iFile)
        intCurrRow = 8
        intNext = 8
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)    'default first sheet   
        sheet2 = appExcel.Worksheets.Create(Trim(oRow("HIBTCD")))
        Call SetPageProperties(sheet2)
        intFile = 0
        sheet.Range("A1:V8").CopyTo(sheet2.Range("A1"))
        Call PrintH()
    End Sub
    Protected Sub DataDetail(ByRef sumcurINV As Integer, ByRef sumhibtcd As Integer)
        sheet2.Range(intCurrRow, 2).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft
        sheet2.Range(intCurrRow, 2).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
        sheet2.Range(intCurrRow, 2).Value = "'" & Right(oRow("HIBLDT"), 2) & "/" & Mid(oRow("HIBLDT"), 5, 2) & "/" & _
                                            Left(oRow("HIBLDT"), 4)  'ChgDate(rs("hibldt").value)  'Change Format dd/mm/yyyy
        sheet2.Range(intCurrRow, 4).Value = Trim(oRow("DIFACD"))
        sheet2.Range(intCurrRow, 3).Value = "'" & Trim(oRow("HIDONO"))
        sheet2.Range(intCurrRow, 5).Value = "'" & Trim(oRow("DIPROD"))
        sheet2.Range(intCurrRow, 6).Value = "'" & Trim(oRow("DIMKSF"))
        sheet2.Range(intCurrRow, 7).Value = Trim(oRow("QTYAM"))
        sheet2.Range(intCurrRow, 8).Value = Trim(oRow("DISPUN"))
        sheet2.Range(intCurrRow, 9).Value = Trim(oRow("SALEAM"))
        sheet2.Range(intCurrRow, 10).Value = Trim(oRow("DIFCUN"))
        sheet2.Range(intCurrRow, 11).Value = Trim(oRow("FCAM"))
        sheet2.Range(intCurrRow, 14).Value = Trim(oRow("DIPDCD"))
        sheet2.Range(intCurrRow, 15).Value = Trim(oRow("CARLINE"))
        sheet2.Range(intCurrRow, 16).Value = Trim(oRow("HIREFE"))
        sheet2.Range(intCurrRow, 17).Value = "T"
        If Trim(oRow("DIFCTP")) = "1" Then
            sheet2.Range(intCurrRow, 17).Value = "F"
        Else
            sheet2.Range(intCurrRow, 17).Value = "T"
        End If
        If Trim(oRow("DISPTP")) = "1" Then
            sheet2.Range(intCurrRow, 18).Value = "F"
        Else
            sheet2.Range(intCurrRow, 18).Value = "T"
        End If
        sheet2.Range(intCurrRow, 19).Value = Trim(oRow("DICOPP"))
        sheet2.Range(intCurrRow, 21).Value = Trim(oRow("DIMANH"))
        '-------- For Summary SubTotal------------
        GrandCol7 += sheet2.Range(intCurrRow, 7).Number
        GrandCol9 += sheet2.Range(intCurrRow, 9).Number
        GrandCol11 += sheet2.Range(intCurrRow, 11).Number
        '-------- For Summary GrandTotal----------
        SubCol7 += sheet2.Range(intCurrRow, 7).Number
        SubCol8 += sheet2.Range(intCurrRow, 8).Number
        SubCol9 += sheet2.Range(intCurrRow, 9).Number
        SubCol11 += sheet2.Range(intCurrRow, 11).Number
        '-----------------------------------------
        intNext = intNext + 1
        sheet.Range("A8:V8").CopyTo(sheet2.Range("A" & intNext))
        intCurrRow = intNext

        curINV = Trim(oRow("HIDONO"))
        hibtcd = Trim(oRow("HIBTCD"))
        sumcurINV = sumcurINV + 1
        sumhibtcd = sumhibtcd + 1

        If intCurrRow >= 65000 Then
            sheet2.Name = Trim(oRow("HIBTCD")) & "_0" & intFile
            intFile = intFile + 1
            intCurrRow = 8
            intNext = 8
            sumcurINV = 0
            sheet2 = appExcel.Worksheets.Create(Trim(oRow("HIBTCD")) & "_0" & intFile)
            Call SetPageProperties(sheet2)
            sheet.Range("A1:V8").CopyTo(sheet2.Range("A1"))
            Call PrintH()
        End If
    End Sub
    Protected Sub SubTotal(ByVal sheet As IWorksheet, ByVal sheet2 As IWorksheet, ByVal intCurrRow As Integer)
        sheet.Range("A9:V9").CopyTo(sheet2.Range("A" & intCurrRow))
        sheet2.Range(intCurrRow, 2).Value = "SUB TOTAL         :        " & curINV
        If Left(sheet2.Range(intCurrRow, 2).Text, 3) = "SUB" Then
            sheet2.Range(intCurrRow, 7).Number = SubCol7
            sheet2.Range(intCurrRow, 8).Number = SubCol8
            sheet2.Range(intCurrRow, 9).Number = SubCol9
            sheet2.Range(intCurrRow, 10).Number = SubCol10
            sheet2.Range(intCurrRow, 11).Number = SubCol11
        Else
            sheet2.Range(intCurrRow, 7).Number = SubCol7
            sheet2.Range(intCurrRow, 8).Number = SubCol8
            sheet2.Range(intCurrRow, 9).Number = SubCol9
            sheet2.Range(intCurrRow, 11).Number = SubCol11
        End If
        sheet2.Range(intCurrRow, 7).NumberFormat = "##,###"
        sheet2.Range(intCurrRow, 8).NumberFormat = "##,###.#0"
        sheet2.Range(intCurrRow, 9).NumberFormat = "##,###.#0"
        sheet2.Range(intCurrRow, 11).NumberFormat = "##,###.#0"
        Call ClearValueTotal("SubTotal")
    End Sub
    Protected Sub GrandTotal(ByVal intCurrRow As Integer)
        sheet.Range("A9:V9").CopyTo(sheet2.Range("A" & intCurrRow))
        sheet2.Range(intCurrRow, 2).Value = "GRAND TOTAL"
        sheet2.Range(intCurrRow, 7).Number = GrandCol7
        sheet2.Range(intCurrRow, 9).Number = GrandCol9
        sheet2.Range(intCurrRow, 11).Number = GrandCol11
        Call ClearValueTotal("GrandTotal")
    End Sub
    Protected Sub ClearValueTotal(ByVal iCase As String)
        Select Case iCase
            Case "SubTotal"
                SubCol7 = 0.0
                SubCol8 = 0.0
                SubCol9 = 0.0
                SubCol11 = 0.0
            Case "GrandTotal"
                GrandCol7 = 0.0
                GrandCol9 = 0.0
                GrandCol11 = 0.0
        End Select
    End Sub
    Protected Sub GrandTotal2(ByRef sumhibtcd As Integer)
        If sumhibtcd > 0 Then
            intNext = intNext + 1
            intCurrRow = intNext
            Call GrandTotal(intCurrRow)
            sumhibtcd = 0
            intNext = 0
            intCurrRow = 0
        End If
    End Sub

    Protected Sub SaveFile(ByRef iFile As Integer)
        If iFile >= 1 Then
            FileName = ReportName & "-" & hibtcd & "-" & ClassFunctionVar.STRDate & ".xls"
        Else
            FileName = ReportName & "-" & Trim(oRow("hibtcd")) & "-" & ClassFunctionVar.STRDate & ".xls"
        End If

        CommonFunction.SaveTmpFile(CurTempPath, FileName)
        iFile = iFile + 1
        FileName = GetFileName(FileName)
        FileName = CurTempPath & FileName
        sheet.Remove()
        Workbook.SaveAs(FileName)
        Workbook.Close(True)
    End Sub

    Private Function GetFileName(ByRef Filename As String) As String
        Call CommonFunction.CreateFileNameSession(idx, dt, drow, Filename)
        Return Filename
    End Function

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        Response.Redirect("~\download.aspx?file=" & GridView2.SelectedValue & "&path=" & DefaultPath)
    End Sub

    Protected Sub SetPageProperties(ByVal sheet2 As IWorksheet)
        With sheet2
            .SetColumnWidth(1, 0.5)
            .SetColumnWidth(2, 16.17)               'COLUMN B -- SHIPPING DATE
            .SetColumnWidth(3, 24.86)               'COLUMN C -- INVOICE NO.
            .SetColumnWidth(4, 7.57)                'COLUMN D -- FACTORY
            .SetColumnWidth(5, 18.43)               'COLUMN E -- PART NO.
            .SetColumnWidth(6, 6.29)                'COLUMN F -- SUFFIX
            .SetColumnWidth(7, 8)                   'COLUMN G -- QTY
            .SetColumnWidth(8, 10.57)               'COLUMN H -- UNIT PRICE (SALES)
            .SetColumnWidth(9, 15.43)               'COLUMN I -- SALE AMOUNT
            .SetColumnWidth(10, 11.71)              'COLUMN J -- UNIT COST (FACTORY)
            .SetColumnWidth(11, 12.57)              'COLUMN K -- COST AMOUNT
            .SetColumnWidth(12, 13.14)              'COLUMN L -- PROFIT AMOUN
            .SetColumnWidth(13, 5.87)               'COLUMN M -- %
            .SetColumnWidth(14, 7.14)               'COLUMN N -- PRODUCT CODE
            .SetColumnWidth(15, 7.57)               'COLUMN O -- CARLINE CODE
            .SetColumnWidth(16, 7.14)               'COLUMN P -- DN NO.
            .SetColumnWidth(17, 7.14)               'COLUMN Q -- FC TYPE
            .SetColumnWidth(18, 8.43)               'COLUMN R -- PRICE TYPE
            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Landscape
            .PageSetup.LeftMargin = 0.18
            .PageSetup.RightMargin = 0.16
            .PageSetup.TopMargin = 0.22
            .PageSetup.BottomMargin = 0.23
            .PageSetup.Zoom = 80
            sheet2.PageSetup.PrintTitleRows = "$1:$7"
            sheet2.Range("A8").FreezePanes()
            'sheet2.Range("D1").FreezePanes()
        End With
    End Sub
End Class
