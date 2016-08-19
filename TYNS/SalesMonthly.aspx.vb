Imports System.IO
Imports System.Data.OleDb
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Partial Class SalesMonthly
    Inherits System.Web.UI.Page
    Private Conn As New OleDbConnection(Classconn.strConnSql)
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("TYNS-SALEREPORT1.xlt")
    Private ReportName As String = "TYNS-SALEREPORT"
    Private crtDate As String = Val(Date.Now.ToString("ddMMyyyy"))
    Private crtTime As String = Val(Date.Now.ToString("HHmm"))
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private oTable As New DataTable
    Private oRow As DataRow
    Private filename As String
    Private intcurrow, intno, intstart As Integer
    Private Libl As String = "tynssale_d"
    Private SumTotAmt, SumAmt As Double


    Protected Sub sumtotal(ByVal lastRecord As Boolean)
        With sheet2
            For i As Integer = intstart To intcurrow - 1
                For a As Integer = 8 To 14
                    If a = 12 Then
                        Continue For
                    End If
                    '.Range(intcurrow, a).Value = .Range(intcurrow, a).Value + .Range(i, a).Value
                    .Range(intcurrow, a).FormulaR1C1 = "= SUM(R[" & -intstart & "]C:R[-1]C)"

                Next
            Next
            If lastRecord Then
                Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow + 1))
                .Range("A" & intcurrow + 1 & ":O" & intcurrow + 1).CellStyle.ColorIndex = 33
                .Range("A" & intcurrow + 1 & ":O" & intcurrow + 1).CellStyle.Font.Bold = True
                .Range(intcurrow + 1, 2).Value = "  SUB  TOTAL SUMMARY : "
                .Range(intcurrow + 1, 9).Value = SumAmt
                .Range(intcurrow + 1, 11).Value = SumTotAmt
            End If
        End With
    End Sub

    Public Function GetMonth(ByVal i As Integer) As String
        If i = 1 Then
            Return "January"
        ElseIf i = 2 Then
            Return "Febuary"
        ElseIf i = 3 Then
            Return "March"
        ElseIf i = 4 Then
            Return "April"
        ElseIf i = 5 Then
            Return "May"
        ElseIf i = 6 Then
            Return "June"
        ElseIf i = 7 Then
            Return "July"
        ElseIf i = 8 Then
            Return "August"
        ElseIf i = 9 Then
            Return "September"
        ElseIf i = 10 Then
            Return "October"
        ElseIf i = 11 Then
            Return "November"
        ElseIf i = 12 Then
            Return "December"
        End If
        Return ""
    End Function
    Protected Sub printreport()
        Dim customer As String
        Dim customerName As String
        oTable = getdata()

        If oTable.Rows.Count <= 0 Then
            '  lblmsg.Text = "No Data"
            Exit Sub
        End If
        oRow = oTable.Rows(0)
        customer = Trim(oRow("CUSTOMER"))
        customerName = Trim(oRow("customername"))
        intcurrow = 6
        intstart = 0
        intno = 1
        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create("TYL")
        With sheet2
            Workbook.Worksheets(0).Range("A1:Q5").CopyTo(.Range("A1"))
            .Range(2, 13).Value = "Printdate	: " & Now()
            .Range(3, 13).Value = "Print by		: " & Session("userId") & "  -  " & Session("username")
            '     .Range(3, 2).Value = "As " & Right(Trim(txtfromdate.Text), 2) & "/" & Mid(Trim(txtfromdate.Text), 5, 2) & "/" & Mid(Trim(txtfromdate.Text), 3, 2) & " - " & Right(Trim(txttodate.Text), 2) & "/" & Mid(Trim(txttodate.Text), 5, 2) & "/" & Mid(Trim(txttodate.Text), 3, 2) & ""

            For Each oRow In oTable.Rows

                'If customer <> Trim(oRow("customer")) Then
                '    Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow))
                '    .Range("A" & intcurrow & ":O" & intcurrow).CellStyle.ColorIndex = 34
                '    .Range("A" & intcurrow & ":O" & intcurrow).CellStyle.Font.Bold = True
                '    .Range(intcurrow, 2).Value = "SUB Total ( " & customer & "-" & customerName & " )"
                '    Call sumtotal(False)
                '    customer = Trim(oRow("customer").ToString())
                '    customerName = Trim(oRow("customername").ToString())
                '    intcurrow += 1
                '    intstart = 0

                'End If
                Workbook.Worksheets(0).Range("A6:Q6").CopyTo(.Range("A" & intcurrow))
                .Range(intcurrow, 1).Value = intno
                .Range(intcurrow, 2).Text = oRow("invoiceno").ToString()
                .Range(intcurrow, 3).Value = Right(oRow("invoicedate"), 2) & "/" & Mid(oRow("invoicedate"), 3, 2) & "/20" & Left(oRow("invoicedate"), 2)
                .Range(intcurrow, 4).Value = oRow("customer") & "-" & Trim(oRow("customername"))
                .Range(intcurrow, 5).Value = oRow("deto").ToString()
                .Range(intcurrow, 6).Value = oRow("desc").ToString()
                .Range(intcurrow, 7).Value = oRow("price").ToString()
                .Range(intcurrow, 8).Value = oRow("qty").ToString()
                .Range(intcurrow, 9).Value = IIf(oRow("qty") < 0, CStr(-1 * CInt(oRow("amount"))), oRow("amount"))
                .Range(intcurrow, 10).Value = IIf(oRow("qty") < 0, CStr(-1 * CInt(oRow("vat"))), oRow("vat"))
                .Range(intcurrow, 11).Value = IIf(oRow("qty") < 0, CStr(((-1) * CInt(oRow("amount")))) + (IIf(oRow("qty") < 0, CInt(-1 * CInt(oRow("vat"))), CInt(oRow("vat")))), oRow("amount") + oRow("vat"))
                .Range(intcurrow, 12).Value = oRow("fc").ToString()
                .Range(intcurrow, 13).Value = oRow("amountfc").ToString()
                .Range(intcurrow, 14).Value = CInt(.Range(intcurrow, 9).Value) - oRow("amountfc")
                .Range(intcurrow, 16).Value = Mid(oRow("invoicedate"), 3, 2) 'GetMonth(CInt(Mid(oRow("invoicedate"), 3, 2)))
                .Range(intcurrow, 17).Value = "20" & Left(oRow("invoicedate"), 2)
                '.Range(intcurrow, 7).Value.Format("{0:#.#0}", oRow("price"))
                'SumTotAmt = SumTotAmt + CDbl(.Range(intcurrow, 11).Value)
                'SumAmt = SumAmt + CDbl(.Range(intcurrow, 9).Value)
                intno += 1
                intcurrow += 1
                intstart += 1
            Next

            'Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow))
            '.Range("A" & intcurrow & ":O" & intcurrow).CellStyle.ColorIndex = 34
            '.Range("A" & intcurrow & ":O" & intcurrow).CellStyle.Font.Bold = True
            '.Range(intcurrow, 2).Value = "SUB Total ( " & oRow("customer") & "-" & Trim(oRow("customername")) & " )"
            'Call sumtotal(True)
            'intcurrow += 1
        End With
        Call SetPageProperties()
        'sheet.Remove()
        Workbook.Worksheets("TYL").Range("A6:Q1000").CopyTo(Workbook.Worksheets(0).Range("A6:Q1000"))
        Workbook.Worksheets(1).Range(1, 6).Value = "Printdate	: " & Now()
        Workbook.Worksheets(1).Range(2, 6).Value = "Print by	: " & Session("userId") & "  -  " & Session("username")
        Workbook.Worksheets("TYL").Remove()
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)

    End Sub
    Protected Function getdata()
        Dim strdatestart As String = datestart.Text.Replace("-", "").Substring(2) & "01"
        Dim strdateend As String = dateto.Text.Replace("-", "").Substring(2) & "32"
        Dim sql As String

        'sql = "Select * FROM OPENQUERY(AS400,'Select ''S'' || right(''0000'' || h.siinvn,5) as invoiceNo,h.siinvd as InvoiceDate,CYGSSC as Customer,"
        'sql += "c.cnme as CustomerName, h.siatn as DeTo,p.IDSCE as desc,d.ilnet as Price,d.ilqty as QTY,"
        'sql += "h.sitot-h.sitax as Amount,h.sitax as VAt,d.ilta10 as FC,d.ilta10* d.ilqty as AmountFC From " & Libl & ".sih h inner "
        'sql += "join " & Libl & ".sil d On siinvn = ilinvn Left Join " & Libl & ".trcm c On h.sicust = c.ccust "
        'sql += "Left Join " & Libl & ".tiim p On left(d.ilprod,5) = left(p.iprod,5)  Inner Join " & Libl & ".ech o On h.siord = o.hord "
        'sql += "WHERE hstat in (''7'',''8'') AND (siinvd >= " & strdatestart & ") AND (siinvd < " & strdateend & ") Order By "
        'sql += " Customer, DeTo, invoiceno, InvoiceDate ')"

        Using sr As StreamReader = New StreamReader(Server.MapPath("..") & "/scriptQuery/queryMonthly.lbs")
            sql = String.Format(sr.ReadToEnd(), Libl, strdatestart, strdateend)
        End Using

        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(sql, Conn)
        da.Fill(ds, "TYL")
        Return ds.Tables("TYL")

    End Function


    Protected Sub SetPageProperties()
        With sheet2
            .SetColumnWidth(1, 4)
            .SetColumnWidth(2, 10.8)
            .SetColumnWidth(3, 10.4)
            .SetColumnWidth(4, 21.7)
            .SetColumnWidth(5, 9)
            .SetColumnWidth(6, 22)
            .SetColumnWidth(7, 7)
            For i As Integer = 8 To 12
                .SetColumnWidth(i, 10)
            Next
            .SetColumnWidth(13, 11)
            .SetColumnWidth(14, 10)
            .SetColumnWidth(15, 26.4)
            For i As Integer = 6 To intcurrow
                .SetRowHeight(i, 15.75)
            Next
            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Landscape
            .PageSetup.LeftMargin = 0.5
            .PageSetup.RightMargin = 0.5
            .PageSetup.TopMargin = 0.5
            .PageSetup.BottomMargin = 0.5
            .PageSetup.Zoom = 70
            .PageSetup.PrintTitleRows = "$5:$15"
            .Range("A6").FreezePanes()
        End With
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub btnGenerate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGenerate.Click
        If datestart.Text = "" Or dateto.Text = "" Then
            'lblmsg.Text = "Please enter valided From Month and To Month  field  and try again."
            Exit Sub
        ElseIf Mid(Trim(dateto.Text), 5, 2) > Mid(Trim(datestart.Text), 5, 2) And Mid(Trim(dateto.Text), 1, 4) <= Mid(Trim(datestart.Text), 1, 4) Then
            '    lblmsg.Text = "Please check month." ', vbOKOnly, "Invalid Data")
            Exit Sub
        End If

        Call printreport()
    End Sub
End Class
