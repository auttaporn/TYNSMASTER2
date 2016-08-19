Imports System.Data
Imports System.Data.OleDb
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Imports System.IO
Partial Class TYL_TYLINV
    Inherits System.Web.UI.Page
    Private Conn As New OleDbConnection(Classconn.strConnSql)
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("TYE-SALEREPORT.xlt")
    Private ReportName As String = "KEISO-SALEREPORT"
    Private crtDate As String = Val(Date.Now.ToString("ddMMyyyy"))
    Private crtTime As String = Val(Date.Now.ToString("HHmm"))
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = ExcelEngine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private oTable As New DataTable
    Private oRow As DataRow
    Private filename As String
    Private intcurrow, intno, intstart As Integer
    Private Libl As String = "tynssale_d"


    Protected Sub sumtotal()
        With sheet2
            For i As Integer = intstart To intcurrow - 1
                'For a As Integer = 7 To 14
                '    '.Range(intcurrow, a).Value = .Range(intcurrow, a).Value + .Range(i, a).Value
                '    .Range(intcurrow, a).FormulaR1C1 = "= SUM(R[" & -intstart & "]C:R[-1]C)"
                'Next
                For a As Integer = 8 To 14
                    If a = 12 Then
                        Continue For
                    End If
                    '.Range(intcurrow, a).Value = .Range(intcurrow, a).Value + .Range(i, a).Value
                    .Range(intcurrow, a).FormulaR1C1 = "= SUM(R[" & -intstart & "]C:R[-1]C)"
                Next
            Next
        End With
    End Sub

    Protected Sub printreport()
        Dim customer As String

        oTable = getdata()

        If oTable.Rows.Count <= 0 Then
            lblmsg.Text = "No Data"
            Exit Sub
        End If
        oRow = oTable.Rows(0)
        customer = Trim(oRow("CUSTOMER"))
        intcurrow = 6
        intstart = 0
        intno = 1
        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create("KEISO")
        With sheet2
            Workbook.Worksheets(0).Range("A1:O5").CopyTo(.Range("A1"))
            .Range(2, 13).Value = "Printdate	: " & Now()
            .Range(3, 13).Value = "Print by		: " & Session("userId") & "  -  " & Session("username")
            .Range(3, 2).Value = "As " & Right(Trim(txtfromdate.Text), 2) & "/" & Mid(Trim(txtfromdate.Text), 5, 2) & "/" & Mid(Trim(txtfromdate.Text), 3, 2) & " - " & Right(Trim(txttodate.Text), 2) & "/" & Mid(Trim(txttodate.Text), 5, 2) & "/" & Mid(Trim(txttodate.Text), 3, 2) & ""

            For Each oRow In oTable.Rows

                If customer <> Trim(oRow("customer")) Then
                    Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow))
                    .Range("A" & intcurrow & ":O" & intcurrow).CellStyle.ColorIndex = 35
                    .Range("A" & intcurrow & ":O" & intcurrow).CellStyle.Font.Bold = True
                    .Range(intcurrow, 2).Value = "SUB Total ( " & oRow("customer") & "-" & Trim(oRow("customername")) & " )"
                    Call sumtotal()
                    customer = Trim(oRow("customer"))
                    intcurrow += 1
                    intstart = 0

                End If
                Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow))
                .Range(intcurrow, 1).Value = intno
                .Range(intcurrow, 2).Text = oRow("invoiceno")
                .Range(intcurrow, 3).Value = Right(oRow("invoicedate"), 2) & "/" & Mid(oRow("invoicedate"), 3, 2) & "/20" & Left(oRow("invoicedate"), 2)
                .Range(intcurrow, 4).Value = oRow("customer") & "-" & Trim(oRow("customername"))
                .Range(intcurrow, 5).Value = oRow("deto")
                .Range(intcurrow, 6).Value = oRow("desc")
                .Range(intcurrow, 7).Value = oRow("price")
                .Range(intcurrow, 8).Value = oRow("qty")
                .Range(intcurrow, 9).Value = oRow("amount")
                .Range(intcurrow, 10).Value = oRow("vat")
                .Range(intcurrow, 11).Value = oRow("amount") + (.Range(intcurrow, 10).Value)
                .Range(intcurrow, 12).Value = oRow("fc")
                .Range(intcurrow, 13).Value = oRow("amountfc")
                .Range(intcurrow, 14).Value = oRow("amount") - oRow("amountfc")
                intno += 1
                intcurrow += 1
                intstart += 1
            Next

            Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow))
            .Range("A" & intcurrow & ":O" & intcurrow).CellStyle.ColorIndex = 35
            .Range("A" & intcurrow & ":O" & intcurrow).CellStyle.Font.Bold = True
            .Range(intcurrow, 2).Value = "SUB Total ( " & oRow("customer") & "-" & Trim(oRow("customername")) & " )"
            Call sumtotal()
            intcurrow += 1
        End With
        Call SetPageProperties()
        sheet.Remove()
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)

    End Sub
    Protected Function getdata()
        Dim sql As String
        Using r As StreamReader = New StreamReader(Server.MapPath("..") & "/scriptQuery/qrySalesumaryKEISO.lbs")
            sql = String.Format(r.ReadToEnd(), Libl, Right(Trim(txtfromdate.Text), 6), Right(Trim(txttodate.Text), 6))
        End Using
        'sql = "Select * FROM OPENQUERY(AS400,'Select ''NS'' || Right(''0000'' || h.stinvn,4) as invoiceNo,h.stinvd as InvoiceDate,CYGSSC as Customer,"
        'sql += "c.cnme as CustomerName, h.statn as DeTo,p.IDSCE as desc,d.tlnet as Price,d.tlqty as QTY,"
        'sql += "h.sttot-h.sttax as Amount,h.sttax as VAt,d.tlta10 as FC,d.tlta10* d.tlqty as AmountFC From " & Libl & ".siht8 h inner "
        'sql += "join " & Libl & ".silt8 d On stinvn = tlinvn Left Outer Join " & Libl & ".trcm c On h.stbtcd = c.ccust "
        'sql += "Left Outer Join " & Libl & ".tiim p On d.tlprod = p.iprod "
        'sql += "inner join tynssale_d.sih th on th.siinvn = h.stord "
        'sql += "inner Join tynssale_d.ech o On th.siord = o.hord "
        'sql += "WHERE hstat in (''7'',''8'') AND (stinvd >= " & Right(Trim(txtfromdate.Text), 6) & ") AND (stinvd <= " & Right(Trim(txttodate.Text), 6) & ") Order By "
        'sql += " Customer, DeTo, invoiceno, InvoiceDate ')"

        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(sql, Conn)
        da.Fill(ds, "TYL")
        Return ds.Tables("TYL")

    End Function

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        If txtfromdate.Text = "" Or txttodate.Text = "" Then
            lblmsg.Text = "Please enter valided From Month and To Month  field  and try again."
            Exit Sub
        ElseIf Mid(Trim(txttodate.Text), 5, 2) > Mid(Trim(txtfromdate.Text), 5, 2) Then
            'lblmsg.Text = "Please check month." ', vbOKOnly, "Invalid Data")
            'Exit Sub
        End If

        Call printreport()

    End Sub

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

End Class

