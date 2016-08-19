Imports DataAccess
Imports System.IO
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Partial Class Default2
    Inherits System.Web.UI.Page
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("StatementTemplate.xlt")
    Private ReportName As String = "Statement-Report"
    Private crtDate As String = Val(Date.Now.ToString("ddMMyyyy"))
    Private crtTime As String = Val(Date.Now.ToString("HHmm"))
    Private _DBconnection As DBConnection
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = ExcelEngine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Public ReadOnly Property DBConnect() As DBConnection
        Get
            If _DBconnection Is Nothing Then
                _DBconnection = New DBConnection
            End If
            Return _DBconnection
        End Get
    End Property
    Protected Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        printreport()
    End Sub

    Private Function GetData() As DataTable
        Dim strsql As String
        Dim startDate As String
        Dim endDate As String
        startDate = Right(ddlYear.Text, 2) & ddlMonth.Text & "01"
        endDate = Right(ddlYear.Text, 2) & ddlMonth.Text & "31"
        Using sr As StreamReader = New StreamReader(Server.MapPath("..") & "/scriptQuery/qryStatement.lbs")
            strsql = String.Format(sr.ReadToEnd(), startDate, endDate)
        End Using
        'strsql = String.Format("select 'NS' || RIGHT('0000' || H.STINVN,4) INVNO,h.STINVD INVDATE,STCST COST,TLQTY,STTAX,STTOT Amount " & _
        '        "from TYNSSALE_D.siht8 H inner join TYNSSALE_D.silt8 D ON H.STINVN = D.TLINVN " & _
        '        "Where STINVD between {0} and {1}", startDate, endDate)
        Dim dt As New DataTable
        dt = DBConnect.ExcuteQueryString(strsql, DBConnection.DatabaseType.AS400)
        Return dt
    End Function
    Private Function GenerateStatementNo() As String
        Return "SK" & (CInt(ddlMonth.SelectedItem.Value) + (12 * (CInt(ddlYear.SelectedItem.Value) - 2015)) - 5).ToString().PadLeft(5, "0")
    End Function
    Protected Sub printreport()
        Dim filename As String
        Dim dt As New DataTable
        Dim intcurrow As Integer
        dt = GetData()
        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create("Statement")
        intcurrow = 5
        With sheet2
            Workbook.Worksheets(0).Range("A1:G5").CopyTo(.Range("A1"))
            .Range(4, 7).Value = String.Format("{0:dd/MM/yyyy}", Now().Date)
            '.Range(1, 7).Value = String.Format("{0:dd/MM/yyyy}", Now().Date)
            '.Range(3, 13).Value = "Print by		: " & Session("userId") & "  -  " & Session("username")
            .Range(1, 7).Value = "As " & ddlMonth.SelectedItem.Text & "'" & Right(ddlYear.Text, 2)
            .Range(3, 7).Value = dt.Rows(0)("SDNO")
            For Each oRow As DataRow In dt.Rows

                intcurrow = intcurrow + 1
                Workbook.Worksheets(0).Range("A7:G7").CopyTo(.Range("A" & intcurrow))
                '.Range(intcurrow, 1).Value = intno
                .Range(intcurrow, 2).Text = oRow("INVNO")

                .Range(intcurrow, 3).Value = Right(oRow("INVDATE"), 2) & "/" & Mid(oRow("INVDATE"), 3, 2) & "/" & Left(oRow("INVDATE"), 4)
                .Range(intcurrow, 4).Value = oRow("COST")
                .Range(intcurrow, 5).Value = oRow("TLQTY")
                .Range(intcurrow, 6).Value = oRow("STTAX")
                .Range(intcurrow, 7).Value = oRow("Amount")

            Next

            'Workbook.Worksheets(0).Range("A6:G6").CopyTo(.Range("A" & intcurrow))
            '.Range("A" & intcurrow & ":G" & intcurrow).CellStyle.ColorIndex = 35
            '.Range("A" & intcurrow & ":G" & intcurrow).CellStyle.Font.Bold = True
            '.Range(intcurrow, 2).Value = "SUB Total ( " & oRow("customer") & "-" & Trim(oRow("customername")) & " )"
            'Call sumtotal()
            intcurrow += 1
        End With
        sumtotal(intcurrow, intcurrow - 1)
        Call SetPageProperties(intcurrow)
        sheet.Remove()
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)

    End Sub
    Protected Sub sumtotal(ByVal intcurrow As Integer, ByVal intstart As Integer)
        With sheet2
            Workbook.Worksheets(0).Range("A30:G30").CopyTo(.Range("A" & intcurrow))
            For i As Integer = intstart To intcurrow - 1
                For a As Integer = 4 To 7
                    .Range(intcurrow, a).FormulaR1C1 = "= SUM(R[" & -intstart + 5 & "]C:R[-1]C)"
                Next
            Next
        End With
    End Sub

    Protected Sub SetPageProperties(ByVal intcurrow As Integer)
        With sheet2
            .SetColumnWidth(1, 11)
            .SetColumnWidth(2, 13)
            .SetColumnWidth(3, 14)
            .SetColumnWidth(4, 13)
            .SetColumnWidth(5, 11)
            .SetColumnWidth(6, 11)
            .SetColumnWidth(7, 13)
            'For i As Integer = 8 To 12
            '    .SetColumnWidth(i, 10)
            'Next
            '.SetColumnWidth(13, 11)
            '.SetColumnWidth(14, 10)
            '.SetColumnWidth(15, 26.4)
            'For i As Integer = 6 To intcurrow
            '    .SetRowHeight(i, 15.75)
            'Next
            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Portrait
            .PageSetup.LeftMargin = 0.5
            .PageSetup.RightMargin = 0.5
            .PageSetup.TopMargin = 0.5
            .PageSetup.BottomMargin = 0.5
            .PageSetup.Zoom = 100
            .PageSetup.PrintTitleRows = "$5:$7"
            .Range("A6").FreezePanes()
        End With
    End Sub
End Class
