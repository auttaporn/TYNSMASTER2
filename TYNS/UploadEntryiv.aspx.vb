Imports System.IO
Imports System.Data.OleDb
Imports DataAccess
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Partial Class UploadEntryiv
    Inherits System.Web.UI.Page
    Private _DBConnection As DBConnection
    Public ReadOnly Property DBConnect() As DBConnection
        Get
            If _DBConnection Is Nothing Then
                _DBConnection = New DBConnection()
            End If
            Return _DBConnection
        End Get
    End Property

    Private strLib As String = "TYNSSALE_D"
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("TYNS-EXPORTINVOICE.xlt")
    Private ReportName As String = "ExportInvoiceEntry"
    Private crtDate As String = Val(Date.Now.ToString("ddMMyyyy"))
    Private crtTime As String = Val(Date.Now.ToString("HHmm"))
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = ExcelEngine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet

    Protected Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
       
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FileUpload1.Attributes.Add("onchange", "return checkFileExtension(this);")
    End Sub

    Protected Sub FileUpload1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles FileUpload1.Load
        If FileUpload1.HasFile Then

            Dim uplod As Boolean = True
            Dim fleUpload As String = Path.GetExtension(FileUpload1.FileName.ToString())
               
                   
            FileUpload1.SaveAs(Server.MapPath("~/UploadFile/" & FileUpload1.FileName.ToString()))
            Dim uploadedFile As String = Server.MapPath("~/UploadFile/" + FileUpload1.FileName.ToString())
            'Try
            InsertFile(uploadedFile)
            'Catch ex As Exception
            '  MessageBox(ex.Message)
            'End Try
        End If
    End Sub

    Public Sub InsertFile(ByVal pth As String)
        Dim strcon As String = String.Empty
        If Path.GetExtension(pth).ToLower().Equals(".xls") Then
            strcon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pth & ";Extended Properties=""Excel 8.0;HDR=YES;"""
        ElseIf Path.GetExtension(pth).ToLower().Equals(".xlsx") Then
            strcon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pth & ";Extended Properties=""Excel 12.0;HDR=YES;"""
        End If
        Dim strselect As String = "Select * from [Sheet1$]"
        Dim excelconn As New OleDbConnection(strcon)
        Dim da As New OleDbDataAdapter(strselect, excelconn)
        Dim ds As New DataSet

        'Try
        da.Fill(ds, "tb1")
        AssignData(ds.Tables(0))
        'Catch ex As OleDbException
        'Throw New Exception(ex.Message.ToString())
        'End Try

    End Sub

    Protected Function getdata() As DataTable
        Dim dt As New DataTable
        Dim sql As String
        'sql = "Select * FROM OPENQUERY(AS400,'Select ''S'' || right(''0000'' || h.siinvn,5) as invoiceNo,h.siinvd as InvoiceDate,h.sicust as Customer,"
        'sql += "c.cnme as CustomerName, h.siatn as DeTo,p.IDSCE as desc,round(((d.ilnet+d.ilnet)/0.02)*0.01,2) as Price,d.ilqty as QTY,"
        'sql += "h.sitot-h.sitax as Amount,Round((h.sitot-h.sitax)*.07,2) as VAt,d.ilta10 as FC,d.ilta10* d.ilqty as AmountFC From " & Libl & ".sih h inner "
        'sql += "join " & Libl & ".sil d On siinvn = ilinvn Left Join " & Libl & ".trcm c On h.sicust = c.ccust "
        'sql += "Left Join " & Libl & ".tiim p On d.ilprod = p.iprod Inner Join " & Libl & ".ech o On h.siord = o.hord "
        'sql += "WHERE hstat in (''7'',''8'') AND (siinvd >= " & Right(Trim(txtfromdate.Text), 6) & ") AND (siinvd <= " & Right(Trim(txttodate.Text), 6) & ") Order By "
        'sql += " Customer, DeTo, invoiceno, InvoiceDate ')"
        sql = "Select 'S' || right('0000' || h.siinvn,5) as invoiceNo,h.siinvd as InvoiceDate,CYGSSC as Customer,EEINVN,EEEXIV,EECSNO,EECPDT,"
        sql += "c.cnme as CustomerName, h.siatn as DeTo,p.IDESC,p.IDSCE as desc,d.ilnet as Price,d.ilqty as QTY,"
        sql += "h.sitot-h.sitax as Amount,h.sitax as VAt,d.ilta10 as FC,d.ilta10* d.ilqty as AmountFC From " & strLib & ".sih h inner "
        sql += "join " & strLib & ".sil d On siinvn = ilinvn Left Join " & strLib & ".trcm c On h.sicust = c.ccust "
        sql += "Left Join " & strLib & ".tiim p On d.ilprod = p.iprod Inner Join " & strLib & ".ech o On h.siord = o.hord "
        sql += "inner join " & strLib & ".NSIVEE ON SIAD2 = trim(EEINVN) "
        sql += "WHERE hstat in ('7','8') AND (siinvd >= " & Right(Trim(txtDateS.Text), 6) & ") AND (siinvd <= " & Right(Trim(txtDateE.Text), 6) & ") Order By "
        sql += " Customer, DeTo, invoiceno, InvoiceDate"


        Dim ds As New DataSet
        dt = DBConnect.ExcuteQueryString(sql, DBConnection.DatabaseType.AS400)

        Return dt

    End Function

    Public Function AssignData(ByVal dt As DataTable) As Boolean
        Dim exiv As New EntryInvoice
        Dim createDate As Decimal
        Dim lotno As String
        lotno = String.Format("{0:yyyy/MM/dd hh:mm}", DateTime.Now)
        lotno = lotno.Replace("/", "")
        lotno = lotno.Replace(":", "")
        lotno = lotno.Replace(" ", "")
        createDate = CDec(lotno.Substring(0, 8))
        lotno = "LOT" & lotno
        Dim strsql As String
        Dim i As Integer = 0
        For Each r As DataRow In dt.Rows
            i = 0
            For Each c As DataColumn In dt.Columns
                If c.ColumnName.Length < 4 Then
                    Continue For
                End If
                If c.ColumnName.Substring(0, 3).ToLower() = "ord" Then
                    exiv.EEORID = CDec(r(c.ColumnName))
                    i = i + 1
                ElseIf c.ColumnName.Substring(0, 3).ToLower() = "pds" Then
                    exiv.EEPROD = r(c.ColumnName)
                    i = i + 1
                ElseIf c.ColumnName.Substring(0, 3).ToLower() = "del" Then
                    exiv.EEDLDT = CDec(String.Format("{0:yyyy/MM/dd}", CDate(r(c.ColumnName))).Replace("/", ""))
                    i = i + 1
                ElseIf c.ColumnName.Substring(0, 4).ToLower() = "q'ty" Then
                    exiv.EEQTY = CDec(r(c.ColumnName))
                    i = i + 1
                ElseIf c.ColumnName.Substring(0, 7).ToLower() = "exp_inv" Then  'Export Invoice No
                    exiv.EEINVN = r(c.ColumnName)
                    i = i + 1
                
                ElseIf c.ColumnName.Substring(0, 6).ToLower() = "custom" Then
                    exiv.EECSNO = r(c.ColumnName)
                    i = i + 1
                ElseIf c.ColumnName.Substring(0, 8).ToLower() = "complete" Then
                    exiv.EECPDT = CDec(String.Format("{0:yyyy/MM/dd}", CDate(r(c.ColumnName))).Replace("/", ""))
                    i = i + 1
                ElseIf c.ColumnName.Substring(0, 9).ToLower() = "exp_entry" Then
                    exiv.EEEXIV = r(c.ColumnName)
                    i = i + 1
                End If
            Next
            If i < 8 Then
                MessageBox("Excel File is incorrect format. Please check your imported file.")
                Return False
            End If

            strsql = "insert into " & strLib & ".NSIVEE(EEORID,EEPROD,EEDLDT,EEQTY,EEINVN,EEEXIV,EECSNO,EECPDT,EELOTN,EEIMDT)"
            strsql &= "values ({0},'{1}',{2},{3},'{4}','{5}','{6}',{7},'{8}',{9})"
            strsql = String.Format(strsql, exiv.EEORID, exiv.EEPROD, exiv.EEDLDT, exiv.EEQTY, exiv.EEINVN, exiv.EEEXIV, exiv.EECSNO, exiv.EECPDT, lotno, createDate)

            If DBConnect.ExcuteNonQueryString(strsql, DBConnection.DatabaseType.AS400) Then
                MessageBox(String.Format("Import Data Complete. Lot Number = {0}.", lotno))
                Return True
            Else
                strsql = String.Format("delete " & strLib & ".NSIVEE where EETNM1 = '{0}'", lotno)
                DBConnect.ExcuteNonQueryString(strsql, DBConnection.DatabaseType.AS400)
                MessageBox("Cannot Import Data.")
                Return False
            End If

        Next
    End Function
    Protected Sub lnkView_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lot_id As String
        lot_id = CType(sender, LinkButton).Attributes("LOTNO")

        Dim strsql As String
        strsql = "select * from " & strLib & ".NSIVEE where EELOTN = '" & lot_id & "'"
        Dim dt As New DataTable
        dt = DBConnect.ExcuteQueryString(strsql, DBConnection.DatabaseType.AS400)
        printReport(dt)
        sheet.Remove()
        Dim filename As String = "stockcompare"
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
    End Sub


    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub

    Protected Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        'Dim strsql As String
        'Dim startDt As String
        'Dim endDt As String
        'startDt = IIf(txtDateS.Text = "", "0", txtDateS.Text)
        'endDt = IIf(txtDateE.Text = "", "0", txtDateE.Text)
        'strsql = "select * from " & strLib & ".NSIVEE where (EEIMDT >= {0} or {0} = 0)  and (EEIMDT <= {1} or {1} = 0) "
        'strsql &= "and (EELOTN = '{2}' or '{2}' = '')"
        'strsql = String.Format(strsql, startDt, endDt, txtLotNo.Text)
        'Dim dt As New DataTable
        'dt = DBConnect.ExcuteQueryString(strsql, DBConnection.DatabaseType.AS400)
        'GridView1.DataSource = dt
        'GridView1.DataBind()
        printReport()
    End Sub

   
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim control As New LinkButton
        'Dim clbl As New Label
        'Dim strSpilt As String

        If e.Row.RowIndex > -1 And e.Row.RowType = DataControlRowType.DataRow Then
            control = CType(e.Row.FindControl("lnkView"), LinkButton)
            If Not control Is Nothing Then

                control.Attributes("LOTNO") = DataBinder.Eval(e.Row.DataItem, "EELOTN").ToString()

            End If

        End If
    End Sub
    Protected Sub printReport(ByVal dt As DataTable)
        Dim iRow As Integer
        Dim iCol As Integer = 1

        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)

        'sheet.Range(1, 9).Value = "'" & String.Format("{0:dd/MM/yyyy}", DateTime.Now)
        'sheet.Range(2, 9).Value = "'" & String.Format("{0}:{1}:{2}", DateTime.Now.Hour.ToString().PadLeft(2, "0"), DateTime.Now.Minute.ToString().PadLeft(2, "0"), DateTime.Now.Second.ToString().PadLeft(2, "0"))
        'For Each dt As DataTable In ds.Tables
        sheet2 = Workbook.Worksheets.Create("salereport")
        sheet.Range("A1:T1").CopyTo(sheet2.Range("A1"))
        iRow = 1
        Dim count As Integer = 0

        For Each r As DataRow In dt.Rows

            iRow = iRow + 1
            count = count + 1
            sheet.Range("A2:T2").CopyTo(sheet2.Range(iRow, 1))
            sheet2.Range(iRow, 1).Value = r("EEORID")
            sheet2.Range(iRow, 2).Value = r("EEPROD")
            sheet2.Range(iRow, 3).Value = "'" & r("EEDLDT").ToString().Substring(6, 2) & "/" & r("EEDLDT").ToString().Substring(4, 2) & "/" & r("EEDLDT").ToString().Substring(0, 4)
            sheet2.Range(iRow, 4).Value = r("EEQTY")
            sheet2.Range(iRow, 5).Value = r("EEINVN")
            sheet2.Range(iRow, 6).Value = r("EEEXIV")
            sheet2.Range(iRow, 7).Value = r("EECSNO")
            sheet2.Range(iRow, 8).Value = "'" & r("EECPDT").ToString().Substring(6, 2) & "/" & r("EECPDT").ToString().Substring(4, 2) & "/" & r("EECPDT").ToString().Substring(0, 4)


        Next

        Call SetPageProperties()
        'Next

    End Sub

    Protected Sub printreport()
        Dim customer As String
        Dim customerName As String
        Dim oTable As New DataTable
        oTable = getdata()

        If oTable.Rows.Count <= 0 Then

            Exit Sub
        End If
        Dim oRow As DataRow
        oRow = oTable.Rows(0)
        customer = Trim(oRow("CUSTOMER"))
        customerName = Trim(oRow("customername"))
        Dim intcurrow As Integer = 6
        Dim intstart As Integer = 0
        Dim intno As Integer = 1
        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create("TYL")
        With sheet2
            Workbook.Worksheets(0).Range("A1:O5").CopyTo(.Range("A1"))
            .Range(2, 11).Value = "Printdate	: " & Now()
            .Range(3, 11).Value = "Print by		: " & Session("userId") & "  -  " & Session("username")
            .Range(3, 2).Value = "As " & Right(Trim(txtDateS.Text), 2) & "/" & Mid(Trim(txtDateS.Text), 5, 2) & "/" & Mid(Trim(txtDateS.Text), 3, 2) & " - " & Right(Trim(txtDateE.Text), 2) & "/" & Mid(Trim(txtDateS.Text), 5, 2) & "/" & Mid(Trim(txtDateS.Text), 3, 2) & ""

            For Each oRow In oTable.Rows

                If customer <> Trim(oRow("customer")) Then
                    Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow))
                    .Range("A" & intcurrow & ":O" & intcurrow).CellStyle.ColorIndex = 34
                    .Range("A" & intcurrow & ":O" & intcurrow).CellStyle.Font.Bold = True
                    .Range(intcurrow, 2).Value = "SUB Total ( " & customer & "-" & customerName & " )"
                    'Call sumtotal(False)
                    customer = Trim(oRow("customer"))
                    customerName = Trim(oRow("customername"))
                    intcurrow += 1
                    intstart = 0

                End If
                Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow))
                .Range(intcurrow, 1).Value = intno
                .Range(intcurrow, 2).Text = oRow("invoiceno")
                .Range(intcurrow, 3).Value = "20" & Left(oRow("invoicedate"), 2) & "-" & Mid(oRow("invoicedate"), 3, 2) & "-" & Right(oRow("invoicedate"), 2)
                .Range(intcurrow, 4).Value = oRow("customer") & "-" & Trim(oRow("customername"))
                .Range(intcurrow, 5).Value = oRow("deto")
                .Range(intcurrow, 6).Value = IIf(oRow("desc").ToString().Trim() = "", oRow("idesc"), oRow("desc"))

                .Range(intcurrow, 7).Value = oRow("EEINVN")
                .Range(intcurrow, 8).Value = oRow("EEEXIV")
                .Range(intcurrow, 9).Value = oRow("EECSNO")
                .Range(intcurrow, 10).Value = "'" & oRow("EECPDT").ToString().Substring(6, 2) & "/" & oRow("EECPDT").ToString().Substring(4, 2) & "/" & oRow("EECPDT").ToString().Substring(0, 4)


                '.Range(intcurrow, 7).Value = oRow("price")
                '.Range(intcurrow, 8).Value = oRow("qty")
                '.Range(intcurrow, 9).Value = IIf(oRow("qty") < 0, CStr(-1 * CInt(oRow("amount"))), oRow("amount"))
                '.Range(intcurrow, 10).Value = IIf(oRow("qty") < 0, CStr(-1 * CInt(oRow("vat"))), oRow("vat"))
                '.Range(intcurrow, 11).Value = IIf(oRow("qty") < 0, CStr(((-1) * CInt(oRow("amount")))) + (IIf(oRow("qty") < 0, CInt(-1 * CInt(oRow("vat"))), CInt(oRow("vat")))), oRow("amount") + oRow("vat"))
                '.Range(intcurrow, 12).Value = oRow("fc")
                '.Range(intcurrow, 13).Value = oRow("amountfc")
                '.Range(intcurrow, 14).Value = CInt(.Range(intcurrow, 9).Value) - oRow("amountfc")
                '.Range(intcurrow, 7).Value.Format("{0:#.#0}", oRow("price"))
                'SumTotAmt = SumTotAmt + CDbl(.Range(intcurrow, 11).Value)
                'SumAmt = SumAmt + CDbl(.Range(intcurrow, 9).Value)
                intno += 1
                intcurrow += 1
                intstart += 1
            Next

            ' Workbook.Worksheets(0).Range("A6:O6").CopyTo(.Range("A" & intcurrow))
            '.Range("A" & intcurrow & ":O" & intcurrow).CellStyle.ColorIndex = 34
            '.Range("A" & intcurrow & ":O" & intcurrow).CellStyle.Font.Bold = True
            '.Range(intcurrow, 2).Value = "SUB Total ( " & oRow("customer") & "-" & Trim(oRow("customername")) & " )"
            'Call sumtotal(True)
            intcurrow += 1
        End With
        Call SetPageProperties()
        sheet.Remove()
        Dim filename As String
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)

    End Sub
    Protected Sub SetPageProperties()
        With sheet2
            .SetColumnWidth(1, 8)
            .SetColumnWidth(2, 13)
            .SetColumnWidth(3, 13)
            .SetColumnWidth(4, 40)
            .SetColumnWidth(5, 14)
            .SetColumnWidth(6, 20)
            .SetColumnWidth(7, 17)

            .SetColumnWidth(8, 17)
            .SetColumnWidth(9, 17)
            .SetColumnWidth(10, 17)
            .SetColumnWidth(11, 25)

            For i As Integer = 6 To 9999
                .SetRowHeight(i, 22)
            Next

            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Landscape
            .PageSetup.LeftMargin = 0.3
            .PageSetup.RightMargin = 0.3
            .PageSetup.TopMargin = 0.3
            .PageSetup.BottomMargin = 0.3
            .PageSetup.Zoom = 100
            '.PageSetup.PrintTitleRows = "$1:$5"
            '.Range("A6").FreezePanes()
        End With
    End Sub
End Class
