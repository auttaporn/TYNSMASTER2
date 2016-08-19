Imports System.Data
Imports System.Data.OleDb
Imports Syncfusion.XlsIO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Web.UI
Imports CommonFunction
Partial Class SALEDO_SUMMARYREPORT02
    Inherits System.Web.UI.Page
    Private conn As New OleDbConnection(Classconn.strConnSql)
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private TemplateFile As String = CurReportPath & ("R2-REPORT.xlt")
    Private ReportName As String = "DOMESTIC-REPORT2"
    Private otable As DataTable
    Private dt As New DataTable
    Private idx As Integer = 0
    Private drow As DataRow
    Private orow As DataRow
    Private intCurrRow, intNext As Integer
    Private strSale As String
    Private Crtdate As String = Val(Date.Now.ToString("yyyyMMdd"))
    Private filename As String
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private Workbook As IWorkbook
    Private sheet, sheet2 As IWorksheet

    Protected Sub SetPlant()
        Dim oTable As New DataTable
        Dim oRow As DataRow
        oTable = ClassFunctionVar.GetDataUser(Session("USERID"), "MTAPDADR02")
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
    Private Function GetData() As DataTable
        Dim condition, sql As String

        condition = "  HICAFG <> 'Y' AND HIIVFG = 'D' "
        If (ddlCompany.SelectedValue <> "ALL") And (ddlCompany.SelectedValue <> "") Then
            condition += " AND HICOCD = '" & Trim(ddlCompany.SelectedValue) & "' "
        End If
        If txtFrom.Text <> "" Then
            condition = condition & " AND HIBLDT >= '" & Right(txtFrom.Text, 4) & Mid(txtFrom.Text, 3, 2) & Left(txtFrom.Text, 2) & "' "
        End If
        If txtTo.Text <> "" And Left(txtTo.Text, 2) <= "31" Then
            condition = condition & " AND HIBLDT <= '" & Right(txtTo.Text, 4) & Mid(txtTo.Text, 3, 2) & Left(txtTo.Text, 2) & "' "
        End If
        condition = condition & "AND HICOCD = '" & Trim(ddlCompany.SelectedValue) & "' AND DISQTY <> 0 AND CTL1.CTYPE = 'cus' "
        If ddlBillto.SelectedValue <> "0" Then
            condition += "AND HIRPFG='" & ddlBillto.SelectedValue & "' "
        End If
        sql = "SELECT HIBTCD AS Cus, DIPDGP, "
        sql += "(ISNULL(SUM((CASE WHEN H.HIDTTP = 'I' THEN DISPAM  ELSE 0 END)"
        sql += "+ (CASE WHEN H.HIDTTP = 'D' THEN D.DISPAM  ELSE 0 END)"
        sql += "- (CASE WHEN H.HIDTTP = 'C' THEN D.DISPAM  ELSE 0 END)), 0))AS DISPAM, "
        sql += "(ISNULL(SUM((CASE WHEN H.HIDTTP = 'I' THEN D .DIFCAM ELSE 0 END) "
        sql += "+ (CASE WHEN H.HIDTTP = 'D' THEN D.DIFCAM ELSE 0 END) "
		sql += "- (CASE WHEN H.HIDTTP = 'C' THEN D.DIFCAM ELSE 0 END)), 0)) AS FC, "
        'sql += "- (CASE WHEN H.HIDTTP = 'C' THEN D.DIFCAM ELSE 0 END)), 0)) AS FC, "
        sql += "SUM(DISPAM) - SUM(Difcam) AS PROFIT, "
        sql += "ISNULL(CTL1.CNAME, '-') AS CUSNAME " 
        sql += "FROM  YSSH H "
        sql += "INNER JOIN YSSD D ON HIDONO = DIDONO  AND HICOCD = DICOCD "
        sql += "LEFT OUTER JOIN YSCTL CTL1 ON HIBTCD = CTL1.CKEY "
'        sql += "WHERE " & condition & " AND HIFACD = '32T1'  "
        sql += "WHERE " & condition & "  "
        sql += "GROUP BY HIBTCD, CTL1.CNAME, DIPDGP "
        sql += "ORDER BY HIBTCD, DIPDGP "
       ' RESPONSE.WRITE(sql)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Dim da As New OleDbDataAdapter(sql, conn)
        Dim ds As New DataSet
        da.Fill(ds, "SaleSummary")
        Return ds.Tables("SaleSummary")
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Crtdate = Right(Crtdate, 2) & Mid(Crtdate, 5, 2) & Left(Crtdate, 4)
        If Session("userid") = "" Then
            Response.Redirect("~/")
        End If

        If Not Page.IsPostBack Then
            txtFrom.Text = Crtdate
            txtTo.Text = Crtdate
            Call SetPlant()
            MultiView1.SetActiveView(ViewCondition)
            ViewCondition.DataBind()
        End If
    End Sub

    Protected Sub btnsubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmit.Click
        Call WriteData()
    End Sub

    Protected Sub PrintH()
        sheet.Range("A1:F9").CopyTo(sheet2.Range("A1"))
        sheet2.Range(2, 1).Value = CStr(Left(txtFrom.Text, 2)) & "/" & CStr(Mid(txtFrom.Text, 3, 2)) & "/" & CStr(Right(txtFrom.Text, 4)) & _
                                  "  -  " & CStr(Left(txtTo.Text, 2)) & "/" & CStr(Mid(txtTo.Text, 3, 2)) & "/" & CStr(Right(txtTo.Text, 4))
        If ddlCompany.SelectedValue = "E" Then
            sheet2.Range(4, 2).Value = "YIC Asia Pacific Corporation Ltd."
        Else
            sheet2.Range(4, 2).Value = "Thai Arrow Products Co.,Ltd."
        End If
        sheet2.Range(3, 1).Value = "DOMESTIC"
        sheet2.Range(4, 6).Value = "Printdate  :  " & CStr(Now())
        sheet2.Range(5, 6).Value = "Print by     :  " & Session("userId") & " - " & Session("username")
    End Sub
 
    Sub WriteData()
        If txtFrom.Text = "" Or txtTo.Text = "" Then
            lblmessage.Text = "Please input data"
            Exit Sub
        End If
        lblmessage.Text = ""

        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create("REPORT No. 2")
        otable = GetData()

        If otable.Rows.Count < 1 Then
            lblmessage.Text = "NO DATA..."
            Exit Sub
        Else
            Call SetPageProperties()
            intNext = 10
            Call PrintH()
            Call PrintDetail()
        End If
    End Sub
    Protected Sub PrintDetail()
        Dim currCus As String
        Dim Count As Integer = 1
        intCurrRow = 7
        orow = otable.Rows(0)
        currCus = Trim(orow("Cus"))
        'sheet2.Name = "REPORT No. 2"
        strSale = "0|7"

        For Each orow In otable.Rows
            If currCus <> Trim(orow("Cus")) Then
                sheet.Range("A7:F9").CopyTo(sheet2.Range("A" & intNext))
                intCurrRow = intNext
                intNext = intNext + 3
                currCus = Trim(orow("Cus"))
                strSale = strSale & "|" & intCurrRow
            End If
            sheet2.Range(intCurrRow, 1).Value = Trim(orow("Cus")) '& "::   " & Trim(orow("Cusname"))
            sheet2.Range(intCurrRow, 2).Value = ":: " & Trim(orow("Cusname"))
            Select Case Trim(orow("dipdgp"))
                Case "01"
                    sheet2.Range(intCurrRow, 4).Value = Trim(orow("DISPAM"))
                    sheet2.Range(intCurrRow + 1, 4).Value = CDbl(Trim(orow("DISPAM"))) - CDbl(Trim(orow("FC")))
                Case "02"
                    sheet2.Range(intCurrRow, 5).Value = Trim(orow("DISPAM"))
                    sheet2.Range(intCurrRow + 1, 5).Value = CDbl(Trim(orow("DISPAM"))) - CDbl(Trim(orow("FC")))
            End Select
        Next
        intCurrRow = intNext
        Call PrintTotal()
        Call CheckFFileName()
        GridView2.DataSource = dt
        GridView2.DataBind()
        If GridView2.Rows.Count = 0 Then
            lblmessage.Text = "Message : No Data"
            MultiView1.SetActiveView(ViewCondition)
        Else
            ViewOpen.DataBind()
            MultiView1.SetActiveView(ViewOpen)
            lblMsg.Text = "Save " & GridView2.Rows.Count & " Files Completed"
        End If
    End Sub
    Protected Sub PrintTotal()
        Dim ab() As String = Split(strSale, "|")
        Dim i As Integer
        With sheet2
            sheet.Range("A10:F12").CopyTo(.Range("A" & intNext))
            .Range(intCurrRow, 2).Value = "DOMESTIC   ::  GRAND TOTAL "
            For i = 1 To UBound(ab)
                .Range("D" & intCurrRow).Number = CDbl(.Range("D" & intCurrRow).Value) + CDbl(.Range(CInt(ab(i)), 4).Value)
                .Range("D" & intCurrRow + 1).Number = CDbl(.Range("D" & intCurrRow + 1).Value) + CDbl(.Range(CInt(ab(i)) + 1, 4).Value)
                .Range("E" & intCurrRow).Number = CDbl(.Range("E" & intCurrRow).Value) + CDbl(.Range(CInt(ab(i)), 5).Value)
                .Range("E" & intCurrRow + 1).Number = CDbl(.Range("E" & intCurrRow + 1).Value) + CDbl(.Range(CInt(ab(i)) + 1, 5).Value)
            Next
            .Range("F" & intCurrRow).Number = CDbl(.Range("D" & intCurrRow).Number) + CDbl(.Range("E" & intCurrRow).Number)
            .Range("F" & intCurrRow + 1).Number = CDbl(.Range("D" & intCurrRow + 1).Number) + CDbl(.Range("E" & intCurrRow + 1).Number)
        End With
    End Sub

    Protected Sub CheckFFileName()
        filename = ReportName & "-" & ClassFunctionVar.CRTDT & ".xls"
        CommonFunction.SaveTmpFile(CurTempPath, filename)
        filename = GetFileName(filename)
        filename = CurTempPath & filename
        sheet.Remove()
        workbook.SaveAs(filename)
        workbook.Close(True)
    End Sub

    Protected Sub SetPageProperties()
        With sheet2
            .SetColumnWidth(1, 6)                   'COLUMN A  -- MAKER CODE
            .SetColumnWidth(2, 45)                  'COLUMN B  -- CUSTOMER
            .SetColumnWidth(3, 12.43)               'COLUMN C  -- DATA TYPE (SALE,PROFIT,%)
            .SetColumnWidth(4, 19.5)                'COLUMN D  -- BUHIN TOTAL
            .SetColumnWidth(5, 19.5)                'COLUMN E  -- KEIKI TOTAL
            .SetColumnWidth(6, 19.5)                'COLUMN F  -- GRAND TOTAL
            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Portrait
            .PageSetup.LeftMargin = 0.18
            .PageSetup.RightMargin = 0.16
            .PageSetup.TopMargin = 0.22
            .PageSetup.BottomMargin = 0.22
            .PageSetup.Zoom = 65
            .PageSetup.PrintTitleRows = "$1:$6"
            .Range("A7").FreezePanes()
        End With
    End Sub
  
    Private Function GetFileName(ByRef Filename As String) As String
        Call CommonFunction.CreateFileNameSession(idx, dt, drow, Filename)
        Return Filename
    End Function

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        Response.Redirect("~\download.aspx?file=" & GridView2.SelectedValue & "&path=" & DefaultPath)
    End Sub
End Class
