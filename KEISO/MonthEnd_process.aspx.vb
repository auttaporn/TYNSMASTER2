Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI
Imports Syncfusion.XlsIO

Partial Class MonthEnd_process
    Inherits System.Web.UI.Page
    Private oConn As New OleDbConnection(ClassConn.strConnSql)
    Private conn400 As New OleDbConnection(ClassConn.strCon400)
    Private Rptpath As String = Server.MapPath("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("TAP-MonthEndCheck.xlt")
    Private strDate As String = Date.Now.ToString("yyyyMMdd")
    Private strTime As String = Date.Now.ToString("HHmm")
    Private ReportName As String = "TAP Month End Report"
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private oTable As New DataTable
    Private oRow As DataRow
    Private Filename As String
    Private intcurrow As Integer = 6
    Private checksheet As String = ""
    Private library_TAPB As String = "TAPSALEF7"
    Private library_TAPC As String = "TAPSALEF"
    Private library_SILT8 As String

    Function getdata() As DataTable
        If oConn.State = ConnectionState.Closed Then
            oConn.Open()
        End If
        '--PAN 3/9/2013 hidono=left(hidono,1) in (''C'',''P'')-- TAP-P SOLUTION
        Dim strsql As String
        strsql = "select * into #a from openquery(as400,' select hifacd as facd,hicocd,hidttp as typeYIC,(CASE WHEN left(hidono,1) in (''C'',''P'') THEN right(rtrim(hidono),6)  ELSE hidono END)as hidono ,hidono as inv, hifcam as fc "
        strsql += " from ygss.yssh where   hibldt  like ''" & txtmonth.Text.Trim & "%'' and  hicocd = ''E''    order by hidttp desc , hidono ')  "

        strsql += " select * into #b from openquery(as400,' select hifacd,hicocd,hidttp as typeTAP,substring(hidono,3,10) as invno,hidono as invTAP, hivbam as amt "
        strsql += " from ygss.yssh where   hibldt  like ''" & txtmonth.Text.Trim & "%'' and  hicocd = ''A''    and hiivfg in(''D'',''K'') and hidttp not in(''C'',''D'') order by hidttp desc , hidono')  "

        Dim ocom As New OleDbCommand(strsql, oConn)
        ocom.CommandTimeout = 350
        ocom.ExecuteNonQuery()

        Dim sql As String
        sql = "select a.*,b.*,(a.fc - b.amt) as total from #a a full outer join #b b on invno=hidono and facd=hifacd where (a.fc - b.amt)  <> 0 or (a.fc - b.amt) is null "
        sql += " order by total desc ,typeYIC desc,typeTAP desc ,  invno ,hidono "

        Dim da As New OleDbDataAdapter(sql, oConn)
        Dim ds As New DataSet
        da.Fill(ds, "data")
        Return ds.Tables("data")

        oConn.Close()
    End Function
    Function getcancel() As DataTable
        Dim query As String

        If oConn.State = ConnectionState.Closed Then
            oConn.Open()
        End If

        query = "select *  from openquery(as400,' select hifacd as facd,hicocd,hidttp as typeYIC,(CASE WHEN left(hidono,1) = ''C'' THEN right(rtrim(hidono),6)  ELSE hidono END)as hidono ,hidono as inv, hifcam as fc "
        query += " from ygss.yssh where   hibldt  like ''" & txtmonth.Text.Trim & "%'' and  hicocd = ''E'' and HICAFG = ''Y''  order by hidttp desc , hidono ')  "

        Dim da As New OleDbDataAdapter(query, oConn)
        Dim ds As New DataSet
        da.Fill(ds, "data")
        Return ds.Tables("data")
    End Function

    Protected Sub deleteinv(ByVal fac As String, ByVal invno As String)
        Dim sql As String
        If fac = "32T1" Or fac = "32T5" Then
            fac = library_TAPB
        ElseIf fac = "32T2" Then
            fac = library_TAPC
        End If

        If conn400.State = ConnectionState.Closed Then
            conn400.Open()
        End If

        sql = " delete from " & fac & ".siht8 where stinvn='" & invno & "' "

        ''Dim ocom As New OleDbCommand(sql, conn400)
        ''ocom.ExecuteNonQuery()
        sql = " delete from " & fac & ".silt8 where tlinvn='" & invno & "' "


    End Sub

    Protected Sub btnEnter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnter.Click
        If txtmonth.Text = "" Then
            lblmsg.Text = "Please Input Month!!!"
            Exit Sub
        ElseIf txtmonth.Text.Trim.Length <> 6 Then
            lblmsg.Text = "Please Input Month Inform YYYYMM !!!"
            Exit Sub
        ElseIf Left(txtmonth.Text.Trim, 2) <> "20" Then
            lblmsg.Text = "Please Input Year Inform 20XX !!!"
            Exit Sub
        End If

        oTable = getdata()
        If oTable.Rows.Count <= 0 Then
            lblmsg.Text = "No Data!!"
            Exit Sub
        End If
      
        Call PrintData()
        Call printcancel()

    
        sheet.Remove()
        Filename = ReportName & "-" & strDate & "-" & strTime & ".xls"
        workbook.SaveAs(Filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
        Response.End()
    End Sub
    Sub printcancel()
        Dim dtcancel As New DataTable
        Dim drow As DataRow

        dtcancel = getcancel()

        If dtcancel.Rows.Count <= 0 Then
            Exit Sub
        End If

        sheet2 = workbook.Worksheets.Create("Cancel YIC Invoice")
        sheet.Range("A1:D5").CopyTo(sheet2.Range("A1"))
        sheet2.Range(1, 2).Text = "Cancel YIC Invoice"
        sheet2.Range(2, 4).Text = Right(txtmonth.Text.Trim, 2) & "/" & Left(txtmonth.Text.Trim, 4)
        sheet2.Range(3, 6).Text = "Print Date : " & Now()
        intcurrow = 6
        Call SetPageProperties()

        For Each drow In dtcancel.Rows
            sheet.Range("A7:D7").CopyTo(sheet2.Range(intcurrow, 1))
            sheet2.Range(intcurrow, 2).Text = drow("FACD")      'Factory
            sheet2.Range(intcurrow, 3).Text = drow("inv")       'YIC INVOICE
            sheet2.Range(intcurrow, 4).Value = drow("FC")       'YIC FC
            intcurrow += 1
        Next

    End Sub
    Sub PrintData()

        appExcel.DefaultFilePath = Server.MapPath(".")
        workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = workbook.Worksheets(0)

        For Each oRow In oTable.Rows

            If oRow("total") IsNot System.DBNull.Value Then   '*****case total invoice diff > 0 ****
                printdiff()
                If oRow("total") <> 0 Then '******** DELETE TAP INVOICE ( SILT8 )
                    deleteinv(Trim(oRow("facd")), Trim(oRow("invno")))
                End If
            Else
                If oRow("typeTAP") IsNot System.DBNull.Value Then

                    If Left(Trim(oRow("invtap")), 3) = "TTT" Then    '***** INVOICE TYE ***********
                        printinvTYE()
                    Else
                        If Trim(oRow("typeTAP")) <> "I" Then    '***** TAP Credit & Debit ***********
                            TAPcredit()
                        Else
                            invtap()
                        End If
                    End If

                Else
                    If Trim(oRow("typeYIC")) <> "I" Then   '***** YIC Credit & Debit ***********
                        YICcredit()
                    Else
                        'Dim CNOREC2 As Int16
                        'CNOREC2 = CHKFCTAP()

                        'If CNOREC2 = 0 Then
                        invyic()
                        'Else
                        'NOFC()
                        'End If
                    End If

                End If
            End If
        Next
    End Sub
    'Sub NOFC() ' PAN 3/9/2013

    '    Dim SQLFC As String
    '    SQLFC = "select TLINVN,TLPROD,TRIM(TLMKSF)||TLYZSF AS SUFFIX,TLTA10 AS FC,TLDATE  from tapsalef7.silt8 where tlinvn='" & oRow("inv") & "' AND TLTA10=0"
    '    Dim DA As New OleDbDataAdapter(SQLFC, conn400)
    '    Dim DT As New DataTable
    '    DA.Fill(DT)
    '    Dim DR As DataRow

    '    For Each DR In DT.Rows
    '        If checksheet <> "NO FC TAP" Then
    '            sheet2 = workbook.Worksheets.Create("NO FC TAP")
    '            sheet.Range("A11:G12").CopyTo(sheet2.Range("A4"))
    '            sheet2.Range(1, 2).Text = "NO FC TAP"
    '            sheet2.Range(2, 4).Text = Right(txtmonth.Text.Trim, 2) & "/" & Left(txtmonth.Text.Trim, 4)
    '            sheet2.Range(3, 6).Text = "Print Date : " & Now()
    '            checksheet = "NO FC TAP"
    '            intcurrow = 6
    '            Call SetPageProperties()
    '        End If
    '        sheet.Range("A13:G13").CopyTo(sheet2.Range(intcurrow, 1))
    '        sheet2.Range(intcurrow, 2).Text = oRow("FACD")      'Factory
    '        sheet2.Range(intcurrow, 3).Text = DR("TLINVN")       'TAP INVOICE
    '        sheet2.Range(intcurrow, 4).Value = DR("TLPROD")       'PART NO.
    '        sheet2.Range(intcurrow, 5).Text = DR("SUFFIX")    'SUFFIX
    '        sheet2.Range(intcurrow, 6).Value = DR("FC")      'FC
    '        sheet2.Range(intcurrow, 7).Value = DR("TLDATE")    'DATE

    '        intcurrow += 1
    '    Next
    'End Sub
    'Function CHKFCTAP() ' PAN 3/9/2013
    '    Dim CNOREC As Int16
    '    If Trim(oRow("FACD")).ToString = "32T1" Then
    '        library_SILT8 = "tapsalef7.silt8"
    '    ElseIf Trim(oRow("FACD")).ToString = "32T2" Then
    '        library_SILT8 = "tapsalef.silt8"
    '    ElseIf Trim(oRow("FACD")).ToString = "32T3" Then
    '        library_SILT8 = "tapsalefP.silt8"
    '    End If

    '    Dim SQLCHK As String
    '    SQLCHK = "SELECT COUNT(*) FROM " & library_SILT8 & " WHERE TLINVN='" & oRow("inv") & "' AND TLTA10=0"

    '    Dim CMD As New OleDbCommand(SQLCHK, conn400)
    '    CNOREC = CMD.ExecuteScalar()

    '    Return CNOREC

    'End Function
    Sub printdiff()
        If checksheet = "" Then
            sheet2 = workbook.Worksheets.Create("Compare Invoice")
            sheet.Range("A1:G5").CopyTo(sheet2.Range("A1"))
            sheet2.Range(2, 4).Text = Right(txtmonth.Text.Trim, 2) & "/" & Left(txtmonth.Text.Trim, 4)
            sheet2.Range(3, 6).Text = "Print Date : " & Now()
            checksheet = "Compare Invoice"
            intcurrow = 6
            Call SetPageProperties()
        End If
        sheet.Range("A7:G7").CopyTo(sheet2.Range(intcurrow, 1))
        sheet2.Range(intcurrow, 2).Text = oRow("FACD")      'Factory
        sheet2.Range(intcurrow, 3).Text = oRow("inv")       'YIC INVOICE
        sheet2.Range(intcurrow, 4).Value = oRow("FC")       'YIC FC
        sheet2.Range(intcurrow, 5).Text = oRow("invTAP")    'TAP INVOICE
        sheet2.Range(intcurrow, 6).Value = oRow("amt")      'TAP AMOUNT
        sheet2.Range(intcurrow, 7).Value = oRow("total")    'DIFF
        intcurrow += 1
    End Sub
    Sub invtap()
        If checksheet = "" Then
            sheet2 = workbook.Worksheets.Create("Compare Invoice")
            sheet.Range("A1:G5").CopyTo(sheet2.Range("A1"))
            sheet2.Range(2, 4).Text = Right(txtmonth.Text.Trim, 2) & "/" & Left(txtmonth.Text.Trim, 4)
            sheet2.Range(3, 6).Text = "Print Date : " & Now()
            checksheet = "Compare Invoice"
            intcurrow = 6
            Call SetPageProperties()
        End If
        sheet.Range("A7:G7").CopyTo(sheet2.Range(intcurrow, 1))
        sheet2.Range(intcurrow, 2).Text = oRow("HIFACD")      'Factory
        sheet2.Range(intcurrow, 5).Text = oRow("invTAP")    'TAP INVOICE
        sheet2.Range(intcurrow, 6).Value = oRow("amt")      'TAP AMOUNT
        sheet2.Range(intcurrow, 7).Value = oRow("amt")      'DIFF
        intcurrow += 1
    End Sub
    Sub invyic()
        If checksheet = "" Then
            sheet2 = workbook.Worksheets.Create("Compare Invoice")
            sheet.Range("A1:G5").CopyTo(sheet2.Range("A1"))
            sheet2.Range(2, 4).Text = Right(txtmonth.Text.Trim, 2) & "/" & Left(txtmonth.Text.Trim, 4)
            sheet2.Range(3, 6).Text = "Print Date : " & Now()
            checksheet = "Compare Invoice"
            intcurrow = 6
            Call SetPageProperties()
        End If
        sheet.Range("A7:G7").CopyTo(sheet2.Range(intcurrow, 1))
        sheet2.Range(intcurrow, 2).Text = oRow("FACD")      'Factory
        sheet2.Range(intcurrow, 3).Text = oRow("inv")       'YIC INVOICE
        sheet2.Range(intcurrow, 4).Value = oRow("FC")       'YIC FC
        sheet2.Range(intcurrow, 7).Value = oRow("FC")       'DIFF
        intcurrow += 1
    End Sub
    Sub printinvTYE()
        If checksheet <> "Invoice TYE" Then
            sheet2 = workbook.Worksheets.Create("Invoice TYE")
            sheet.Range("A1:D5").CopyTo(sheet2.Range("A1"))
            sheet2.Range(1, 2).Text = "Invoice TYE"
            sheet2.Range(2, 4).Text = Right(txtmonth.Text.Trim, 2) & "/" & Left(txtmonth.Text.Trim, 4)
            sheet2.Range(3, 6).Text = "Print Date : " & Now()
            sheet2.Range(3, 2).Text = "TYE INNOICE"
            sheet2.Range(4, 3).Text = "TYE INVOICE"
            sheet2.Range(4, 4).Text = "AMOUNT"
            checksheet = "Invoice TYE"
            intcurrow = 6
            Call SetPageProperties()
        End If
        sheet.Range("A7:D7").CopyTo(sheet2.Range(intcurrow, 1))
        sheet2.Range(intcurrow, 2).Text = oRow("HIFACD") 'Factory
        sheet2.Range(intcurrow, 3).Text = oRow("invTAP")    'TAP INVOICE
        sheet2.Range(intcurrow, 4).Text = oRow("AMT")     'TAP PRICE
        intcurrow += 1
    End Sub
    Sub TAPcredit()
        If checksheet <> "Credit & Debit" Then
            sheet2 = workbook.Worksheets.Create("Credit & Debit")
            sheet.Range("A1:E5").CopyTo(sheet2.Range("A1"))
            sheet2.Range(1, 2).Text = "Credit & Debit"
            sheet2.Range(2, 4).Text = Right(txtmonth.Text.Trim, 2) & "/" & Left(txtmonth.Text.Trim, 4)
            sheet2.Range(3, 6).Text = "Print Date : " & Now()
            sheet2.Range(3, 2).Text = "TAP INNOICE"
            sheet2.Range(4, 3).Text = "INVOICE"
            sheet2.Range(4, 4).Text = "AMOUNT"
            sheet2.Range(4, 5).Text = "TYPE"
            checksheet = "Credit & Debit"
            intcurrow = 6
            Call SetPageProperties()
        End If
        sheet.Range("A7:E7").CopyTo(sheet2.Range(intcurrow, 1))
        sheet2.Range(intcurrow, 2).Text = oRow("HIFACD")    'Factory
        sheet2.Range(intcurrow, 3).Text = oRow("invTAP")    'TAP INVOICE
        sheet2.Range(intcurrow, 4).Text = oRow("AMT")       'AMOUNT
        sheet2.Range(intcurrow, 5).Text = oRow("typeTAP")   'TYPE Credit Debit
        intcurrow += 1
    End Sub
    Sub YICcredit()
        If checksheet = "Credit & Debit" Then
            intcurrow += 1
            sheet.Range(3, 2).CopyTo(sheet2.Range(intcurrow, 2))
            sheet2.Range(intcurrow, 2).Text = "YIC INNOICE"
            intcurrow += 1
            sheet.Range("A4:E5").CopyTo(sheet2.Range(intcurrow, 1))
            sheet2.Range(intcurrow, 3).Text = "INVOICE"
            sheet2.Range(intcurrow, 5).Text = "TYPE"
            intcurrow += 2
            checksheet = "Credit & DebitYIC"
        ElseIf checksheet <> "Credit & DebitYIC" And checksheet <> "Credit & Debit" Then
            sheet2 = workbook.Worksheets.Create("Credit & Debit")
            sheet.Range("A1:E5").CopyTo(sheet2.Range("A1"))
            sheet2.Range(1, 2).Text = "Credit & Debit"
            sheet2.Range(2, 4).Text = Right(txtmonth.Text.Trim, 2) & "/" & Left(txtmonth.Text.Trim, 4)
            sheet2.Range(3, 6).Text = "Print Date : " & Now()
            sheet2.Range(3, 2).Text = "YIC INNOICE"
            sheet2.Range(4, 3).Text = "INVOICE"
            sheet2.Range(4, 5).Text = "TYPE"
            checksheet = "Credit & DebitYIC"
            intcurrow = 6
            Call SetPageProperties()
        End If
        sheet.Range("A7:E7").CopyTo(sheet2.Range(intcurrow, 1))
        sheet2.Range(intcurrow, 2).Text = oRow("FACD")      'Factory
        sheet2.Range(intcurrow, 3).Text = oRow("inv")       'YIC INVOICE
        sheet2.Range(intcurrow, 4).Text = oRow("FC")        'YIC FC
        sheet2.Range(intcurrow, 5).Text = oRow("typeYIC")   'Credit Debit
        intcurrow += 1
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If Session("USERID") = "" Then Response.Redirect("Login.aspx")
        If Not Page.IsPostBack Then
            'Dim checkpermis As Integer
            'checkpermis = ClassMaster.getpermis("YICMAIN", Session("USERID"))
            'If checkpermis = 0 Then
            '    Response.Redirect("~\NotAuthority.aspx")
            'End If
            txtmonth.Text = Now.Year.ToString & Right("00" & (Now.Month.ToString - 1), 2)
        End If
        If chkview.SelectedIndex = 0 Then
            MultiView1.SetActiveView(ViewMonthEnd)
        ElseIf chkview.SelectedIndex = 1 Then
            MultiView1.SetActiveView(ViewPRDCD)

        End If


    End Sub

    Protected Sub SetPageProperties()
        With sheet2
            .SetColumnWidth(1, 0.5)                       'COLUMN A  -- 
            .SetColumnWidth(2, 10)                        'COLUMN B  -- Factory
            .SetColumnWidth(3, 15)                        'COLUMN C  -- invoice  
            .SetColumnWidth(4, 12)                        'COLUMN D  -- price,FC
            .SetColumnWidth(5, 13)                        'COLUMN E  -- invoice,credit & debit
            .SetColumnWidth(6, 12)                        'COLUMN F  -- price
            .SetColumnWidth(7, 11)                        'COLUMN G  -- diff

        End With
    End Sub

    Protected Sub btncancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btncancel.Click
        lblmsg.Text = ""
        txtmonth.Text = ""
    End Sub

    Protected Sub btnPRDCD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPRDCD.Click
        Dim ptable As New DataTable
        ptable = getPRDCD()
        If ptable.Rows.Count > 0 Then
            gdvPRDCD.DataSource = ptable
            gdvPRDCD.DataBind()
        Else
            lblerr.Text = "NO DATA."

        End If


    End Sub
    Function getPRDCD()
        If oConn.State = ConnectionState.Closed Then
            oConn.Open()
        End If
        Dim sql As String
        sql = "select * from openquery(as400,'select  distinct(tlprod),tlpdcd,pprdcd,''TAP-B''as fac,tldate+20000000 as tldate,stcust,cnme  "
        sql += " from tapsalef7.silt8 inner join tapsalef7.siht8 on stinvn=tlinvn  "
        sql += " inner join vptapbd7.trcm on ccust=stcust  "
        sql += " left outer join ygss.yprd on tlprod=pprod "
        sql += " where tldate >= ''" & Right(txtfrm.Text.Trim, 6) & "'' and tldate <= ''" & Right(txtto.Text.Trim, 6) & "'' and tlpdcd=''''   and stcanc <> ''Y'' order by fac,tlpdcd,tlprod')"
        sql += "  union "
        sql += " select * from openquery(as400,'select  distinct(tlprod),tlpdcd,pprdcd,''TAP-C''as fac,tldate+20000000 as tldate,stcust,cnme  "
        sql += " from tapsalef.silt8 inner join tapsalef.siht8 on stinvn=tlinvn  "
        sql += " inner join vptapbd.trcm on ccust=stcust  "
        sql += " left outer join ygss.yprd on tlprod=pprod "
        sql += " where tldate >= ''" & Right(txtfrm.Text.Trim, 6) & "'' and tldate <= ''" & Right(txtto.Text.Trim, 6) & "'' and tlpdcd=''''   and stcanc <> ''Y'' order by fac,tlpdcd,tlprod')"
   sql += "  union "
        sql += " select * from openquery(as400,'select  distinct(tlprod),tlpdcd,pprdcd,''TAP-P''as fac,tldate+20000000 as tldate,stcust,cnme  "
        sql += " from tapsalefp.silt8 inner join tapsalefp.siht8 on stinvn=tlinvn  "
        sql += " inner join vptapbdp.trcm on ccust=stcust  "
        sql += " left outer join ygss.yprd on tlprod=pprod "
        sql += " where tldate >= ''" & Right(txtfrm.Text.Trim, 6) & "'' and tldate <= ''" & Right(txtto.Text.Trim, 6) & "'' and tlpdcd=''''   and stcanc <> ''Y'' order by fac,tlpdcd,tlprod')"

        Dim da As New OleDbDataAdapter(sql, oConn)
        Dim ds As New DataSet
        da.Fill(ds, "data")
        Return ds.Tables("data")

        '  da.SelectCommand.CommandTimeout = 120


        oConn.Close()
    End Function

    Protected Sub btnPRDCDcancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPRDCDcancel.Click
        lblerr.Text = ""
        txtfrm.Text = ""
        txtto.Text = ""
        oConn.Close()
    End Sub
End Class


