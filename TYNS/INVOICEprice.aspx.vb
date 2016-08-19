Imports System.Data
Imports System.Data.OleDb
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine

Partial Class TAPPRG_INVOICEprice
    Inherits System.Web.UI.Page
    Private Conn As New OleDbConnection(Classconn.strConnSql)
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/rpt/ExcelReport/")
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private TemplateFile As String
    Private ReportName As String = "INVOICEprice"
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private sheet3 As IWorksheet
    Private filename As String
    Private Crtdate As String = Val(Date.Now.ToString("ddmmyyyy"))
    Private crtTime As String = Val(Date.Now.ToString("HHmm"))
    Private intCurrow As Integer
    Private oTable, oTable2 As DataTable
    Private orow, orow2 As DataRow
    Private recchk As Integer
    Dim library_fac As String
    Dim YM
    Protected Function getdataisuzu() As DataTable

        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If

        Dim strSql As String = ""

        strSql = "Select * From openquery(as400,'Select ilprod,ilmksf,ilyzsf,sum(ilqty) as ilqty from  " & library_fac & ".sil "
        strSql += "Inner Join " & library_fac & ".sih On ilinvn = siinvn inner JOIN " & library_fac & ".ech ON ilinvn = hinvn "
        strSql += " left join ygss.tcust on tycust = sicust  "
        strSql += "Where ildate <= " & YM & "31 And ildate >= " & YM & "01  and ilcusb=''110201'' "
        strSql += "and siinvn not in (select RInvc from " & library_fac & ".rar where rnxt = ''2'')"
        strSql += "group by ilprod,ilmksf,ilyzsf order by ilprod')"

        Dim da As New OleDbDataAdapter(strSql, Conn)
        Dim ds As New DataSet
        da.Fill(ds, "detailisuzu")
        Return ds.Tables("detailisuzu")

    End Function


    Protected Function getdata() As DataTable

        Dim strSql As String = ""

        If Conn.State = ConnectionState.Open Then
            Conn.Close()
        End If
        Conn.Open()

        If (ddlCompany.SelectedValue = "TAP-B") Then 'PAN 16/8/2013
            library_fac = "tapsalef7"
        ElseIf (ddlCompany.SelectedValue = "TAP-C") Then
            library_fac = "tapsalef"
        ElseIf (ddlCompany.SelectedValue = "TAP-P") Then
            library_fac = "tapsalefp"
        End If

        YM = Right(Trim(txtFrYear.Text), 2) & Right("00" & (ddlMonth.SelectedValue), 2)

        Select Case ddlreport.SelectedValue

            Case "non"
                strSql += "Select * From openquery(as400,'Select Distinct sicust,ilprod,ilmksf,ilyzsf,ilpdcd,hdtyp,ttcust "
                If (cbinvoiceno.Checked) Then
                    strSql += ",siinvn ,20000000+siinvd as siinvd"
                End If
                strSql += " From " & library_fac & ".sil "
                strSql += "Inner Join " & library_fac & ".sih On ilinvn = siinvn inner JOIN " & library_fac & ".ech ON ilinvn = hinvn "
                '**** pui update 10/1/2012 check customer code ******
                strSql += " left join ygss.tcust on tycust = sicust  "
                '********************* end 10/1/2012 ****************
                strSql += "Where siinvd <= " & YM & "31 And siinvd >= " & YM & "01 And ilta10 =0 "
                strSql += "and siinvn not in (select RInvc from " & library_fac & ".rar where rnxt = ''2'')"
                strSql += "ORDER BY ilpdcd')"
            Case "price"
                strSql += "select * from  openquery(as400,'select distinct ilprod,ilmksf,ilyzsf,ilta10,ilpdcd,hdtyp "
                strSql += "from " & library_fac & ".sil inner JOIN " & library_fac & ".ech ON ilinvn = hinvn  "
                strSql += " where ildate >= " & YM & "01 and ildate <= " & YM & "31 and ilta10 > 0 order by ilpdcd')"
            Case "FC"
                strSql += "Select * From openquery(as400,'Select Distinct tlprod,tlmksf,tlyzsf,tlpdcd,hdtyp,tlcusb "
                If (cbinvoiceno.Checked) Or ddlCompany.SelectedValue = "TAP-C" Then
                    strSql += ",stinvn ,20000000+stinvd as stinvd"
                End If
                strSql += " From " & library_fac & ".silt8 "
                strSql += "Inner Join " & library_fac & ".siht8 On tlinvn = stinvn inner JOIN " & library_fac & ".ech ON tlinvn = hinvn "
                strSql += "Where stinvd <= " & YM & "31 And stinvd >= " & YM & "01 And tlta10 =0 "
                strSql += "and stcanc <> ''Y'' and  tlpdcd <> ''231'' "
                strSql += "ORDER BY tlpdcd')"

        End Select


        ' response.write(strSql)
        'response.end()

        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If

        Dim da As New OleDbDataAdapter(strSql, Conn)
        Dim ds As New DataSet
        '  da.SelectCommand.CommandTimeout = 120
        da.Fill(ds, "detail")
        Return ds.Tables("detail")

        If Conn.State = ConnectionState.Open Then  '********** update 28/3/2012 ************
            Conn.Close()
        End If
    End Function
    'Protected Sub NOFC()
    '    Dim STRSQL As String
    '    Dim iRec As Integer = 0

    '    oTable = getdata()

    '    If oTable.Rows.Count = 0 Then
    '        lblmessage.Text = "NO DATA"
    '        Exit Sub
    '    End If

    '    Dim dtshow As New DataTable
    '    Dim drshow As DataRow
    '    dtshow.Columns.Add("Invoice Part No")
    '    dtshow.Columns.Add("Part No")
    '    dtshow.Columns.Add("MK Suffix")
    '    dtshow.Columns.Add("YZ Suffix")
    '    dtshow.Columns.Add("Product code")
    '    dtshow.Columns.Add("ADD FC")

    '    Dim strDate As String = Now.Year.ToString & Right("00" & Now.Month.ToString, 2) & Right("00" & Now.Day.ToString, 2)
    '    For Each orow In oTable.Rows

    '        drshow = dtshow.NewRow
    '        drshow("Invoice Part No") = Trim(orow("tlprod"))

    '        '***JB,PRESS,RUBBER,INJ
    '        If Trim(orow("tlpdcd")) <> "210" And Trim(orow("tlpdcd")) <> "220" And Trim(orow("tlpdcd")) <> "270" Then
    '            If Trim(orow("tlprod")).Length > 8 Then
    '                drshow("Part no") = Left(Trim(orow("tlprod")), 4) & "-" & Mid(Trim(orow("tlprod")), 5, 4) & "-" & Mid(Trim(orow("tlprod")), 9, 2)
    '            Else
    '                drshow("Part no") = Left(Trim(orow("tlprod")), 4) & "-" & Mid(Trim(orow("tlprod")), 5, 4)
    '            End If

    '        End If

    '        '*** TUBE *****
    '        If Trim(orow("tlpdcd")) = "270" Then
    '            drshow("Part no") = "450" & Trim(orow("tlprod"))
    '        End If


    '        '*** AW *****
    '        If Trim(orow("tlpdcd")) = "210AW" Then       'AW
    '            STRSQL = "INSERT INTO YIC_VDMY "
    '            STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T4'',''1996'',rtrim(ntypat)AS ntypat,''210'',''999999'','''','''',"
    '            STRSQL += " nttfdt,ntaspr/10000 AS ntaspr,''BAH'',''0'',''KG'',''0''," & strDate & ",''" & Session("USERID") & "'',"
    '            STRSQL += "''0'','''',''0'','''',''0'','''',''0''"
    '            STRSQL += " FROM  "
    '            STRSQL += " vc1tapaw.vcnt "
    '            STRSQL += " WHERE ntac = ''FCDOM'' AND nttfdt >= ''" & Left(strDate, 6) & "01'' AND  SUBSTRING(NTYPAT,4,11) =''" & Trim(orow("tlprod")) & "'' ') "
    '            Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
    '            Try
    '                ''   iRec = oCom.ExecuteNonQuery()
    '            Catch ex As Exception
    '                iRec = -1
    '                lblmessage.Text = STRSQL
    '            End Try
    '            'iRec = oCom.ExecuteNonQuery
    '            If iRec > 0 Then
    '                drshow("ADD FC") = "ADD"
    '            ElseIf iRec < 0 Then
    '                drshow("ADD FC") = "ADD ERR"
    '            Else
    '                drshow("ADD FC") = "Not Add"
    '            End If

    '        End If '**** end if 210 

    '        drshow("MK Suffix") = Trim(orow("tlmksf"))
    '        drshow("YZ Suffix") = Trim(orow("tlyzsf"))
    '        drshow("Product code") = Trim(orow("tlpdcd"))
    '        drshow("ADD FC") = ""

    '        dtshow.Rows.Add(drshow)
    '    Next

    '    grdshow.DataSource = dtshow
    '    grdshow.DataBind()

    '    Dim gdRow As GridViewRow
    '    For Each gdRow In grdshow.Rows
    '        If grdshow.Rows(gdRow.RowIndex).Cells(5).Text.Trim = "220" Or grdshow.Rows(gdRow.RowIndex).Cells(5).Text.Trim = "210" Then
    '            grdshow.Rows(gdRow.RowIndex).Cells(0).Enabled = False
    '        End If
    '    Next

    'End Sub

    Protected Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        lblmessage.Text = ""

        ''If ddlCompany.SelectedValue = "TAP-C" And ddlreport.SelectedValue = "FC" Then
        ''      NOFC()
        ''      Exit Sub
        ''End If

        appExcel.DefaultFilePath = Server.MapPath(".")
        If ddlreport.SelectedValue = "non" Or ddlreport.SelectedValue = "FC" Then
            TemplateFile = CurReportPath & ("TAP-NONprice.XLT")
        Else
            TemplateFile = CurReportPath & ("TAP-INVOICEprice.XLT")
        End If

        workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = workbook.Worksheets(0)
        sheet2 = workbook.Worksheets.Create("INVOICE")

        If txtFrYear.Text = "" Then
            lblmessage.Text = "Please input YEAR!!"
            Exit Sub
        End If

        Call PrintH()
        Call PrintDetail()
        If lblmessage.Text = "NO DATA" Then
            Exit Sub
        Else
            lblmessage.Text = " "
        End If

        Call SetPageProperties()

        sheet.Remove()
        filename = ReportName & "-" & Crtdate & "-" & crtTime & ".xls"
        workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
        Response.End()
        If Conn.State = ConnectionState.Open Then  '********** update 28/3/2012 ************
            Conn.Close()
        End If
    End Sub
    Sub chkyic(ByVal inv As String, ByVal prdcd As String)
        Dim strchk As String
        Dim FACT As String
        If ddlCompany.SelectedValue = "TAP-B" Then
            FACT = "32T1"
        ElseIf ddlCompany.SelectedValue = "TAP-C" Then
            FACT = "32T4"
        End If

        strchk = "select count(*) from yic_priceadj where yicfact='" & FACT & "' and yiprdcd='" & prdcd & "' and  yicprod='" & inv & "' and yicstdts='" & txtFrYear.Text & "" & ddlMonth.SelectedValue & "01'"
        Dim cmdchk As New OleDbCommand(strchk, Conn)
        recchk = cmdchk.ExecuteScalar()
    End Sub
    Protected Sub PrintDetail()
        Dim STRSQL As String
        oTable = getdata()

        If oTable.Rows.Count = 0 Then
            lblmessage.Text = "NO DATA"
            Exit Sub
        End If
        Dim tabletemp As Integer = 0
        Dim iRec As Integer = 0
        intCurrow = 6
        ''  orow = oTable.Rows(0)
        If Conn.State = ConnectionState.Closed Then  '********** update 28/3/2012 ************
            Conn.Open()
        End If

        Dim chkpart As String = ""  '****** update 29/3/2012 for check part in TAP-C 
        Dim yiccust As String = "999999" '**** pui update 21/11/2011 : new customer (JOHOKU)
        Dim chkcust As String = "" '****** pui update 10/1/2012 : check customer 
        Select Case ddlreport.SelectedValue
            Case "non"
                Dim strDate As String = Now.Year.ToString & Right("00" & Now.Month.ToString, 2) & Right("00" & Now.Day.ToString, 2)
                For Each orow In oTable.Rows
                    If (cbinvoiceno.Checked) Then
                        sheet.Range("A8:I8").CopyTo(sheet2.Range(intCurrow, 1))
                        sheet2.Range(intCurrow, 2).Text = orow("siinvn")
                        sheet2.Range(intCurrow, 3).Text = Right(orow("siinvd"), 2) & "/" & Mid(orow("siinvd"), 5, 2) & "/" & Left(orow("siinvd"), 4)
                        sheet2.Range(intCurrow, 4).Text = orow("sicust")
                        sheet2.Range(intCurrow, 5).Text = orow("ilprod")
                        sheet2.Range(intCurrow, 6).Text = orow("ilmksf")
                        sheet2.Range(intCurrow, 7).Text = orow("ilyzsf")
                        sheet2.Range(intCurrow, 8).Text = orow("ilpdcd")
                        sheet2.Range(intCurrow, 9).Text = orow("hdtyp")
                    Else
                        sheet.Range("A8:I8").CopyTo(sheet2.Range(intCurrow, 1))
                        sheet2.Range(intCurrow, 4).Text = orow("sicust")
                        sheet2.Range(intCurrow, 5).Text = orow("ilprod")
                        sheet2.Range(intCurrow, 6).Text = orow("ilmksf")
                        sheet2.Range(intCurrow, 7).Text = orow("ilyzsf")
                        sheet2.Range(intCurrow, 8).Text = orow("ilpdcd")
                        sheet2.Range(intCurrow, 9).Text = orow("hdtyp")
                    End If

                    '**** pui update 10/1/2012 check customer code ************
                    If orow("TTCUST") Is System.DBNull.Value Then
                        chkcust = ""
                    Else
                        chkcust = orow("TTCUST")
                    End If

                    If chkcust <> "" Then
                        '**** ถ้าไม่เท่ากับ ช่องว่าง คือ johoku ยังไม่รู้ว่าถ้าเจอแล้วจะให้ทำอะไร  10/1/2012 By:pui ******  
                    End If
                    '************ end 10/1/2012 ****************

                    If Trim(orow("ilpdCD")) = "210" And ddlCompany.SelectedValue <> "TAP-B" Then
                        chkyic(Trim(orow("ilprod")), Trim(orow("ilpdCD")))
                        If Trim(orow("sicust")) = "333610" Then 'KIDEN 
                            If recchk = 0 Then

                                STRSQL = "INSERT INTO YIC_PRICEADJ " & _
                               "SELECT YICFACT, YICCUST, YICDOEX, YICNMSP, YIMKSF, YIYZSF,'" & Trim(orow("ilprod")) & "', YICPRDK, YICSTDT, YICBCRN, YICTYPE, YICUPTR, YICREAS, YIRMK1, " & _
                                " YIQONO, YIUPFG, YICLDFL," & strDate & "," & Session("USERID") & ", 0, '', YICDLDT, YICDLUS, YIEXCHG, YICOPBS, YICOPAD, YIYEAR, YIMONTH," & _
                                " YIPDCOP, YIPRICE, YIPRDCD, YIFREF, YIMAKER, YIMODEL, YICSTDTS, YIUPrice, '', '0', " & _
                                " yicqid,yicpak,yicmat,yicprcs,yipakcos,yinopac " & _
                                " FROM YIC_PRICEADJ WHERE YIYEAR = '" & Trim(txtFrYear.Text) & "' AND YIMONTH = '" & ddlMonth.SelectedValue & "' AND substring(YICPROD,4, 6) = '" & Left(Trim(orow("ilprod")), 6) & "' "

                                Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
                                Try
                                    iRec = oCom.ExecuteNonQuery()
                                Catch ex As Exception
                                    iRec = -1
                                    lblmessage.Text = STRSQL
                                End Try

                            End If
                        Else
                            If recchk = 0 Then

                                STRSQL = "INSERT INTO YIC_PRICEADJ " & _
                               "SELECT YICFACT, YICCUST, YICDOEX, YICNMSP, YIMKSF, YIYZSF,'" & Trim(orow("ilprod")) & "', YICPRDK, YICSTDT, YICBCRN, YICTYPE, YICUPTR, YICREAS, YIRMK1, " & _
                                " YIQONO, YIUPFG, YICLDFL," & strDate & "," & Session("USERID") & ", 0, '', YICDLDT, YICDLUS, YIEXCHG, YICOPBS, YICOPAD, YIYEAR, YIMONTH," & _
                                " YIPDCOP, YIPRICE, YIPRDCD, YIFREF, YIMAKER, YIMODEL, YICSTDTS, YIUPrice, '', '0', " & _
                                " yicqid,yicpak,yicmat,yicprcs,yipakcos,yinopac " & _
                                " FROM YIC_PRICEADJ WHERE YIYEAR = '" & Trim(txtFrYear.Text) & "' AND YIMONTH = '" & ddlMonth.SelectedValue & "' AND substring(YICPROD,4, 6) = '" & Left(Trim(orow("ilprod")), 6) & "' "

                                Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
                                Try
                                    iRec = oCom.ExecuteNonQuery()
                                Catch ex As Exception
                                    iRec = -1
                                    lblmessage.Text = STRSQL
                                End Try
                            End If
                        End If


                        If iRec > 0 Then
                            sheet2.Range(intCurrow, 6).Text = "ADD"
                        ElseIf recchk > 0 Then
                            sheet2.Range(intCurrow, 6).Text = "DUPLICATE"
                            sheet2.Range(intCurrow, 6, intCurrow, 7).Merge()
                        ElseIf iRec < 0 Then
                            sheet2.Range(intCurrow, 6).Text = "ADD ERR"
                        Else
                            sheet2.Range(intCurrow, 6).Text = "Not Add"
                        End If

                    End If '*****end if check product code and factory For Automatic ADD new price ******
                    intCurrow += 1
                    iRec = 0
                Next

                oTable2 = getdataisuzu()

                If oTable2.Rows.Count = 0 Then
                    lblmessage.Text = "NO DATA ISUZU TRIPETCH"
                    Exit Sub
                End If

                sheet3 = workbook.Worksheets.Create("ISUZU TRIPETCH")
                sheet.Range("A10:E14").CopyTo(sheet3.Range("A1"))
                sheet3.Range(3, 3).Text = ddlCompany.SelectedItem.Text
                sheet3.Range(2, 5).Text = (ddlMonth.SelectedValue) & "/" & (txtFrYear.Text)

                sheet3.Range(2, 2).Text = "No TAP Price Report(Isuzu Tripetch)"

                tabletemp = 0
                iRec = 0
                intCurrow = 6

                For Each orow2 In oTable2.Rows
                    sheet.Range("A16:E16").CopyTo(sheet3.Range(intCurrow, 1))
                    sheet3.Range(intCurrow, 2).Text = orow2("ilprod")
                    sheet3.Range(intCurrow, 3).Text = orow2("ilmksf")
                    sheet3.Range(intCurrow, 4).Text = orow2("ilyzsf")
                    sheet3.Range(intCurrow, 5).Value = orow2("ilqty")
                    intCurrow += 1
                    iRec = 0
                Next

                sheet3.Range(intCurrow - 1, 2, intCurrow - 1, 5).CellStyle.Borders(ExcelBordersIndex.EdgeBottom).LineStyle = ExcelLineStyle.Medium

            Case "price"
                For Each orow In oTable.Rows
                    iRec += 1
                    sheet.Range("A7:H7").CopyTo(sheet2.Range(intCurrow, 1))
                    sheet2.Range(intCurrow, 2).Text = iRec
                    sheet2.Range(intCurrow, 3).Text = orow("ilprod")
                    sheet2.Range(intCurrow, 4).Text = orow("ilmksf")
                    sheet2.Range(intCurrow, 5).Text = orow("ilyzsf")
                    sheet2.Range(intCurrow, 6).Text = orow("ilta10")
                    sheet2.Range(intCurrow, 7).Text = orow("ilpdcd")
                    sheet2.Range(intCurrow, 8).Text = orow("hdtyp")
                    intCurrow += 1
                Next
            Case "FC"
                Dim strDate As String = Now.Year.ToString & Right("00" & Now.Month.ToString, 2) & Right("00" & Now.Day.ToString, 2)
                Dim coladd As Integer = 0
                For Each orow In oTable.Rows

                    If chkpart <> Trim(orow("tlprod")) & Trim(orow("tlmksf")) & Trim(orow("tlyzsf")) Then  '***** update 29/3/2012 check part double *****



                        If (cbinvoiceno.Checked) Then '*** pui remark 22/11/2011 *****
                            coladd = 6
                            sheet.Range("A8:I8").CopyTo(sheet2.Range(intCurrow, 1))
                            '' sheet.Range("D8:I8").CopyTo(sheet2.Range(intCurrow, 4))
                            sheet2.Range(intCurrow, 2).Text = orow("stinvn")
                            sheet2.Range(intCurrow, 3).Text = Right(orow("stinvd"), 2) & "/" & Mid(orow("stinvd"), 5, 2) & "/" & Left(orow("stinvd"), 4)
                            sheet2.Range(intCurrow, 4).Text = orow("tlcusb")
                            sheet2.Range(intCurrow, 5).Text = orow("tlprod")
                            sheet2.Range(intCurrow, 6).Text = orow("tlmksf")
                            sheet2.Range(intCurrow, 7).Text = orow("tlyzsf")
                            sheet2.Range(intCurrow, 8).Text = orow("tlpdcd")
                            sheet2.Range(intCurrow, 9).Text = orow("hdtyp")
                        Else
                            coladd = 6
                            sheet.Range("A8:I8").CopyTo(sheet2.Range(intCurrow, 1))
                            ''sheet.Range("D8:I8").CopyTo(sheet2.Range(intCurrow, 2))
                            sheet2.Range(intCurrow, 4).Text = orow("tlcusb")
                            sheet2.Range(intCurrow, 5).Text = orow("tlprod")
                            sheet2.Range(intCurrow, 6).Text = orow("tlmksf")
                            sheet2.Range(intCurrow, 7).Text = orow("tlyzsf")
                            sheet2.Range(intCurrow, 8).Text = orow("tlpdcd")
                            sheet2.Range(intCurrow, 9).Text = orow("hdtyp")
                        End If

                        If Trim(orow("tlpdcd")) = "210" Then       'AW
                            STRSQL = "INSERT INTO YIC_VDMY "

                            If ddlCompany.SelectedValue = "TAP-B" Then
                                '  STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T1'',''1996'',rtrim(ntypat)AS ntypat,''210'',''999999'','''','''',"
                            Else 'TAP-C
                                STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T2'',''1996'',rtrim(ntypat)AS ntypat,''210'',''999999'','''','''',"

                                STRSQL += " nttfdt,ntaspr/10000 AS ntaspr,''BAH'',''0'',''KG'',''0''," & strDate & ",''" & Session("USERID") & "'',"
                                STRSQL += "''0'','''',''0'','''',''0'','''',''0''"
                                STRSQL += " FROM  "
                                STRSQL += " vc1tapaw.vcnt "
                                STRSQL += " WHERE ntac = ''FCDOM'' AND nttfdt >= ''" & Left(strDate, 6) & "01'' AND  SUBSTRING(NTYPAT,4,11) =''" & Trim(orow("tlprod")) & "'' and NTASPR>0') "

                                Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
                                Try
                                    iRec = oCom.ExecuteNonQuery()
                                Catch ex As Exception
                                    iRec = -1
                                    lblmessage.Text = STRSQL
                                End Try
                                'iRec = oCom.ExecuteNonQuery

                                If iRec > 0 Then
                                    sheet2.Range(intCurrow, coladd).Text = "ADD"
                                ElseIf iRec < 0 Then
                                    sheet2.Range(intCurrow, coladd).Text = "ADD ERR"
                                Else
                                    sheet2.Range(intCurrow, coladd).Text = "Not Add"
                                End If

                            End If  ' **** end if TAP-B 
                            '**** HONDA TAP-B 
                        ElseIf (Trim(orow("tlpdcd")) = "450" Or Trim(orow("tlpdcd")) = "490") And Left(Trim(orow("tlcusb")), 4) = "1104" And ddlCompany.SelectedValue = "TAP-B" Then
                            STRSQL = "INSERT INTO YIC_VDMY "

                            STRSQL = STRSQL & "SELECT top 1 * FROM openquery(as400,'select ''32T1'',''0'',replace(rtrim(ntypat),''-'','''') AS ntypat,''280'',''999999'','''','''',"

                            '*** pui update 14/11/2011 problem normal/spare *******
                            '' STRSQL += " nttfdt,ntaspr/1000 AS ntaspr,''BAH'',''0'',''KG'',''0''," & strDate & ",''" & Session("USERID") & "'',"
                            STRSQL += " nttfdt,ntaspr/1000 AS ntaspr,''BAH'',ntaspr/1000 AS ntaspr,''KG'',''0''," & strDate & ",''" & Session("USERID") & "'',"
                            '********** end 14/11/2011 *************

                            STRSQL += "''0'','''',''0'','''',''0'','''',''0''"
                            STRSQL += " FROM  "
                            STRSQL += " vc1tapij.vcnt "
                            STRSQL += " WHERE ntac = ''FCDOM'' AND nttfdt < ''" & strDate & "'' AND  replace(ntypat,''-'','''') =''" & Trim(orow("tlprod")) & "''  and NTASPR>0 order by nttfdt desc ') "
                            Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
                            Try
                                iRec = oCom.ExecuteNonQuery()
                            Catch ex As Exception
                                iRec = -1
                                lblmessage.Text = STRSQL
                                '**** pui add 14/2/2012 *****
                                ''If Trim(orow("tlpdcd")) = "490" And Left(Trim(orow("tlprod")), 5) = "38221" Then

                                STRSQL = "SELECT top 1 * into #a" & tabletemp & " FROM openquery(as400,'select replace(rtrim(ntypat),''-'','''') AS ntypat, "
                                STRSQL += " nttfdt,ntaspr/1000 AS ntaspr "
                                STRSQL += " FROM  "
                                STRSQL += " vc1tapij.vcnt "
                                STRSQL += " WHERE ntac = ''FCDOM'' AND nttfdt < ''" & strDate & "'' AND  replace(ntypat,''-'','''') =''" & Trim(orow("tlprod")) & "'' order by nttfdt desc ') "


                                STRSQL += " IF EXISTS (SELECT * FROM YIC_VDMY WHERE MYUPTR=0 AND TRANSAS='Y' AND MYSTDT=(select nttfdt from #a" & tabletemp & ") AND MYPROD='" & Trim(orow("tlprod")) & "' AND PRDC='280')"
                                STRSQL += "UPDATE YIC_VDMY SET TRANSAS='',TRANDTD='0',MYUPTR=(select ntaspr from #a" & tabletemp & "),MYCWET=(select ntaspr from #a" & tabletemp & "),MYUPDT='" & strDate & "',MYUPUS='" & Session("USERID") & "' "
                                STRSQL += " WHERE MYUPTR=0 AND TRANSAS='Y' AND MYSTDT=(select nttfdt from #a" & tabletemp & ") AND MYPROD='" & Trim(orow("tlprod")) & "' AND PRDC='280' "

                                ''End If
                                Dim oCom2 As New Data.OleDb.OleDbCommand(STRSQL, Conn)
                                iRec = oCom2.ExecuteNonQuery()
                                tabletemp += 1
                            End Try

                            If iRec > 0 Then
                                sheet2.Range(intCurrow, coladd).Text = "ADD"
                                ''sheet2.Range(intCurrow, 7).Text = orow("nttfdt")
                            ElseIf iRec < 0 Then
                                sheet2.Range(intCurrow, coladd).Text = "ADD ERR"
                            Else
                                sheet2.Range(intCurrow, coladd).Text = "Not Add"
                            End If
                            'End If '**** end if  HONDA TAP-B 

                        ElseIf Trim(orow("tlpdcd")) = "230" Or Trim(orow("tlpdcd")) = "231" Then   ' VTA ALL FACTORY   ***** pui update 22/11/2011 *****
                            STRSQL = "INSERT INTO YIC_VDMY "

                            If ddlCompany.SelectedValue = "TAP-B" Then
                                STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T1'',''1996''," & Trim(orow("tlprod")) & ",''230'',''999999'','''','''',"
                            Else 'TAP-C
                                STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T2'',''1996''," & Trim(orow("tlprod")) & ",''230'',''999999'','''','''',"
                            End If
                            STRSQL += " nttfdt,ntaspr/1000 AS ntaspr,''BAH'',''0'',''KG'',''0''," & strDate & ",''" & Session("USERID") & "'',"
                            STRSQL += "''0'','''',''0'','''',''0'','''',''0''"
                            STRSQL += " FROM  "
                            STRSQL += " VC1TAPVTA.vcnt "
                            STRSQL += " WHERE ntac = ''FCDOM''  and NTASPR>0 AND nttfdt >= ''" & Left(strDate, 6) & "01'' AND  substring(ntypat,4," & Len(Trim(orow("tlprod"))) & " )=''" & Trim(orow("tlprod")) & "'' ') "
                            Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
                            Try
                                iRec = oCom.ExecuteNonQuery()
                            Catch ex As Exception
                                iRec = -1
                                lblmessage.Text = STRSQL
                            End Try

                            If iRec > 0 Then
                                sheet2.Range(intCurrow, coladd).Text = "ADD"
                            ElseIf iRec < 0 Then
                                sheet2.Range(intCurrow, coladd).Text = "ADD ERR"
                            Else
                                sheet2.Range(intCurrow, coladd).Text = "Not Add"
                            End If
                            'End If '**** end if VTA                 

                        ElseIf Trim(orow("tlpdcd")) = "270" Then   ' TUBE   ALL FACTORY   ***** pui update 22/11/2011 *****
                            STRSQL = "INSERT INTO YIC_VDMY "

                            If ddlCompany.SelectedValue = "TAP-B" Then
                                STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T1'',''1996'',rtrim(ntypat)AS ntypat,''270'',''999999'','''','''',"
                            Else 'TAP-C
                                STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T2'',''1996'',rtrim(ntypat)AS ntypat,''270'',''999999'','''','''',"
                            End If
                            STRSQL += " nttfdt,ntaspr/10000 AS ntaspr,''BAH'',''0'',''KG'',''0''," & strDate & ",''" & Session("USERID") & "'',"
                            STRSQL += "''0'','''',''0'','''',''0'','''',''0''"
                            STRSQL += " FROM  "
                            STRSQL += " vc1tapaw.vcnt "
                            STRSQL += " WHERE ntac = ''FCDOM''  and NTASPR>0 AND nttfdt >= ''" & Left(strDate, 6) & "01'' AND REPLACE(NTYPAT,''-'','''')=''" & Trim(orow("tlprod")) & "'' ') "
                            Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
                            Try
                                iRec = oCom.ExecuteNonQuery()
                            Catch ex As Exception
                                iRec = -1
                                lblmessage.Text = STRSQL
                            End Try

                            If iRec > 0 Then
                                sheet2.Range(intCurrow, coladd).Text = "ADD"
                            ElseIf iRec < 0 Then
                                sheet2.Range(intCurrow, coladd).Text = "ADD ERR"
                            Else
                                sheet2.Range(intCurrow, coladd).Text = "Not Add"
                            End If
                            ' End If '**** end if TUBE 

                        ElseIf ddlCompany.SelectedValue = "TAP-C" And Trim(orow("tlpdcd")) <> "220" Then

                            STRSQL = "INSERT INTO YIC_VDMY "

                            STRSQL = STRSQL & "select '32T2','1996',rtrim(yicprod),yiprdcd,'999999','','',"

                            STRSQL += " yicstdt,yicuptr/1000 ,'BAH','0','KG','0'," & strDate & ",'" & Session("USERID") & "',"
                            STRSQL += "'0','','0','','0','','0' "
                            STRSQL += " FROM  "
                            STRSQL += " yic_priceadj "
                            STRSQL += " WHERE   yicreas >= '" & Trim(orow("stinvd")) & "'  AND rtrim(yicprod)='" & Trim(orow("tlprod")) & "' and tranas='Y'  "

                            Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
                            Try
                                iRec = oCom.ExecuteNonQuery()
                            Catch ex As Exception
                                iRec = -1
                                lblmessage.Text = STRSQL
                            End Try

                            If iRec > 0 Then
                                sheet2.Range(intCurrow, coladd).Text = "ADD"
                            ElseIf iRec < 0 Then
                                sheet2.Range(intCurrow, coladd).Text = "ADD ERR"
                            Else
                                sheet2.Range(intCurrow, coladd).Text = "Not Add"
                            End If

                        End If

                        chkpart = Trim(orow("tlprod")) & Trim(orow("tlmksf")) & Trim(orow("tlyzsf")) '***** check part double *****
                        intCurrow += 1
                    End If '***** END check part double *****

                Next
        End Select

        If Conn.State = ConnectionState.Open Then  '********** update 28/3/2012 ************
            Conn.Close()
        End If
    End Sub
    Protected Sub PrintH()

        With sheet

            If ddlreport.SelectedValue = "non" Then
                sheet.Range("A1:I5").CopyTo(sheet2.Range("A1"))
                sheet2.Range(3, 4).Text = ddlCompany.SelectedItem.Text
                sheet2.Range(2, 4).Text = (ddlMonth.SelectedValue) & "/" & (txtFrYear.Text)
                If (cbinvoiceno.Checked) Then
                    sheet.Range("B6:I7").CopyTo(sheet2.Range(6, 2))
                Else
                    sheet.Range("C6:I7").CopyTo(sheet2.Range(6, 2))
                End If
                sheet2.Range(2, 2).Text = "No TAP Price Report"
            ElseIf ddlreport.SelectedValue = "FC" Then
                sheet.Range("A1:H3").CopyTo(sheet2.Range("A1"))
                sheet2.Range(1, 2).Text = "FC"
                sheet2.Range(3, 2).Text = "FACTORY :"
                sheet2.Range(3, 3).Text = ddlCompany.SelectedItem.Text
                sheet2.Range(2, 4).Text = (ddlMonth.SelectedValue) & "/" & (txtFrYear.Text)
                '' If (cbinvoiceno.Checked) Then
                sheet.Range("A4:C5").CopyTo(sheet2.Range(4, 1))
                sheet.Range("D4:I5").CopyTo(sheet2.Range(4, 4))
                ''Else
                ''    sheet.Range("D4:I5").CopyTo(sheet2.Range(4, 2))
                ''End If
                sheet2.Range(2, 2).Text = "No TAP FC Report"
            Else
                sheet.Range("A1:G5").CopyTo(sheet2.Range("A1"))
                sheet2.Range(3, 3).Text = ddlCompany.SelectedItem.Text
                sheet2.Range(2, 3).Text = (ddlMonth.SelectedValue) & "/" & (txtFrYear.Text)
            End If
        End With


    End Sub
    Protected Sub SetPageProperties()
        With sheet2
            Select Case ddlreport.SelectedValue
                Case "non"
                    .SetColumnWidth(1, 0.5)                       'COLUMN A  -- 
                    .SetColumnWidth(2, 12)                        'COLUMN B  -- invoice no.
                    .SetColumnWidth(3, 12)                      'COLUMN C  -- product
                    .SetColumnWidth(4, 12)                      'COLUMN D  -- 
                    .SetColumnWidth(5, 20)
                    .SetColumnWidth(6, 8)                        'COLUMN E  -- 
                    .SetColumnWidth(7, 8)                        'COLUMN F  --

                    If oTable2.Rows.Count > 0 Then
                        sheet3.SetColumnWidth(1, 0.5)
                        sheet3.SetColumnWidth(2, 20)
                        sheet3.SetColumnWidth(3, 8)
                        sheet3.SetColumnWidth(4, 8)
                        sheet3.SetColumnWidth(5, 6.43)
                    End If

                Case "price"
                    .SetColumnWidth(1, 0.5)                       'COLUMN A  -- 
                    .SetColumnWidth(2, 9.5)                        'COLUMN B  -- invoice no.
                    .SetColumnWidth(3, 22)                        'COLUMN C  -- product
                    .SetColumnWidth(4, 8.5)                        'COLUMN D  -- 
                    .SetColumnWidth(5, 8.5)                        'COLUMN E  -- 
                    .SetColumnWidth(6, 11)                        'COLUMN F  --
                    .SetColumnWidth(7, 9.5)
                Case "FC"
                    .SetColumnWidth(1, 0.5)
                    '' If (cbinvoiceno.Checked) Then
                    .SetColumnWidth(2, 12)
                    .SetColumnWidth(3, 12)
                    .SetColumnWidth(4, 12)
                    .SetColumnWidth(5, 20)
                    .SetColumnWidth(6, 8)
                    .SetColumnWidth(7, 8)
                    .SetColumnWidth(8, 8)
                    ''Else
                    ''.SetColumnWidth(2, 12)
                    ''.SetColumnWidth(3, 20)
                    ''.SetColumnWidth(4, 8)
                    ''.SetColumnWidth(5, 8)
                    ''.SetColumnWidth(6, 8)
                    ''.SetColumnWidth(7, 8)
                    ''End If

            End Select


            ' .SetRowHeight(1, 10)
            ' .SetRowHeight(2, 10)
            ' .SetRowHeight(3, 21)

            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Portrait
            .PageSetup.LeftMargin = 0.5
            .PageSetup.RightMargin = 0.5
            .PageSetup.TopMargin = 0.5
            .PageSetup.BottomMargin = 0.5
            .PageSetup.Zoom = 70
            .PageSetup.PrintTitleRows = "$4:$5"
            .Range("A6").FreezePanes()

        End With
    End Sub

    Protected Sub grdshow_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdshow.SelectedIndexChanged
        If grdshow.SelectedRow.Cells(5).Text = "220" Or grdshow.SelectedRow.Cells(5).Text = "200" Then
            Exit Sub
        End If

        Dim dtadd As New DataTable
        Dim dradd As DataRow

        dtadd = getmaster(grdshow.SelectedRow.Cells(1).Text.Trim, grdshow.SelectedRow.Cells(5).Text.Trim)
        If dtadd.Rows.Count <= 0 Then
            lblerr.Text = "Part No. : " & grdshow.SelectedRow.Cells(1).Text & " No MASTER"
            grddetail.DataSource = Nothing
            grddetail.DataBind()
            Exit Sub
        Else
            lblerr.Text = ""
        End If
        grddetail.DataSource = dtadd
        grddetail.DataBind()

    End Sub

    Function getmaster(ByVal part As String, ByVal pdcd As String)
        If Conn.State = ConnectionState.Closed Then  '********** update 28/3/2012 ************
            Conn.Open()
        End If
        Dim sql As String
        sql = "SELECT * FROM OPENQUERY(as400,'SELECT rtrim(ntypat) AS ntypat,nttfdt,ntaspr/1000 AS ntaspr,''" & pdcd.Trim & "''as PDCD  FROM  "
        ' sql += " vc1tapij.vcnt "
        sql += " vc1tapaw.vcnt "
        'sql += " WHERE ntac = ''FCDOM'' and ntypat=''" & part & "'' ')"
        sql += " WHERE ntac = ''FCDOM'' and SUBSTRING(NTYPAT,4,11) =''" & part & "'' ')"
        Dim da As New OleDbDataAdapter(sql, Conn)
        Dim dt As New DataTable
        da.Fill(dt)
        Return dt
        If Conn.State = ConnectionState.Open Then  '********** update 28/3/2012 ************
            Conn.Close()
        End If
    End Function

    Protected Sub grddetail_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grddetail.SelectedIndexChanged

        If Conn.State = ConnectionState.Closed Then  '********** update 28/3/2012 ************
            Conn.Open()
        End If

        Dim STRSQL As String
        Dim iRec As Integer = 0
        Dim strDate As String = Now.Year.ToString & Right("00" & Now.Month.ToString, 2) & Right("00" & Now.Day.ToString, 2)

        If grddetail.SelectedRow.Cells(2).Text <> "210" And grddetail.SelectedRow.Cells(2).Text <> "270" And grddetail.SelectedRow.Cells(2).Text <> "220" Then       'INJ,JB,PRESS,RUBBER
            STRSQL = "INSERT INTO YIC_VDMY "
            STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T5'',''1996'',REPLACE(rtrim(NTYPAT),''-'','''') AS ntypat,''" & grddetail.SelectedRow.Cells(2).Text.Trim & "'',''999999'','''','''',"
            STRSQL += " nttfdt,ntaspr/1000 AS ntaspr,''BAH'',''0'',''KG'',''0''," & strDate & ",''" & Session("USERID") & "'',"
            STRSQL += "''0'','''',''0'','''',''0'','''',''0''"
            STRSQL += " FROM  "
            STRSQL += " vc1tapij.vcnt A LEFT JOIN YGSS.YPRD B ON A.NTYPAT=B.PPROD "
            STRSQL += " WHERE ntac = ''FCDOM'' AND nttfdt = ''" & grddetail.SelectedRow.Cells(3).Text.Trim & "'' AND PPRDCD=''" & grddetail.SelectedRow.Cells(2).Text.Trim & "'' AND  NTYPAT =''" & grddetail.SelectedRow.Cells(1).Text.Trim & "'' ') "
            Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
            Try
                ''iRec = oCom.ExecuteNonQuery()
            Catch ex As Exception
                iRec = -1
                ' lblmessage.Text = STRSQL
            End Try
            If iRec > 0 Then
                lblerr.Text = "ADD " & grddetail.SelectedRow.Cells(1).Text & " Complete."
            ElseIf iRec < 0 Then
                lblerr.Text = "ADD ERROR"
            Else
                lblerr.Text = "Not Add"
            End If
        End If '**** end if INJ,JB,PRESS,RUBBER 


        If grddetail.SelectedRow.Cells(2).Text = "270" Then       '****TUBE********
            STRSQL = "INSERT INTO YIC_VDMY "
            STRSQL = STRSQL & "SELECT * FROM openquery(as400,'select ''32T4'',''1996'',rtrim(ntypat) AS ntypat,''" & grddetail.SelectedRow.Cells(2).Text.Trim & "'',''999999'','''','''',"
            STRSQL += " nttfdt,ntaspr/10000 AS ntaspr,''BAH'',''0'',''KG'',''0''," & strDate & ",''" & Session("USERID") & "'',"
            STRSQL += "''0'','''',''0'','''',''0'','''',''0''"
            STRSQL += " FROM  "
            STRSQL += " vc1tapaw.vcnt "
            STRSQL += " WHERE ntac = ''FCDOM'' AND nttfdt = ''" & grddetail.SelectedRow.Cells(3).Text.Trim & "'' "
            STRSQL += " AND LEFT(NTYPAT,3)=''450'' AND SUBSTRING(NTYPAT,10,11) <> '''' AND SUBSTRING(NTYPAT,25,25) <> ''B'' AND  NTYPAT =''" & grddetail.SelectedRow.Cells(1).Text.Trim & "'' ') "
            Dim oCom As New Data.OleDb.OleDbCommand(STRSQL, Conn)
            Try
                ''iRec = oCom.ExecuteNonQuery()
            Catch ex As Exception
                iRec = -1

            End Try
            If iRec > 0 Then
                lblerr.Text = "ADD " & grddetail.SelectedRow.Cells(1).Text & " Complete."
            ElseIf iRec < 0 Then
                lblerr.Text = "ADD ERROR"
            Else
                lblerr.Text = "Not Add"
            End If
        End If '**** end if 270 


        grddetail.DataSource = Nothing
        grddetail.DataBind()

        If Conn.State = ConnectionState.Open Then  '********** update 28/3/2012 ************
            Conn.Close()
        End If
    End Sub



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If Session("USERID") = "" Then Response.Redirect("Login.aspx")
        'If Not Page.IsPostBack Then
        '    Dim checkpermis As Integer
        '    checkpermis = ClassMaster.getpermis("YICREPORT", Session("USERID"))
        '    If checkpermis = 0 Then
        '        Response.Redirect("~\NotAuthority.aspx")
        '    End If
        '    txtFrYear.Text = Right(Crtdate, 4)

        '    If Conn.State = ConnectionState.Open Then  '********** update 28/3/2012 ************
        '        Conn.Close()
        '    End If
        'End If
    End Sub
End Class
