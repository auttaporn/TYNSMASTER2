Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Partial Class UploadWH
    Inherits System.Web.UI.Page
    Private savePath As String = Server.MapPath("~/uploads/OtherJob/")
    Private aRecord(), Detail()
    Private aHeader(), aHeader1()
    Private strHeader, File_Type, strFact, StrLibPACK, Factory, Period, StrLib, StrLibPart, Customer, PartNo, ModelYear, AirTerm, BoatTerm As String
    Private StrLibPRICE As String = "YGSS.TMVDMC"
    Private StrLibTSPRATE As String = "YGSS.TSPRATE"
    Private EffDate, CntRec, ii, CntRec2, cntRec3, Errnum As Integer
    Private AirPrice, BoatPrice, FCCOST, PACKING, MPACKING As Double
    Private strConn400 = Classconn.strCon400
    Private oConn400 As New OleDb.OleDbConnection(strConn400)
    Private cmd400 As New OleDbCommand()
    Private strCmd, strFileError, strFileNameError, suffix As String
    Private strCond, YSFX, MSFX, Ftype, Cntmark, uom, UserId, ErrorFlag, BASEFC, COPBASE, COPADJ, EXCH As String
    Private strdate As String = Date.Now.ToString("yyyyMMdd")
    Private strDetail, sDate, strDat, ChKDigit1, ChKDigit2, PartNo1, MAKER, MODEL, MAKER2, MAKER3, MODEL2, MODEL3, QID As String
    Private Fcost, CuWeight As Double
    Private markup As Double = 0
    Private oTable, oTable2 As New DataTable
    Private oROW, oROW2 As DataRow
    Private objStreamReader
    Private ASCII As Encoding = Encoding.ASCII
    Private TSwriteError
    Private i, SeqNo, LL As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'message.Text = ""

        'If Session("UID") = "" Then Response.Redirect("Login.aspx")
        'Dim checkpermis As Integer
        'checkpermis = ManageFile.getpermis("TAPCOST03", Session("UID"))
        'If checkpermis = 0 Then
        '    Response.Redirect("~\NotAuthority.aspx")
        'End If

        'txtUserCode.Value = Session("UID")
    End Sub

    Sub ClearData()
        Dim cntRec1 As Double
        On Error GoTo MsgErr1

        strCond = " WHERE MYPROD = '" & PartNo & "' AND MYSTDT = " & EffDate & _
         "  AND MYMLYY = " & ModelYear & " AND MYMSFX = '" & MSFX & "' AND " & _
         " MYYSFX = '" & YSFX & _
         "' AND MYFACT = '" & Factory & "'"
        strCond = strCond & " AND MYLDFL = '1' "
        strCmd = "Select COUNT(MYPROD)  FROM   " & StrLib & strCond
        Dim Cmd1 As New OleDbCommand()
        Cmd1.Connection = oConn400
        Cmd1.CommandText = strCmd
        cntRec1 = Cmd1.ExecuteScalar()

        If cntRec1 > 0 Then
            strCond = " WHERE MYPROD = '" & PartNo & "' AND MYSTDT = " & EffDate & _
           "  AND MYMLYY = " & ModelYear & " AND MYMSFX = '" & MSFX & "' AND " & _
           " MYYSFX = '" & YSFX & _
           "' AND MYFACT = '" & Factory & "'"
            strCond = strCond & " AND MYLDFL = '1' "
            strCmd = "Delete   FROM   " & StrLib & strCond
            Dim Cmd2 As New OleDbCommand()
            Cmd2.Connection = oConn400
            Cmd2.CommandText = strCmd
            Cmd2.ExecuteNonQuery()
        End If
        Exit Sub
MsgErr1:
        Response.Write(cntRec1 & " " & strCmd)
        Response.End()
    End Sub

    Sub CheckData()
        Call ClearData()
        strCond = " WHERE MYPROD = '" & PartNo & "' AND MYSTDT = " & EffDate & _
         "   AND MYMSFX = '" & MSFX & "' AND " & _
         " MYYSFX = '" & YSFX & "' AND MYMLYY='" & ModelYear & "'"

        strCmd = "SELECT COUNT(MYPROD)   FROM   " & StrLIB & strCond
        Dim Cmd3 As New OleDbCommand()
        Cmd3.connection = oConn400
        Cmd3.CommandText = strCMD
        CntRec = Cmd3.ExecuteScalar

        Dim strsql As String
        strsql = "select count(*) from " & StrLibPRICE & " where TMPROD = '" & PartNo & "' AND TMSTDT = " & EffDate & _
         "   AND TMMKSF = '" & MSFX & "' AND TMFACT='" & Factory & "' AND " & _
         " TMYZSF = '" & YSFX & "'"
        Dim Cmd4 As New OleDbCommand()
        Cmd4.Connection = oConn400
        Cmd4.CommandText = strsql
        CntRec2 = Cmd4.ExecuteScalar

    End Sub

    Sub CheckProd()
        strCond = " WHERE MAFACT = '" & Factory & "' AND MAPROD = '" & PartNo & "'  AND MAMLYY = " & ModelYear
        strcond = strCond & " AND MAMSFX = '" & MSFX & "' AND MAYSFX = '" & YSFX & "' AND MALDFL =  '0'"

        Dim cmd As New OleDbCommand()
        cmd.connection = oConn400
        strCmd = "SELECT MASFLD  FROM   " & StrLIBPart & strCond

        cmd.CommandText = strCMD
        cntRec3 = cmd.ExecuteScalar()
    End Sub

    Sub SplitToArray(ByVal strFile)
        On Error GoTo msgerr3

        Dim objFile As New FileInfo(strFile)
        objStreamReader = objFile.OpenText()
        Dim data As String = objStreamReader.ReadToEnd()

        oConn400.Open()
        aRecord = data.Split(Chr(10))

        Headerset(aRecord(0))
        Call DetailText()
        objStreamReader.Close()
        oConn400.Close()

        If ii > 0 Then
            GetAllFiles()
        End If

        Exit Sub

msgerr3:
        message.Text = "Wrong Format!!!!!!!!!!"
    End Sub
    Sub SaveToDB2(ByVal strCmd As String)

        On Error GoTo msgerr2

        strCmd = "INSERT INTO " & StrLib & " VALUES ( " & strCmd


        Dim CMD As New OleDbCommand(strCmd, oConn400)
        Dim CntRec4 = CMD.ExecuteNonQuery()
        If CntRec4 <> 0 Then
            ErrorFlag = "Insert Complete"
        Else
            message.Text = "Insert Data Failed"
            Exit Sub
        End If
        Exit Sub

msgerr2:

        ErrorFlag = "INSERT ERROR"
        If ii = 0 Then
            strDat = "Factory," & Factory & ", Effective Date," & EffDate & ",UOM," & uom
            TSwriteError.Write(strDat)
            TSwriteError.WriteLine()
            strDat = "No,Part no,SFX,Model,TYPE,F/C,Cu Weight,Packing,Base FC,Copp base,Copp Adj,Exch Rate,QID,Maker,Model,MarkUp,Remark"
            TSwriteError.Write(strDat)
            TSwriteError.WriteLine()
        End If
        strDat = SeqNo & "," & PartNo & ",'" & suffix & "," & ModelYear & "," & Ftype & "," & Fcost & "," & CuWeight & "," & PACKING & "," & BASEFC & "," & COPBASE & "," & COPADJ & "," & EXCH & "," & QID & "," & MAKER & "," & MODEL & "," & markup & "," & ErrorFlag & ""
        TSwriteError.Write(strDat)
        TSwriteError.WriteLine()
        ii = ii + 1

    End Sub
    Sub SAVETMVDMC()



        Dim strSQL As String
        strSQL = "INSERT INTO " & StrLibPRICE & " "
        strSQL += "VALUES( '" & Factory & "', " & ModelYear & ", "
        strSQL += "" & EffDate & ", " 'TMSTDT
        strSQL += " 'NM' ," 'TMPKCD
        strSQL += " '" & MAKER.ToUpper & "' ," 'TMMAKE
        strSQL += " '" & MODEL & "' ," 'TMMODL
        strSQL += " '" & QID & "' ," 'TMQID
        strSQL += "'" & PartNo & "', " 'TMPROD
        strSQL += "'" & MSFX & "', " 'TMMKSF
        strSQL += "'" & YSFX & "', " 'TMYKSF 
        strSQL += "" & Fcost & ", " 'TMUPTR
        strSQL += "" & CuWeight & ", " 'TMWETH
        strSQL += "" & BASEFC & ", " 'TMWETH
        strSQL += "0, " 'TMPRIC
        strSQL += "" & COPBASE & ", " 'COPPER BASE
        strSQL += "" & COPADJ & ", " 'COPPER ADJUST
        strSQL += "" & EXCH & ", " 'EXCHANG RATE
        strSQL += "" & markup & ", " 'TMMARK
        strSQL += "" & PACKING & ", " 'TMPBAS
        strSQL += "0, " 'TMPPRI
        strSQL += " " & EffDate & " , " 'TMSTRD
        strSQL += " " & Left(EffDate, 6) & "31 ," 'TMEXPD
        strSQL += " '" & uom & "' ," 'TMUOM
        strSQL += " 'N' ," 'TMCFFG
        strSQL += "" & strdate & ", " 'MCADDT
        strSQL += "'" & Session("UID") & "', " 'MCADUS
        strSQL += " 0, " 'MCUPDT
        strSQL += "'') " 'MCUPUS


        Dim cmd As New OleDbCommand(strSQL, oConn400)
        Dim CntRec5 = cmd.ExecuteNonQuery()
        If CntRec5 <> 0 Then
            ErrorFlag = "Insert Complete"
        Else
            message.Text = "Insert Data Failed"
            Exit Sub
        End If
    End Sub
    Sub UPDATEFC()
        strCmd = "UPDATE " & StrLib & " SET MYUPTR=" & Fcost & ",MYUPDT=" & strdate & ",MYUPUS='" & Session("UID") & "',MYCWET='" & CuWeight & "' WHERE MYPROD='" & PartNo & "' "
        strCmd += "AND MYMSFX='" & MSFX & "'  AND MYYSFX='" & YSFX & "'  AND MYSTDT='" & EffDate & "'"
        Dim CMD As New OleDbCommand(strCmd, oConn400)
        Dim cntrec6 = CMD.ExecuteNonQuery()
        If cntrec6 = 0 Then
            message.Text = "Update Data Failed"
            Exit Sub
        Else
            ErrorFlag = "Update Complete"
        End If
    End Sub
    Sub UPDATETMVDMC()

        Dim strSQL As String
        strSQL = "UPDATE " & StrLibPRICE & " SET TMUPTR=" & Fcost & ",TMCOPP='" & CuWeight & "',TMPRIC=0,TMPPRI=0,"
        strSQL += "TMEXCH=" & EXCH & ",TMCOPA=" & COPADJ & ",TMCOPB=" & COPBASE & ",TMPBAS=" & PACKING & ",TMUPTB=" & BASEFC & ","
        strSQL += "TMMARK=" & markup & ",TMUPDT='" & strdate & "',TMUPUS='" & Session("UID") & "',TMCFFG='',"
        strSQL += "tmmake='" & MAKER & "',tmmodl='" & MODEL & "',tmqid='" & QID & "'"
        strSQL += " WHERE TMPROD='" & PartNo & "' AND TMMKSF='" & MSFX & "'  AND TMYZSF='" & YSFX & "'  AND TMSTDT='" & EffDate & "' AND TMFACT='" & Factory & "'"
        Dim CMD As New OleDbCommand(strSQL, oConn400)
        Dim cntrec7 = CMD.ExecuteNonQuery()
        If cntrec7 = 0 Then
            message.Text = "Update Data Failed"
            Exit Sub
        Else
            ErrorFlag = "Update Complete"
        End If
    End Sub

    Sub Headerset(ByVal strHead As String)
        aHeader = strHead.Split(",")
        Factory = Left(Trim(aHeader(1)), 5)
        EffDate = aHeader(3)
        uom = Left(Trim(aHeader(5)), 3)

    End Sub

    Sub DetailText()
        Errnum = 0


        Dim DATEDT As New Date

        On Error GoTo msgerr

        Call ManageFile.DeleteFileName(savePath & "ErrorFile\")
        strFileError = ManageFile.GenFileName("FC" & Factory, "", (savePath & "ErrorFile\"), ".csv")
        strFileNameError = savePath & "ErrorFile\" & strFileError
        Dim objFSO = Server.CreateObject("Scripting.FileSystemObject")
        TSwriteError = objFSO.CreateTextFile(strFileNameError, True)
        Dim iNum, iNum2, iNum3 As Integer
        MAKER2 = ""
        ii = 0
        DATEDT = DateTime.Now
        sDate = DATEDT.ToString("u")
        sDate = Left(sDate, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2)

        For i = 2 To UBound(aRecord) - 1

            Dim clause As String

            iNum2 = InStr(1, Trim(aRecord(i).ToString), ",")

            If iNum2 <> 0 Then

                For H As Integer = 1 To 2
                    iNum = InStr(iNum2, Trim(aRecord(i).ToString), Chr(34))
                    iNum2 = InStr(iNum + 1, Trim(aRecord(i).ToString), Chr(34))
                    iNum3 = iNum2 - iNum + 1

                    If iNum > 0 And iNum2 > 0 Then
                        clause = Mid((aRecord(i)), iNum, iNum3)
                        If InStr(1, clause, ",") > 0 Then
                            MAKER3 = Replace(clause, ",", "")
                            MAKER3 = Replace(MAKER3, Chr(34), "")
                            aRecord(i) = Replace(aRecord(i), clause, MAKER3)
                            MAKER2 = True
                        End If
                        iNum2 = iNum2 - 2
                    Else
                        Exit For
                    End If
                Next


            End If


            Detail = aRecord(i).Split(",")
            If (Detail(0) <> "") And Detail.Length > 3 Then
                ChKDigit1 = Left(Detail(5), 1)
                ChKDigit2 = Left(Detail(6), 1)
                If (ChKDigit1 <> Chr(34)) And (ChKDigit2 <> Chr(34)) And (Right(Trim(Detail(2)), 1) <> "'") And ChKDigit1 <> "" And ChKDigit2 <> "" Then
                    SeqNo = Detail(0)
                    PartNo1 = Trim(Detail(1))
                    LL = InStr(PartNo1, "#")
                    If LL > 0 Then
                        PartNo = Left(PartNo1, LL - 1)
                    Else
                        PartNo = PartNo1
                    End If
                    PartNo = Replace(PartNo, "�", "")

                    If Len(Trim(Detail(2))) = 5 Or Len(Trim(Detail(2))) = 3 Or Trim(Detail(2)) = "" Or Detail(2) Is DBNull.Value Then

                        MSFX = Left(Trim(Detail(2)), 2)
                        YSFX = Right(Trim(Detail(2)), 1)
                        suffix = MSFX & YSFX
                        ModelYear = Trim(Detail(3))
                        Ftype = Trim(Detail(4))

                        If Trim(Detail(5)) <> 0 Then
                            Fcost = Format(CDbl(Detail(5)), "#.##")
                        Else
                            Fcost = Trim(Detail(5))
                        End If

                        CuWeight = Trim(Detail(6))

                        '---PAN UPDATE 17/10/2013---
                        If Detail.Length > 7 Then 'check column 

                            Dim p As String = Detail(7).ToString.Trim
                            If p = "" Then
                                PACKING = 0
                            Else
                                If IsNumeric(Detail(7)) Then
                                    PACKING = Detail(7)
                                Else
                                    PACKING = 0
                                End If
                            End If

                            If IsNumeric(Detail(8)) And Detail(8).ToString.Trim <> "" Then
                                BASEFC = Trim(Detail(8))
                            Else
                                BASEFC = 0
                            End If

                            If IsNumeric(Detail(9)) And Detail(9).ToString.Trim <> "" Then
                                COPBASE = Trim(Detail(9))
                            Else
                                COPBASE = 0
                            End If

                            If IsNumeric(Detail(10)) And Detail(10).ToString.Trim <> "" Then
                                COPADJ = Trim(Detail(10))
                            Else
                                COPADJ = 0
                            End If

                            If IsNumeric(Detail(11)) And Detail(11).ToString.Trim <> "" Then
                                EXCH = Trim(Detail(11))
                            Else
                                EXCH = 0
                            End If


                            QID = IIf(Detail(12) IsNot DBNull.Value, Trim(Detail(12)), "")
                            MAKER = Left(IIf(Detail(13) IsNot DBNull.Value, Trim(Detail(13)), ""), 10)
                            MODEL = Left(IIf(Detail(14) IsNot DBNull.Value, Trim(Detail(14)), ""), 20)
                            MODEL = Replace(MODEL, Chr(13), "")

                            If IsNumeric(Detail(15)) And Detail(15).ToString.Trim <> "" Then
                                markup = Trim(Detail(15))
                            Else
                                markup = 0
                            End If

                        End If
                        Dim Effdate2 As String

                        If Mid(EffDate, 5, 2) = "12" Then
                            Effdate2 = Left(EffDate, 4) + 1
                            Effdate2 = Effdate2 + "01"
                        Else
                            Effdate2 = Left(EffDate, 6) + 1
                        End If

                        ErrorFlag = ""
                        CheckProd() 'check data in vdma
                        If cntRec3 = 0 Then
                            ErrorFlag = "No Master"
                        Else
                            CheckData() 'check data in vdmy & tmvdmc

                            If CntRec = 0 Then    ' if No has Data in VDMY
                                strDetail = "'" & Factory & "'," & ModelYear & "," & _
                                "'" & PartNo & "'," & cntRec3 & "," & _
                                "'" & MSFX & "','" & YSFX & "'," & EffDate & _
                                "," & Fcost & ",'BAH'," & CuWeight & ",'" & uom & "','0'," & _
                                sDate & ",'" & UserId & Ftype & "',0,'',0,'',0)"
                                SaveToDB2(strDetail)
                            Else

                                If strdate <= Effdate2 & 10 Then 'check month >> Update ได้เฉพาะเดือนเดียวกัน
                                    UPDATEFC()
                                End If
                            End If

                            If Detail.Length > 7 Then

                                If markup = 0 And CntRec2 = 1 Then 'ถ้าตอนแรกมี markup แต่ Update เป็นไม่มี markup ให้ลบออก 25/6/14 PAN
                                    deletedata()

                                ElseIf markup > 0 Then 'ถ้าเป็น Export/TAP-TAP ไม่ต้องคำนวณราคา 29/5/14 PAN

                                    If CntRec2 = 0 Then ' if No has Data in TMVDMC
                                        SAVETMVDMC()
                                    Else


                                        If strdate <= Effdate2 & 10 Then 'check month >> Update ได้เฉพาะเดือนเดียวกัน
                                            UPDATETMVDMC()
                                        End If
                                    End If

                                End If


                            End If


                        End If
                    Else
                        suffix = Detail(2)
                        ErrorFlag = "Wrong Suffix"
                    End If

                End If



                'CREATE ERRFILE >> NEW,DUPLICATE,COMPLETE
                If (ErrorFlag <> "") And ErrorFlag <> "INSERT ERROR" Then
                    If ii = 0 Then
                        strDat = "Factory," & Factory & ", Effective Date," & EffDate & ",UOM," & uom
                        TSwriteError.Write(strDat)
                        TSwriteError.WriteLine()
                        strDat = "No,Part no,SFX,Model,TYPE,F/C,Cu Weight,Packing,Base FC,Copp base,Copp Adj,Exch Rate,QID,Maker,Model,MarkUp,Remark"
                        TSwriteError.Write(strDat)
                        TSwriteError.WriteLine()
                    End If
                    strDat = SeqNo & "," & PartNo & ",'" & suffix & "," & ModelYear & "," & Ftype & "," & Fcost & "," & CuWeight & "," & PACKING & "," & BASEFC & "," & COPBASE & "," & COPADJ & "," & EXCH & "," & QID & "," & MAKER & "," & MODEL & "," & markup & "," & ErrorFlag & ""
                    TSwriteError.Write(strDat)
                    TSwriteError.WriteLine()
                    ii = ii + 1

                End If
            Else
                'If (Trim(Detail(1)) <> "") Then
                '    strDat = aRecord(i) & "," & "Format Error"
                '    TSwriteError.Write(strDat)
                '    TSwriteError.WriteLine()
                '    ii = ii + 1
                'End If
                Errnum += 1
            End If
            strDetail = ""

            If Errnum = 5 Then
                Exit For
            End If
        Next

        calculate()
        objStreamReader.CLOSE()
        TSwriteError.Close()
        oConn400.Close()
        Exit Sub

msgerr:
        message.Text = "Wrong Format!!!!!!!!"
        objStreamReader.CLOSE()
        TSwriteError.Close()
        oConn400.Close()

        Exit Sub
        Response.Write(UBound(aRecord))
        Response.End()
    End Sub
    Sub deletedata() '25/6/14
        Dim strdel As String
        strdel = "delete from " & StrLibPRICE & "  "
        strdel += " WHERE TMPROD='" & PartNo & "' AND TMMKSF='" & MSFX & "'  "
        strdel += " AND TMYZSF='" & YSFX & "'  AND TMSTDT='" & EffDate & "' AND TMFACT='" & Factory & "'"
        Dim CMD As New OleDbCommand(strdel, oConn400)
        Dim cntdel = CMD.ExecuteNonQuery()

        If cntdel = 0 Then
            message.Text = "Delete Data Failed"
            Exit Sub
        Else
            ErrorFlag = "Delete Complete"
        End If
    End Sub

    Private Sub Upload_Click(ByVal source As Object, ByVal e As EventArgs) Handles upload.ServerClick
        If Not (uploadedFile.PostedFile Is Nothing) Then
            Dim postedFile = uploadedFile.PostedFile
            Dim filename As String = Path.GetFileName(postedFile.FileName)
            Dim contentType As String = postedFile.ContentType
            Dim contentLength As Integer = postedFile.ContentLength
            Dim FileType As String = UCase(Left(filename, 2))
            Dim FileType1 As String = UCase(Right(filename, 3))
            File_Type = UCase(Mid(filename, 7, 2))
            If Len(txtUserCode.Value) < 9 Then
                UserId = Trim(txtUserCode.Value) & Space((9 - Len(Trim(txtUserCode.Value))))
            Else
                UserId = txtUserCode.Value
            End If

            If (FileType = "FC") And (FileType1 = "CSV") Then
                Factory = Mid(filename, 3, 4)
                setfact()

                Call ManageFile.DeleteFileName(savePath & "Data/")
                postedFile.SaveAs(savePath & "/Data/" & filename)
                SplitToArray(savePath & "/Data/" & filename)


                message.Text = message.Text & postedFile.Filename & " uploaded" & _
                "&nbsp; content length:  " & contentLength.ToString() & "<br>"
            Else

                message.Text = "Invalid File Name = " & filename & " <BR> "
            End If
        End If
    End Sub
    Sub setfact()

        Select Case Factory
            Case "32T1"
                If File_Type = "EX" Then
                    StrLib = "VD1TAPBE.VDMY" '"#PAN.VDMYP"
                    StrLibPart = "VD1TAPBE.VDMA"
                Else
                    StrLib = "VD1TAPBD.VDMY" '"#PAN.VDMYP"
                    StrLibPart = "VD1TAPBD.VDMA"
                End If
                savePath = Server.MapPath("Uploads/OtherJob/32T1/")
            Case "32T2"
                StrLib = "VD1TAPC.VDMY" '"#PAN.VDMYP"
                StrLibPart = "VD1TAPC.VDMA"
                savePath = Server.MapPath("Uploads/OtherJob/32T2/")
            Case "32T3"
                StrLib = "VD1TAPP.VDMY" '"#PAN.VDMYP"
                StrLibPart = "VD1TAPP.VDMA"
                savePath = Server.MapPath("Uploads/OtherJob/32T3/")
        End Select

    End Sub

    Sub GetAllFiles()
        Dim oFS, oFolder, oFile
        Dim FILE_FOLDER As String = savePath & "ErrorFile\"
        Dim dtshow As New DataTable
        Dim drshow As DataRow

        dtshow.Columns.Add("strFileName")
        dtshow.Columns.Add("strFileType")
        dtshow.Columns.Add("strFileSize")
        dtshow.Columns.Add("strFileDtMod")
        dtshow.Columns.Add("Link")
        oFS = Server.CreateObject("Scripting.FileSystemObject")
        oFolder = oFS.getFolder(FILE_FOLDER)   'Set Folder Object To Proper File Directory

        Dim intCounter
        intCounter = 0

        For Each oFile In oFolder.Files 'CREATE ERRFILE

            If Left(oFile.Name, 2) = "FC" Then
                drshow = dtshow.NewRow
                drshow("strFileName") = oFile.Name
                drshow("strFileType") = "CSV"
                drshow("strFileSize") = oFile.Size
                drshow("strFileDtMod") = oFile.DateLastModified
                drshow("Link") = "uploads\otherjob\" & Factory & "\Errorfile\" & oFile.Name
                dtshow.Rows.Add(drshow)
            End If


        Next

        grdshow.DataSource = dtshow
        grdshow.DataBind()

        lblmsg.Text = oFolder.Files.Count & " Files Available"

    End Sub

    Sub calculate() 'PAN UPDATE -PACKING 15/10/14

        Dim strSQL As String
        strSQL = "UPDATE " & StrLibPRICE & " SET tmcffg='C',TMPRIC=round((TMUPTB * (TMMARK / 100)) + ((TMCOPA - TMCOPB) / 1000000) * TMEXCH * TMCOPP-(tmpbas*tmmark/100),2),TMPPRI=round((TMPBAS*TMMARK)/100,2) "
        strSQL += " WHERE TMSTDT='" & EffDate & "' AND TMFACT='" & Factory & "' AND TMPRIC=0 AND TMPPRI=0"
        Dim CMD As New OleDbCommand(strSQL, oConn400)
        CntRec = CMD.ExecuteNonQuery()


    End Sub
    Function chkmarkup()

        strCmd = "select tsmark from " & StrLibTSPRATE & " where tsprod='" & PartNo & "' and tsfact='" & Factory & "'"
        Dim da As New OleDbDataAdapter(strCmd, oConn400)
        Dim dt As New DataTable
        da.Fill(dt)
        Return dt
    End Function

    Protected Sub LinkButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkButton1.Click
        Dim filename As String
        filename = "FC32T2 5.5.14 TEST.csv"
        Response.Redirect("~\Docs\" & filename)
    End Sub
End Class
