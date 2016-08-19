Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports System.Windows.Forms.Design
Imports Syncfusion.XlsIO

Public Class CommonFunction
    Inherits FolderNameEditor
    Private mUserName As String
    Private mUserID As String
    Private mUserPWD As String
    Private _Description As String = "Please select a directory below:"
    Private _Path As String = String.Empty
    Private FuncBrowse As New Syncfusion.Windows.Forms.FolderBrowser()

    Public Property Username() As String
        Get
            Return mUserName
        End Get
        Set(ByVal value As String)
            mUserName = value
        End Set
    End Property
    Public ReadOnly Property Path() As String
        Get
            Return _Path
        End Get
    End Property
    Public Property Description() As String
        Get
            Return _Description
        End Get
        Set(ByVal Value As String)
            _Description = Value
        End Set
    End Property

    Public Function ShowBrowser() As System.Windows.Forms.DialogResult
        With FuncBrowse
            .Description = _Description
            .StartLocation = Syncfusion.Windows.Forms.FolderBrowserFolder.MyDocuments
            Dim result As Windows.Forms.DialogResult = .ShowDialog
            If result = Windows.Forms.DialogResult.OK Then
                _Path = .DirectoryPath
            Else
                _Path = String.Empty
            End If
            Return result
            Return _Path
        End With
    End Function

    Public Shared Sub DeleteFileName(ByVal Path As String) 'Del 1 Month
        Dim dt As New Date
        Dim sDate, MMYY_Date, MMYY_FILE As String
        Dim ii As Integer
        ii = 1
        dt = DateTime.Now
        sDate = dt.ToString("u")
        sDate = Left(sDate, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2)
        MMYY_Date = Left(sDate, 6)  'Create Year,Month Now
        Dim ObjDir As New DirectoryInfo(Path)
        Dim FileName As FileInfo
        For Each FileName In ObjDir.GetFiles("*.*")
            MMYY_FILE = Year(FileName.CreationTime) & Format(Month(FileName.CreationTime), "0#")   'Create Year,Month of .Getfiles
            If MMYY_FILE < MMYY_Date Then
                FileName.Delete()
            End If
        Next
    End Sub

    Public Sub DelFileName(ByVal Path As String)   'Del 15 days
        Dim dt As New Date
        Dim sDate As String
        Dim ii, MMYY_Date, MMYY_FILE As Integer
        ii = 1
        dt = DateTime.Now
        sDate = dt.ToString("u")
        MMYY_Date = Left(sDate, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2) '20081201
        '= 'Left(sDate, 8)  'Create Year,Month Now
        Dim ObjDir As New DirectoryInfo(Path)
        Dim FileName As FileInfo
        For Each FileName In ObjDir.GetFiles("*.*")
            MMYY_FILE = Year(FileName.CreationTime) & Format(Month(FileName.CreationTime), "0#") & Day(FileName.CreationTime)   'Create Year,Month of .Getfiles
            If (MMYY_Date - MMYY_FILE) >= 15 Then
                FileName.Delete()
            End If
        Next
    End Sub

    Public Sub DelDirName(ByVal sPath As String, ByVal pType As String)
        Dim dt As New Date
        Dim sDate As String
        Dim ii As Integer
        Dim MMYY_Date, MMYY_FILE As String

        ii = 1
        dt = DateTime.Now
        sDate = dt.ToString("u")
        sDate = Left(sDate, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2)
        '	MMYY_Date = LEFT(sDate,4) & mid(sDate,6,2) 
        MMYY_Date = Left(sDate, 6)
        Dim ObjDir As New DirectoryInfo(sPath & pType)
        Dim ObjFile As New DirectoryInfo(sPath & pType)
        For Each ObjFile In ObjDir.GetDirectories()
            MMYY_FILE = Year(ObjFile.CreationTime) & Format(Month(ObjFile.CreationTime), "0#")
            If MMYY_FILE < MMYY_Date Then
                ObjFile.Delete()
            End If
        Next
    End Sub
    Public Sub DeleteAllFile(ByVal Path As String)
        Dim dt As New Date
        Dim sDate, MMYY_Date As String
        Dim ii As Integer
        ii = 1
        dt = DateTime.Now
        sDate = dt.ToString("u")
        sDate = Left(sDate, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2)
        '	MMYY_Date = LEFT(sDate,4) & mid(sDate,6,2) 
        MMYY_Date = Left(sDate, 6)
        Dim ObjDir As New DirectoryInfo(Path)
        Dim FileName As FileInfo
        For Each FileName In ObjDir.GetFiles("*.*")
            FileName.Delete()
        Next
    End Sub
    Public Shared Function GenFileName(ByVal PreFixFile As String, ByVal TypeFile As String, ByVal sPath As String, ByVal pFile As String)
        Dim dt As New Date
        Dim sDate As String = ""
        Dim strcmd As String = ""
        Dim ii As Integer

        Call DeleteFileName(sPath)
        ii = 1
        dt = DateTime.Now
        sDate = dt.ToString("u")
        sDate = Left(sDate, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2)
        GenFileName = PreFixFile & TypeFile & "_" & sDate & "_" & Format(ii, "0##") & pFile
        Dim ObjDir As New DirectoryInfo(sPath)
        Dim ObjFile As FileInfo
        For Each ObjFile In ObjDir.GetFiles("*" & PreFixFile & "*" & pFile)
            If GenFileName = ObjFile.Name Then
                ii = ii + 1
                GenFileName = PreFixFile & TypeFile & "_" & sDate & "_" & _
                              Format(ii, "0##") & pFile
            End If
        Next

    End Function

    Public Shared Function GenDirName(ByVal sPath As String, ByVal pType As String) As String
        Dim dt As New Date
        Dim sDate As String
        Dim ii As Integer
        ii = 1
        dt = DateTime.Now
        sDate = dt.ToString("u")
        sDate = Left(sDate, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2)

        Dim ObjDir As New DirectoryInfo(sPath & pType)
        Dim ObjFile As DirectoryInfo

        GenDirName = pType & sDate & "_" & Format(ii, "0##")
        For Each ObjFile In ObjDir.GetDirectories()

            If GenDirName = ObjFile.Name Then
                ii = ii + 1
                GenDirName = pType & sDate & "_" & Format(ii, "0##")
            End If
        Next
        ObjDir.CreateSubdirectory(GenDirName)
    End Function

    Public Sub del_Old_Directory(ByVal sPath As String)
        Dim dirName As String
        Dim curMon, dirMon As Integer

        Dim objdir As New DirectoryInfo(sPath)
        Dim tmpdir As DirectoryInfo
        'Response.Write(sPath & "<BR>")
        curMon = Month(DateTime.Now)

        For Each tmpdir In objdir.GetDirectories()
            dirName = tmpdir.Name
            dirMon = Val(Mid(dirName, 8, 2))
            'Response.Write(dirMon & "D C" & curMon & "<BR>")
            If dirMon < curMon Then
                'Response.Write(dirMon & "D C" & curMon & "<BR>")
                tmpdir.Delete(1)
            End If
        Next
    End Sub

    Public Function GenNoOfDirName(ByVal sPath As String) As String

        Dim objdir As New DirectoryInfo(sPath)

        Dim tmpdir As DirectoryInfo
        Dim dd, mm, NoFolder As Integer
        Dim strcmd, dirName As String
        Dim DDtoday, MMtoday As Integer
        Dim maxNoFolder As Integer = 0

        DDtoday = Day(DateTime.Now)
        MMtoday = Month(DateTime.Now)

        For Each tmpdir In objdir.GetDirectories()
            dirName = tmpdir.Name
            NoFolder = Val(Right(dirName, 3))
            strcmd = Mid(dirName, 4, 8)                         'EXP20070511_001 >>>>  20070511				

            dd = Val(Right(strcmd, 2))                         'day of folder
            mm = Val(Mid(strcmd, 5, 2))                         'mon of folder

            If mm = MMtoday Then
                If dd = DDtoday Then
                    If maxNoFolder < NoFolder Then
                        maxNoFolder = NoFolder
                    End If
                End If
            End If
        Next

        Dim numstr As String
        maxNoFolder = maxNoFolder + 1

        numstr = maxNoFolder.ToString("000")

        Return numstr
    End Function

    Function CheckOfDirName(ByVal sPath As String) As Boolean

        Dim objdir As New DirectoryInfo(sPath)

        Dim tmpdir As DirectoryInfo
        Dim dd, mm As Integer
        'Dim dd, mm, yy, NoFolder As Integer
        Dim strcmd, dirName As String
        Dim DDtoday, MMtoday As Integer
        Dim maxNoFolder As Integer = 0
        Dim HaveDir As Boolean
        DDtoday = Day(DateTime.Now)
        MMtoday = Month(DateTime.Now)
        HaveDir = False
        For Each tmpdir In objdir.GetDirectories()
            dirName = tmpdir.Name
            'NoFolder =val(right(dirName,3))
            strcmd = Mid(dirName, 4, 8)                         'EXP20070511_001 >>>>  20070511				

            dd = Val(Right(strcmd, 2))                         'day of folder
            mm = Val(Mid(strcmd, 5, 2))                         'mon of folder

            If mm = MMtoday Then
                If dd = DDtoday Then
                    HaveDir = True
                End If
            End If
        Next

        'dim numstr as String
        'maxNoFolder=maxNoFolder+1

        'numstr = maxNoFolder.toString("000")

        Return HaveDir
    End Function

    Function CheckOfDirUser(ByVal sPath As String, ByVal SubUser As String) As Boolean

        Dim objdir As New DirectoryInfo(sPath)
        Dim strUser, dirName As String
        Dim tmpdir As DirectoryInfo
        Dim HaveDir As Boolean
        HaveDir = False
        For Each tmpdir In objdir.GetDirectories()
            dirName = tmpdir.Name
            strUser = Trim(Right(dirName, 10))
            If strUser = SubUser Then
                HaveDir = True
            End If
        Next

        'dim numstr as String
        'maxNoFolder=maxNoFolder+1

        'numstr = maxNoFolder.toString("000")

        Return HaveDir
    End Function

    Public Function genDir(ByVal sPath As String, ByVal strFile As String) As String

        'Dim dirName, pType As String
        Dim curDate As Date
        Dim sDate As String
        'Dim tmpdir As DirectoryInfo
        Dim curDay As Integer
        'dim  numTmp, dirDay As Integer
        curDay = Day(DateTime.Now)
        'Response.Write(curDay & "<BR>" & sPath &  "<BR>"  & strFile &  "<BR>" ) 
        Dim objdir As New DirectoryInfo(sPath)

        curDate = DateTime.Now                  ' date format are dd/m/yyyy 3:55:34 PM
        sDate = curDate.ToString("u")           ' string format are yyyy-mm-dd 3:55:34 PM
        sDate = sDate.Substring(0, 10)           ' string format are yyyy-mm-dd 
        sDate = sDate.Replace("-", "")           ' string format are yyyymmdd 

        strFile = strFile & sDate & "_" & GenNoOfDirName(sPath)

        objdir.CreateSubdirectory(strFile)

        Return strFile

    End Function

    Function CheckDir(ByVal sPath As String, ByVal strFile As String, ByVal SubUser As String) As String

        'Dim dirName, pType, strcmd As String
        Dim curDate As Date
        Dim sDate, strFile1 As String
        ' Dim tmpdir As DirectoryInfo
        Dim curDay As Integer ', numTmp, dirDay
        curDay = Day(DateTime.Now)

        Dim objdir As New DirectoryInfo(sPath)

        curDate = DateTime.Now                  ' date format are dd/m/yyyy 3:55:34 PM
        sDate = curDate.ToString("u")           ' string format are yyyy-mm-dd 3:55:34 PM
        sDate = sDate.Substring(0, 10)           ' string format are yyyy-mm-dd 
        sDate = sDate.Replace("-", "")           ' string format are yyyymmdd 
        strFile = strFile & sDate
        '	response.write (SubUser)
        '	response.end()
        If Not CheckOfDirName(sPath) Then
            objdir.CreateSubdirectory(strFile)
        Else
            If (SubUser <> "") Then
                If Not CheckOfDirUser(sPath & strFile, SubUser) Then
                    Dim objdir1 As New DirectoryInfo(sPath & strFile)
                    strFile1 = SubUser
                    objdir1.CreateSubdirectory(strFile1)
                End If
            End If
        End If
        strFile = strFile & "\" & SubUser
        Return strFile
    End Function
    Public Sub GetBeginCol(ByVal OldProdMonth As String, ByVal NewProdMonth As String, ByRef Nocol As Integer)
        Dim OldYear, NewYear As String
        Dim OldMonth As Integer = 0
        Dim NewMonth As Integer = 0
        Dim DifOldMonth As Integer = 0
        OldYear = Left(OldProdMonth, 4)
        NewYear = Left(NewProdMonth, 4)
        OldMonth = CInt(Right(OldProdMonth, 2))
        NewMonth = CInt(Right(NewProdMonth, 2))
        If NewYear > OldYear Then
            NewMonth = 12 + CInt(Right(NewProdMonth, 2))
        End If
        Nocol = (NewMonth - OldMonth) + 1
    End Sub
    Public Shared Sub CreateFileNameSession(ByRef idx As Integer, ByRef dt As DataTable, _
                                            ByRef drow As DataRow, ByVal Filename As String)
        If idx = 0 Then
            dt.Columns.Add(New DataColumn("FileName", GetType(String)))
        End If
        drow = dt.NewRow()
        drow("Filename") = Filename
        dt.Rows.Add(drow)
        idx += 1
    End Sub

    Public Shared Sub SaveTmpFile(ByVal CurTempPath As String, ByRef Filename As String)
        Dim iRun As Integer = 1
        Dim ObjDir As New DirectoryInfo(CurTempPath)
        Dim ObjFile As FileInfo
        For Each ObjFile In ObjDir.GetFiles("*.*")
            If ObjFile.Name = Filename Then
                iRun = iRun + 1
                Filename = Filename.Insert(Len(Trim(Filename)) - 4, "(" & Format(iRun, "0#") & ")")
            End If
        Next
    End Sub

    Public Shared Function CheckDate(ByRef strDate As String, ByRef oMSG As String, ByVal iFormat As String, ByVal iCase As String)
        Dim iDate As String = ""
        Dim iMonth As String = ""
        Dim iYear As String = ""
        If strDate.Length = 8 Then
            Select Case iFormat
                Case "ddmmyyyy"
                    iDate = Left(strDate, 2)
                    iMonth = Mid(strDate, 3, 2)
                    iYear = Right(strDate, 4)
                Case "yyyymmdd"
                    iDate = Right(strDate, 2)
                    iMonth = Mid(strDate, 5, 2)
                    iYear = Left(strDate, 4)
            End Select
            If iDate > 31 Or iDate < 1 Then
                oMSG = "DATE INVALID, DATE BETWEEN 1-31"
            End If
            If iMonth > 12 Or iMonth < 1 Then
                oMSG = "MONTH INVALID, MONTH BETWEEN 1-12"
            End If
            Select Case iCase
                Case "YTHA1"
                    'Format 01122009 --> 01122552
                    If iYear < 2500 Then
                        iYear += 543
                    End If
                    strDate = iDate & iMonth & iYear
                Case "YENG1"
                    'Format 01122552 --> 01122009
                    If iYear > 2500 Then
                        iYear -= 543
                    End If
                    strDate = iDate & iMonth & iYear
                Case "YTHA2"
                    'Format 20091201 --> 25521201
                    If iYear < 2500 Then
                        iYear += 543
                    End If
                    strDate = iYear & iMonth & iDate
                Case "YENG2"
                    'Format 25521201 --> 20091201
                    If iYear > 2500 Then
                        iYear -= 543
                    End If
                    strDate = iYear & iMonth & iDate
            End Select
        End If
        Return strDate
    End Function
    Public Shared Function CheckYear(ByRef iYear As String, ByVal iCase As String)
        Select Case iCase
            Case "YTHA"
                If iYear < 2500 Then
                    iYear += 543
                End If
            Case "YENG"
                If iYear > 2500 Then
                    iYear -= 543
                End If
        End Select
        Return iYear
    End Function
    Public Shared Function DiffDate_FinalDay(ByVal strDate As String, ByVal iFormat As String, ByVal iDiff As Integer)
        'หาวันที่ปัจจุบัน + 10
        Dim strResult As String = ""
        Dim iDate As String = ""
        Dim iMONTH As String = ""
        Dim iYEAR As String = ""
        If strDate.Length = 8 Then
            Select Case iFormat
                Case "ddmmyyyy"
                    iDate = Left(strDate, 2)
                    iMONTH = Mid(strDate, 3, 2)
                    iYEAR = Right(strDate, 4)
                Case "yyyymmdd"
                    iDate = Right(strDate, 2)
                    iMONTH = Mid(strDate, 5, 2)
                    iYEAR = Left(strDate, 4)
            End Select

            Dim Final_Day As String = ""
            Select Case iMONTH
                Case "01", "03", "05", "07", "08", "10", "12"
                    Final_Day = 31
                Case "04", "06", "09", "11"
                    Final_Day = 30
                Case "02"
                    Final_Day = 28
            End Select
            If (strDate <= Final_Day) And (strDate >= (Final_Day - (iDiff - 1))) Then
                Dim Diff As String = Final_Day - strDate
                Diff = (iDiff - Diff)
                If iMONTH < 12 Then
                    strResult = iYEAR & iMONTH + 1 & Format(Diff, "0#")
                Else
                    strResult = iYEAR + 1 & iMONTH + 1 & Format(Diff, "0#")
                End If
            Else
                strResult += iYEAR & iMONTH & iDate + 10
            End If
        End If
        Return strResult
    End Function
    Public Shared Sub NextDate(ByRef strResult1 As String, ByRef strresult2 As String, ByVal strDate As String, ByVal iFormat As String)
        'หาวันที่ระหว่าง  + 1 เดือน  และ - 1 เดือน จากวันที่ปัจจุบัน
        Dim iDate As String = ""
        Dim iMONTH As String = ""
        Dim iYEAR As String = ""
        If strDate.Length = 8 Then
            Select Case iFormat
                Case "ddmmyyyy"
                    iDate = Left(strDate, 2)
                    iMONTH = Mid(strDate, 3, 2)
                    iYEAR = Right(strDate, 4)
                Case "yyyymmdd"
                    iDate = Right(strDate, 2)
                    iMONTH = Mid(strDate, 5, 2)
                    iYEAR = Left(strDate, 4)
            End Select

            If iMONTH > 1 And iMONTH < 12 Then
                strResult1 = iYEAR & iMONTH - 1 & iDate
                strresult2 = iYEAR & iMONTH + 1 & iDate
            ElseIf iMONTH = 12 Then
                strResult1 = iYEAR & iMONTH - 1 & iDate
                strresult2 = iYEAR + 1 & "01" & iDate
            ElseIf iMONTH = 1 Then
                strResult1 = iYEAR - 1 & 12 & iDate
                strresult2 = iYEAR & iMONTH + 1 & iDate
            End If
        End If
    End Sub

    Public Shared Sub ThaiMonth(ByRef strDate As String, ByRef Result As String)
        'Format DDMMYYY --> EX. 01122551
        Dim iMonth As String = ""
        If strDate.Length = 8 Then
            iMonth = Mid(strDate, 3, 2)
            Select Case iMonth
                Case "01"
                    iMonth = "มกราคม"
                Case "02"
                    iMonth = "กุมภาพันธ์"
                Case "03"
                    iMonth = "มีนาคม"
                Case "04"
                    iMonth = "เมษายน"
                Case "05"
                    iMonth = "พฤษภาคม"
                Case "06"
                    iMonth = "มิถุนายน"
                Case "07"
                    iMonth = "กรกฏาคม"
                Case "08"
                    iMonth = "สิงหาคม"
                Case "09"
                    iMonth = "กันยายน"
                Case "10"
                    iMonth = "ตุลาคม"
                Case "11"
                    iMonth = "พฤศจิกายน"
                Case "12"
                    iMonth = "ธันวาคม"
            End Select
        End If
        Result = Left(strDate, 2) & " " & iMonth & " " & Right(strDate, 4)
    End Sub
    Public Shared Function DaysofMonth(ByRef iMonth As Integer, ByRef Result As Integer)
        If iMonth > 0 Then
            Select Case iMonth
                Case 1, 3, 5, 7, 8, 10, 12
                    Result = 31
                Case 2
                    Result = 28
                Case 4, 6, 9, 11
                    Result = 30
            End Select
        End If
        Return Result
    End Function
    Public Shared Sub Customs_Company(ByVal Library As String, ByRef oTable As DataTable)
        Dim strSql As String = ""
        Dim Conn As New OleDb.OleDbConnection(Classconn.strCon400)
        Dim dt As New DataTable
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
        strSql = "SELECT * FROM " & Library & ".TGCO "
        Dim da As New OleDb.OleDbDataAdapter(strSql, Conn)
        da.Fill(dt)
        oTable = dt
    End Sub

    Public Shared Sub Customs_Type(ByVal Type As String, ByVal TypeDTPV As String, ByRef strType As String)
        Select Case Type
            Case "1"
                strType = "นำเข้าจากต่างประเทศ(IMPORT)"
            Case "2"
                Select Case TypeDTPV                                   'ประเภทรายการ
                    Case "2"
                        strType = "รับโอนจากคลังสินค้าทัณฑ์บน(BOND)"
                    Case "3"
                        strType = "รับโอนจาก 19 ทวิ(BIS 19)"
                    Case "4"
                        strType = "รับโอนจาก FZ(FZ)"
                    Case "5"
                        strType = "รับโอนจาก EPZ(EPZ)"
                    Case "6"
                        strType = "รับโอนจาก BOI(BOI)"
                    Case Else
                        strType = "รับโอนไม่ระบุประเภท"
                End Select
            Case "3"
                Select Case TypeDTPV                                    'ประเภทรายการ
                    Case "1"
                        strType = "โอนย้ายออก"
                    Case "2"
                        strType = "โอนย้ายออกจากคลังสินค้าทัณฑ์บน(BOND)"
                    Case "3"
                        strType = "โอนย้ายออกจาก 19 ทวิ(BIS 19)"
                    Case "4"
                        strType = "โอนย้ายออกจาก FZ(FZ)"
                    Case "5"
                        strType = "โอนย้ายออกจาก EPZ(EPZ)"
                    Case "6"
                        strType = "โอนย้ายออกจาก BOI(BOI)"
                    Case Else
                        strType = "โอนย้ายออกไม่ระบุประเภท"
                End Select
            Case "4"
                strType = "ชำระภาษีอากร"
            Case "5"
                strType = "ส่งออกต่างประเทศ"
            Case "6"
                strType = "ทำลาย/บริจาค"
        End Select
        If strType = "" Then
            strType = "ไม่ระบุประเภท"
        End If
    End Sub
    Public Shared Function RoundFloor(ByVal OriginNum As String) As Double
        Dim Num As Double = CDbl(OriginNum)
        If Right(OriginNum, 2) > 0 Then
            Num = CDbl(Mid(OriginNum, 1, (InStr(1, OriginNum, ".") - 1)))
        End If
        Return Num
    End Function

End Class
