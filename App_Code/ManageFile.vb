Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Public Class ManageFile
    Public Shared Function DeleteFileName(ByVal sPath As String)
        Dim dt As New Date
        Dim sDate, MMYY_Date, MMYY_FILE As String
        Dim ii As Integer
        ii = 1
        dt = DateTime.Now
        sDate = dt.ToString("u")
        sDate = Left(sDate, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2)
        '	MMYY_Date = LEFT(sDate,4) & mid(sDate,6,2) 
        MMYY_Date = Left(sDate, 6)
        Dim ObjDir As New DirectoryInfo(sPath)
        Dim FileName As FileInfo
        For Each FileName In ObjDir.GetFiles("*.*")
            MMYY_FILE = Year(FileName.CreationTime) & Format(Month(FileName.CreationTime), "0#")
            If MMYY_FILE < MMYY_Date Then
                FileName.Delete()
            End If
        Next
    End Function

    Public Shared Function GenFileName(ByVal PreFixFile As String, ByVal TypeFile As String, ByVal sPath As String, ByVal pFile As String)
        Dim dt As New Date
        ''   Dim sDate, strcmd As String
        Dim sDate As String
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
    Public Shared Function getpermis(ByVal menutype As String, ByVal userid As String)
        Dim strsql As String
        strsql = " select *  from websalepermis where sysname='TAPCOST' and menuID='" & menutype & "' and uid='" & userid & "'" '   YICREPORT / YICMAIN
        Dim DA As New OleDbDataAdapter(strsql, Classcon.OpenSqlConn())
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT.Rows.Count
    End Function

End Class
