Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb

Public Class dllYSIYSIV
    Private mYWBTCD As String
    Private mYWPLNR As String
    Private mYWPREC As String
    Private mYWPROC As String
    Private mYWRDOC As String
    Private mYWDOCK As String
    Private mYWINS As String
    Private mYWPKC As String
    Private mYWPKV As Decimal
    Private mYWPART As String
    Private mYWCHG As String
    Private mYWPNM As String
    Private mYWCALN As String
    Private mYWPDCD As String
    Private mYWPDTP As String
    Private mYWPDGP As String
    Private mYWODNO As String
    'Private mYWPMON As Decimal
    Private mYWIVNO As String
    Private mYWIVDT As Decimal
    Private mYWRCDT As Decimal
    Private mYWEFDT As Decimal
    Private mYWCFDT As Decimal
    Private mYWRQTY As Decimal
    Private mYWRPRC As Decimal
    Private mYWIPRC As Decimal
    Private mYWNPRC As Decimal
    Private mYWCPRC As Decimal
    Private mYWDFPR As Decimal
    Private mYWDRPR As Decimal
    Private mYWCRPR As Decimal
    Private mYWPACK As Decimal
    Private mYWFLAG As String
    Private mYWDOCN As String
    Private mYWCRBY As String
    Private mYWCRTD As Decimal
    Private mYWCRTT As Decimal
    Private mYWCMBY As String
    Private mYWCMDT As Decimal
    Private mYWCMTM As Decimal

    Private mAccStatus As Boolean

    Public Property YWBTCD() As String
        Get
            Return mYWBTCD
        End Get
        Set(ByVal value As String)
            mYWBTCD = value
        End Set
    End Property
    Public Property YWPLNR() As String
        Get
            Return mYWPLNR
        End Get
        Set(ByVal value As String)
            mYWPLNR = value
        End Set
    End Property
    Public Property YWPREC() As String
        Get
            Return mYWPREC
        End Get
        Set(ByVal value As String)
            mYWPREC = value
        End Set
    End Property
    Public Property YWPROC() As String
        Get
            Return mYWPROC
        End Get
        Set(ByVal value As String)
            mYWPROC = value
        End Set
    End Property
    Public Property YWRDOC() As String
        Get
            Return mYWRDOC
        End Get
        Set(ByVal value As String)
            mYWRDOC = value
        End Set
    End Property
    Public Property YWDOCK() As String
        Get
            Return mYWDOCK
        End Get
        Set(ByVal value As String)
            mYWDOCK = value
        End Set
    End Property
    Public Property YWINS() As String
        Get
            Return mYWINS
        End Get
        Set(ByVal value As String)
            mYWINS = value
        End Set
    End Property
    Public Property YWPKC() As String
        Get
            Return mYWPKC
        End Get
        Set(ByVal value As String)
            mYWPKC = value
        End Set
    End Property
    Public Property YWPKV() As Decimal
        Get
            Return mYWPKV
        End Get
        Set(ByVal value As Decimal)
            mYWPKV = value
        End Set
    End Property
    Public Property YWPART() As String
        Get
            Return mYWPART
        End Get
        Set(ByVal value As String)
            value = Left(value, value.LastIndexOf("-"))
            'value = (Replace(Trim(value), "-", ""))
            value = Replace(Trim(value), " ", "")
            mYWPART = value
        End Set
    End Property
    Public Property YWCHG() As String
        Get
            Return mYWCHG
        End Get
        Set(ByVal value As String)
            mYWCHG = value
        End Set
    End Property
    Public Property YWPNM() As String
        Get
            Return mYWPNM
        End Get
        Set(ByVal value As String)
            mYWPNM = value
        End Set
    End Property
    Public Property YWCALN() As String
        Get
            Return mYWCALN
        End Get
        Set(ByVal value As String)
            mYWCALN = value
        End Set
    End Property
    Public Property YWPDCD() As String
        Get
            Return mYWPDCD
        End Get
        Set(ByVal value As String)
            mYWPDCD = value
        End Set
    End Property
    Public Property YWPDTP() As String
        Get
            Return mYWPDTP
        End Get
        Set(ByVal value As String)
            mYWPDTP = value
        End Set
    End Property
    Public Property YWPDGP() As String
        Get
            Return mYWPDGP
        End Get
        Set(ByVal value As String)
            mYWPDGP = value
        End Set
    End Property
    Public Property YWODNO() As String
        Get
            Return mYWODNO
        End Get
        Set(ByVal value As String)

            If InStr(value, "TAP", CompareMethod.Text) > 0 Then
                mYWODNO = Mid(value, InStr(value, "TAP", CompareMethod.Text), Len(value))
            Else
                mYWODNO = value
            End If


        End Set
    End Property
    'Public Property YWPMON() As Decimal
    '    Get
    '        Return mYWPMON
    '    End Get
    '    Set(ByVal value As Decimal)
    '        mYWPMON = value
    '    End Set
    'End Property
    Public Property YWIVNO() As String
        Get
            Return mYWIVNO
        End Get
        Set(ByVal value As String)
            mYWIVNO = value
        End Set
    End Property
    Public Property YWIVDT() As Decimal
        Get
            Return mYWIVDT
        End Get
        Set(ByVal value As Decimal)
            mYWIVDT = value
        End Set
    End Property
    Public Property YWRCDT() As Decimal
        Get
            Return mYWRCDT
        End Get
        Set(ByVal value As Decimal)
            mYWRCDT = value
        End Set
    End Property
    Public Property YWEFDT() As Decimal
        Get
            Return mYWEFDT
        End Get
        Set(ByVal value As Decimal)
            mYWEFDT = value
        End Set
    End Property
    Public Property YWCFDT() As Decimal
        Get
            Return mYWCFDT
        End Get
        Set(ByVal value As Decimal)
            mYWCFDT = value
        End Set
    End Property
    Public Property YWRQTY() As Decimal
        Get
            Return mYWRQTY
        End Get
        Set(ByVal value As Decimal)
            mYWRQTY = value
        End Set
    End Property
    Public Property YWRPRC() As Decimal
        Get
            Return mYWRPRC
        End Get
        Set(ByVal value As Decimal)
            mYWRPRC = value
        End Set
    End Property
    Public Property YWIPRC() As Decimal
        Get
            Return mYWIPRC
        End Get
        Set(ByVal value As Decimal)
            mYWIPRC = value
        End Set
    End Property
    Public Property YWNPRC() As Decimal
        Get
            Return mYWNPRC
        End Get
        Set(ByVal value As Decimal)
            mYWNPRC = value
        End Set
    End Property
    Public Property YWCPRC() As Decimal
        Get
            Return mYWCPRC
        End Get
        Set(ByVal value As Decimal)
            mYWCPRC = value
        End Set
    End Property
    Public Property YWDFPR() As Decimal
        Get
            Return mYWDFPR
        End Get
        Set(ByVal value As Decimal)
            mYWDFPR = value
        End Set
    End Property
    Public Property YWDRPR() As Decimal
        Get
            Return mYWDRPR
        End Get
        Set(ByVal value As Decimal)
            mYWDRPR = value
        End Set
    End Property
    Public Property YWCRPR() As Decimal
        Get
            Return mYWCRPR
        End Get
        Set(ByVal value As Decimal)
            mYWCRPR = value
        End Set
    End Property
    Public Property YWPACK() As Decimal
        Get
            Return mYWPACK
        End Get
        Set(ByVal value As Decimal)
            mYWPACK = value
        End Set
    End Property
    Public Property YWFLAG() As String
        Get
            Return mYWFLAG
        End Get
        Set(ByVal value As String)
            mYWFLAG = value

        End Set
    End Property
    Public Property YWDOCN() As String
        Get
            Return mYWDOCN
        End Get
        Set(ByVal value As String)
            mYWDOCN = value
        End Set
    End Property
    Public Property YWCRBY() As String
        Get
            Return mYWCRBY
        End Get
        Set(ByVal value As String)
            mYWCRBY = value
        End Set
    End Property
    Public Property YWCRTD() As Decimal
        Get
            Return mYWCRTD
        End Get
        Set(ByVal value As Decimal)

            mYWCRTD = value
        End Set
    End Property
    Public Property YWCRTT() As Decimal
        Get
            Return mYWCRTT
        End Get
        Set(ByVal value As Decimal)
            mYWCRTT = value
        End Set
    End Property
    Public Property YWCMBY() As String
        Get
            Return mYWCMBY
        End Get
        Set(ByVal value As String)
            mYWCMBY = value
        End Set
    End Property
    Public Property YWCMDT() As Decimal
        Get
            Return mYWCMDT
        End Get
        Set(ByVal value As Decimal)
            mYWCMDT = value
        End Set
    End Property
    Public Property YWCMTM() As Decimal
        Get
            Return mYWCMTM
        End Get
        Set(ByVal value As Decimal)
            mYWCMTM = value
        End Set
    End Property

    Public ReadOnly Property AccStatus() As Boolean
        Get
            Return mAccStatus
        End Get
    End Property


    Public Sub insert()
        Dim strsql As String
        strsql = "INSERT INTO " & Classconn.LIBRARY_FILE & " "
        strsql += "( YWBTCD,YWPLNR,YWPREC,YWPROC,YWRDOC,YWDOCK,YWINS,YWPKC"
        strsql += ",YWPKV,YWPART,YWCHG,YWPNM,YWCALN,YWPDCD,YWPDTP,YWPDGP"
        strsql += ",YWODNO,YWIVNO,YWIVDT,YWRCDT,YWEFDT,YWCFDT,YWRQTY"           ' DELETE PmON 09/08/08
        strsql += ",YWRPRC,YWIPRC,YWNPRC,YWCPRC,YWDFPR,YWDRPR,YWCRPR,YWPACK"
        strsql += ",YWFLAG,YWDOCN,YWCRBY,YWCRTD,YWCRTT,YWCMBY,YWCMDT,YWCMTM ) "
        strsql += "VALUES (?,?,?,?,?,?,?,?,?,?,"
        strsql += "?,?,?,?,?,?,?,?,?,"                                        ' DELETE ? PmON 09/08/08
        strsql += "?,?,?,?,?,?,?,?,?,?,"
        strsql += "?,?,?,?,?,?,?,?,?,? )"

        Dim conn As New OleDbConnection(Classconn.strCon400)
        Dim insertDA As New OleDbDataAdapter
        conn.Open()
        Dim insertCMD As New OleDbCommand(strsql, conn)

        insertDA.InsertCommand = insertCMD

        insertCMD.Parameters.Add("YWBTCD", OleDbType.VarChar, 15).Value = Me.mYWBTCD
        insertCMD.Parameters.Add("YWPLNR", OleDbType.VarChar, 5).Value = Me.mYWPLNR
        insertCMD.Parameters.Add("YWPREC", OleDbType.VarChar, 6).Value = Me.mYWPREC
        insertCMD.Parameters.Add("YWPROC", OleDbType.VarChar, 6).Value = Me.mYWPROC
        insertCMD.Parameters.Add("YWRDOC", OleDbType.VarChar, 2).Value = Me.mYWRDOC
        insertCMD.Parameters.Add("YWDOCK", OleDbType.VarChar, 5).Value = Me.mYWDOCK
        insertCMD.Parameters.Add("YWINS", OleDbType.VarChar, 2).Value = Me.mYWINS
        insertCMD.Parameters.Add("YWPKC", OleDbType.VarChar, 5).Value = Me.mYWPKC
        insertCMD.Parameters.Add("YWPKV", OleDbType.Decimal, 5).Value = Me.mYWPKV
        insertCMD.Parameters.Add("YWPART", OleDbType.VarChar, 25).Value = Me.mYWPART
        insertCMD.Parameters.Add("YWCHG", OleDbType.VarChar, 3).Value = Me.mYWCHG
        insertCMD.Parameters.Add("YWPNM", OleDbType.VarChar, 25).Value = Me.mYWPNM
        insertCMD.Parameters.Add("YWCALN", OleDbType.VarChar, 6).Value = Me.mYWCALN
        insertCMD.Parameters.Add("YWPDCD", OleDbType.VarChar, 4).Value = Me.mYWPDCD
        insertCMD.Parameters.Add("YWPDTP", OleDbType.VarChar, 3).Value = Me.mYWPDTP
        insertCMD.Parameters.Add("YWPDGP", OleDbType.VarChar, 2).Value = Me.mYWPDGP
        insertCMD.Parameters.Add("YWODNO", OleDbType.VarChar, 20).Value = Me.mYWODNO
        'insertCMD.Parameters.Add("YWPMON", OleDbType.Decimal, 6).Value = Me.mYWPMON
        insertCMD.Parameters.Add("YWIVNO", OleDbType.VarChar, 15).Value = Me.mYWIVNO
        insertCMD.Parameters.Add("YWIVDT", OleDbType.Decimal, 8).Value = Me.mYWIVDT
        insertCMD.Parameters.Add("YWRCDT", OleDbType.Decimal, 8).Value = Me.mYWRCDT
        insertCMD.Parameters.Add("YWEFDT", OleDbType.Decimal, 8).Value = Me.mYWEFDT
        insertCMD.Parameters.Add("YWCFDT", OleDbType.Decimal, 8).Value = Me.mYWCFDT
        insertCMD.Parameters.Add("YWRQTY", OleDbType.Decimal, 6).Value = Me.mYWRQTY
        insertCMD.Parameters.Add("YWRPRC", OleDbType.Double, 15).Value = Me.mYWRPRC
        insertCMD.Parameters.Add("YWIPRC", OleDbType.Double, 15).Value = Me.mYWIPRC
        insertCMD.Parameters.Add("YWNPRC", OleDbType.Double, 15).Value = Me.mYWNPRC
        insertCMD.Parameters.Add("YWCPRC", OleDbType.Double, 15).Value = Me.mYWCPRC
        insertCMD.Parameters.Add("YWDFPR", OleDbType.Double, 15).Value = Me.mYWDFPR
        insertCMD.Parameters.Add("YWDRPR", OleDbType.Double, 15).Value = Me.mYWDRPR
        insertCMD.Parameters.Add("YWCRPR", OleDbType.Double, 15).Value = Me.mYWCRPR
        insertCMD.Parameters.Add("YWPACK", OleDbType.Decimal, 5).Value = Me.mYWPACK
        insertCMD.Parameters.Add("YWFLAG", OleDbType.VarChar, 1).Value = Me.mYWFLAG
        insertCMD.Parameters.Add("YWDOCN", OleDbType.VarChar, 15).Value = Me.mYWDOCN
        insertCMD.Parameters.Add("YWCRBY", OleDbType.VarChar, 10).Value = Me.mYWCRBY
        insertCMD.Parameters.Add("YWCRTD", OleDbType.Decimal, 8).Value = Me.mYWCRTD
        insertCMD.Parameters.Add("YWCRTT", OleDbType.Decimal, 6).Value = Me.mYWCRTT
        insertCMD.Parameters.Add("YWCMBY", OleDbType.VarChar, 10).Value = Me.mYWCMBY
        insertCMD.Parameters.Add("YWCMDT", OleDbType.Decimal, 8).Value = Me.mYWCMDT
        insertCMD.Parameters.Add("YWCMTM", OleDbType.Decimal, 6).Value = Me.mYWCMTM



        Try
            insertCMD.ExecuteNonQuery()
            Me.mAccStatus = True
        Catch ex As System.Exception
            Me.mAccStatus = False
            insertDA.InsertCommand.Connection.Close()
            conn.Close()
        End Try


     

        insertDA.InsertCommand.Connection.Close()
        conn.Close()



    End Sub
    Public Sub delete()
        Dim strsql As String
        strsql = "delete from " & Classconn.LIBRARY_FILE & " "
        'strsql += "Where YWBTCD = ? and YWPLNR = ? "

        Dim conn As New OleDbConnection(Classconn.strCon400)
        Dim delDA As New OleDbDataAdapter
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.Open()

        Dim delCMD As New OleDbCommand(strsql, conn)
        delDA.InsertCommand = delCMD

        'delCMD.Parameters.Add("YWBTCD", OleDbType.VarChar, 15).Value = Me.mYWBTCD
        'delCMD.Parameters.Add("YWPLNR", OleDbType.VarChar, 5).Value = Me.mYWPLNR
        'insertCMD.Parameters.Add("YWPREC", OleDbType.VarChar, 6).Value = Me.mYWPREC
        'insertCMD.Parameters.Add("YWPROC", OleDbType.VarChar, 6).Value = Me.mYWPROC
        'insertCMD.Parameters.Add("YWRDOC", OleDbType.VarChar, 2).Value = Me.mYWRDOC
        'insertCMD.Parameters.Add("YWDOCK", OleDbType.VarChar, 5).Value = Me.mYWDOCK
        'insertCMD.Parameters.Add("YWINS", OleDbType.VarChar, 2).Value = Me.mYWINS
        'insertCMD.Parameters.Add("YWPKC", OleDbType.VarChar, 5).Value = Me.mYWPKC
        'insertCMD.Parameters.Add("YWPKV", OleDbType.Decimal, 5).Value = Me.mYWPKV
        'insertCMD.Parameters.Add("YWPART", OleDbType.VarChar, 25).Value = Me.mYWPART
        'insertCMD.Parameters.Add("YWCHG", OleDbType.VarChar, 3).Value = Me.mYWCHG
        'insertCMD.Parameters.Add("YWPNM", OleDbType.VarChar, 25).Value = Me.mYWPNM
        'insertCMD.Parameters.Add("YWCALN", OleDbType.VarChar, 6).Value = Me.mYWCALN
        'insertCMD.Parameters.Add("YWPDCD", OleDbType.VarChar, 4).Value = Me.mYWPDCD
        'insertCMD.Parameters.Add("YWPDTP", OleDbType.VarChar, 3).Value = Me.mYWPDTP
        'insertCMD.Parameters.Add("YWPDGP", OleDbType.VarChar, 2).Value = Me.mYWPDGP
        'insertCMD.Parameters.Add("YWODNO", OleDbType.VarChar, 20).Value = Me.mYWODNO
        'insertCMD.Parameters.Add("YWPMON", OleDbType.Decimal, 6).Value = Me.mYWPMON
        'insertCMD.Parameters.Add("YWIVNO", OleDbType.VarChar, 15).Value = Me.mYWIVNO
        'insertCMD.Parameters.Add("YWIVDT", OleDbType.Decimal, 8).Value = Me.mYWIVDT
        'insertCMD.Parameters.Add("YWRCDT", OleDbType.Decimal, 8).Value = Me.mYWRCDT
        'insertCMD.Parameters.Add("YWEFDT", OleDbType.Decimal, 8).Value = Me.mYWEFDT
        'insertCMD.Parameters.Add("YWCFDT", OleDbType.Decimal, 8).Value = Me.mYWCFDT
        'insertCMD.Parameters.Add("YWRQTY", OleDbType.Decimal, 6).Value = Me.mYWRQTY
        'insertCMD.Parameters.Add("YWRPRC", OleDbType.Double, 15).Value = Me.mYWRPRC
        'insertCMD.Parameters.Add("YWIPRC", OleDbType.Double, 15).Value = Me.mYWIPRC
        'insertCMD.Parameters.Add("YWNPRC", OleDbType.Double, 15).Value = Me.mYWNPRC
        'insertCMD.Parameters.Add("YWCPRC", OleDbType.Double, 15).Value = Me.mYWCPRC
        'insertCMD.Parameters.Add("YWDFPR", OleDbType.Double, 15).Value = Me.mYWDFPR
        'insertCMD.Parameters.Add("YWDRPR", OleDbType.Double, 15).Value = Me.mYWDRPR
        'insertCMD.Parameters.Add("YWCRPR", OleDbType.Double, 15).Value = Me.mYWCRPR
        'insertCMD.Parameters.Add("YWPACK", OleDbType.Decimal, 5).Value = Me.mYWPACK
        'insertCMD.Parameters.Add("YWFLAG", OleDbType.VarChar, 1).Value = Me.mYWFLAG
        'insertCMD.Parameters.Add("YWDOCN", OleDbType.VarChar, 15).Value = Me.mYWDOCN
        'insertCMD.Parameters.Add("YWCRBY", OleDbType.VarChar, 10).Value = Me.mYWCRBY
        'insertCMD.Parameters.Add("YWCRTD", OleDbType.Decimal, 8).Value = Me.mYWCRTD
        'insertCMD.Parameters.Add("YWCRTT", OleDbType.Decimal, 6).Value = Me.mYWCRTT
        'insertCMD.Parameters.Add("YWCMBY", OleDbType.VarChar, 10).Value = Me.mYWCMBY
        'insertCMD.Parameters.Add("YWCMDT", OleDbType.Decimal, 8).Value = Me.mYWCMDT
        'insertCMD.Parameters.Add("YWCMTM", OleDbType.Decimal, 6).Value = Me.mYWCMTM



        Try
            delCMD.ExecuteNonQuery()
            Me.mAccStatus = True
        Catch ex As System.Exception
            Me.mAccStatus = False
            conn.Close()
        End Try


        delCMD.Dispose()
        conn.Close()


    End Sub

    Public Sub selectdata()

        Dim dt As New DataTable
        Dim strsql As String
        strsql = "select * from " & Classconn.LIBRARY_FILE
        Dim da As New OleDbDataAdapter(strsql, Classconn.strCon400)
        da.Fill(dt)

        If dt.Rows.Count < 0 Then
            Me.mAccStatus = False
        Else
            Me.mAccStatus = True
        End If



    End Sub
    








End Class
