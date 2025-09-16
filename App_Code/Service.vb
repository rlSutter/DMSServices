Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.Data.SqlClient
Imports System.IO
Imports System.Collections
Imports System.Collections.Generic
Imports System.Net
Imports System.Configuration
Imports System.Net.Mail
Imports System.Math
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports System.Threading
Imports System.Runtime.InteropServices
Imports Dms.AsynchronousOperations
Imports System.Data
Imports System.Xml.Serialization
Imports log4net
Imports Amazon
Imports Amazon.S3
Imports System.Threading.Tasks
Imports System.Security.Cryptography.X509Certificates
Imports System.Runtime.Caching


<WebService(Namespace:="http://dms.hq.local/svc/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService

    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum

    Public Class clsSSL
        Public Function AcceptAllCertifications(ByVal sender As Object, ByVal certification As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
            Return True
        End Function
    End Class

    Public Class profile
        Public UID As String
        Public USessID As String
        Public CONTACT_ID As String
        Public SUB_ID As String
        Public LANG_CD As String
        Public OWNER_EMAIL As String
        Public CONTACT_OU_ID As String
        Public READ_ONLY As String
        Public SHARED_FLG As String
        Public PRIVACY_FLG As String
        Public SITE_ONLY As String
        Public TERM_FLG As String
        Public START_FLG As String
        Public SUPER_FLG As String
        Public LOGOURL As String
        Public USERNAME As String
        Public PART_ACC_FLG As String
        Public TRAINER_ACC_FLG As String
        Public REG_AS_EMP_ID As String
        Public DOMAIN As String
        Public DOMAIN_FLG As String
        Public SYSADMIN_FLG As String
        Public TRAINING_FLG As String
        Public TRAINER_FLG As String
        Public TRAINER_ID As String
        Public PART_FLG As String
        Public PART_ID As String
        Public MT_FLG As String
        Public MT_ID As String
        Public SVC_TYPE As String
        Public CYCLE_DAYS As String
        Public JURISDICTION_ID As String
        Public REPORTS_FLG As String
        Public EMP_ID As String
        Public LOGGED_IN As String
        Public SUB_CURRCLM As String
        Public PProgram As String
        Public PHEADER As String
        Public PFOOTER As String
        Public DMS_SESSION_ID As String
        Public DMS_USER_ID As String
        Public DMS_USER_AID As String
        Public DMS_SUB_ID As String
        Public DMS_DOMAIN_ID As String
        Public HOME_URL As String
        Public DEF_SUB_ID As String
        Public UNSUB_URL As String
        Public LOGOUT_URL As String
        Public numrights As String
        Public cached As String
    End Class

    <WebMethod(Description:="Publish a document to a specified individual and optionally notify them")> _
    Public Function PublishDMSDoc(ByVal DocId As String, ByVal ContactId As String, ByVal NotifyFlg As String, _
            ByVal ReqdFlag As String, ByVal Domain As String, ByVal Expiration As String, ByVal Debug As String) As String

        ' This function creates an association for the individual specified, and optionally sends them an
        ' email notice with portal access instructions

        '   DocId   	- The "DMS.Documents.row_id" of the document (req.)
        '   ContactId   - The "siebeldb.S_CONTACT.ROW_ID" of the individual (req.)
        '   NotifyFlg   - Set to "Y" to send an automated email notice to the individual (opt)
        '   ReqdFlag    - The document is marked as "Required" reading (opt)
        '   Domain      - The Domain for the user.  (opt)
        '   Expiration  - The expiration date of the document, as a string (opt)

        ' web.config Parameters used:
        '   dms        	    - connection string to DMS.dms database

        ' Variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String
        Dim bResponse As Boolean

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' HCIDB Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim ConnS As String

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("PDDDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service
        Dim EmailService As New basic.com.certegrity.cloudsvc.Service
        Dim DmsService As New local.hq.dms.Service
        Dim sXML As String

        ' Other declarations
        Dim temp, loginpath, EOL As String
        Dim doc_title, doc_desc, doc_count As String
        Dim FST_NAME, LAST_NAME, UID, EMAIL_ADDR, ReplyTo As String
        Dim PART_ID, TRAINER_NUM, MT_ID As String
        Dim Subject, Body As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        doc_title = "Document"
        doc_desc = ""
        doc_count = "0"
        EOL = Chr(13)
        sXML = ""
        FST_NAME = ""
        LAST_NAME = ""
        EMAIL_ADDR = ""
        UID = ""
        PART_ID = ""
        TRAINER_NUM = ""
        MT_ID = ""
        ReplyTo = ""
        Subject = ""
        Body = ""
        bResponse = False
        results = "Success"

        ' ============================================
        ' Fix parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "" Then Debug = "N"
        If Debug = "T" Then
            DocId = "1889609"
            ContactId = "21120611WE0"
            NotifyFlg = "N"
            ReqdFlag = "N"
            Expiration = ""
        Else
            DocId = Trim(HttpUtility.UrlEncode(DocId))
            If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))
            If InStr(DocId, "%") > 0 Then DocId = Trim(DocId)
            ContactId = Trim(HttpUtility.UrlEncode(ContactId))
            If InStr(ContactId, "%") > 0 Then ContactId = Trim(HttpUtility.UrlDecode(ContactId))
            If InStr(ContactId, "%") > 0 Then ContactId = Trim(ContactId)
            If InStr(ContactId, " ") > 0 Then ContactId = ContactId.Replace(" ", "+")
            If NotifyFlg = "" Then NotifyFlg = "N"
            If ReqdFlag = "" Then ReqdFlag = "N"
        End If

        ' ============================================
        ' Get system defaults
        ' hcidb
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("siebeldb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("PublishDMSDoc_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
            results = "Failure"
            GoTo CloseOut2
        End Try
        ' dms
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;Min Pool Size=3;Max Pool Size=3;Connect Timeout=5;database=DMS"
        Catch ex As Exception
            errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\PublishDMSDoc.log"
            Try
                log4net.GlobalContext.Properties("PDDLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & "Error Opening Log. " & vbCrLf
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  ContactId: " & ContactId)
                mydebuglog.Debug("  NotifyFlg: " & NotifyFlg)
                mydebuglog.Debug("  Domain: " & Domain)
            End If
        End If

        ' ============================================
        ' Check required parameters
        If (DocId = "" Or ContactId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & "Invalid parameter(s) " & vbCrLf
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open database connections
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Get details of the document and verify that the specified document exists
        SqlS = "SELECT D.name, D.description " & _
        "FROM DMS.dbo.Documents D  " & _
        "WHERE D.row_id=" & DocId
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get document information: " & SqlS)
        d_cmd.CommandText = SqlS
        d_dr = d_cmd.ExecuteReader()
        If Not d_dr Is Nothing Then
            While d_dr.Read()
                Try
                    doc_title = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                    doc_desc = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType)).ToString
                    If doc_title = "" Then results = "Failure"
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error getting document information. " & ex.ToString & vbCrLf
                    GoTo CloseOut
                End Try
            End While
        Else
            errmsg = errmsg & "Error getting document information. " & vbCrLf
            results = "Failure"
        End If
        d_dr.Close()
        If Debug = "Y" Then mydebuglog.Debug("   > doc_title: " & doc_title)
        If doc_title = "" Then
            results = "Failure"
            errmsg = errmsg & "Document does not exist. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Get Contact information and verify that the specified contact exists 
        SqlS = "SELECT C.FST_NAME, C.LAST_NAME, C.EMAIL_ADDR, C.X_REGISTRATION_NUM, D.DOMAIN, D.CS_EMAIL, " & _
        "C.X_PART_ID, C.X_TRAINER_NUM, C.REG_AS_EMP_ID " & _
        "FROM siebeldb.dbo.S_CONTACT C " & _
        "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID=C.ROW_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_DOMAIN D ON D.DOMAIN=S.DOMAIN " & _
        "WHERE C.ROW_ID='" & ContactId & "'"
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Getting contact information: " & SqlS)
        cmd.CommandText = SqlS
        dr = cmd.ExecuteReader()
        If Not dr Is Nothing Then
            While dr.Read()
                Try
                    FST_NAME = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                    LAST_NAME = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                    EMAIL_ADDR = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                    UID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                    If Domain = "" Then Domain = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                    ReplyTo = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                    PART_ID = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                    TRAINER_NUM = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                    MT_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error getting contact information. " & ex.ToString & vbCrLf
                    GoTo CloseOut
                End Try
            End While
            dr.Close()
        Else
            errmsg = errmsg & "Error getting contact information. " & vbCrLf
            dr.Close()
            results = "Failure"
        End If
        If Debug = "Y" Then mydebuglog.Debug("   > name: " & FST_NAME & " " & LAST_NAME & vbCrLf & "   > email address: " & EMAIL_ADDR)
        If doc_title = "" Then
            results = "Failure"
            errmsg = errmsg & "Contact does not exist. "
            GoTo CloseOut
        End If
        If Domain = "" Then Domain = "TIPS"

        ' ============================================
        ' Remove existing association if test mode
        If Debug = "T" Then
            SqlS = "DELETE FROM DMS.dbo.Document_Associations " & _
            "WHERE doc_id=" & DocId & " AND fkey='" & ContactId & "' AND association_id=3"
            d_cmd.CommandText = SqlS
            Try
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Check to see if association exists already
        Dim numrecs As Integer = 0
        SqlS = "SELECT COUNT(*) FROM DMS.dbo.Document_Associations WHERE doc_id=" & DocId & " AND fkey='" & ContactId & "' AND association_id=3"
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Checking to see if association exists: " & SqlS)
        cmd.CommandText = SqlS
        dr = cmd.ExecuteReader()
        If Not dr Is Nothing Then
            While dr.Read()
                Try
                    numrecs = CheckDBNull(dr(0), enumObjectType.IntType)
                Catch ex As Exception
                End Try
            End While
            dr.Close()
        Else
            errmsg = errmsg & "Error getting contact information. " & vbCrLf
            dr.Close()
            results = "Failure"
        End If
        If Debug = "Y" Then mydebuglog.Debug("   > numrecs found: " & numrecs.ToString)

        ' ============================================
        ' Create DMS association record to allow contact to see document
        If numrecs = 0 Then
            temp = "null"
            If Expiration <> "" Then temp = "'" & Expiration & "'"
            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag, expiration) " &
                "SELECT TOP 1 1, 1,3," & DocId & ",'" & ContactId & "','Y','Y','" & ReqdFlag & "'," & temp & " " &
                "FROM DMS.dbo.Document_Associations " &
                "WHERE NOT EXISTS (SELECT doc_id FROM DMS.dbo.Document_Associations " &
                "WHERE doc_id=" & DocId & " AND fkey='" & ContactId & "' AND association_id=3)"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Create association: " & SqlS)
            d_cmd.CommandText = SqlS
            Try
                returnv = d_cmd.ExecuteNonQuery()
                If returnv = 0 Then
                    results = "Failure"
                End If
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & "Error creating the association record. " & ex.ToString & vbCrLf
            End Try
        End If

        ' ============================================
        ' Get current document count if the user has a subscription
        If results <> "Failure" Then
            Try
                DmsService.UpdDMSDocCountAsync(ContactId, TRAINER_NUM, PART_ID, MT_ID, Debug)
                'doc_count = DmsService.UpdDMSDocCount(ContactId, TRAINER_NUM, PART_ID, MT_ID, Debug)
                'If Debug = "Y" Then mydebuglog.Debug("  > doc_count: " & doc_count)
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & "Unable to update document count. " & ex.ToString & vbCrLf
            End Try
        End If

        ' ============================================
        ' Set service loginpath
        loginpath = ""
        If results <> "Failure" Then
            Select Case Domain
                Case "TIPS"
                    loginpath = "http://www.gettips.com/servicelogin.html?RNL="
                Case "PBSA"
                    loginpath = "http://www.gettips.com/servicelogin.html?RNL="
                Case Else
                    loginpath = "http://www.compliancetracking.com/servicelogin.html?RNL="
            End Select
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Loginpath: " & loginpath)
        End If

        ' ============================================
        ' Send Notification if applicable
        '  MESSAGE 0083
        If NotifyFlg = "Y" And results <> "Failure" Then
            If EMAIL_ADDR <> "" And ReplyTo <> "" Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Sending notice")
                If UID <> "" Then
                    Subject = "A document was published to you in your " & Domain & " training portal"
                    Body = FST_NAME & " " & LAST_NAME & "," & EOL & EOL & "The document '" & doc_title & "' is now available to you in your " & Domain & " document library located within your Certification Manager training portal.  " & EOL & EOL & "This document is described as: " & EOL & " " & Trim(doc_desc) & EOL & EOL
                    Body = Body & "To access your Certification Manager document library and this document, click on the following link:" &
                        EOL & EOL & loginpath & "mydocs&DOM=" & Domain & "&DLF=Y" & EOL & EOL
                Else
                    Subject = "A document was published to you"
                    Body = FST_NAME & " " & LAST_NAME & "," & EOL & EOL & "Attached please find the document '" & doc_title & "'.  " & EOL & EOL & "This document is described as: " & EOL & " " & Trim(doc_desc) & EOL & EOL
                End If
                Body = Body & "Please respond to this email if you have any questions. Thank you."
                sXML = "<EMailMessageList><EMailMessage>"
                sXML = sXML & "<debug>" & Debug & "</debug>"
                sXML = sXML & "<database>C</database>"
                sXML = sXML & "<Id></Id>"
                sXML = sXML & "<SourceId></SourceId>"
                sXML = sXML & "<From>" & ReplyTo & "</From>"
                sXML = sXML & "<FromId></FromId>"
                sXML = sXML & "<FromName>Customer Service</FromName>"
                sXML = sXML & "<To>" & EMAIL_ADDR & "</To>"
                sXML = sXML & "<ToId>" & ContactId & "</ToId>"
                sXML = sXML & "<Cc></Cc>"
                sXML = sXML & "<Bcc></Bcc>"
                sXML = sXML & "<ReplyTo>" & ReplyTo & "</ReplyTo>"
                sXML = sXML & "<Subject>" & Subject & "</Subject>"
                sXML = sXML & "<Body>" & HttpUtility.UrlEncode(Body) & "</Body>"
                sXML = sXML & "<Format></Format>"
                If UID = "" Then
                    sXML = sXML & "<AttachmentList>"
                    sXML = sXML & "<Attachment>" & DocId & "</Attachment>"
                    sXML = sXML & "</AttachmentList>"
                End If
                sXML = sXML & "</EMailMessage></EMailMessageList>"
                If Debug = "Y" Then mydebuglog.Debug("  > Email XML: " & sXML & vbCrLf)
                Try
                    EmailService.SendMailAsync(sXML)
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & "Unable to send notification. " & ex.ToString & vbCrLf
                End Try
            End If
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            errmsg = errmsg & CloseDBConnection(con, cmd, dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the hcidb database connection. " & vbCrLf
        End Try
        Try
            errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for Doc #" & DocId & " to contact id " & ContactId
        myeventlog.Info("PublishDMSDoc : Results: " & ltemp)
        If Trim(errmsg) <> "" Then myeventlog.Error("PublishDMSDoc : Error: " & Trim(errmsg))
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                Else
                    mydebuglog.Debug("  Results: " & ltemp)
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return results

    End Function

    <WebMethod(Description:="Updates a specified contact portal document count")> _
    Public Function AsyncUpdDMSDocCount(ByVal CONTACT_ID As String, ByVal TRAINER_NUM As String, ByVal PART_ID As String, _
         ByVal MT_ID As String, ByVal Debug As String) As String

        ' This function creates an association for the individual specified, and optionally sends them an
        ' email notice with portal access instructions

        '   CONTACT_ID	- The "siebeldb.S_CONTACT.ROW_ID" of the individual (req)
        '   TRAINER_NUM	- The "siebeldb.S_CONTACT.X_TRAINER_NUM" of the individual (opt)
        '   PART_ID	- The "siebeldb.S_CONTACT.X_PART_ID" of the individual (opt)
        '   MT_ID	- The "siebeldb.S_CONTACT.REG_AS_EMP_ID" of the individual (opt)

        ' web.config Parameters used:
        '   dms        	    - connection string to DMS.dms database

        ' Variables
        Dim results, temp As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String
        Dim bResponse As Boolean
        Dim SUB_ID, DOMAIN, USER_AID, SUB_AID, DOMAIN_AID, UID As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("AUDDCDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        bResponse = False
        results = "Success"
        SUB_ID = ""
        DOMAIN = ""
        USER_AID = ""
        SUB_AID = ""
        DOMAIN_AID = ""
        UID = ""

        ' ============================================
        ' Fix parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "" Then Debug = "N"
        If Debug = "T" Then
            CONTACT_ID = "21120611WE0"
            PART_ID = "732632"
            MT_ID = ""
            TRAINER_NUM = "22"
        Else
            CONTACT_ID = Trim(HttpUtility.UrlEncode(CONTACT_ID))
            If InStr(CONTACT_ID, "%") > 0 Then CONTACT_ID = Trim(HttpUtility.UrlDecode(CONTACT_ID))
            CONTACT_ID = Trim(EncodeParamSpaces(CONTACT_ID))
            PART_ID = Trim(HttpUtility.UrlEncode(PART_ID))
            If InStr(PART_ID, "%") > 0 Then PART_ID = Trim(HttpUtility.UrlDecode(PART_ID))
            PART_ID = Trim(EncodeParamSpaces(PART_ID))
            MT_ID = Trim(HttpUtility.UrlEncode(MT_ID))
            If InStr(MT_ID, "%") > 0 Then MT_ID = Trim(HttpUtility.UrlDecode(MT_ID))
            MT_ID = Trim(EncodeParamSpaces(MT_ID))
        End If

        ' ============================================
        ' Get Parameters
        Try
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("UpdDMSDocCount_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\AsyncUpdDMSDocCount.log"
            Try
                log4net.GlobalContext.Properties("AUDDCLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & "Error Opening Log. " & vbCrLf
                results = "Failure"
                GoTo CloseOut
            End Try
        End If

        ' ============================================
        ' Check required parameters
        If (CONTACT_ID = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & "Invalid parameter(s) " & vbCrLf
            GoTo CloseOut
        End If

        ' ============================================
        ' Update Document Count
        ' Create an instance of the test class.
        Dim ad As New AsyncMain()

        ' Create the delegate.
        Dim caller As New AsynchUpdDMSDoc(AddressOf ad.UpdDMSDoc)

        ' Initiate the asynchronous call.
        Dim result As IAsyncResult = caller.BeginInvoke(CONTACT_ID, TRAINER_NUM, PART_ID, MT_ID, Debug, Nothing, Nothing)

CloseOut:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for Contact id " & CONTACT_ID
        If Trim(errmsg) <> "" Then myeventlog.Error("AsyncUpdDMSDocCount :  Error: " & Trim(errmsg))
        myeventlog.Info("AsyncUpdDMSDocCount : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return results

    End Function

    <WebMethod(Description:="Updates a specified contact portal document count")> _
    Public Function UpdDMSDocCount(ByVal CONTACT_ID As String, ByVal TRAINER_NUM As String, ByVal PART_ID As String, _
         ByVal MT_ID As String, ByVal Debug As String) As String

        ' This function creates an association for the individual specified, and optionally sends them an
        ' email notice with portal access instructions

        '   CONTACT_ID	- The "siebeldb.S_CONTACT.ROW_ID" of the individual (req)
        '   TRAINER_NUM	- The "siebeldb.S_CONTACT.X_TRAINER_NUM" of the individual (opt)
        '   PART_ID	- The "siebeldb.S_CONTACT.X_PART_ID" of the individual (opt)
        '   MT_ID	- The "siebeldb.S_CONTACT.REG_AS_EMP_ID" of the individual (opt)

        ' web.config Parameters used:
        '   dms        	    - connection string to DMS.dms database

        ' Variables
        Dim results, temp As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String
        Dim bResponse As Boolean
        Dim doc_count As String
        Dim SUB_ID, DOMAIN, USER_AID, SUB_AID, DOMAIN_AID, UID As String
        Dim Category_Constraint, TRAINER_FLG, MT_FLG, PART_FLG, TRAINING_FLG, TRAINER_ACC_FLG, SITE_ONLY, SYSADMIN_FLG, EMP_ID As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' HCIDB Database declarations
        Dim con, RO_con As SqlConnection
        Dim cmd, RO_cmd As SqlCommand
        Dim dr, RO_dr As SqlDataReader
        Dim ConnS, RO_ConnS As String

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("UDDCDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        temp = ""
        errmsg = ""
        bResponse = False
        doc_count = "0"
        results = "Success"
        SUB_ID = ""
        DOMAIN = ""
        USER_AID = ""
        SUB_AID = ""
        DOMAIN_AID = ""
        UID = ""
        TRAINER_FLG = ""
        MT_FLG = ""
        PART_FLG = ""
        TRAINING_FLG = ""
        TRAINER_ACC_FLG = ""
        SITE_ONLY = ""
        SYSADMIN_FLG = ""
        EMP_ID = ""

        ' ============================================
        ' Fix parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "" Then Debug = "N"
        If Debug = "T" Then
            CONTACT_ID = "21120611WE0"
            PART_ID = "732632"
            MT_ID = ""
            TRAINER_NUM = "22"
        Else
            CONTACT_ID = Trim(HttpUtility.UrlEncode(CONTACT_ID))
            If InStr(CONTACT_ID, "%") > 0 Then CONTACT_ID = Trim(HttpUtility.UrlDecode(CONTACT_ID))
            CONTACT_ID = Trim(EncodeParamSpaces(CONTACT_ID))
            PART_ID = Trim(HttpUtility.UrlEncode(PART_ID))
            If InStr(PART_ID, "%") > 0 Then PART_ID = Trim(HttpUtility.UrlDecode(PART_ID))
            PART_ID = Trim(EncodeParamSpaces(PART_ID))
            MT_ID = Trim(HttpUtility.UrlEncode(MT_ID))
            If InStr(MT_ID, "%") > 0 Then MT_ID = Trim(HttpUtility.UrlDecode(MT_ID))
            MT_ID = Trim(EncodeParamSpaces(MT_ID))
        End If

        ' ============================================
        ' Get system defaults
        ' hcidb1
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("siebeldb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            RO_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidbro").ConnectionString
            If RO_ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;ApplicationIntent=ReadOnly;"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("UpdDMSDocCount_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
            results = "Failure"
            GoTo CloseOut2
        End Try
        ' dms
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;database=DMS"
        Catch ex As Exception
            errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\UpdDMSDocCount.log"
            Try
                log4net.GlobalContext.Properties("UDDCLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & "Error Opening Log. " & vbCrLf
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  CONTACT_ID: " & CONTACT_ID)
                mydebuglog.Debug("  PART_ID: " & PART_ID)
                mydebuglog.Debug("  MT_ID: " & MT_ID)
                mydebuglog.Debug("  TRAINER_NUM: " & TRAINER_NUM)
            End If
        End If

        ' ============================================
        ' Check required parameters
        If (CONTACT_ID = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & "Invalid parameter(s) " & vbCrLf
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open database connections
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If
        errmsg = OpenDBConnection(RO_ConnS, RO_con, RO_cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Get Subscription and Domain Info
        SqlS = "SELECT S.ROW_ID, S.DOMAIN, C.X_REGISTRATION_NUM, C.X_TRAINER_FLG, C.X_MAST_TRNR_FLG, " &
            "(SELECT CASE WHEN C.X_PART_ID IS NOT NULL AND C.X_PART_ID<>'' THEN 'Y' ELSE 'N' END) AS PART_FLG, S.SVC_TYPE, " &
            "SC.TRAINER_ACC_FLG, SC.SITE_ONLY_FLG, SC.SYSADMIN_FLG, E.ROW_ID " &
            "FROM siebeldb.dbo.CX_SUB_CON SC " &
            "INNER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID " &
            "INNER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=SC.CON_ID " &
            "LEFT OUTER JOIN siebeldb.dbo.S_EMPLOYEE E ON E.X_CON_ID=C.ROW_ID AND E.CNTRCTR_EMPLR_ID IS NULL " &
            "WHERE SC.CON_ID='" & CONTACT_ID & "'"
        If Debug = "Y" Then mydebuglog.Debug("  Get subscription info: " & SqlS)
        Try
            RO_cmd.CommandText = SqlS
            RO_dr = RO_cmd.ExecuteReader()
            If Not RO_dr Is Nothing Then
                While RO_dr.Read()
                    Try
                        SUB_ID = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType)).ToString
                        DOMAIN = Trim(CheckDBNull(RO_dr(1), enumObjectType.StrType)).ToString
                        UID = Trim(CheckDBNull(RO_dr(2), enumObjectType.StrType)).ToString
                        TRAINER_FLG = Trim(CheckDBNull(RO_dr(3), enumObjectType.StrType)).ToString
                        MT_FLG = Trim(CheckDBNull(RO_dr(4), enumObjectType.StrType)).ToString
                        PART_FLG = Trim(CheckDBNull(RO_dr(5), enumObjectType.StrType)).ToString
                        TRAINING_FLG = Trim(CheckDBNull(RO_dr(6), enumObjectType.StrType)).ToString
                        Select Case TRAINING_FLG.ToUpper()
                            Case "CERTIFICATION MANAGER REG DB"
                                TRAINING_FLG = "N"
                            Case "CERTIFICATION MANAGER REPORTS"
                                TRAINING_FLG = "N"
                            Case Else
                                TRAINING_FLG = "Y"
                        End Select
                        TRAINER_ACC_FLG = Trim(CheckDBNull(RO_dr(7), enumObjectType.StrType)).ToString
                        SITE_ONLY = Trim(CheckDBNull(RO_dr(8), enumObjectType.StrType)).ToString
                        SYSADMIN_FLG = Trim(CheckDBNull(RO_dr(9), enumObjectType.StrType)).ToString
                        EMP_ID = Trim(CheckDBNull(RO_dr(10), enumObjectType.StrType)).ToString
                    Catch ex As Exception
                        'results = "Failure"
                        'errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
                        GoTo CloseOut
                    End Try
                End While
            Else
                errmsg = errmsg & "Error getting document count. " & vbCrLf
                results = "Failure"
            End If
            Try
                RO_dr.Close()
                RO_dr = Nothing
            Catch ex As Exception
            End Try
        Catch ex As Exception
        End Try
        If Debug = "Y" Then
            mydebuglog.Debug("      > Sub_Id/Domain/UID: " & SUB_ID & "/" & DOMAIN & "/" & UID)
            mydebuglog.Debug("      > TRAINER_FLG:" & TRAINER_FLG)
            mydebuglog.Debug("      > MT_FLG: " & MT_FLG)
            mydebuglog.Debug("      > PART_FLG: " & PART_FLG)
            mydebuglog.Debug("      > TRAINING_FLG: " & TRAINING_FLG)
            mydebuglog.Debug("      > TRAINER_ACC_FLG: " & TRAINER_ACC_FLG)
            mydebuglog.Debug("      > SITE_ONLY: " & SITE_ONLY)
            mydebuglog.Debug("      > SYSADMIN_FLG: " & SYSADMIN_FLG)
            mydebuglog.Debug("      > EMP_ID: " & EMP_ID)
        End If

        ' If no subscription, no point
        If SUB_ID = "" Or UID = "" Then
            'errmsg = errmsg & "No UID " & UID & " or SUB_ID " & SUB_ID & " to update"
            GoTo CloseOut
        End If

        ' ============================================
        ' Get DMS security information
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get DMS security information")

        ' -----
        ' User AID
        SqlS = "SELECT UA.row_id " & _
        "FROM DMS.dbo.User_Group_Access UA " & _
        "INNER JOIN DMS.dbo.Users U ON U.row_id=UA.access_id " & _
        "WHERE UA.type_id='U' AND U.ext_user_id='" & CONTACT_ID & "'"
        If Debug = "Y" Then mydebuglog.Debug("  .. Get user security: " & SqlS)
        Try
            RO_cmd.CommandText = SqlS
            RO_dr = RO_cmd.ExecuteReader()
            If Not RO_dr Is Nothing Then
                While RO_dr.Read()
                    Try
                        USER_AID = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType)).ToString
                    Catch ex As Exception
                    End Try
                End While
            End If
            Try
                RO_dr.Close()
                RO_dr = Nothing
            Catch ex As Exception
            End Try
        Catch ex As Exception
        End Try
        If Debug = "Y" Then mydebuglog.Debug("      > USER_AID: " & USER_AID)

        ' -----
        ' Subscription AID
        SqlS = "SELECT UA.row_id " & _
        "FROM DMS.dbo.User_Group_Access UA " & _
        "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " & _
        "WHERE UA.type_id='G' AND G.name='" & SUB_ID & "'"
        If Debug = "Y" Then mydebuglog.Debug("  .. Get subscription security: " & SqlS)
        Try
            RO_cmd.CommandText = SqlS
            RO_dr = RO_cmd.ExecuteReader()
            If Not RO_dr Is Nothing Then
                While RO_dr.Read()
                    Try
                        SUB_AID = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType)).ToString
                    Catch ex As Exception
                    End Try
                End While
            End If
            Try
                RO_dr.Close()
                RO_dr = Nothing
            Catch ex As Exception
            End Try
        Catch ex As Exception
        End Try
        If Debug = "Y" Then mydebuglog.Debug("      > SUB_AID: " & SUB_AID)

        ' -----
        ' Domain AID
        SqlS = "SELECT UA.row_id " & _
        "FROM DMS.dbo.User_Group_Access UA " & _
        "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " & _
        "WHERE UA.type_id='G' AND G.name='" & DOMAIN & "'"
        If Debug = "Y" Then mydebuglog.Debug("  .. Get domain security: " & SqlS)
        Try
            RO_cmd.CommandText = SqlS
            RO_dr = RO_cmd.ExecuteReader()
            If Not RO_dr Is Nothing Then
                While RO_dr.Read()
                    Try
                        DOMAIN_AID = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType)).ToString
                    Catch ex As Exception
                    End Try
                End While
            End If
            Try
                RO_dr.Close()
                RO_dr = Nothing
            Catch ex As Exception
            End Try
        Catch ex As Exception
        End Try
        If Debug = "Y" Then mydebuglog.Debug("      > DOMAIN_AID: " & DOMAIN_AID)

        ' ============================================
        ' Generate Category Constraint
        Category_Constraint = "CK.key_id IN ("
        If TRAINER_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "3,"
        End If
        If MT_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "5,"
        End If
        If PART_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "7,"
        End If
        If TRAINING_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "8,"
        End If
        If TRAINER_ACC_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "10,"
        End If
        If SITE_ONLY = "Y" Then
            Category_Constraint = Category_Constraint & "12,"
        End If
        Category_Constraint = Category_Constraint & "13,"
        If SYSADMIN_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "15,"
        End If
        If EMP_ID <> "" Then
            Category_Constraint = Category_Constraint & "16,"
        End If
        Category_Constraint = Category_Constraint & "14) "
        If Debug = "Y" Then mydebuglog.Debug("  Category_Constraint: " & Category_Constraint)

        ' ============================================
        ' Get current document count if the user has a subscription
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Generate document count")

        SqlS = "SELECT count(1) AS NUM_DOC " & _
        "FROM (" & _
        "SELECT D.row_id " & _
        "FROM DMS.dbo.Documents D " & _
        "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " & _
        "LEFT OUTER JOIN DMS.dbo.Category_Keywords CK ON CK.cat_id=DC.cat_id " & _
        "WHERE DC.pr_flag='Y' AND (CK.key_id IN (3,5,7,8,13,15,16,14)) " & _
        "GROUP BY D.row_id " & _
        "INTERSECT " & _
        "SELECT DISTINCT DA.doc_id " & _
        "FROM DMS.dbo.Document_Associations DA " & _
        "INNER JOIN DMS.dbo.Documents D on D.row_id=DA.doc_id " & _
        "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id "
        SqlS = SqlS & "WHERE ((DA.association_id='3' AND DA.fkey='" & CONTACT_ID & "' AND DA.pr_flag='Y') or "
        If TRAINER_NUM <> "" Then SqlS = SqlS & "(DA.association_id='5' AND DA.fkey='" & TRAINER_NUM & "' AND DA.pr_flag='Y') or "
        If PART_ID <> "" Then SqlS = SqlS & "(DA.association_id='4' AND DA.fkey='" & PART_ID & "' AND DA.pr_flag='Y') or "
        If MT_ID <> "" Then SqlS = SqlS & "(DA.association_id='37' AND DA.fkey='" & MT_ID & "' AND DA.pr_flag='Y') or "
        SqlS = Left(SqlS, Len(SqlS) - 4) & ") AND D.deleted IS NULL AND ("
        If USER_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & USER_AID & " OR "
        If SUB_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & SUB_AID & " OR "
        If DOMAIN_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & DOMAIN_AID & " OR "
        SqlS = Left(SqlS, Len(SqlS) - 4) & ") GROUP BY DA.doc_id ) d "
        If Debug = "Y" Then mydebuglog.Debug("  .. Get document count: " & SqlS)
        Try
            RO_cmd.CommandText = SqlS
            RO_dr = RO_cmd.ExecuteReader()
            If Not RO_dr Is Nothing Then
                While RO_dr.Read()
                    Try
                        doc_count = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType)).ToString
                    Catch ex As Exception
                        results = "Failure"
                        errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
                        GoTo CloseOut
                    End Try
                End While
            Else
                errmsg = errmsg & "Error getting document count. " & vbCrLf
                results = "Failure"
            End If
            RO_dr.Close()
            RO_dr = Nothing
        Catch ex As Exception
        End Try
        If Debug = "Y" Then mydebuglog.Debug("      > doc_count: " & doc_count)

        ' -----
        ' Update CX_SUB_CON.NEW_DOC with document count if applicable
        If doc_count <> "" Then
            SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
            "SET NEW_DOC=" & doc_count & _
            " WHERE CON_ID='" & CONTACT_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  .. Update contact document count in CX_SUB_CON: " & SqlS)
            Try
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                If returnv = 0 Then results = "Failure"
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & "Error setting the document count. " & ex.ToString & vbCrLf
            End Try

            SqlS = "UPDATE siebeldb.dbo.S_CONTACT " & _
                "SET DCKING_NUM=" & doc_count & " " & _
                "WHERE ROW_ID='" & CONTACT_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  .. Update contact document count in S_CONTACT: " & SqlS)
            Try
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                If returnv = 0 Then results = "Failure"
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & "Error setting the document count. " & ex.ToString & vbCrLf
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            errmsg = errmsg & CloseDBConnection(con, cmd, dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the hcidb database connection. " & vbCrLf
        End Try
        Try
            errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
        End Try
        Try
            errmsg = errmsg & CloseDBConnection(RO_con, RO_cmd, RO_dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " : Contact id " & CONTACT_ID & " has " & doc_count & " documents"
        If Trim(errmsg) <> "" Then myeventlog.Error("UpdDMSDocCount :  Error: " & Trim(errmsg))
        myeventlog.Info("UpdDMSDocCount : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                Else
                    mydebuglog.Debug("  Results: " & ltemp)
                End If
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Log Performance Data
        If Debug <> "T" Then
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
            If results = "Success" Then results = doc_count
        End If

        ' ============================================
        ' Return results        
        Return results

    End Function

    <WebMethod(Description:="Save or update the provided DMS document association")> _
    Public Function SaveDMSDocAssoc(ByVal DocId As String, ByVal Association As String, _
        ByVal AssocKey As String, ByVal PrFlag As String, ByVal ReqdFlag As String, ByVal Rights As String, _
        ByVal Debug As String) As Boolean

        ' This function creates a Document_Associations record for the document and association specified

        ' The input parameters are as follows:
        '
        '   DocId   	- The "DMS.Documents.row_id" of the document (req.)
        '   Association	- The "DMS.Associations.name" of the item to be stored. (req.)
        '   AssocKey    - The "DMS.Document_Associations.fkey" of the record to be created (req.)
        '   PrFlag      - The "DMS.Document_Associations.pr_flag" of the record to be created (req.)
        '   ReqdFlag    - The "DMS.Document_Associations.reqd_flag" of the record to be created (req.)
        '   Rights      - The "DMS.Document_Associations.access_type" of the record to be created (req.)
        '                   Currently, this translates into the "access_flag" setting
        '   Debug	    - The debug mode flag: "Y", "N" or "T" 
        '
        ' The results are as follows:
        '
        '   DocAssocId    - The "DMS.Document_Association.row_id" of the record created

        ' web.config Parameters used:
        '   dms        	    - connection string to DMS.dms database

        ' Variables
        Dim temp As String
        Dim results As Boolean
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SDDADebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' Local Cache declarations
        Dim DMSCache As New CachingWrapper.LocalCache

        ' Association declarations
        Dim AssocId, DocAssocId, AssocAccess As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = False
        SqlS = ""
        returnv = 0
        AssocId = ""
        DocAssocId = ""
        AssocAccess = ""
        temp = ""

        ' ============================================
        ' Get parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            DocId = "1"
            Association = "Account"
            Rights = "Y"
            AssocKey = "KEY"
            PrFlag = "N"
            ReqdFlag = "N"
        Else
            DocId = Trim(HttpUtility.UrlEncode(DocId))
            If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))

            Association = Trim(HttpUtility.UrlDecode(Association))
            If InStr(Association, "%") > 0 Then Association = Trim(HttpUtility.UrlDecode(Association))

            If InStr(AssocKey, "%") > 0 Then AssocKey = Trim(HttpUtility.UrlDecode(AssocKey))
            If InStr(AssocKey, " ") > 0 Then AssocKey = EncodeParamSpaces(AssocKey)

            Rights = Trim(HttpUtility.UrlDecode(Rights))
            If InStr(Rights, "%") > 0 Then Rights = Trim(HttpUtility.UrlDecode(Rights))

            PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
            If InStr(PrFlag, "%") > 0 Then PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))

            ReqdFlag = Trim(HttpUtility.UrlDecode(ReqdFlag))
            If InStr(ReqdFlag, "%") > 0 Then ReqdFlag = Trim(HttpUtility.UrlDecode(ReqdFlag))
        End If

        ' ============================================
        ' Get system defaults
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDocAssoc_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\SaveDMSDocAssoc.log"
            Try
                log4net.GlobalContext.Properties("SDDALogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  Association: " & Association)
                mydebuglog.Debug("  AssocKey: " & AssocKey)
                mydebuglog.Debug("  Rights: " & Rights)
                mydebuglog.Debug("  PrFlag: " & PrFlag)
                mydebuglog.Debug("  ReqdFlag: " & ReqdFlag)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(Association) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No association specified. "
            GoTo CloseOut2
        End If
        If Trim(DocId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No document specified. "
            GoTo CloseOut2
        End If
        If Trim(AssocKey) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No association key specified. "
            GoTo CloseOut2
        End If
        If Trim(Rights) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No association rights specified. "
            GoTo CloseOut2
        End If
        If Trim(PrFlag) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No association primary flag specified. "
            GoTo CloseOut2
        End If
        If Trim(ReqdFlag) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No association required flag specified. "
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open SQL Server database connections
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not connect to the database. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Load Associations from and/or into cache
        Dim dt As DataTable = New DataTable
        If Not DMSCache.GetCachedItem("Associations") Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Associations found in cache")
            Try
                dt = DMSCache.GetCachedItem("Associations")
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                GoTo CloseOut
            End Try
        Else
            SqlS = "SELECT name, row_id FROM DMS.dbo.Associations ORDER BY name"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Associations into cache: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If d_dr.HasRows Then
                    dt.Load(d_dr)
                    DMSCache.AddToCache("Associations", dt, CachingWrapper.CachePriority.NotRemovable)
                End If
                d_dr.Close()
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                GoTo CloseOut
            End Try
        End If
        If dt Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not retrieve associations. "
            GoTo CloseOut
        End If

        ' Debug output
        If Debug = "Y" Then
            mydebuglog.Debug(" Associations Columns found: " & dt.Columns.Count.ToString)
            mydebuglog.Debug(" Associations Rows found: " & dt.Rows.Count.ToString)
        End If

        ' ============================================
        ' Locate Association Id in datatable
        Dim dr() As DataRow = dt.Select("name='" & Association & "'")
        If dr Is Nothing Or dr.Length = 0 Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Association not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not find association. "
            GoTo CloseOut
        End If
        AssocId = dr(0).Item("row_id").ToString
        If Debug = "Y" Then
            mydebuglog.Debug(" Association row_id: " & AssocId)
        End If

        ' ============================================
        ' Close data objects
        Try
            dr = Nothing
            dt = Nothing
            DMSCache = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Concert rights to access_flag as needed
        If Rights = "Y" Or Rights = "N" Then AssocAccess = Rights

        ' ============================================
        ' Write Document Association record
        If AssocAccess <> "" Then
            SqlS = "INSERT INTO DMS.dbo.Document_Associations " & _
                "(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                "VALUES (1, 1, " & AssocId & ", " & DocId & ", '" & AssocKey & "', '" & PrFlag & "', '" & AssocAccess & "', '" & ReqdFlag & "')"
        Else
            SqlS = "INSERT INTO DMS.dbo.Document_Associations " & _
                "(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_type, reqd_flag) " & _
                "VALUES (1, 1, " & AssocId & ", " & DocId & ", '" & AssocKey & "', '" & PrFlag & "', '" & Rights & "', '" & ReqdFlag & "')"
        End If
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Contact: " & vbCrLf & SqlS)
        If Debug <> "T" Then
            d_cmd.CommandText = SqlS
            Try
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try
        End If
        results = True

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for association " & Association & ", with key '" & AssocKey & "' and document " & DocId
        If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDocAssoc :  Error: " & Trim(errmsg))
        myeventlog.Info("SaveDMSDocAssoc : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return results
    End Function

    <WebMethod(Description:="Save or update the provided DMS document category")> _
    Public Function SaveDMSDocCat(ByVal DocId As String, ByVal Category As String, ByVal PrFlag As String, _
        ByVal Debug As String) As Boolean

        ' This function creates a Document_Categories record for the document and category specified

        ' The input parameters are as follows:
        '
        '   DocId   	- The "DMS.Documents.row_id" of the document 
        '   Category	- The "DMS.Categories.name" of the item to be stored. (req.)
        '   PrFlag      - The primary category flag for the record to be created
        '   Debug	    - The debug mode flag: "Y", "N" or "T" 
        '
        ' The results are as follows:
        '
        '   Boolean     - True/False

        ' web.config Parameters used:
        '   dms        	    - connection string to DMS.dms database

        ' Variables
        Dim temp As String
        Dim results As Boolean
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SDDCDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' Local Cache declarations
        Dim DMSCache As New CachingWrapper.LocalCache

        ' Category declarations
        Dim CatId, DocCatId As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = False
        SqlS = ""
        returnv = 0
        CatId = ""
        DocCatId = ""
        temp = ""

        ' ============================================
        ' Get parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            DocId = "1"
            Category = "Account"
            PrFlag = "N"
        Else
            DocId = Trim(HttpUtility.UrlEncode(DocId))
            If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))

            Category = Trim(HttpUtility.UrlDecode(Category))
            If InStr(Category, "%") > 0 Then Category = Trim(HttpUtility.UrlDecode(Category))

            PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
            If InStr(PrFlag, "%") > 0 Then PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
        End If

        ' ============================================
        ' Get system defaults
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDocCat_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\SaveDMSDocCat.log"
            Try
                log4net.GlobalContext.Properties("SDDCLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  Category: " & Category)
                mydebuglog.Debug("  PrFlag: " & PrFlag)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(Category) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No category specified. "
            GoTo CloseOut2
        End If
        If Trim(DocId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No document specified. "
            GoTo CloseOut2
        End If
        If PrFlag = "" Then PrFlag = "N"

        ' ============================================
        ' Open SQL Server database connections
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not connect to the database. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Load Categories from and/or into cache
        Dim dt As DataTable = New DataTable
        If Not DMSCache.GetCachedItem("Categories") Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Categories found in cache")
            Try
                dt = DMSCache.GetCachedItem("Categories")
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                GoTo CloseOut
            End Try
        Else
            SqlS = "SELECT name, row_id FROM DMS.dbo.Categories ORDER BY name"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Categories into cache: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If d_dr.HasRows Then
                    dt.Load(d_dr)
                    DMSCache.AddToCache("Categories", dt, CachingWrapper.CachePriority.NotRemovable)
                End If
                d_dr.Close()
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                GoTo CloseOut
            End Try
        End If
        If dt Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not retrieve categories. "
            GoTo CloseOut
        End If

        ' Debug output
        If Debug = "Y" Then
            mydebuglog.Debug(" Categories Columns found: " & dt.Columns.Count.ToString)
            mydebuglog.Debug(" Categories Rows found: " & dt.Rows.Count.ToString)
        End If

        ' ============================================
        ' Locate Category Id in datatable
        Dim dr() As DataRow = dt.Select("name='" & Category & "'")
        If dr Is Nothing Or dr.Length = 0 Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Category not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not find category. "
            GoTo CloseOut
        End If
        CatId = dr(0).Item("row_id").ToString
        If Debug = "Y" Then
            mydebuglog.Debug(" Category row_id: " & CatId)
        End If

        ' ============================================
        ' Close data objects
        Try
            dr = Nothing
            dt = Nothing
            DMSCache = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Write Document_Categories record
        SqlS = "INSERT INTO DMS.dbo.Document_Categories " & _
                "(created_by, last_upd_by, doc_id, cat_id, pr_flag) " & _
                "VALUES (1, 1, " & DocId & ", " & CatId & ", '" & PrFlag & "')"
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Categories record for document: " & vbCrLf & SqlS)
        If Debug <> "T" Then
            d_cmd.CommandText = SqlS
            Try
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try
        End If
        results = True

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for category " & Category & " and document " & DocId
        If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDocCat :  Error: " & Trim(errmsg))
        myeventlog.Info("SaveDMSDocCat : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return results
    End Function

    <WebMethod(Description:="Save or update the provided DMS document keyword")> _
    Public Function SaveDMSDocKey(ByVal DocId As String, ByVal DocKey As String, ByVal KeyVal As String, _
        ByVal PrFlag As String, ByVal Debug As String) As Boolean

        ' This function creates a Document_Keywords record for the document and Keyword specified

        ' The input parameters are as follows:
        '
        '   DocId   	- The "DMS.Documents.row_id" of the document (req.)
        '   DocKey	    - The "DMS.Keywords.name" of the item to be stored. (req.)
        '   KeyVal	    - The "DMS.Document_Keywords.val" of the keyword to be created (opt.)
        '   PrFlag      - The primary DocKey flag for the record to be created
        '   Debug	    - The debug mode flag: "Y", "N" or "T" 
        '
        ' The results are as follows:
        '
        '   Boolean     - True/False

        ' web.config Parameters used:
        '   dms        	    - connection string to DMS.dms database

        ' Variables
        Dim temp As String
        Dim results As Boolean
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SDDKDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' Local Cache declarations
        Dim DMSCache As New CachingWrapper.LocalCache

        ' DocKey declarations
        Dim KeyId, DocKeyId As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = False
        SqlS = ""
        returnv = 0
        KeyId = ""
        DocKeyId = ""
        temp = ""

        ' ============================================
        ' Get parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            DocId = "1"
            DocKey = "Shared"
            KeyVal = ""
            PrFlag = "N"
        Else
            DocId = Trim(HttpUtility.UrlEncode(DocId))
            If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))

            DocKey = Trim(HttpUtility.UrlDecode(DocKey))
            If InStr(DocKey, "%") > 0 Then DocKey = Trim(HttpUtility.UrlDecode(DocKey))

            KeyVal = Trim(HttpUtility.UrlDecode(KeyVal))
            If InStr(KeyVal, "%") > 0 Then KeyVal = Trim(HttpUtility.UrlDecode(KeyVal))
            If InStr(KeyVal, " ") > 0 Then KeyVal = Replace(KeyVal, " ", "+")

            PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
            If InStr(PrFlag, "%") > 0 Then PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
        End If

        ' ============================================
        ' Get system defaults
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDocKey_debug")
            If temp = "Y" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\SaveDMSDocKey.log"
            Try
                log4net.GlobalContext.Properties("SDDKLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  DocKey: " & DocKey)
                mydebuglog.Debug("  KeyVal: " & KeyVal)
                mydebuglog.Debug("  PrFlag: " & PrFlag)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(DocKey) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No keyword specified. "
            GoTo CloseOut2
        End If
        If Trim(DocId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No document specified. "
            GoTo CloseOut2
        End If
        If PrFlag = "" Then PrFlag = "N"

        ' ============================================
        ' Open SQL Server database connections
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not connect to the database. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Load Keywords from and/or into cache
        Dim dt As DataTable = New DataTable
        If Not DMSCache.GetCachedItem("Keywords") Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Keywords found in cache")
            Try
                dt = DMSCache.GetCachedItem("Keywords")
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                GoTo CloseOut
            End Try
        Else
            SqlS = "SELECT name, row_id FROM DMS.dbo.Keywords ORDER BY name"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Keywords into cache: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If d_dr.HasRows Then
                    dt.Load(d_dr)
                    DMSCache.AddToCache("Keywords", dt, CachingWrapper.CachePriority.NotRemovable)
                End If
                d_dr.Close()
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                GoTo CloseOut
            End Try
        End If
        If dt Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not retrieve Keywords. "
            GoTo CloseOut
        End If

        ' Debug output
        If Debug = "Y" Then
            mydebuglog.Debug(" Keywords Columns found: " & dt.Columns.Count.ToString)
            mydebuglog.Debug(" Keywords Rows found: " & dt.Rows.Count.ToString)
        End If

        ' ============================================
        ' Locate DocKey Id in datatable
        Dim dr() As DataRow = dt.Select("name='" & DocKey & "'")
        If dr Is Nothing Or dr.Length = 0 Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Keyword not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not find Keyword. "
            GoTo CloseOut
        End If
        KeyId = dr(0).Item("row_id").ToString
        If Debug = "Y" Then
            mydebuglog.Debug(" Keyword row_id: " & KeyId)
        End If

        ' ============================================
        ' Close data objects
        Try
            dr = Nothing
            dt = Nothing
            DMSCache = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Write Document_Keywords record
        SqlS = "INSERT INTO DMS.dbo.Document_Keywords " & _
                "(created_by, last_upd_by, doc_id, key_id, pr_flag, val) " & _
                "VALUES (1, 1, " & DocId & ", " & KeyId & ", '" & PrFlag & "','" & KeyVal & "')"
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Keywords record for document: " & vbCrLf & SqlS)
        If Debug <> "T" Then
            d_cmd.CommandText = SqlS
            Try
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try
        End If
        results = True

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for Keyword '" & DocKey & "' with value '" & KeyVal & "' and document " & DocId
        If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDocKey :  Error: " & Trim(errmsg))
        myeventlog.Info("SaveDMSDocKey : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return results
    End Function

    <WebMethod(Description:="Create DMS document user record(s)")> _
    Public Function SaveDMSDocUser(ByVal DocId As String, ByVal Domain As String, _
        ByVal DomainOwner As String, ByVal DomainRights As String, ByVal SubId As String, _
        ByVal SubOwner As String, ByVal SubRights As String, ByVal ConId As String, ByVal ConOwner As String, _
        ByVal ConRights As String, ByVal RegId As String, ByVal RegOwner As String, ByVal RegRights As String, _
        ByVal Debug As String) As Boolean

        ' This function creates Document_Users record(s) for the document and types of users specified

        ' The input parameters are as follows:
        '
        '   DocId		    - The "DMS.Documents.row_id" of the document (req.)
        '   DocId		    - A required document id (*DMS.Documents.row_id*)
        '   Domain		    - The Domain (hcidb1.CX_DOMAIN.DOMAIN) associated with a document - optional
        '   DomainOwner		- A flag to indicate the domain is the owner of the document - optional
        '   DomainRights	- The CRUD rights of the domain to the document - optional
        '   SubId		    - The Subscription Id (hcidb1.CX_SUBSCRIPTIONS.ROW_ID) associated with a document - optional
        '   SubOwner		- A flag to indicate the subscription is the owner of the document - optional
        '   SubRights		- The CRUD rights of the subscription to the document - optional
        '   ConId		    - The Contact Id (hcidb1.S_CONTACT.ROW_ID) associated with a document - optional
        '   ConOwner		- A flag to indicate the contact is the owner of the document - optional
        '   ConRights		- The CRUD rights of the contact to the document - optional
        '   RegId		    - The Contact registration (hcidb1.S_CONTACT.X_REGISTRATION_NUM) associated with a document - optional
        '   RegOwner		- A flag to indicate the registration is the owner of the document - optional
        '   RegRights		- The CRUD rights of the registration to the document - optional
        '   Debug		    - The debug mode flag: "Y", "N" or "T" 
        '
        ' The results are as follows:
        '
        '   Boolean    		- True/False to indicate success of the operation

        ' web.config Parameters used:
        '   dms			- connection string to DMS.dms database

        ' Variables
        Dim temp As String
        Dim results As Boolean
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SDDUDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' Local Cache declarations
        Dim DMSCache As New CachingWrapper.LocalCache

        ' User declarations
        Dim SubUGA, DomainUGA, ConUGA, RegUGA As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = False
        SqlS = ""
        returnv = 0
        SubUGA = ""     ' Subscription User Group Access Id
        DomainUGA = ""  ' Domain User Group Access Id
        ConUGA = ""     ' Contact User Group Access Id
        RegUGA = ""     ' Web Registration User Group Access Id
        temp = ""

        ' ============================================
        ' Get parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            DocId = "1"
            Domain = "TIPS"
            DomainOwner = ""
            DomainRights = "R"
            SubId = ""
            SubOwner = ""
            SubRights = ""
            ConId = ""
            ConOwner = ""
            ConRights = ""
            RegId = ""
            RegOwner = ""
            RegRights = ""
        Else
            DocId = Trim(HttpUtility.UrlDecode(DocId.Trim))
            If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))

            Domain = Trim(HttpUtility.UrlDecode(Domain.Trim)).ToUpper
            If InStr(Domain, "%") > 0 Then Domain = Trim(HttpUtility.UrlDecode(Domain))

            DomainOwner = Trim(HttpUtility.UrlDecode(DomainOwner.Trim)).ToUpper
            If InStr(DomainOwner, "%") > 0 Then DomainOwner = Trim(HttpUtility.UrlDecode(DomainOwner))
            If DomainOwner = "" Then DomainOwner = "N"

            DomainRights = Trim(HttpUtility.UrlDecode(DomainRights.Trim)).ToUpper
            If InStr(DomainRights, "%") > 0 Then DomainRights = Trim(HttpUtility.UrlDecode(DomainRights))

            SubId = Trim(HttpUtility.UrlDecode(SubId.Trim)).ToUpper
            If InStr(SubId, "%") > 0 Then SubId = Trim(HttpUtility.UrlDecode(SubId))

            SubOwner = Trim(HttpUtility.UrlDecode(SubOwner.Trim)).ToUpper
            If InStr(SubOwner, "%") > 0 Then SubOwner = Trim(HttpUtility.UrlDecode(SubOwner))
            If SubOwner = "" Then SubOwner = "N"

            SubRights = Trim(HttpUtility.UrlDecode(SubRights.Trim)).ToUpper
            If InStr(SubRights, "%") > 0 Then SubRights = Trim(HttpUtility.UrlDecode(SubRights))

            ConId = Trim(HttpUtility.UrlDecode(ConId.Trim)).ToUpper
            If InStr(ConId, "%") > 0 Then ConId = Trim(HttpUtility.UrlDecode(ConId))
            If InStr(ConId, " ") > 0 Then ConId = ConId.Replace(" ", "+")

            ConOwner = Trim(HttpUtility.UrlDecode(ConOwner.Trim)).ToUpper
            If InStr(ConOwner, "%") > 0 Then ConOwner = Trim(HttpUtility.UrlDecode(ConOwner))
            If ConOwner = "" Then ConOwner = "N"

            ConRights = Trim(HttpUtility.UrlDecode(ConRights.Trim)).ToUpper
            If InStr(ConRights, "%") > 0 Then ConRights = Trim(HttpUtility.UrlDecode(ConRights))

            RegId = Trim(HttpUtility.UrlDecode(RegId.Trim)).ToUpper
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))

            RegOwner = Trim(HttpUtility.UrlDecode(RegOwner.Trim)).ToUpper
            If InStr(RegOwner, "%") > 0 Then RegOwner = Trim(HttpUtility.UrlDecode(RegOwner))
            If RegOwner = "" Then RegOwner = "N"

            RegRights = Trim(HttpUtility.UrlDecode(RegRights.Trim)).ToUpper
            If InStr(RegRights, "%") > 0 Then RegRights = Trim(HttpUtility.UrlDecode(RegRights))
        End If

        ' ============================================
        ' Get system defaults
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDocUser_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = False
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\SaveDMSDocUser.log"
            Try
                log4net.GlobalContext.Properties("SDDULogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = False
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  Domain:" & Domain)
                mydebuglog.Debug("  DomainOwner:" & DomainOwner)
                mydebuglog.Debug("  DomainRights:" & DomainRights)
                mydebuglog.Debug("  SubId:" & SubId)
                mydebuglog.Debug("  SubOwner:" & SubOwner)
                mydebuglog.Debug("  SubRights:" & SubRights)
                mydebuglog.Debug("  ConId:" & ConId)
                mydebuglog.Debug("  ConOwner:" & ConOwner)
                mydebuglog.Debug("  ConRights:" & ConRights)
                mydebuglog.Debug("  RegId:" & RegId)
                mydebuglog.Debug("  RegOwner:" & RegOwner)
                mydebuglog.Debug("  RegRights:" & RegRights)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(DocId) = "" Then
            results = False
            errmsg = errmsg & vbCrLf & "No document specified. "
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open SQL Server database connections
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Then
            results = False
            errmsg = errmsg & vbCrLf & "Could not connect to the database. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Load Subscription Users from and/or into cache
        Dim dt1 As DataTable = New DataTable
        Dim dr1() As DataRow
        If SubId <> "" Then
            If Not DMSCache.GetCachedItem("Subscriptions") Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Subscriptions found in cache")
                Try
                    dt1 = DMSCache.GetCachedItem("Subscriptions")
                Catch ex As Exception
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                    GoTo CloseOut
                End Try
            Else
                SqlS = "SELECT DISTINCT G.name, A.row_id " & _
                "FROM DMS.dbo.Groups G " & _
                "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=G.row_id " & _
                "WHERE G.type_cd='Subscription' AND A.type_id='G'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Subscriptions into cache: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If d_dr.HasRows Then
                        dt1.Load(d_dr)
                        DMSCache.AddToCache("Subscriptions", dt1, CachingWrapper.CachePriority.Default)
                    End If
                    d_dr.Close()
                Catch ex As Exception
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                    GoTo CloseOut
                End Try
            End If
            If dt1 Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                results = False
                errmsg = errmsg & vbCrLf & "Could not retrieve Subscriptions. "
                GoTo CloseOut
            End If

            ' Debug output
            If Debug = "Y" Then
                mydebuglog.Debug(" Subscriptions Columns found: " & dt1.Columns.Count.ToString)
                mydebuglog.Debug(" Subscriptions Rows found: " & dt1.Rows.Count.ToString)
            End If

            ' Locate Subscription UGA in datatable
            Try
                dr1 = dt1.Select("name='" & SubId & "'")
                If dr1 Is Nothing Or dr1.Length = 0 Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Subscription not found")
                Else
                    SubUGA = dr1(0).Item("row_id").ToString
                End If
            Catch ex As Exception
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug(" Subscription UGA row_id: " & SubUGA)
            End If
        End If

        ' ============================================
        ' Load Domain Users from and/or into cache
        Dim dt2 As DataTable = New DataTable
        Dim dr2() As DataRow
        If Domain <> "" Then
            If Not DMSCache.GetCachedItem("Domains") Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Domains found in cache")
                Try
                    dt2 = DMSCache.GetCachedItem("Domains")
                Catch ex As Exception
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                    GoTo CloseOut
                End Try
            Else
                SqlS = "SELECT DISTINCT G.name, A.row_id " & _
                "FROM DMS.dbo.Groups G " & _
                "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=G.row_id " & _
                "WHERE G.type_cd='Domain' AND A.type_id='G'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Domains into cache: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If d_dr.HasRows Then
                        dt2.Load(d_dr)
                        DMSCache.AddToCache("Domains", dt2, CachingWrapper.CachePriority.Default)
                    End If
                    d_dr.Close()
                Catch ex As Exception
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                    GoTo CloseOut
                End Try
            End If
            If dt2 Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                results = False
                errmsg = errmsg & vbCrLf & "Could not retrieve Domains. "
                GoTo CloseOut
            End If

            ' Debug output
            If Debug = "Y" Then
                mydebuglog.Debug(" Domains Columns found: " & dt2.Columns.Count.ToString)
                mydebuglog.Debug(" Domains Rows found: " & dt2.Rows.Count.ToString)
            End If

            ' Locate Domain UGA in datatable
            Try
                dr2 = dt2.Select("name='" & Domain & "'")
                If dr2 Is Nothing Or dr2.Length = 0 Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Domain not found")
                    errmsg = errmsg & vbCrLf & "Could not find Domain. "
                Else
                    DomainUGA = dr2(0).Item("row_id").ToString
                End If
            Catch ex As Exception
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug(" Domain UGA row_id: " & DomainUGA)
            End If
        End If

        ' ============================================
        ' Load Contact Users from and/or into cache
        Dim dt3 As DataTable = New DataTable
        Dim dr3() As DataRow
        If ConId <> "" Then
            If Not DMSCache.GetCachedItem("Contacts") Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Contacts found in cache")
                Try
                    dt3 = DMSCache.GetCachedItem("Contacts")
                Catch ex As Exception
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                    GoTo CloseOut
                End Try
            Else
                SqlS = "SELECT DISTINCT U.ext_user_id as name, A.row_id " & _
                "FROM DMS.dbo.Users U " & _
                "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=U.row_id " & _
                "WHERE A.type_id='U' AND U.ext_user_id IS NOT NULL"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Contacts into cache: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If d_dr.HasRows Then
                        dt3.Load(d_dr)
                        DMSCache.AddToCache("Contacts", dt3, CachingWrapper.CachePriority.Default)
                    End If
                    d_dr.Close()
                Catch ex As Exception
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                    GoTo CloseOut
                End Try
            End If
            If dt3 Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                results = False
                errmsg = errmsg & vbCrLf & "Could not retrieve Contacts. "
                GoTo CloseOut
            End If

            ' Debug output
            If Debug = "Y" Then
                mydebuglog.Debug(" Contacts Columns found: " & dt3.Columns.Count.ToString)
                mydebuglog.Debug(" Contacts Rows found: " & dt3.Rows.Count.ToString)
            End If

            ' Locate Domain UGA in datatable
            Try
                dr3 = dt3.Select("name='" & ConId & "'")
                If dr3 Is Nothing Or dr3.Length = 0 Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Contact not found")
                    'errmsg = errmsg & vbCrLf & "Could not find Contact. "
                    ConUGA = ""
                Else
                    ConUGA = dr3(0).Item("row_id").ToString
                End If
            Catch ex As Exception
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug(" Contact UGA row_id: " & ConUGA)
            End If
        End If

        ' ============================================
        ' Load Registration Contact Users from and/or into cache
        Dim dt4 As DataTable = New DataTable
        Dim dr4() As DataRow
        If RegId <> "" Then
            If Not DMSCache.GetCachedItem("Registrations") Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Registrations found in cache")
                Try
                    dt4 = DMSCache.GetCachedItem("Registrations")
                Catch ex As Exception
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                    GoTo CloseOut
                End Try
            Else
                SqlS = "SELECT DISTINCT U.ext_id as name, A.row_id " & _
                "FROM DMS.dbo.Users U " & _
                "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=U.row_id " & _
                "WHERE A.type_id='U' AND U.ext_id IS NOT NULL"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Registrations into cache: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If d_dr.HasRows Then
                        dt4.Load(d_dr)
                        DMSCache.AddToCache("Registrations", dt4, CachingWrapper.CachePriority.Default)
                    End If
                    d_dr.Close()
                Catch ex As Exception
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                    GoTo CloseOut
                End Try
            End If
            If dt4 Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                results = False
                errmsg = errmsg & vbCrLf & "Could not retrieve Registrations. "
                GoTo CloseOut
            End If

            ' Debug output
            If Debug = "Y" Then
                mydebuglog.Debug(" Registrations Columns found: " & dt4.Columns.Count.ToString)
                mydebuglog.Debug(" Registrations Rows found: " & dt4.Rows.Count.ToString)
            End If

            ' Locate Domain UGA in datatable
            Try
                dr4 = dt4.Select("name='" & RegId & "'")
                If dr4 Is Nothing Or dr4.Length = 0 Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Registration not found")
                    errmsg = errmsg & vbCrLf & "Could not find Registration. "
                Else
                    RegUGA = dr4(0).Item("row_id").ToString
                End If
            Catch ex As Exception
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug(" Registration UGA row_id: " & RegUGA)
            End If
        End If

        ' ============================================
        ' Close data objects
        Try
            dr1 = Nothing
            dt1 = Nothing
        Catch ex As Exception
        End Try
        Try
            dr2 = Nothing
            dt2 = Nothing
        Catch ex As Exception
        End Try
        Try
            dr3 = Nothing
            dt3 = Nothing
        Catch ex As Exception
        End Try
        Try
            dr4 = Nothing
            dt4 = Nothing
        Catch ex As Exception
        End Try
        Try
            DMSCache = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Write Document User records
        '
        ' Subscription User Group Access Id
        If SubUGA <> "" Then
            SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
            "VALUES (1, 1, " & DocId & ", " & SubUGA & ", '" & SubOwner & "', '" & SubRights & "')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Subscription: " & vbCrLf & SqlS)
            If Debug <> "T" Then
                d_cmd.CommandText = SqlS
                Try
                    returnv = d_cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End If
        End If

        ' Domain User Group Access Id
        If DomainUGA <> "" Then
            SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
            "VALUES (1, 1, " & DocId & ", " & DomainUGA & ", '" & DomainOwner & "', '" & DomainRights & "')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Domain: " & vbCrLf & SqlS)
            If Debug <> "T" Then
                d_cmd.CommandText = SqlS
                Try
                    returnv = d_cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End If
        End If

        ' Contact User Group Access Id
        If ConUGA <> "" Then
            SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
            "VALUES (1, 1, " & DocId & ", " & ConUGA & ", '" & ConOwner & "', '" & ConRights & "')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Contact: " & vbCrLf & SqlS)
            If Debug <> "T" Then
                d_cmd.CommandText = SqlS
                Try
                    returnv = d_cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End If
        End If

        ' Web Registration User Group Access Id
        If RegUGA <> "" Then
            SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
            "VALUES (1, 1, " & DocId & ", " & RegUGA & ", '" & RegOwner & "', '" & RegRights & "')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Registration: " & vbCrLf & SqlS)
            If Debug <> "T" Then
                d_cmd.CommandText = SqlS
                Try
                    returnv = d_cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End If
        End If
        results = True

        ' Create access for supervisor if no contact was specified just to make sure someone "owns" the document
        If ConUGA = "" And RegUGA = "" Then
            SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                   "VALUES (1, 1, " & DocId & ", 1, 'Y', 'REDO')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Supervisor: " & vbCrLf & SqlS)
            If Debug <> "T" Then
                d_cmd.CommandText = SqlS
                Try
                    returnv = d_cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End If
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for document " & DocId
        If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDocUser :  Error: " & Trim(errmsg))
        myeventlog.Info("SaveDMSDocUser : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return results
    End Function

    <WebMethod(Description:="Save or update the provided document into the DMS")> _
    Public Function SaveDMSDoc(ByVal DImage As String, ByVal DocId As String, ByVal ReportId As String, ByVal ItemName As String, ByVal DFileName As String, ByVal DURL As String, _
        ByVal Description As String, ByVal FileExt As String, ByVal ExtId As String, ByVal Domain As String, ByVal DRights As String, ByVal DAccess As String, _
        ByVal UserId As String, ByVal ConId As String, ByVal URights As String, ByVal UAccess As String, ByVal SubId As String, ByVal SRights As String, ByVal SAccess As String, ByVal Debug As String) As String

        ' This function locates the specified item, and returns it to the calling system as a binary

        ' The input parameters are as follows:
        '
        '   DImage	    - The Base64 encoded binary of the image to be stored (req.)
        '   DURL        - The URL to the binary of the image to be stored (req.)
        '   DocId   	- The base64/reversed version of "DMS.Documents.row_id" of the document to be stored - 
        '		            if blank, then create a new document, otherwise update this one. (opt.)
        '   ReportId    - The CX_REP_ENT_SCHED.ROW_ID of a report document to be stored (opt.)
        '   ItemName	- The "DMS.Documents.name" of the document to be stored. (req.)
        '   DFileName   - The "DMS.Documents.dfilename" of the document to be stored. (req.)
        '   Description - The description of the document (req.)
        '   FileExt     - The file extension of the document to be stored (req.)
        '   ExtId	    - The external id of the document to be stored (opt.)

        '   Domain      - The Domain Groups.ext_id of the owner (opt.)
        '   DRights	    - The Domain access rights (opt.)
        '   DAccess     - The Domain user access flag
        '   UserId	    - The User Users.ext_id - the registration id of the person who created the document (opt.)
        '   ConId       - The S_CONTACT.ROW_ID of the person who the document is directly linked to (opt.)
        '   URights	    - The User access rights (opt.)
        '   UAccess     - The User access flag
        '   SubId	    - The Subscription Groups.ext_id - the subscription for this document (opt.)
        '   SRights	    - The Subscription access rights (opt.)
        '   SAccess     - The Subscription user access flag
        '   Debug	    - The debug mode flag: "Y", "N" or "T" 

        ' web.config Parameters used:
        '   dms        	    - connection string to DMS.dms database

        ' Variables
        Dim results, temp As String
        Dim mypath, errmsg, logging As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' HCIDB Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SDDDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service
        Dim BasicService As New basic.com.certegrity.cloudsvc.Service
        Dim DmsService As New local.hq.dms.Service

        ' Local Cache declarations
        Dim DMSCache As New CachingWrapper.LocalCache

        ' File handling declarations
        Dim bfs As FileStream
        Dim bw As BinaryWriter
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim VerifiedSize As Long
        Dim rsize As String

        ' Document declarations
        Dim d_dsize, SaveDest As String
        Dim d_ext As String
        Dim DummyKey As String
        Dim DecodedDocId, VerifiedDocId As String
        Dim tempfile, AddlDesc As String
        Dim data_type_id As String
        Dim DomainGroupId, SubGroupId, UserDMSId As String
        Dim Trainer, PartId As String
        Dim DocVersionId As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        d_ext = ""
        d_dsize = ""
        DomainGroupId = ""
        SubGroupId = ""
        UserDMSId = ""
        data_type_id = ""
        BinaryFile = ""
        d_ConnS = ""
        DecodedDocId = ""
        VerifiedDocId = ""
        tempfile = ""
        DummyKey = ""
        VerifiedSize = 0
        Trainer = ""
        PartId = ""
        DocVersionId = ""
        AddlDesc = ""

        ' ============================================
        ' Get and fix parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            Domain = ""
            ItemName = ""
        Else
            If DocId.Trim <> "" Then
                DocId = Trim(HttpUtility.UrlEncode(DocId))
                If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))
                If InStr(DocId, "%") > 0 Then DocId = Trim(DocId)
                DecodedDocId = Trim(FromBase64(ReverseString(DocId)))
            Else
                DecodedDocId = ""
            End If
            ' Meta data            
            If InStr(ItemName, "%") > 0 Then ItemName = Trim(HttpUtility.UrlDecode(ItemName))
            If InStr(DFileName, "%") > 0 Then DFileName = Trim(HttpUtility.UrlDecode(DFileName))
            If ItemName = "" Then ItemName = DFileName ' If Document.name is blank make it the same as the filename
            If InStr(Description, "%") > 0 Then Description = Trim(HttpUtility.UrlDecode(Description))
            If InStr(FileExt, "%") > 0 Then FileExt = Trim(HttpUtility.UrlDecode(FileExt))
            FileExt = Trim(FileExt.ToLower)
            If InStr(LCase(DFileName), FileExt) = 0 Then DFileName = DFileName & "." & FileExt
            If InStr(ExtId, "%") > 0 Then ExtId = Trim(HttpUtility.UrlDecode(ExtId))
            If InStr(ExtId, " ") > 0 Then ExtId = EncodeParamSpaces(ExtId)
            ' Security data
            If InStr(Domain, "%") > 0 Then Domain = Trim(HttpUtility.UrlDecode(Domain))
            If InStr(Domain, " ") > 0 Then Domain = EncodeParamSpaces(Domain)
            If InStr(DRights, "%") > 0 Then DRights = Trim(HttpUtility.UrlDecode(DRights))
            If InStr(DAccess, "%") > 0 Then DAccess = Trim(HttpUtility.UrlDecode(DAccess))
            UserId = Trim(UserId)
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, " ") > 0 Then UserId = EncodeParamSpaces(UserId)
            If InStr(URights, "%") > 0 Then URights = Trim(HttpUtility.UrlDecode(URights))
            If InStr(UAccess, "%") > 0 Then UAccess = Trim(HttpUtility.UrlDecode(UAccess))
            SubId = Trim(SubId)
            If InStr(SubId, "%") > 0 Then SubId = Trim(HttpUtility.UrlDecode(SubId))
            If InStr(SubId, " ") > 0 Then SubId = EncodeParamSpaces(SubId)
            If InStr(SRights, "%") > 0 Then SRights = Trim(HttpUtility.UrlDecode(SRights))
            If InStr(SAccess, "%") > 0 Then SAccess = Trim(HttpUtility.UrlDecode(SAccess))
            ConId = Trim(ConId)
            If InStr(ConId, "%") > 0 Then ConId = Trim(HttpUtility.UrlDecode(ConId))
            If InStr(ConId, " ") > 0 Then ConId = EncodeParamSpaces(ConId)
        End If

        ' ============================================
        ' Get system defaults
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDoc_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
            AccessKey = System.Configuration.ConfigurationManager.AppSettings("minio-key")
            If AccessKey = "" Then AccessKey = "dms"
            AccessSecret = System.Configuration.ConfigurationManager.AppSettings("minio-secret")
            If AccessSecret <> "" Then AccessSecret = System.Web.HttpUtility.HtmlDecode(AccessSecret)
            If AccessSecret = "" Then AccessSecret = "SfI@aUE$?=&KcAOI?C5NU|-c*Oec7ZPJ"
            AccessRegion = System.Configuration.ConfigurationManager.AppSettings("minio-region")
            If AccessRegion = "" Then AccessRegion = "us-east"
            AccessBucket = System.Configuration.ConfigurationManager.AppSettings("minio-bucket")
            If AccessBucket = "" Then AccessBucket = "default"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\SaveDMSDoc.log"
            Try
                log4net.GlobalContext.Properties("SDDLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug(" -Document Data-")
                mydebuglog.Debug("  DURL: " & DURL)
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  ReportId: " & ReportId)
                mydebuglog.Debug("  DecodedDocId: " & DecodedDocId)
                mydebuglog.Debug("  ItemName: " & ItemName)
                mydebuglog.Debug("  DFileName: " & DFileName)
                mydebuglog.Debug("  Description: " & Description)
                mydebuglog.Debug("  FileExt: " & FileExt)
                mydebuglog.Debug("  ExtId: " & ExtId)
                mydebuglog.Debug(" -Access Control-")
                mydebuglog.Debug("  Domain: " & Domain)
                mydebuglog.Debug("  DRights: " & DRights)
                mydebuglog.Debug("  DAccess: " & DAccess)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  ConId: " & ConId)
                mydebuglog.Debug("  URights: " & URights)
                mydebuglog.Debug("  UAccess: " & UAccess)
                mydebuglog.Debug("  SubId: " & SubId)
                mydebuglog.Debug("  SRights: " & SRights)
                mydebuglog.Debug("  SAccess: " & SAccess)
                mydebuglog.Debug(" -Object Access-")
                mydebuglog.Debug("  AccessBucket: " & AccessBucket)
                mydebuglog.Debug("  AccessRegion: " & AccessRegion & vbCrLf)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(ItemName) = "" And ReportId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No document name specified. "
            GoTo CloseOut2
        End If
        If Trim(Description) = "" And ReportId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No description specified. "
            GoTo CloseOut2
        End If
        If Trim(DFileName) = "" And ReportId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No document filename specified. "
            GoTo CloseOut2
        End If
        If Trim(FileExt) = "" And ReportId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No extension specified. "
            GoTo CloseOut2
        End If
        If DocId.Trim <> "" Then
            If Not IsNumeric(DecodedDocId) Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Incorrect document id specified. "
                GoTo CloseOut2
            End If
        End If

        ' ============================================
        ' Open SQL Server database connection to DMS
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Or d_cmd Is Nothing Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Unable to open DMS connection. "
            GoTo CloseOut
        End If

        ' ============================================
        ' If ReportId is provided, then retrieve from this database and store in a buffer
        '   Retrieve other values from the report schedule record
        If ReportId <> "" Then

            ' If dealing with reports, open HCIDB connection
            errmsg = OpenDBConnection(ConnS, con, cmd)
            cmd.CommandTimeout = 120            ' Set timeout to 2 mins

            If errmsg <> "" Or d_cmd Is Nothing Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Unable to open HCIDB connection. "
                GoTo CloseOut
            End If

            ' Retrieve report information
            rsize = ""
            SqlS = "SELECT S.DSIZE, S.DIMAGE, R.DESCRIPTION+(SELECT CASE WHEN S.ADDL_DESC IS NOT NULL THEN S.ADDL_DESC ELSE '' END), S.FORMAT, " & _
            "(SELECT CASE WHEN S.XFER_ID IS NULL THEN S.ROW_ID ELSE S.XFER_ID END)+'.'+LOWER(S.FORMAT) AS DFILENAME, " & _
            "R.NAME AS ITEMNAME " & _
            "FROM siebeldb.dbo.CX_REP_ENT_SCHED S " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_REPORT_ENT E ON E.ROW_ID=S.ENT_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_REPORTS R ON R.ROW_ID=E.REPORT_ID " & _
            "WHERE S.ROW_ID='" & ReportId & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Retrieving report: " & vbCrLf & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            rsize = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            temp = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            If UCase(Trim(temp)) <> UCase(Trim(Description)) Then Description = Description & temp
                            If Debug = "Y" Then mydebuglog.Debug("  >Description: " & Description)
                            If FileExt = "" Then FileExt = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            DFileName = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            If ItemName = "" Then ItemName = Trim(CheckDBNull(dr(5), enumObjectType.StrType))

                            ' Get binary and attach to the object outbyte if found, not cached or updated recently
                            '   retval will be "0" if this is not the case
                            If rsize <> "" Then
                                ReDim outbyte(Val(rsize) - 1)
                                startIndex = 0
                                retval = dr.GetBytes(1, 0, outbyte, 0, rsize)
                            End If
                        Catch obug As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting report - read failure. " & obug.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Error getting report - datarecord failure." & vbCrLf
                    results = "Failure"
                End If

            Catch ex As Exception
                errmsg = errmsg & "Error getting report - command failure. " & vbCrLf & ex.Message
            End Try
            Try
                dr.Close()
            Catch ex As Exception
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "Report Information-")
                mydebuglog.Debug("  Size: " & rsize)
                mydebuglog.Debug("  ItemName: " & ItemName)
                mydebuglog.Debug("  DFileName: " & DFileName)
                mydebuglog.Debug("  Description: " & Description)
                mydebuglog.Debug("  FileExt: " & FileExt)
            End If

            ' If unable to locate the report, then error out
            If retval = 0 Or FileExt = "" Then
                results = "Failure"
                errmsg = errmsg & "Error getting report." & vbCrLf
                GoTo CloseOut
            End If

            ' Lookup the web registration id of the contact id if not previously looked up
            If UserId = "" And ConId <> "" Then
                SqlS = "SELECT TOP 1 C.X_REGISTRATION_NUM, SC.SUB_ID, S.DOMAIN, C.X_TRAINER_NUM, C.X_PART_ID " & _
                "FROM siebeldb.dbo.S_CONTACT C " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID=C.ROW_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID " & _
                "WHERE C.ROW_ID='" & ConId & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Get contact information: " & vbCrLf & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                ' Set basic access rights as well 
                                UserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                URights = "REDO"
                                UAccess = "Y"
                                SubId = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                SRights = "R"
                                SAccess = "N"
                                Domain = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                DRights = "R"
                                DAccess = "N"
                                Trainer = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                PartId = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                            Catch ex As Exception
                                errmsg = errmsg & "The domain contact information was not found. " & ex.ToString
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "The domain contact information was not found."
                    End If
                Catch ex As Exception
                End Try
                Try
                    dr.Close()
                Catch ex As Exception
                End Try
            End If

            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "Contact Information-")
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  SubId: " & SubId)
                mydebuglog.Debug("  Domain: " & Domain)
                mydebuglog.Debug("  Trainer: " & Trainer)
                mydebuglog.Debug("  PartId: " & PartId)
            End If
        End If

        ' ============================================
        ' Create output directory for temp file caching if needed
        SaveDest = mypath & "work_dir\" & FileExt
        Try
            Directory.CreateDirectory(SaveDest)
        Catch
        End Try

        ' ============================================
        ' Determine data_type_id based on the following parameter:
        '   FileExt - The file extension of the document to be stored (req.)
        Dim dt As DataTable = New DataTable
        If Not DMSCache.GetCachedItem("DocumentTypes") Is Nothing Then
            ' Get document types from cache
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Document Types found in cache")
            Try
                dt = DMSCache.GetCachedItem("DocumentTypes")
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                GoTo CloseOut
            End Try
        Else
            ' Load document types into cache
            SqlS = "SELECT extension, row_id FROM DMS.dbo.Document_Types"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Document Types into cache: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If d_dr.HasRows Then
                    dt.Load(d_dr)
                    DMSCache.AddToCache("DocumentTypes", dt, CachingWrapper.CachePriority.NotRemovable)
                End If
                d_dr.Close()
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                GoTo CloseOut
            End Try
        End If
        If dt Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not retrieve Document Types. "
            GoTo CloseOut
        End If

        ' Debug output
        If Debug = "Y" Then
            mydebuglog.Debug(" Document Types Columns found: " & dt.Columns.Count.ToString)
            mydebuglog.Debug(" Document Types Rows found: " & dt.Rows.Count.ToString)
        End If

        ' ============================================
        ' Locate Document Type Id in datatable
        Dim drow() As DataRow = dt.Select("extension='" & FileExt & "'")
        If drow Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Document Type not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not find association. "
            GoTo CloseOut
        End If
        data_type_id = drow(0).Item("row_id").ToString
        If Debug = "Y" Then mydebuglog.Debug(" Document Types row_id: " & data_type_id)
        Try
            dr = Nothing
            dt = Nothing
        Catch ex As Exception
        End Try

        ' Check data type id
        If data_type_id = "" Then
            results = "Failure"
            errmsg = errmsg & "Unknown data type id. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Create or verify basic document record in the DMS if needed and determine document id
        '   If a document id was supplied, it is assumed that we are overwriting.
        If DecodedDocId <> "" Then

            ' Query DMS for existing document id
            SqlS = "SELECT TOP 1 row_id " & _
                "FROM DMS.dbo.Documents " & _
                "WHERE row_id=" & DecodedDocId
            If Debug = "Y" Then mydebuglog.Debug("  Verify provided DocId: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            VerifiedDocId = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                            If Debug = "Y" Then mydebuglog.Debug("  > Verified Id=" & VerifiedDocId)
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & vbCrLf & "Error verifying supplied doc id. " & ex.ToString
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & vbCrLf & "Error verifying supplied doc id."
                    d_dr.Close()
                    results = "Failure"
                    GoTo CloseOut
                End If
                d_dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error verifying supplied doc id: " & oBug.ToString)
                results = "Failure"
            End Try

        Else
            ' Generate a dummy key to use when creating a document
            DummyKey = BasicService.GeneratePassword(Debug)
            If Debug = "Y" Then mydebuglog.Debug("  Generated DummyKey: " & DummyKey)

            ' Create new record in DMS and verify that it was created
            SqlS = "INSERT DMS.dbo.Documents (created, created_by, last_upd, last_upd_by, name, ext_id) " & _
            "VALUES (GETDATE(), 1, GETDATE(), 1, '" & SqlString(ItemName) & "','" & DummyKey & "') "
            If Debug = "Y" Then mydebuglog.Debug("  Create basic document: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try

            ' Locate created record id
            SqlS = "SELECT row_id " & _
            "FROM DMS.dbo.Documents " & _
            "WHERE ext_id='" & DummyKey & "' and name='" & SqlString(ItemName) & "'"
            If Debug = "Y" Then mydebuglog.Debug("  Locate basic DocId: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            VerifiedDocId = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                            If Debug = "Y" Then mydebuglog.Debug("  > Verified Id=" & VerifiedDocId)
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error locating dummy doc id. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Error locating basic doc id." & vbCrLf
                    d_dr.Close()
                    results = "Failure"
                End If
                d_dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error locating basic doc id: " & oBug.ToString)
                results = "Failure"
            End Try
        End If

        ' Check document id
        If VerifiedDocId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Unable to verify or create document"
            GoTo CloseOut
        End If

        ' ============================================
        ' Extract and store binary locally

        ' Generate and remove temp file if applicable
        '   The filename of the temp file is in the format [Document Id]+.+[Extension]
        tempfile = SaveDest & "\" & VerifiedDocId & "." & FileExt
        If Debug = "Y" Then mydebuglog.Debug("   Temp file: " & tempfile & vbCrLf)
        Try
            If (My.Computer.FileSystem.FileExists(tempfile)) Then Kill(tempfile)
        Catch ex As Exception
        End Try

        ' If a Report stored in a buffer, cache in a local file
        If retval > 0 Then
            Try
                bfs = New FileStream(tempfile, FileMode.Create, FileAccess.Write)
                bw = New BinaryWriter(bfs)
                bw.Write(outbyte)
                bw.Flush()
                bw.Close()
                bfs.Close()
                Try
                    bfs.Dispose()
                    bfs = Nothing
                Catch ex2 As Exception
                End Try
                Try
                    bw.Dispose()
                    bw = Nothing
                Catch ex2 As Exception
                End Try
            Catch ex As Exception
                errmsg = errmsg & "Unable to write the report file to a temp file." & ex.ToString & vbCrLf
                results = "Failure"
                retval = 0
            End Try
            d_dsize = retval.ToString
        End If

        ' If a URL is provided, retrieve the data from that URL and cache locally
        If DURL <> "" Then
            Dim oRequest As System.Net.HttpWebRequest = CType(System.Net.HttpWebRequest.Create(DURL), System.Net.HttpWebRequest)
            Using oResponse As System.Net.WebResponse = CType(oRequest.GetResponse, System.Net.WebResponse)
                Using responseStream As System.IO.Stream = oResponse.GetResponseStream
                    Using bfs2 As New FileStream(tempfile, FileMode.Create, FileAccess.Write)
                        Dim buffer(2047) As Byte
                        Dim read As Integer
                        Do
                            read = responseStream.Read(buffer, 0, buffer.Length)
                            bfs2.Write(buffer, 0, read)
                        Loop Until read = 0
                        responseStream.Close()
                        bfs2.Flush()
                        bfs2.Close()
                        Try
                            bfs2.Dispose()
                        Catch ex As Exception
                        End Try
                        Try
                            buffer = Nothing
                        Catch ex As Exception
                        End Try
                    End Using
                    responseStream.Close()
                    responseStream.Dispose()
                End Using
                oResponse.Close()
            End Using
            Try
                oRequest = Nothing
            Catch ex As Exception
            End Try
            d_dsize = tempfile.Length.ToString
        End If

        ' If base64 encoded binary is supplied, retrieve and cache locally
        If Len(DImage) > 0 Then
            ' Convert input string into byte array, and then into binary
            Dim imagebuffer As Byte() = Convert.FromBase64String(DImage)
            If imagebuffer.Length = 0 Then
                ' No attachment found
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No attachment error. "
                GoTo CloseOut2
            End If
            Try
                bfs = New FileStream(tempfile, FileMode.Create, FileAccess.Write)
                bw = New BinaryWriter(bfs)
                bw.Write(imagebuffer)
                bw.Flush()
                bw.Close()
                bfs.Close()
                Try
                    bfs.Dispose()
                    bfs = Nothing
                Catch ex2 As Exception
                End Try
                Try
                    bw.Dispose()
                    bw = Nothing
                Catch ex2 As Exception
                End Try
            Catch ex As Exception
                errmsg = errmsg & "Unable to write the file to a temp file." & ex.ToString & vbCrLf
                results = "Failure"
                retval = 0
            End Try
            d_dsize = Len(DImage).ToString
        End If

        ' ============================================
        ' Create Document_Versions record
        If ReportId <> "" Then
            ' If a report, set to not backup
            SqlS = "INSERT DMS.dbo.Document_Versions (doc_id, created, created_by, last_upd, last_upd_by, backed_up, dsize, version, minio_flg) " &
            "SELECT " & VerifiedDocId & ", GETDATE(), 1, GETDATE(), 1, GETDATE(), '" & d_dsize & "', ISNULL(MAX([version]),0)+1, 'Y' FROM DMS.dbo.Document_Versions WHERE doc_id=" & VerifiedDocId & "; select Scope_Identity();"
        Else
            SqlS = "INSERT DMS.dbo.Document_Versions (doc_id, created, created_by, last_upd, last_upd_by, dsize, version, minio_flg) " &
            "SELECT " & VerifiedDocId & ", GETDATE(), 1, GETDATE(), 1, '" & d_dsize & "', ISNULL(MAX([version]),0)+1, 'Y' FROM DMS.dbo.Document_Versions WHERE doc_id=" & VerifiedDocId & "; select Scope_Identity();"
        End If
        If Debug = "Y" Then mydebuglog.Debug("  Create basic document versions record: " & SqlS)
        Try
            d_cmd.CommandText = SqlS
            DocVersionId = d_cmd.ExecuteScalar()
            If Debug = "Y" Then mydebuglog.Debug("  > DocVersionId=" & DocVersionId & vbCrLf)
        Catch ex As Exception
            results = "Failure"
            errmsg = errmsg & "Error creating basic document versions record. " & ex.ToString & vbCrLf
            GoTo CloseOut
        End Try

        ' Locate the Document_Versions record created
        'SqlS = "SELECT row_id " & _
        '"FROM DMS.dbo.Document_Versions " & _
        '"WHERE doc_id=" & VerifiedDocId
        'If Debug = "Y" Then mydebuglog.Debug("  Locate DocVersionId: " & SqlS)
        'Try
        '    d_cmd.CommandText = SqlS
        '    d_dr = d_cmd.ExecuteReader()
        '    If Not d_dr Is Nothing Then
        '        While d_dr.Read()
        '            Try
        '                DocVersionId = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
        '                If Debug = "Y" Then mydebuglog.Debug("  > DocVersionId=" & DocVersionId)
        '            Catch ex As Exception
        '                results = "Failure"
        '                errmsg = errmsg & "Error locating DocVersionId. " & ex.ToString & vbCrLf
        '                GoTo CloseOut
        '            End Try
        '        End While
        '    Else
        '        errmsg = errmsg & "Error locating DocVersionId." & vbCrLf
        '        d_dr.Close()
        '        results = "Failure"
        '    End If
        '    d_dr.Close()
        'Catch oBug As Exception
        '   If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error locating DocVersionId: " & oBug.ToString)
        '   results = "Failure"
        'End Try

        ' ============================================
        ' Retrieve document record and update with information supplied
        If DocVersionId <> "" Then

            ' Set configuration
            Dim MConfig As AmazonS3Config = New AmazonS3Config()
            'MConfig.RegionEndpoint = RegionEndpoint.USEast1
            MConfig.ServiceURL = "https://192.168.5.134"
            MConfig.ForcePathStyle = True
            MConfig.EndpointDiscoveryEnabled = False

            Dim Minio As AmazonS3Client = New AmazonS3Client(AccessKey, AccessSecret, MConfig)
            Try
                ServicePointManager.ServerCertificateValidationCallback = AddressOf sslhttps.AcceptAllCertifications
                Dim fileTransfer As Amazon.S3.Transfer.TransferUtility = New Amazon.S3.Transfer.TransferUtility(Minio)

                fileTransfer.Upload(tempfile, AccessBucket, DocVersionId & "-" & DocVersionId)
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Unload complete of " & tempfile & " to bucket " & AccessBucket & " and object id " & DocVersionId)
                Try
                    fileTransfer = Nothing
                Catch ex2 As Exception
                    errmsg = errmsg & "Error closing fileTransfer: " & ex2.Message & vbCrLf
                End Try
            Catch ex As Exception
                errmsg = errmsg & "Error writing to Minio: " & ex.Message & vbCrLf
                results = "Failure"
            End Try

            Try
                Minio = Nothing
            Catch ex As Exception
                errmsg = errmsg & "Error closing Minio: " & ex.Message & vbCrLf
            End Try

            'Dim da As New SqlDataAdapter("SELECT TOP 1 * FROM DMS.dbo.Document_Versions WHERE row_id=" & DocVersionId, d_con)
            'Dim MyCB As SqlCommandBuilder = New SqlCommandBuilder(da)
            'Dim ds As New Data.DataSet()

            ' Open database
            'da.MissingSchemaAction = Data.MissingSchemaAction.AddWithKey
            'Try
            'd_con.Open()
            'Catch ex As Exception
            'End Try
            'da.Fill(ds, "Document_Versions")

            ' Get file and attach to MyData object (can use imagebuffer instead?)
            'Dim mstream As New System.IO.FileStream(tempfile, FileMode.OpenOrCreate, FileAccess.Read)
            'd_dsize = mstream.Length
            'Dim MyData(d_dsize) As Byte
            'mstream.Read(MyData, 0, d_dsize)
            'mstream.Close()

            'ds.Tables("Document_Versions").Rows(0)("dimage") = MyData
            'ds.Tables("Document_Versions").Rows(0)("dsize") = d_dsize
            'da.Update(ds, "Document_Versions")
            'If Debug = "Y" Or logging = "Y" Then mydebuglog.Debug("  Saved " & Trim(d_dsize.ToString) & " bytes in file of type " & data_type_id & " with ext_id " & ExtId)
            'VerifiedSize = d_dsize

            ' Close objects created
            'Try
            ' mstream.Close()
            'mstream.Dispose()
            'mstream = Nothing
            'Catch ex As Exception
            'End Try
            'Try
            'da.Dispose()
            'da = Nothing
            'Catch ex As Exception
            'End Try
            'Try
            'MyCB.Dispose()
            'MyCB = Nothing
            'Catch ex As Exception
            'End Try
            'Try
            'ds.Dispose()
            'ds = Nothing
            'Catch ex As Exception
            'End Try
        End If

        ' ============================================
        ' Update Document record
        '   ItemName	    - The "DMS.Documents.name" of the document to be stored. (req.)
        '   DFileName       - The "DMS.Documents.dfilename" of the document to be stored. (req.)
        '   Description     - The description of the document (req.)
        '   ExtId	        - The external id of the document to be stored (opt.)
        '   DocVersionId    - The FK to Document_Versions
        SqlS = "UPDATE DMS.dbo.Documents " & _
            "SET name='" & SqlString(ItemName) & "', dfilename='" & SqlString(DFileName) & "', " & _
            "description='" & SqlString(Description) & "', ext_id='" & SqlString(ExtId) & "', data_type_id=" & data_type_id & _
            ", last_version_id=" & DocVersionId & " WHERE row_id=" & VerifiedDocId
        If Debug = "Y" Then mydebuglog.Debug("  Update Document record: " & SqlS & vbCrLf)
        Try
            d_cmd.CommandText = SqlS
            returnv = d_cmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        ' ============================================
        ' Create security for the document if applicable
        If VerifiedDocId <> "" Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Creating Document_Users records")
            If Domain <> "" Or SubId <> "" Or UserId <> "" Or ConId <> "" Then
                Try
                    If ConId <> "" Then
                        SaveDMSDocUser(VerifiedDocId, Domain, DAccess, DRights, SubId, SAccess, SRights, ConId, UAccess, URights, "", "", "", Debug)
                    Else
                        SaveDMSDocUser(VerifiedDocId, Domain, DAccess, DRights, SubId, SAccess, SRights, "", "", "", UserId, UAccess, URights, Debug)
                    End If
                Catch ex As Exception
                    errmsg = errmsg & "Unable to create access rights." & ex.ToString & vbCrLf
                End Try
            End If
        End If

        ' ============================================
        ' Process supplied contact id
        If ConId <> "" Then

            ' Create contact association for the document if applicable
            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                "VALUES (1, 1, 3, " & VerifiedDocId & ", '" & ConId & "', 'Y', 'Y', 'N')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Creating Document_Associations record for Contact: " & vbCrLf & SqlS & vbCrLf)
            d_cmd.CommandText = SqlS
            Try
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
                errmsg = errmsg & "Could not update the Document_Associations table" & vbCrLf & ex.Message
                If Debug = "Y" Then mydebuglog.Debug(errmsg)
            End Try

            ' Update document count
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Updating document count for document id: " & VerifiedDocId & " and contact id '" & ConId & "'")
            UpdDMSDocCount(ConId, Trainer, PartId, "", Debug)

        End If

        ' ============================================
        ' If report, set document category and keywords
        If ReportId <> "" Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Creating Document_Categories records")
            SaveDMSDocCat(VerifiedDocId, "Report Documents", "Y", Debug)

            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Creating Document_Keywords records")
            SaveDMSDocKey(VerifiedDocId, "Report Only", ConId, "Y", Debug)
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        If Debug = "Y" Then mydebuglog.Debug("Closing database connections " & vbCrLf)
        '   DMS
        Try
            d_dr = Nothing
            d_con.Dispose()
            d_con = Nothing
            d_cmd.Dispose()
            d_cmd = Nothing
        Catch ex As Exception
        End Try

        '   HCIDB
        Try
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Delete cached local temp file
        If tempfile <> "" Then
            If Debug = "Y" Then mydebuglog.Debug("Attempting to remove temp file: " & tempfile & vbCrLf)
            Try
                If Debug <> "Y" Then
                    If (My.Computer.FileSystem.FileExists(tempfile)) Then Kill(tempfile)
                End If
            Catch ex As Exception
            End Try
        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for item '" & ItemName & "', in '" & Domain & "' domain with filename " & DFileName & "." & FileExt & " and report id " & ReportId & " to document id " & VerifiedDocId
        If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDoc :  Error: " & Trim(errmsg))
        myeventlog.Info("SaveDMSDoc : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return item
        If VerifiedDocId <> "" And Debug <> "T" Then
            Return VerifiedDocId
        Else
            Return Nothing
        End If
    End Function

    <WebMethod(Description:="Logs a user out of the DMS")> _
    Public Function UserLogout(ByVal SessionId As String, ByVal Debug As String) As Boolean

        ' This function closes a DMS session for a user

        ' The input parameters are as follows:
        '
        '   Domain	    - The users subscription domain, Opt.
        '   Debug	    - The debug mode flag: "Y", "N" or "T" 

        ' The result is a boolean indicating success or failure

        ' Variables
        Dim results As String
        Dim success As Boolean
        Dim mypath, errmsg, logging, temp As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("ULDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        success = False
        temp = ""

        ' ============================================
        ' Get and fix parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            SessionId = "1111111111"
        End If

        ' ============================================
        ' Get system defaults
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("UserLogout_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\UserLogout.log"
            Try
                log4net.GlobalContext.Properties("ULLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  Debug: " & Debug)
                mydebuglog.Debug("  SessionId: " & SessionId)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(SessionId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No session id specified. "
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open SQL Server database connection to DMS
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Or d_cmd Is Nothing Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Unable to open DMS connection. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Perform Logout
        Try
            SqlS = "DELETE FROM DMS.dbo.User_Sessions WHERE row_id=" & SessionId
            If Debug = "Y" Then mydebuglog.Debug("  Removing session: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                returnv = d_cmd.ExecuteNonQuery()
                If Debug = "Y" Then mydebuglog.Debug("    > returnv: " & returnv.ToString)
                If returnv > 0 Then success = True
            Catch ex As Exception
                errmsg = errmsg & "Error removing session. " & ex.ToString & vbCrLf
            End Try
        Catch oBug As Exception
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error removing session: " & oBug.ToString)
            results = "Failure"
        End Try

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            d_dr = Nothing
            d_con.Dispose()
            d_con = Nothing
            d_cmd.Dispose()
            d_cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = success.ToString & " for logout of " & SessionId
        If Trim(errmsg) <> "" Then myeventlog.Error("UserLogout :  Error: " & Trim(errmsg))
        myeventlog.Info("UserLogout : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return item
        Return success
    End Function

    <WebMethod(Description:="Check Document Access")> _
    Public Function CheckDocAccess(ByVal ContactId As String, ByVal UserId As String, ByVal SessionId As String, _
        ByVal DocId As String, ByVal LocalCache As String, ByVal Debug As String) As Boolean

        ' This function determines whether the specified user can "see" the document specified

        ' The input parameters are as follows:
        '
        '   ContactId  	- The users S_CONTACT.ROW_ID, Reqd.
        '   UserId  	- The users S_CONTACT.X_REGISTRATION_NUM, Reqd.
        '   SessionId  	- The users web site session id, Reqd.
        '   DocId  	    - The Document.row_id of the record to check, Reqd.
        '   LocalCache  - Locally cache profiles mode
        '   Debug	    - The debug mode flag: "Y", "N" or "T" 

        ' The result is a boolean indicating "yes" or "no"

        ' web.config Parameters used:
        '   dms        	- connection string to DMS.dms database

        ' Variables
        Dim DocAccess As Boolean
        Dim results As String
        Dim mypath, errmsg, logging, temp As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("CDADebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Local Cache declarations
        Dim DMSCache As New CachingWrapper.LocalCache

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service
        Dim http As New simplehttp()
        Dim hciscormsvc_farm As String

        ' Profile declarations
        Dim ProfileFile As String
        Dim serializer As XmlSerializer = New XmlSerializer(GetType(profile))
        Dim UProfile As profile

        ' Misc. declarations
        Dim Category_Constraint, Cat_Accessible As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        temp = ""
        Category_Constraint = ""
        Cat_Accessible = ""
        DocAccess = False

        ' ============================================
        ' Get and fix parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            ContactId = "21120611WE0"
            UserId = "RTO31123036OA"
            SessionId = "11111111"
            DocId = "47"
        Else
            If InStr(ContactId, "%") > 0 Then ContactId = Trim(HttpUtility.UrlDecode(ContactId))
            If InStr(ContactId, " ") > 0 Then ContactId = EncodeParamSpaces(ContactId)
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, " ") > 0 Then UserId = EncodeParamSpaces(UserId)
            If InStr(SessionId, "%") > 0 Then SessionId = Trim(HttpUtility.UrlDecode(SessionId))
            If InStr(SessionId, " ") > 0 Then SessionId = EncodeParamSpaces(SessionId)
            If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))
            If InStr(DocId, " ") > 0 Then DocId = EncodeParamSpaces(DocId)
            If InStr(LocalCache, "%") > 0 Then LocalCache = Trim(HttpUtility.UrlDecode(LocalCache))
            If InStr(LocalCache, " ") > 0 Then LocalCache = EncodeParamSpaces(LocalCache)
            If LocalCache = "" Then LocalCache = "N"
        End If

        ' ============================================
        ' Get system defaults
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("CheckDocAccess_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
            hciscormsvc_farm = System.Configuration.ConfigurationManager.AppSettings.Get("hciscormsvc_farm")
            If hciscormsvc_farm = "" Then hciscormsvc_farm = "192.168.7.43"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\CheckDocAccess.log"
            Try
                log4net.GlobalContext.Properties("CDALogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  ContactId: " & ContactId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  SessionId: " & SessionId)
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  hciscormsvc_farm: " & hciscormsvc_farm)
                mydebuglog.Debug("  LocalCache: " & LocalCache & vbCrLf)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(ContactId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No contact id specified. "
            GoTo CloseOut2
        End If
        If Trim(UserId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No user id specified. "
            GoTo CloseOut2
        End If
        If Trim(SessionId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No session id specified. "
            GoTo CloseOut2
        End If
        If Trim(DocId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No document id specified. "
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open SQL Server database connection to DMS
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Or d_cmd Is Nothing Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Unable to open DMS connection. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Setup for local caching of serialized profiles
        Try
            temp = mypath & "profiles"
            Directory.CreateDirectory(temp)
        Catch ex As Exception
        End Try
        ProfileFile = mypath & "profiles\" & UserId & "-" & SessionId & ".xml"
        If Debug = "Y" And LocalCache = "Y" Then mydebuglog.Debug(" Local Cache ProfileFile: " & ProfileFile & vbCrLf)

        ' ============================================
        ' Retrieve and deserialize profile
        Try
            Try
                UProfile = DMSCache.GetCachedItem(UserId & "." & SessionId & ".xml")
                If Not UProfile Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(" Found profile in memory cache" & vbCrLf)
                    GoTo BuildConstraint
                End If
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug(" Did not find profile in memory cache" & vbCrLf)

            If LocalCache = "Y" And File.Exists(ProfileFile) Then
                If Debug = "Y" Then mydebuglog.Debug(" Retrieving profile from local file cache" & vbCrLf)
                ' Check to see if the profile was already serialized locally and retrieve
                Dim osr As New StreamReader(ProfileFile)
                UProfile = serializer.Deserialize(osr)
                osr.Close()
                osr = Nothing

                ' Add to memory cache
                If DMSCache.GetCachedItem(UserId & "." & SessionId & ".xml") Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(" Caching profile to memory cache" & vbCrLf)
                    DMSCache.AddToCache(UserId & "." & SessionId & ".xml", UProfile, CachingWrapper.CachePriority.Default)
                End If
            Else
                If Debug = "Y" Then mydebuglog.Debug(" Retrieving profile from CouchBase" & vbCrLf)
                ' Retrieve Login Information from Couchbase
                Dim StringUserProfile As String
                StringUserProfile = http.geturl("http://hciscormsvc.certegrity.com/cmp/WebService.asmx/GetUserProfile?UserId=" & UserId & "&SessId=" & SessionId & "&Debug=" & Debug, hciscormsvc_farm, 80, "", "")
                Dim XmlUserProfile As XmlDocument = New XmlDocument()
                XmlUserProfile.LoadXml(StringUserProfile)

                Dim settings As XmlWriterSettings = New XmlWriterSettings()
                settings.OmitXmlDeclaration = True
                settings.ConformanceLevel = ConformanceLevel.Fragment
                settings.CloseOutput = False

                ' Create the XmlWriter object and serialize the profile to it 
                Dim strm As MemoryStream = New MemoryStream()
                Dim writer As XmlWriter = XmlWriter.Create(strm, settings)
                XmlUserProfile.WriteTo(writer)
                writer.Flush()
                writer.Close()
                strm.Position = 0
                writer = Nothing

                ' Serialize to a temp file
                If LocalCache = "Y" Then
                    If Debug = "Y" Then mydebuglog.Debug(" Storing profile to local file cache" & vbCrLf)
                    Dim xwrite As TextWriter = New StreamWriter(ProfileFile)
                    Dim xwriter As XmlTextWriter = New XmlTextWriter(xwrite)
                    XmlUserProfile.WriteTo(xwriter)
                    xwriter.Close()
                    xwriter = Nothing
                    xwrite = Nothing
                End If

                ' Deserialize the memory stream to an instance of a profile class object
                UProfile = serializer.Deserialize(strm)
                strm.Close()
                strm = Nothing

                ' Add to memory cache
                If DMSCache.GetCachedItem(UserId & "." & SessionId & ".xml") Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(" Caching profile to memory cache" & vbCrLf)
                    DMSCache.AddToCache(UserId & "." & SessionId & ".xml", UProfile, CachingWrapper.CachePriority.Default)
                End If

                XmlUserProfile = Nothing
            End If

            Try
                If UProfile.CONTACT_ID = "" Then GoTo CloseOut
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(" Error reading profile: " & ex.Message)
                GoTo CloseOut
            End Try

        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug(" Error reading profile: " & ex.Message)
            GoTo CloseOut
        End Try

        ' ============================================
        ' Create Category Keyword constraint
        ' Build a category keyword constraint for use in association queries
BuildConstraint:
        If Debug = "Y" Then mydebuglog.Debug(" Computing category constraint")
        Category_Constraint = "CK.key_id IN ("
        If UProfile.TRAINER_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "3,"
            Cat_Accessible = Cat_Accessible & "#3"
        End If
        If UProfile.MT_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "5,"
            Cat_Accessible = Cat_Accessible & "#5"
        End If
        If UProfile.PART_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "7,"
            Cat_Accessible = Cat_Accessible & "#7"
        End If
        If UProfile.TRAINING_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "8,"
            Cat_Accessible = Cat_Accessible & "#8"
        End If
        If UProfile.TRAINER_ACC_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "10,"
            Cat_Accessible = Cat_Accessible & "#10"
        End If
        If UProfile.SITE_ONLY = "Y" Then
            Category_Constraint = Category_Constraint & "12,"
            Cat_Accessible = Cat_Accessible & "#12"
        End If
        If UProfile.REPORTS_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "13,"
            Cat_Accessible = Cat_Accessible & "#13"
        End If
        If UProfile.SYSADMIN_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "15,"
            Cat_Accessible = Cat_Accessible & "#15"
        End If
        If UProfile.EMP_ID <> "" Then
            Category_Constraint = Category_Constraint & "16,"
            Cat_Accessible = Cat_Accessible & "#16"
        End If
        Category_Constraint = Category_Constraint & "14) "
        Cat_Accessible = Cat_Accessible & "#14"
        If Debug = "Y" Then
            mydebuglog.Debug("  > Category_Constraint: " & Category_Constraint)
            mydebuglog.Debug("  > Cat_Accessible: " & Cat_Accessible & vbCrLf)
        End If

        ' ============================================
        ' Prepare document query
        If Debug = "Y" Then mydebuglog.Debug(" Computing query")
        SqlS = "SELECT row_id " & _
            "FROM DMS.dbo.Documents " & _
            "WHERE row_id=" & DocId & _
            " INTERSECT ( " & _
            "    SELECT D.row_id " & _
            "    FROM DMS.dbo.Documents D " & _
            "    LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id " & _
            "    WHERE ((DA.association_id='3' AND DA.fkey='" & UProfile.CONTACT_ID & "')"
        If UProfile.TRAINER_ID <> "" Then
            SqlS = SqlS & " OR (DA.association_id='5' AND DA.fkey='" & UProfile.TRAINER_ID & "')"
        End If
        If UProfile.PART_ID <> "" Then
            SqlS = SqlS & " OR (DA.association_id='4' AND DA.fkey='" & UProfile.PART_ID & "')"
        End If
        SqlS = SqlS & ") "

        If UProfile.DMS_USER_AID <> "" Then
            SqlS = SqlS & " UNION " & _
            "    SELECT D.row_id " & _
            "    FROM DMS.dbo.Documents D " & _
            "    INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id " & _
            "    WHERE DU.user_access_id=" & UProfile.DMS_USER_AID & " "
        End If

        If Category_Constraint <> "" Then
            SqlS = SqlS & " UNION " & _
            "    SELECT D.row_id " & _
            "    FROM DMS.dbo.Documents D " & _
            "    LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " & _
            "    LEFT OUTER JOIN DMS.dbo.Category_Keywords CK ON CK.cat_id=DC.cat_id AND (" & Category_Constraint & ") " & _
            "    LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id " & _
            "    WHERE DC.pr_flag='Y' AND C.public_flag='Y' AND CK.key_id IS NOT NULL " & _
            "    GROUP BY D.row_id "
        End If
        SqlS = SqlS & ")"

        If Debug = "Y" Then mydebuglog.Debug("  > Document Access Query: " & SqlS)
        Try
            d_cmd.CommandText = SqlS
            d_dr = d_cmd.ExecuteReader()
            If Not d_dr Is Nothing Then
                While d_dr.Read()
                    Try
                        results = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                        If Debug = "Y" Then mydebuglog.Debug("  > Located Document Id: " & results)
                    Catch ex As Exception
                        results = "Failure"
                        errmsg = errmsg & vbCrLf & "Error locating Document Id. " & ex.ToString
                        GoTo CloseOut
                    End Try
                End While
            End If
            d_dr.Close()
        Catch oBug As Exception
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error verifying supplied doc id: " & oBug.ToString)
            results = "Failure"
        End Try

        If results <> "" Then DocAccess = True
        If Debug = "Y" Then mydebuglog.Debug(" Results: " & DocAccess)
CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            d_dr = Nothing
            d_con.Dispose()
            d_con = Nothing
            d_cmd.Dispose()
            d_cmd = Nothing
            serializer = Nothing
            UProfile = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = DocAccess & " for user " & UserId & " and document " & DocId
        If Trim(errmsg) <> "" Then myeventlog.Error("CheckDocAccess :  Error: " & Trim(errmsg))
        myeventlog.Info("CheckDocAccess : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return access
        Return DocAccess
    End Function

    <WebMethod(Description:="Log a user into the DMS")> _
    Public Function UserLogin(ByVal ContactId As String, ByVal UserId As String, ByVal SessionId As String, _
        ByVal TrainerId As String, ByVal PartId As String, ByVal MtId As String, ByVal SubId As String, _
        ByVal Domain As String, ByVal Debug As String) As XmlDocument

        ' This function locates the specified item, and returns it to the calling system as a binary

        ' The input parameters are as follows:
        '
        '   ContactId  	- The users S_CONTACT.ROW_ID, Reqd.
        '   UserId  	- The users S_CONTACT.X_REGISTRATION_NUM, Reqd.
        '   SessionId  	- The users web site session id, Reqd.
        '   TrainerId  	- The users S_CONTACT.X_TRAINER_NUM, Opt.
        '   PartId  	- The users S_CONTACT.X_PART_ID, Opt.
        '   MtId  	    - The users S_CONTACT.REG_AS_EMP_ID, Opt.
        '   SubId	    - The users CX_SUB_CON.SUB_ID, Opt.
        '   Domain	    - The users subscription domain, Opt.
        '   Debug	    - The debug mode flag: "Y", "N" or "T" 

        ' The results are an XML document containing IDs

        ' web.config Parameters used:
        '   dms        	- connection string to DMS.dms database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging, temp As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' Read-Only Database declarations
        Dim RO_con As SqlConnection
        Dim RO_cmd As SqlCommand
        Dim RO_dr As SqlDataReader
        Dim RO_ConnS As String

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("ULODebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' DMS ID declarations
        Dim DMS_SESSION_ID, DMS_USER_ID, DMS_USER_AID As String
        Dim DMS_SUB_ID, DMS_DOMAIN_ID As String

        ' Misc. declarations

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DMS_SESSION_ID = ""
        DMS_USER_ID = ""
        DMS_USER_AID = ""
        DMS_SUB_ID = ""
        DMS_DOMAIN_ID = ""
        temp = ""

        ' ============================================
        ' Get and fix parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            ContactId = "21120611WE0"
            UserId = "RTO31123036OA"
            SessionId = "1111111111"
        Else
            If InStr(ContactId, "%") > 0 Then ContactId = Trim(HttpUtility.UrlDecode(ContactId))
            If InStr(ContactId, " ") > 0 Then ContactId = EncodeParamSpaces(ContactId)
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, " ") > 0 Then UserId = EncodeParamSpaces(UserId)
            If InStr(SessionId, "%") > 0 Then SessionId = Trim(HttpUtility.UrlDecode(SessionId))
            If InStr(SessionId, " ") > 0 Then SessionId = EncodeParamSpaces(SessionId)
            If InStr(TrainerId, "%") > 0 Then TrainerId = Trim(HttpUtility.UrlDecode(TrainerId))
            If InStr(TrainerId, " ") > 0 Then TrainerId = EncodeParamSpaces(TrainerId)
            If InStr(PartId, "%") > 0 Then PartId = Trim(HttpUtility.UrlDecode(PartId))
            If InStr(PartId, " ") > 0 Then PartId = EncodeParamSpaces(PartId)
            If InStr(MtId, "%") > 0 Then MtId = Trim(HttpUtility.UrlDecode(MtId))
            If InStr(MtId, " ") > 0 Then MtId = EncodeParamSpaces(MtId)
            If InStr(SubId, "%") > 0 Then SubId = Trim(HttpUtility.UrlDecode(SubId))
            If InStr(SubId, " ") > 0 Then SubId = EncodeParamSpaces(SubId)
            If InStr(Domain, "%") > 0 Then Domain = Trim(HttpUtility.UrlDecode(Domain))
            If InStr(Domain, " ") > 0 Then Domain = EncodeParamSpaces(Domain)
        End If

        ' ============================================
        ' Get system defaults
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            RO_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidbro").ConnectionString
            If RO_ConnS = "" Then RO_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;ApplicationIntent=ReadOnly;"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("UserLogin_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\UserLogin.log"
            Try
                log4net.GlobalContext.Properties("ULOLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  Debug: " & Debug)
                mydebuglog.Debug("  ContactId: " & ContactId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  SessionId: " & SessionId)
                mydebuglog.Debug("  TrainerId: " & TrainerId)
                mydebuglog.Debug("  PartId: " & PartId)
                mydebuglog.Debug("  MtId: " & MtId)
                mydebuglog.Debug("  SubId: " & SubId)
                mydebuglog.Debug("  Domain: " & Domain)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(ContactId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No contact id specified. "
            GoTo CloseOut2
        End If
        If Trim(UserId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No user id specified. "
            GoTo CloseOut2
        End If
        If Trim(SessionId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No session id specified. "
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open SQL Server database connection to DMS
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Or d_cmd Is Nothing Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Unable to open DMS connection. "
            GoTo CloseOut
        End If

        errmsg = OpenDBConnection(RO_ConnS, RO_con, RO_cmd)
        If errmsg <> "" Or RO_cmd Is Nothing Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Unable to open DMS ReadOnly connection. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Update Document Count
        ' Create an instance of the test class.
        Dim ad As New AsyncMain()

        ' Create the delegate.
        Dim caller As New AsynchUpdDMSDoc(AddressOf ad.UpdDMSDoc)

        ' Initiate the asynchronous call.
        Dim result As IAsyncResult = caller.BeginInvoke(ContactId, TrainerId, PartId, MtId, Debug, Nothing, Nothing)

        ' ============================================
        ' Locate Existing user session
        If results <> "Failure" Then

            ' Query DMS for existing user session
            SqlS = "SELECT S.row_id, U.row_id  " & _
                "FROM DMS.dbo.User_Sessions S  " & _
                "LEFT OUTER JOIN DMS.dbo.Users U ON U.ext_user_id='" & ContactId & "' " & _
                "WHERE S.user_id='" & UserId & "' AND S.session_key='" & SessionId & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locate Existing user session: " & SqlS)
            Try
                RO_cmd.CommandText = SqlS
                RO_dr = RO_cmd.ExecuteReader()
                If Not RO_dr Is Nothing Then
                    While RO_dr.Read()
                        Try
                            DMS_USER_ID = Trim(CheckDBNull(RO_dr(1), enumObjectType.StrType))
                            DMS_SESSION_ID = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType))
                            If Debug = "Y" Then
                                mydebuglog.Debug("  > Located User Id: " & DMS_USER_ID)
                                mydebuglog.Debug("  > Located Session Id: " & DMS_SESSION_ID)
                            End If
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & vbCrLf & "Error locating session id. " & ex.ToString
                            GoTo CloseOut
                        End Try
                    End While
                End If
                RO_dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error locating session id: " & oBug.ToString)
            End Try
        End If

        ' ============================================
        ' Perform Login if not logged in already        
        If DMS_SESSION_ID = "" Then

            ' Remove old session, if any
            SqlS = "DELETE FROM DMS.dbo.User_Sessions WHERE user_id='" & UserId & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Removing old sessions: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try

            ' Create User_Sessions record
            SqlS = "INSERT INTO DMS.dbo.User_Sessions(user_id, session_key, machine_id) " &
            "VALUES('" & UserId & "','" & SessionId & "','CM')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Create User_Sessions record: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Unable to create User_Sessions record. " & ex.ToString
                GoTo CloseOut
            End Try

            ' Locate created User_Sessions record
            SqlS = "SELECT S.row_id, U.row_id " &
                "FROM DMS.dbo.User_Sessions S " &
                "LEFT OUTER JOIN DMS.dbo.Users U ON U.ext_user_id='" & ContactId & "' " &
                "WHERE S.user_id='" & UserId & "' AND S.session_key='" & SessionId & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locate created User_Sessions record: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            DMS_SESSION_ID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                            DMS_USER_ID = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                            If Debug = "Y" Then mydebuglog.Debug("  > Located User_Sessions Id: " & DMS_SESSION_ID)
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & vbCrLf & "Unable to locate created User_Sessions record. " & ex.ToString
                            GoTo CloseOut
                        End Try
                    End While
                End If
                d_dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error Locate created User_Sessions record: " & oBug.ToString)
                results = "Failure"
            End Try

            ' Check session id
            If DMS_SESSION_ID = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Unable to verify or create user session"
                GoTo CloseOut
            End If
        End If

        ' ============================================
        ' Get user access id
        SqlS = "SELECT UA.row_id " & _
            "FROM DMS.dbo.User_Group_Access UA " & _
            "INNER JOIN DMS.dbo.Users U ON U.row_id=UA.access_id " & _
            "WHERE UA.type_id='U' AND U.ext_user_id='" & ContactId & "'"
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Getting users user access id: " & SqlS)
        Try
            RO_cmd.CommandText = SqlS
            RO_dr = RO_cmd.ExecuteReader()
            If Not RO_dr Is Nothing Then
                While RO_dr.Read()
                    Try
                        DMS_USER_AID = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType))
                        If Debug = "Y" Then mydebuglog.Debug("  > DMS_USER_AID: " & DMS_USER_AID)
                    Catch ex As Exception
                    End Try
                End While
            End If
            RO_dr.Close()
        Catch oBug As Exception
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error locating users user access id: " & oBug.ToString)
            results = "Failure"
        End Try

        ' ============================================
        ' Get subscription user access id
        If SubId <> "" Then
            SqlS = "SELECT UA.row_id " & _
                "FROM DMS.dbo.User_Group_Access UA " & _
                "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " & _
                "WHERE UA.type_id='G' AND G.name='" & SubId & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Getting subscription user access id: " & SqlS)
            Try
                RO_cmd.CommandText = SqlS
                RO_dr = RO_cmd.ExecuteReader()
                If Not RO_dr Is Nothing Then
                    While RO_dr.Read()
                        Try
                            DMS_SUB_ID = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType))
                            If Debug = "Y" Then mydebuglog.Debug("  > DMS_SUB_ID: " & DMS_SUB_ID)
                        Catch ex As Exception
                        End Try
                    End While
                End If
                RO_dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error locating subscription user access id: " & oBug.ToString)
                results = "Failure"
            End Try
        End If

        ' ============================================
        ' Get domain user access id
        If Domain <> "" Then
            SqlS = "SELECT UA.row_id " & _
                "FROM DMS.dbo.User_Group_Access UA " & _
                "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " & _
                "WHERE UA.type_id='G' AND G.name='" & Domain & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Getting domain user access id: " & SqlS)
            Try
                RO_cmd.CommandText = SqlS
                RO_dr = RO_cmd.ExecuteReader()
                If Not RO_dr Is Nothing Then
                    While RO_dr.Read()
                        Try
                            DMS_DOMAIN_ID = Trim(CheckDBNull(RO_dr(0), enumObjectType.StrType))
                            If Debug = "Y" Then mydebuglog.Debug("  > DMS_DOMAIN_ID: " & DMS_DOMAIN_ID)
                        Catch ex As Exception
                        End Try
                    End While
                End If
                RO_dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error locating domain user access id: " & oBug.ToString)
                results = "Failure"
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            d_dr = Nothing
            d_con.Dispose()
            d_con = Nothing
            d_cmd.Dispose()
            d_cmd = Nothing
            RO_dr = Nothing
            RO_con.Dispose()
            RO_con = Nothing
            RO_cmd.Dispose()
            RO_cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for UserId: " & UserId & ", ContactId: " & ContactId & ", SessionId: " & SessionId & ", in Domain: " & Domain
        If Trim(errmsg) <> "" Then myeventlog.Error("UserLogin :  Error: " & Trim(errmsg))
        myeventlog.Info("UserLogin : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug(vbCrLf & "Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return the data for this session
        '   <SESSION>
        '       <DMS_SESSION_ID>       - The User_Sessions.row_id for the user
        '       <DMS_USER_ID>          - The Users.row_id for the user
        '       <DMS_USER_AID>         - The User_Group_Access.row_id for the user
        '       <DMS_SUB_ID>           - The User_Group_Access.row_id for the user's subscription
        '       <DMS_DOMAIN_ID>        - The User_Group_Access.row_id for the user's domain        
        '   </SESSION>
        ' ============================================
        ' Return the address to the service consumer as an XML document
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("SESSION")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            If Debug <> "T" Then
                AddXMLChild(odoc, resultsRoot, "DMS_SESSION_ID", DMS_SESSION_ID)
                AddXMLChild(odoc, resultsRoot, "DMS_USER_ID", DMS_USER_ID)
                AddXMLChild(odoc, resultsRoot, "DMS_USER_AID", DMS_USER_AID)
                AddXMLChild(odoc, resultsRoot, "DMS_SUB_ID", DMS_SUB_ID)
                AddXMLChild(odoc, resultsRoot, "DMS_DOMAIN_ID", DMS_DOMAIN_ID)
            End If
            If Debug = "T" Then AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return item
        Return odoc
    End Function

    <WebMethod(Description:="Retrieves the specified asset")>
    Public Function GetDocument(ByVal RegId As String, ByVal UserId As String, ByVal Type As String, ByVal Debug As String, ByVal Asset As String, ByVal DocId As String) As Byte()

        ' This function locates the specified item, and returns it to the calling system 

        ' The input parameters are as follows:
        '
        '   RegId       - The CX_SESS_REG.ROW_ID of the attendee if the Type specified is a course
        '                   "Asset" or "Resource".  Otherwise this maps to the value in the field
        '                   "DMS.Document_Associations.fkey"
        '
        '   UserId      - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                   user if the Type specified is "Media" or "Resource".  Otherwise this can
        '                   be left blank.
        '
        '   Type        - A keyword to indicate the category of asset to retrieve. Currently
        '                   "Media" or "Resource".  This translates into the query used to 
        '                   locate the asset specified.  If any other value than this parameter
        '                   maps to the field "DMS.Association.name"
        '
        '   Debug       - "Y", "N" or "T"
        '   
        '   Asset       - The DMS.Documents.dfilename of the asset to be retrieved, or if "default.jpg",
        '                   the first associated item in the category "Images" will be returned.
        '   
        '   DocId       - The DMS.Documents.row_id of the document to be retrieved

        ' web.config Parameters used:
        '   hcidb           - connection string to hcidb1.siebeldb database
        '   dms               - connection string to DMS.dms database
        '   cache           - connection string to cache.sdf database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String
        Dim DecodedUserId, ValidatedUserId As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer
        Dim TypeTrans As String

        ' Cache database declarations
        Dim c_ConnS As String
        Dim CacheHit As Integer
        Dim dAsset, dCrseId, dFileName, minio_flg As String
        Dim LastUpd As DateTime
        Dim d_last_updated As DateTime

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String
        Dim dms_cache_age As String

        ' Logging declarations
        Dim myeventlog = log4net.LogManager.GetLogger("EventLog")
        Dim mydebuglog = log4net.LogManager.GetLogger("GMDebugLog")
        Dim logfile, temp As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' File handling declarations
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim d_dsize, d_doc_id, d_verid, SaveDest As String
        Dim dLastUpd As DateTime
        Dim CRSE_ID As String
        Dim killcount As Double

        Dim filecache As ObjectCache = MemoryCache.Default
        Dim fileContents(1000) As Byte
        Dim cacheItemName As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        ValidatedUserId = ""
        CRSE_ID = ""
        d_dsize = ""
        BinaryFile = ""
        c_ConnS = ""
        dAsset = ""
        d_doc_id = ""
        dCrseId = ""
        dFileName = ""
        TypeTrans = ""
        'Debug = "Y"
        killcount = 0
        temp = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If RegId = "" And DocId = "" And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
            Asset = "2.A.5.c.3.swf"
            Type = "Asset"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
            If InStr(Asset, "%") > 0 Then Asset = Trim(HttpUtility.UrlDecode(Asset))
        End If
        If Trim(Asset) = "" And DocId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No item specified. "
            GoTo CloseOut2
        End If

        ' 07-09-2015;Ren Hou; Added to remove special characters;
        Dim RegExStr As String = "[\\/:*?""<>|]"  'For eliminating Characters: \ / : * ? "  |
        Asset = Regex.Replace(Asset, RegExStr, "")

        ' Translate category name if the type is "Asset" or "Resource", otherwise assume no category or default category
        Select Case Trim(Type.ToLower)
            Case "asset"
                TypeTrans = "8"
                Asset = Replace(Asset, "media_", "")
            Case "resource"
                TypeTrans = "10"
            Case Else
                TypeTrans = "8"
                Asset = Replace(Asset, "media_", "")
        End Select

        Type = Trim(Type.ToLower)
        If Type = "media" Then Type = "asset"

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidbro").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=DotNet;pwd=FS6vdEje,#D7;database=siebeldb;ApplicationIntent=ReadOnly"
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIBDB;uid=DotNet;pwd=FS6vdEje,#D7;database=DMS;ApplicationIntent=ReadOnly"
            dms_cache_age = Trim(System.Configuration.ConfigurationManager.AppSettings("dmscacheage"))
            If dms_cache_age = "" Or Not IsNumeric(dms_cache_age) Then dms_cache_age = "30"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetDocument_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
            AccessKey = System.Configuration.ConfigurationManager.AppSettings("minio-key")
            If AccessKey = "" Then AccessKey = "dms"
            AccessSecret = System.Configuration.ConfigurationManager.AppSettings("minio-secret")
            If AccessSecret <> "" Then AccessSecret = System.Web.HttpUtility.HtmlDecode(AccessSecret)
            If AccessSecret = "" Then AccessSecret = "TptbjrNTVQDRYFJzNmw27BV5"
            AccessRegion = System.Configuration.ConfigurationManager.AppSettings("minio-region")
            If AccessRegion = "" Then AccessRegion = "us-east"
            AccessBucket = System.Configuration.ConfigurationManager.AppSettings("minio-bucket")
            If AccessBucket = "" Then AccessBucket = "dms"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetDocument.log"
            Try
                log4net.GlobalContext.Properties("GMLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Type: " & Type)
                mydebuglog.Debug("  Asset: " & Asset)
                mydebuglog.Debug("  TypeTrans: " & TypeTrans)
                mydebuglog.Debug("  AccessBucket: " & AccessBucket)
                mydebuglog.Debug("  AccessRegion: " & AccessRegion)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  Appsetting dms_cache_age: " & dms_cache_age & vbCrLf)
            End If
        End If

        ' ============================================
        ' Open SQL Server database connections
        errmsg = OpenDBConnection(ConnS, con, cmd)          ' hcidb1
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)    ' dms
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Validate identity if needed
        If Trim(Type.ToLower) = "asset" Or Type = "resource" Then
            If Not cmd Is Nothing Then
                ' -----
                ' Query registration
                Try
                    SqlS = "SELECT R.CRSE_ID, C.X_REGISTRATION_NUM " &
                    "FROM siebeldb.dbo.CX_SESS_REG R " &
                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " &
                    "WHERE R.ROW_ID='" & RegId & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                                CRSE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                ValidatedUserId = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "The record was not found." & vbCrLf
                        dr.Close()
                        results = "Failure"
                    End If
                    dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                    results = "Failure"
                End Try
                If Debug = "Y" Then mydebuglog.Debug("   ... CRSE_ID: " & CRSE_ID)

                ' -----
                ' Verify the user
                If ValidatedUserId <> DecodedUserId Then
                    results = "Failure"
                    CRSE_ID = ""
                    errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                    GoTo CloseOut
                End If
            Else
                results = "Failure"
                GoTo CloseOut
            End If
        End If

        ' ============================================
        ' Create output directory for asset caching
        Select Case Trim(Type.ToLower)
            Case "asset"
                SaveDest = mypath & "course_temp\" & CRSE_ID
            Case "resource"
                SaveDest = mypath & "course_temp\" & CRSE_ID
            Case Else
                If Type.ToLower <> "" Then
                    SaveDest = mypath & Replace(Trim(Type.ToLower), " ", "_") & "_temp\" & RegId
                Else
                    SaveDest = mypath & "document_temp\" & RegId
                End If
        End Select
        Try
            Directory.CreateDirectory(SaveDest)
        Catch
        End Try
        If Debug = "Y" Then mydebuglog.Debug("   Asset caching: " & SaveDest & vbCrLf)

        ' ============================================
        ' Get the name of the asset if necessary
        If Debug = "Y" Then mydebuglog.Debug("   Looking for: " & LCase(Trim(System.IO.Path.GetFileNameWithoutExtension(Asset))) & vbCrLf)
        If LCase(Trim(System.IO.Path.GetFileNameWithoutExtension(Asset))) = "default" Or DocId <> "" Then
            If Not d_cmd Is Nothing Then
                ' Query DMS
                If DocId <> "" Then
                    SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd " &
                    "FROM DMS.dbo.Documents D  " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id  " &
                    "WHERE D.row_id IS NOT NULL AND D.deleted IS NULL AND LOWER(A.name)='" & Type & "' AND C.name='Images' " &
                    "AND D.row_id=" & DocId & " " &
                    "ORDER BY D.last_upd DESC"
                Else
                    SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd " &
                    "FROM DMS.dbo.Documents D  " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id  " &
                    "WHERE D.row_id IS NOT NULL AND D.deleted IS NULL AND LOWER(A.name)='" & Type & "' AND C.name='Images' " &
                    "AND DA.fkey='" & RegId & "' " &
                    "ORDER BY D.last_upd DESC"
                End If
                If Debug = "Y" Then mydebuglog.Debug("  Get default item for type specified: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                Asset = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                                d_last_updated = CheckDBNull(d_dr(2), enumObjectType.DteType)
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_doc_id=" & d_doc_id & "  Asset=" & Asset)
                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error getting default item. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "Error getting default item." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting default item: " & oBug.ToString)
                    results = "Failure"
                End Try
            End If
        End If

        ' -----
        ' Locate last updated
        If Asset = "" And DocId <> "" Then
            If Not d_cmd Is Nothing Then
                ' Query DMS
                SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd " &
                    "FROM DMS.dbo.Documents D  " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id  " &
                    "WHERE D.row_id=" & DocId & " AND D.deleted IS NULL " &
                    "ORDER BY D.last_upd DESC"
                If Debug = "Y" Then mydebuglog.Debug("  Get document information for document id specified: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                Asset = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                                d_last_updated = CheckDBNull(d_dr(2), enumObjectType.DteType)
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_doc_id=" & d_doc_id & "  Asset=" & Asset)

                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error getting default item. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "Error getting default item." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting default item: " & oBug.ToString)
                    results = "Failure"
                End Try
            End If
        End If

        ' -----
        ' Generate cache filename
        BinaryFile = SaveDest & "\" & Asset
        'BinaryFile = BinaryFile.Replace(mypath, ".\")
        If Debug = "Y" Then mydebuglog.Debug("  Cache filename: " & BinaryFile & vbCrLf)

        ' ============================================
        ' Get DMS record containing asset 
        '  If the cache was hit, make sure the entry is current, otherwise restore it
        If Not d_cmd Is Nothing Then
            ' -----
            ' Get information from versions record
            SqlS = ""
            If d_doc_id <> "" Then
                SqlS = "SELECT TOP 1 v.dsize, v.row_id, v.last_upd, v.minio_flg, v.dimage " &
                    "FROM DMS.dbo.Documents d " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                    "WHERE d.row_id=" & d_doc_id
            Else
                Select Case Type
                    Case "asset"
                        SqlS = "SELECT TOP 1 v.dsize, v.row_id, v.last_upd, d.row_id, v.minio_flg, v.dimage " &
                            "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id " &
                            "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=" &
                            TypeTrans & " and (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                    Case "resource"
                        SqlS = "SELECT TOP 1 v.dsize, v.row_id, v.last_upd, d.row_id, v.minio_flg, v.dimage " &
                            "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id " &
                            "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=" &
                            TypeTrans & " and (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                    Case Else
                        SqlS = "SELECT TOP 1 v.dsize, v.row_id, v.last_upd, d.row_id, v.minio_flg, v.dimage " &
                            "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da on da.doc_id=d.row_id  " &
                            "LEFT OUTER JOIN DMS.dbo.Associations a on a.row_id=da.association_id  " &
                            "WHERE da.fkey='" & CRSE_ID & "' and lower(a.name)='" & Trim(Type.ToLower) & "' and " &
                            "(d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                End Select
            End If

            'Load content of cached object
            Dim last_upt As Date = Today.AddYears(-50)
            ' Check to see if the document is in the in-memory cache
            cacheItemName = BinaryFile
            Dim docNotInDB As Boolean = 0
            If filecache(cacheItemName) Is Nothing Then
                fileContents = Nothing
            Else
                Dim tmpObj As Object
                tmpObj = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
                ReDim fileContents(tmpObj.Length)
                fileContents = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
                CacheHit = 1
                mydebuglog.Debug(vbCrLf & " Retrieved Cached object " & cacheItemName & " !")
            End If

            If Not IsNothing(filecache(cacheItemName)) Or IsNothing(fileContents) Then
                'Check if the cached item need to be renewed;
                Try
                    Dim upt_SqlStr As String = ""
                    upt_SqlStr = "Select TOP 1 v.last_upd " & SqlS.Substring(SqlS.IndexOf("FROM "))
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check last updated date for document: " & vbCrLf & upt_SqlStr)
                    d_cmd.CommandText = upt_SqlStr
                    last_upt = d_cmd.ExecuteScalar()
                Catch ex As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
                    results = "Failure"
                End Try

                If last_upt = Today.AddYears(-50) Then  'document no longer exists in the database
                    docNotInDB = True
                    filecache.Remove(cacheItemName)
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   " & Asset & " is not found in database! " & vbCrLf)
                    'results = "Failure"
                Else
                    If Not IsNothing(fileContents) Then
                        If last_upt > TryCast(filecache(cacheItemName), HciDMSDocument).UpdateDate Then
                            'Remove if the update_date on the cache is before the last updated date on DB record.
                            filecache.Remove(cacheItemName)
                            mydebuglog.Debug(vbCrLf & "  Cached object " & cacheItemName & " expired.")
                        End If
                    End If
                End If

                'Load content of cached object
                If filecache(cacheItemName) Is Nothing Then
                    fileContents = Nothing
                Else
                    Dim tmpObj As Object
                    tmpObj = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
                    ReDim fileContents(tmpObj.Length)
                    fileContents = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
                    CacheHit = 1
                    mydebuglog.Debug(vbCrLf & " Retrieved Cached object " & cacheItemName & " !")
                End If

                If (IsNothing(fileContents) Or Debug = "R") And docNotInDB = False Then
                    If Debug = "Y" Then mydebuglog.Debug("  Checking item with DMS query: " & SqlS)
                    Try
                        d_cmd.CommandText = SqlS
                        d_dr = d_cmd.ExecuteReader()
                        If Not d_dr Is Nothing Then
                            While d_dr.Read()
                                Try
                                    d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                    d_verid = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                                    dLastUpd = d_dr(2)
                                    d_doc_id = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                                    minio_flg = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                    If Debug = "Y" Then mydebuglog.Debug("  > Record found on query id " & d_verid & ":  d_dsize=" & d_dsize & ", minio_flg=" & minio_flg & ", dLastUpd=" & Format(dLastUpd) & ", cLastUpd=" & Convert.ToString(LastUpd) & ", CacheHit=" & Format(CacheHit))

                                    If minio_flg = "Y" Then
                                        ' Get binary from Minio
                                        If Debug = "Y" Then mydebuglog.Debug("  Getting binary from Minio")
                                        Dim MConfig As AmazonS3Config = New AmazonS3Config()
                                        'MConfig.RegionEndpoint = RegionEndpoint.USEast1
                                        MConfig.ServiceURL = "https://192.168.5.134"
                                        MConfig.ForcePathStyle = True
                                        MConfig.EndpointDiscoveryEnabled = False
                                        Dim Minio As AmazonS3Client = New AmazonS3Client(AccessKey, AccessSecret, MConfig)
                                        ServicePointManager.ServerCertificateValidationCallback = AddressOf sslhttps.AcceptAllCertifications
                                        Try
                                            Dim mobj2 = Minio.GetObject(AccessBucket, d_doc_id & "-" & d_verid)
                                            retval = mobj2.ContentLength
                                            If retval > 0 Then
                                                ReDim outbyte(Val(retval))
                                                Dim intval As Integer
                                                For i = 0 To retval
                                                    intval = mobj2.ResponseStream.ReadByte()
                                                    If intval < 255 And intval > 0 Then
                                                        outbyte(i) = intval
                                                    End If
                                                    If intval = 255 Then outbyte(i) = 255
                                                    If intval < 0 Then
                                                        outbyte(i) = 0
                                                        If Debug = "Y" Then mydebuglog.Debug(" .. " & i.ToString & "   intval: " & intval.ToString)
                                                    End If
                                                Next
                                                If Debug = "Y" Then
                                                    BinaryFile = SaveDest & Asset
                                                    If System.IO.File.Exists(BinaryFile) Then System.IO.File.Delete(BinaryFile)
                                                    File.WriteAllBytes(BinaryFile, outbyte)
                                                End If
                                            End If
                                            mobj2 = Nothing
                                        Catch ex2 As Exception
                                            results = "Failure"
                                            errmsg = errmsg & "Error getting object. " & ex2.ToString & vbCrLf
                                            GoTo CloseOut
                                        End Try

                                        Try
                                            Minio = Nothing
                                        Catch ex As Exception
                                            errmsg = errmsg & "Error closing Minio: " & ex.Message & vbCrLf
                                        End Try
                                    Else
                                        ' Get binary from document_versions
                                        If Debug = "Y" Then mydebuglog.Debug("  Getting binary from Document_Versions")
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        Try
                                            retval = d_dr.GetBytes(5, 0, outbyte, 0, d_dsize)
                                        Catch ex As Exception
                                            results = "Failure"
                                            errmsg = errmsg & "Error getting item. " & ex.ToString & vbCrLf
                                            GoTo CloseOut
                                        End Try
                                    End If

                                Catch ex As Exception
                                    results = "Failure"
                                    errmsg = errmsg & "Error getting asset. " & ex.ToString & vbCrLf
                                    GoTo CloseOut
                                End Try
                            End While
                        Else
                            errmsg = errmsg & "Error getting asset." & vbCrLf
                            d_dr.Close()
                            results = "Failure"
                        End If
                        d_dr.Close()
                    Catch oBug As Exception
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting document: " & oBug.ToString)
                        results = "Failure"
                    End Try

                    If Debug = "Y" Then
                        mydebuglog.Debug("Document reported size: " & d_dsize & ", Object reported size: " & Str(retval) & vbCrLf)
                    End If

                    '***** Set cache object; 2/28/17; Ren Hou;  ****
                    Dim policy As New CacheItemPolicy()
                    policy.SlidingExpiration = TimeSpan.FromDays(CDbl(dms_cache_age))
                    filecache.Set(cacheItemName, New HciDMSDocument(dLastUpd, outbyte), policy)
                    ReDim fileContents(outbyte.Length)
                    fileContents = outbyte
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Caching DMS doc to key: " & cacheItemName)
                    '****  ****
                End If
            End If
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
            d_dr = Nothing
            d_con.Dispose()
            d_con = Nothing
            d_cmd.Dispose()
            d_cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetDocument : Error: " & Trim(errmsg))
        If CacheHit = 1 Then
            If Debug <> "T" Then myeventlog.Info("GetDocument : Results: " & results & " for CACHED " & Type & " file " & Asset & ", by UserId # " & DecodedUserId & " with RegId " & RegId & ", Doc Id: " & d_doc_id)
        Else
            If Debug <> "T" Then myeventlog.Info("GetDocument : Results: " & results & " for " & Type & " file " & Asset & ", by UserId # " & DecodedUserId & " with RegId " & RegId & ", Doc Id: " & d_doc_id)
        End If
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If CacheHit = 1 Then
                    mydebuglog.Debug("Results: " & results & " for CACHED " & Type & " file " & Asset & ", by UserId # " & DecodedUserId & " with RegId " & RegId & " at " & Now.ToString)
                Else
                    mydebuglog.Debug("Results: " & results & " for " & Type & " file " & Asset & ", by UserId # " & DecodedUserId & " with RegId " & RegId & " at " & Now.ToString)
                End If
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return asset
        Try
            'bfs = File.Open(BinaryFile, FileMode.Open, FileAccess.Read)
            'Dim lngLen As Long = bfs.Length
            'ReDim outbyte(CInt(lngLen - 1))
            'bfs.Read(outbyte, 0, CInt(lngLen))
            Return fileContents
        Catch exp As Exception
            Return Nothing
        Finally
            'bfs.Close()
            'bfs = Nothing
            outbyte = Nothing
            fileContents = Nothing
        End Try
    End Function

    <WebMethod(Description:="Verify the specified document")>
    Public Function VerifyDocument(ByVal Asset As String, ByVal DocId As String, ByVal Debug As String) As String

        ' This function locates the specified document, and confirms it exists

        ' The input parameters are as follows:
        '   Asset       - The DMS.Documents.dfilename of the asset to be retrieved, or if "default.jpg",
        '                   the first associated item in the category "Images" will be returned.
        '   
        '   DocId       - The DMS.Documents.row_id of the document to be retrieved
        '
        '   Debug       - "Y", "N" or "T"

        ' web.config Parameters used:
        '   hcidb           - connection string to hcidb1.siebeldb database
        '   dms               - connection string to DMS.dms database
        '   cache           - connection string to cache.sdf database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String
        Dim DecodedUserId, ValidatedUserId As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer
        Dim TypeTrans As String

        ' Cache database declarations
        Dim c_ConnS As String
        Dim dAsset, dCrseId, dFileName, minio_flg, d_ext As String
        Dim d_last_updated As DateTime

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim myeventlog = log4net.LogManager.GetLogger("EventLog")
        Dim mydebuglog = log4net.LogManager.GetLogger("GMDebugLog")
        Dim logfile, temp As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' File handling declarations
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim d_dsize, d_doc_id, d_verid, vd_dsize As String
        Dim dLastUpd As DateTime
        Dim CRSE_ID As String
        Dim killcount As Double

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        ValidatedUserId = ""
        CRSE_ID = ""
        d_dsize = ""
        BinaryFile = ""
        c_ConnS = ""
        dAsset = ""
        d_doc_id = ""
        dCrseId = ""
        dFileName = ""
        TypeTrans = ""
        'Debug = "Y"
        killcount = 0
        temp = ""
        d_verid = ""
        minio_flg = "N"

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If Asset = "" And DocId = "" And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            Asset = "2.A.5.c.3.swf"
        Else
            If InStr(Asset, "%") > 0 Then Asset = Trim(HttpUtility.UrlDecode(Asset))
            Dim RegExStr As String = "[\\/:*?""<>|]"  'For eliminating Characters: \ / : * ? "  |
            Asset = Regex.Replace(Asset, RegExStr, "")
        End If
        If Trim(Asset) = "" And DocId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No item specified. "
            GoTo CloseOut2
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidbro").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb;ApplicationIntent=ReadOnly"
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIBDB;uid=DMS;pwd=5241200;database=DMS;ApplicationIntent=ReadOnly"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("VerifyDocument_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
            AccessKey = System.Configuration.ConfigurationManager.AppSettings("minio-key")
            If AccessKey = "" Then AccessKey = "dms"
            AccessSecret = System.Configuration.ConfigurationManager.AppSettings("minio-secret")
            If AccessSecret <> "" Then AccessSecret = System.Web.HttpUtility.HtmlDecode(AccessSecret)
            If AccessSecret = "" Then AccessSecret = "SfI@aUE$?=&KcAOI?C5NU|-c*Oec7ZPJ"
            AccessRegion = System.Configuration.ConfigurationManager.AppSettings("minio-region")
            If AccessRegion = "" Then AccessRegion = "us-east"
            AccessBucket = System.Configuration.ConfigurationManager.AppSettings("minio-bucket")
            If AccessBucket = "" Then AccessBucket = "default"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\VerifyDocument.log"
            Try
                log4net.GlobalContext.Properties("GMLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  Asset: " & Asset)
                mydebuglog.Debug("  AccessBucket: " & AccessBucket)
                mydebuglog.Debug("  AccessRegion: " & AccessRegion & vbCrLf)
            End If
        End If

        ' ============================================
        ' Open SQL Server database connections
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)    ' dms
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Get the asset 
        If Debug = "Y" Then mydebuglog.Debug("   Looking for: " & LCase(Trim(System.IO.Path.GetFileNameWithoutExtension(Asset))) & vbCrLf)
        If Asset <> "" And DocId = "" Then
            If Not d_cmd Is Nothing Then
                SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd, V.dsize, V.row_id, V.last_upd, V.minio_flg, DT.extension " &
                "FROM DMS.dbo.Documents D  " &
                "LEFT OUTER JOIN DMS.dbo.Document_Types DT on D.data_type_id=DT.row_id " &
                "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=d.last_version_id " &
                "WHERE D.row_id IS NOT NULL AND D.deleted IS NULL AND D.dfilename='" & Asset & "' " &
                "ORDER BY D.last_upd DESC"
                If Debug = "Y" Then mydebuglog.Debug("  Get document information for asset specified: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                Asset = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                DocId = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                                d_last_updated = CheckDBNull(d_dr(2), enumObjectType.DteType)
                                d_dsize = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                                d_verid = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                dLastUpd = CheckDBNull(d_dr(5), enumObjectType.DteType)
                                minio_flg = Trim(CheckDBNull(d_dr(6), enumObjectType.StrType))
                                d_ext = Trim(CheckDBNull(d_dr(7), enumObjectType.StrType))
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_doc_id=" & d_doc_id & ",  Asset=" & Asset & ",  d_ext=" & d_ext & ",  d_verid=" & d_verid & ",  minio_flg=" & minio_flg & ", dLastUpd=" & dLastUpd.ToString & vbCrLf)
                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error getting default item. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "Error getting default item." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting default item: " & oBug.ToString)
                    results = "Failure"
                End Try
            End If
        End If

        ' -----
        ' Locate last updated
        If Asset = "" And DocId <> "" Then
            If Not d_cmd Is Nothing Then
                ' Query DMS
                SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd, V.dsize, V.row_id, V.last_upd, V.minio_flg, DT.extension " &
                    "FROM DMS.dbo.Documents D " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Types DT on D.data_type_id=DT.row_id " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id " &
                    "WHERE D.row_id=" & DocId & " AND D.row_id IS NOT NULL AND D.deleted IS NULL " &
                    "ORDER BY D.last_upd DESC"
                If Debug = "Y" Then mydebuglog.Debug("  Get document information for document id specified: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                Asset = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                DocId = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                                d_last_updated = CheckDBNull(d_dr(2), enumObjectType.DteType)
                                d_dsize = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                                d_verid = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                dLastUpd = CheckDBNull(d_dr(5), enumObjectType.DteType)
                                minio_flg = Trim(CheckDBNull(d_dr(6), enumObjectType.StrType))
                                d_ext = Trim(CheckDBNull(d_dr(7), enumObjectType.StrType))
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_doc_id=" & d_doc_id & ",  Asset=" & Asset & ",  d_ext=" & d_ext & ",  d_verid=" & d_verid & ",  minio_flg=" & minio_flg & ", dLastUpd=" & dLastUpd.ToString & vbCrLf)

                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error getting default item. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "Error getting default item." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting default item: " & oBug.ToString)
                    results = "Failure"
                End Try
            End If
        End If

        ' ============================================
        ' Get asset from object store 
        '  If the cache was hit, make sure the entry is current, otherwise restore it
        If Not d_cmd Is Nothing And d_verid <> "" Then

            If minio_flg = "Y" Then
                ' Get binary from Minio
                Dim MConfig As AmazonS3Config = New AmazonS3Config()
                'MConfig.RegionEndpoint = RegionEndpoint.USEast1
                MConfig.ServiceURL = "https://192.168.5.134"
                MConfig.ForcePathStyle = True
                MConfig.EndpointDiscoveryEnabled = False
                Dim Minio As AmazonS3Client = New AmazonS3Client(AccessKey, AccessSecret, MConfig)
                ServicePointManager.ServerCertificateValidationCallback = AddressOf sslhttps.AcceptAllCertifications
                Try
                    If Debug = "Y" Then mydebuglog.Debug("  Retrieving binary from Minio")
                    Dim mobj2 = Minio.GetObject(AccessBucket, DocId & "-" & d_verid)
                    retval = mobj2.ContentLength
                    If retval = d_dsize Then results = "Success"
                    mobj2 = Nothing
                Catch ex2 As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error getting object. " & ex2.ToString & vbCrLf
                    GoTo CloseOut
                End Try

                Try
                    Minio = Nothing
                Catch ex As Exception
                    errmsg = errmsg & "Error closing Minio: " & ex.Message & vbCrLf
                End Try
                If Debug = "Y" Then
                    mydebuglog.Debug("  > Found image in Minio.   results: " & results)
                    mydebuglog.Debug("                            Size: " & d_dsize)
                    mydebuglog.Debug("                            Bytes retrieved: " & retval.ToString & vbCrLf)
                End If

            Else
                ' Get image from document_versions
                SqlS = "SELECT TOP 1 V.dimage, datalength(V.dimage) " &
                    "FROM DMS.dbo.Document_Versions V " &
                    "WHERE V.row_id=" & d_verid
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            vd_dsize = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                            If vd_dsize <> "" And Val(vd_dsize) > 0 Then
                                If Debug = "Y" Then mydebuglog.Debug("  Retrieving binary from Document_Versions")
                                ReDim outbyte(Val(vd_dsize) - 1)
                                startIndex = 0
                                Try
                                    retval = d_dr.GetBytes(0, 0, outbyte, 0, vd_dsize)
                                Catch ex As Exception
                                    results = "Failure"
                                    errmsg = errmsg & "Error getting item. " & ex.ToString & vbCrLf
                                    GoTo CloseOut
                                End Try
                            End If
                            If retval = vd_dsize Then results = "Success"
                            If Debug = "Y" Then
                                mydebuglog.Debug("  > Found image on query.   results: " & results)
                                mydebuglog.Debug("                            Size: " & vd_dsize)
                                mydebuglog.Debug("                            Bytes retrieved: " & retval.ToString & vbCrLf)
                            End If
                        End While
                    End If
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error getting image. " & ex.ToString & vbCrLf
                    GoTo CloseOut
                End Try
                d_dr.Close()
            End If
        Else
            results = "Failure"
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            d_dr = Nothing
            d_con.Dispose()
            d_con = Nothing
            d_cmd.Dispose()
            d_cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("VerifyDocument : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("VerifyDocument : Results: " & results & " for file " & Asset & " with document id " & DocId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & results & " for file " & Asset & " with document id " & DocId & " at " & Now.ToString)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return asset
        Try
            Return results
        Catch exp As Exception
            Return Nothing
        Finally
            outbyte = Nothing
        End Try
    End Function

    <WebMethod(Description:="Update the provided document into the DMS")>
    Public Function UpdateDoc(ByVal DImage As String, ByVal DURL As String, ByVal DocId As String, ByVal ReportId As String, ByVal UpdType As String, ByVal Debug As String) As String

        ' This function locates the specified item and updates it as applicable

        ' The input parameters are as follows:
        '
        '   DImage	    - The Base64 encoded binary of the image to be stored (req.)
        '   DURL        - The URL to the binary of the image to be stored (req.)
        '   DocId   	- The document id of "DMS.Documents.row_id" of the document to be stored (req)
        '   ReportId    - The CX_REP_ENT_SCHED.ROW_ID of a report document to be stored (opt.)
        '   UpdType	    - "N"ew version or "R"eplace existing version
        '   Debug	    - The debug mode flag: "Y", "N" or "T" 

        ' web.config Parameters used:
        '   dms        	    - connection string to DMS.dms database

        ' Variables
        Dim results, temp As String
        Dim mypath, errmsg, logging As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' HCIDB Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim ConnS As String

        ' Logging declarations
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("UDDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service
        Dim BasicService As New basic.com.certegrity.cloudsvc.Service
        Dim DmsService As New local.hq.dms.Service

        ' Local Cache declarations
        Dim DMSCache As New CachingWrapper.LocalCache

        ' File handling declarations
        Dim bfs As FileStream
        Dim bw As BinaryWriter
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim VerifiedSize As Long
        Dim rsize As String

        ' Document declarations
        Dim d_dsize, SaveDest As String
        Dim d_ext As String
        Dim DummyKey As String
        Dim DecodedDocId, VerifiedDocId As String
        Dim tempfile, AddlDesc, minio_flg As String
        Dim data_type_id As String
        Dim DomainGroupId, SubGroupId, UserDMSId, Description, FileExt, DFileName, ItemName As String
        Dim Trainer, PartId As String
        Dim DocVersionId As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        d_ext = ""
        d_dsize = ""
        DomainGroupId = ""
        SubGroupId = ""
        UserDMSId = ""
        data_type_id = ""
        BinaryFile = ""
        d_ConnS = ""
        DecodedDocId = ""
        VerifiedDocId = ""
        tempfile = ""
        DummyKey = ""
        VerifiedSize = 0
        Trainer = ""
        PartId = ""
        DocVersionId = ""
        AddlDesc = ""
        Description = ""
        FileExt = ""
        DFileName = ""
        ItemName = ""
        minio_flg = "N"

        ' Fix parameters
        Debug = UCase(Left(Debug, 1))
        If UpdType <> "R" Then UpdType = "N"

        ' ============================================
        ' Get system defaults
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("UpdateDoc_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
            AccessKey = System.Configuration.ConfigurationManager.AppSettings("minio-key")
            If AccessKey = "" Then AccessKey = "dms"
            AccessSecret = System.Configuration.ConfigurationManager.AppSettings("minio-secret")
            If AccessSecret <> "" Then AccessSecret = System.Web.HttpUtility.HtmlDecode(AccessSecret)
            If AccessSecret = "" Then AccessSecret = "SfI@aUE$?=&KcAOI?C5NU|-c*Oec7ZPJ"
            AccessRegion = System.Configuration.ConfigurationManager.AppSettings("minio-region")
            If AccessRegion = "" Then AccessRegion = "us-east"
            AccessBucket = System.Configuration.ConfigurationManager.AppSettings("minio-bucket")
            If AccessBucket = "" Then AccessBucket = "default"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\UpdateDoc.log"
            Try
                log4net.GlobalContext.Properties("UDLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug(" -Document Data-")
                mydebuglog.Debug("  DURL: " & DURL)
                mydebuglog.Debug("  DocId: " & DocId)
                mydebuglog.Debug("  ReportId: " & ReportId)
                mydebuglog.Debug("  UpdType: " & UpdType)
                mydebuglog.Debug("  AccessBucket: " & AccessBucket)
                mydebuglog.Debug("  AccessRegion: " & AccessRegion & vbCrLf)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(DImage) = "" And ReportId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No document specified. "
            GoTo CloseOut2
        End If
        If DocId.Trim <> "" Then
            If Not IsNumeric(DocId) Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Incorrect document id specified. "
                GoTo CloseOut2
            End If
        End If

        ' ============================================
        ' Open SQL Server database connection to DMS
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
        If errmsg <> "" Or d_cmd Is Nothing Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Unable to open DMS connection. "
            GoTo CloseOut
        End If

        ' ============================================
        ' If ReportId is provided, then retrieve from this database and store in a buffer
        '   Retrieve other values from the report schedule record
        If ReportId <> "" Then

            ' If dealing with reports, open HCIDB connection
            errmsg = OpenDBConnection(ConnS, con, cmd)
            cmd.CommandTimeout = 120            ' Set timeout to 2 mins

            If errmsg <> "" Or d_cmd Is Nothing Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Unable to open HCIDB connection. "
                GoTo CloseOut
            End If

            ' Retrieve report information
            rsize = ""
            SqlS = "SELECT S.DSIZE, S.DIMAGE, R.DESCRIPTION+(SELECT CASE WHEN S.ADDL_DESC IS NOT NULL THEN S.ADDL_DESC ELSE '' END), S.FORMAT, " &
            "(SELECT CASE WHEN S.XFER_ID IS NULL THEN S.ROW_ID ELSE S.XFER_ID END)+'.'+LOWER(S.FORMAT) AS DFILENAME, " &
            "R.NAME AS ITEMNAME " &
            "FROM siebeldb.dbo.CX_REP_ENT_SCHED S " &
            "LEFT OUTER JOIN siebeldb.dbo.CX_REPORT_ENT E ON E.ROW_ID=S.ENT_ID " &
            "LEFT OUTER JOIN siebeldb.dbo.CX_REPORTS R ON R.ROW_ID=E.REPORT_ID " &
            "WHERE S.ROW_ID='" & ReportId & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Retrieving report: " & vbCrLf & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            rsize = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            temp = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            If UCase(Trim(temp)) <> UCase(Trim(Description)) Then Description = Description & temp
                            If Debug = "Y" Then mydebuglog.Debug("  >Description: " & Description)
                            If FileExt = "" Then FileExt = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            DFileName = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            If ItemName = "" Then ItemName = Trim(CheckDBNull(dr(5), enumObjectType.StrType))

                            ' Get binary and attach to the object outbyte if found, not cached or updated recently
                            '   retval will be "0" if this is not the case
                            If rsize <> "" Then
                                ReDim outbyte(Val(rsize) - 1)
                                startIndex = 0
                                retval = dr.GetBytes(1, 0, outbyte, 0, rsize)
                            End If
                        Catch obug As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting report - read failure. " & obug.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Error getting report - datarecord failure." & vbCrLf
                    results = "Failure"
                End If

            Catch ex As Exception
                errmsg = errmsg & "Error getting report - command failure. " & vbCrLf & ex.Message
            End Try
            Try
                dr.Close()
            Catch ex As Exception
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "Report Information-")
                mydebuglog.Debug("  Size: " & rsize)
                mydebuglog.Debug("  ItemName: " & ItemName)
                mydebuglog.Debug("  DFileName: " & DFileName)
                mydebuglog.Debug("  Description: " & Description)
                mydebuglog.Debug("  FileExt: " & FileExt)
                mydebuglog.Debug("  ItemName: " & ItemName & vbCrLf)
            End If

            ' If unable to locate the report, then error out
            If retval = 0 Or FileExt = "" Then
                results = "Failure"
                errmsg = errmsg & "Error getting report." & vbCrLf
                GoTo CloseOut
            End If
        End If

        ' ============================================
        ' Create output directory for temp file caching if needed
        SaveDest = mypath & "work_dir\" & FileExt
        Try
            Directory.CreateDirectory(SaveDest)
        Catch
        End Try

        ' ============================================
        ' Determine data_type_id based on the following parameter:
        '   FileExt - The file extension of the document to be stored (req.)
        Dim dt As DataTable = New DataTable
        If Not DMSCache.GetCachedItem("DocumentTypes") Is Nothing Then
            ' Get document types from cache
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Document Types found in cache")
            Try
                dt = DMSCache.GetCachedItem("DocumentTypes")
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                GoTo CloseOut
            End Try
        Else
            ' Load document types into cache
            SqlS = "SELECT extension, row_id FROM DMS.dbo.Document_Types"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Document Types into cache: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If d_dr.HasRows Then
                    dt.Load(d_dr)
                    DMSCache.AddToCache("DocumentTypes", dt, CachingWrapper.CachePriority.NotRemovable)
                End If
                d_dr.Close()
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                GoTo CloseOut
            End Try
        End If
        If dt Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not retrieve Document Types. "
            GoTo CloseOut
        End If

        ' Debug output
        If Debug = "Y" Then
            mydebuglog.Debug(" Document Types Columns found: " & dt.Columns.Count.ToString)
            mydebuglog.Debug(" Document Types Rows found: " & dt.Rows.Count.ToString)
        End If

        ' ============================================
        ' Locate Document Type Id in datatable
        Dim drow() As DataRow = dt.Select("extension='" & FileExt & "'")
        If drow Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Document Type not found")
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Could not find association. "
            GoTo CloseOut
        End If
        data_type_id = drow(0).Item("row_id").ToString
        If Debug = "Y" Then mydebuglog.Debug(" Document Types row_id: " & data_type_id)
        Try
            dr = Nothing
            dt = Nothing
        Catch ex As Exception
        End Try

        ' Check data type id
        If data_type_id = "" Then
            results = "Failure"
            errmsg = errmsg & "Unknown data type id. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Locate document record in the DMS
        If DocId <> "" Then

            ' Query DMS for existing document id
            SqlS = "SELECT TOP 1 D.row_id, V.row_id, T.extension, D.dfilename, V.minio_flg " &
                "FROM DMS.dbo.Documents D " &
                "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id  " &
                "LEFT OUTER JOIN DMS.dbo.Document_Types T ON T.row_id=D.data_type_id " &
                "WHERE D.row_id=" & DocId
            If Debug = "Y" Then mydebuglog.Debug("  Verify provided DocId: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            VerifiedDocId = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                            DocVersionId = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                            If FileExt = "" Then FileExt = Trim(CheckDBNull(d_dr(2), enumObjectType.StrType))
                            If DFileName = "" Then DFileName = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                            minio_flg = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & vbCrLf & "Error verifying supplied doc id. " & ex.ToString
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & vbCrLf & "Error verifying supplied doc id."
                    d_dr.Close()
                    results = "Failure"
                    GoTo CloseOut
                End If
                d_dr.Close()

                If Debug = "Y" Then
                    mydebuglog.Debug(vbCrLf & "Document Information-")
                    mydebuglog.Debug("  VerifiedDocId: " & VerifiedDocId)
                    mydebuglog.Debug("  DocVersionId: " & DocVersionId)
                    mydebuglog.Debug("  FileExt: " & FileExt)
                    mydebuglog.Debug("  minio_flg: " & minio_flg)
                    mydebuglog.Debug("  DFileName: " & DFileName & vbCrLf)
                End If

                ' If not found then there is an error
                If VerifiedDocId = "" Then
                    results = "Failure"
                    errmsg = errmsg & vbCrLf & "Unable to verify document"
                    GoTo CloseOut
                End If

            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error verifying supplied doc id: " & oBug.ToString)
                results = "Failure"
            End Try
        Else
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Unable to verify document"
            GoTo CloseOut
        End If

        ' ============================================
        ' Extract and store binary locally of updated image
        ' Generate and remove temp file if applicable
        '   The filename of the temp file is in the format [Document Id]+.+[Extension]
        tempfile = SaveDest & "\" & VerifiedDocId & "." & FileExt
        If Debug = "Y" Then mydebuglog.Debug("   Temp file: " & tempfile & vbCrLf)
        Try
            If (My.Computer.FileSystem.FileExists(tempfile)) Then Kill(tempfile)
        Catch ex As Exception
        End Try

        ' If a Report stored in a buffer, cache in a local file
        If retval > 0 Then
            Try
                bfs = New FileStream(tempfile, FileMode.Create, FileAccess.Write)
                bw = New BinaryWriter(bfs)
                bw.Write(outbyte)
                bw.Flush()
                bw.Close()
                bfs.Close()
                Try
                    bfs.Dispose()
                    bfs = Nothing
                Catch ex2 As Exception
                End Try
                Try
                    bw.Dispose()
                    bw = Nothing
                Catch ex2 As Exception
                End Try
            Catch ex As Exception
                errmsg = errmsg & "Unable to write the report file to a temp file." & ex.ToString & vbCrLf
                results = "Failure"
                retval = 0
            End Try
            d_dsize = retval.ToString
        End If

        ' If a URL is provided, retrieve the data from that URL and cache locally
        If DURL <> "" Then
            Dim oRequest As System.Net.HttpWebRequest = CType(System.Net.HttpWebRequest.Create(DURL), System.Net.HttpWebRequest)
            Using oResponse As System.Net.WebResponse = CType(oRequest.GetResponse, System.Net.WebResponse)
                Using responseStream As System.IO.Stream = oResponse.GetResponseStream
                    Using bfs2 As New FileStream(tempfile, FileMode.Create, FileAccess.Write)
                        Dim buffer(2047) As Byte
                        Dim read As Integer
                        Do
                            read = responseStream.Read(buffer, 0, buffer.Length)
                            bfs2.Write(buffer, 0, read)
                        Loop Until read = 0
                        responseStream.Close()
                        bfs2.Flush()
                        bfs2.Close()
                        Try
                            bfs2.Dispose()
                        Catch ex As Exception
                        End Try
                        Try
                            buffer = Nothing
                        Catch ex As Exception
                        End Try
                    End Using
                    responseStream.Close()
                    responseStream.Dispose()
                End Using
                oResponse.Close()
            End Using
            Try
                oRequest = Nothing
            Catch ex As Exception
            End Try
            d_dsize = tempfile.Length.ToString
        End If

        ' If base64 encoded binary is supplied, retrieve and cache locally
        If Len(DImage) > 0 Then
            ' Convert input string into byte array, and then into binary
            Dim imagebuffer As Byte() = Convert.FromBase64String(DImage)
            If imagebuffer.Length = 0 Then
                ' No attachment found
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No attachment error. "
                GoTo CloseOut2
            End If
            Try
                bfs = New FileStream(tempfile, FileMode.Create, FileAccess.Write)
                bw = New BinaryWriter(bfs)
                bw.Write(imagebuffer)
                bw.Flush()
                bw.Close()
                bfs.Close()
                Try
                    bfs.Dispose()
                    bfs = Nothing
                Catch ex2 As Exception
                End Try
                Try
                    bw.Dispose()
                    bw = Nothing
                Catch ex2 As Exception
                End Try
            Catch ex As Exception
                errmsg = errmsg & "Unable to write the file to a temp file." & ex.ToString & vbCrLf
                results = "Failure"
                retval = 0
            End Try
            d_dsize = Len(DImage).ToString
        End If

        ' ============================================
        ' Document Versions record
        If UpdType = "N" Then
            ' Create Document_Versions record
            If ReportId <> "" Then
                ' If a report, set to not backup
                SqlS = "INSERT DMS.dbo.Document_Versions (doc_id, created, created_by, last_upd, last_upd_by, backed_up, dsize, version, minio_flg) " &
                    "SELECT " & VerifiedDocId & ", GETDATE(), 1, GETDATE(), 1, GETDATE(), '" & d_dsize & "', ISNULL(MAX([version]),0)+1, 'Y' FROM DMS.dbo.Document_Versions WHERE doc_id=" & VerifiedDocId & "; select Scope_Identity();"
            Else
                SqlS = "INSERT DMS.dbo.Document_Versions (doc_id, created, created_by, last_upd, last_upd_by, dsize, version, minio_flg) " &
                    "SELECT " & VerifiedDocId & ", GETDATE(), 1, GETDATE(), 1, '" & d_dsize & "', ISNULL(MAX([version]),0)+1, 'Y' FROM DMS.dbo.Document_Versions WHERE doc_id=" & VerifiedDocId & "; select Scope_Identity();"
            End If
            If Debug = "Y" Then mydebuglog.Debug("  Create basic document versions record: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                DocVersionId = d_cmd.ExecuteScalar()
                If Debug = "Y" Then mydebuglog.Debug("  > DocVersionId=" & DocVersionId)
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & "Error creating basic document versions record. " & ex.ToString & vbCrLf
                GoTo CloseOut
            End Try
        End If

        ' ============================================
        ' Retrieve document record and update with information supplied
        If DocVersionId <> "" Then

            ' Set configuration
            Dim MConfig As AmazonS3Config = New AmazonS3Config()
            'MConfig.RegionEndpoint = RegionEndpoint.USEast1
            MConfig.ServiceURL = "https://192.168.5.134"
            MConfig.ForcePathStyle = True
            MConfig.EndpointDiscoveryEnabled = False

            Dim Minio As AmazonS3Client = New AmazonS3Client(AccessKey, AccessSecret, MConfig)
            ServicePointManager.ServerCertificateValidationCallback = AddressOf sslhttps.AcceptAllCertifications

            ' If update delete existing object
            If UpdType = "R" Then
                Try
                    Dim mobj2 = Minio.DeleteObject(AccessBucket, VerifiedDocId & "-" & DocVersionId)
                    mobj2 = Nothing
                Catch ex2 As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error getting object. " & ex2.ToString & vbCrLf
                    GoTo CloseOut
                End Try
            End If

            ' Store object - new or updated
            Try
                Dim fileTransfer As Amazon.S3.Transfer.TransferUtility = New Amazon.S3.Transfer.TransferUtility(Minio)

                fileTransfer.Upload(tempfile, AccessBucket, DocVersionId)
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Unload complete of " & tempfile & " to bucket " & AccessBucket & " and object id " & DocVersionId)
                Try
                    fileTransfer = Nothing
                Catch ex2 As Exception
                    errmsg = errmsg & "Error closing fileTransfer: " & ex2.Message & vbCrLf
                End Try
            Catch ex As Exception
                errmsg = errmsg & "Error writing to Minio: " & ex.Message & vbCrLf
                results = "Failure"
            End Try

            Try
                Minio = Nothing
            Catch ex As Exception
                errmsg = errmsg & "Error closing Minio: " & ex.Message & vbCrLf
            End Try
        End If

        ' ============================================
        ' Update Document record
        '   ItemName	    - The "DMS.Documents.name" of the document to be stored. (req.)
        '   DFileName       - The "DMS.Documents.dfilename" of the document to be stored. (req.)
        '   Description     - The description of the document (req.)
        '   ExtId	        - The external id of the document to be stored (opt.)
        '   DocVersionId    - The FK to Document_Versions
        If UpdType = "N" Then
            SqlS = "UPDATE DMS.dbo.Documents " &
            "SET last_version_id=" & DocVersionId & " WHERE row_id=" & VerifiedDocId
            If Debug = "Y" Then mydebuglog.Debug("  Update Document record: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        If Debug = "Y" Then mydebuglog.Debug("Closing database connections " & vbCrLf)
        '   DMS
        Try
            d_dr = Nothing
            d_con.Dispose()
            d_con = Nothing
            d_cmd.Dispose()
            d_cmd = Nothing
        Catch ex As Exception
        End Try

        '   HCIDB
        Try
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Delete cached local temp file
        If tempfile <> "" Then
            If Debug = "Y" Then mydebuglog.Debug("Attempting to remove temp file: " & tempfile & vbCrLf)
            Try
                If Debug <> "Y" Then
                    If (My.Computer.FileSystem.FileExists(tempfile)) Then Kill(tempfile)
                End If
            Catch ex As Exception
            End Try
        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ltemp = results & " for filename " & DFileName & "." & FileExt & " and report id " & ReportId & " to document id " & VerifiedDocId
        If Trim(errmsg) <> "" Then myeventlog.Error("UpdateDoc :  Error: " & Trim(errmsg))
        myeventlog.Info("UpdateDoc : Results: " & ltemp)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return item
        If VerifiedDocId <> "" And Debug <> "T" Then
            Return VerifiedDocId
        Else
            Return Nothing
        End If
    End Function

    <WebMethod(Description:="Remove the specified document")>
    Public Function DelDocument(ByVal UserId As String, ByVal Asset As String, ByVal DocId As String, ByVal DelType As String, ByVal Debug As String) As String

        ' This function locates the specified item, and deletes it 

        ' The input parameters are as follows:
        '
        '   UserId      - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                   user.
        '
        '   Asset       - The DMS.Documents.dfilename of the asset to be retrieved, or if "default.jpg",
        '                   the first associated item in the category "Images" will be returned.
        '   
        '   DocId       - The DMS.Documents.row_id of the document to be retrieved
        '
        '   DelType     - The type of deletion = "P"=permanent, anything else soft deletion
        '   
        '   Debug       - "Y", "N" or "T"

        ' web.config Parameters used:
        '   hcidb           - connection string to hcidb1.siebeldb database
        '   dms             - connection string to DMS.dms database
        '   cache           - connection string to cache.sdf database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String
        Dim DecodedUserId, ValidatedUserId As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer
        Dim TypeTrans As String

        ' Cache database declarations
        Dim c_ConnS As String
        Dim CacheHit As Integer
        Dim dAsset, dCrseId, dFileName As String
        Dim LastUpd As DateTime
        Dim d_last_updated As DateTime

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String
        Dim dms_cache_age As String

        ' Logging declarations
        Dim myeventlog = log4net.LogManager.GetLogger("EventLog")
        Dim mydebuglog = log4net.LogManager.GetLogger("DDDebugLog")
        Dim logfile, temp As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

        ' File handling declarations
        'Dim bfs As FileStream
        'Dim bw As BinaryWriter
        Dim outbyte(1000) As Byte
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim d_dsize, d_doc_id, d_verid, SaveDest, DMS_USER_ID, DMS_UA_ID As String
        Dim dLastUpd As DateTime
        Dim CRSE_ID, minio_flg As String
        Dim killcount As Double

        Dim filecache As ObjectCache = MemoryCache.Default
        Dim fileContents(1000) As Byte

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        ValidatedUserId = ""
        CRSE_ID = ""
        d_dsize = ""
        BinaryFile = ""
        c_ConnS = ""
        dAsset = ""
        d_doc_id = ""
        dCrseId = ""
        dFileName = ""
        TypeTrans = ""
        'Debug = "Y"
        killcount = 0
        temp = ""
        DMS_USER_ID = ""
        DMS_UA_ID = ""
        minio_flg = "N"

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If Asset = "" And DocId = "" And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        UserId = Trim(HttpUtility.UrlEncode(UserId))
        If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
        If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
        DecodedUserId = FromBase64(ReverseString(UserId))
        If InStr(Asset, "%") > 0 Then Asset = Trim(HttpUtility.UrlDecode(Asset))
        If DelType <> "P" Then DelType = ""
        If Trim(Asset) = "" And DocId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No item specified. "
            GoTo CloseOut2
        End If
        Dim RegExStr As String = "[\\/:*?""<>|]"  'For eliminating Characters: \ / : * ? "  |
        Asset = Regex.Replace(Asset, RegExStr, "")

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidbro").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb;ApplicationIntent=ReadOnly"
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIBDB;uid=DMS;pwd=5241200;database=DMS;ApplicationIntent=ReadOnly"
            dms_cache_age = Trim(System.Configuration.ConfigurationManager.AppSettings("dmscacheage"))
            If dms_cache_age = "" Or Not IsNumeric(dms_cache_age) Then dms_cache_age = "30"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("DelDocument_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
            AccessKey = System.Configuration.ConfigurationManager.AppSettings("minio-key")
            If AccessKey = "" Then AccessKey = "dms"
            AccessSecret = System.Configuration.ConfigurationManager.AppSettings("minio-secret")
            If AccessSecret <> "" Then AccessSecret = System.Web.HttpUtility.HtmlDecode(AccessSecret)
            If AccessSecret = "" Then AccessSecret = "SfI@aUE$?=&KcAOI?C5NU|-c*Oec7ZPJ"
            AccessRegion = System.Configuration.ConfigurationManager.AppSettings("minio-region")
            If AccessRegion = "" Then AccessRegion = "us-east"
            AccessBucket = System.Configuration.ConfigurationManager.AppSettings("minio-bucket")
            If AccessBucket = "" Then AccessBucket = "default"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\DelDocument.log"
            Try
                log4net.GlobalContext.Properties("DDLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Asset: " & Asset)
                mydebuglog.Debug("  DelType: " & DelType)
                mydebuglog.Debug("  AccessBucket: " & AccessBucket)
                mydebuglog.Debug("  AccessRegion: " & AccessRegion)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  Appsetting dms_cache_age: " & dms_cache_age & vbCrLf)
            End If
        End If

        ' ============================================
        ' Open SQL Server database connections
        errmsg = OpenDBConnection(ConnS, con, cmd)          ' hcidb1
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)    ' dms
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Validate identity if needed
        If UserId <> "" Then
            If Not cmd Is Nothing Then
                ' -----
                ' Query registration
                SqlS = "SELECT C.ROW_ID " &
                    "FROM siebeldb.dbo.S_CONTACT C " &
                    "WHERE C.X_REGISTRATION_NUM='" & UserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get contact: " & SqlS)
                cmd.CommandText = SqlS
                Try
                    cmd.CommandText = SqlS
                    DecodedUserId = cmd.ExecuteScalar()
                Catch ex As Exception
                    errmsg = errmsg & "Error getting user. " & ex.ToString & vbCrLf
                End Try

                ' Get DMS user id
                If DecodedUserId <> "" Then
                    SqlS = "SELECT row_id FROM DMS.dbo.Users WHERE ext_user_id='" & DecodedUserId & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Get DMS_USER_ID: " & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        DMS_USER_ID = cmd.ExecuteScalar()
                    Catch ex As Exception
                        errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                    End Try
                End If

                ' Get DMS_UA_ID
                If DMS_USER_ID <> "" Then
                    SqlS = "SELECT DISTINCT A.row_id " &
                    "FROM DMS.dbo.Users U " &
                    "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=U.row_id " &
                    "WHERE A.type_id='U' AND U.row_id=" & DMS_USER_ID
                    If Debug = "Y" Then mydebuglog.Debug("  Get DMS_UA_ID: " & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        DMS_UA_ID = cmd.ExecuteScalar()
                    Catch ex As Exception
                        errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                    End Try
                End If

            Else
                results = "Failure"
                GoTo CloseOut
            End If
        End If
        If DMS_USER_ID = "" Then DMS_USER_ID = "1"

        ' ============================================
        ' Create output directory for asset caching
        SaveDest = mypath & "document_temp"
        Try
            Directory.CreateDirectory(SaveDest)
        Catch
        End Try
        If Debug = "Y" Then mydebuglog.Debug("   Asset caching directory: " & SaveDest & vbCrLf)

        ' ============================================
        ' Get asset information if necessary
        If Debug = "Y" Then mydebuglog.Debug("   Looking for: " & LCase(Trim(System.IO.Path.GetFileNameWithoutExtension(Asset))) & vbCrLf)
        If LCase(Trim(System.IO.Path.GetFileNameWithoutExtension(Asset))) = "default" Then
            If Not d_cmd Is Nothing Then
                ' Query DMS
                SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd " &
                "FROM DMS.dbo.Documents D  " &
                "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id  " &
                "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id  " &
                "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id  " &
                "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id  " &
                "WHERE D.row_id IS NOT NULL AND D.deleted IS NULL AND LOWER(A.name)='Contact' " &
                "AND DA.fkey='" & DecodedUserId & "' AND D.name='" & Asset & "'" &
                "ORDER BY D.last_upd DESC"
                If Debug = "Y" Then mydebuglog.Debug("  Get item: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                Asset = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                                d_last_updated = CheckDBNull(d_dr(2), enumObjectType.DteType)
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_doc_id=" & d_doc_id & "  Asset=" & Asset)
                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error getting default item. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "Error getting default item." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting default item: " & oBug.ToString)
                    results = "Failure"
                End Try
            End If
        End If

        ' -----
        ' Locate last updated
        If Asset = "" And DocId <> "" And d_doc_id = "" Then
            If Not d_cmd Is Nothing Then
                ' Query DMS
                SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd " &
                    "FROM DMS.dbo.Documents D  " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id  " &
                    "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id  " &
                    "WHERE D.row_id=" & DocId & " AND D.deleted IS NULL " &
                    "ORDER BY D.last_upd DESC"
                If Debug = "Y" Then mydebuglog.Debug("  Get document information for document id specified: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                Asset = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                                d_last_updated = CheckDBNull(d_dr(2), enumObjectType.DteType)
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_doc_id=" & d_doc_id & "  Asset=" & Asset)

                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error getting default item. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "Error getting default item." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting default item: " & oBug.ToString)
                    results = "Failure"
                End Try
            End If
        End If

        ' ============================================
        ' Process deletion 
        If Not d_cmd Is Nothing And d_doc_id <> "" Then

            ' -----
            ' Remove from the cache
            BinaryFile = SaveDest & "\" & Asset
            'BinaryFile = BinaryFile.Replace(mypath, "")
            If Debug = "Y" Then mydebuglog.Debug("  Cache filename: " & BinaryFile & vbCrLf)
            Try
                filecache.Remove(BinaryFile)
            Catch ex As Exception
                errmsg = errmsg & "Error removing from cache. " & ex.ToString & vbCrLf
            End Try

            ' -----
            ' Get information from versions record
            d_verid = ""
            SqlS = "SELECT TOP 1 v.dsize, v.row_id, v.last_upd, v.minio_flg " &
                    "FROM DMS.dbo.Documents d " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                    "WHERE d.row_id=" & d_doc_id
            If Debug = "Y" Then mydebuglog.Debug("  Get information from versions record: " & SqlS)
            d_cmd.CommandText = SqlS
            d_dr = d_cmd.ExecuteReader()
            If Not d_dr Is Nothing Then
                While d_dr.Read()
                    Try
                        d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                        d_verid = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                        dLastUpd = CheckDBNull(d_dr(2), enumObjectType.DteType)
                        minio_flg = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                        If Debug = "Y" Then mydebuglog.Debug("  > Record found on query id " & d_verid & ":  d_dsize=" & d_dsize & ", minio_flg=" & minio_flg & ", dLastUpd=" & Format(dLastUpd) & ", cLastUpd=" & Convert.ToString(LastUpd) & ", CacheHit=" & Format(CacheHit))
                    Catch ex As Exception
                        results = "Failure"
                        errmsg = errmsg & "Error getting asset. " & ex.ToString & vbCrLf
                        GoTo CloseOut
                    End Try
                End While
            Else
                errmsg = errmsg & "Error getting asset." & vbCrLf
                d_dr.Close()
                results = "Failure"
            End If
            d_dr.Close()

            ' -----
            ' Remove binary from versions record
            If DelType = "P" And d_verid <> "" And minio_flg = "Y" Then
                Try
                    ' Remove binary from Minio
                    Dim MConfig As AmazonS3Config = New AmazonS3Config()
                    'MConfig.RegionEndpoint = RegionEndpoint.USEast1
                    MConfig.ServiceURL = "https://192.168.5.134"
                    MConfig.ForcePathStyle = True
                    MConfig.EndpointDiscoveryEnabled = False
                    Dim Minio As AmazonS3Client = New AmazonS3Client(AccessKey, AccessSecret, MConfig)
                    ServicePointManager.ServerCertificateValidationCallback = AddressOf sslhttps.AcceptAllCertifications
                    Try
                        Dim mobj2 = Minio.DeleteObject(AccessBucket, d_doc_id & "-" & d_verid)
                        mobj2 = Nothing
                    Catch ex2 As Exception
                        results = "Failure"
                        errmsg = errmsg & "Error getting object. " & ex2.ToString & vbCrLf
                        GoTo CloseOut
                    End Try
                    Try
                        Minio = Nothing
                    Catch ex As Exception
                        errmsg = errmsg & "Error closing Minio: " & ex.Message & vbCrLf
                    End Try
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error removing document: " & oBug.ToString)
                    results = "Failure"
                End Try
            End If

            ' -----
            ' Process record deletion as applicable
            If DelType = "P" Then
                Try
                    SqlS = "DELETE FROM DMS.dbo. Document_Categories WHERE doc_id=" & d_doc_id
                    If Debug = "Y" Then mydebuglog.Debug("  Remove Document_Categories: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "DELETE FROM DMS.dbo.Document_Keywords WHERE doc_id=" & d_doc_id
                    If Debug = "Y" Then mydebuglog.Debug("  Remove Document_Keywords: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "DELETE FROM DMS.dbo.Document_Associations WHERE doc_id=" & d_doc_id
                    If Debug = "Y" Then mydebuglog.Debug("  Remove Document_Associations: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "DELETE FROM DMS.dbo.Document_Users WHERE doc_id=" & d_doc_id
                    If Debug = "Y" Then mydebuglog.Debug("  Remove Document_Users: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "DELETE FROM DMS.dbo.Document_Versions WHERE doc_id=" & d_doc_id
                    If Debug = "Y" Then mydebuglog.Debug("  Remove Document_Versions: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "DELETE FROM DMS.dbo.Documents WHERE row_id=" & d_doc_id
                    If Debug = "Y" Then mydebuglog.Debug("  Remove Documents: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "UPDATE siebeldb.dbo.CX_SESSIONS_X SET IMAGE_KEY=NULL WHERE IMAGE_KEY='" & d_doc_id & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Update CX_SESSIONS_X: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "DELETE FROM siebeldb.dbo.CX_REP_ENT_SCHED " &
                    "WHERE DMS_DOC_ID='" & d_doc_id & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Remove CX_REP_ENT_SCHED: " & SqlS & vbCrLf)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()
                    results = "Success"
                Catch ex As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error removing document: " & ex.ToString)
                    results = "Failure"
                End Try
            Else
                Try
                    SqlS = "DELETE FROM DMS.dbo.Document_Users WHERE doc_id=" & d_doc_id
                    If Debug = "Y" Then mydebuglog.Debug("  Remove Document_Users records: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "UPDATE DMS.dbo.Documents " &
                    "SET deleted=GETDATE(), last_upd_by=" & DMS_USER_ID &
                    " WHERE row_id=" & d_doc_id
                    If Debug = "Y" Then mydebuglog.Debug("  Set Documents record to deleted: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    If d_verid <> "" Then
                        SqlS = "UPDATE DMS.dbo.Document_Versions " &
                        "SET deleted=GETDATE(), last_upd_by=" & DMS_USER_ID &
                        " WHERE row_id=" & d_verid
                        If Debug = "Y" Then mydebuglog.Debug("  Set Document_Versions record to deleted: " & SqlS)
                        d_cmd.CommandText = SqlS
                        returnv = d_cmd.ExecuteNonQuery()
                    End If

                    If DMS_UA_ID <> "" Then
                        SqlS = "INSERT INTO DMS.dbo.Document_Users " &
                        "(doc_id,user_access_id,created_by,last_upd_by,owner_flag,access_type) " &
                        "VALUES (" & d_doc_id & "," & DMS_UA_ID & "," & DMS_USER_ID & "," & DMS_USER_ID & ",'Y','REDO')"
                        If Debug = "Y" Then mydebuglog.Debug("  Created deleting user document access for undelete: " & SqlS)
                        d_cmd.CommandText = SqlS
                        returnv = d_cmd.ExecuteNonQuery()
                    End If

                    SqlS = "DELETE FROM DMS.dbo.Document_Associations " &
                    "WHERE doc_id=" & d_doc_id & " and association_id=3 and fkey is not null and fkey<>''"
                    If Debug = "Y" Then mydebuglog.Debug("  Delete Contact Document_Associations: " & SqlS)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()

                    SqlS = "DELETE FROM siebeldb.dbo.CX_REP_ENT_SCHED " &
                    "WHERE DMS_DOC_ID='" & d_doc_id & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Remove CX_REP_ENT_SCHED: " & SqlS & vbCrLf)
                    d_cmd.CommandText = SqlS
                    returnv = d_cmd.ExecuteNonQuery()
                    results = "Success"
                Catch ex As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error removing document: " & ex.ToString)
                    results = "Failure"
                End Try
            End If
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
            d_dr = Nothing
            d_con.Dispose()
            d_con = Nothing
            d_cmd.Dispose()
            d_cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("DelDocument : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("DelDocument : Results: " & results & " for file " & Asset & ", by UserId # " & DecodedUserId & " and doc id " & d_doc_id)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & results & " for file " & Asset & ", by UserId # " & DecodedUserId & " and doc id " & d_doc_id & " at " & Now.ToString)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Try
            Return results
        Catch exp As Exception
            Return Nothing
        Finally
            outbyte = Nothing
            fileContents = Nothing
        End Try
    End Function

    <WebMethod(Description:="Rewrite and executes a query to a cache table")>
    Public Function EmpQuery(ByVal EmpId As String, ByVal ConId As String, ByVal FunctionId As String, ByVal TranId As String, ByVal Query As String, ByVal DMSFlag As String,
        ByVal Refresh As String, ByVal Debug As String) As XmlDocument

        ' Given a valid query, this function generates a result table from the query.

        ' The parameters are as follows:
        '   EmpId           - A FK to S_EMPLOYEE.ROW_ID
        '   ConId           - A FK to S_CONTACT.ROW_ID
        '   FunctionId      - A key representing the function that executed the query
        '   TranId          - A key representing the transaction id of the query
        '   Query           - The query itself
        '   DMSFlag         - A flag to indicate that the query is against the DMS server
        '   Refresh         - A flag to force a refresh
        '   Debug           - A flag to indicate the service is to run in Debug mode or not
        '                       "Y"  - Yes for debug mode on.. logging on
        '                       "N"  - No for debug mode off.. logging off
        '                       "T"  - Test mode on.. logging off

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Service dependencies:
        '   n/a

        ' Variables
        Dim results As String
        Dim mypath, errmsg, ErrorDesc, temp As String
        Dim QAttempt As Integer
        Dim UnionFlag As Boolean

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv, returnvd As Integer
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim logfile, Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim Dt1 As DateTime
        Dim TimeSpent As Double
        Dim ltemp As String
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        Dim VersionNum As String = "100"
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("QDebugLog")

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        Dim QueryLogService As New com.certegrity.cloudsvc.cm.Service

        ' Results variables
        Dim TableName, TempSessionId, LoggedOut, Tmp_TableName As String
        Dim SavedDMSFlag As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        errmsg = ""
        results = "Success"
        Logging = "Y"
        SqlS = ""
        TempSessionId = ""
        TableName = ""
        Tmp_TableName = ""
        LoggedOut = "N"
        SavedDMSFlag = ""
        ErrorDesc = ""
        QAttempt = 0
        returnvd = 0
        returnv = 0
        temp = ""
        UnionFlag = False

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb;Min Pool Size=3;Max Pool Size=5"
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS;Min Pool Size=3;Max Pool Size=5"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("EmpQuery_Debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Fix parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            EmpId = "1-2T"
            ConId = "21120611WE0"
            FunctionId = "N"
            Query = "SELECT TOP 1 ROW_ID FROM siebeldb.dbo.S_CONTACT"
            DMSFlag = "N"
            Refresh = "N"
        Else
            If InStr(EmpId, "%") > 0 Then EmpId = Trim(HttpUtility.UrlDecode(EmpId)) Else EmpId = EmpId.Trim
            If InStr(ConId, "%") > 0 Then ConId = Trim(HttpUtility.UrlDecode(ConId)) Else ConId = ConId.Trim
            If InStr(FunctionId, "%") > 0 Then FunctionId = Trim(HttpUtility.UrlDecode(FunctionId)) Else FunctionId = FunctionId.Trim
            If InStr(TranId, "%") > 0 Then TranId = Trim(HttpUtility.UrlDecode(TranId)) Else TranId = TranId.Trim
            If InStr(Query, "%20") > 0 Then Query = Trim(HttpUtility.UrlDecode(Query)) Else Query = Query.Trim
            If Refresh.ToLower = "true" Or Refresh.ToLower = "y" Then Refresh = "Y" Else Refresh = "N"
            If DMSFlag = "" Then DMSFlag = "N"
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            logfile = "C:\logs\EmpQuery.log"
            Try
                log4net.GlobalContext.Properties("QLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & "Error Opening Log. " & vbCrLf
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                Try
                    mydebuglog.Debug("----------------------------------")
                    mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                    mydebuglog.Debug("Parameters-")
                    mydebuglog.Debug("  EmpId: " & EmpId)
                    mydebuglog.Debug("  ConId: " & ConId)
                    mydebuglog.Debug("  FunctionId: " & FunctionId)
                    mydebuglog.Debug("  TranId: " & TranId)
                    mydebuglog.Debug("  Query: " & Query)
                    mydebuglog.Debug("  DMSFlag: " & DMSFlag)
                    mydebuglog.Debug("  Refresh: " & Refresh)
                Catch ex As Exception
                End Try
            End If
        End If

        ' ============================================
        ' Check parameters
        '   UserId and SessionId are required under all circumstances
        If EmpId = "" And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & "Invalid parameter(s) " & vbCrLf
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open database connection and retrieve record
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process request 
        If Not cmd Is Nothing Then
            Try
                ' -----
                ' Create the tablename and drop the table
                If DMSFlag = "Y" Then
                    TableName = "CM.dbo.[" & EmpId & "-" & EmpId & "]"
                    If FunctionId = "GetContent" Then
                        Tmp_TableName = "CM.dbo.[" & EmpId & "-" & EmpId & "-T]"
                    End If
                Else
                    TableName = "CM.dbo.[" & EmpId & "-" & EmpId & "]"
                End If

                ' -----
                ' Check to see if the table has already been queried or not
                Dim QueryId, NumRows, RefreshFlg, SavedQuery As String
                QueryId = ""
                NumRows = "0"
                RefreshFlg = ""
                SavedQuery = ""
                SqlS = "SELECT ROW_ID, ROWS, REFRESH_FLAG, QUERY, DMS_FLAG " &
                "FROM reports.dbo.CM_QUERIES " &
                "WHERE USER_ID='" & EmpId & "' AND SESSION_ID='" & EmpId & "' AND TRAN_ID='" & TranId & "' AND DMS_FLAG='" & DMSFlag & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Checking for existing result set: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            QueryId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            NumRows = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            RefreshFlg = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            SavedQuery = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            SavedDMSFlag = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        Catch ex As Exception
                        End Try
                    End While
                End If
                If RefreshFlg <> "" Then RefreshFlg = Refresh
                If Debug = "Y" Then
                    mydebuglog.Debug("   ... TableName: " & TableName)
                    mydebuglog.Debug("   ... Tmp_TableName: " & Tmp_TableName)
                    mydebuglog.Debug("   ... QueryId: " & QueryId)
                    mydebuglog.Debug("   ... NumRows: " & NumRows)
                    mydebuglog.Debug("   ... RefreshFlg: " & RefreshFlg)
                    mydebuglog.Debug("   ... SavedQuery: " & SavedQuery)
                    mydebuglog.Debug("   ... SavedDMSFlag: " & SavedDMSFlag)
                End If
                dr.Close()

                ' -----
                ' Set the query if none was supplied
                If Query = "" And SavedQuery <> "" Then Query = SavedQuery
                If Query = "" Then
                    results = "Failure"
                    errmsg = errmsg & "Invalid parameter(s) " & vbCrLf
                    TableName = ""
                    GoTo CloseOut
                End If

                ' -----
                ' Check found temp table to see if it contains the same records
                If QueryId <> "" Then
                    Dim CheckTranId As String
                    CheckTranId = ""
                    SqlS = "SELECT TOP 1 CM_TRAN_ID " &
                    "FROM " & TableName
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Checking existing temp table: " & SqlS)
                    If DMSFlag = "Y" Then
                        Try
                            cmd.CommandText = SqlS
                            dr = cmd.ExecuteReader()
                            If Not dr Is Nothing Then
                                While dr.Read()
                                    Try
                                        CheckTranId = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                    Catch ex As Exception
                                    End Try
                                End While
                            End If
                            dr.Close()
                        Catch ex As Exception
                            CheckTranId = ""
                            RefreshFlg = "N"
                        End Try
                    Else
                        Try
                            cmd.CommandText = SqlS
                            dr = cmd.ExecuteReader()
                            If Not dr Is Nothing Then
                                While dr.Read()
                                    Try
                                        CheckTranId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                    Catch ex As Exception
                                    End Try
                                End While
                            End If
                            dr.Close()
                        Catch ex As Exception
                            CheckTranId = ""
                            RefreshFlg = "N"
                        End Try
                    End If
                    If Debug = "Y" Then mydebuglog.Debug("   ... CheckTranId: " & CheckTranId)
                    If CheckTranId.Trim <> TranId.Trim Then QueryId = ""
                End If

                ' -----
                ' If Refresh or New, then Drop Table on both dms and hcidb1
DropTables:
                If QueryId = "" Or RefreshFlg = "Y" Or NumRows = "0" Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Dropping temp tables: [" & EmpId & "-" & EmpId & "]")
                    Try
                        cmd.CommandText = "IF OBJECT_ID('CM.dbo.[" & EmpId & "-" & EmpId & "]') IS NOT NULL DROP TABLE CM.dbo.[" & EmpId & "-" & EmpId & "]"
                        returnv = cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End If

                ' -----
                ' Check for requery of same result set
                If QueryId <> "" And NumRows <> "0" And RefreshFlg = "N" Then
                    'TableName = ""
                    returnv = NumRows
                    GoTo CloseOut
                End If

                ' -----
                ' Generate the temp table query
                If FunctionId = "GetContent" And Tmp_TableName <> "" Then
                    SqlS = ", '" & TranId & "' AS CM_TRAN_ID INTO " & Tmp_TableName & " "
                Else
                    SqlS = ", '" & TranId & "' AS CM_TRAN_ID INTO " & TableName & " "
                End If

                ' <<< New code
                If FunctionId = "GetContent" Then
                    SqlS = Left(Query, InStr(1, Query, "FROM ") - 1) & SqlS
                    SqlS = SqlS & Right(Query, Len(Query) - InStr(1, Query, "FROM ") + 1)
                Else
                    If InStr(1, Query.ToUpper, "UNION") > 0 Then
                        Dim SplitQuery() As String = Split(Query, "FROM", -1, CompareMethod.Text)
                        If SplitQuery.Length > 1 Then
                            If Debug = "Y" Then
                                mydebuglog.Debug("  SplitQuery(0): " & SplitQuery(0))
                                mydebuglog.Debug("  SplitQuery(1): " & SplitQuery(1))
                            End If
                            SqlS = SplitQuery(0) & SqlS & " FROM " & SplitQuery(1)
                            For I As Integer = 2 To SplitQuery.Length - 1
                                SqlS = SqlS & ", '" & TranId & "' AS CM_TRAN_ID " & " FROM " & SplitQuery(I)
                                If Debug = "Y" Then mydebuglog.Debug("  SplitQuery(" & I.ToString & "): " & SplitQuery(I))
                            Next
                        End If
                    Else
                        SqlS = Left(Query, InStr(1, Query, "FROM ") - 1) & SqlS
                        SqlS = SqlS & Right(Query, Len(Query) - InStr(1, Query, "FROM ") + 1)

                    End If
                End If
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Generating query: " & SqlS)

                ' Remove cross-apply temp table
                If Tmp_TableName <> "" Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Dropping GetContent temp table: " & Tmp_TableName & vbCrLf)
                    Try
                        cmd.CommandText = "IF OBJECT_ID('" & Tmp_TableName & "') IS NOT NULL DROP TABLE " & Tmp_TableName
                        returnv = cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End If

                ' -----
                ' Execute temp-table query
                returnv = 0
                Try
                    If DMSFlag = "Y" Then
                        cmd.CommandText = SqlS
                        returnv = cmd.ExecuteNonQuery()
                    Else
                        cmd.CommandText = SqlS
                        returnv = cmd.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                    ' If this fails, it likely means the temp table already exists
                    ErrorDesc = ex.ToString
                    If InStr(ErrorDesc, "There is already an object named") > 0 Then
                        QAttempt = QAttempt + 1
                        RefreshFlg = "Y"
                        If QAttempt < 3 Then GoTo DropTables
                    End If
                    results = "Failure"
                    errmsg = errmsg & "Error writing table. " & ErrorDesc & vbCrLf
                End Try

                ' -----
                ' Perform cross-join on GetContent results to collapse detail records
                ' 
                If FunctionId = "GetContent" And Tmp_TableName <> "" And TableName <> "" Then
                    SqlS = "SELECT DISTINCT row_id, name, description, type, dfilename, category=left(cat.list, len(cat.list)-1), " &
                    "cat_id=left(cattype.list, len(cattype.list)-1), cat_pr_flag=left(catpr.list, len(catpr.list)-1), created, last_upd,  " &
                    "type_id=left(acctype.list, len(acctype.list)-1), access_type=left(acc.list, len(acc.list)-1), " &
                    "key_id=left(keys.list, len(keys.list)-1), CM_TRAN_ID " &
                    "INTO CM.dbo.[" & EmpId & "-" & EmpId & "] " &
                    "FROM CM.dbo.[" & EmpId & "-" & EmpId & "-T] b " &
                    "CROSS APPLY  " &
                    "( " &
                    "   SELECT category+',' AS [text()] " &
                    "      FROM " & Tmp_TableName & " d " &
                    "      WHERE b.row_id = d.row_id and category is not null " &
                    "      GROUP BY category " &
                    "      ORDER BY category " &
                    "      FOR XML PATH('') " &
                    ") cat (list) " &
                    "CROSS APPLY  " &
                    "( " &
                    "   SELECT cast(cat_id as varchar)+',' AS [text()] " &
                    "      FROM " & Tmp_TableName & " e " &
                    "      WHERE b.row_id = e.row_id and cat_id is not null " &
                    "      GROUP BY cat_id " &
                    "      ORDER BY cat_id " &
                    "      FOR XML PATH('') " &
                    "   ) cattype (list) " &
                    "CROSS APPLY  " &
                    "( " &
                    "   SELECT cat_pr_flag+',' AS [text()] " &
                    "      FROM " & Tmp_TableName & " h " &
                    "      WHERE b.row_id = h.row_id and cat_pr_flag is not null " &
                    "      GROUP BY cat_pr_flag " &
                    "      ORDER BY cat_pr_flag " &
                    "      FOR XML PATH('') " &
                    "   ) catpr (list) " &
                    "CROSS APPLY  " &
                    "( " &
                    "   SELECT access_type+',' AS [text()] " &
                    "      FROM " & Tmp_TableName & " f " &
                    "      WHERE b.row_id = f.row_id and access_type is not null " &
                    "      GROUP BY access_type " &
                    "      ORDER BY access_type " &
                    "      FOR XML PATH('') " &
                    "   ) acc (list) " &
                    "CROSS APPLY  " &
                    "( " &
                    "   SELECT type_id+',' AS [text()] " &
                    "      FROM " & Tmp_TableName & " g " &
                    "      WHERE b.row_id = g.row_id and type_id is not null " &
                    "      GROUP BY type_id " &
                    "      ORDER BY type_id " &
                    "      FOR XML PATH('') " &
                    ") acctype (list) " &
                    "CROSS APPLY  " &
                    "( " &
                    "   SELECT cast(key_id as varchar)+',' AS [text()] " &
                    "      FROM " & Tmp_TableName & " j " &
                    "      WHERE b.row_id = j.row_id and key_id is not null " &
                    "      GROUP BY key_id " &
                    "      ORDER BY key_id " &
                    "      FOR XML PATH('') " &
                    ") keys (list) " &
                    "WHERE b.row_id IS NOT NULL"
                    If Debug = "Y" Then mydebuglog.Debug("  Performing cross join: " & vbCrLf & SqlS & vbCrLf)
                    Try
                        cmd.CommandText = SqlS
                        returnv = cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try

                    ' Remove cross-apply temp table
                    SqlS = "IF OBJECT_ID('" & Tmp_TableName & "') IS NOT NULL DROP TABLE " & Tmp_TableName
                    If Debug = "Y" Then mydebuglog.Debug("  Dropping GetContent temp table: " & SqlS & vbCrLf)
                    Try
                        cmd.CommandText = SqlS
                        returnvd = cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End If

                ' If returned no record count, then the query failed, so do not return a tablename
                If Debug = "Y" Then mydebuglog.Debug("   ... returnv: " & returnv.ToString)
                'If returnv = 0 Then TableName = ""
                If Debug = "Y" Then mydebuglog.Debug("   ... TableName: " & TableName & vbCrLf)

            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString & vbCrLf)
                results = "Failure"
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            errmsg = errmsg & CloseDBConnection(con, cmd, dr)
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' -----
        ' Log the query using the logging service
        If TableName <> "" Then
            Try
                If temp = "Y" Then
                    results = QueryLogService.CMLogQuery(EmpId, TableName, EmpId, FunctionId, TranId, Query, DMSFlag, Debug)
                Else
                    results = QueryLogService.CMLogQuery(EmpId, TableName, EmpId, FunctionId, TranId, Query, DMSFlag, Debug)
                End If
                If results <> "Success" Then errmsg = errmsg & "Logging error. " & vbCrLf
            Catch ex As Exception
                errmsg = errmsg & "Logging error. " & ex.ToString & vbCrLf
            End Try
        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        ' Compute time spent
        Dt1 = DateTime.Parse(LogStartTime)
        Dim runLength As Global.System.TimeSpan = Now.Subtract(Dt1)
        TimeSpent = runLength.TotalMilliseconds
        ltemp = "Generated " & HttpUtility.HtmlEncode(TableName) & " at " & Format(Now) & " for " & FunctionId & " transaction id " & TranId & " in " & TimeSpent.ToString & " milliseconds"
        If Trim(errmsg) <> "" Then myeventlog.Error("EmpQuery :  Error: " & Trim(errmsg))
        myeventlog.Info("EmpQuery : Results: " & ltemp)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                ' Log results
                If Trim(errmsg) <> "" Then mydebuglog.Debug("  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results : " & ltemp)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)

        ' ============================================
        ' Return results
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("results")
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        If Debug = "T" Then
            If TableName = "" Then
                AddXMLAttribute(odoc, resultsRoot, "msg", "Failure")
            Else
                AddXMLAttribute(odoc, resultsRoot, "msg", "Success")
            End If
        Else
            AddXMLAttribute(odoc, resultsRoot, "tablename", TableName)
            AddXMLAttribute(odoc, resultsRoot, "records", returnv.ToString)
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        End If

        'results = HttpUtility.HtmlEncode(TableName)
        'Return results
        Return odoc

    End Function

    ' =================================================
    ' EMAIL
    Public Function PrepareMail(ByVal FromEmail As String, ByVal ToEmail As String, ByVal Subject As String, _
        ByVal Body As String, ByVal Debug As String, ByRef mydebuglog As ILog) As Boolean
        ' This function wraps message info into the XML necessary to call the SendMail web service function.
        ' This is used by other services executing from this application.
        ' Assumptions:  Create a record in MESSAGES and IDs are unknown 
        Dim wp As String

        ' Web service declarations
        Dim EmailService As New basic.com.certegrity.cloudsvc.Service

        wp = "<EMailMessageList><EMailMessage>"
        wp = wp & "<debug>" & Debug & "</debug>"
        wp = wp & "<database>C</database>"
        wp = wp & "<Id> </Id>"
        wp = wp & "<SourceId></SourceId>"
        wp = wp & "<From>" & FromEmail & "</From>"
        wp = wp & "<FromId></FromId>"
        wp = wp & "<FromName></FromName>"
        wp = wp & "<To>" & ToEmail & "</To>"
        wp = wp & "<ToId></ToId>"
        wp = wp & "<Cc></Cc>"
        wp = wp & "<Bcc></Bcc>"
        wp = wp & "<ReplyTo>" & FromEmail & "</ReplyTo>"
        wp = wp & "<Subject>" & Subject & "</Subject>"
        wp = wp & "<Body>" & Body & "</Body>"
        wp = wp & "<Format></Format>"
        wp = wp & "</EMailMessage></EMailMessageList>"
        If Debug = "Y" Then mydebuglog.Debug("  > Email XML: " & wp)

        PrepareMail = EmailService.SendMail(wp)

    End Function

    ' =================================================
    ' NUMERIC
    Public Function Round(ByVal nValue As Double, ByVal nDigits As Integer) As Double
        Round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
    End Function

    ' =================================================
    ' XML DOCUMENT MANAGEMENT
    Private Sub AddXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
        root.AppendChild(resultsItem)
    End Sub

    Private Sub CreateXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
    End Sub

    Private Sub AddXMLAttribute(ByVal xmldoc As XmlDocument, _
        ByVal xmlnode As XmlElement, ByVal attribute As String, _
        ByVal attributevalue As String)
        ' Used to add an attribute to a specified node

        Dim newAtt As XmlAttribute

        newAtt = xmldoc.CreateAttribute(attribute)
        newAtt.Value = attributevalue
        xmlnode.Attributes.Append(newAtt)
    End Sub

    Private Function GetNodeValue(ByVal sNodeName As String, ByVal oParentNode As XmlNode) As String
        ' Generic function to return the value of a node in an XML document
        Dim oNode As XmlNode = oParentNode.SelectSingleNode(".//" + sNodeName)
        If oNode Is Nothing Then
            Return String.Empty
        Else
            Return oNode.InnerText
        End If
    End Function

    ' =================================================
    ' COLLECTIONS 
    ' This class implements a simple dictionary using an array of DictionaryEntry objects (key/value pairs).
    Public Class SimpleDictionary
        Implements IDictionary

        ' The array of items
        Dim items() As DictionaryEntry
        Dim ItemsInUse As Integer = 0

        ' Construct the SimpleDictionary with the desired number of items.
        ' The number of items cannot change for the life time of this SimpleDictionary.
        Public Sub New(ByVal numItems As Integer)
            items = New DictionaryEntry(numItems - 1) {}
        End Sub

        ' IDictionary Members
        Public ReadOnly Property IsReadOnly() As Boolean Implements IDictionary.IsReadOnly
            Get
                Return False
            End Get
        End Property

        Public Function Contains(ByVal key As Object) As Boolean Implements IDictionary.Contains
            Dim index As Integer
            Return TryGetIndexOfKey(key, index)
        End Function

        Public ReadOnly Property IsFixedSize() As Boolean Implements IDictionary.IsFixedSize
            Get
                Return False
            End Get
        End Property

        Public Sub Remove(ByVal key As Object) Implements IDictionary.Remove
            If key = Nothing Then
                Throw New ArgumentNullException("key")
            End If
            ' Try to find the key in the DictionaryEntry array
            Dim index As Integer
            If TryGetIndexOfKey(key, index) Then

                ' If the key is found, slide all the items up.
                Array.Copy(items, index + 1, items, index, (ItemsInUse - index) - 1)
                ItemsInUse = ItemsInUse - 1
            Else

                ' If the key is not in the dictionary, just return. 
            End If
        End Sub

        Public Sub Clear() Implements IDictionary.Clear
            ItemsInUse = 0
        End Sub

        Public Sub Add(ByVal key As Object, ByVal value As Object) Implements IDictionary.Add

            ' Add the new key/value pair even if this key already exists in the dictionary.
            If ItemsInUse = items.Length Then
                Throw New InvalidOperationException("The dictionary cannot hold any more items.")
            End If
            items(ItemsInUse) = New DictionaryEntry(key, value)
            ItemsInUse = ItemsInUse + 1
        End Sub

        Public ReadOnly Property Keys() As ICollection Implements IDictionary.Keys
            Get

                ' Return an array where each item is a key.
                ' Note: Declaring keyArray() to have a size of ItemsInUse - 1
                '       ensures that the array is properly sized, in VB.NET
                '       declaring an array of size N creates an array with
                '       0 through N elements, including N, as opposed to N - 1
                '       which is the default behavior in C# and C++.
                Dim keyArray() As Object = New Object(ItemsInUse - 1) {}
                Dim n As Integer
                For n = 0 To ItemsInUse - 1
                    keyArray(n) = items(n).Key
                Next n

                Return keyArray
            End Get
        End Property

        Public ReadOnly Property Values() As ICollection Implements IDictionary.Values
            Get
                ' Return an array where each item is a value.
                Dim valueArray() As Object = New Object(ItemsInUse - 1) {}
                Dim n As Integer
                For n = 0 To ItemsInUse - 1
                    valueArray(n) = items(n).Value
                Next n

                Return valueArray
            End Get
        End Property

        Default Public Property Item(ByVal key As Object) As Object Implements IDictionary.Item
            Get

                ' If this key is in the dictionary, return its value.
                Dim index As Integer
                If TryGetIndexOfKey(key, index) Then

                    ' The key was found return its value.
                    Return items(index).Value
                Else

                    ' The key was not found return null.
                    Return Nothing
                End If
            End Get

            Set(ByVal value As Object)
                ' If this key is in the dictionary, change its value. 
                Dim index As Integer
                If TryGetIndexOfKey(key, index) Then

                    ' The key was found change its value.
                    items(index).Value = value
                Else

                    ' This key is not in the dictionary add this key/value pair.
                    Add(key, value)
                End If
            End Set
        End Property

        Private Function TryGetIndexOfKey(ByVal key As Object, ByRef index As Integer) As Boolean
            For index = 0 To ItemsInUse - 1
                ' If the key is found, return true (the index is also returned).
                If items(index).Key.Equals(key) Then
                    Return True
                End If
            Next index

            ' Key not found, return false (index should be ignored by the caller).
            Return False
        End Function

        Private Class SimpleDictionaryEnumerator
            Implements IDictionaryEnumerator

            ' A copy of the SimpleDictionary object's key/value pairs.
            Dim items() As DictionaryEntry
            Dim index As Integer = -1

            Public Sub New(ByVal sd As SimpleDictionary)
                ' Make a copy of the dictionary entries currently in the SimpleDictionary object.
                items = New DictionaryEntry(sd.Count - 1) {}
                Array.Copy(sd.items, 0, items, 0, sd.Count)
            End Sub

            ' Return the current item.
            Public ReadOnly Property Current() As Object Implements IDictionaryEnumerator.Current
                Get
                    ValidateIndex()
                    Return items(index)
                End Get
            End Property

            ' Return the current dictionary entry.
            Public ReadOnly Property Entry() As DictionaryEntry Implements IDictionaryEnumerator.Entry
                Get
                    Return Current
                End Get
            End Property

            ' Return the key of the current item.
            Public ReadOnly Property Key() As Object Implements IDictionaryEnumerator.Key
                Get
                    ValidateIndex()
                    Return items(index).Key
                End Get
            End Property

            ' Return the value of the current item.
            Public ReadOnly Property Value() As Object Implements IDictionaryEnumerator.Value
                Get
                    ValidateIndex()
                    Return items(index).Value
                End Get
            End Property

            ' Advance to the next item.
            Public Function MoveNext() As Boolean Implements IDictionaryEnumerator.MoveNext
                If index < items.Length - 1 Then
                    index = index + 1
                    Return True
                End If

                Return False
            End Function

            ' Validate the enumeration index and throw an exception if the index is out of range.
            Private Sub ValidateIndex()
                If index < 0 Or index >= items.Length Then
                    Throw New InvalidOperationException("Enumerator is before or after the collection.")
                End If
            End Sub

            ' Reset the index to restart the enumeration.
            Public Sub Reset() Implements IDictionaryEnumerator.Reset
                index = -1
            End Sub

        End Class

        Public Function GetEnumerator() As IDictionaryEnumerator Implements IDictionary.GetEnumerator

            'Construct and return an enumerator.
            Return New SimpleDictionaryEnumerator(Me)
        End Function


        ' ICollection Members
        Public ReadOnly Property IsSynchronized() As Boolean Implements IDictionary.IsSynchronized
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property SyncRoot() As Object Implements IDictionary.SyncRoot
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Public ReadOnly Property Count() As Integer Implements IDictionary.Count
            Get
                Return ItemsInUse
            End Get
        End Property

        Public Sub CopyTo(ByVal array As Array, ByVal index As Integer) Implements IDictionary.CopyTo
            Throw New NotImplementedException()
        End Sub

        ' IEnumerable Members
        Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator

            ' Construct and return an enumerator.
            Return Me.GetEnumerator()
        End Function
    End Class

    ' =================================================
    ' STRING FUNCTIONS
    Public Shared Function Soundex(ByVal Word As String) As String
        Return Soundex(Word, 4)
    End Function

    Public Shared Function Soundex(ByVal Word As String, ByVal Length As Integer) As String
        ' Value to return
        Dim Value As String = ""
        ' Size of the word to process
        Dim Size As Integer = Word.Length
        ' Make sure the word is at least two characters in length
        If (Size > 1) Then
            ' Convert the word to all uppercase
            Word = Word.ToUpper()
            ' Conver to the word to a character array for faster processing
            Dim Chars() As Char = Word.ToCharArray()
            ' Buffer to build up with character codes
            Dim Buffer As New System.Text.StringBuilder
            Buffer.Length = 0
            ' The current and previous character codes
            Dim PrevCode As Integer = 0
            Dim CurrCode As Integer = 0
            ' Append the first character to the buffer
            Buffer.Append(Chars(0))
            ' Prepare variables for loop
            Dim i As Integer
            Dim LoopLimit As Integer = Size - 1
            ' Loop through all the characters and convert them to the proper character code
            For i = 1 To LoopLimit
                Select Case Chars(i)
                    Case "A", "E", "I", "O", "U", "H", "W", "Y"
                        CurrCode = 0
                    Case "B", "F", "P", "V"
                        CurrCode = 1
                    Case "C", "G", "J", "K", "Q", "S", "X", "Z"
                        CurrCode = 2
                    Case "D", "T"
                        CurrCode = 3
                    Case "L"
                        CurrCode = 4
                    Case "M", "N"
                        CurrCode = 5
                    Case "R"
                        CurrCode = 6
                End Select
                ' Check to see if the current code is the same as the last one
                If (CurrCode <> PrevCode) Then
                    ' Check to see if the current code is 0 (a vowel); do not proceed
                    If (CurrCode <> 0) Then
                        Buffer.Append(CurrCode)
                    End If
                End If
                ' If the buffer size meets the length limit, then exit the loop
                If (Buffer.Length = Length) Then
                    Exit For
                End If
            Next
            ' Padd the buffer if required
            Size = Buffer.Length
            If (Size < Length) Then
                Buffer.Append("0", (Length - Size))
            End If
            ' Set the return value
            Value = Buffer.ToString()
        End If
        ' Return the computed soundex
        Return Value
    End Function

    Public Function ReverseString(ByVal InputString As String) As String
        ' Reverses a string
        Dim lLen As Long, lCtr As Long
        Dim sChar As String
        Dim sAns As String
        sAns = ""
        lLen = Len(InputString)
        For lCtr = lLen To 1 Step -1
            sChar = Mid(InputString, lCtr, 1)
            sAns = sAns & sChar
        Next
        ReverseString = sAns
    End Function

    Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        ' Validate email address

        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If

    End Function

    Function SqlString(ByVal Instring As String) As String
        ' Make a string safe for use in a SQL query
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            SqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 1) = "'" Then
                outstring = outstring & "''"
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        SqlString = outstring
    End Function

    Function CheckNull(ByVal Instring As String) As String
        ' Check to see if a string is null
        If Instring Is Nothing Then
            CheckNull = ""
        Else
            CheckNull = Instring
        End If
    End Function

    Public Function CheckDBNull(ByVal obj As Object, _
    Optional ByVal ObjectType As enumObjectType = enumObjectType.StrType) As Object
        ' Checks an object to determine if its null, and if so sets it to a not-null empty value
        Dim objReturn As Object
        objReturn = obj
        If ObjectType = enumObjectType.StrType And IsDBNull(obj) Then
            objReturn = ""
        ElseIf ObjectType = enumObjectType.IntType And IsDBNull(obj) Then
            objReturn = 0
        ElseIf ObjectType = enumObjectType.DblType And IsDBNull(obj) Then
            objReturn = 0.0
        ElseIf ObjectType = enumObjectType.DteType And IsDBNull(obj) Then
            objReturn = Now
        End If
        Return objReturn
    End Function

    Public Function NumString(ByVal strString As String) As String
        ' Remove everything but numbers from a string
        Dim bln As Boolean
        Dim i As Integer
        Dim iv As String
        NumString = ""

        'Can array element be evaluated as a number?
        For i = 1 To Len(strString)
            iv = Mid(strString, i, 1)
            bln = IsNumeric(iv)
            If bln Then NumString = NumString & iv
        Next

    End Function

    Public Function ToBase64(ByVal data() As Byte) As String
        ' Encode a Base64 string
        If data Is Nothing Then Throw New ArgumentNullException("data")
        Return Convert.ToBase64String(data)
    End Function

    Public Function FromBase64(ByVal base64 As String) As String
        ' Decode a Base64 string
        Dim results As String
        If base64 Is Nothing Then Throw New ArgumentNullException("base64")
        results = System.Text.Encoding.ASCII.GetString(Convert.FromBase64String(base64))
        Return results
    End Function

    Function DeSqlString(ByVal Instring As String) As String
        ' Convert a string from SQL query encoded to non-encoded
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        CheckDBNull(Instring, enumObjectType.StrType)
        If Len(Instring) = 0 Then
            DeSqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 2) = "''" Then
                outstring = outstring & "'"
                i = i + 1
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        DeSqlString = outstring
    End Function

    Public Function StringToBytes(ByVal str As String) As Byte()
        ' Convert a random string to a byte array
        ' e.g. "abcdefg" to {a,b,c,d,e,f,g}
        Dim s As Char()
        s = str.ToCharArray
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            b(i) = Convert.ToByte(s(i))
        Next
        Return b
    End Function

    Public Function EncodeParamSpaces(ByVal InVal As String) As String
        ' If given a urlencoded parameter value, replace spaces with "+" signs

        Dim temp As String
        Dim i As Integer

        If InStr(InVal, " ") > 0 Then
            temp = ""
            For i = 1 To Len(InVal)
                If Mid(InVal, i, 1) = " " Then
                    temp = temp & "+"
                Else
                    temp = temp & Mid(InVal, i, 1)
                End If
            Next
            EncodeParamSpaces = temp
        Else
            EncodeParamSpaces = InVal
        End If
    End Function

    Public Function DecodeParamSpaces(ByVal InVal As String) As String
        ' If given an encoded parameter value, replace "+" signs with spaces

        Dim temp As String
        Dim i As Integer

        If InStr(InVal, "+") > 0 Then
            temp = ""
            For i = 1 To Len(InVal)
                If Mid(InVal, i, 1) = "+" Then
                    temp = temp & " "
                Else
                    temp = temp & Mid(InVal, i, 1)
                End If
            Next
            DecodeParamSpaces = temp
        Else
            DecodeParamSpaces = InVal
        End If
    End Function

    Public Function NumStringToBytes(ByVal str As String) As Byte()
        ' Convert a string containing numbers to a byte array
        ' e.g. "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16" to 
        '  {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16}
        Dim s As String()
        s = str.Split(" ")
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            b(i) = Convert.ToByte(s(i))
        Next
        Return b
    End Function

    Public Function BytesToString(ByVal b() As Byte) As String
        ' Convert a byte array to a string
        Dim i As Integer
        Dim s As New System.Text.StringBuilder()
        For i = 0 To b.Length - 1
            Console.WriteLine(b(i))
            If i <> b.Length - 1 Then
                s.Append(b(i) & " ")
            Else
                s.Append(b(i))
            End If
        Next
        Return s.ToString
    End Function

    ' =================================================
    ' DATABASE FUNCTIONS
    Public Function OpenDBConnection(ByVal ConnS As String, ByRef con As SqlConnection, ByRef cmd As SqlCommand) As String
        ' Function to open a database connection with extreme error-handling
        ' Returns an error message if unable to open the connection
        Dim SqlS As String
        SqlS = ""
        OpenDBConnection = ""

        Try
            con = New SqlConnection(ConnS)
            con.Open()
            If Not con Is Nothing Then
                Try
                    cmd = New SqlCommand(SqlS, con)
                    cmd.CommandTimeout = 300
                Catch ex2 As Exception
                    OpenDBConnection = "Error opening the command string: " & ex2.ToString
                End Try
            End If
        Catch ex As Exception
            If con.State <> Data.ConnectionState.Closed Then con.Dispose()
            ConnS = ConnS & ";Pooling=false"
            Try
                con = New SqlConnection(ConnS)
                con.Open()
                If Not con Is Nothing Then
                    Try
                        cmd = New SqlCommand(SqlS, con)
                        cmd.CommandTimeout = 300
                    Catch ex2 As Exception
                        OpenDBConnection = "Error opening the command string: " & ex2.ToString
                    End Try
                End If
            Catch ex2 As Exception
                OpenDBConnection = "Unable to open database connection for connection string: " & ConnS & vbCrLf & "Windows error: " & vbCrLf & ex2.ToString & vbCrLf
            End Try
        End Try

    End Function

    Public Function CloseDBConnection(ByRef con As SqlConnection, ByRef cmd As SqlCommand, ByRef dr As SqlDataReader) As String
        ' This function closes a database connection safely
        Dim ErrMsg As String
        ErrMsg = ""

        ' Handle datareader
        Try
            dr.Close()
        Catch ex As Exception
        End Try
        Try
            dr = Nothing
        Catch ex As Exception
        End Try

        ' Handle command
        Try
            cmd.Dispose()
        Catch ex As Exception
        End Try
        Try
            cmd = Nothing
        Catch ex As Exception
        End Try

        ' Handle connection
        Try
            con.Close()
        Catch ex As Exception
        End Try
        Try
            SqlConnection.ClearPool(con)
        Catch ex As Exception
        End Try
        Try
            con.Dispose()
        Catch ex As Exception
        End Try
        Try
            con = Nothing
        Catch ex As Exception
        End Try

        ' Exit
        Return ErrMsg
    End Function

    ' =================================================
    ' DEBUG FUNCTIONS
    Public Sub writeoutput(ByVal fs As StreamWriter, ByVal instring As String)
        ' This function writes a line to a previously opened streamwriter, and then flushes it
        ' promptly.  This assists in debugging services
        fs.WriteLine(instring)
        fs.Flush()
    End Sub

    Public Sub writeoutputfs(ByVal fs As FileStream, ByVal instring As String)
        ' This function writes a line to a previously opened filestream, and then flushes it
        ' promptly.  This assists in debugging services
        fs.Write(StringToBytes(instring), 0, Len(instring))
        fs.Write(StringToBytes(vbCrLf), 0, 2)
        fs.Flush()
    End Sub

End Class

Namespace Dms.AsynchronousOperations
    Public Class AsyncMain

        Enum enumObjectType
            StrType = 0
            IntType = 1
            DblType = 2
            DteType = 3
        End Enum

        Public Function UpdDMSDoc(ByVal CONTACT_ID As String, ByVal TRAINER_NUM As String, ByVal PART_ID As String, _
            ByVal MT_ID As String, ByVal Debug As String) As String

            ' This function generates a count of visible documents for the specific individual.  This is a local version
            ' of UpdDMSDocCount

            '   CONTACT_ID	- The "siebeldb.S_CONTACT.ROW_ID" of the individual (req)
            '   TRAINER_NUM	- The "siebeldb.S_CONTACT.X_TRAINER_NUM" of the individual (opt)
            '   PART_ID	- The "siebeldb.S_CONTACT.X_PART_ID" of the individual (opt)
            '   MT_ID	- The "siebeldb.S_CONTACT.REG_AS_EMP_ID" of the individual (opt)

            ' web.config Parameters used:
            '   dms        	    - connection string to DMS.dms database

            ' Variables
            Dim results, temp As String
            Dim iDoc As XmlDocument = New XmlDocument()
            Dim mypath, errmsg, logging As String
            Dim bResponse As Boolean
            Dim doc_count As String
            Dim SUB_ID, DOMAIN, USER_AID, SUB_AID, DOMAIN_AID, UID As String
            Dim Category_Constraint, TRAINER_FLG, MT_FLG, PART_FLG, TRAINING_FLG, TRAINER_ACC_FLG, SITE_ONLY, SYSADMIN_FLG, EMP_ID As String

            ' Database declarations
            Dim SqlS As String
            Dim returnv As Integer

            ' HCIDB Database declarations
            Dim con As SqlConnection
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            Dim ConnS As String

            ' DMS Database declarations
            Dim d_con As SqlConnection
            Dim d_cmd As SqlCommand
            Dim d_dr As SqlDataReader
            Dim d_ConnS As String

            ' Logging declarations
            Dim ltemp As String
            Dim myeventlog As log4net.ILog
            Dim mydebuglog As log4net.ILog
            myeventlog = log4net.LogManager.GetLogger("EventLog")
            mydebuglog = log4net.LogManager.GetLogger("UDDDebugLog")
            Dim logfile As String
            Dim LogStartTime As String = Now.ToString
            Dim VersionNum As String = "100"

            ' Web service declarations
            Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

            ' ============================================
            ' Variable setup
            mypath = HttpRuntime.AppDomainAppPath
            logging = "Y"
            temp = ""
            errmsg = ""
            bResponse = False
            doc_count = "0"
            results = "Success"
            SUB_ID = ""
            DOMAIN = ""
            USER_AID = ""
            SUB_AID = ""
            DOMAIN_AID = ""
            UID = ""
            TRAINER_FLG = ""
            MT_FLG = ""
            PART_FLG = ""
            TRAINING_FLG = ""
            TRAINER_ACC_FLG = ""
            SITE_ONLY = ""
            SYSADMIN_FLG = ""
            EMP_ID = ""

            ' ============================================
            ' Fix parameters
            Debug = UCase(Left(Debug, 1))
            If Debug = "" Then Debug = "N"
            'Debug = "Y"
            If Debug = "T" Then
                CONTACT_ID = "21120611WE0"
                PART_ID = "732632"
                MT_ID = ""
                TRAINER_NUM = "22"
            Else
                CONTACT_ID = Trim(HttpUtility.UrlEncode(CONTACT_ID))
                If InStr(CONTACT_ID, "%") > 0 Then CONTACT_ID = Trim(HttpUtility.UrlDecode(CONTACT_ID))
                CONTACT_ID = Trim(EncodeParamSpaces(CONTACT_ID))
                PART_ID = Trim(HttpUtility.UrlEncode(PART_ID))
                If InStr(PART_ID, "%") > 0 Then PART_ID = Trim(HttpUtility.UrlDecode(PART_ID))
                PART_ID = Trim(EncodeParamSpaces(PART_ID))
                MT_ID = Trim(HttpUtility.UrlEncode(MT_ID))
                If InStr(MT_ID, "%") > 0 Then MT_ID = Trim(HttpUtility.UrlDecode(MT_ID))
                MT_ID = Trim(EncodeParamSpaces(MT_ID))
            End If

            ' ============================================
            ' Get system defaults
            ' hcidb1
            Try
                ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("siebeldb").ConnectionString
                If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
                temp = System.Configuration.ConfigurationManager.AppSettings.Get("UpdDMSDoc_debug")
                If temp = "Y" And Debug <> "T" Then Debug = "Y"
            Catch ex As Exception
                errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
                results = "Failure"
                GoTo CloseOut2
            End Try
            ' dms
            Try
                d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
                If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;database=DMS"
            Catch ex As Exception
                errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
                results = "Failure"
                GoTo CloseOut2
            End Try

            ' ============================================
            ' Open log file if applicable
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                logfile = "C:\Logs\UpdDMSDoc.log"
                Try
                    log4net.GlobalContext.Properties("UDDDLogFileName") = logfile
                    log4net.Config.XmlConfigurator.Configure()
                Catch ex As Exception
                    errmsg = errmsg & "Error Opening Log. " & vbCrLf
                    results = "Failure"
                    GoTo CloseOut2
                End Try

                If Debug = "Y" Then
                    mydebuglog.Debug("----------------------------------")
                    mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                    mydebuglog.Debug("Parameters-")
                    mydebuglog.Debug("  CONTACT_ID: " & CONTACT_ID)
                    mydebuglog.Debug("  PART_ID: " & PART_ID)
                    mydebuglog.Debug("  MT_ID: " & MT_ID)
                    mydebuglog.Debug("  TRAINER_NUM: " & TRAINER_NUM)
                End If
            End If

            ' ============================================
            ' Check required parameters
            If (CONTACT_ID = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "Invalid parameter(s) " & vbCrLf
                GoTo CloseOut2
            End If

            ' ============================================
            ' Open database connections
            errmsg = OpenDBConnection(ConnS, con, cmd)
            If errmsg <> "" Then
                results = "Failure"
                GoTo CloseOut
            End If
            errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
            If errmsg <> "" Then
                results = "Failure"
                GoTo CloseOut
            End If

            ' ============================================
            ' Get Subscription and Domain Info
            'SqlS = "SELECT S.ROW_ID, S.DOMAIN, C.X_REGISTRATION_NUM " & _
            '"FROM siebeldb.dbo.CX_SUB_CON SC " & _
            '"INNER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID " & _
            '"INNER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=SC.CON_ID " & _
            '"WHERE SC.CON_ID='" & CONTACT_ID & "'"
            SqlS = "SELECT S.ROW_ID, S.DOMAIN, C.X_REGISTRATION_NUM, C.X_TRAINER_FLG, C.X_MAST_TRNR_FLG, " & _
            "(SELECT CASE WHEN C.X_PART_ID IS NOT NULL AND C.X_PART_ID<>'' THEN 'Y' ELSE 'N' END) AS PART_FLG, S.SVC_TYPE, " & _
            "SC.TRAINER_ACC_FLG, SC.SITE_ONLY_FLG, SC.SYSADMIN_FLG, E.ROW_ID " & _
            "FROM siebeldb.dbo.CX_SUB_CON SC  " & _
            "INNER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID  " & _
            "INNER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=SC.CON_ID  " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_EMPLOYEE E ON E.X_CON_ID=C.ROW_ID AND E.CNTRCTR_EMPLR_ID IS NULL  " & _
            "WHERE SC.CON_ID='" & CONTACT_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  Get subscription info: " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            SUB_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                            DOMAIN = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                            UID = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                            TRAINER_FLG = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                            MT_FLG = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                            PART_FLG = Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString
                            TRAINING_FLG = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                            Select Case TRAINING_FLG.ToUpper()
                                Case "CERTIFICATION MANAGER REG DB"
                                    TRAINING_FLG = "N"
                                Case "CERTIFICATION MANAGER REPORTS"
                                    TRAINING_FLG = "N"
                                Case Else
                                    TRAINING_FLG = "Y"
                            End Select
                            TRAINER_ACC_FLG = Trim(CheckDBNull(dr(7), enumObjectType.StrType)).ToString
                            SITE_ONLY = Trim(CheckDBNull(dr(8), enumObjectType.StrType)).ToString
                            SYSADMIN_FLG = Trim(CheckDBNull(dr(9), enumObjectType.StrType)).ToString
                            EMP_ID = Trim(CheckDBNull(dr(10), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                            'results = "Failure"
                            'errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Error getting document count. " & vbCrLf
                    results = "Failure"
                End If
                Try
                    dr.Close()
                    dr = Nothing
                Catch ex As Exception
                End Try
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
                GoTo CloseOut
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug("      > Sub_Id/Domain/UID: " & SUB_ID & "/" & DOMAIN & "/" & UID)
                mydebuglog.Debug("      > TRAINER_FLG:" & TRAINER_FLG)
                mydebuglog.Debug("      > MT_FLG: " & MT_FLG)
                mydebuglog.Debug("      > PART_FLG: " & PART_FLG)
                mydebuglog.Debug("      > TRAINING_FLG: " & TRAINING_FLG)
                mydebuglog.Debug("      > TRAINER_ACC_FLG: " & TRAINER_ACC_FLG)
                mydebuglog.Debug("      > SITE_ONLY: " & SITE_ONLY)
                mydebuglog.Debug("      > SYSADMIN_FLG: " & SYSADMIN_FLG)
                mydebuglog.Debug("      > EMP_ID: " & EMP_ID)
            End If

            ' If no subscription, no point
            If SUB_ID = "" Or UID = "" Then
                If Debug = "Y" Then mydebuglog.Debug("No subscription to update for UID '" & UID & "' and SUB_ID '" & SUB_ID & "'")
                GoTo CloseOut
            End If

            ' ============================================
            ' Get DMS security information
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get DMS security information")

            ' -----
            ' User AID
            SqlS = "SELECT UA.row_id " & _
            "FROM DMS.dbo.User_Group_Access UA " & _
            "INNER JOIN DMS.dbo.Users U ON U.row_id=UA.access_id " & _
            "WHERE UA.type_id='U' AND U.ext_user_id='" & CONTACT_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  .. Get user security: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            USER_AID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                        End Try
                    End While
                End If
                Try
                    d_dr.Close()
                    d_dr = Nothing
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug("      > USER_AID: " & USER_AID)

            ' -----
            ' Subscription AID
            SqlS = "SELECT UA.row_id " & _
            "FROM DMS.dbo.User_Group_Access UA " & _
            "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " & _
            "WHERE UA.type_id='G' AND G.name='" & SUB_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  .. Get subscription security: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            SUB_AID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                        End Try
                    End While
                End If
                Try
                    d_dr.Close()
                    d_dr = Nothing
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug("      > SUB_AID: " & SUB_AID)

            ' -----
            ' Domain AID
            SqlS = "SELECT UA.row_id " & _
            "FROM DMS.dbo.User_Group_Access UA " & _
            "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " & _
            "WHERE UA.type_id='G' AND G.name='" & DOMAIN & "'"
            If Debug = "Y" Then mydebuglog.Debug("  .. Get domain security: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            DOMAIN_AID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                        End Try
                    End While
                End If
                Try
                    d_dr.Close()
                    d_dr = Nothing
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug("      > DOMAIN_AID: " & DOMAIN_AID)

            ' ============================================
            ' Generate Category Constraint
            Category_Constraint = "CK.key_id IN ("
            If TRAINER_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "3,"
            End If
            If MT_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "5,"
            End If
            If PART_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "7,"
            End If
            If TRAINING_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "8,"
            End If
            If TRAINER_ACC_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "10,"
            End If
            If SITE_ONLY = "Y" Then
                Category_Constraint = Category_Constraint & "12,"
            End If
            Category_Constraint = Category_Constraint & "13,"
            If SYSADMIN_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "15,"
            End If
            If EMP_ID <> "" Then
                Category_Constraint = Category_Constraint & "16,"
            End If
            Category_Constraint = Category_Constraint & "14) "
            If Debug = "Y" Then mydebuglog.Debug("  Category_Constraint: " & Category_Constraint)

            ' ============================================
            ' Get current document count if the user has a subscription
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Generate document count")
            SqlS = "SELECT count(1) AS NUM_DOC " & _
            "FROM (" & _
            "SELECT D.row_id " & _
            "FROM DMS.dbo.Documents D " & _
            "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " & _
            "LEFT OUTER JOIN DMS.dbo.Category_Keywords CK ON CK.cat_id=DC.cat_id " & _
            "WHERE DC.pr_flag='Y' AND (CK.key_id IN (3,5,7,8,13,15,16,14)) " & _
            "GROUP BY D.row_id " & _
            "INTERSECT " & _
            "SELECT DISTINCT DA.doc_id " & _
            "FROM DMS.dbo.Document_Associations DA " & _
            "INNER JOIN DMS.dbo.Documents D on D.row_id=DA.doc_id " & _
            "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id "
            SqlS = SqlS & "WHERE ((DA.association_id='3' AND DA.fkey='" & CONTACT_ID & "' AND DA.pr_flag='Y') or "
            If TRAINER_NUM <> "" Then SqlS = SqlS & "(DA.association_id='5' AND DA.fkey='" & TRAINER_NUM & "' AND DA.pr_flag='Y') or "
            If PART_ID <> "" Then SqlS = SqlS & "(DA.association_id='4' AND DA.fkey='" & PART_ID & "' AND DA.pr_flag='Y') or "
            If MT_ID <> "" Then SqlS = SqlS & "(DA.association_id='37' AND DA.fkey='" & MT_ID & "' AND DA.pr_flag='Y') or "
            SqlS = Left(SqlS, Len(SqlS) - 4) & ") AND D.deleted IS NULL AND ("
            If USER_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & USER_AID & " OR "
            If SUB_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & SUB_AID & " OR "
            If DOMAIN_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & DOMAIN_AID & " OR "
            SqlS = Left(SqlS, Len(SqlS) - 4) & ") GROUP BY DA.doc_id ) d "
            If Debug = "Y" Then mydebuglog.Debug("  .. Get document count: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            doc_count = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Error getting document count. " & vbCrLf
                    results = "Failure"
                End If
                d_dr.Close()
                d_dr = Nothing
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug("      > doc_count: " & doc_count)

            ' -----
            ' Drop temp table
            'SqlS = "DROP TABLE DMS.dbo.[" & UID & "]"
            'If Debug = "Y" Then mydebuglog.Debug("  .. Drop temp table for count: " & SqlS)
            'Try
            'd_cmd.CommandText = SqlS
            'returnv = d_cmd.ExecuteNonQuery()
            'Catch ex As Exception
            'End Try

            ' -----
            ' Get list of visible documents and store in temp table
            'SqlS = "SELECT DISTINCT DA.doc_id " & _
            '"INTO DMS.dbo.[" & UID & "] " & _
            '"FROM DMS.dbo.Document_Associations DA " & _
            '"INNER JOIN DMS.dbo.Documents D on D.row_id=DA.doc_id " & _
            '"INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id " & _
            '"WHERE ((DA.association_id='3' AND DA.fkey='" & CONTACT_ID & "' AND DA.pr_flag='Y') or "
            'If TRAINER_NUM <> "" Then SqlS = SqlS & "(DA.association_id='5' AND DA.fkey='" & TRAINER_NUM & "' AND DA.pr_flag='Y') or "
            'If PART_ID <> "" Then SqlS = SqlS & "(DA.association_id='4' AND DA.fkey='" & PART_ID & "' AND DA.pr_flag='Y') or "
            'If MT_ID <> "" Then SqlS = SqlS & "(DA.association_id='37' AND DA.fkey='" & MT_ID & "' AND DA.pr_flag='Y') or "
            'SqlS = Left(SqlS, Len(SqlS) - 4) & ") AND D.deleted IS NULL AND ("
            'If USER_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & USER_AID & " OR "
            'If SUB_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & SUB_AID & " OR "
            'If DOMAIN_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & DOMAIN_AID & " OR "
            'SqlS = Left(SqlS, Len(SqlS) - 4) & ")"
            'SqlS = SqlS & " GROUP BY DA.doc_id"
            'If Debug = "Y" Then mydebuglog.Debug("  .. Create temp table for count: " & SqlS)
            'Try
            'd_cmd.CommandText = SqlS
            'returnv = d_cmd.ExecuteNonQuery()
            'Catch ex As Exception
            'End Try

            ' -----
            ' Get count of documents in temp table
            'SqlS = "SELECT COUNT(*) AS NUM_DOC " & _
            '"FROM DMS.dbo.[" & UID & "]"
            'If Debug = "Y" Then mydebuglog.Debug("  .. Get document count from temp table: " & SqlS)
            'Try
            'd_cmd.CommandText = SqlS
            'd_dr = d_cmd.ExecuteReader()
            'If Not d_dr Is Nothing Then
            'While d_dr.Read()
            'Try
            'doc_count = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
            'Catch ex As Exception
            'results = "Failure"
            'errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
            'GoTo CloseOut
            'End Try
            'End While
            'Else
            'errmsg = errmsg & "Error getting document count. " & vbCrLf
            'results = "Failure"
            'End If
            'd_dr.Close()
            'Catch ex As Exception
            'End Try
            'If Debug = "Y" Then mydebuglog.Debug("      > doc_count: " & doc_count)

            ' -----
            ' Drop temp table
            'SqlS = "DROP TABLE DMS.dbo.[" & UID & "]"
            'If Debug = "Y" Then mydebuglog.Debug("  .. Drop temp table for count: " & SqlS)
            'Try
            'd_cmd.CommandText = SqlS
            'returnv = d_cmd.ExecuteNonQuery()
            'Catch ex As Exception
            'End Try

            ' -----
            ' Update CX_SUB_CON.NEW_DOC with document count if applicable
            If doc_count <> "" Then
                SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
                "SET NEW_DOC=" & doc_count & _
                " WHERE CON_ID='" & CONTACT_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug("  .. Update contact document count in CX_SUB_CON: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    returnv = cmd.ExecuteNonQuery()
                    If returnv = 0 Then results = "Failure"
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error setting the document count. " & ex.ToString & vbCrLf
                End Try

                SqlS = "UPDATE siebeldb.dbo.S_CONTACT " & _
                "SET DCKING_NUM=" & doc_count & " " & _
                "WHERE ROW_ID='" & CONTACT_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug("  .. Update contact document count in S_CONTACT: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    returnv = cmd.ExecuteNonQuery()
                    If returnv = 0 Then results = "Failure"
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error setting the document count. " & ex.ToString & vbCrLf
                End Try
            End If

CloseOut:
            ' ============================================
            ' Close database connections and objects
            Try
                errmsg = errmsg & CloseDBConnection(con, cmd, dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the hcidb database connection. " & vbCrLf
            End Try
            Try
                errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
            End Try

CloseOut2:
            ' ============================================
            ' Close the log file if any
            ltemp = results & " : Contact id " & CONTACT_ID & " has " & doc_count & " documents"
            If Trim(errmsg) <> "" Then myeventlog.Error("UpdDMSDoc : Error: " & Trim(errmsg))
            myeventlog.Info("UpdDMSDoc : Results: " & ltemp)
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                Try
                    If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                    If Debug = "Y" Then
                        mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                        mydebuglog.Debug("----------------------------------")
                    Else
                        mydebuglog.Debug("  Results: " & ltemp)
                    End If
                Catch ex As Exception
                End Try
            End If

            ' ============================================
            ' Log Performance Data
            If Debug <> "T" Then
                ' Send the web request
                Try
                    LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
                Catch ex As Exception
                End Try
                If results = "Success" Then results = doc_count
            End If

            ' ============================================
            ' Return results        
            Return results

        End Function

        Public Function SaveDMSDocAssoc(ByVal DocId As String, ByVal Association As String, _
            ByVal AssocKey As String, ByVal PrFlag As String, ByVal ReqdFlag As String, ByVal Rights As String, _
            ByVal Debug As String) As Boolean

            ' This function creates a Document_Associations record for the document and association specified

            ' The input parameters are as follows:
            '
            '   DocId   	- The "DMS.Documents.row_id" of the document (req.)
            '   Association	- The "DMS.Associations.name" of the item to be stored. (req.)
            '   AssocKey    - The "DMS.Document_Associations.fkey" of the record to be created (req.)
            '   PrFlag      - The "DMS.Document_Associations.pr_flag" of the record to be created (req.)
            '   ReqdFlag    - The "DMS.Document_Associations.reqd_flag" of the record to be created (req.)
            '   Rights      - The "DMS.Document_Associations.access_type" of the record to be created (req.)
            '                   Currently, this translates into the "access_flag" setting
            '   Debug	    - The debug mode flag: "Y", "N" or "T" 
            '
            ' The results are as follows:
            '
            '   DocAssocId    - The "DMS.Document_Association.row_id" of the record created

            ' web.config Parameters used:
            '   dms        	    - connection string to DMS.dms database

            ' Variables
            Dim temp As String
            Dim results As Boolean
            Dim iDoc As XmlDocument = New XmlDocument()
            Dim mypath, errmsg, logging As String

            ' Database declarations
            Dim SqlS As String
            Dim returnv As Integer

            ' DMS Database declarations
            Dim d_con As SqlConnection
            Dim d_cmd As SqlCommand
            Dim d_dr As SqlDataReader
            Dim d_ConnS As String

            ' Logging declarations
            Dim ltemp As String
            Dim myeventlog As log4net.ILog
            Dim mydebuglog As log4net.ILog
            myeventlog = log4net.LogManager.GetLogger("EventLog")
            mydebuglog = log4net.LogManager.GetLogger("SDDADebugLog")
            Dim logfile As String
            Dim LogStartTime As String = Now.ToString
            Dim VersionNum As String = "100"

            ' Web service declarations
            Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

            ' Local Cache declarations
            Dim DMSCache As New CachingWrapper.LocalCache

            ' Association declarations
            Dim AssocId, DocAssocId, AssocAccess As String

            ' ============================================
            ' Variable setup
            mypath = HttpRuntime.AppDomainAppPath
            logging = "Y"
            errmsg = ""
            results = False
            SqlS = ""
            returnv = 0
            AssocId = ""
            DocAssocId = ""
            AssocAccess = ""
            temp = ""

            ' ============================================
            ' Get parameters
            Debug = UCase(Left(Debug, 1))
            If Debug = "T" Then
                DocId = "1"
                Association = "Account"
                Rights = "Y"
                AssocKey = "KEY"
                PrFlag = "N"
                ReqdFlag = "N"
            Else
                DocId = Trim(HttpUtility.UrlEncode(DocId))
                If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))

                Association = Trim(HttpUtility.UrlDecode(Association))
                If InStr(Association, "%") > 0 Then Association = Trim(HttpUtility.UrlDecode(Association))

                If InStr(AssocKey, "%") > 0 Then AssocKey = Trim(HttpUtility.UrlDecode(AssocKey))
                If InStr(AssocKey, " ") > 0 Then AssocKey = EncodeParamSpaces(AssocKey)

                Rights = Trim(HttpUtility.UrlDecode(Rights))
                If InStr(Rights, "%") > 0 Then Rights = Trim(HttpUtility.UrlDecode(Rights))

                PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
                If InStr(PrFlag, "%") > 0 Then PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))

                ReqdFlag = Trim(HttpUtility.UrlDecode(ReqdFlag))
                If InStr(ReqdFlag, "%") > 0 Then ReqdFlag = Trim(HttpUtility.UrlDecode(ReqdFlag))
            End If

            ' ============================================
            ' Get system defaults
            Try
                d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
                If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
                temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDocAssoc_debug")
                If temp = "Y" Then Debug = "Y"
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            ' ============================================
            ' Open log file if applicable
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                logfile = "C:\Logs\SaveDMSDocAssoc.log"
                Try
                    log4net.GlobalContext.Properties("SDDALogFileName") = logfile
                    log4net.Config.XmlConfigurator.Configure()
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error Opening Log. "
                    results = "Failure"
                    GoTo CloseOut2
                End Try

                If Debug = "Y" Then
                    mydebuglog.Debug("----------------------------------")
                    mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                    mydebuglog.Debug("Parameters-")
                    mydebuglog.Debug("  DocId: " & DocId)
                    mydebuglog.Debug("  Association: " & Association)
                    mydebuglog.Debug("  AssocKey: " & AssocKey)
                    mydebuglog.Debug("  Rights: " & Rights)
                    mydebuglog.Debug("  PrFlag: " & PrFlag)
                    mydebuglog.Debug("  ReqdFlag: " & ReqdFlag)
                End If
            End If

            ' ============================================
            ' Validate Parameters
            If Trim(Association) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No association specified. "
                GoTo CloseOut2
            End If
            If Trim(DocId) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No document specified. "
                GoTo CloseOut2
            End If
            If Trim(AssocKey) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No association key specified. "
                GoTo CloseOut2
            End If
            If Trim(Rights) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No association rights specified. "
                GoTo CloseOut2
            End If
            If Trim(PrFlag) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No association primary flag specified. "
                GoTo CloseOut2
            End If
            If Trim(ReqdFlag) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No association required flag specified. "
                GoTo CloseOut2
            End If

            ' ============================================
            ' Open SQL Server database connections
            errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
            If errmsg <> "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not connect to the database. "
                GoTo CloseOut
            End If

            ' ============================================
            ' Load Associations from and/or into cache
            Dim dt As DataTable = New DataTable
            If Not DMSCache.GetCachedItem("Associations") Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Associations found in cache")
                Try
                    dt = DMSCache.GetCachedItem("Associations")
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                    GoTo CloseOut
                End Try
            Else
                SqlS = "SELECT name, row_id FROM DMS.dbo.Associations ORDER BY name"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Associations into cache: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If d_dr.HasRows Then
                        dt.Load(d_dr)
                        DMSCache.AddToCache("Associations", dt, CachingWrapper.CachePriority.NotRemovable)
                    End If
                    d_dr.Close()
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                    GoTo CloseOut
                End Try
            End If
            If dt Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not retrieve associations. "
                GoTo CloseOut
            End If

            ' Debug output
            If Debug = "Y" Then
                mydebuglog.Debug(" Associations Columns found: " & dt.Columns.Count.ToString)
                mydebuglog.Debug(" Associations Rows found: " & dt.Rows.Count.ToString)
            End If

            ' ============================================
            ' Locate Association Id in datatable
            Dim dr() As DataRow = dt.Select("name='" & Association & "'")
            If dr Is Nothing Or dr.Length = 0 Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Association not found")
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not find association. "
                GoTo CloseOut
            End If
            AssocId = dr(0).Item("row_id").ToString
            If Debug = "Y" Then
                mydebuglog.Debug(" Association row_id: " & AssocId)
            End If

            ' ============================================
            ' Close data objects
            Try
                dr = Nothing
                dt = Nothing
                DMSCache = Nothing
            Catch ex As Exception
            End Try

            ' ============================================
            ' Concert rights to access_flag as needed
            If Rights = "Y" Or Rights = "N" Then AssocAccess = Rights

            ' ============================================
            ' Write Document Association record
            If AssocAccess <> "" Then
                SqlS = "INSERT INTO DMS.dbo.Document_Associations " & _
                    "(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                    "VALUES (1, 1, " & AssocId & ", " & DocId & ", '" & AssocKey & "', '" & PrFlag & "', '" & AssocAccess & "', '" & ReqdFlag & "')"
            Else
                SqlS = "INSERT INTO DMS.dbo.Document_Associations " & _
                    "(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_type, reqd_flag) " & _
                    "VALUES (1, 1, " & AssocId & ", " & DocId & ", '" & AssocKey & "', '" & PrFlag & "', '" & Rights & "', '" & ReqdFlag & "')"
            End If
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Contact: " & vbCrLf & SqlS)
            If Debug <> "T" Then
                d_cmd.CommandText = SqlS
                Try
                    returnv = d_cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End If
            results = True

CloseOut:
            ' ============================================
            ' Close database connections and objects
            Try
                errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
            End Try

CloseOut2:
            ' ============================================
            ' Close the log file if any
            ltemp = results & " for association " & Association & ", with key '" & AssocKey & "' and document " & DocId
            If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDocAssoc :  Error: " & Trim(errmsg))
            myeventlog.Info("SaveDMSDocAssoc : Results: " & ltemp)
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                Try
                    If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                    mydebuglog.Debug("Results: " & ltemp)
                    If Debug = "Y" Then
                        mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                        mydebuglog.Debug("----------------------------------")
                    End If
                Catch ex As Exception
                End Try
            End If

            ' Log Performance Data
            If Debug <> "T" Then
                ' ============================================
                ' Send the web request
                Try
                    LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
                Catch ex As Exception
                End Try
            End If

            ' ============================================
            ' Return results
            Return results
        End Function

        Public Function SaveDMSDocCat(ByVal DocId As String, ByVal Category As String, ByVal PrFlag As String, _
            ByVal Debug As String) As Boolean

            ' This function creates a Document_Categories record for the document and category specified

            ' The input parameters are as follows:
            '
            '   DocId   	- The "DMS.Documents.row_id" of the document 
            '   Category	- The "DMS.Categories.name" of the item to be stored. (req.)
            '   PrFlag      - The primary category flag for the record to be created
            '   Debug	    - The debug mode flag: "Y", "N" or "T" 
            '
            ' The results are as follows:
            '
            '   Boolean     - True/False

            ' web.config Parameters used:
            '   dms        	    - connection string to DMS.dms database

            ' Variables
            Dim temp As String
            Dim results As Boolean
            Dim iDoc As XmlDocument = New XmlDocument()
            Dim mypath, errmsg, logging As String

            ' Database declarations
            Dim SqlS As String
            Dim returnv As Integer

            ' DMS Database declarations
            Dim d_con As SqlConnection
            Dim d_cmd As SqlCommand
            Dim d_dr As SqlDataReader
            Dim d_ConnS As String

            ' Logging declarations
            Dim ltemp As String
            Dim myeventlog As log4net.ILog
            Dim mydebuglog As log4net.ILog
            myeventlog = log4net.LogManager.GetLogger("EventLog")
            mydebuglog = log4net.LogManager.GetLogger("SDDCDebugLog")
            Dim logfile As String
            Dim LogStartTime As String = Now.ToString
            Dim VersionNum As String = "100"

            ' Web service declarations
            Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

            ' Local Cache declarations
            Dim DMSCache As New CachingWrapper.LocalCache

            ' Category declarations
            Dim CatId, DocCatId As String

            ' ============================================
            ' Variable setup
            mypath = HttpRuntime.AppDomainAppPath
            logging = "Y"
            errmsg = ""
            results = False
            SqlS = ""
            returnv = 0
            CatId = ""
            DocCatId = ""
            temp = ""

            ' ============================================
            ' Get parameters
            Debug = UCase(Left(Debug, 1))
            If Debug = "T" Then
                DocId = "1"
                Category = "Account"
                PrFlag = "N"
            Else
                DocId = Trim(HttpUtility.UrlEncode(DocId))
                If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))

                Category = Trim(HttpUtility.UrlDecode(Category))
                If InStr(Category, "%") > 0 Then Category = Trim(HttpUtility.UrlDecode(Category))

                PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
                If InStr(PrFlag, "%") > 0 Then PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
            End If

            ' ============================================
            ' Get system defaults
            Try
                d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
                If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
                temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDocCat_debug")
                If temp = "Y" Then Debug = "Y"
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            ' ============================================
            ' Open log file if applicable
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                logfile = "C:\Logs\SaveDMSDocCat.log"
                Try
                    log4net.GlobalContext.Properties("SDDCLogFileName") = logfile
                    log4net.Config.XmlConfigurator.Configure()
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error Opening Log. "
                    results = "Failure"
                    GoTo CloseOut2
                End Try

                If Debug = "Y" Then
                    mydebuglog.Debug("----------------------------------")
                    mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                    mydebuglog.Debug("Parameters-")
                    mydebuglog.Debug("  DocId: " & DocId)
                    mydebuglog.Debug("  Category: " & Category)
                    mydebuglog.Debug("  PrFlag: " & PrFlag)
                End If
            End If

            ' ============================================
            ' Validate Parameters
            If Trim(Category) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No category specified. "
                GoTo CloseOut2
            End If
            If Trim(DocId) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No document specified. "
                GoTo CloseOut2
            End If
            If PrFlag = "" Then PrFlag = "N"

            ' ============================================
            ' Open SQL Server database connections
            errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
            If errmsg <> "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not connect to the database. "
                GoTo CloseOut
            End If

            ' ============================================
            ' Load Categories from and/or into cache
            Dim dt As DataTable = New DataTable
            If Not DMSCache.GetCachedItem("Categories") Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Categories found in cache")
                Try
                    dt = DMSCache.GetCachedItem("Categories")
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                    GoTo CloseOut
                End Try
            Else
                SqlS = "SELECT name, row_id FROM DMS.dbo.Categories ORDER BY name"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Categories into cache: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If d_dr.HasRows Then
                        dt.Load(d_dr)
                        DMSCache.AddToCache("Categories", dt, CachingWrapper.CachePriority.NotRemovable)
                    End If
                    d_dr.Close()
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                    GoTo CloseOut
                End Try
            End If
            If dt Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not retrieve categories. "
                GoTo CloseOut
            End If

            ' Debug output
            If Debug = "Y" Then
                mydebuglog.Debug(" Categories Columns found: " & dt.Columns.Count.ToString)
                mydebuglog.Debug(" Categories Rows found: " & dt.Rows.Count.ToString)
            End If

            ' ============================================
            ' Locate Category Id in datatable
            Dim dr() As DataRow = dt.Select("name='" & Category & "'")
            If dr Is Nothing Or dr.Length = 0 Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Category not found")
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not find category. "
                GoTo CloseOut
            End If
            CatId = dr(0).Item("row_id").ToString
            If Debug = "Y" Then
                mydebuglog.Debug(" Category row_id: " & CatId)
            End If

            ' ============================================
            ' Close data objects
            Try
                dr = Nothing
                dt = Nothing
                DMSCache = Nothing
            Catch ex As Exception
            End Try

            ' ============================================
            ' Write Document_Categories record
            SqlS = "INSERT INTO DMS.dbo.Document_Categories " & _
                    "(created_by, last_upd_by, doc_id, cat_id, pr_flag) " & _
                    "VALUES (1, 1, " & DocId & ", " & CatId & ", '" & PrFlag & "')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Categories record for document: " & vbCrLf & SqlS)
            If Debug <> "T" Then
                d_cmd.CommandText = SqlS
                Try
                    returnv = d_cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End If
            results = True

CloseOut:
            ' ============================================
            ' Close database connections and objects
            Try
                errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
            End Try

CloseOut2:
            ' ============================================
            ' Close the log file if any
            ltemp = results & " for category " & Category & " and document " & DocId
            If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDocCat :  Error: " & Trim(errmsg))
            myeventlog.Info("SaveDMSDocCat : Results: " & ltemp)
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                Try
                    If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                    mydebuglog.Debug("Results: " & ltemp)
                    If Debug = "Y" Then
                        mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                        mydebuglog.Debug("----------------------------------")
                    End If
                Catch ex As Exception
                End Try
            End If

            ' Log Performance Data
            If Debug <> "T" Then
                ' ============================================
                ' Send the web request
                Try
                    LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
                Catch ex As Exception
                End Try
            End If

            ' ============================================
            ' Return results
            Return results
        End Function

        Public Function SaveDMSDocKey(ByVal DocId As String, ByVal DocKey As String, ByVal KeyVal As String, _
        ByVal PrFlag As String, ByVal Debug As String) As Boolean

            ' This function creates a Document_Keywords record for the document and Keyword specified

            ' The input parameters are as follows:
            '
            '   DocId       - The "DMS.Documents.row_id" of the document (req.)
            '   DocKey      - The "DMS.Keywords.name" of the item to be stored. (req.)
            '   KeyVal      - The "DMS.Document_Keywords.val" of the keyword to be created (opt.)
            '   PrFlag      - The primary DocKey flag for the record to be created
            '   Debug       - The debug mode flag: "Y", "N" or "T" 
            '
            ' The results are as follows:
            '
            '   Boolean     - True/False

            ' web.config Parameters used:
            '   dms        	    - connection string to DMS.dms database

            ' Variables
            Dim temp As String
            Dim results As Boolean
            Dim iDoc As XmlDocument = New XmlDocument()
            Dim mypath, errmsg, logging As String

            ' Database declarations
            Dim SqlS As String
            Dim returnv As Integer

            ' DMS Database declarations
            Dim d_con As SqlConnection
            Dim d_cmd As SqlCommand
            Dim d_dr As SqlDataReader
            Dim d_ConnS As String

            ' Logging declarations
            Dim ltemp As String
            Dim myeventlog As log4net.ILog
            Dim mydebuglog As log4net.ILog
            myeventlog = log4net.LogManager.GetLogger("EventLog")
            mydebuglog = log4net.LogManager.GetLogger("SDDKDebugLog")
            Dim logfile As String
            Dim LogStartTime As String = Now.ToString
            Dim VersionNum As String = "100"

            ' Web service declarations
            Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

            ' Local Cache declarations
            Dim DMSCache As New CachingWrapper.LocalCache

            ' DocKey declarations
            Dim KeyId, DocKeyId As String

            ' ============================================
            ' Variable setup
            mypath = HttpRuntime.AppDomainAppPath
            logging = "Y"
            errmsg = ""
            results = False
            SqlS = ""
            returnv = 0
            KeyId = ""
            DocKeyId = ""
            temp = ""

            ' ============================================
            ' Get parameters
            Debug = UCase(Left(Debug, 1))
            If Debug = "T" Then
                DocId = "1"
                DocKey = "Shared"
                KeyVal = ""
                PrFlag = "N"
            Else
                DocId = Trim(HttpUtility.UrlEncode(DocId))
                If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))

                DocKey = Trim(HttpUtility.UrlDecode(DocKey))
                If InStr(DocKey, "%") > 0 Then DocKey = Trim(HttpUtility.UrlDecode(DocKey))

                KeyVal = Trim(HttpUtility.UrlDecode(KeyVal))
                If InStr(KeyVal, "%") > 0 Then KeyVal = Trim(HttpUtility.UrlDecode(KeyVal))

                PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
                If InStr(PrFlag, "%") > 0 Then PrFlag = Trim(HttpUtility.UrlDecode(PrFlag))
            End If

            ' ============================================
            ' Get system defaults
            Try
                d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
                If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
                temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDocCat_debug")
                If temp = "Y" Then Debug = "Y"
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            ' ============================================
            ' Open log file if applicable
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                logfile = "C:\Logs\SaveDMSDocKey.log"
                Try
                    log4net.GlobalContext.Properties("SDDKLogFileName") = logfile
                    log4net.Config.XmlConfigurator.Configure()
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error Opening Log. "
                    results = "Failure"
                    GoTo CloseOut2
                End Try

                If Debug = "Y" Then
                    mydebuglog.Debug("----------------------------------")
                    mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                    mydebuglog.Debug("Parameters-")
                    mydebuglog.Debug("  DocId: " & DocId)
                    mydebuglog.Debug("  DocKey: " & DocKey)
                    mydebuglog.Debug("  KeyVal: " & KeyVal)
                    mydebuglog.Debug("  PrFlag: " & PrFlag)
                End If
            End If

            ' ============================================
            ' Validate Parameters
            If Trim(DocKey) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No keyword specified. "
                GoTo CloseOut2
            End If
            If Trim(DocId) = "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "No document specified. "
                GoTo CloseOut2
            End If
            If PrFlag = "" Then PrFlag = "N"

            ' ============================================
            ' Open SQL Server database connections
            errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
            If errmsg <> "" Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not connect to the database. "
                GoTo CloseOut
            End If

            ' ============================================
            ' Load Keywords from and/or into cache
            Dim dt As DataTable = New DataTable
            If Not DMSCache.GetCachedItem("Keywords") Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Keywords found in cache")
                Try
                    dt = DMSCache.GetCachedItem("Keywords")
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                    GoTo CloseOut
                End Try
            Else
                SqlS = "SELECT name, row_id FROM DMS.dbo.Keywords ORDER BY name"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Keywords into cache: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If d_dr.HasRows Then
                        dt.Load(d_dr)
                        DMSCache.AddToCache("Keywords", dt, CachingWrapper.CachePriority.NotRemovable)
                    End If
                    d_dr.Close()
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                    GoTo CloseOut
                End Try
            End If
            If dt Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not retrieve Keywords. "
                GoTo CloseOut
            End If

            ' Debug output
            If Debug = "Y" Then
                mydebuglog.Debug(" Keywords Columns found: " & dt.Columns.Count.ToString)
                mydebuglog.Debug(" Keywords Rows found: " & dt.Rows.Count.ToString)
            End If

            ' ============================================
            ' Locate DocKey Id in datatable
            Dim dr() As DataRow = dt.Select("name='" & DocKey & "'")
            If dr Is Nothing Or dr.Length = 0 Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Keyword not found")
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Could not find Keyword. "
                GoTo CloseOut
            End If
            KeyId = dr(0).Item("row_id").ToString
            If Debug = "Y" Then
                mydebuglog.Debug(" Keyword row_id: " & KeyId)
            End If

            ' ============================================
            ' Close data objects
            Try
                dr = Nothing
                dt = Nothing
                DMSCache = Nothing
            Catch ex As Exception
            End Try

            ' ============================================
            ' Write Document_Keywords record
            SqlS = "INSERT INTO DMS.dbo.Document_Keywords " & _
                    "(created_by, last_upd_by, doc_id, key_id, pr_flag, val) " & _
                    "VALUES (1, 1, " & DocId & ", " & KeyId & ", '" & PrFlag & "','" & KeyVal & "')"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Keywords record for document: " & vbCrLf & SqlS)
            If Debug <> "T" Then
                d_cmd.CommandText = SqlS
                Try
                    returnv = d_cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End If
            results = True

CloseOut:
            ' ============================================
            ' Close database connections and objects
            Try
                errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
            End Try

CloseOut2:
            ' ============================================
            ' Close the log file if any
            ltemp = results & " for Keyword " & DocKey & " with value '" & KeyVal & "' and document " & DocId
            If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDocKey :  Error: " & Trim(errmsg))
            myeventlog.Info("SaveDMSDocKey : Results: " & ltemp)
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                Try
                    If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                    mydebuglog.Debug("Results: " & ltemp)
                    If Debug = "Y" Then
                        mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                        mydebuglog.Debug("----------------------------------")
                    End If
                Catch ex As Exception
                End Try
            End If

            ' Log Performance Data
            If Debug <> "T" Then
                ' ============================================
                ' Send the web request
                Try
                    LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
                Catch ex As Exception
                End Try
            End If

            ' ============================================
            ' Return results
            Return results
        End Function

        Public Function SaveDMSDocUser(ByVal DocId As String, ByVal Domain As String, _
        ByVal DomainOwner As String, ByVal DomainRights As String, ByVal SubId As String, _
        ByVal SubOwner As String, ByVal SubRights As String, ByVal ConId As String, ByVal ConOwner As String, _
        ByVal ConRights As String, ByVal RegId As String, ByVal RegOwner As String, ByVal RegRights As String, _
        ByVal Debug As String) As Boolean

            ' This function creates Document_Users record(s) for the document and types of users specified

            ' The input parameters are as follows:
            '
            '   DocId		    - The "DMS.Documents.row_id" of the document (req.)
            '   DocId		    - A required document id (*DMS.Documents.row_id*)
            '   Domain		    - The Domain (hcidb1.CX_DOMAIN.DOMAIN) associated with a document - optional
            '   DomainOwner		- A flag to indicate the domain is the owner of the document - optional
            '   DomainRights	- The CRUD rights of the domain to the document - optional
            '   SubId		    - The Subscription Id (hcidb1.CX_SUBSCRIPTIONS.ROW_ID) associated with a document - optional
            '   SubOwner		- A flag to indicate the subscription is the owner of the document - optional
            '   SubRights		- The CRUD rights of the subscription to the document - optional
            '   ConId		    - The Contact Id (hcidb1.S_CONTACT.ROW_ID) associated with a document - optional
            '   ConOwner		- A flag to indicate the contact is the owner of the document - optional
            '   ConRights		- The CRUD rights of the contact to the document - optional
            '   RegId		    - The Contact registration (hcidb1.S_CONTACT.X_REGISTRATION_NUM) associated with a document - optional
            '   RegOwner		- A flag to indicate the registration is the owner of the document - optional
            '   RegRights		- The CRUD rights of the registration to the document - optional
            '   Debug		    - The debug mode flag: "Y", "N" or "T" 
            '
            ' The results are as follows:
            '
            '   Boolean    		- True/False to indicate success of the operation

            ' web.config Parameters used:
            '   dms			- connection string to DMS.dms database

            ' Variables
            Dim temp As String
            Dim results As Boolean
            Dim iDoc As XmlDocument = New XmlDocument()
            Dim mypath, errmsg, logging As String

            ' Database declarations
            Dim SqlS As String
            Dim returnv As Integer

            ' DMS Database declarations
            Dim d_con As SqlConnection
            Dim d_cmd As SqlCommand
            Dim d_dr As SqlDataReader
            Dim d_ConnS As String

            ' Logging declarations
            Dim ltemp As String
            Dim myeventlog As log4net.ILog
            Dim mydebuglog As log4net.ILog
            myeventlog = log4net.LogManager.GetLogger("EventLog")
            mydebuglog = log4net.LogManager.GetLogger("SDDUDebugLog")
            Dim logfile As String
            Dim LogStartTime As String = Now.ToString
            Dim VersionNum As String = "100"

            ' Web service declarations
            Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

            ' Local Cache declarations
            Dim DMSCache As New CachingWrapper.LocalCache

            ' User declarations
            Dim SubUGA, DomainUGA, ConUGA, RegUGA As String

            ' ============================================
            ' Variable setup
            mypath = HttpRuntime.AppDomainAppPath
            logging = "Y"
            errmsg = ""
            results = False
            SqlS = ""
            returnv = 0
            SubUGA = ""     ' Subscription User Group Access Id
            DomainUGA = ""  ' Domain User Group Access Id
            ConUGA = ""     ' Contact User Group Access Id
            RegUGA = ""     ' Web Registration User Group Access Id
            temp = ""

            ' ============================================
            ' Get parameters
            Debug = UCase(Left(Debug, 1))
            If Debug = "T" Then
                DocId = "1"
                Domain = "TIPS"
                DomainOwner = ""
                DomainRights = "R"
                SubId = ""
                SubOwner = ""
                SubRights = ""
                ConId = ""
                ConOwner = ""
                ConRights = ""
                RegId = ""
                RegOwner = ""
                RegRights = ""
            Else
                DocId = Trim(HttpUtility.UrlDecode(DocId.Trim))
                If InStr(DocId, "%") > 0 Then DocId = Trim(HttpUtility.UrlDecode(DocId))

                Domain = Trim(HttpUtility.UrlDecode(Domain.Trim)).ToUpper
                If InStr(Domain, "%") > 0 Then Domain = Trim(HttpUtility.UrlDecode(Domain))

                DomainOwner = Trim(HttpUtility.UrlDecode(DomainOwner.Trim)).ToUpper
                If InStr(DomainOwner, "%") > 0 Then DomainOwner = Trim(HttpUtility.UrlDecode(DomainOwner))
                If DomainOwner = "" Then DomainOwner = "N"

                DomainRights = Trim(HttpUtility.UrlDecode(DomainRights.Trim)).ToUpper
                If InStr(DomainRights, "%") > 0 Then DomainRights = Trim(HttpUtility.UrlDecode(DomainRights))

                SubId = Trim(HttpUtility.UrlDecode(SubId.Trim)).ToUpper
                If InStr(SubId, "%") > 0 Then SubId = Trim(HttpUtility.UrlDecode(SubId))

                SubOwner = Trim(HttpUtility.UrlDecode(SubOwner.Trim)).ToUpper
                If InStr(SubOwner, "%") > 0 Then SubOwner = Trim(HttpUtility.UrlDecode(SubOwner))
                If SubOwner = "" Then SubOwner = "N"

                SubRights = Trim(HttpUtility.UrlDecode(SubRights.Trim)).ToUpper
                If InStr(SubRights, "%") > 0 Then SubRights = Trim(HttpUtility.UrlDecode(SubRights))

                ConId = Trim(HttpUtility.UrlDecode(ConId.Trim)).ToUpper
                If InStr(ConId, "%") > 0 Then ConId = Trim(HttpUtility.UrlDecode(ConId))
                If InStr(ConId, " ") > 0 Then ConId = ConId.Replace(" ", "+")

                ConOwner = Trim(HttpUtility.UrlDecode(ConOwner.Trim)).ToUpper
                If InStr(ConOwner, "%") > 0 Then ConOwner = Trim(HttpUtility.UrlDecode(ConOwner))
                If ConOwner = "" Then ConOwner = "N"

                ConRights = Trim(HttpUtility.UrlDecode(ConRights.Trim)).ToUpper
                If InStr(ConRights, "%") > 0 Then ConRights = Trim(HttpUtility.UrlDecode(ConRights))

                RegId = Trim(HttpUtility.UrlDecode(RegId.Trim)).ToUpper
                If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))

                RegOwner = Trim(HttpUtility.UrlDecode(RegOwner.Trim)).ToUpper
                If InStr(RegOwner, "%") > 0 Then RegOwner = Trim(HttpUtility.UrlDecode(RegOwner))
                If RegOwner = "" Then RegOwner = "N"

                RegRights = Trim(HttpUtility.UrlDecode(RegRights.Trim)).ToUpper
                If InStr(RegRights, "%") > 0 Then RegRights = Trim(HttpUtility.UrlDecode(RegRights))
            End If

            ' ============================================
            ' Get system defaults
            Try
                d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
                If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
                temp = System.Configuration.ConfigurationManager.AppSettings.Get("SaveDMSDocUser_debug")
                If temp = "Y" And Debug <> "T" Then Debug = "Y"
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
                results = False
                GoTo CloseOut2
            End Try

            ' ============================================
            ' Open log file if applicable
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                logfile = "C:\Logs\SaveDMSDocUser.log"
                Try
                    log4net.GlobalContext.Properties("SDDULogFileName") = logfile
                    log4net.Config.XmlConfigurator.Configure()
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error Opening Log. "
                    results = False
                    GoTo CloseOut2
                End Try

                If Debug = "Y" Then
                    mydebuglog.Debug("----------------------------------")
                    mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                    mydebuglog.Debug("Parameters-")
                    mydebuglog.Debug("  DocId: " & DocId)
                    mydebuglog.Debug("  Domain:" & Domain)
                    mydebuglog.Debug("  DomainOwner:" & DomainOwner)
                    mydebuglog.Debug("  DomainRights:" & DomainRights)
                    mydebuglog.Debug("  SubId:" & SubId)
                    mydebuglog.Debug("  SubOwner:" & SubOwner)
                    mydebuglog.Debug("  SubRights:" & SubRights)
                    mydebuglog.Debug("  ConId:" & ConId)
                    mydebuglog.Debug("  ConOwner:" & ConOwner)
                    mydebuglog.Debug("  ConRights:" & ConRights)
                    mydebuglog.Debug("  RegId:" & RegId)
                    mydebuglog.Debug("  RegOwner:" & RegOwner)
                    mydebuglog.Debug("  RegRights:" & RegRights)
                End If
            End If

            ' ============================================
            ' Validate Parameters
            If Trim(DocId) = "" Then
                results = False
                errmsg = errmsg & vbCrLf & "No document specified. "
                GoTo CloseOut2
            End If

            ' ============================================
            ' Open SQL Server database connections
            errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
            If errmsg <> "" Then
                results = False
                errmsg = errmsg & vbCrLf & "Could not connect to the database. "
                GoTo CloseOut
            End If

            ' ============================================
            ' Load Subscription Users from and/or into cache
            Dim dt1 As DataTable = New DataTable
            Dim dr1() As DataRow
            If SubId <> "" Then
                If Not DMSCache.GetCachedItem("Subscriptions") Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Subscriptions found in cache")
                    Try
                        dt1 = DMSCache.GetCachedItem("Subscriptions")
                    Catch ex As Exception
                        results = False
                        errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                        GoTo CloseOut
                    End Try
                Else
                    SqlS = "SELECT DISTINCT G.name, A.row_id " & _
                    "FROM DMS.dbo.Groups G " & _
                    "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=G.row_id " & _
                    "WHERE G.type_cd='Subscription' AND A.type_id='G'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Subscriptions into cache: " & SqlS)
                    Try
                        d_cmd.CommandText = SqlS
                        d_dr = d_cmd.ExecuteReader()
                        If d_dr.HasRows Then
                            dt1.Load(d_dr)
                            DMSCache.AddToCache("Subscriptions", dt1, CachingWrapper.CachePriority.Default)
                        End If
                        d_dr.Close()
                    Catch ex As Exception
                        results = False
                        errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                        GoTo CloseOut
                    End Try
                End If
                If dt1 Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not retrieve Subscriptions. "
                    GoTo CloseOut
                End If

                ' Debug output
                If Debug = "Y" Then
                    mydebuglog.Debug(" Subscriptions Columns found: " & dt1.Columns.Count.ToString)
                    mydebuglog.Debug(" Subscriptions Rows found: " & dt1.Rows.Count.ToString)
                End If

                ' Locate Subscription UGA in datatable
                Try
                    dr1 = dt1.Select("name='" & SubId & "'")
                    If dr1 Is Nothing Or dr1.Length = 0 Then
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Subscription not found")
                        errmsg = errmsg & vbCrLf & "Could not find Subscription. "
                    Else
                        SubUGA = dr1(0).Item("row_id").ToString
                    End If
                Catch ex As Exception
                End Try
                If Debug = "Y" Then
                    mydebuglog.Debug(" Subscription UGA row_id: " & SubUGA)
                End If
            End If

            ' ============================================
            ' Load Domain Users from and/or into cache
            Dim dt2 As DataTable = New DataTable
            Dim dr2() As DataRow
            If Domain <> "" Then
                If Not DMSCache.GetCachedItem("Domains") Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Domains found in cache")
                    Try
                        dt2 = DMSCache.GetCachedItem("Domains")
                    Catch ex As Exception
                        results = False
                        errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                        GoTo CloseOut
                    End Try
                Else
                    SqlS = "SELECT DISTINCT G.name, A.row_id " & _
                    "FROM DMS.dbo.Groups G " & _
                    "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=G.row_id " & _
                    "WHERE G.type_cd='Domain' AND A.type_id='G'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Domains into cache: " & SqlS)
                    Try
                        d_cmd.CommandText = SqlS
                        d_dr = d_cmd.ExecuteReader()
                        If d_dr.HasRows Then
                            dt2.Load(d_dr)
                            DMSCache.AddToCache("Domains", dt2, CachingWrapper.CachePriority.Default)
                        End If
                        d_dr.Close()
                    Catch ex As Exception
                        results = False
                        errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                        GoTo CloseOut
                    End Try
                End If
                If dt2 Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not retrieve Domains. "
                    GoTo CloseOut
                End If

                ' Debug output
                If Debug = "Y" Then
                    mydebuglog.Debug(" Domains Columns found: " & dt2.Columns.Count.ToString)
                    mydebuglog.Debug(" Domains Rows found: " & dt2.Rows.Count.ToString)
                End If

                ' Locate Domain UGA in datatable
                Try
                    dr2 = dt2.Select("name='" & Domain & "'")
                    If dr2 Is Nothing Or dr2.Length = 0 Then
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Domain not found")
                        errmsg = errmsg & vbCrLf & "Could not find Domain. "
                    Else
                        DomainUGA = dr2(0).Item("row_id").ToString
                    End If
                Catch ex As Exception
                End Try
                If Debug = "Y" Then
                    mydebuglog.Debug(" Domain UGA row_id: " & DomainUGA)
                End If
            End If

            ' ============================================
            ' Load Contact Users from and/or into cache
            Dim dt3 As DataTable = New DataTable
            Dim dr3() As DataRow
            If ConId <> "" Then
                If Not DMSCache.GetCachedItem("Contacts") Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Contacts found in cache")
                    Try
                        dt3 = DMSCache.GetCachedItem("Contacts")
                    Catch ex As Exception
                        results = False
                        errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                        GoTo CloseOut
                    End Try
                Else
                    SqlS = "SELECT DISTINCT U.ext_user_id as name, A.row_id " & _
                    "FROM DMS.dbo.Users U " & _
                    "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=U.row_id " & _
                    "WHERE A.type_id='U' AND U.ext_user_id IS NOT NULL"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Contacts into cache: " & SqlS)
                    Try
                        d_cmd.CommandText = SqlS
                        d_dr = d_cmd.ExecuteReader()
                        If d_dr.HasRows Then
                            dt3.Load(d_dr)
                            DMSCache.AddToCache("Contacts", dt3, CachingWrapper.CachePriority.Default)
                        End If
                        d_dr.Close()
                    Catch ex As Exception
                        results = False
                        errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                        GoTo CloseOut
                    End Try
                End If
                If dt3 Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not retrieve Contacts. "
                    GoTo CloseOut
                End If

                ' Debug output
                If Debug = "Y" Then
                    mydebuglog.Debug(" Contacts Columns found: " & dt3.Columns.Count.ToString)
                    mydebuglog.Debug(" Contacts Rows found: " & dt3.Rows.Count.ToString)
                End If

                ' Locate Domain UGA in datatable
                Try
                    dr3 = dt3.Select("name='" & ConId & "'")
                    If dr3 Is Nothing Or dr3.Length = 0 Then
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Contact not found")
                        'errmsg = errmsg & vbCrLf & "Could not find Contact. "
                        ConUGA = ""
                    Else
                        ConUGA = dr3(0).Item("row_id").ToString
                    End If
                Catch ex As Exception
                End Try
                If Debug = "Y" Then
                    mydebuglog.Debug(" Contact UGA row_id: " & ConUGA)
                End If
            End If

            ' ============================================
            ' Load Registration Contact Users from and/or into cache
            Dim dt4 As DataTable = New DataTable
            Dim dr4() As DataRow
            If RegId <> "" Then
                If Not DMSCache.GetCachedItem("Registrations") Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Registrations found in cache")
                    Try
                        dt4 = DMSCache.GetCachedItem("Registrations")
                    Catch ex As Exception
                        results = False
                        errmsg = errmsg & vbCrLf & "Could not get data from cache: " & ex.Message
                        GoTo CloseOut
                    End Try
                Else
                    SqlS = "SELECT DISTINCT U.ext_id as name, A.row_id " & _
                    "FROM DMS.dbo.Users U " & _
                    "LEFT OUTER JOIN DMS.dbo.User_Group_Access A ON A.access_id=U.row_id " & _
                    "WHERE A.type_id='U' AND U.ext_id IS NOT NULL"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Loading Registrations into cache: " & SqlS)
                    Try
                        d_cmd.CommandText = SqlS
                        d_dr = d_cmd.ExecuteReader()
                        If d_dr.HasRows Then
                            dt4.Load(d_dr)
                            DMSCache.AddToCache("Registrations", dt4, CachingWrapper.CachePriority.Default)
                        End If
                        d_dr.Close()
                    Catch ex As Exception
                        results = False
                        errmsg = errmsg & vbCrLf & "Could not retrieve data from SQL or load to datatable: " & ex.Message
                        GoTo CloseOut
                    End Try
                End If
                If dt4 Is Nothing Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Datatable not found")
                    results = False
                    errmsg = errmsg & vbCrLf & "Could not retrieve Registrations. "
                    GoTo CloseOut
                End If

                ' Debug output
                If Debug = "Y" Then
                    mydebuglog.Debug(" Registrations Columns found: " & dt4.Columns.Count.ToString)
                    mydebuglog.Debug(" Registrations Rows found: " & dt4.Rows.Count.ToString)
                End If

                ' Locate Domain UGA in datatable
                Try
                    dr4 = dt4.Select("name='" & RegId & "'")
                    If dr4 Is Nothing Or dr4.Length = 0 Then
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Registration not found")
                        errmsg = errmsg & vbCrLf & "Could not find Registration. "
                    Else
                        RegUGA = dr4(0).Item("row_id").ToString
                    End If
                Catch ex As Exception
                End Try
                If Debug = "Y" Then
                    mydebuglog.Debug(" Registration UGA row_id: " & RegUGA)
                End If
            End If

            ' ============================================
            ' Close data objects
            Try
                dr1 = Nothing
                dt1 = Nothing
            Catch ex As Exception
            End Try
            Try
                dr2 = Nothing
                dt2 = Nothing
            Catch ex As Exception
            End Try
            Try
                dr3 = Nothing
                dt3 = Nothing
            Catch ex As Exception
            End Try
            Try
                dr4 = Nothing
                dt4 = Nothing
            Catch ex As Exception
            End Try
            Try
                DMSCache = Nothing
            Catch ex As Exception
            End Try

            ' ============================================
            ' Write Document User records
            '
            ' Subscription User Group Access Id
            If SubUGA <> "" Then
                SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                "VALUES (1, 1, " & DocId & ", " & SubUGA & ", '" & SubOwner & "', '" & SubRights & "')"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Subscription: " & vbCrLf & SqlS)
                If Debug <> "T" Then
                    d_cmd.CommandText = SqlS
                    Try
                        returnv = d_cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End If
            End If

            ' Domain User Group Access Id
            If DomainUGA <> "" Then
                SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                "VALUES (1, 1, " & DocId & ", " & DomainUGA & ", '" & DomainOwner & "', '" & DomainRights & "')"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Domain: " & vbCrLf & SqlS)
                If Debug <> "T" Then
                    d_cmd.CommandText = SqlS
                    Try
                        returnv = d_cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End If
            End If

            ' Contact User Group Access Id
            If ConUGA <> "" Then
                SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                "VALUES (1, 1, " & DocId & ", " & ConUGA & ", '" & ConOwner & "', '" & ConRights & "')"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Contact: " & vbCrLf & SqlS)
                If Debug <> "T" Then
                    d_cmd.CommandText = SqlS
                    Try
                        returnv = d_cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End If
            End If

            ' Web Registration User Group Access Id
            If RegUGA <> "" Then
                SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                "VALUES (1, 1, " & DocId & ", " & RegUGA & ", '" & RegOwner & "', '" & RegRights & "')"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Registration: " & vbCrLf & SqlS)
                If Debug <> "T" Then
                    d_cmd.CommandText = SqlS
                    Try
                        returnv = d_cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End If
            End If
            results = True

            ' Create access for supervisor if no contact was specified just to make sure someone "owns" the document
            If ConUGA = "" And RegUGA = "" Then
                SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                       "VALUES (1, 1, " & DocId & ", 1, 'Y', 'REDO')"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Users record for Supervisor: " & vbCrLf & SqlS)
                If Debug <> "T" Then
                    d_cmd.CommandText = SqlS
                    Try
                        returnv = d_cmd.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End If
            End If

CloseOut:
            ' ============================================
            ' Close database connections and objects
            Try
                errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
            End Try

CloseOut2:
            ' ============================================
            ' Close the log file if any
            ltemp = results & " for document " & DocId
            If Trim(errmsg) <> "" Then myeventlog.Error("SaveDMSDocUser :  Error: " & Trim(errmsg))
            myeventlog.Info("SaveDMSDocUser : Results: " & ltemp)
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                Try
                    If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                    mydebuglog.Debug("Results: " & ltemp)
                    If Debug = "Y" Then
                        mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                        mydebuglog.Debug("----------------------------------")
                    End If
                Catch ex As Exception
                End Try
            End If

            ' Log Performance Data
            If Debug <> "T" Then
                ' ============================================
                ' Send the web request
                Try
                    LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
                Catch ex As Exception
                End Try
            End If

            ' ============================================
            ' Return results
            Return results
        End Function

        Public Function UpdDMSDocCount(ByVal CONTACT_ID As String, ByVal TRAINER_NUM As String, ByVal PART_ID As String, _
               ByVal MT_ID As String, ByVal Debug As String) As String

            ' This function creates an association for the individual specified, and optionally sends them an
            ' email notice with portal access instructions

            '   CONTACT_ID	- The "siebeldb.S_CONTACT.ROW_ID" of the individual (req)
            '   TRAINER_NUM	- The "siebeldb.S_CONTACT.X_TRAINER_NUM" of the individual (opt)
            '   PART_ID	- The "siebeldb.S_CONTACT.X_PART_ID" of the individual (opt)
            '   MT_ID	- The "siebeldb.S_CONTACT.REG_AS_EMP_ID" of the individual (opt)

            ' web.config Parameters used:
            '   dms        	    - connection string to DMS.dms database

            ' Variables
            Dim results, temp As String
            Dim iDoc As XmlDocument = New XmlDocument()
            Dim mypath, errmsg, logging As String
            Dim bResponse As Boolean
            Dim doc_count As String
            Dim SUB_ID, DOMAIN, USER_AID, SUB_AID, DOMAIN_AID, UID As String
            Dim Category_Constraint, TRAINER_FLG, MT_FLG, PART_FLG, TRAINING_FLG, TRAINER_ACC_FLG, SITE_ONLY, SYSADMIN_FLG, EMP_ID As String

            ' Database declarations
            Dim SqlS As String
            Dim returnv As Integer

            ' HCIDB Database declarations
            Dim con As SqlConnection
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            Dim ConnS As String

            ' DMS Database declarations
            Dim d_con As SqlConnection
            Dim d_cmd As SqlCommand
            Dim d_dr As SqlDataReader
            Dim d_ConnS As String

            ' Logging declarations
            Dim ltemp As String
            Dim myeventlog As log4net.ILog
            Dim mydebuglog As log4net.ILog
            myeventlog = log4net.LogManager.GetLogger("EventLog")
            mydebuglog = log4net.LogManager.GetLogger("UDDCDebugLog")
            Dim logfile As String
            Dim LogStartTime As String = Now.ToString
            Dim VersionNum As String = "100"

            ' Web service declarations
            Dim LoggingService As New basic.com.certegrity.cloudsvc.Service

            ' ============================================
            ' Variable setup
            mypath = HttpRuntime.AppDomainAppPath
            logging = "Y"
            temp = ""
            errmsg = ""
            bResponse = False
            doc_count = "0"
            results = "Success"
            SUB_ID = ""
            DOMAIN = ""
            USER_AID = ""
            SUB_AID = ""
            DOMAIN_AID = ""
            UID = ""
            TRAINER_FLG = ""
            MT_FLG = ""
            PART_FLG = ""
            TRAINING_FLG = ""
            TRAINER_ACC_FLG = ""
            SITE_ONLY = ""
            SYSADMIN_FLG = ""
            EMP_ID = ""

            ' ============================================
            ' Fix parameters
            Debug = UCase(Left(Debug, 1))
            If Debug = "" Then Debug = "N"
            If Debug = "T" Then
                CONTACT_ID = "21120611WE0"
                PART_ID = "732632"
                MT_ID = ""
                TRAINER_NUM = "22"
            Else
                CONTACT_ID = Trim(HttpUtility.UrlEncode(CONTACT_ID))
                If InStr(CONTACT_ID, "%") > 0 Then CONTACT_ID = Trim(HttpUtility.UrlDecode(CONTACT_ID))
                CONTACT_ID = Trim(EncodeParamSpaces(CONTACT_ID))
                PART_ID = Trim(HttpUtility.UrlEncode(PART_ID))
                If InStr(PART_ID, "%") > 0 Then PART_ID = Trim(HttpUtility.UrlDecode(PART_ID))
                PART_ID = Trim(EncodeParamSpaces(PART_ID))
                MT_ID = Trim(HttpUtility.UrlEncode(MT_ID))
                If InStr(MT_ID, "%") > 0 Then MT_ID = Trim(HttpUtility.UrlDecode(MT_ID))
                MT_ID = Trim(EncodeParamSpaces(MT_ID))
            End If

            ' ============================================
            ' Get system defaults
            ' hcidb1
            Try
                ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("siebeldb").ConnectionString
                If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
                temp = System.Configuration.ConfigurationManager.AppSettings.Get("UpdDMSDocCount_debug")
                If temp = "Y" And Debug <> "T" Then Debug = "Y"
            Catch ex As Exception
                errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
                results = "Failure"
                GoTo CloseOut2
            End Try
            ' dms
            Try
                d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
                If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;database=DMS"
            Catch ex As Exception
                errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
                results = "Failure"
                GoTo CloseOut2
            End Try

            ' ============================================
            ' Open log file if applicable
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                logfile = "C:\Logs\UpdDMSDocCount.log"
                Try
                    log4net.GlobalContext.Properties("UDDCLogFileName") = logfile
                    log4net.Config.XmlConfigurator.Configure()
                Catch ex As Exception
                    errmsg = errmsg & "Error Opening Log. " & vbCrLf
                    results = "Failure"
                    GoTo CloseOut2
                End Try

                If Debug = "Y" Then
                    mydebuglog.Debug("----------------------------------")
                    mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                    mydebuglog.Debug("Parameters-")
                    mydebuglog.Debug("  CONTACT_ID: " & CONTACT_ID)
                    mydebuglog.Debug("  PART_ID: " & PART_ID)
                    mydebuglog.Debug("  MT_ID: " & MT_ID)
                    mydebuglog.Debug("  TRAINER_NUM: " & TRAINER_NUM)
                End If
            End If

            ' ============================================
            ' Check required parameters
            If (CONTACT_ID = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "Invalid parameter(s) " & vbCrLf
                GoTo CloseOut2
            End If

            ' ============================================
            ' Open database connections
            errmsg = OpenDBConnection(ConnS, con, cmd)
            If errmsg <> "" Then
                results = "Failure"
                GoTo CloseOut
            End If
            errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)
            If errmsg <> "" Then
                results = "Failure"
                GoTo CloseOut
            End If

            ' ============================================
            ' Get Subscription and Domain Info
            'SqlS = "SELECT S.ROW_ID, S.DOMAIN, C.X_REGISTRATION_NUM " & _
            '"FROM siebeldb.dbo.CX_SUB_CON SC  " & _
            '"INNER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID " & _
            '"INNER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=SC.CON_ID " & _
            '"WHERE SC.CON_ID='" & CONTACT_ID & "'"
            SqlS = "SELECT S.ROW_ID, S.DOMAIN, C.X_REGISTRATION_NUM, C.X_TRAINER_FLG, C.X_MAST_TRNR_FLG, " & _
            "(SELECT CASE WHEN C.X_PART_ID IS NOT NULL AND C.X_PART_ID<>'' THEN 'Y' ELSE 'N' END) AS PART_FLG, S.SVC_TYPE, " & _
            "SC.TRAINER_ACC_FLG, SC.SITE_ONLY_FLG, SC.SYSADMIN_FLG, E.ROW_ID " & _
            "FROM siebeldb.dbo.CX_SUB_CON SC  " & _
            "INNER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID  " & _
            "INNER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=SC.CON_ID  " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_EMPLOYEE E ON E.X_CON_ID=C.ROW_ID AND E.CNTRCTR_EMPLR_ID IS NULL  " & _
            "WHERE SC.CON_ID='" & CONTACT_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  Get subscription info: " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            SUB_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                            DOMAIN = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                            UID = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                            TRAINER_FLG = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                            MT_FLG = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                            PART_FLG = Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString
                            TRAINING_FLG = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                            Select Case TRAINING_FLG.ToUpper()
                                Case "CERTIFICATION MANAGER REG DB"
                                    TRAINING_FLG = "N"
                                Case "CERTIFICATION MANAGER REPORTS"
                                    TRAINING_FLG = "N"
                                Case Else
                                    TRAINING_FLG = "Y"
                            End Select
                            TRAINER_ACC_FLG = Trim(CheckDBNull(dr(7), enumObjectType.StrType)).ToString
                            SITE_ONLY = Trim(CheckDBNull(dr(8), enumObjectType.StrType)).ToString
                            SYSADMIN_FLG = Trim(CheckDBNull(dr(9), enumObjectType.StrType)).ToString
                            EMP_ID = Trim(CheckDBNull(dr(10), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                            'results = "Failure"
                            'errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Error getting document count. " & vbCrLf
                    results = "Failure"
                End If
                Try
                    dr.Close()
                    dr = Nothing
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug("      > Sub_Id/Domain/UID: " & SUB_ID & "/" & DOMAIN & "/" & UID)
                mydebuglog.Debug("      > TRAINER_FLG:" & TRAINER_FLG)
                mydebuglog.Debug("      > MT_FLG: " & MT_FLG)
                mydebuglog.Debug("      > PART_FLG: " & PART_FLG)
                mydebuglog.Debug("      > TRAINING_FLG: " & TRAINING_FLG)
                mydebuglog.Debug("      > TRAINER_ACC_FLG: " & TRAINER_ACC_FLG)
                mydebuglog.Debug("      > SITE_ONLY: " & SITE_ONLY)
                mydebuglog.Debug("      > SYSADMIN_FLG: " & SYSADMIN_FLG)
                mydebuglog.Debug("      > EMP_ID: " & EMP_ID)
            End If

            ' If no subscription, no point
            If SUB_ID = "" Or UID = "" Then
                errmsg = errmsg & "No subscription to update for UID " & UID
                results = "Failure"
                GoTo CloseOut
            End If

            ' ============================================
            ' Get DMS security information
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get DMS security information")

            ' -----
            ' User AID
            SqlS = "SELECT UA.row_id " & _
            "FROM DMS.dbo.User_Group_Access UA " & _
            "INNER JOIN DMS.dbo.Users U ON U.row_id=UA.access_id " & _
            "WHERE UA.type_id='U' AND U.ext_user_id='" & CONTACT_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  .. Get user security: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            USER_AID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                        End Try
                    End While
                End If
                Try
                    d_dr.Close()
                    d_dr = Nothing
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug("      > USER_AID: " & USER_AID)

            ' -----
            ' Subscription AID
            SqlS = "SELECT UA.row_id " & _
            "FROM DMS.dbo.User_Group_Access UA " & _
            "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " & _
            "WHERE UA.type_id='G' AND G.name='" & SUB_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  .. Get subscription security: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            SUB_AID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                        End Try
                    End While
                End If
                Try
                    d_dr.Close()
                    d_dr = Nothing
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug("      > SUB_AID: " & SUB_AID)

            ' -----
            ' Domain AID
            SqlS = "SELECT UA.row_id " & _
            "FROM DMS.dbo.User_Group_Access UA " & _
            "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " & _
            "WHERE UA.type_id='G' AND G.name='" & DOMAIN & "'"
            If Debug = "Y" Then mydebuglog.Debug("  .. Get domain security: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            DOMAIN_AID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                        End Try
                    End While
                End If
                Try
                    d_dr.Close()
                    d_dr = Nothing
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug("      > DOMAIN_AID: " & DOMAIN_AID)

            ' ============================================
            ' Generate Category Constraint
            Category_Constraint = "CK.key_id IN ("
            If TRAINER_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "3,"
            End If
            If MT_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "5,"
            End If
            If PART_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "7,"
            End If
            If TRAINING_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "8,"
            End If
            If TRAINER_ACC_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "10,"
            End If
            If SITE_ONLY = "Y" Then
                Category_Constraint = Category_Constraint & "12,"
            End If
            Category_Constraint = Category_Constraint & "13,"
            If SYSADMIN_FLG = "Y" Then
                Category_Constraint = Category_Constraint & "15,"
            End If
            If EMP_ID <> "" Then
                Category_Constraint = Category_Constraint & "16,"
            End If
            Category_Constraint = Category_Constraint & "14) "
            If Debug = "Y" Then mydebuglog.Debug("  Category_Constraint: " & Category_Constraint)

            ' ============================================
            ' Get current document count if the user has a subscription
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Generate document count")

            SqlS = "SELECT count(1) AS NUM_DOC " & _
            "FROM (" & _
            "SELECT D.row_id " & _
            "FROM DMS.dbo.Documents D " & _
            "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " & _
            "LEFT OUTER JOIN DMS.dbo.Category_Keywords CK ON CK.cat_id=DC.cat_id " & _
            "WHERE DC.pr_flag='Y' AND (CK.key_id IN (3,5,7,8,13,15,16,14)) " & _
            "GROUP BY D.row_id " & _
            "INTERSECT " & _
            "SELECT DISTINCT DA.doc_id " & _
            "FROM DMS.dbo.Document_Associations DA " & _
            "INNER JOIN DMS.dbo.Documents D on D.row_id=DA.doc_id " & _
            "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id "
            SqlS = SqlS & "WHERE ((DA.association_id='3' AND DA.fkey='" & CONTACT_ID & "' AND DA.pr_flag='Y') or "
            If TRAINER_NUM <> "" Then SqlS = SqlS & "(DA.association_id='5' AND DA.fkey='" & TRAINER_NUM & "' AND DA.pr_flag='Y') or "
            If PART_ID <> "" Then SqlS = SqlS & "(DA.association_id='4' AND DA.fkey='" & PART_ID & "' AND DA.pr_flag='Y') or "
            If MT_ID <> "" Then SqlS = SqlS & "(DA.association_id='37' AND DA.fkey='" & MT_ID & "' AND DA.pr_flag='Y') or "
            SqlS = Left(SqlS, Len(SqlS) - 4) & ") AND D.deleted IS NULL AND ("
            If USER_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & USER_AID & " OR "
            If SUB_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & SUB_AID & " OR "
            If DOMAIN_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & DOMAIN_AID & " OR "
            SqlS = Left(SqlS, Len(SqlS) - 4) & ") GROUP BY DA.doc_id ) d "
            If Debug = "Y" Then mydebuglog.Debug("  .. Get document count: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            doc_count = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Error getting document count. " & vbCrLf
                    results = "Failure"
                End If
                d_dr.Close()
                d_dr = Nothing
            Catch ex As Exception
            End Try
            If Debug = "Y" Then mydebuglog.Debug("      > doc_count: " & doc_count)

            ' -----
            ' Drop temp table
            'SqlS = "DROP TABLE DMS.dbo.[" & UID & "]"
            'If Debug = "Y" Then mydebuglog.Debug("  .. Drop temp table for count: " & SqlS)
            'Try
            'd_cmd.CommandText = SqlS
            'returnv = d_cmd.ExecuteNonQuery()
            'Catch ex As Exception
            'End Try

            ' -----
            ' Get list of visible documents and store in temp table
            'SqlS = "SELECT DISTINCT DA.doc_id " & _
            '"INTO DMS.dbo.[" & UID & "] " & _
            '"FROM DMS.dbo.Document_Associations DA " & _
            '"INNER JOIN DMS.dbo.Documents D on D.row_id=DA.doc_id " & _
            '"INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id " & _
            '"WHERE ((DA.association_id='3' AND DA.fkey='" & CONTACT_ID & "' AND DA.pr_flag='Y') or "
            'If TRAINER_NUM <> "" Then SqlS = SqlS & "(DA.association_id='5' AND DA.fkey='" & TRAINER_NUM & "' AND DA.pr_flag='Y') or "
            'If PART_ID <> "" Then SqlS = SqlS & "(DA.association_id='4' AND DA.fkey='" & PART_ID & "' AND DA.pr_flag='Y') or "
            'If MT_ID <> "" Then SqlS = SqlS & "(DA.association_id='37' AND DA.fkey='" & MT_ID & "' AND DA.pr_flag='Y') or "
            'SqlS = Left(SqlS, Len(SqlS) - 4) & ") AND D.deleted IS NULL AND ("
            'If USER_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & USER_AID & " OR "
            'If SUB_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & SUB_AID & " OR "
            'If DOMAIN_AID <> "" Then SqlS = SqlS & "DU.user_access_id=" & DOMAIN_AID & " OR "
            'SqlS = Left(SqlS, Len(SqlS) - 4) & ")"
            'SqlS = SqlS & " GROUP BY DA.doc_id"
            'If Debug = "Y" Then mydebuglog.Debug("  .. Create temp table for count: " & SqlS)
            'Try
            'd_cmd.CommandText = SqlS
            'returnv = d_cmd.ExecuteNonQuery()
            'Catch ex As Exception
            'End Try

            ' -----
            ' Get count of documents in temp table
            'SqlS = "SELECT COUNT(*) AS NUM_DOC " & _
            '"FROM DMS.dbo.[" & UID & "]"
            'If Debug = "Y" Then mydebuglog.Debug("  .. Get document count from temp table: " & SqlS)
            'Try
            'd_cmd.CommandText = SqlS
            'd_dr = d_cmd.ExecuteReader()
            'If Not d_dr Is Nothing Then
            'While d_dr.Read()
            'Try
            'doc_count = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType)).ToString
            'Catch ex As Exception
            'results = "Failure"
            'errmsg = errmsg & "Error getting document count. " & ex.ToString & vbCrLf
            'GoTo CloseOut
            'End Try
            'End While
            'Else
            'errmsg = errmsg & "Error getting document count. " & vbCrLf
            'results = "Failure"
            'End If
            'd_dr.Close()
            'd_dr = Nothing
            'Catch ex As Exception
            'End Try
            'If Debug = "Y" Then mydebuglog.Debug("      > doc_count: " & doc_count)

            ' -----
            ' Drop temp table
            'SqlS = "DROP TABLE DMS.dbo.[" & UID & "]"
            'If Debug = "Y" Then mydebuglog.Debug("  .. Drop temp table for count: " & SqlS)
            'Try
            'd_cmd.CommandText = SqlS
            'returnv = d_cmd.ExecuteNonQuery()
            'Catch ex As Exception
            'End Try

            ' -----
            ' Update CX_SUB_CON.NEW_DOC with document count if applicable
            If doc_count <> "" Then
                SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
                "SET NEW_DOC=" & doc_count & _
                " WHERE CON_ID='" & CONTACT_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug("  .. Update contact document count in CX_SUB_CON: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    returnv = cmd.ExecuteNonQuery()
                    If returnv = 0 Then results = "Failure"
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error setting the document count. " & ex.ToString & vbCrLf
                End Try

                SqlS = "UPDATE siebeldb.dbo.S_CONTACT " & _
                    "SET DCKING_NUM=" & doc_count & " " & _
                    "WHERE ROW_ID='" & CONTACT_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug("  .. Update contact document count in S_CONTACT: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    returnv = cmd.ExecuteNonQuery()
                    If returnv = 0 Then results = "Failure"
                Catch ex As Exception
                    results = "Failure"
                    errmsg = errmsg & "Error setting the document count. " & ex.ToString & vbCrLf
                End Try
            End If

CloseOut:
            ' ============================================
            ' Close database connections and objects
            Try
                errmsg = errmsg & CloseDBConnection(con, cmd, dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the hcidb database connection. " & vbCrLf
            End Try
            Try
                errmsg = errmsg & CloseDBConnection(d_con, d_cmd, d_dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the dms database connection. " & vbCrLf
            End Try

CloseOut2:
            ' ============================================
            ' Close the log file if any
            ltemp = results & " : Contact id " & CONTACT_ID & " has " & doc_count & " documents"
            If Trim(errmsg) <> "" Then myeventlog.Error("UpdDMSDocCount :  Error: " & Trim(errmsg))
            myeventlog.Info("UpdDMSDocCount : Results: " & ltemp)
            If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
                Try
                    If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                    If Debug = "Y" Then
                        mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                        mydebuglog.Debug("----------------------------------")
                    Else
                        mydebuglog.Debug("  Results: " & ltemp)
                    End If
                Catch ex As Exception
                End Try
            End If

            ' ============================================
            ' Log Performance Data
            If Debug <> "T" Then
                ' Send the web request
                Try
                    LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
                Catch ex As Exception
                End Try
                If results = "Success" Then results = doc_count
            End If

            ' ============================================
            ' Return results        
            Return results

        End Function

        ' ====================================================
        ' Minio functions

        ' ====================================================
        ' Database support functions
        Public Function OpenDBConnection(ByVal ConnS As String, ByRef con As SqlConnection, ByRef cmd As SqlCommand) As String
            ' Function to open a database connection with extreme error-handling
            ' Returns an error message if unable to open the connection
            Dim SqlS As String
            SqlS = ""
            OpenDBConnection = ""

            Try
                con = New SqlConnection(ConnS)
                con.Open()
                If Not con Is Nothing Then
                    Try
                        cmd = New SqlCommand(SqlS, con)
                        cmd.CommandTimeout = 300
                    Catch ex2 As Exception
                        OpenDBConnection = "Error opening the command string: " & ex2.ToString
                    End Try
                End If
            Catch ex As Exception
                If con.State <> Data.ConnectionState.Closed Then con.Dispose()
                ConnS = ConnS & ";Pooling=false"
                Try
                    con = New SqlConnection(ConnS)
                    con.Open()
                    If Not con Is Nothing Then
                        Try
                            cmd = New SqlCommand(SqlS, con)
                            cmd.CommandTimeout = 300
                        Catch ex2 As Exception
                            OpenDBConnection = "Error opening the command string: " & ex2.ToString
                        End Try
                    End If
                Catch ex2 As Exception
                    OpenDBConnection = "Unable to open database connection for connection string: " & ConnS & vbCrLf & "Windows error: " & vbCrLf & ex2.ToString & vbCrLf
                End Try
            End Try

        End Function

        Public Function CloseDBConnection(ByRef con As SqlConnection, ByRef cmd As SqlCommand, ByRef dr As SqlDataReader) As String
            ' This function closes a database connection safely
            Dim ErrMsg As String
            ErrMsg = ""

            ' Handle datareader
            Try
                dr.Close()
            Catch ex As Exception
            End Try
            Try
                dr = Nothing
            Catch ex As Exception
            End Try

            ' Handle command
            Try
                cmd.Dispose()
            Catch ex As Exception
            End Try
            Try
                cmd = Nothing
            Catch ex As Exception
            End Try

            ' Handle connection
            Try
                con.Close()
            Catch ex As Exception
            End Try
            Try
                SqlConnection.ClearPool(con)
            Catch ex As Exception
            End Try
            Try
                con.Dispose()
            Catch ex As Exception
            End Try
            Try
                con = Nothing
            Catch ex As Exception
            End Try

            ' Exit
            Return ErrMsg
        End Function

        ' ====================================================
        ' Other support functions
        Public Function CheckDBNull(ByVal obj As Object, _
          Optional ByVal ObjectType As enumObjectType = enumObjectType.StrType) As Object
            ' Checks an object to determine if its null, and if so sets it to a not-null empty value
            Dim objReturn As Object
            objReturn = obj
            If ObjectType = enumObjectType.StrType And IsDBNull(obj) Then
                objReturn = ""
            ElseIf ObjectType = enumObjectType.IntType And IsDBNull(obj) Then
                objReturn = 0
            ElseIf ObjectType = enumObjectType.DblType And IsDBNull(obj) Then
                objReturn = 0.0
            ElseIf ObjectType = enumObjectType.DteType And IsDBNull(obj) Then
                objReturn = Now
            End If
            Return objReturn
        End Function

        Public Sub writeoutput(ByVal fs As StreamWriter, ByVal instring As String)
            ' This function writes a line to a previously opened streamwriter, and then flushes it
            ' promptly.  This assists in debugging services
            fs.WriteLine(instring)
            fs.Flush()
        End Sub

        Public Sub writeoutputfs(ByVal fs As FileStream, ByVal instring As String)
            ' This function writes a line to a previously opened filestream, and then flushes it
            ' promptly.  This assists in debugging services
            fs.Write(StringToBytes(instring), 0, Len(instring))
            fs.Write(StringToBytes(vbCrLf), 0, 2)
            fs.Flush()
        End Sub

        Public Function StringToBytes(ByVal str As String) As Byte()
            ' Convert a random string to a byte array
            ' e.g. "abcdefg" to {a,b,c,d,e,f,g}
            Dim s As Char()
            s = str.ToCharArray
            Dim b(s.Length - 1) As Byte
            Dim i As Integer
            For i = 0 To s.Length - 1
                b(i) = Convert.ToByte(s(i))
            Next
            Return b
        End Function

        Public Function EncodeParamSpaces(ByVal InVal As String) As String
            ' If given a urlencoded parameter value, replace spaces with "+" signs

            Dim temp As String
            Dim i As Integer

            If InStr(InVal, " ") > 0 Then
                temp = ""
                For i = 1 To Len(InVal)
                    If Mid(InVal, i, 1) = " " Then
                        temp = temp & "+"
                    Else
                        temp = temp & Mid(InVal, i, 1)
                    End If
                Next
                EncodeParamSpaces = temp
            Else
                EncodeParamSpaces = InVal
            End If
        End Function
    End Class

    ' =================================================
    ' In Memory Cache Class
    Public Class HciDMSDocument
        Private _dataType As Type
        Private _updateDate As Date
        Private _docObj As Object
        Private _cachedObj As ObjectCache = MemoryCache.Default
        Public Sub New(uptDate As Date, docObj As Object)
            _dataType = docObj.GetType()
            _updateDate = uptDate
            _docObj = docObj
        End Sub
        Public Property DataType() As Type
            Get
                Return Me._dataType
            End Get
            Set(value As Type)
                _dataType = value
            End Set
        End Property
        Public Property UpdateDate() As Date
            Get
                Return Me._updateDate
            End Get
            Set(value As Date)
                _updateDate = value
            End Set
        End Property
        Public Property CachedObj() As Object
            Get
                Return System.Convert.ChangeType(_docObj, _dataType)
            End Get
            Set(value As Object)
                _docObj = value
            End Set
        End Property
    End Class

    ' =================================================
    ' HTTP PROXY CLASS
    Class simplehttp
        Public Function geturl(ByVal url As String, ByVal proxyip As String, ByVal port As Integer, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 5000
            req.CookieContainer = New CookieContainer()
            req.Referer = ""
            req.Headers.[Set]("Accept-Language", "en,en-us")
            Dim stream_in As StreamReader

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don’t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            req.Proxy = proxy

            Dim response As String = ""
            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                stream_in = New StreamReader(resp.GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
            End Try
            Return response
        End Function

        Public Function getposturl(ByVal url As String, ByVal postdata As String, ByVal proxyip As String, ByVal port As Short, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 5000
            req.CookieContainer = New CookieContainer()
            req.Method = "POST"
            req.ContentType = "application/x-www-form-urlencoded"
            req.ContentLength = postdata.Length
            req.Referer = ""

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don’t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            req.Proxy = proxy

            Dim stream_out As New StreamWriter(req.GetRequestStream(), System.Text.Encoding.ASCII)
            stream_out.Write(postdata)
            stream_out.Close()
            Dim response As String = ""

            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                Dim resStream As Stream = resp.GetResponseStream()
                Dim stream_in As New StreamReader(req.GetResponse().GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
            End Try
            Return response
        End Function
    End Class

    Public Delegate Function AsynchUpdDMSDocCount(ByVal CONTACT_ID As String, ByVal TRAINER_NUM As String, ByVal PART_ID As String, _
         ByVal MT_ID As String, ByVal Debug As String) As String

    Public Delegate Function AsynchUpdDMSDoc(ByVal CONTACT_ID As String, ByVal TRAINER_NUM As String, ByVal PART_ID As String, _
        ByVal MT_ID As String, ByVal Debug As String) As String

    Public Delegate Function AsynchSaveDMSDocAssoc(ByVal DocId As String, ByVal Association As String, _
        ByVal AssocKey As String, ByVal PrFlag As String, ByVal ReqdFlag As String, ByVal Rights As String, _
        ByVal Debug As String) As Boolean

    Public Delegate Function AsynchSaveDMSDocCat(ByVal DocId As String, ByVal Category As String, ByVal PrFlag As String, _
        ByVal Debug As String) As Boolean

    Public Delegate Function AsynchSaveDMSDocKey(ByVal DocId As String, ByVal DocKey As String, ByVal KeyVal As String, _
        ByVal PrFlag As String, ByVal Debug As String) As Boolean

    Public Delegate Function AsynchSaveDMSDocUser(ByVal DocId As String, ByVal Domain As String, _
        ByVal DomainOwner As String, ByVal DomainRights As String, ByVal SubId As String, _
        ByVal SubOwner As String, ByVal SubRights As String, ByVal ConId As String, ByVal ConOwner As String, _
        ByVal ConRights As String, ByVal RegId As String, ByVal RegOwner As String, ByVal RegRights As String, _
        ByVal Debug As String) As Boolean
End Namespace
