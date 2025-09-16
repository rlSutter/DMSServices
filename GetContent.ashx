<%@ WebHandler Language="VB" Class="GetContent" %>

Imports System
Imports System.Web
Imports System.Configuration
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Web.Script.Serialization
Imports System.Xml
Imports System.Net
Imports System.Text
Imports Amazon
Imports Amazon.S3
Imports System.Threading.Tasks
Imports log4net

Public Class GetContent : Implements IHttpHandler

    ' Service to retrieve a list of content from the DMS based on the supplied parameters
    ' If no parameters are supplied, this agent returns a list of all documents to which the user has access
    ' Based on the user access rights to a document (or the document function) the user will be given CRUD functionality

    ' Search Criteria:
    '		Name	===				Para.	Related to===								Notes===
    '		Employee Id				EID	
    '		Row Id					ID		DMS.Documents.row_id						The document id
    '	Affiliation
    '		Category Id				CTI		DMS.Categories.row_id						The id of the category
    '		Category Name			CTN		DMS.Categories.name							The name of the category
    '		Association Name		ASN		DMS.Associations.name						The name of the association
    '		Association Id			ASI		DMS.Associations.row_id						The id of the association
    '		Association Key			ASK		DMS.Document_Associations.fkey				The record key of the association
    '		Related Assoc			APM		DMS.Associations.row_id						Optional association id
    '		Related Assoc			APT		DMS.Associations.name						Optional association name
    '		Related Assoc			ADF		DMS.Document_Associations.fkey				Optional association key
    '		Description				DSC													Description
    '		Document Name			NAM													Document Name
    '	User ContextpUBL
    '		Domain					DMU		DMS.Groups.name								Flag to indicate domain-wide search
    '		Subscription			SBU		DMS.Groups.name								Flag to indicate sub-wide search
    '		User					USR		DMS.Users.ext_user_id						Flag to indicate user privileges
    '		Owner					OWN		DMS.Documents.created_by/last_upd_by		Flag to indicate user is creator/updater
    '		Alternate Sub			ASB		CX_SUBSCRIPTION.ROW_ID						Alternative subscription id
    '	Attributes
    '		Document Type			DCT		DMS.Document_Types.extension				Filename extension of document
    '		Deleted					DCR		DMS.Documents.deleted						Flag to indicate only deleted documents
    '	Full-text
    '		Description				FTX		DMS.Documents.description					Description of document contains specified text
    '	Context
    '		NOT_TYPE				NTY		Calling screen type
    '		CALL_SCREEN				CIS		Calling screen type
    '		CALL_ID					CID		Calling screen id

    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        ' Declare variables
        Dim ErrMsg, ErrLvl As String
        Dim PublicKey As String
        Dim Debug As String
        Dim outstring As String

        ' Search declarations
        Dim Extension, LANG_CD, temp As String
        Dim SDOC_ID, CAT_ID, KEY_ID, CAT_NAME, ASN_NAME, ASN_ID, ASN_KEY, ASN_OPT, KEY_NAME, CAT_NAME_S As String
        Dim DOC_EXT, DLT_FLG, FTEXT, StartDt, EndDt, DOC_NAME, DESC_TEXT, DOC_EXT_ID, TYPE_ADMIN As String
        Dim PUBLIC_FLG, DMS_FLG, SUB_FLG, USR_FLG, OWN_FLG, ALT_SUB_ID, DOC_LIB_FLG, DOC_AREA, ADMIN_ACCESS As String
        Dim OptAssoc, OptAssocKey, OptAssocId, PRIMARY_ASSOC, SearchScope, relink, PHeader, ButtonList, PTITLE As String
        Dim SortOrd, SortDir, GroupBy, OrderBy As String
        Dim DOMAIN, POPUP_FLG, CALL_ID, CALL_SCREEN, NOT_TYPE, CST_FLG, CAT_FLG, ASC_FLG, ASSOC_PARAM, ASSOC_PARAM_TYPE, ASSOC_DEF, ASSOC_RSTR As String
        Dim CONTACT_ID, DMS_DOMAIN_ID, DMS_ASUB_ID, SUB_ID, DMS_SUB_ID, EMP_ID, DMS_USER_ID, DMS_UA_ID, DMS_USER_AID, UID, OWNER_FLG, EDIT_FLG, DEL_FLG, OUTPUT_FLG As String
        Dim SYSADMIN_FLG, TRAINING_ACCESS, TRAINER_ACC_FLG, MT_FLG, SVC_TYPE, TRAINING_FLG, TRAINER_FLG, TRAINER_ID, PART_ID, PART_FLG, SITE_ONLY As String
        Dim SQuery As String = ""
        Dim TotalDocs As Integer
        Dim OrderID As String = ""
        Dim Refresh As Boolean = False
        Dim TableName As String = ""

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("PDDDebugLog")
        Dim logfile, tempdebug, temp3 As String
        Dim Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"
        Dim callip As String = context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If callip Is Nothing Then
            callip = context.Request.UserHostAddress
        Else
            If callip.Contains(",") Then
                callip = Left(callip, callip.IndexOf(",") - 1)
            Else
                callip = callip
            End If
        End If

        ' Context declarations
        Dim NextLink, BasePath As String
        Dim PrevLink As String = Trim(context.Request.ServerVariables("HTTP_REFERER"))
        Dim BROWSER As String = Trim(context.Request.ServerVariables("HTTP_USER_AGENT"))

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        Dim DMSServices As New local.hq.datafluxapp.dms.Service

        ' ============================================
        ' Variable setup
        Debug = "Y"
        Logging = "Y"
        PublicKey = ""
        Extension = ""
        BROWSER = ""
        Debug = "N"
        ErrMsg = ""
        LANG_CD = "ENU"
        ErrLvl = "Warning"
        DOMAIN = ""
        SDOC_ID = ""
        CAT_ID = ""
        KEY_ID = ""
        CAT_NAME = ""
        CAT_NAME_S = ""
        ASN_NAME = ""
        ASN_ID = ""
        ASN_KEY = ""
        ASN_OPT = ""
        KEY_NAME = ""
        DOC_AREA = ""
        DOC_EXT = ""
        DLT_FLG = ""
        DOC_EXT_ID = ""
        DMS_USER_ID = ""
        DMS_UA_ID = ""
        DMS_USER_AID = ""
        DMS_SUB_ID = ""
        SUB_ID = ""
        FTEXT = ""
        StartDt = ""
        EndDt = ""
        PRIMARY_ASSOC = ""
        OptAssoc = ""
        OptAssocKey = ""
        OptAssocId = ""
        DOC_NAME = ""
        DESC_TEXT = ""
        PUBLIC_FLG = ""
        DMS_FLG = ""
        SUB_FLG = ""
        USR_FLG = ""
        OWN_FLG = ""
        ALT_SUB_ID = ""
        DOC_LIB_FLG = ""
        EMP_ID = ""
        CONTACT_ID = ""
        SortOrd = ""
        SortDir = ""
        DOMAIN = ""
        POPUP_FLG = ""
        CALL_ID = ""
        CALL_SCREEN = ""
        NOT_TYPE = ""
        CST_FLG = ""
        CAT_FLG = ""
        ASC_FLG = ""
        ASSOC_PARAM = ""
        ASSOC_PARAM_TYPE = ""
        ASSOC_DEF = ""
        ASSOC_RSTR = ""
        TYPE_ADMIN = "N"
        DMS_DOMAIN_ID = ""
        DMS_ASUB_ID = ""
        OWNER_FLG = "N"
        UID = ""
        SYSADMIN_FLG = "N"
        TRAINING_ACCESS = ""
        TRAINER_ACC_FLG = "N"
        TRAINER_FLG = "N"
        MT_FLG = "N"
        SVC_TYPE = ""
        TRAINING_FLG = "N"
        TRAINER_ID = ""
        PART_ID = ""
        PART_FLG = "N"
        SITE_ONLY = "Y"
        SearchScope = "P"
        EDIT_FLG = "N"
        DEL_FLG = "N"
        OUTPUT_FLG = "N"
        ADMIN_ACCESS = "N"
        temp = ""
        temp3 = ""
        GroupBy = ""
        OrderBy = ""
        NextLink = ""
        outstring = ""
        PHeader = ""
        ButtonList = ""
        PTITLE = ""
        BasePath = "https://w4.certegrity.com/dmsa.nsf"

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidbro").ConnectionString
            If ConnS = "" Then ConnS = "server="
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server="
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("GetContent_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            ErrMsg = ErrMsg & vbCrLf & "Unable to get defaults from web.config: " & ex.Message
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\GetContent.log"
            Try
                log4net.GlobalContext.Properties("PDDLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                ErrMsg = ErrMsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If

        ' ============================================
        ' GET PARAMETERS
        ' Refresh
        If Not context.Request.QueryString("RFR") Is Nothing Then temp = DeURL(context.Request.QueryString("RFR"))
        If temp <> "" Then
            Refresh = True
            OrderID = temp
        Else
            Randomize()
            OrderID = "DCL" & Chr(Str(Int(Rnd() * 26)) + 65) & Trim(Str(Hour(Now))) & Trim(Str(Minute(Now))) & Trim(Str(Second(Now))) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65)
        End If

        '	Criteria filters
        If Not context.Request.QueryString("ID") Is Nothing Then SDOC_ID = context.Request.QueryString("ID")                ' Document id
        If Not context.Request.QueryString("DOC_ID") Is Nothing Then SDOC_ID = context.Request.QueryString("DOC_ID")        ' Document id
        If Not context.Request.QueryString("CTI") Is Nothing Then CAT_ID = DeURL(context.Request.QueryString("CTI"))        ' The id of the category
        If Not context.Request.QueryString("KEY") Is Nothing Then KEY_ID = DeURL(context.Request.QueryString("KEY"))        ' The id of the keyword	
        If Not context.Request.QueryString("KYN") Is Nothing Then KEY_NAME = DeURL(context.Request.QueryString("KYN"))      ' The keyword	
        If Not context.Request.QueryString("CTN") Is Nothing Then CAT_NAME = context.Request.QueryString("CTN")             ' The name of the category
        If Not context.Request.QueryString("ASN") Is Nothing Then ASN_NAME = context.Request.QueryString("ASN")             ' The name of the association
        If Not context.Request.QueryString("ASI") Is Nothing Then ASN_ID = context.Request.QueryString("ASI")               ' The id of the association
        If Not context.Request.QueryString("ASK") Is Nothing Then ASN_KEY = DeURL(Trim(context.Request.QueryString("ASK"))) ' The record key of the association
        If Not context.Request.QueryString("AOP") Is Nothing Then ASN_OPT = context.Request.QueryString("AOP")              ' The association is optional
        If Not context.Request.QueryString("DCT") Is Nothing Then DOC_EXT = LCase(context.Request.QueryString("DCT"))       ' Filename extension of document
        If Not context.Request.QueryString("DCR") Is Nothing Then DLT_FLG = UCase(context.Request.QueryString("DCR"))       ' Flag to indicate only deleted documents
        If Not context.Request.QueryString("FTX") Is Nothing Then FTEXT = context.Request.QueryString("FTX")                ' Description of document contains specified text
        If Not context.Request.QueryString("DTH") Is Nothing Then StartDt = context.Request.QueryString("DTH")              ' Created start date
        If Not context.Request.QueryString("EDT") Is Nothing Then EndDt = context.Request.QueryString("EDT")                ' Created end date	
        If Not context.Request.QueryString("NME") Is Nothing Then DOC_NAME = DeURL(context.Request.QueryString("NME"))      ' Document Name
        If Not context.Request.QueryString("DSC") Is Nothing Then DESC_TEXT = DeURL(context.Request.QueryString("DSC"))     ' Description Text	
        If Not context.Request.QueryString("SS") Is Nothing Then SearchScope = DeURL(context.Request.QueryString("SS"))     ' Search Scope = "M"/"P"/"A"

        '	Access filters
        If Not context.Request.QueryString("PUB") Is Nothing Then PUBLIC_FLG = UCase(context.Request.QueryString("PUB"))    ' Flag to indicate public only search
        If Not context.Request.QueryString("DMU") Is Nothing Then DMS_FLG = UCase(context.Request.QueryString("DMU"))       ' Flag to indicate domain-wide search
        If Not context.Request.QueryString("SBU") Is Nothing Then SUB_FLG = UCase(context.Request.QueryString("SBU"))       ' Flag to indicate sub-wide search
        If Not context.Request.QueryString("USR") Is Nothing Then USR_FLG = UCase(context.Request.QueryString("USR"))       ' Flag to indicate user specific search
        If Not context.Request.QueryString("OWN") Is Nothing Then OWN_FLG = UCase(context.Request.QueryString("OWN"))       ' Flag to indicate that user is creator
        If Not context.Request.QueryString("SUB_ID") Is Nothing Then SUB_ID = context.Request.QueryString("SUB_ID")         ' Alternative subscription id for searches		
        If Not context.Request.QueryString("ASB") Is Nothing Then ALT_SUB_ID = context.Request.QueryString("ASB")           ' Alternative subscription id for searches		
        If Not context.Request.QueryString("DLF") Is Nothing Then DOC_LIB_FLG = context.Request.QueryString("DLF")          ' Flag to indicate that we are seeking all docs associated with the current user
        If Not context.Request.QueryString("EID") Is Nothing Then EMP_ID = context.Request.QueryString("EID")               ' Employee Id of user	

        '	Sorting 
        If Not context.Request.QueryString("SO") Is Nothing Then SortOrd = context.Request.QueryString("SO")                ' Sort Order
        If Not context.Request.QueryString("SD") Is Nothing Then SortDir = context.Request.QueryString("SD")                ' Sort Direction

        '	Context
        If Not context.Request.QueryString("DOM") Is Nothing Then DOMAIN = context.Request.QueryString("DOM")               ' Domain
        If Not context.Request.QueryString("POP") Is Nothing Then POPUP_FLG = context.Request.QueryString("POP")            ' Popup window flag
        If Not context.Request.QueryString("CID") Is Nothing Then CALL_ID = context.Request.QueryString("CID")              ' Calling Id
        If Not context.Request.QueryString("CIS") Is Nothing Then CALL_SCREEN = context.Request.QueryString("CIS")          ' Calling Screen/Type
        If Not context.Request.QueryString("NTY") Is Nothing Then NOT_TYPE = context.Request.QueryString("NTY")             ' Calling Screen/Type	
        If Not context.Request.QueryString("CST") Is Nothing Then CST_FLG = context.Request.QueryString("CST")              ' Customer Service Documents flag
        If Not context.Request.QueryString("CFL") Is Nothing Then CAT_FLG = context.Request.QueryString("CFL")              ' Display categories flag
        If Not context.Request.QueryString("AFL") Is Nothing Then ASC_FLG = context.Request.QueryString("AFL")              ' Display associations flag
        If Not context.Request.QueryString("APM") Is Nothing Then ASSOC_PARAM = context.Request.QueryString("APM")          ' Optional Association Id
        If Not context.Request.QueryString("APT") Is Nothing Then ASSOC_PARAM_TYPE = DeURL(context.Request.QueryString("APT")) ' Optional Association Name
        If Not context.Request.QueryString("ADF") Is Nothing Then ASSOC_DEF = DeURL(context.Request.QueryString("ADF"))     ' Optional Association key
        If Not context.Request.QueryString("ASR") Is Nothing Then ASSOC_RSTR = context.Request.QueryString("ASR")           ' Associated Restricted 

        ' ============================================
        ' Fix parameters
        If DOC_LIB_FLG = "Y" Then
            ASN_NAME = "Contact"
            DMS_FLG = "Y"
            SUB_FLG = "Y"
            USR_FLG = "Y"
        End If
        If LCase(CAT_ID) = "any" Then CAT_ID = ""
        If LCase(KEY_ID) = "any" Then KEY_ID = ""
        If LCase(DOC_EXT) = "any" Then DOC_EXT = ""
        If InStr(1, CAT_NAME, "%20") > 0 Then CAT_NAME = rspURL(CAT_NAME)
        If InStr(1, ASN_NAME, "%20") > 0 Then ASN_NAME = rspURL(ASN_NAME)
        If InStr(1, ASN_KEY, " ") > 0 Then ASN_KEY = EnURL(ASN_KEY)
        If InStr(1, ASN_KEY, "%2B") > 0 Then ASN_KEY = Replace(ASN_KEY, "%2B", "+")
        If InStr(1, ASSOC_DEF, " ") > 0 Then ASSOC_DEF = EnURL(ASSOC_DEF)
        If InStr(1, ASSOC_DEF, "%2B") > 0 Then ASSOC_DEF = Replace(ASSOC_DEF, "%2B", "+")
        If InStr(1, FTEXT, "%20") > 0 Then FTEXT = rspURL(FTEXT)
        If DMS_FLG <> "Y" And SUB_FLG <> "Y" And USR_FLG <> "Y" Then OWN_FLG = "Y"
        If DMS_FLG = "" Then DMS_FLG = "N"
        If SUB_FLG = "" Then SUB_FLG = "N"
        If USR_FLG = "" Then USR_FLG = "N"
        If POPUP_FLG = "" Then POPUP_FLG = "N"
        If CAT_FLG = "" Then CAT_FLG = "N"
        If ASC_FLG = "" Then ASC_FLG = "N"
        If CALL_SCREEN = "1" Or CALL_SCREEN = "" Then CALL_SCREEN = "CNT"
        If NOT_TYPE = "" And CALL_SCREEN <> "" Then NOT_TYPE = CALL_SCREEN
        If NOT_TYPE <> "" And CALL_SCREEN = "" Then CALL_SCREEN = NOT_TYPE
        If ASN_OPT = "" Then ASN_OPT = "N"

        ' ============================================
        ' VALIDATE PARAMETERS
        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  OrderID: " & OrderID)
            mydebuglog.Debug("  Refresh: " & Refresh.ToString)
            mydebuglog.Debug("  SDOC_ID: " & SDOC_ID)
            mydebuglog.Debug("  CAT_ID: " & CAT_ID)
            mydebuglog.Debug("  KEY_ID: " & KEY_ID)
            mydebuglog.Debug("  CAT_NAME: " & CAT_NAME)
            mydebuglog.Debug("  ASN_NAME: " & ASN_NAME)
            mydebuglog.Debug("  ASN_ID: " & ASN_ID)
            mydebuglog.Debug("  ASN_KEY: " & ASN_KEY)
            mydebuglog.Debug("  ASN_OPT: " & ASN_OPT)
            mydebuglog.Debug("  DOC_EXT: " & DOC_EXT)
            mydebuglog.Debug("  DLT_FLG: " & DLT_FLG)
            mydebuglog.Debug("  FTEXT: " & FTEXT)
            mydebuglog.Debug("  StartDt: " & StartDt)
            mydebuglog.Debug("  EndDt: " & EndDt)
            mydebuglog.Debug("  DOC_NAME: " & DOC_NAME)
            mydebuglog.Debug("  DESC_TEXT: " & DESC_TEXT)
            mydebuglog.Debug("  PUBLIC_FLG: " & PUBLIC_FLG)
            mydebuglog.Debug("  DMS_FLG: " & DMS_FLG)
            mydebuglog.Debug("  SUB_FLG: " & SUB_FLG)
            mydebuglog.Debug("  USR_FLG: " & USR_FLG)
            mydebuglog.Debug("  OWN_FLG: " & OWN_FLG)
            mydebuglog.Debug("  ALT_SUB_ID: " & ALT_SUB_ID)
            mydebuglog.Debug("  DOC_LIB_FLG: " & DOC_LIB_FLG)
            mydebuglog.Debug("  EMP_ID: " & EMP_ID)
            mydebuglog.Debug("  SortOrd: " & SortOrd)
            mydebuglog.Debug("  SortDir: " & SortDir)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  POPUP_FLG: " & POPUP_FLG)
            mydebuglog.Debug("  CALL_ID: " & CALL_ID)
            mydebuglog.Debug("  CALL_SCREEN: " & CALL_SCREEN)
            mydebuglog.Debug("  NOT_TYPE: " & NOT_TYPE)
            mydebuglog.Debug("  CST_FLG: " & CST_FLG)
            mydebuglog.Debug("  CAT_FLG: " & CAT_FLG)
            mydebuglog.Debug("  ASC_FLG: " & ASC_FLG)
            mydebuglog.Debug("  ASSOC_PARAM: " & ASSOC_PARAM)
            mydebuglog.Debug("  ASSOC_PARAM_TYPE: " & ASSOC_PARAM_TYPE)
            mydebuglog.Debug("  ASSOC_DEF: " & ASSOC_DEF)
            mydebuglog.Debug("  ASSOC_RSTR: " & ASSOC_RSTR)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  BROWSER: " & BROWSER & vbCrLf)
        End If

        If EMP_ID = "" And Not Refresh Then
            ErrLvl = "Error"
            Select Case LANG_CD
                Case "ESN"
                    ErrMsg = "Solicitud no v&aacute;lida. No se ha especificado ning&uacute;n ID de documento o versi&oacute;n."
                Case Else
                    ErrMsg = "Invalid request. No employee id specified."
            End Select
            GoTo DisplayErrorMsg
        End If

        ' ============================================
        ' START QUERY
        ' If invoked via a parameter, an ASN_NAME or ASN_ID is provided and no ASN_KEY is provided, then redirect to a 
        ' search form
        If (ASN_NAME <> "" Or ASN_ID <> "") And ASN_KEY = "" Then
            If SDOC_ID <> "" Then temp3 = temp3 & "&ID=" & SDOC_ID
            If CAT_ID <> "" Then temp3 = temp3 & "&CTI=" & CAT_ID
            If KEY_ID <> "" Then temp3 = temp3 & "&KEY=" & KEY_ID
            If CAT_NAME <> "" Then temp3 = temp3 & "&CTN=" & EnURL(CAT_NAME)
            If ASN_NAME <> "" Then temp3 = temp3 & "&ASN=" & EnURL(ASN_NAME)
            If ASN_ID <> "" Then temp3 = temp3 & "&ASI=" & EnURL(ASN_ID)
            If ASN_KEY <> "" Then temp3 = temp3 & "&ASK=" & EnURL(ASN_KEY)
            If DOC_EXT <> "" Then temp3 = temp3 & "&DCT=" & DOC_EXT
            If DLT_FLG <> "" Then temp3 = temp3 & "&DCR=" & DLT_FLG
            If ALT_SUB_ID <> "" Then temp3 = temp3 & "&ASB=" & ALT_SUB_ID
            If FTEXT <> "" Then temp3 = temp3 & "&FTX=" & EnURL(FTEXT)
            If StartDt <> "" Then temp3 = temp3 & "&DTH=" & StartDt
            If EndDt <> "" Then temp3 = temp3 & "&EDT=" & EndDt
            If DOC_NAME <> "" Then temp3 = temp3 & "&NME=" & EnURL(DOC_NAME)
            If DESC_TEXT <> "" Then temp3 = temp3 & "&DSC=" & EnURL(DESC_TEXT)
            If DMS_FLG <> "" Then temp3 = temp3 & "&DMU=" & DMS_FLG
            If SUB_FLG <> "" Then temp3 = temp3 & "&SBU=" & SUB_FLG
            If USR_FLG <> "" Then temp3 = temp3 & "&USR=" & USR_FLG
            If OWN_FLG <> "" Then temp3 = temp3 & "&OWN=" & OWN_FLG
            If SortOrd <> "" Then temp3 = temp3 & "&SO=" & EnURL(SortOrd)
            If SortDir <> "" Then temp3 = temp3 & "&SD=" & EnURL(SortDir)
            If DOMAIN <> "" Then temp3 = temp3 & "&DOM=" & DOMAIN
            If POPUP_FLG <> "" Then temp3 = temp3 & "&POP=" & POPUP_FLG
            If OrderID <> "" Then temp3 = temp3 & "&CID=" & OrderID
            If CST_FLG <> "" Then temp3 = temp3 & "&CST=" & CST_FLG
            If CAT_FLG <> "" Then temp3 = temp3 & "&CFL=" & CAT_FLG
            If ASC_FLG <> "" Then temp3 = temp3 & "&AFL=" & ASC_FLG
            If ASSOC_PARAM <> "" Then temp3 = temp3 & "&APM=" & ASSOC_PARAM
            If ASSOC_PARAM_TYPE <> "" Then temp3 = temp3 & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
            If ASSOC_DEF <> "" Then temp3 = temp3 & "&ADF=" & ASSOC_DEF
            If ASSOC_RSTR <> "" Then temp3 = temp3 & "&ASR=" & ASSOC_RSTR
            If EMP_ID <> "" Then temp3 = temp3 & "&EID=" & EMP_ID
            If CALL_ID <> "" Then temp3 = temp3 & "&CID=" & CALL_ID
            If CALL_SCREEN <> "" Then temp3 = temp3 & "&CIS=" & CALL_SCREEN
            If NOT_TYPE <> "" Then temp3 = temp3 & "&NTY=" & NOT_TYPE
            NextLink = BasePath & "/ReviewDocument?EditDocument" & temp3

            outstring = "Content-Type:text/plain" & vbCrLf
            outstring = outstring & "Content-Type:text/html" & vbCrLf
            outstring = outstring & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">" & vbCrLf
            outstring = outstring & "<HTML>" & vbCrLf
            outstring = outstring & "<HEAD>" & vbCrLf
            outstring = outstring & "<TITLE>Certification Manager</TITLE>" & vbCrLf
            outstring = outstring & "<META HTTP-EQUIV=""Refresh"" CONTENT=""0; URL=" & NextLink & """>" & vbCrLf
            outstring = outstring & "</HEAD>" & vbCrLf
            outstring = outstring & "</HTML>" & vbCrLf
            GoTo CloseOut
        End If

PerformSearch:
        ' ----
        ' Compute invocation link
        temp3 = ""
        If SDOC_ID <> "" Then temp3 = temp3 & "&ID=" & SDOC_ID
        If CAT_ID <> "" Then temp3 = temp3 & "&CTI=" & CAT_ID
        If KEY_ID <> "" Then temp3 = temp3 & "&KEY=" & KEY_ID
        If CAT_NAME <> "" Then temp3 = temp3 & "&CTN=" & EnURL(CAT_NAME)
        If ASN_NAME <> "" Then temp3 = temp3 & "&ASN=" & EnURL(ASN_NAME)
        If ASN_ID <> "" Then temp3 = temp3 & "&ASI=" & EnURL(ASN_ID)
        If ASN_KEY <> "" Then temp3 = temp3 & "&ASK=" & EnURL(ASN_KEY)
        If DOC_EXT <> "" Then temp3 = temp3 & "&DCT=" & DOC_EXT
        If DLT_FLG <> "" Then temp3 = temp3 & "&DCR=" & DLT_FLG
        If ALT_SUB_ID <> "" Then temp3 = temp3 & "&ASB=" & ALT_SUB_ID
        If FTEXT <> "" Then temp3 = temp3 & "&FTX=" & EnURL(FTEXT)
        If StartDt <> "" Then temp3 = temp3 & "&DTH=" & StartDt
        If EndDt <> "" Then temp3 = temp3 & "&EDT=" & EndDt
        If DOC_NAME <> "" Then temp3 = temp3 & "&NME=" & EnURL(DOC_NAME)
        If DESC_TEXT <> "" Then temp3 = temp3 & "&DSC=" & EnURL(DESC_TEXT)
        If DMS_FLG <> "" Then temp3 = temp3 & "&DMU=" & DMS_FLG
        If SUB_FLG <> "" Then temp3 = temp3 & "&SBU=" & SUB_FLG
        If USR_FLG <> "" Then temp3 = temp3 & "&USR=" & USR_FLG
        If OWN_FLG <> "" Then temp3 = temp3 & "&OWN=" & OWN_FLG
        If SortOrd <> "" Then temp3 = temp3 & "&SO=" & EnURL(SortOrd)
        If SortDir <> "" Then temp3 = temp3 & "&SD=" & EnURL(SortDir)
        If DOMAIN <> "" Then temp3 = temp3 & "&DOM=" & DOMAIN
        If POPUP_FLG <> "" Then temp3 = temp3 & "&POP=" & POPUP_FLG
        If OrderID <> "" Then temp3 = temp3 & "&CID=" & OrderID
        If CST_FLG <> "" Then temp3 = temp3 & "&CST=" & CST_FLG
        If CAT_FLG <> "" Then temp3 = temp3 & "&CFL=" & CAT_FLG
        If ASC_FLG <> "" Then temp3 = temp3 & "&AFL=" & ASC_FLG
        If ASSOC_PARAM <> "" Then temp3 = temp3 & "&APM=" & ASSOC_PARAM
        If ASSOC_PARAM_TYPE <> "" Then temp3 = temp3 & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
        If ASSOC_DEF <> "" Then temp3 = temp3 & "&ADF=" & ASSOC_DEF
        If ASSOC_RSTR <> "" Then temp3 = temp3 & "&ASR=" & ASSOC_RSTR
        If EMP_ID <> "" Then temp3 = temp3 & "&EID=" & EMP_ID
        If CALL_ID <> "" Then temp3 = temp3 & "&CID=" & OrderID
        If CALL_SCREEN <> "" Then temp3 = temp3 & "&CIS=" & CALL_SCREEN
        If NOT_TYPE <> "" Then temp3 = temp3 & "&NTY=" & NOT_TYPE
        NextLink = "GetContent.ashx?" & temp3
        If Debug = "Y" Then
            mydebuglog.Debug("Computed NextLink: " & NextLink & vbCrLf)
        End If

        ' ============================================
        ' Open database connection 
        ErrMsg = OpenDBConnection(ConnS, con, cmd)
        If ErrMsg <> "" Then
            GoTo DBError
        End If
        If cmd Is Nothing Then GoTo DBError
        ErrMsg = OpenDBConnection(d_ConnS, d_con, d_cmd)    ' dms
        If ErrMsg <> "" Then
            GoTo DBError
        End If
        If d_cmd Is Nothing Then GoTo DBError

        ' ================================================
        ' GET ACCESS INFORMATION
        ' Set Document Library flags
        If DOC_LIB_FLG = "Y" Then ASN_KEY = EMP_ID

        ' ================================================
        ' HANDLE REFRESH
        If Refresh And OrderID <> "" Then
            SqlS = "SELECT USER_ID, FULL_QUERY, EDIT_FLG, DEL_FLG, ADMIN_ACCESS, SUB_FLG, USR_FLG " &
            "FROM reports.dbo.CMREP " &
            "WHERE ORDER_ID='" & OrderID & "'"
            If Debug = "Y" Then mydebuglog.Debug("  Get query to refresh: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        EMP_ID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                        SQuery = Trim(CheckDBNull(d_dr(1), enumObjectType.StrType))
                        EDIT_FLG = Trim(CheckDBNull(d_dr(2), enumObjectType.StrType))
                        DEL_FLG = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                        ADMIN_ACCESS = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                        SUB_FLG = Trim(CheckDBNull(d_dr(5), enumObjectType.StrType))
                        USR_FLG = Trim(CheckDBNull(d_dr(6), enumObjectType.StrType))
                    End While
                Else
                    ErrMsg = ErrMsg & "Error getting refresh query." & vbCrLf
                End If
                d_dr.Close()
            Catch ex As Exception
                ErrMsg = ErrMsg & "Error getting refresh query: " & ex.Message & vbCrLf
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug("  > EMP_ID: " & EMP_ID)
                mydebuglog.Debug("  > EDIT_FLG, DEL_FLG, ADMIN_ACCESS, SUB_FLG, USR_FLG: " & EDIT_FLG & "/" & DEL_FLG & "/" & ADMIN_ACCESS & "/" & SUB_FLG & "/" & USR_FLG)
                mydebuglog.Debug("  > SQuery: " & SQuery & vbCrLf)
            End If
        End If

        ' ================================================
        ' GET DMS USER INFORMATION
        ' Get user id
        If DMS_USER_ID = "" Then
            SqlS = "SELECT U.row_id, UGA.row_id, E.X_CON_ID, SC.SUB_ID, S.DOMAIN, C.X_REGISTRATION_NUM, SC.SYSADMIN_FLG, SC.TRAINING_ACCESS, SC.TRAINER_ACC_FLG, " &
                "C.X_MAST_TRNR_FLG, S.SVC_TYPE, C.X_TRAINER_NUM, C.X_PART_ID, SC.SITE_ONLY_FLG " &
                "FROM DMS.dbo.Users U " &
                "LEFT OUTER JOIN DMS.dbo.User_Group_Access UGA on UGA.access_id=U.row_id And UGA.type_id='U' " &
                "LEFT OUTER JOIN siebeldb.dbo.S_EMPLOYEE E ON E.X_CON_ID=U.ext_user_id " &
                "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID=E.X_CON_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=E.X_CON_ID " &
                "WHERE E.ROW_ID='" & EMP_ID & "' AND UGA.row_id IS NOT NULL"
            If Debug = "Y" Then mydebuglog.Debug("Get user information: " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        DMS_USER_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        DMS_UA_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        DMS_USER_AID = DMS_UA_ID
                        CONTACT_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        SUB_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        DOMAIN = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        UID = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        SYSADMIN_FLG = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                        TRAINING_ACCESS = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                        TRAINER_ACC_FLG = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                        MT_FLG = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                        SVC_TYPE = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                        TRAINER_ID = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                        If TRAINER_ID <> "" Then TRAINER_FLG = "Y"
                        PART_ID = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                        If PART_ID <> "" Then PART_FLG = "Y"
                        Select Case SVC_TYPE
                            Case "CERTIFICATION MANAGER REG DB"
                                TRAINING_FLG = "N"
                            Case "PUBLIC ACCESS"
                                TRAINING_FLG = "Y"
                            Case Else
                                TRAINING_FLG = "Y"
                        End Select
                        SITE_ONLY = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                    End While
                Else
                    ErrMsg = ErrMsg & "Error getting user information." & vbCrLf
                End If
                dr.Close()
            Catch ex As Exception
                ErrMsg = ErrMsg & "Error getting user information: " & ex.Message & vbCrLf
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug("  > CONTACT_ID: " & CONTACT_ID)
                mydebuglog.Debug("  > UID: " & UID)
                mydebuglog.Debug("  > SUB_ID: " & SUB_ID)
                mydebuglog.Debug("  > SYSADMIN_FLG: " & SYSADMIN_FLG)
                mydebuglog.Debug("  > TRAINING_ACCESS: " & TRAINING_ACCESS)
                mydebuglog.Debug("  > TRAINER_ACC_FLG: " & TRAINER_ACC_FLG)
                mydebuglog.Debug("  > MT_FLG: " & MT_FLG)
                mydebuglog.Debug("  > SVC_TYPE: " & SVC_TYPE)
                mydebuglog.Debug("  > SITE_ONLY: " & SITE_ONLY)
                mydebuglog.Debug("  > DOMAIN: " & DOMAIN)
                mydebuglog.Debug("  > DMS_USER_ID: " & DMS_USER_ID)
                mydebuglog.Debug("  > DMS_UA_ID: " & DMS_UA_ID)
            End If
            If DMS_USER_ID = "" Then DMS_USER_ID = "1"
        End If
        If DMS_UA_ID = "" Or CONTACT_ID = "" Then GoTo AccessError

        ' Get subscription user access id 
        If DMS_SUB_ID = "" Then
            SqlS = "SELECT UA.row_id " &
            "FROM DMS.dbo.User_Group_Access UA " &
            "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " &
            "WHERE UA.type_id='G' AND G.name='" & SUB_ID & "'"
            DMS_SUB_ID = GetSingleRecord("Get dms subscription id", "Groups", cmd, SqlS, mydebuglog, Debug)
            If Debug = "Y" Then mydebuglog.Debug("  > DMS_SUB_ID: " & DMS_SUB_ID)
        End If

        ' Get alt subscription user access id        
        If ALT_SUB_ID <> "" Then
            SqlS = "SELECT UA.row_id " &
            "FROM DMS.dbo.User_Group_Access UA " &
            "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " &
            "WHERE UA.type_id='G' AND G.name='" & ALT_SUB_ID & "'"
            DMS_ASUB_ID = GetSingleRecord("Get dms alt subscription id", "Groups", cmd, SqlS, mydebuglog, Debug)
            If Debug = "Y" Then mydebuglog.Debug("  > DMS_ASUB_ID: " & DMS_ASUB_ID)
        End If

        ' Get domain user access id
        If DOMAIN <> "" Then
            SqlS = "SELECT UA.row_id " &
            "FROM DMS.dbo.User_Group_Access UA " &
            "INNER JOIN DMS.dbo.Groups G ON G.row_id=UA.access_id " &
            "WHERE UA.type_id='G' AND G.name='" & DOMAIN & "'"
            DMS_DOMAIN_ID = GetSingleRecord("Get dms domain id", "Groups", cmd, SqlS, mydebuglog, Debug)
            If Debug = "Y" Then mydebuglog.Debug("  > DMS_DOMAIN_ID: " & DMS_DOMAIN_ID & vbCrLf)
        End If
        If DMS_DOMAIN_ID = "" Then GoTo AccessError

        ' Redirect on refresh
        If Refresh Then GoTo GetRecords

        ' ================================================
        ' ENRICH SUPPLIED PARAMETERS AS NEEDED - QUERIES FASTER FROM KEYS
        ' Get association
        If Debug = "Y" Then mydebuglog.Debug("CAT_NAME: " & CAT_NAME)
        If ASN_NAME = "" Or ASN_ID = "" Then
            ' Get association id
            If ASN_NAME = "" And ASN_ID <> "" Then
                SqlS = "SELECT name FROM DMS.dbo.Associations WHERE row_id='" & ASN_ID & "'"
                ASN_NAME = GetSingleRecord("Get association name", "Associations", cmd, SqlS, mydebuglog, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  > ASN_NAME: " & ASN_NAME)
            End If
            ' Get association name
            If ASN_NAME <> "" And ASN_ID = "" Then
                SqlS = "SELECT row_id FROM DMS.dbo.Associations WHERE name='" & ASN_NAME & "'"
                ASN_ID = GetSingleRecord("Get association id", "Associations", cmd, SqlS, mydebuglog, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  > ASN_ID: " & ASN_ID)
            End If
        End If
        ' Get optional association
        If ASSOC_PARAM = "" Or ASSOC_PARAM_TYPE = "" Then
            ' Get optional association id
            If ASSOC_PARAM_TYPE = "" And ASSOC_PARAM <> "" Then
                SqlS = "SELECT name FROM DMS.dbo.Associations WHERE row_id='" & ASSOC_PARAM & "'"
                ASSOC_PARAM_TYPE = GetSingleRecord("Get opt association name", "Associations", cmd, SqlS, mydebuglog, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  > ASSOC_PARAM_TYPE: " & ASSOC_PARAM_TYPE)
            End If
            ' Get optional association name
            If ASSOC_PARAM_TYPE <> "" And ASSOC_PARAM = "" Then
                SqlS = "SELECT row_id FROM DMS.dbo.Associations WHERE name='" & ASSOC_PARAM_TYPE & "'"
                ASSOC_PARAM = GetSingleRecord("Get opt association id", "Associations", cmd, SqlS, mydebuglog, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  > ASSOC_PARAM: " & ASSOC_PARAM)
            End If
        End If
        ' Get category
        If CAT_ID = "" Or CAT_NAME = "" Then
            ' Get category id
            If CAT_ID = "" And CAT_NAME <> "" And CAT_NAME <> "any" Then
                SqlS = "SELECT row_id FROM DMS.dbo.Categories WHERE name='" & HttpUtility.UrlDecode(CAT_NAME) & "'"
                CAT_ID = GetSingleRecord("Get category id", "Categories", cmd, SqlS, mydebuglog, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  > CAT_ID: " & CAT_ID)
            End If
            ' Get category name
            If CAT_ID <> "" And CAT_NAME = "" Then
                SqlS = "SELECT name FROM DMS.dbo.Categories WHERE row_id='" & CAT_ID & "'"
                CAT_NAME = GetSingleRecord("Get category name", "Categories", cmd, SqlS, mydebuglog, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  > CAT_NAME: " & CAT_NAME)
            End If
            If CAT_ID = "" And CAT_NAME <> "" And CAT_NAME <> "any" Then GoTo CatNotFound
        End If
        ' Get keyword
        If KEY_ID = "" Or KEY_NAME = "" Then
            ' Get Keyword Id
            If KEY_ID = "" And KEY_NAME <> "" And KEY_NAME <> "any" Then
                SqlS = "SELECT row_id FROM DMS.dbo.Keywords WHERE name='" & HttpUtility.UrlDecode(KEY_NAME) & "'"
                KEY_ID = GetSingleRecord("Get keyword id", "Keywords", cmd, SqlS, mydebuglog, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  > KEY_ID: " & KEY_ID)
            End If
            ' Get Keyword name
            If KEY_ID <> "" And KEY_NAME = "" Then
                SqlS = "SELECT name FROM DMS.dbo.Keywords WHERE row_id=" & KEY_ID
                KEY_NAME = GetSingleRecord("Get keyword name", "Keywords", cmd, SqlS, mydebuglog, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  > KEY_NAME: " & KEY_NAME)
            End If
        End If
        ' Get category singular name
        If CAT_ID <> "" And CAT_NAME_S = "" Then
            SqlS = "SELECT sname FROM DMS.dbo.Categories WHERE row_id='" & CAT_ID & "'"
            CAT_NAME_S = GetSingleRecord("Get category singular name", "Categories", cmd, SqlS, mydebuglog, Debug)
            If Debug = "Y" Then mydebuglog.Debug("  > CAT_NAME_S: " & CAT_NAME_S)
        End If
        ' Get extension 
        If DOC_EXT <> "" Then
            SqlS = "SELECT row_id FROM DMS.dbo.Document_Types WHERE extension='" & DOC_EXT & "'"
            DOC_EXT_ID = GetSingleRecord("Get extension id", "Categories", cmd, SqlS, mydebuglog, Debug)
            If Debug = "Y" Then mydebuglog.Debug("  > DOC_EXT_ID: " & DOC_EXT_ID & vbCrLf)
        End If

        ' ================================================
        ' GET PRIMARY ASSOCIATION IF SUPPLIED              
        Dim AssocFilter As String = ""
        If ASN_NAME <> "" And ASN_KEY <> "" Then
            SqlS = ""
            PRIMARY_ASSOC = GetDocFilter(cmd, dr, ASN_NAME, ASN_KEY, ASN_ID, UID, DOMAIN, SUB_ID, CONTACT_ID, OptAssoc, OptAssocKey, OptAssocId, OWNER_FLG, SqlS, PUBLIC_FLG, EDIT_FLG, DEL_FLG, ADMIN_ACCESS, OUTPUT_FLG, "N", DOC_AREA, mydebuglog, Debug)
            AssocFilter = SqlS
        End If
        If SYSADMIN_FLG = "Y" Then
            ADMIN_ACCESS = "Y"
            EDIT_FLG = "Y"
            DEL_FLG = "Y"
            OUTPUT_FLG = "Y"
        End If
        If Debug = "Y" Then
            mydebuglog.Debug("  ... AssocFilter: " & AssocFilter)
            mydebuglog.Debug("  ... DOC_AREA: " & DOC_AREA)
            mydebuglog.Debug("  ... EDIT_FLG, DEL_FLG, OUTPUT_FLG: " & EDIT_FLG & "/" & DEL_FLG & "/" & OUTPUT_FLG)
            mydebuglog.Debug("  ... ADMIN_ACCESS: " & ADMIN_ACCESS)
            mydebuglog.Debug("  ... ASSOCIATION OWNER_FLG: " & OWNER_FLG)
            mydebuglog.Debug("  ... PRIMARY_ASSOC: " & PRIMARY_ASSOC & vbCrLf)
        End If

        ' ================================================
        ' GET CATEGORY CONSTRAINT	
        ' Check primary category rights stored in Category_Keywords
        Dim Category_Constraint, Public_Category, Cat_Accessible, temp1 As String
        Public_Category = "N"
        Cat_Accessible = ""
        Category_Constraint = ""
        temp1 = ""
        If CAT_ID <> "" Then
            SqlS = "SELECT CK.key_id, C.public_flag " &
            "FROM DMS.dbo.Category_Keywords CK " &
            "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=CK.cat_id " &
            "WHERE CK.cat_id=" & CAT_ID
            If Debug = "Y" Then mydebuglog.Debug("  Get category information : " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        temp1 = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        Public_Category = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        If Debug = "Y" Then
                            mydebuglog.Debug("  > Public_Category: " & Public_Category)
                            mydebuglog.Debug("  > Category with access: " & temp1)
                            mydebuglog.Debug("  > TRAINER_FLG: " & TRAINER_FLG)
                            mydebuglog.Debug("  > MT_FLG: " & MT_FLG)
                            mydebuglog.Debug("  > PART_FLG: " & PART_FLG)
                            mydebuglog.Debug("  > TRAINING_FLG: " & TRAINING_FLG)
                            mydebuglog.Debug("  > TRAINER_ACC_FLG: " & TRAINER_ACC_FLG & vbCrLf)
                        End If
                        Select Case temp1
                            Case "3"
                                If TRAINER_FLG <> "Y" Then
                                    TotalDocs = 0
                                    GoTo RetrieveRecords
                                End If
                            Case "5"
                                If MT_FLG <> "Y" Then
                                    TotalDocs = 0
                                    GoTo RetrieveRecords
                                End If
                            Case "7"
                                If PART_FLG <> "Y" Then
                                    TotalDocs = 0
                                    GoTo RetrieveRecords
                                End If
                            Case "8"
                                If TRAINING_FLG <> "Y" Then
                                    TotalDocs = 0
                                    GoTo RetrieveRecords
                                End If
                            Case "10"
                                If TRAINER_ACC_FLG <> "Y" And TRAINER_FLG <> "Y" Then
                                    TotalDocs = 0
                                    GoTo RetrieveRecords
                                End If
                            Case "12"
                                If SITE_ONLY <> "Y" Then
                                    TotalDocs = 0
                                    GoTo RetrieveRecords
                                End If
                            Case "15"
                                If SYSADMIN_FLG <> "Y" Then
                                    TotalDocs = 0
                                    GoTo RetrieveRecords
                                End If
                            Case "16"
                                If EMP_ID = "" Then
                                    TotalDocs = 0
                                    GoTo RetrieveRecords
                                End If
                        End Select

                    End While
                Else
                    ErrMsg = ErrMsg & "Error getting user information." & vbCrLf
                End If
                dr.Close()
            Catch ex As Exception
                ErrMsg = ErrMsg & "Error getting user information: " & ex.Message & vbCrLf
            End Try
        End If

        ' Build a category keyword constraint for use in association queries
        Category_Constraint = "CK.key_id IN ("
        If TRAINER_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "3,"
            Cat_Accessible = Cat_Accessible & "#3"
        End If
        If MT_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "5,"
            Cat_Accessible = Cat_Accessible & "#5"
        End If
        If PART_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "7,"
            Cat_Accessible = Cat_Accessible & "#7"
        End If
        If TRAINING_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "8,"
            Cat_Accessible = Cat_Accessible & "#8"
        End If
        If TRAINER_ACC_FLG = "Y" Or TRAINER_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "10,"
            Cat_Accessible = Cat_Accessible & "#10"
        End If
        If SITE_ONLY = "Y" Then
            Category_Constraint = Category_Constraint & "12,"
            Cat_Accessible = Cat_Accessible & "#12"
        End If
        Category_Constraint = Category_Constraint & "13,"
        Cat_Accessible = Cat_Accessible & "#13"
        If SYSADMIN_FLG = "Y" Then
            Category_Constraint = Category_Constraint & "15,"
            Cat_Accessible = Cat_Accessible & "#15"
        End If
        If EMP_ID <> "" Then
            Category_Constraint = Category_Constraint & "16,"
            Cat_Accessible = Cat_Accessible & "#16"
        End If
        Category_Constraint = Category_Constraint & "14) "
        Cat_Accessible = Cat_Accessible & "#14"
        If Debug = "Y" Then mydebuglog.Debug(" Computed Category_Constraint: " & Category_Constraint & " - " & Cat_Accessible & vbCrLf)

        ' ================================================
        '  GENERATE AND EXECUTE QUERY
        ' Reset "Where" clause fields	
        temp1 = ""
        Dim temp2 As String = ""
        temp3 = ""
        Dim temp4 As String = ""
        Dim temp5 As String = ""
        Dim temp6 As String = ""
        Dim temp7 As String = ""
        Dim temp8 As String = ""
        Dim temp9 As String = ""

        ' -----
        ' Set selection clause
        temp7 = "SELECT TOP 10000 D.row_id, D.name, D.description, DT.name as type, " &
        "C.name as category, C.row_id as cat_id, DC.pr_flag as cat_pr_flag, " &
        "D.created, D.last_upd, UA.type_id, DU.access_type, CK.key_id, D.dfilename"

        ' Set source clause
        SqlS = "FROM DMS.dbo.Documents D " &
        "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " &
        "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id " &
        "LEFT OUTER JOIN DMS.dbo.Category_Keywords CK ON CK.cat_id=DC.cat_id " &
        "LEFT OUTER JOIN DMS.dbo.Document_Types DT on DT.row_id=D.data_type_id "
        If DLT_FLG <> "Y" Then
            SqlS = SqlS & "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id " &
            "LEFT OUTER JOIN DMS.dbo.User_Group_Access UA ON UA.row_id=DU.user_access_id "
        Else
            SqlS = SqlS & "LEFT OUTER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id " &
            "LEFT OUTER JOIN DMS.dbo.User_Group_Access UA ON UA.row_id=DU.user_access_id "
        End If
        SqlS = SqlS & "WHERE D.row_id IN ("

        GroupBy = "GROUP BY D.row_id, D.name, D.description, DT.name, C.name, C.row_id, " &
        "DC.pr_flag, D.created, D.last_upd, UA.type_id, DU.access_type, CK.key_id, D.dfilename"

        ' << added from GetAssocDocs
        If SearchScope = "A" Then SqlS = SqlS & "("
        'If SearchScope = "A" Or SearchScope = "M" Then
        'SqlS = SqlS & "SELECT D.row_id " &
        '   "FROM DMS.dbo.Documents D " &
        '    "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " &
        '    "LEFT OUTER JOIN DMS.dbo.Category_Keywords CK ON CK.cat_id=DC.cat_id AND (" & Category_Constraint & ") " &
        '    "WHERE DC.pr_flag='Y' GROUP BY D.row_id " &
        '    "INTERSECT "
        'End If
        ' -----
        ' Pass 1 - Add category access constraint
        If Category_Constraint <> "" Then
            If SearchScope = "P" Or Public_Category = "Y" Then
                SqlS = SqlS & "SELECT D.row_id " &
                "FROM DMS.dbo.Documents D " &
                "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " &
                "LEFT OUTER JOIN DMS.dbo.Category_Keywords CK ON CK.cat_id=DC.cat_id AND (" & Category_Constraint & ") " &
                "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id " &
                "WHERE DC.pr_flag='Y' AND CK.key_id IS NOT NULL GROUP BY D.row_id " &
                "INTERSECT "
            Else
                If CAT_ID = "" Then
                    SqlS = SqlS & "SELECT D.row_id " &
                    "FROM DMS.dbo.Documents D " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " &
                    "LEFT OUTER JOIN DMS.dbo.Category_Keywords CK ON CK.cat_id=DC.cat_id AND (" & Category_Constraint & ") " &
                    "WHERE DC.pr_flag='Y' GROUP BY D.row_id " &
                    "INTERSECT "
                End If
            End If
        End If

        ' -----
        ' Pass 1 - Add basic security filter
        temp1 = " "
        If OWN_FLG = "Y" Then
            temp1 = temp1 & "AND D.created_by=" & DMS_USER_ID
        End If

        ' Deleted flag
        If DLT_FLG = "Y" Then
            temp1 = temp1 & "AND D.deleted IS NOT NULL "
        Else
            temp1 = temp1 & "AND D.deleted IS NULL "
        End If

        ' Domain access
        If DMS_FLG <> "" And DLT_FLG <> "Y" And DOMAIN <> "" Then       ' Flag to indicate domain-wide search
            If DMS_FLG = "Y" And DMS_DOMAIN_ID <> "" And DOMAIN <> "TIPS" Then
                temp2 = "DU.user_access_id=" & DMS_DOMAIN_ID & " or "
            End If
            If DOMAIN = "TIPS" Then
                temp2 = temp2 & "DU.user_access_id IN (11,112,828) or "
            Else
                temp2 = temp2 & "DU.user_access_id=11 or "
            End If
        End If

        ' Subscription access
        If SUB_FLG <> "" And DLT_FLG <> "Y" And (DMS_ASUB_ID <> "" Or DMS_SUB_ID <> "") Then            ' Flag to indicate sub-wide search
            If SUB_FLG = "Y" Then
                If ALT_SUB_ID <> "" And DMS_ASUB_ID <> "" Then
                    temp3 = temp3 & "DU.user_access_id=" & DMS_ASUB_ID & " or "
                Else
                    If DMS_SUB_ID <> "" Then
                        temp3 = temp3 & "DU.user_access_id=" & DMS_SUB_ID & " or "
                    End If
                End If
            End If
        End If

        ' User access
        If (USR_FLG <> "" And DLT_FLG <> "Y") Or EDIT_FLG = "Y" Then            ' Flag to indicate user specific search
            If DMS_USER_ID <> "" And PUBLIC_FLG <> "Y" Then
                If (USR_FLG = "Y" Or EDIT_FLG = "Y") And DMS_USER_AID <> "" Then
                    temp4 = temp4 & "DU.user_access_id=" & DMS_USER_AID & " "
                End If
            End If
        End If

        ' Compile rights query
        temp5 = Trim(temp2 & temp3 & temp4)
        If Right(temp5, 2) = "or" Then temp5 = Left(temp5, Len(temp5) - 2)
        temp5 = Trim(temp5)
        If temp5 <> "" Then
            temp9 = temp1 & " AND (" & temp5 & ") "
        Else
            temp9 = temp1
        End If
        If SDOC_ID <> "" Then
            temp9 = temp9 & " AND D.row_id=" & SDOC_ID & " "
        End If
        SqlS = temp7 & " " & SqlS
        If Debug = "Y" Then
            mydebuglog.Debug(" User access: " & temp5 & vbCrLf)
            mydebuglog.Debug(" User access constraint: " & temp9 & vbCrLf)
        End If

        ' -----
        ' Pass 2a - Generate search criteria Intersect filter for user document searches
        If DOC_LIB_FLG = "Y" Then
            temp1 = ""
            SqlS = SqlS & " " &
            "SELECT D.row_id " &
            "FROM DMS.dbo.Documents D " &
            "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id " &
            "WHERE ("
            If CONTACT_ID <> "" Then temp1 = temp1 & " (DA.association_id='3' AND DA.fkey='" & CONTACT_ID & "' and DA.pr_flag='Y') or "
            If TRAINER_FLG = "Y" And TRAINER_ID <> "" Then temp1 = temp1 & " (DA.association_id='5' AND DA.fkey='" & TRAINER_ID & "' and DA.pr_flag='Y') or "
            If PART_ID <> "" Then temp1 = temp1 & " (DA.association_id='4' AND DA.fkey='" & PART_ID & "' and DA.pr_flag='Y') or "
            If MT_FLG = "Y" And EMP_ID <> "" Then temp1 = temp1 & " (DA.association_id='37' AND DA.fkey='" & EMP_ID & "' and DA.pr_flag='Y') or "
            temp1 = Left(temp1, Len(temp1) - 4)
            SqlS = SqlS & temp1 & ") "

            ' -----
            ' Pass 2b - Generate search criteria Intersect filter for all other searches
        Else
            ' Prepare document criteria selection query
            temp1 = " " &
                    "SELECT D.row_id " &
                    "FROM DMS.dbo.Documents D "
            If CAT_ID <> "" Or InStr(AssocFilter, "C.") > 0 Or InStr(AssocFilter, "DC.") > 0 Then
                temp1 = temp1 & "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id "
            End If
            If InStr(AssocFilter, " C.") > 0 Then
                temp1 = temp1 & "LEFT OUTER JOIN DMS.dbo.Categories C ON C.doc_id=DC.cat_id "
            End If
            If AssocFilter <> "" Then
                temp1 = temp1 & "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id "
                If ASN_ID = "" Or InStr(AssocFilter, " A.") > 0 Then
                    temp1 = temp1 & "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id "
                End If
            End If
            If InStr(AssocFilter, "DU.") > 0 Then
                temp1 = temp1 & "LEFT OUTER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id "
            End If
            temp1 = temp1 & "WHERE "

            ' Category parameter
            If CAT_ID <> "" And CALL_SCREEN = "RWD" Then
                temp8 = temp8 & "DC.cat_id=" & CAT_ID & " AND "
            End If

            ' Other selection parameters
            If DOC_EXT_ID <> "" Then
                temp8 = temp8 & "D.data_type_id='" & DOC_EXT_ID & "' AND "
            End If
            If StartDt <> "" Then
                temp8 = temp8 & "D.created>='" & StartDt & "' AND "
            End If
            If EndDt <> "" Then
                temp8 = temp8 & "D.created<='" & EndDt & "' AND "
            End If
            If FTEXT <> "" Then
                temp8 = temp8 & "PATINDEX('%" & FTEXT & "%',D.description)>0 AND "
            End If
            If DOC_NAME <> "" Then
                temp8 = temp8 & "PATINDEX('%" & UCase(DOC_NAME) & "%',UPPER(D.name))>0 AND "
            End If
            If DESC_TEXT <> "" Then
                temp8 = temp8 & "PATINDEX('%" & UCase(DESC_TEXT) & "%',UPPER(D.description))>0 AND "
            End If

            ' Combine parameters
            If temp8 <> "" Then
                If AssocFilter <> "" Then
                    If ASN_OPT = "N" Then
                        temp8 = temp8 & "(" & Right(AssocFilter, Len(AssocFilter) - 4) & ") "
                    Else
                        temp8 = Left(temp8, Len(temp8) - 4)
                        temp8 = " ( (" & temp8 & ") OR (" & Right(AssocFilter, Len(AssocFilter) - 4) & ") )"
                    End If
                Else
                    temp8 = Left(temp8, Len(temp8) - 4)
                End If
            Else
                If Len(AssocFilter) > 0 Then
                    temp8 = Right(AssocFilter, Len(AssocFilter) - 4) & " "
                End If
            End If
            If Debug = "Y" Then mydebuglog.Debug(" Document selection criteria: " & temp8 & vbCrLf)

            ' Add the criteria if applicable
            If temp8 <> "" Then
                SqlS = SqlS & temp1 & temp8 & " "
            Else
                If InStr(SqlS, "INTERSECT") > 0 Then SqlS = Replace(SqlS, "INTERSECT", "")
            End If
            If InStr(temp8, "PATINDEX") > 0 Then SqlS = SqlS & " GROUP BY D.row_id "

            ' If there is a keyword, add a key constraint
            If Debug = "Y" Then
                mydebuglog.Debug(" KEY_ID: " & KEY_ID)
                mydebuglog.Debug(" KEY_NAME: " & KEY_NAME & vbCrLf)
            End If
            If (KEY_ID <> "" And KEY_NAME <> "") Then
                If temp1 <> "" Then
                    SqlS = SqlS & "INTERSECT (" &
                        "SELECT D.row_id " &
                        "FROM DMS.dbo.Documents D " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Keywords K on K.doc_id=D.row_id " &
                        "WHERE K.key_id=" & KEY_ID & ") "
                Else
                    SqlS = SqlS & "" &
                        "SELECT D.row_id " &
                        "FROM DMS.dbo.Documents D " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Keywords K on K.doc_id=D.row_id " &
                        "WHERE K.key_id=" & KEY_ID & " "
                End If
            End If

            ' If there is a category, add user constraint 
            If (CAT_ID <> "" And Public_Category = "N") Or (temp8 <> "" And ASN_ID <> "3" And InStr(temp8, "C.public_flag='Y'") = 0) Then
                If Debug = "Y" Then
                    mydebuglog.Debug(" SearchScope: " & SearchScope)
                    mydebuglog.Debug(" Public_Category: " & Public_Category & vbCrLf)
                End If
                If SearchScope <> "P" And Public_Category <> "Y" Then
                    temp1 = ""
                    temp2 = ""
                    If DMS_USER_AID <> "" Then
                        ' Select records where there is an association between the user and the document
                        temp2 = temp2 & "INTERSECT (" &
                        "SELECT D.row_id " &
                        "FROM DMS.dbo.Documents D " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id " &
                        "WHERE ("
                        If CONTACT_ID <> "" Then temp1 = temp1 & " (DA.association_id='3' AND DA.fkey='" & CONTACT_ID & "') or "
                        If TRAINER_FLG = "Y" And TRAINER_ID <> "" Then temp1 = temp1 & " (DA.association_id='5' AND DA.fkey='" & TRAINER_ID & "') or "
                        If PART_ID <> "" Then temp1 = temp1 & " (DA.association_id='4' AND DA.fkey='" & PART_ID & "') or "
                        If MT_FLG = "Y" And MT_FLG <> "" Then temp1 = temp1 & " (DA.association_id='37' AND DA.fkey='" & EMP_ID & "') or "
                        temp1 = Left(temp1, Len(temp1) - 4)
                        temp2 = temp2 & Trim(temp1) & ") "

                        ' Also select records where the user is an owner of the document				
                        temp2 = temp2 & "	UNION " &
                        "SELECT D.row_id " &
                        "FROM DMS.dbo.Documents D " &
                        "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id  " &
                        "WHERE DU.user_access_id=" & DMS_USER_AID & " " &
                        ") "
                    Else
                        ' Only records where there is an association between the user and the document
                        temp2 = temp2 & "INTERSECT " &
                        "SELECT D.row_id " &
                        "FROM DMS.dbo.Documents D " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id " &
                        "WHERE ("
                        If CONTACT_ID <> "" Then temp1 = temp1 & " (DA.association_id='3' AND DA.fkey='" & CONTACT_ID & "') or "
                        If TRAINER_FLG = "Y" And TRAINER_ID <> "" Then temp1 = temp1 & " (DA.association_id='5' AND DA.fkey='" & TRAINER_ID & "') or "
                        If PART_ID <> "" Then temp1 = temp1 & " (DA.association_id='4' AND DA.fkey='" & PART_ID & "') or "
                        If MT_FLG = "Y" And EMP_ID <> "" Then temp1 = temp1 & " (DA.association_id='37' AND DA.fkey='" & EMP_ID & "') or "
                        temp1 = Left(temp1, Len(temp1) - 4)
                        temp2 = temp2 & Trim(temp1) & ") "
                    End If
                    If Debug = "Y" Then mydebuglog.Debug("Private Category selection: " & temp2)
                    SqlS = SqlS & temp2
                End If
            End If
        End If

        ' ----
        ' Finalize query
        SqlS = SqlS & ") " & temp9

        ' Generate sort clause
        Select Case SortOrd
            Case "1"            ' Category name
                OrderBy = "ORDER BY C.name " & SortDir & ", D.name, D.row_id "
            Case "2"            ' Association name
                OrderBy = "ORDER BY D.association " & SortDir & ", D.name, D.row_id "
            Case "3"           ' Document name
                OrderBy = "ORDER BY D.name " & SortDir & ", D.row_id"
            Case "4"            ' Document type
                OrderBy = "ORDER BY D.type " & SortDir & ", D.name, D.row_id "
            Case "5"            ' Created
                OrderBy = "ORDER BY D.created " & SortDir & ", D.name, D.row_id "
            Case "6"            ' Last Updated
                OrderBy = "ORDER BY D.last_upd " & SortDir & ", D.name, D.row_id "
            Case "7"            ' Created Only
                OrderBy = "ORDER BY D.created " & SortDir
            Case Else
                OrderBy = "ORDER BY D.name " & SortDir & ", D.row_id "
        End Select

        ' -----
        ' Add group by and order by clauses to finalize
        SqlS = SqlS & GroupBy & " " & OrderBy
        SQuery = SqlS

        ' ============================================
        ' PERFORM QUERY
GetRecords:
        If Refresh Then SqlS = SQuery
        If Debug = "Y" Then mydebuglog.Debug(" Document Query: " & SQuery & vbCrLf)
        Try
            Dim EmpQueryDoc As XmlElement = DMSServices.EmpQuery(EMP_ID, CONTACT_ID, "GetContent", OrderID, SqlS, "Y", Refresh, Debug)
            If EmpQueryDoc.HasAttribute("tablename") Then TableName = EmpQueryDoc.GetAttribute("tablename").ToString
            If EmpQueryDoc.HasAttribute("records") Then TotalDocs = Val(EmpQueryDoc.GetAttribute("records").ToString)
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug("  > Unable to perform search: " & ex.Message & vbCrLf)
            GoTo NoneFound
        End Try
        If Debug = "Y" Then
            mydebuglog.Debug(" TableName: " & TableName)
            mydebuglog.Debug(" TotalDocs: " & TotalDocs.ToString & vbCrLf)
        End If
        If TableName = "" Or TotalDocs = 0 Then GoTo NoneFound

RetrieveRecords:
        ' Query the temp table if applicable to get the document count
        If TotalDocs = 0 Then
            SqlS = "SELECT COUNT(*) FROM " & TableName
            If Debug = "Y" Then mydebuglog.Debug(" Select results count: " & SqlS & vbCrLf)
            Try
                d_cmd.CommandText = SqlS
                Dim result As Object = d_cmd.ExecuteScalar()
                If result <> Nothing Then
                    TotalDocs = Val(result.ToString)
                End If
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug("  > Unable to Select results count: " & ex.Message & vbCrLf)
                GoTo NoneFound
            End Try
        End If
        If TotalDocs = 0 Then GoTo NoneFound

        ' Store query for export
        Dim returnv As Integer
        If Not Refresh Then
            SqlS = "INSERT reports.dbo.CMREP (USER_ID, ORDER_ID, CREATED, LOC_ID, FULL_QUERY, EDIT_FLG, DEL_FLG, ADMIN_ACCESS, SUB_FLG, USR_FLG) " &
                "VALUES ('" & EMP_ID & "','" & OrderID & "',GETDATE(),'S058','" & Replace(SQuery, "'", "''") & "','" & EDIT_FLG & "','" & DEL_FLG & "','" & ADMIN_ACCESS & "','" & SUB_FLG & "','" & USR_FLG & "')"
            If Debug = "Y" Then mydebuglog.Debug("  Store query for export: " & SqlS & vbCrLf)
            Try
                d_cmd.CommandText = SqlS
                returnv = d_cmd.ExecuteNonQuery()
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug("  > Unable to Store query for export: " & ex.Message & vbCrLf)
            End Try
        End If

        ' ================================================
        ' GENERATE QUERYRESULTS

        ' -----
        ' Build relink query
        relink = ""
        If CAT_ID <> "" Then relink = relink & "&CTI=" & EnURL(CAT_ID)
        If CAT_NAME <> "" Then relink = relink & "&CTN=" & EnURL(CAT_NAME)
        If ASN_NAME <> "" Then relink = relink & "&ASN=" & EnURL(ASN_NAME)
        If ASN_ID <> "" Then relink = relink & "&ASI=" & EnURL(ASN_ID)
        If ASN_KEY <> "" Then relink = relink & "&ASK=" & ASN_KEY
        If ASN_OPT <> "" Then relink = relink & "&AOP=" & EnURL(ASN_OPT)
        If DOC_EXT <> "" Then relink = relink & "&DCT=" & DOC_EXT
        If DLT_FLG <> "" Then relink = relink & "&DCR=" & DLT_FLG
        If FTEXT <> "" Then relink = relink & "&FTX=" & EnURL(FTEXT)
        If StartDt <> "" Then relink = relink & "&DTH=" & StartDt
        If EndDt <> "" Then relink = relink & "&EDT=" & EndDt
        If DOC_NAME <> "" Then relink = relink & "&NME=" & EnURL(DOC_NAME)
        If DESC_TEXT <> "" Then relink = relink & "&DSC=" & EnURL(DESC_TEXT)
        If SearchScope <> "" Then relink = relink & "&SS=" & EnURL(SearchScope)
        If PUBLIC_FLG <> "" Then relink = relink & "&PUB=" & PUBLIC_FLG
        If DMS_FLG <> "" Then relink = relink & "&DMU=" & DMS_FLG
        If SUB_FLG <> "" Then relink = relink & "&SBU=" & SUB_FLG
        If USR_FLG <> "" Then relink = relink & "&USR=" & USR_FLG
        If OWN_FLG <> "" Then relink = relink & "&OWN=" & OWN_FLG
        If ALT_SUB_ID <> "" Then relink = relink & "&ASB=" & ALT_SUB_ID
        If DOC_LIB_FLG <> "" Then relink = relink & "&DLF=" & DOC_LIB_FLG
        If SortOrd <> "" Then relink = relink & "&SO=" & SortOrd
        If SortDir <> "" Then relink = relink & "&SD=" & SortDir
        If DOMAIN <> "" Then relink = relink & "&DOM=" & DOMAIN
        If POPUP_FLG <> "" Then relink = relink & "&POP=" & POPUP_FLG
        'If CALL_ID<>"" Then relink = relink & "&CID=" & CALL_ID
        If CALL_SCREEN <> "" Then relink = relink & "&CIS=" & CALL_SCREEN
        If NOT_TYPE <> "" Then relink = relink & "&NTY=" & NOT_TYPE
        If CST_FLG <> "" Then relink = relink & "&CST=" & CST_FLG
        If CAT_FLG <> "" Then relink = relink & "&CFL=" & CAT_FLG
        If ASC_FLG <> "" Then relink = relink & "&AFL=" & ASC_FLG
        If ASSOC_PARAM <> "" Then relink = relink & "&APM=" & ASSOC_PARAM
        If ASSOC_PARAM_TYPE <> "" Then relink = relink & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
        If ASSOC_DEF <> "" Then relink = relink & "&ADF=" & ASSOC_DEF
        If ASSOC_RSTR <> "" Then relink = relink & "&ASR=" & ASSOC_RSTR
        If EMP_ID <> "" Then relink = relink & "&EID=" & EMP_ID
        If Debug = "Y" Then mydebuglog.Debug(" relink: " & relink & vbCrLf)

        ' ================================================
        ' Generate Page Header
        PHeader = "<table width=""1000"" border=""0"" cellpadding=""2"" cellspacing=""1"" bgcolor=""#FFFFFF"">" & vbCrLf
        PHeader = PHeader & "<tr VALIGN=top><td BGCOLOR=""#FFFFFF"" class=""body"">" & vbCrLf

        ' Buttons
        If ADMIN_ACCESS = "Y" Then
            If DLT_FLG <> "Y" And NOT_TYPE <> "DCT" Then
                ButtonList = ButtonList & "<td class=""otherlinks""><a href=""javascript:openNewWindow('" & BasePath & "/ListRecs?OpenAgent&TYP=DMCT&POP=Y&EID=" & EMP_ID & "',800,600)"" class=""buttons2"">Edit Categories</a></td>"
            End If
            If ALT_SUB_ID = "" Then
                temp3 = ""
                If SDOC_ID <> "" Then temp3 = temp3 & "&ID=" & SDOC_ID
                If ASN_NAME <> "" Then temp3 = temp3 & "&ASN=" & EnURL(ASN_NAME)
                If ASN_ID <> "" Then temp3 = temp3 & "&ASI=" & EnURL(ASN_ID)
                If ASN_KEY <> "" Then temp3 = temp3 & "&ASK=" & EnURL(ASN_KEY)
                If CAT_NAME <> "" Then temp3 = temp3 & "&CTN=" & EnURL(CAT_NAME)
                If CAT_ID <> "" Then temp3 = temp3 & "&CTD=" & EnURL(CAT_ID)
                If KEY_ID <> "" Then temp3 = temp3 & "&KEY=" & EnURL(KEY_ID)
                If DOC_EXT <> "" Then temp3 = temp3 & "&DCT=" & DOC_EXT
                If DLT_FLG <> "" Then temp3 = temp3 & "&DCR=" & DLT_FLG
                If ALT_SUB_ID <> "" Then temp3 = temp3 & "&ASB=" & EnURL(ALT_SUB_ID)
                If FTEXT <> "" Then temp3 = temp3 & "&FTX=" & EnURL(FTEXT)
                If StartDt <> "" Then temp3 = temp3 & "&DTH=" & StartDt
                If EndDt <> "" Then temp3 = temp3 & "&EDT=" & EndDt
                If DOC_NAME <> "" Then temp3 = temp3 & "&NME=" & EnURL(DOC_NAME)
                If DESC_TEXT <> "" Then temp3 = temp3 & "&DSC=" & EnURL(DESC_TEXT)
                If DMS_FLG <> "" Then temp3 = temp3 & "&DMU=" & DMS_FLG
                If SUB_FLG <> "" Then temp3 = temp3 & "&SBU=" & SUB_FLG
                If USR_FLG <> "" Then temp3 = temp3 & "&USR=" & USR_FLG
                If OWN_FLG <> "" Then temp3 = temp3 & "&OWN=" & OWN_FLG
                If SortOrd <> "" Then temp3 = temp3 & "&SO=" & EnURL(SortOrd)
                If SortDir <> "" Then temp3 = temp3 & "&SD=" & EnURL(SortDir)
                If DOMAIN <> "" Then temp3 = temp3 & "&DOM=" & DOMAIN
                If POPUP_FLG <> "" Then temp3 = temp3 & "&POP=" & POPUP_FLG
                If OrderID <> "" Then temp3 = temp3 & "&CID=" & OrderID
                If CST_FLG <> "" Then temp3 = temp3 & "&CST=" & CST_FLG
                If CAT_FLG <> "" Then temp3 = temp3 & "&CFL=" & CAT_FLG
                If ASC_FLG <> "" Then temp3 = temp3 & "&AFL=" & ASC_FLG
                If ASSOC_PARAM <> "" Then temp3 = temp3 & "&APM=" & ASSOC_PARAM
                If ASSOC_PARAM_TYPE <> "" Then temp3 = temp3 & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
                If ASSOC_DEF <> "" Then temp3 = temp3 & "&ADF=" & ASSOC_DEF
                If ASSOC_RSTR <> "" Then temp3 = temp3 & "&ASR=" & ASSOC_RSTR
                If EMP_ID <> "" Then temp3 = temp3 & "&EID=" & EMP_ID
                If CALL_ID <> "" Then temp3 = temp3 & "&CID=" & CALL_ID
                If CALL_SCREEN <> "" Then temp3 = temp3 & "&CIS=" & CALL_SCREEN
                If NOT_TYPE <> "" Then temp3 = temp3 & "&NTY=" & NOT_TYPE
                If DLT_FLG = "N" Or DLT_FLG = "" Then
                    temp3 = temp3 & "&DCR=Y&CIS=1"
                    ButtonList = ButtonList & "<td class=""otherlinks""><a href='GetContent.ashx?" & temp3 & "' class=""buttons2"">Restore Deleted</a></td>"
                Else
                    temp3 = temp3 & "&DCR=N&CIS=1"
                    ButtonList = ButtonList & "<td class=""otherlinks""><a href='GetContent.ashx?" & temp3 & "' class=""buttons2"">Review Non-Deleted</a></td>"
                End If
            End If
        End If
        If (EDIT_FLG = "Y" And DLT_FLG <> "Y" And ALT_SUB_ID = "") Or (CST_FLG = "Y" Or ADMIN_ACCESS = "Y") Then
            If CST_FLG = "Y" Or ADMIN_ACCESS = "Y" Then
                temp3 = "&AID=&CTD=9&CIS=" & NOT_TYPE & "&CID=" & OrderID & "&POP=N&EID=" & EMP_ID
                If ASN_NAME <> "" Then temp3 = temp3 & "&ASN=" & EnURL(ASN_NAME)
                If ASN_ID <> "" Then temp3 = temp3 & "&ASI=" & EnURL(ASN_ID)
                If ASN_KEY <> "" Then temp3 = temp3 & "&ASK=" & EnURL(ASN_KEY)
                If Not (ASSOC_RSTR = "Y" And OWNER_FLG <> "Y") Then
                    ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/OpenContent?OpenAgent" & temp3 & "' class=""buttons2"">Add a Document</a></td>"
                End If
            Else
                temp3 = "&CID=" & OrderID
                If CAT_NAME <> "" Then temp3 = temp3 & "&CTN=" & EnURL(CAT_NAME)
                If CAT_ID <> "" Then temp3 = temp3 & "&CTD=" & EnURL(CAT_ID)
                If KEY_ID <> "" Then temp3 = temp3 & "&KEY=" & EnURL(KEY_ID)
                If ASN_NAME <> "" Then temp3 = temp3 & "&ASN=" & EnURL(ASN_NAME)
                If ASN_ID <> "" Then temp3 = temp3 & "&AID=" & EnURL(ASN_ID)
                If ASN_KEY <> "" Then temp3 = temp3 & "&ASK=" & EnURL(ASN_KEY)
                If ALT_SUB_ID <> "" Then temp3 = temp3 & "&ASB=" & EnURL(ALT_SUB_ID)
                If ASSOC_PARAM <> "" Then temp3 = temp3 & "&APM=" & EnURL(ASSOC_PARAM)
                If ASSOC_PARAM_TYPE <> "" Then temp3 = temp3 & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
                If ASSOC_DEF <> "" Then temp3 = temp3 & "&ADF=" & EnURL(ASSOC_DEF)
                If ASSOC_RSTR <> "" Then temp3 = temp3 & "&ASR=" & ASSOC_RSTR
                If DOMAIN <> "" Then temp3 = temp3 & "&DOM=" & DOMAIN
                If CST_FLG <> "" Then temp3 = temp3 & "&CST=" & CST_FLG
                If POPUP_FLG <> "" Then temp3 = temp3 & "&POP=" & POPUP_FLG
                If PUBLIC_FLG <> "" Then temp3 = temp3 & "&PUB=" & PUBLIC_FLG
                If EMP_ID <> "" Then temp3 = temp3 & "&EID=" & EMP_ID
                If CALL_SCREEN <> "" Then temp3 = temp3 & "&CIS=" & CALL_SCREEN
                If NOT_TYPE <> "" Then temp3 = temp3 & "&NTY=" & NOT_TYPE
                If CST_FLG <> "Y" Then
                    If Not (ASSOC_RSTR = "Y" And OWNER_FLG <> "Y") Then
                        ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/OpenContent?OpenAgent&NEW=Y&ID=" & temp3 & "' class=""buttons2"">Add a Document</a></td>"
                    End If
                End If
            End If
        End If
        If ASN_NAME = "Order" And ASN_KEY <> "" Then
            ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/SendReceipt?OpenAgent&ID=" & ASN_KEY & "&EID=" & EMP_ID & "&CID=" & OrderID & "&CIS=" & CALL_SCREEN & "' class=""buttons2"">Send Receipt</a></td>"
        End If
        If NOT_TYPE = "TRC" Then
            ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/UpdRecs?OpenAgent&CID=" & OrderID & "&TYP=GTC&STP=Card&RID=" & ASN_KEY & "' class=""buttons2"">Generate Card</a></td>"
            ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/UpdRecs?OpenAgent&CID=" & OrderID & "&TYP=GTCR&STP=Card&RID=" & ASN_KEY & "' class=""buttons2"">Generate Cert</a></td>"
        End If
        'If OUTPUT_FLG = "Y" Then
        ' ButtonList = ButtonList & "<td class=""otherlinks""><a href=""javascript:openNewWindow('" & BasePath & "/ExportScreen?OpenAgent&ID=" & OrderID & "&POP=Y&DB=DMS&EID=" & EMP_ID & "&TYP=DOC',800,600)"" class=""buttons2"">Export</a></td>"
        'End If
        ButtonList = ButtonList & "<td class=""otherlinks""><a href='GetContent.ashx?RFR=" & OrderID & "' class=""buttons2"">Refresh</a></td>"

        ' Previous screen
        If CALL_SCREEN = "RWD" Then
            ButtonList = ButtonList & "<td class=""otherlinks""><a href='javascript:history.go(-1)' class=""buttons2"">Previous</a></td>"
        End If
        If Debug = "Y" Then mydebuglog.Debug(" ButtonList: " & ButtonList & vbCrLf)
        If ButtonList <> "" Then
            PHeader = PHeader & "<table border=""0"" cellpadding=""0"" cellspacing=""1""><tr valign=""top"">" & ButtonList & "</tr></table>"
        End If
        PHeader = PHeader & "</td></tr>"

        ' Title
        PHeader = PHeader & "<tr><td width=""1152"" valign=""top"" bgcolor=""#F5ECDA"" class=""bodyKO"">" & vbCrLf
        PHeader = PHeader & "<div align=""left""><img src=""/images/spacer2.gif"" width=""3"" height=""8"">" & vbCrLf
        If CAT_NAME <> "" Then
            If UCase(Right(CAT_NAME, 1)) = "S" Then
                PTITLE = Trim(StrConv(DeURL(CAT_NAME), 3))
            Else
                PTITLE = Trim(StrConv(DeURL(CAT_NAME), 3) & " items ")
            End If
        Else
            PTITLE = "Documents "
        End If
        If DLT_FLG = "Y" Then PTITLE = "Deleted " & PTITLE
        If PRIMARY_ASSOC <> "" Then PTITLE = PTITLE & " found related to <i>" & PRIMARY_ASSOC & "</i>"
        If PRIMARY_ASSOC = "" Then PTITLE = PTITLE & " found "
        If Debug = "Y" Then mydebuglog.Debug(" PTITLE: " & PTITLE & vbCrLf)
        PHeader = PHeader & "<span class=""BigHeader""><strong>"
        If DOC_LIB_FLG = "Y" Then PHeader = PHeader & "Your "
        PHeader = PHeader & PTITLE & "</strong></span></div></td>" & vbCrLf
        PHeader = PHeader & "</tr></table>" & vbCrLf
        If Debug = "Y" Then mydebuglog.Debug(" PHeader: " & PHeader & vbCrLf)

        ' Generate Page
        outstring = outstring & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & vbCrLf
        outstring = outstring & "<html>" & vbCrLf
        outstring = outstring & "<head>" & vbCrLf
        outstring = outstring & "<title>DMSA</title>" & vbCrLf
        outstring = outstring & "<meta http-equiv=""Pragma"" content=""no-cache"">" & vbCrLf
        outstring = outstring & "<meta http-equiv=""Cache-Control"" content=""no-cache, no-store, must-revalidate"">" & vbCrLf
        outstring = outstring & "<meta http-equiv=""Expires"" content=""-1"">" & vbCrLf
        outstring = outstring & "<link rel=""stylesheet"" href=""//datafluxapp/css/facebox.css"" media=""screen"" type=""text/css""/>" & vbCrLf
        outstring = outstring & "<style type=""text/css"" title=""currentStyle""> " & vbCrLf
        outstring = outstring & "@import ""//datafluxapp/themes/smoothness/jquery-ui-1.7.2.custom.css""; " & vbCrLf
        outstring = outstring & "</style> " & vbCrLf
        outstring = outstring & "<link rel=""stylesheet"" href=""//datafluxapp/css/TableTools.css"" media=""screen"" type=""text/css""/>" & vbCrLf
        outstring = outstring & "<script type=""text/javascript"" src=""//datafluxapp/js/jquery-1.6.2.min.js""></script> " & vbCrLf
        outstring = outstring & "<script type=""text/javascript"" src=""//datafluxapp/js/jquery.dataTables.min.js""></script> " & vbCrLf
        outstring = outstring & "<script type=""text/javascript"" src=""//datafluxapp/js/pipelining.js""></script>" & vbCrLf
        outstring = outstring & "<script type=""text/javascript"" src=""//datafluxapp/js/datatables.delay.js""></script>" & vbCrLf
        outstring = outstring & "<link rel=""stylesheet"" href=""//datafluxapp/cs2/tips.css"" media=""screen"" type=""text/css""/>" & vbCrLf
        outstring = outstring & "<style>#pgbody { top: 10px; }</style>" & vbCrLf
        outstring = outstring & "<style type=""text/css""> " & vbCrLf
        outstring = outstring & "#ClassFrame2 {position:absolute; " & vbCrLf
        outstring = outstring & "        left: 4px; " & vbCrLf
        outstring = outstring & "        top: 24px; " & vbCrLf
        outstring = outstring & "        width: 600px; " & vbCrLf
        outstring = outstring & "        height: 800px; " & vbCrLf
        outstring = outstring & "        z-index: 0; " & vbCrLf
        outstring = outstring & "        resize: both; " & vbCrLf
        outstring = outstring & "       } " & vbCrLf
        outstring = outstring & "#ClassFrame3 {position:absolute; " & vbCrLf
        outstring = outstring & "        left: 4px; " & vbCrLf
        outstring = outstring & "        top: 2px; " & vbCrLf
        outstring = outstring & "        width: 200px; " & vbCrLf
        outstring = outstring & "        height: 834px; " & vbCrLf
        outstring = outstring & "        z-index: 0; " & vbCrLf
        outstring = outstring & "        resize: both; " & vbCrLf
        outstring = outstring & "       } " & vbCrLf
        outstring = outstring & "#wrap {" & vbCrLf
        outstring = outstring & "        width: 750px;" & vbCrLf
        outstring = outstring & "        height: 1500px;" & vbCrLf
        outstring = outstring & "        padding: 0;" & vbCrLf
        outstring = outstring & "        overflow: hidden;" & vbCrLf
        outstring = outstring & "}" & vbCrLf
        outstring = outstring & "#class {" & vbCrLf
        outstring = outstring & "        width: 1000px;" & vbCrLf
        outstring = outstring & "        height: 1350px;" & vbCrLf
        outstring = outstring & "        border: 0px;" & vbCrLf
        outstring = outstring & "        zoom: 1.00;" & vbCrLf
        outstring = outstring & "        background-color: white;" & vbCrLf
        outstring = outstring & "        -moz-transform: scale(1.00);" & vbCrLf
        outstring = outstring & "        -moz-transform-origin: 0 0;" & vbCrLf
        outstring = outstring & "        -o-transform: scale(1.00);" & vbCrLf
        outstring = outstring & "        -o-transform-origin: 0 0;" & vbCrLf
        outstring = outstring & "        -webkit-transform: scale(1.00);" & vbCrLf
        outstring = outstring & "        -webkit-transform-origin: 0 0;" & vbCrLf
        outstring = outstring & "a.buttons1:link, a.buttons1:visited {" & vbCrLf
        outstring = outstring & "    background-color: #F5ECDA;" & vbCrLf
        outstring = outstring & "    border: dotted 1px;" & vbCrLf
        outstring = outstring & "    border-color: #BA3A13 #BA3A13 #BA3A13 #BA3A13;" & vbCrLf
        outstring = outstring & "    color: #BA3A13;" & vbCrLf
        outstring = outstring & "    font-family: arial,helvetica,sans-serif;" & vbCrLf
        outstring = outstring & "    font-size: 7pt;" & vbCrLf
        outstring = outstring & "    font-weight: bold;" & vbCrLf
        outstring = outstring & "    letter-spacing: normal;" & vbCrLf
        outstring = outstring & "    padding: 1px;" & vbCrLf
        outstring = outstring & "    text-decoration: none;" & vbCrLf
        outstring = outstring & "    width: 60px;" & vbCrLf
        outstring = outstring & "    display: block;" & vbCrLf
        outstring = outstring & "    overflow: hidden;" & vbCrLf
        outstring = outstring & "    line-height: 1.5;" & vbCrLf
        outstring = outstring & "    margin-right: 2px;" & vbCrLf
        outstring = outstring & "    margin-bottom: 2px;" & vbCrLf
        outstring = outstring & "    text-align: center;}" & vbCrLf
        outstring = outstring & "a.buttons1:hover {" & vbCrLf
        outstring = outstring & "    background-color: #283747; color: white;}" & vbCrLf
        outstring = outstring & "a.buttons1:active, a.buttons1:focus {" & vbCrLf
        outstring = outstring & "    border: dotted 1px;" & vbCrLf
        outstring = outstring & "    border-color: #BA3A13 #BA3A13 #BA3A13 #BA3A13;" & vbCrLf
        outstring = outstring & "    letter-spacing: normal;" & vbCrLf
        outstring = outstring & "    padding: 1px;}" & vbCrLf
        outstring = outstring & "}" & vbCrLf
        outstring = outstring & "@media screen and (-webkit-min-device-pixel-ratio:0) { #class { zoom: 1; } }" & vbCrLf
        outstring = outstring & "</style> " & vbCrLf
        outstring = outstring & "<script type=""text/javascript""> " & vbCrLf
        'outstring = outstring & "function changewidth() { var x = document.getElementById('Class'); x.style.width = window.innerWidth; }" & vbCrLf
        outstring = outstring & "function openNewWindow(fileName,theWidth,theHeight) { " & vbCrLf
        outstring = outstring & "   var win = window.open("""", ""Document"", ""width=600,height=800,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1"");" & vbCrLf
        outstring = outstring & "   if (win) { win.location=fileName; }" & vbCrLf
        outstring = outstring & "   if (win) { win.focus(); }" & vbCrLf
        outstring = outstring & "} " & vbCrLf
        outstring = outstring & "</script>" & vbCrLf
        outstring = outstring & "</head>" & vbCrLf
        outstring = outstring & "<body text=""#000000"" bgcolor=""#FFFFFF"">" & vbCrLf
        outstring = outstring & "<div id=""pgbody"">" & vbCrLf
        outstring = outstring & "<center>" & vbCrLf
        outstring = outstring & PHeader & vbCrLf
        outstring = outstring & "<script type=""text/javascript"" charset=""utf-8""> " & vbCrLf
        outstring = outstring & "$(document).ready(function() { " & vbCrLf
        outstring = outstring & "	oTable = $('#ListResults').dataTable({ " & vbCrLf
        outstring = outstring & "		""bJQueryUI"": true, " & vbCrLf
        outstring = outstring & "		""bProcessing"": true, " & vbCrLf
        outstring = outstring & "		""bServerSide"": true, " & vbCrLf
        outstring = outstring & "		""sAjaxSource"": ""GetContentBE.ashx?RG=" & EMP_ID & "&SESS=" & EMP_ID & "&CID=" & OrderID & relink & "&BP=https://w4.certegrity.com/dmsa.nsf"", " & vbCrLf
        outstring = outstring & "		""fnServerData"": fnDataTablesPipeline, " & vbCrLf
        outstring = outstring & "		""iDisplayLength"": 10, " & vbCrLf
        outstring = outstring & "		""sPaginationType"": ""full_numbers"", " & vbCrLf
        outstring = outstring & "		""aaSorting"": [[ 1, ""DESC"" ]], " & vbCrLf
        outstring = outstring & "		""sDom"": '<""H""lfr>t<""F""ip>' , " & vbCrLf
        outstring = outstring & "		""bStateSave"" : true, " & vbCrLf
        outstring = outstring & "		""oLanguage"": { " & vbCrLf
        outstring = outstring & "			""sInfo"": ""Total of _TOTAL_  documents to show (_START_ to _END_)"", " & vbCrLf
        outstring = outstring & "			""sProcessing"": ""<img src='/images/ajax-loader.gif' border='1'>"" " & vbCrLf
        outstring = outstring & "		}, " & vbCrLf
        outstring = outstring & "		""aoColumnDefs"": [  " & vbCrLf
        outstring = outstring & "			{ ""sType"": ""html"",  ""aTargets"": [ 0 ] } " & vbCrLf
        outstring = outstring & "		] " & vbCrLf
        outstring = outstring & "	}).fnSetFilteringDelay(1000); " & vbCrLf
        outstring = outstring & "}); " & vbCrLf
        outstring = outstring & "</script> " & vbCrLf
        outstring = outstring & "<div id=""queryResults""><div id=""container""><div class=""demo_jui""> " & vbCrLf
        outstring = outstring & "<table width=""1000"" cellpadding=""0"" cellspacing=""0"" border=""0"" class=""display compact"" id=""ListResults""> " & vbCrLf
        outstring = outstring & "<thead> " & vbCrLf
        outstring = outstring & "<tr> " & vbCrLf
        outstring = outstring & "<th>Category</th>" & vbCrLf
        outstring = outstring & "<th>Name</th>" & vbCrLf
        outstring = outstring & "<th>Date</th>" & vbCrLf
        outstring = outstring & "<th>Id</th>" & vbCrLf
        outstring = outstring & "<th>Description</th>" & vbCrLf
        If (EDIT_FLG = "Y" Or DEL_FLG = "Y" Or ADMIN_ACCESS = "Y") Then
            outstring = outstring & "<th>Action</th>" & vbCrLf
        End If
        outstring = outstring & "</tr></thead>" & vbCrLf
        outstring = outstring & "<tbody> " & vbCrLf
        outstring = outstring & "</tbody></table><br/><br/><br/>" & vbCrLf
        outstring = outstring & "</div></div></div>" & vbCrLf
        outstring = outstring & "<div id=""ClassFrame3"" style=""background-color#FFFFFF; display:none;"">" & vbCrLf
        outstring = outstring & "<table border=""0"" cellpadding=""0"" cellspacing=""1"" style=""margin-left:2px; margin; width: 1000px; background-color: white;""><tr valign=""top"">" & vbCrLf
        outstring = outstring & "<td class=""otherlinks"" width=""7%""><a href=""#"" onclick=""document.getElementById('Class').style.zoom = '50%';"" class=""buttons1"">50%</a></td>" & vbCrLf
        outstring = outstring & "<td class=""otherlinks"" width=""7%""><a href=""#"" onclick=""document.getElementById('Class').style.zoom = '100%';"" class=""buttons1"">100%</a></td>" & vbCrLf
        outstring = outstring & "<td class=""otherlinks"" width=""7%""><a href=""#"" onclick=""document.getElementById('Class').style.zoom = '150%';"" class=""buttons1"">150%</a></td>" & vbCrLf
        outstring = outstring & "<td class=""otherlinks"" width=""7%""><a href=""#"" onclick=""document.getElementById('ClassFrame3').style.display = 'none';"" class=""buttons1"">Close</a></td>" & vbCrLf
        outstring = outstring & "<td class=""otherlinks"" width=""7%""><a href=""#"" onclick=""document.getElementById('Class').contentWindow.focus();document.getElementById('Class').contentWindow.print();"" class=""buttons1"">Print</a></td>" & vbCrLf
        outstring = outstring & "<td class=""otherlinks"" width=""65%""></td>"
        outstring = outstring & "</tr></table>" & vbCrLf
        outstring = outstring & "<div id=""ClassFrame2"" style=""margin-top 20px;"">" & vbCrLf
        outstring = outstring & "<iframe name=""ClassFrame"" id=""Class"" style=""background-color#FFFFFF; border: groove;"" height=""600"" width=""800"" marginwidth=""0"" marginheight=""0"" scrolling=""no"" src="""" allowtransparency=""True"">One moment please...</iframe>" & vbCrLf
        outstring = outstring & "</center>" & vbCrLf
        outstring = outstring & "<div></div>" & vbCrLf
        outstring = outstring & "</form>" & vbCrLf
        outstring = outstring & "</body>" & vbCrLf
        outstring = outstring & "</html>" & vbCrLf
        GoTo CloseOut

DBError:
        ErrLvl = "Error"
        Select Case LANG_CD
            Case "ESN"
                ErrMsg = ErrMsg & "<br/>El sistema puede no estar disponible ahora. Por favor, int&eacute;ntelo de nuevo m&aacute;s tarde"
            Case Else
                ErrMsg = ErrMsg & "<br/>The system may be unavailable now.  Please try again later"
        End Select
        GoTo CloseOut

AccessError:
        ErrLvl = "Error"
        Select Case LANG_CD
            Case "ESN"
                ErrMsg = ErrMsg & "<br/>No tienes acceso a esta funci&oacute;n"
            Case Else
                ErrMsg = ErrMsg & "<br/>You do not have access to this function"
        End Select
        GoTo CloseOut

NoneFound:
        Select Case LANG_CD
            Case "ESN"
                ErrMsg = ErrMsg & "<br/>No se pueden recuperar los documentos. No se encontr&oacute; ninguno, el criterio de b&uacute;squeda era incorrecto o no puede acceder a esta informaci&oacute;n"
            Case Else
                ErrMsg = ErrMsg & "<br/>Unable to retrieve documents.  Either none were found, the search criteria was incorrect, Or this information Is Not accessible to you"
        End Select
        GoTo ErrorDisplay

CatNotFound:
        Select Case LANG_CD
            Case "ESN"
                ErrMsg = ErrMsg & "<br/>No se encontr&oacute; la categor&iacute;a especificada. Por favor, int&eacute;ntelo de nuevo m&aacute;s tarde"
            Case Else
                ErrMsg = ErrMsg & "<br/>The specified category was Not found.  Please try again later"
        End Select

        ' Generate an error screen
ErrorDisplay:
        PHeader = "<table width=""1000"" border=""0"" cellpadding=""2"" cellspacing=""1"" bgcolor=""#FFFFFF"">" & vbCrLf
        PHeader = PHeader & "<tr VALIGN=top><td BGCOLOR=""#FFFFFF"" class=""body"">" & vbCrLf

        ' Buttons
        If ADMIN_ACCESS = "Y" Then
            If DLT_FLG <> "Y" And NOT_TYPE <> "DCT" Then
                ButtonList = ButtonList & "<td class=""otherlinks""><a href=""javascript:openNewWindow('" & BasePath & "/ListRecs?OpenAgent&TYP=DMCT&POP=Y&EID=" & EMP_ID & "',800,600)"" class=""buttons2"">Edit Categories</a></td>"
            End If
            If ALT_SUB_ID = "" Then
                temp3 = ""
                If SDOC_ID <> "" Then temp3 = temp3 & "&ID=" & SDOC_ID
                If ASN_NAME <> "" Then temp3 = temp3 & "&ASN=" & EnURL(ASN_NAME)
                If ASN_ID <> "" Then temp3 = temp3 & "&ASI=" & EnURL(ASN_ID)
                If ASN_KEY <> "" Then temp3 = temp3 & "&ASK=" & EnURL(ASN_KEY)
                If CAT_NAME <> "" Then temp3 = temp3 & "&CTN=" & EnURL(CAT_NAME)
                If CAT_ID <> "" Then temp3 = temp3 & "&CTD=" & EnURL(CAT_ID)
                If KEY_ID <> "" Then temp3 = temp3 & "&KEY=" & EnURL(KEY_ID)
                If DOC_EXT <> "" Then temp3 = temp3 & "&DCT=" & DOC_EXT
                If DLT_FLG <> "" Then temp3 = temp3 & "&DCR=" & DLT_FLG
                If ALT_SUB_ID <> "" Then temp3 = temp3 & "&ASB=" & EnURL(ALT_SUB_ID)
                If FTEXT <> "" Then temp3 = temp3 & "&FTX=" & EnURL(FTEXT)
                If StartDt <> "" Then temp3 = temp3 & "&DTH=" & StartDt
                If EndDt <> "" Then temp3 = temp3 & "&EDT=" & EndDt
                If DOC_NAME <> "" Then temp3 = temp3 & "&NME=" & EnURL(DOC_NAME)
                If DESC_TEXT <> "" Then temp3 = temp3 & "&DSC=" & EnURL(DESC_TEXT)
                If DMS_FLG <> "" Then temp3 = temp3 & "&DMU=" & DMS_FLG
                If SUB_FLG <> "" Then temp3 = temp3 & "&SBU=" & SUB_FLG
                If USR_FLG <> "" Then temp3 = temp3 & "&USR=" & USR_FLG
                If OWN_FLG <> "" Then temp3 = temp3 & "&OWN=" & OWN_FLG
                If SortOrd <> "" Then temp3 = temp3 & "&SO=" & EnURL(SortOrd)
                If SortDir <> "" Then temp3 = temp3 & "&SD=" & EnURL(SortDir)
                If DOMAIN <> "" Then temp3 = temp3 & "&DOM=" & DOMAIN
                If POPUP_FLG <> "" Then temp3 = temp3 & "&POP=" & POPUP_FLG
                If OrderID <> "" Then temp3 = temp3 & "&CID=" & OrderID
                If CST_FLG <> "" Then temp3 = temp3 & "&CST=" & CST_FLG
                If CAT_FLG <> "" Then temp3 = temp3 & "&CFL=" & CAT_FLG
                If ASC_FLG <> "" Then temp3 = temp3 & "&AFL=" & ASC_FLG
                If ASSOC_PARAM <> "" Then temp3 = temp3 & "&APM=" & ASSOC_PARAM
                If ASSOC_PARAM_TYPE <> "" Then temp3 = temp3 & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
                If ASSOC_DEF <> "" Then temp3 = temp3 & "&ADF=" & ASSOC_DEF
                If ASSOC_RSTR <> "" Then temp3 = temp3 & "&ASR=" & ASSOC_RSTR
                If EMP_ID <> "" Then temp3 = temp3 & "&EID=" & EMP_ID
                If CALL_ID <> "" Then temp3 = temp3 & "&CID=" & CALL_ID
                If CALL_SCREEN <> "" Then temp3 = temp3 & "&CIS=" & CALL_SCREEN
                If NOT_TYPE <> "" Then temp3 = temp3 & "&NTY=" & NOT_TYPE
                If DLT_FLG = "N" Or DLT_FLG = "" Then
                    temp3 = temp3 & "&DCR=Y&CIS=1"
                    ButtonList = ButtonList & "<td class=""otherlinks""><a href='GetContent.ashx?" & temp3 & "' class=""buttons2"">Restore Deleted</a></td>"
                Else
                    temp3 = temp3 & "&DCR=N&CIS=1"
                    ButtonList = ButtonList & "<td class=""otherlinks""><a href='GetContent.ashx?" & temp3 & "' class=""buttons2"">Review Non-Deleted</a></td>"
                End If
            End If
        End If
        If EDIT_FLG = "Y" And DLT_FLG <> "Y" And ALT_SUB_ID = "" Then
            If CST_FLG = "Y" Or ADMIN_ACCESS = "Y" Then
                temp3 = "&AID=&CTD=9&CIS=" & NOT_TYPE & "&CID=" & OrderID & "&POP=N&EID=" & EMP_ID
                If ASN_NAME <> "" Then temp3 = temp3 & "&ASN=" & EnURL(ASN_NAME)
                If ASN_ID <> "" Then temp3 = temp3 & "&ASI=" & EnURL(ASN_ID)
                If ASN_KEY <> "" Then temp3 = temp3 & "&ASK=" & EnURL(ASN_KEY)
                If Not (ASSOC_RSTR = "Y" And OWNER_FLG <> "Y") Then
                    ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/OpenContent?OpenAgent" & temp3 & "' class=""buttons2"">Add a Document</a></td>"
                End If
            Else
                temp3 = "&CID=" & OrderID
                If CAT_NAME <> "" Then temp3 = temp3 & "&CTN=" & EnURL(CAT_NAME)
                If CAT_ID <> "" Then temp3 = temp3 & "&CTD=" & EnURL(CAT_ID)
                If KEY_ID <> "" Then temp3 = temp3 & "&KEY=" & EnURL(KEY_ID)
                If ASN_NAME <> "" Then temp3 = temp3 & "&ASN=" & EnURL(ASN_NAME)
                If ASN_ID <> "" Then temp3 = temp3 & "&AID=" & EnURL(ASN_ID)
                If ASN_KEY <> "" Then temp3 = temp3 & "&ASK=" & EnURL(ASN_KEY)
                If ALT_SUB_ID <> "" Then temp3 = temp3 & "&ASB=" & EnURL(ALT_SUB_ID)
                If ASSOC_PARAM <> "" Then temp3 = temp3 & "&APM=" & EnURL(ASSOC_PARAM)
                If ASSOC_PARAM_TYPE <> "" Then temp3 = temp3 & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
                If ASSOC_DEF <> "" Then temp3 = temp3 & "&ADF=" & EnURL(ASSOC_DEF)
                If ASSOC_RSTR <> "" Then temp3 = temp3 & "&ASR=" & ASSOC_RSTR
                If DOMAIN <> "" Then temp3 = temp3 & "&DOM=" & DOMAIN
                If CST_FLG <> "" Then temp3 = temp3 & "&CST=" & CST_FLG
                If POPUP_FLG <> "" Then temp3 = temp3 & "&POP=" & POPUP_FLG
                If PUBLIC_FLG <> "" Then temp3 = temp3 & "&PUB=" & PUBLIC_FLG
                If EMP_ID <> "" Then temp3 = temp3 & "&EID=" & EMP_ID
                If CALL_SCREEN <> "" Then temp3 = temp3 & "&CIS=" & CALL_SCREEN
                If NOT_TYPE <> "" Then temp3 = temp3 & "&NTY=" & NOT_TYPE
                If CST_FLG <> "Y" Then
                    If Not (ASSOC_RSTR = "Y" And OWNER_FLG <> "Y") Then
                        ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/OpenContent?OpenAgent&NEW=Y&ID=" & temp3 & "' class=""buttons2"">Add a Document</a></td>"
                    End If
                End If
            End If
        End If
        If ASN_NAME = "Order" And ASN_KEY <> "" Then
            ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/SendReceipt?OpenAgent&ID=" & ASN_KEY & "&EID=" & EMP_ID & "&CID=" & OrderID & "&CIS=" & CALL_SCREEN & "' class=""buttons2"">Send Receipt</a></td>"
        End If
        If NOT_TYPE = "TRC" Then
            ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/UpdRecs?OpenAgent&CID=" & OrderID & "&TYP=GTC&STP=Card&RID=" & ASN_KEY & "' class=""buttons2"">Generate Card</a></td>"
            ButtonList = ButtonList & "<td class=""otherlinks""><a href='" & BasePath & "/UpdRecs?OpenAgent&CID=" & OrderID & "&TYP=GTCR&STP=Card&RID=" & ASN_KEY & "' class=""buttons2"">Generate Cert</a></td>"
        End If
        ' Previous screen
        If CALL_SCREEN = "RWD" Then
            ButtonList = ButtonList & "<td class=""otherlinks""><a href='javascript:history.go(-1)' class=""buttons2"">Previous</a></td>"
        End If
        If Debug = "Y" Then mydebuglog.Debug(" ButtonList: " & ButtonList & vbCrLf)
        If ButtonList <> "" Then
            PHeader = PHeader & "<table border=""0"" cellpadding=""0"" cellspacing=""1""><tr valign=""top"">" & ButtonList & "</tr></table>"
        End If
        PHeader = PHeader & "</td></tr>"

        ' Title
        PHeader = PHeader & "<tr><td width=""1152"" valign=""top"" bgcolor=""#F5ECDA"" class=""bodyKO"">" & vbCrLf
        PHeader = PHeader & "<div align=""left""><img src=""/images/spacer2.gif"" width=""3"" height=""8"">" & vbCrLf
        If CAT_NAME <> "" Then
            If UCase(Right(CAT_NAME, 1)) = "S" Then
                PTITLE = Trim(StrConv(DeURL(CAT_NAME), 3))
            Else
                PTITLE = Trim(StrConv(DeURL(CAT_NAME), 3) & " items ")
            End If
        Else
            PTITLE = "Documents "
        End If
        If DLT_FLG = "Y" Then PTITLE = "Deleted " & PTITLE
        If PRIMARY_ASSOC <> "" Then PTITLE = PTITLE & " found related to <i>" & PRIMARY_ASSOC & "</i>"
        If PRIMARY_ASSOC = "" Then PTITLE = PTITLE & " found "
        If Debug = "Y" Then mydebuglog.Debug(" PTITLE: " & PTITLE & vbCrLf)
        PHeader = PHeader & "<span class=""BigHeader""><strong>"
        If DOC_LIB_FLG = "Y" Then PHeader = PHeader & "Your "
        PHeader = PHeader & PTITLE & "</strong></span></div></td>" & vbCrLf
        PHeader = PHeader & "</tr></table>" & vbCrLf
        If Debug = "Y" Then mydebuglog.Debug(" PHeader: " & PHeader & vbCrLf)

        outstring = outstring & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & vbCrLf
        outstring = outstring & "<html>" & vbCrLf
        outstring = outstring & "<head>" & vbCrLf
        outstring = outstring & "<title>DMSA</title>" & vbCrLf
        outstring = outstring & "<meta http-equiv=""Pragma"" content=""no-cache"">" & vbCrLf
        outstring = outstring & "<meta http-equiv=""Cache-Control"" content=""no-cache, no-store, must-revalidate"">" & vbCrLf
        outstring = outstring & "<meta http-equiv=""Expires"" content=""-1"">" & vbCrLf
        outstring = outstring & "<link rel=""stylesheet"" href=""//datafluxapp/css/facebox.css"" media=""screen"" type=""text/css""/>" & vbCrLf
        outstring = outstring & "<link rel=""stylesheet"" href=""//datafluxapp/cs2/tips.css"" media=""screen"" type=""text/css""/>" & vbCrLf
        outstring = outstring & "<style>#pgbody { top: 10px; }</style>" & vbCrLf
        outstring = outstring & "</head>" & vbCrLf
        outstring = outstring & "<body text=""#000000"" bgcolor=""#FFFFFF"">" & vbCrLf
        outstring = outstring & "<div id=""pgbody"">" & vbCrLf
        outstring = outstring & "<center>" & vbCrLf
        outstring = outstring & PHeader & vbCrLf
        outstring = outstring & "<center><table width=""1000"" cellpadding=""0"" cellspacing=""0"" border=""0"" id=""ListResults""> " & vbCrLf
        outstring = outstring & "<tr><td class=""LargeDBody"">" & ErrMsg & "</td></tr>" & vbCrLf
        outstring = outstring & "</table></center><br/><br/><br/>" & vbCrLf
        outstring = outstring & "</body>" & vbCrLf
        outstring = outstring & "</html>" & vbCrLf

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
            ErrMsg = ErrMsg & "Unable to close the database connection(s): " & ex.ToString & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(ErrMsg) <> "" Then myeventlog.Error("GetContent.ashx " & ErrLvl & ": " & Trim(ErrMsg))
        myeventlog.Info("GetContent.ashx  Results for user id: " & EMP_ID & " And order id: " & OrderID)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(ErrMsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error " & Trim(ErrMsg))
                mydebuglog.Debug("Results for user id " & EMP_ID & " And order id: " & OrderID)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "GetContent", LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If
        If ErrMsg <> "" And outstring = "" Then GoTo DisplayErrorMsg

        ' ============================================
        ' WRITE RESULTS
        context.Response.Clear()
        context.Response.AddHeader("Content-Disposition", "inline; extension-token=" & Extension)
        context.Response.ContentType = "text/html"
        context.Response.Write(outstring)
        Exit Sub

DisplayErrorMsg:
        context.Response.ContentType = "text/html"
        context.Response.Write("<h2><b>" & ErrMsg & "</b></h2>")
        If Debug = "Y" Then
            Using writer As New StreamWriter("C:\Logs\GetContent_failed.log", True)
                writer.WriteLine(Now.ToString & " - " & ErrMsg & ", OrderID " & OrderID & ", EmpId: " & EMP_ID & ", BROWSER: " & BROWSER)
            End Using
        End If
    End Sub

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

    ' ================================================
    ' HELPER FUNCTIONS
    Function GetDocFilter(ByVal cmd As SqlCommand, ByVal dr As SqlDataReader, ByVal RecordType As String, ByVal RecordID As String, ByVal AssocId As String, ByVal UID As String,
                          ByVal DOMAIN As String, ByVal SUB_ID As String, ByVal CONTACT_ID As String, ByRef OptAssoc As String, ByRef OptAssocKey As String,
                          ByRef OptAssocId As String, ByRef OWNER_FLG As String, ByRef SqlS As String, ByVal PUBLIC_FLG As String, ByRef EDIT_FLG As String,
                          ByRef DEL_FLG As String, ByRef ADMIN_ACCESS As String, ByRef OUTPUT_FLG As String, ByVal REQD_FLAG As String, ByRef DOC_AREA As String,
                          ByVal mydebuglog As ILog, ByVal Debug As String) As String
        ' Translates a document type into an appropriate query to retrieve a list of these documents from the DMS
        ' INPUTS:
        '	cmd         			- ODBC connection to database
        '	RecordType				- The type of record being searched
        '	RecordID				- The record id of the association to be searched
        '    AssocId				- The id of the type
        '	UID					    - User id of the user
        '	DOMAIN					- Domain of user
        '	SUB_ID					- Subscription of user
        '	CONTACT_ID				- Contact id of user
        '	REQD_FLAG				- Association Required
        '	debug					- Debug flag

        ' OUTPUTS:
        ' 	GetDocFilter			- The name of the associated record	
        '	SqlS					- The SQL query for the DMS generated
        '	OptAssoc				- Optional association name
        '	OptAssocKey				- Optional association key
        '	OptAssocId				- Optional association id
        '	OWNER_FLG			    - Whether the contact is allowed to add documents to the associated record
        '	DOC_AREA				- The document area
        ' -----
        ' Declarations
        Dim Query1, Query2, temp, temp2, temp3, temp4, temp5, AREA, RecOwner, RecSub, ErrMsg As String
        Dim RIGHTS As String
        Dim RIGHTSLST As New Dictionary(Of String, String)
        ' -----
        ' Set Variable Defaults
        Query1 = ""
        Query2 = ""
        temp = ""
        temp2 = ""
        temp3 = ""
        temp4 = ""
        temp5 = ""
        RecOwner = ""
        GetDocFilter = ""
        AREA = ""
        ErrMsg = ""
        RecSub = ""
        RIGHTS = ""

        If Debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "Start GetDocFilter--------------------")
            mydebuglog.Debug(" ..RecordType " & RecordType)
            mydebuglog.Debug(" ..RecordID " & RecordID)
            mydebuglog.Debug(" ..AssocId " & AssocId)
            mydebuglog.Debug(" ..DOMAIN " & DOMAIN)
            mydebuglog.Debug(" ..SUB_ID " & SUB_ID)
            mydebuglog.Debug(" ..CONTACT_ID " & CONTACT_ID)
        End If

        If RecordID = "" Or RecordType = "" Then GoTo CloseOut
        If DOMAIN = "" And SUB_ID = "" And CONTACT_ID = "" Then GoTo CloseOut

        ' -----
        ' Perform action	 
        Select Case RecordType
            Case "Exam"
                AREA = "EXAMS"
                Query1 = "SELECT NAME " &
                "FROM siebeldb.dbo.S_CRSE_TST WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CRSE_TST", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Exam' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Survey"
                AREA = "EXAMS"
                Query1 = "SELECT NAME " &
                "FROM siebeldb.dbo.S_CRSE_TST WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CRSE_TST", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Survey' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Attempt"
                AREA = "EXAMS"
                Query1 = "SELECT 'Exam Registration # ' + LTRIM(CAST(MS_IDENT AS VARCHAR)) " &
                "FROM siebeldb.dbo.S_CRSE_TSTRUN WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CRSE_TST", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Assessment Attempt' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Assessment Attempt"
                AREA = "EXAMS"
                Query1 = "SELECT 'Exam Registration # ' + LTRIM(CAST(MS_IDENT AS VARCHAR)) " &
                "FROM siebeldb.dbo.S_CRSE_TSTRUN WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CRSE_TST", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Assessment Attempt' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Master Trainer"
                AREA = "MASTER TRAINERS"
                ' Locate related record
                Query1 = "SELECT FST_NAME+' '+LAST_NAME " &
                "FROM siebeldb.dbo.S_EMPLOYE WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_EMPLOYEE", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Contact' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Session"
                AREA = "SESSIONS"
                If IsNumeric(RecordID) Then
                    Query1 = "SELECT 'Session # ' +CAST(SESS_ID AS VARCHAR) " &
                    "FROM siebeldb.dbo.CX_SESSIONS_X WHERE SESS_ID='" & RecordID & "'"
                    temp = GetSingleRecord(RecordType, "CX_SESSIONS_X", cmd, Query1, mydebuglog, Debug)
                Else
                    Query1 = "SELECT 'Session # ' +CAST(SESS_ID AS VARCHAR) " &
                    "FROM siebeldb.dbo.CX_SESSIONS_X WHERE ROW_ID='" & RecordID & "'"
                    temp = GetSingleRecord(RecordType, "CX_SESSIONS_X", cmd, Query1, mydebuglog, Debug)

                    Query1 = "SELECT SESS_ID FROM siebeldb.dbo.CX_SESSIONS_X WHERE ROW_ID='" & RecordID & "'"
                    RecordID = GetSingleRecord(RecordType, "CX_SESSIONS_X", cmd, Query1, mydebuglog, Debug)
                    If Debug = "Y" Then mydebuglog.Debug(" .. New RecordID:" & RecordID)
                End If
                If Debug = "Y" Then mydebuglog.Debug(" .. temp:" & temp)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Session' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Workshop"
                AREA = "WORKSHOPS"
                Query1 = "SELECT 'Workshop # '+CAST(X_WSID AS VARCHAR) " &
                "FROM siebeldb.dbo.S_CRSE_OFFR WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CRSE_OFFR", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Workshop' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Contact"
                AREA = "CONTACTS"
                RecOwner = RecordID
                ' Locate related record
                Query1 = "SELECT FST_NAME+' '+LAST_NAME " &
                "FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CONTACT", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Contact' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Trainer"
                AREA = "TRAINERS"
                ' Locate related record
                Query1 = "SELECT FST_NAME+' '+LAST_NAME " &
                "FROM siebeldb.dbo.S_CONTACT WHERE X_TRAINER_NUM='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CONTACT", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Trainer' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Session Registration"
                AREA = "TRAINING MANAGEMENT"
                Query1 = "SELECT '# '+CAST(R.MS_IDENT AS VARCHAR)+', '+C.FST_NAME+' '+C.LAST_NAME " &
                "FROM siebeldb.dbo.CX_SESS_REG R " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " &
                "WHERE R.ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_SESS_REG", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Session Registration' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Workshop Registration"
                AREA = "TRAINING MANAGEMENT"
                Query1 = "SELECT '# '+CAST(R.MS_IDENT AS VARCHAR)+', '+C.FST_NAME+' '+C.LAST_NAME " &
                "FROM siebeldb.dbo.S_CRSE_REG R " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.PERSON_ID " &
                "WHERE R.ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CRSE_REG", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Workshop Registration' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Organization"
                RecordType = "Account"

            Case "Account"
                AREA = "ACCOUNTS"
                ' Set the owner to be the primary billing contact - need to revisit this
                Query1 = "SELECT PR_BL_PER_ID " &
                "FROM siebeldb.dbo.S_ORG_EXT WHERE ROW_ID='" & RecordID & "'"
                RecOwner = GetSingleRecord(RecordType, "S_ORG_EXT", cmd, Query1, mydebuglog, Debug)

                ' Locate related record
                Query1 = "SELECT NAME+' '+(SELECT CASE WHEN LOC IS NOT NULL THEN LOC ELSE '' END) " &
                "FROM siebeldb.dbo.S_ORG_EXT WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_ORG_EXT", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp

                ' Locate documents associated with this account
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Account' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Activity"
                AREA = "SERVICE REQUESTS"
                Query1 = "SELECT CAST(MS_IDENT AS VARCHAR) " &
                "FROM siebeldb.dbo.S_EVT_ACT WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_EVT_ACT", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Activity' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Service Request"
                AREA = "SERVICE REQUESTS"
                Query1 = "SELECT CAST(MS_IDENT AS VARCHAR) " &
                "FROM siebeldb.dbo.S_SRV_REQ WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_SRV_REQ", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Service Request' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Training Request"
                AREA = "TRAINING MANAGEMENT"
                Query1 = "SELECT CAST(MS_IDENT AS VARCHAR) " &
                "FROM siebeldb.dbo.CX_SUB_TRN_REQ WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_SUB_TRN_REQ", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Training Request' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Subscription"
                AREA = "SUPERVISOR"
                Query1 = "SELECT RTRIM(SVC_TYPE)a+' for '+ROW_ID FROM siebeldb.dbo.CX_SUBSCRIPTION WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_SUBSCRIPTION", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Subscription' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Course"
                AREA = "COURSE AUTHORING"
                Query1 = "SELECT (SELECT CASE WHEN CHARINDEX(C.X_SUMMARY_CD,C.NAME)>0 THEN '' ELSE " &
                "(SELECT CASE WHEN C.X_SUMMARY_CD IS NOT NULL THEN C.X_SUMMARY_CD+' ' ELSE '' END) END)+C.NAME " &
                "FROM siebeldb.dbo.S_CRSE C WHERE C.ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CRSE", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Course' AND DA.fkey='" & RecordID & "'"
                End If

                If OptAssoc <> "" And OptAssocKey <> "" Then
                    If OptAssocId <> "" Then
                        Query2 = Query2 & " AND EXISTS (SELECT oda.doc_id  " &
                        "FROM DMS.dbo.Document_Associations oda " &
                        "LEFT OUTER JOIN DMS.dbo.Associations oa ON oa.row_id=oda.association_id " &
                        "WHERE oda.row_id='" & OptAssocId & "' AND oda.fkey='" & OptAssocKey & "' AND oda.doc_id=D.row_id)"
                    Else
                        Query2 = Query2 & " AND EXISTS (SELECT oda.doc_id  " &
                        "FROM DMS.dbo.Document_Associations oda " &
                        "LEFT OUTER JOIN DMS.dbo.Associations oa ON oa.row_id=oda.association_id " &
                        "WHERE oa.name='" & OptAssoc & "' AND oda.fkey='" & OptAssocKey & "' AND oda.doc_id=D.row_id)"
                    End If
                End If

            Case "Curriculum"
                AREA = "COURSE AUTHORING"
                Query1 = "SELECT (SELECT CASE WHEN CHARINDEX(C.X_SUMMARY_CD,C.NAME)>0 THEN '' ELSE " &
                "(SELECT CASE WHEN C.X_SUMMARY_CD IS NOT NULL THEN C.X_SUMMARY_CD+' ' ELSE '' END) END)+C.NAME " &
                "FROM siebeldb.dbo.S_CURRCLM C WHERE C.ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CURRCLM", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Curriculum' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Jurisdiction"
                AREA = "BASIC"
                Query1 = "SELECT NAME FROM siebeldb.dbo.CX_JURISDICTION_X WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_JURISDICTION_X", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Jurisdiction' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Participant"
                AREA = "PARTICIPANTS"
                Query1 = "SELECT FST_NAME+' '+LAST_NAME " &
                "FROM siebeldb.dbo.CX_PARTICIPANT_X WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_PARTICIPANT_X", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Participant' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Participation"
                AREA = "PARTICIPANTS"
                Query1 = "SELECT 'Participation # '+CAST(PART_NUM AS VARCHAR) FROM siebeldb.dbo.CX_SESS_PART_X WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_SESS_PART_X", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Participation' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Trainer Certification"
                AREA = "TRAINERS"
                Query1 = "SELECT 'Trainer Certification # '+CAST(P.MS_IDENT AS VARCHAR) + ' tested on '+CONVERT(VARCHAR,T.TEST_DT,101) " &
                "FROM siebeldb.dbo.S_CURRCLM_PER P  " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN T ON T.ROW_ID=P.X_CRSE_TSTRUN_ID " &
                "WHERE P.ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_CURRCLM_PER", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Trainer Certification' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Order"
                AREA = "ORDERS"
                Query1 = "SELECT 'Invoice # '+CAST(X_INVOICE_NUM AS VARCHAR) " &
                "FROM siebeldb.dbo.S_ORDER  " &
                "WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "S_ORDER", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Order' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Scheduled Session"
                RecordType = "Planned Session"

            Case "Planned Session"
                AREA = "TRAINING MANAGEMENT"
                ' Locate documents associated with the course and this session, as well as related record
                Query2 = "DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "' AND DA.pr_flag<>'Y'"
                ' Locate defails about this session
                Query1 = "SELECT S.CRSE_ID, C.NAME, S.START_DT, S.END_DT, T.FST_NAME, T.LAST_NAME, " &
                "H.NAME, H.LOC, A.CITY, A.STATE, C.TYPE_CD, T.ROW_ID, S.SUB_ID " &
                "FROM siebeldb.dbo.CX_TRAIN_OFFR S " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CRSE C ON C.ROW_ID=S.CRSE_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT T ON T.ROW_ID=S.TRAINER_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT H ON H.ROW_ID=S.HELD_OU_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A ON A.ROW_ID=S.HELD_ADDR_ID " &
                "WHERE S.ROW_ID='" & RecordID & "'"
                If Debug = "Y" Then mydebuglog.Debug("Locate " & RecordType & ": " & Query1)
                Try
                    cmd.CommandText = Query1
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                RecOwner = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                RecSub = CheckDBNull(dr(12), enumObjectType.StrType)
                                If temp <> "" Then
                                    If AssocId <> "" Then
                                        Query2 = "AND ((" & Query2 & ") or ( DA.association_id=15 AND DA.fkey='" & temp & "')) "
                                    Else
                                        Query2 = "AND ((" & Query2 & ") or (A.name='Course' AND DA.fkey='" & temp & "')) "
                                    End If
                                    OptAssoc = "Course"
                                    OptAssocKey = temp
                                End If

                                ' Compute name
                                temp = Trim(CheckDBNull(dr(2), enumObjectType.StrType)) & " scheduled session"
                                If CheckDBNull(dr(10), enumObjectType.StrType) <> "Web Based" Then
                                    temp = temp & " on " & Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                    If Trim(CheckDBNull(dr(3), enumObjectType.StrType)) <> Trim(CheckDBNull(dr(2), enumObjectType.StrType)) Then temp = temp & " to " & Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                    If Trim(CheckDBNull(dr(4), enumObjectType.StrType)) <> "" Then temp = temp & ", Trained by " & Trim(CheckDBNull(dr(4), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                    If Trim(CheckDBNull(dr(6), enumObjectType.StrType)) <> "" Then temp = temp & ", At " & Trim(CheckDBNull(dr(6), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                    If Trim(CheckDBNull(dr(8), enumObjectType.StrType)) <> "" Then temp = temp & ", In " & Trim(CheckDBNull(dr(8), enumObjectType.StrType)) & ", " & Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                                End If
                                If temp <> "" Then GetDocFilter = temp
                            Catch ex2 As Exception
                                ErrMsg = ErrMsg & "Error getting " & RecordType & ": " & ex2.ToString
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        ErrMsg = ErrMsg & "Error getting " & RecordType
                    End If
                    dr.Close()
                Catch ex As Exception
                    ErrMsg = ErrMsg & "Error getting " & RecordType & ": " & ex.ToString
                End Try
                If Debug = "Y" Then mydebuglog.Debug("Found Session: " & GetDocFilter)

            Case "Scheduled Workshop"
                RecordType = "Planned Workshop"

            Case "Planned Workshop"
                If Debug = "Y" Then Call mydebuglog.Debug("Checking Planned Workshop " & RecordID)
                AREA = "TRAINING MANAGEMENT"
                ' Locate documents associated with the course and this workshop, as well as related record
                Query2 = "AND A.name='Planned Workshop' AND DA.fkey='" & RecordID & "'"
                ' Locate details about this workshop
                Query1 = "SELECT S.CRSE_ID, CR.NAME, S.START_DT, S.END_DT, T.FST_NAME, T.LAST_NAME, " &
                "H.NAME, H.LOC, A.CITY, A.STATE, C.ROW_ID " &
                "FROM siebeldb.dbo.S_CRSE_OFFR S " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=S.CRSE_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.S_EMPLOYEE T ON T.ROW_ID=S.INSTRUCTOR_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.REG_AS_EMP_ID=T.ROW_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT H ON H.ROW_ID=S.X_ACCOUNT_HELD " &
                "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A ON A.ROW_ID=S.X_HELD_ADDRESS_ID " &
                "WHERE S.ROW_ID='" & RecordID & "'"
                If Debug = "Y" Then mydebuglog.Debug("Locate " & RecordType & ": " & Query1)
                Try
                    cmd.CommandText = Query1
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        Try
                            temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            RecOwner = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                            If temp <> "" Then
                                If AssocId <> "" Then
                                    Query2 = "AND ((" & Query2 & ") or (DA.association_id=" & AssocId & " AND DA.fkey='" & temp & "'))"
                                Else
                                    Query2 = "AND ((" & Query2 & ") or (A.name='Course' AND DA.fkey='" & temp & "'))"
                                End If
                                OptAssoc = "Course"
                                OptAssocKey = temp
                            End If
                            ' Compute name
                            temp = Trim(CheckDBNull(dr(1), enumObjectType.StrType)) & " scheduled session<br/>&nbsp;On " & Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            If Trim(CheckDBNull(dr(3), enumObjectType.StrType)) <> Trim(CheckDBNull(dr(2), enumObjectType.StrType)) Then temp = temp & " to " & Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            If Trim(CheckDBNull(dr(4), enumObjectType.StrType)) <> "" Then temp = temp & " <br/>&nbsp;Trained by " & Trim(CheckDBNull(dr(4), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                            If Trim(CheckDBNull(dr(6), enumObjectType.StrType)) <> "" Then temp = temp & "<br/>&nbsp;At " & Trim(CheckDBNull(dr(6), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                            If Trim(CheckDBNull(dr(8), enumObjectType.StrType)) <> "" Then temp = temp & "<br/>&nbsp;In " & Trim(CheckDBNull(dr(8), enumObjectType.StrType)) & ", " & Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                            If temp <> "" Then GetDocFilter = temp
                        Catch ex2 As Exception
                            ErrMsg = ErrMsg & "Error getting " & RecordType & ": " & ex2.ToString
                        End Try
                    Else
                        ErrMsg = ErrMsg & "Error getting " & RecordType
                    End If
                Catch ex As Exception
                    ErrMsg = ErrMsg & "Error getting " & RecordType & ": " & ex.ToString
                End Try

            Case "Trainer Permit"
                AREA = "TRAINER PERMITS"
                Query1 = "SELECT 'Permit # '+PERMIT_NO FROM siebeldb.dbo.CX_TRAINER_PERMITS WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_TRAINER_PERMITS", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Trainer Permit' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Organization Violation"
                AREA = "VIOLATIONS"
                Query1 = "SELECT 'Violation Case # '+CASE_NUMBER FROM siebeldb.dbo.CX_OU_VIOLATIONS WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_OU_VIOLATIONS", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Violation' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Organization Permit"
                AREA = "SITE PERMITS"
                Query1 = "SELECT 'Permit # '+PERMIT_NUM FROM siebeldb.dbo.CX_OU_PERMITS WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_OU_PERMITS", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Permit' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Course Feedback"
                AREA = "COURSE AUTHORING"
                Query1 = "SELECT 'Comment # '+CAST(ID AS VARCHAR) FROM elearning.dbo.ELN_FEEDBACK WHERE ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "ELN_FEEDBACK", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Course Feedback' AND DA.fkey='" & RecordID & "'"
                End If

            Case "Participant Trainer"

            Case "Notification Service"
                AREA = "NOTIFICATION SERVICE"
                Query1 = "SELECT 'Notification # '+CAST(MS_IDENT AS VARCHAR) FROM siebeldb.dbo.CX_CON_MSG_SVC WHERE ROW_ID='" & RecordID & "'"
                temp = GetSingleRecord(RecordType, "CX_CON_MSG_SVC", cmd, Query1, mydebuglog, Debug)
                If temp <> "" Then GetDocFilter = temp
                If AssocId <> "" Then
                    Query2 = "AND DA.association_id=" & AssocId & " AND DA.fkey='" & RecordID & "'"
                Else
                    Query2 = "AND A.name='Notification Service' AND DA.fkey='" & RecordID & "'"
                End If
        End Select

        ' -----	
        ' Set the OWNER_FLG field based on access to the area specified
        '  The only people who can add documents are administrators and record owners with "Edit" access to the associated record
CheckOwner:
        EDIT_FLG = "N"
        If AREA <> "" Then
            If UID = "" Then
                OWNER_FLG = "N"
            Else
                If Debug = "Y" Then
                    mydebuglog.Debug(vbCrLf & " AREA TO CHECK: " & AREA)
                    mydebuglog.Debug(" RecOwner: " & RecOwner)
                    mydebuglog.Debug(" CONTACT_ID: " & CONTACT_ID)
                End If
                ' Check admin access
                SqlS = "SELECT P.RIGHTS, SC.SYSADMIN_FLG FROM siebeldb.dbo.CX_SUB_CON_PRIV P " &
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.ROW_ID=P.SUB_CON_ID " &
                    "WHERE CON_ID='" & CONTACT_ID & "' AND UPPER(P.FCTN)='" & AREA & "'"
                If Debug = "Y" Then mydebuglog.Debug(" Check admin access: " & SqlS & vbCrLf)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        Try
                            While dr.Read()
                                RIGHTS = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                ADMIN_ACCESS = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                If InStr(RIGHTS, "A") > -1 Then ADMIN_ACCESS = "Y"
                                If ADMIN_ACCESS = "Y" Then OWNER_FLG = "Y"
                                If InStr(RIGHTS, "E") > -1 Then EDIT_FLG = "Y"
                                If InStr(RIGHTS, "D") > -1 Then DEL_FLG = "Y"
                                If InStr(RIGHTS, "O") > -1 Then OUTPUT_FLG = "Y"
                                If EDIT_FLG <> "Y" Then
                                    If OWNER_FLG <> "Y" Then OWNER_FLG = "N"
                                Else
                                    If SUB_ID = RecSub Then OWNER_FLG = "Y"
                                End If
                            End While
                        Catch ex2 As Exception
                            ErrMsg = ErrMsg & "Error getting privileges: " & ex2.ToString
                        End Try
                    Else
                        ErrMsg = ErrMsg & "Error getting privileges"
                    End If
                    dr.Close()
                Catch ex As Exception
                    ErrMsg = ErrMsg & "Error getting privileges: " & ex.ToString
                End Try
            End If
        End If

CloseOut:
        SqlS = Query2
        DOC_AREA = AREA
        If ErrMsg <> "" Then GetDocFilter = ErrMsg
        If Debug = "Y" Then
            mydebuglog.Debug(" SqlS: " & SqlS)
            mydebuglog.Debug(" RIGHTS: " & RIGHTS)
            mydebuglog.Debug(" EDIT_FLG: " & EDIT_FLG)
            mydebuglog.Debug(" DEL_FLG: " & DEL_FLG)
            mydebuglog.Debug(" OUTPUT_FLG: " & OUTPUT_FLG)
            mydebuglog.Debug(" DOC_AREA: " & DOC_AREA)
            mydebuglog.Debug(" Returning: " & GetDocFilter)
            mydebuglog.Debug("GetDocFilter End----------------------" & vbCrLf)
        End If
        Return GetDocFilter
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

    Public Function ExecQuery(ByVal QType As String, ByVal QRec As String, ByVal cmd As SqlCommand, ByVal SqlS As String, ByVal mydebuglog As ILog, ByVal Debug As String) As String
        Dim returnv As Integer
        Dim errmsg As String
        errmsg = ""
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  " & QType & " " & QRec & " record: " & SqlS)
        Try
            cmd.CommandText = SqlS
            returnv = cmd.ExecuteNonQuery()
            If returnv = 0 Then
                errmsg = errmsg & "The " & QRec & " record was not " & QType & vbCrLf
            End If
        Catch ex As Exception
            errmsg = errmsg & "Error " & QType & " record. " & ex.ToString & vbCrLf & "Query: " & SqlS
        End Try
        Return errmsg
    End Function

    Public Function GetSingleRecord(ByVal QType As String, ByVal QRec As String, ByVal cmd As SqlCommand, ByVal SqlS As String, ByVal mydebuglog As ILog, ByVal Debug As String) As String
        Dim errmsg As String = ""
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  " & QType & " " & QRec & " record: " & SqlS)
        Try
            cmd.CommandText = SqlS
            Dim result As Object = cmd.ExecuteScalar()
            If result <> Nothing Then
                GetSingleRecord = result.ToString()
                Return GetSingleRecord
            Else
                errmsg = errmsg & "The " & QRec & " record was not " & QType & vbCrLf
            End If
        Catch ex As Exception
            errmsg = errmsg & "Error " & QType & " record. " & ex.ToString & vbCrLf & "Query: " & SqlS
        End Try
        Return errmsg
    End Function

    Public Function CheckDBNull(ByVal obj As Object, Optional ByVal ObjectType As enumObjectType = enumObjectType.StrType) As Object
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

    ' String functions
    Function DeURL(ByVal theString As String) As String
        Dim x As String
        x = InStr(1, theString, "+")
        If x > 0 Then
            DeURL = Replace(theString, "+", " ")
        Else
            DeURL = theString
        End If
    End Function

    Function rspURL(ByVal theString As String) As String
        ' Replaces hex space with real space
        Dim x As String
        x = InStr(1, theString, "%20")
        If x > 0 Then
            rspURL = Replace(theString, "%20", " ")
        Else
            rspURL = theString
        End If
    End Function

    Function EnURL(ByVal theString As String) As String
        ' Replaces space with "+"
        Dim x As String
        x = InStr(1, theString, " ")
        If x > 0 Then
            EnURL = Replace(theString, " ", "+")
        Else
            EnURL = theString
        End If
    End Function

End Class