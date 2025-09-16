<%@ WebHandler Language="VB" Class="GetContentBE" %>

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
Imports System.Text
Imports Newtonsoft.Json.Converters
Imports log4net
Imports CachingWrapper.LocalCache
Imports Newtonsoft.Json.Linq

Public Class GetContentBE : Implements IHttpHandler

    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum

    Public Structure DocsRecord
        Dim DocId As String             ' Documents.row_id
        Dim DocName As String           ' Documents.name
        Dim DocFileName As String       ' Documents.dfilename
        Dim DocDesc As String           ' Documents.description
        Dim DocType As String           ' Document_Types.name
        Dim DocCreated As String        ' Documents.created
        Dim DocUpdated As String        ' Documents.last_upd
        Dim DocRights As String         ' Document_Users.access_type
    End Structure

    Public Structure DocCategory
        Dim DocId As String             ' Documents.row_id
        Dim DocCat As String            ' Categories.name
        Dim DocCatId As String          ' Catagories.row_id	
        Dim DocCatPr As String          ' Document_Categories.pr_flag
    End Structure

    Public Sub ProcessRequest(context As HttpContext) Implements IHttpHandler.ProcessRequest

        ' This service looks up order OrderID in the database specified by Source, determines
        ' the weight, and computes the freight amounts by type, which it returns in a JSON document  

        ' Parameter declarations
        Dim Debug, callback As String

        ' Context declarations
        Dim NextLink, BasePath As String
        Dim PrevLink As String = Trim(context.Request.ServerVariables("HTTP_REFERER"))
        Dim BROWSER As String = Trim(context.Request.ServerVariables("HTTP_USER_AGENT"))

        ' Result declarations
        Dim jdoc As String
        Dim results, errmsg, ErrLvl As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("PDDDebugLog")
        Dim logfile, tempdebug As String
        Dim Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"
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

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        Dim DMSServices As New local.hq.datafluxapp.dms.Service

        ' Search declarations
        Dim sStart, sAmount, sEcho, sCol, sDir, sSearch As String
        Dim Extension, LANG_CD, temp As String
        Dim SDOC_ID, CAT_ID, KEY_ID, CAT_NAME, ASN_NAME, ASN_ID, ASN_KEY, ASN_OPT, KEY_NAME, CAT_NAME_S As String
        Dim DOC_EXT, DLT_FLG, FTEXT, StartDt, EndDt, DOC_NAME, DESC_TEXT, DOC_EXT_ID, TYPE_ADMIN As String
        Dim PUBLIC_FLG, DMS_FLG, SUB_FLG, USR_FLG, OWN_FLG, ALT_SUB_ID, DOC_LIB_FLG, DOC_AREA As String
        Dim OptAssoc, OptAssocKey, OptAssocId, PRIMARY_ASSOC, SearchScope, relink As String
        Dim SortOrd, SortDir, GroupBy, OrderBy As String
        Dim DOMAIN, POPUP_FLG, CALL_ID, CALL_SCREEN, NOT_TYPE, CST_FLG, CAT_FLG, ASC_FLG, ASSOC_PARAM, ASSOC_PARAM_TYPE, ASSOC_DEF, ASSOC_RSTR As String
        Dim CONTACT_ID, DMS_DOMAIN_ID, DMS_ASUB_ID, SUB_ID, DMS_SUB_ID, EMP_ID, DMS_USER_ID, DMS_UA_ID, DMS_USER_AID, UID, OWNER_FLG, EDIT_FLG As String
        Dim SYSADMIN_FLG, TRAINING_ACCESS, TRAINER_ACC_FLG, MT_FLG, SVC_TYPE, TRAINING_FLG, TRAINER_FLG, TRAINER_ID, PART_ID, PART_FLG, SITE_ONLY As String
        Dim REG_NUM, pSessID, securebase As String
        Dim SUPER_FLG, PART_ACC_FLG, REG_AS_EMP_ID, DOMAIN_FLG, MT_ID, REPORTS_FLG, READ_FLG, DEL_FLG, OUTPUT_FLG, ADMIN_ACCESS, EXPORT_ACCESS As String
        Dim DOC_ID, DOC_TITLE, DOC_DESC, DOC_DATE, DOC_TYPE, DOC_RIGHTS, DOC_CATEGORY, DOC_FILENAME As String
        Dim editlink, attachlink, dellink, restorelink, publishlink, assoclink, emaillink As String
        Dim TotalDocs As Integer = 0
        Dim OrderID As String = ""
        Dim Refresh As Boolean = False
        Dim TableName As String = ""
        Dim edlink, aslink, pulink As String
        Dim sEnd As String
        Dim FilterCount, sc, TotalRecs As Integer
        Dim SearchClause, CM_TRAN_ID As String
        Dim NoCols As Integer
        Dim FUNCTION_ID, iTotalRecords As String

        ' ============================================
        ' Variable setup
        Debug = "Y"
        OrderID = ""
        jdoc = ""
        Logging = "Y"
        errmsg = ""
        results = "Failure"
        SqlS = ""
        callback = ""
        Extension = ""
        LANG_CD = "ENU"
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
        temp = ""
        GroupBy = ""
        OrderBy = ""
        NextLink = ""
        relink = ""
        BasePath = ""
        ErrLvl = "Warning"
        sStart = "0"
        sAmount = "10"
        sEcho = "1"
        sCol = ""
        sDir = ""
        sSearch = ""
        REG_NUM = ""
        pSessID = ""
        NoCols = 0
        CM_TRAN_ID = ""
        SearchClause = ""
        FilterCount = 0
        TotalRecs = 0
        FUNCTION_ID = ""
        iTotalRecords = ""

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("GetContentBE_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\GetContentBE.log"
            Try
                log4net.GlobalContext.Properties("PDDLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If

        ' ============================================
        ' Get parameters    
        ' Datatable parameters
        If Not context.Request.QueryString("iDisplayStart") Is Nothing Then sStart = HttpUtility.UrlDecode(context.Request.QueryString("iDisplayStart"))
        If Not context.Request.QueryString("iDisplayLength") Is Nothing Then sAmount = HttpUtility.UrlDecode(context.Request.QueryString("iDisplayLength"))
        If Not context.Request.QueryString("start") Is Nothing Then sStart = HttpUtility.UrlDecode(context.Request.QueryString("start"))
        If Not context.Request.QueryString("length") Is Nothing Then sAmount = HttpUtility.UrlDecode(context.Request.QueryString("length"))
        If sAmount = "-1" Then sAmount = "32000000"
        If Not context.Request.QueryString("sEcho") Is Nothing Then sEcho = HttpUtility.UrlDecode(context.Request.QueryString("sEcho"))
        If sEcho = "" Then sEcho = "1"
        If Not context.Request.QueryString("iSortCol_0") Is Nothing Then sCol = HttpUtility.UrlDecode(context.Request.QueryString("iSortCol_0"))
        If Not context.Request.QueryString("sSortDir_0") Is Nothing Then sDir = UCase(HttpUtility.UrlDecode(context.Request.QueryString("sSortDir_0")))
        If Not context.Request.QueryString("sSearch") Is Nothing Then sSearch = UCase(HttpUtility.UrlDecode(context.Request.QueryString("sSearch")))
        If Not context.Request.QueryString("callback") Is Nothing Then callback = HttpUtility.UrlDecode(context.Request.QueryString("callback"))

        ' Context parameters            
        If Not context.Request.QueryString("RG") Is Nothing Then REG_NUM = HttpUtility.UrlDecode(context.Request.QueryString("RG"))
        If Not context.Request.QueryString("SESS") Is Nothing Then pSessID = Trim(DeURL(context.Request.QueryString("SESS")))
        If Not context.Request.QueryString("CID") Is Nothing Then CALL_ID = context.Request.QueryString("CID")
        OrderID = CALL_ID
        If Not context.Request.QueryString("BP") Is Nothing Then BasePath = context.Request.QueryString("BP")
        securebase = Replace(BasePath, "http:", "https:")

        '	Criteria filters
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
        If Not context.Request.QueryString("ASB") Is Nothing Then ALT_SUB_ID = context.Request.QueryString("ASB")           ' Alternative subscription id for searches		
        If Not context.Request.QueryString("DLF") Is Nothing Then DOC_LIB_FLG = context.Request.QueryString("DLF")          ' Flag to indicate that we are seeking all docs associated with the current user
        If Not context.Request.QueryString("EID") Is Nothing Then EMP_ID = context.Request.QueryString("EID")               ' Employee Id of user	

        '	Sorting 
        If Not context.Request.QueryString("SO") Is Nothing Then SortOrd = context.Request.QueryString("SO")                ' Sort Order
        If sCol = "" Then sCol = SortOrd
        If sCol = "" Then sCol = "0"
        If Not context.Request.QueryString("SD") Is Nothing Then SortDir = context.Request.QueryString("SD")                ' Sort Direction
        If sDir = "" Then sDir = SortDir
        If sDir = "" Then sDir = "ASC"

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
            mydebuglog.Debug("  CALL_ID: " & CALL_ID)
            mydebuglog.Debug("  pSessID: " & pSessID)
            mydebuglog.Debug("  REG_NUM: " & REG_NUM)
            mydebuglog.Debug("  sStart: " & sStart)
            mydebuglog.Debug("  sAmount: " & sAmount)
            mydebuglog.Debug("  sEcho: " & sEcho)
            mydebuglog.Debug("  sCol: " & sCol)
            mydebuglog.Debug("  SortOrd: " & SortOrd)
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
        If OrderID = "" Then
            errmsg = errmsg & vbCrLf & "Parameter()s missing"
            GoTo CloseOut2
        End If
        If EMP_ID = "" Then
            ErrLvl = "Error"
            Select Case LANG_CD
                Case "ESN"
                    errmsg = "Solicitud no v&aacute;lida. No se ha especificado ning&uacute;n ID de documento o versi&oacute;n."
                Case Else
                    errmsg = "Invalid request. No employee id specified."
            End Select
            GoTo CloseOut2
        End If

        ' ================================================
        ' BUILD INVOCATION URLS
        If ASN_NAME <> "" Then relink = relink & "&ASN=" & EnURL(ASN_NAME)
        If ASN_ID <> "" Then relink = relink & "&ASI=" & EnURL(ASN_ID)
        If ASN_KEY <> "" Then relink = relink & "&ASK=" & EnURL(ASN_KEY)
        If DOC_EXT <> "" Then relink = relink & "&DCT=" & DOC_EXT
        If DLT_FLG <> "" Then relink = relink & "&DCR=" & DLT_FLG
        If ALT_SUB_ID <> "" Then relink = relink & "&ASB=" & ALT_SUB_ID
        If FTEXT <> "" Then relink = relink & "&FTX=" & EnURL(FTEXT)
        If StartDt <> "" Then relink = relink & "&DTH=" & StartDt
        If EndDt <> "" Then relink = relink & "&EDT=" & EndDt
        If DOC_NAME <> "" Then relink = relink & "&NME=" & EnURL(DOC_NAME)
        If DESC_TEXT <> "" Then relink = relink & "&DSC=" & EnURL(DESC_TEXT)
        If DMS_FLG <> "" Then relink = relink & "&DMU=" & DMS_FLG
        If SUB_FLG <> "" Then relink = relink & "&SBU=" & SUB_FLG
        If USR_FLG <> "" Then relink = relink & "&USR=" & USR_FLG
        If OWN_FLG <> "" Then relink = relink & "&OWN=" & OWN_FLG
        If SortOrd <> "" Then relink = relink & "&SO=" & EnURL(SortOrd)
        If SortDir <> "" Then relink = relink & "&SD=" & EnURL(SortDir)
        If DOMAIN <> "" Then relink = relink & "&DOM=" & DOMAIN
        If POPUP_FLG <> "" Then relink = relink & "&POP=" & POPUP_FLG
        If CALL_ID <> "" Then relink = relink & "&CID=" & CALL_ID & "&CIS=CNT"
        If CST_FLG <> "" Then relink = relink & "&CST=" & CST_FLG
        If CAT_FLG <> "" Then relink = relink & "&CFL=" & CAT_FLG
        If ASC_FLG <> "" Then relink = relink & "&AFL=" & ASC_FLG
        If ASSOC_PARAM <> "" Then relink = relink & "&APM=" & ASSOC_PARAM
        If ASSOC_PARAM_TYPE <> "" Then relink = relink & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
        If ASSOC_DEF <> "" Then relink = relink & "&ADF=" & ASSOC_DEF
        If ASSOC_RSTR <> "" Then relink = relink & "&ASR=" & ASSOC_RSTR
        If Debug = "Y" Then mydebuglog.Debug("Requery URL: " & relink)

        edlink = ""
        If CALL_ID <> "" Then edlink = edlink & "&CID=" & CALL_ID & "&CIS=CNT"
        If CAT_ID <> "" Then edlink = edlink & "&CTD=" & EnURL(CAT_ID)
        If ASN_NAME <> "" Then edlink = edlink & "&ASN=" & EnURL(ASN_NAME)
        If ASN_ID <> "" Then edlink = edlink & "&AID=" & EnURL(ASN_ID)
        If ASN_KEY <> "" Then edlink = edlink & "&ASK=" & EnURL(ASN_KEY)
        If ALT_SUB_ID <> "" Then edlink = edlink & "&ASB=" & EnURL(ALT_SUB_ID)
        If ASSOC_PARAM <> "" Then edlink = edlink & "&APM=" & EnURL(ASSOC_PARAM)
        If ASSOC_PARAM_TYPE <> "" Then edlink = edlink & "&APT=" & EnURL(ASSOC_PARAM_TYPE)
        If ASSOC_DEF <> "" Then edlink = edlink & "&ADF=" & EnURL(ASSOC_DEF)
        If ASSOC_RSTR <> "" Then edlink = edlink & "&ASR=" & ASSOC_RSTR
        If DOMAIN <> "" Then edlink = edlink & "&DOM=" & DOMAIN
        If POPUP_FLG <> "" Then edlink = edlink & "&POP=" & POPUP_FLG
        If CST_FLG <> "" Then edlink = edlink & "&CST=" & CST_FLG
        If PUBLIC_FLG <> "" Then edlink = edlink & "&PUB=" & PUBLIC_FLG
        If Debug = "Y" Then mydebuglog.Debug("Edit URL: " & edlink)

        pulink = ""
        If CALL_ID <> "" Then pulink = pulink & "&CID=" & CALL_ID & "&CIS=CNT"
        If DOMAIN <> "" Then pulink = pulink & "&DOM=" & DOMAIN
        If POPUP_FLG <> "" Then pulink = pulink & "&POP=" & POPUP_FLG
        If CST_FLG <> "" Then pulink = pulink & "&CST=" & CST_FLG
        If PUBLIC_FLG <> "" Then pulink = pulink & "&PUB=" & PUBLIC_FLG
        If Debug = "Y" Then mydebuglog.Debug("Publish URL: " & pulink)

        aslink = ""
        If ASN_ID <> "" Then aslink = aslink & "&ACP=" & EnURL(ASN_ID)
        If ASN_KEY <> "" Then aslink = aslink & "&AC2=" & EnURL(ASN_KEY)
        If CALL_ID <> "" Then aslink = aslink & "&CID=" & CALL_ID & "&CIS=CNT"
        If DOMAIN <> "" Then aslink = aslink & "&DOM=" & DOMAIN
        If POPUP_FLG <> "" Then aslink = aslink & "&POP=" & POPUP_FLG
        If PUBLIC_FLG <> "" Then aslink = aslink & "&PUB=" & PUBLIC_FLG
        If Debug = "Y" Then mydebuglog.Debug("Assoc URL: " & aslink)

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Or cmd Is Nothing Then
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
            GoTo CloseOut
        End If

        ' ================================================	
        ' GET PROFILE
        CONTACT_ID = "1-2T"
        SUPER_FLG = "Y"
        TRAINER_ID = ""
        TRAINER_FLG = "Y"
        PART_ID = ""
        PART_FLG = "Y"
        PART_ACC_FLG = "N"
        TRAINER_ACC_FLG = "N"
        REG_AS_EMP_ID = "1-2T"
        DOMAIN = "TIPS"
        DOMAIN_FLG = "N"
        SYSADMIN_FLG = "Y"
        MT_FLG = "N"
        MT_ID = ""
        SVC_TYPE = ""
        REPORTS_FLG = "N"
        EMP_ID = "1-2T"
        SITE_ONLY = "N"
        Select Case SVC_TYPE
            Case "CERTIFICATION MANAGER REG DB"
                TRAINING_FLG = "N"
            Case "PUBLIC ACCESS"
                TRAINING_FLG = "Y"
            Case Else
                TRAINING_FLG = "Y"
        End Select

        ' Perform rights checks
        READ_FLG = "Y"
        EDIT_FLG = "Y"
        DEL_FLG = "Y"
        OUTPUT_FLG = "N"
        ADMIN_ACCESS = "Y"
        EXPORT_ACCESS = "Y"

        ' ================================================	
        ' GET SAVED QUERY
        SqlS = "SELECT TOP 1 FUNCTION_ID, ROWS " &
        "FROM reports.dbo.CM_QUERIES " &
        "WHERE USER_ID='" & REG_NUM & "' AND SESSION_ID='" & pSessID & "' AND TRAN_ID='" & CALL_ID & "'"
        If Debug = "Y" Then mydebuglog.Debug("Get query: " & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        FUNCTION_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        iTotalRecords = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  FUNCTION_ID= " & FUNCTION_ID & ",  iTotalRecords= " & iTotalRecords & vbCrLf)
                    Catch ex2 As Exception
                        errmsg = errmsg & "Error getting query: " & ex2.ToString
                        If Debug = "Y" Then mydebuglog.Debug("Error getting query: " & ex2.ToString & vbCrLf)
                        GoTo CloseOut
                    End Try
                End While
            Else
                errmsg = errmsg & "Error getting query." & vbCrLf
            End If
            dr.Close()
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug("Error getting query: " & ex.ToString & vbCrLf)
            errmsg = errmsg & "Error getting query: " & ex.ToString
        End Try
        If FUNCTION_ID <> "GetContent" Then
            errmsg = errmsg & vbCrLf & "Parameter()s missing"
            GoTo CloseOut
        End If
        EMP_ID = REG_NUM

        ' ================================================	
        ' BUILD CATEGORY CONSTRAINT
        Dim Category_Constraint, Cat_Accessible As String
        Cat_Accessible = ""
        Category_Constraint = "CK.key_id IN ("
        If TRAINER_FLG = "Y" Or DOC_LIB_FLG = "Y" Then
            Cat_Accessible = Cat_Accessible & "#3"
        End If
        If MT_FLG = "Y" Then
            Cat_Accessible = Cat_Accessible & "#5"
        End If
        If PART_FLG = "Y" Then
            Cat_Accessible = Cat_Accessible & "#7"
        End If
        If TRAINING_FLG = "Y" Then
            Cat_Accessible = Cat_Accessible & "#8"
        End If
        If TRAINER_ACC_FLG = "Y" Or TRAINER_FLG = "Y" Then
            Cat_Accessible = Cat_Accessible & "#10"
        End If
        If SITE_ONLY = "Y" Then
            Cat_Accessible = Cat_Accessible & "#12"
        End If
        Cat_Accessible = Cat_Accessible & "#13"
        If SYSADMIN_FLG = "Y" Then
            Cat_Accessible = Cat_Accessible & "#15"
        End If
        If EMP_ID <> "" Then
            Cat_Accessible = Cat_Accessible & "#16"
        End If
        Cat_Accessible = Cat_Accessible & "#14"
        If Debug = "Y" Then mydebuglog.Debug("  > Cat_Accessible= " & Cat_Accessible & vbCrLf)

        ' ================================================	
        ' RETRIEVE DATA SET
        ' Compute window
        If Val(sStart) = 0 Then sStart = "1"
        sEnd = Str(Val(sStart) + Val(sAmount)) - 1
        If Debug = "Y" Then mydebuglog.Debug("  > sEnd= " & sEnd & vbCrLf)

        ' Generate column list
        Dim SortCol(20) As String

        ' Compute order	
        'If CAT_FLG = "Y" And CAT_ID = "" Then
        SortCol(NoCols + 1) = "category"
        NoCols = NoCols + 1
        'End If
        SortCol(NoCols + 1) = "name"
        SortCol(NoCols + 2) = "created"
        SortCol(NoCols + 3) = "description"
        NoCols = NoCols + 3

        ' Compute order	
        If (sCol + 1) > NoCols Then sCol = 0
        OrderBy = " ORDER BY " & SortCol(sCol + 1) & " " & sDir

        ' String search clause
        If sSearch <> "" Then
            sSearch = LCase(sSearch)
            SearchClause = "WHERE lower(category) LIKE '%" & sSearch & "%' OR " &
            "lower(name) LIKE '%" & sSearch & "%' OR " &
            "convert(VARCHAR,created,101) LIKE '%" & sSearch & "%' OR " &
            "lower(description) LIKE '%" & sSearch & "%'"
        End If

        ' Get filtered records and verify transaction
        FilterCount = 0
        sc = 0
ReQuery:
        SqlS = "SELECT COUNT(*), CM_TRAN_ID FROM CM.dbo.[" & REG_NUM & "-" & pSessID & "] " & SearchClause & " GROUP BY CM_TRAN_ID"
        If Debug = "Y" Then mydebuglog.Debug("Get Filtered Record Count: " & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        FilterCount = CheckDBNull(dr(0), enumObjectType.IntType)
                        CM_TRAN_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        'If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  FilterCount=" & FilterCount.ToString & ",  CM_TRAN_ID=" & CM_TRAN_ID)
                    Catch ex As Exception
                        errmsg = errmsg & "Error getting query: " & ex.ToString
                        If Debug = "Y" Then mydebuglog.Debug("Error getting Filtered Record Count: " & ex.ToString)
                        GoTo CloseOut
                    End Try
                End While
            Else
                errmsg = errmsg & "Error getting Filtered Record Count." & vbCrLf
            End If
            dr.Close()
        Catch ex As Exception
            errmsg = errmsg & "Error getting Filtered Record Count: " & ex.ToString
            If Debug = "Y" Then mydebuglog.Debug("Error getting Filtered Record count: " & ex.ToString & vbCrLf)
        End Try
        If Debug = "Y" Then
            mydebuglog.Debug("  > FilterCount= " & FilterCount.ToString)
            mydebuglog.Debug("  > CM_TRAN_ID= " & CM_TRAN_ID)
        End If

        ' Recompute start if necessary
        If FilterCount < sStart Then
            sStart = "1"
            sEnd = Str(Val(sStart) + Val(sAmount)) - 1
            If Debug = "Y" Then
                mydebuglog.Debug("  > New sStart= " & sStart)
                mydebuglog.Debug("  > New sEnd= " & sEnd)
            End If
        End If

        ' Verify that the query is current
        If CM_TRAN_ID <> CALL_ID And CM_TRAN_ID <> "" Then
            Dim EmpQueryDoc As XmlElement = DMSServices.EmpQuery(EMP_ID, CONTACT_ID, "GetContent", CALL_ID, SqlS, "Y", Refresh, Debug)
            If EmpQueryDoc.HasAttribute("tablename") Then TableName = EmpQueryDoc.GetAttribute("tablename").ToString
            If EmpQueryDoc.HasAttribute("records") Then TotalDocs = Val(EmpQueryDoc.GetAttribute("records").ToString)
            sc = sc + 1
            If Debug = "Y" Then
                mydebuglog.Debug("  > Requery Count= " & sc.ToString)
                mydebuglog.Debug("  > CMQuery output= " & TableName)
            End If
            If sc < 3 Then GoTo ReQuery
            GoTo DBError
        End If

        ' Compute data query
        Dim iTotalDisplayRecords As String = ""
        If Debug = "Y" Then
            mydebuglog.Debug("  > Order By Clause= " & OrderBy)
            mydebuglog.Debug("  > Search Clause= " & SearchClause)
        End If
        SqlS = "SELECT rownum, row_id, name, description, type, cast(category as varchar), cast(cat_id as varchar), cast(cat_pr_flag as varchar), created, last_upd, cast(type_id as varchar), cast(access_type as varchar), cast(key_id as varchar), dfilename, CM_TRAN_ID " &
        "FROM (SELECT ROW_NUMBER() OVER(" & Trim(OrderBy) & ") AS " &
        "rownum, * FROM CM.dbo.[" & REG_NUM & "-" & pSessID & "] " & Trim(SearchClause) & ") AS Records " &
        "WHERE rownum>=" & sStart & " AND rownum<=" & sEnd
        If Debug = "Y" Then
            mydebuglog.Debug("  > Data Query= " & SqlS)
            mydebuglog.Debug("  > Filter Records found= " & iTotalDisplayRecords & vbCrLf)
        End If

        ' ================================================
        ' RETRIEVE AND ENRICH RECORDS, STORE TO A LIST        
        Dim DocRecords(Val(FilterCount)) As DocsRecord
        Dim DocCategories(Val(FilterCount)) As DocCategory
        Dim splitcat, splitcatid, keys As String()
        Dim KeyFound As Boolean = False
        Dim DocId As String = ""
        Dim KeyId As String = ""
        Dim temp1 As String = ""
        Dim temp2 As String = ""
        Dim temp3 As String = ""
        Dim temp4 As String = ""
        Dim temp5 As String = ""
        Dim temp6 As String = ""
        Dim ctr As Integer = 0
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim k As Integer = 0

        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                Try
                    While dr.Read()
                        i = i + 1
                        If Debug = "Y" Then mydebuglog.Debug("  > i: " & i.ToString)
                        TotalRecs = CheckDBNull(dr(0), enumObjectType.IntType)
                        DocId = CheckDBNull(dr(1), enumObjectType.StrType)
                        KeyId = CheckDBNull(dr(12), enumObjectType.StrType)
                        If Debug = "Y" Then mydebuglog.Debug("  > Records found= " & TotalRecs.ToString & ", DocId= " & DocId & ", KeyId= " & KeyId)
                        KeyFound = False

                        ' Check to make sure that we can access records with this keyword
                        If DocId = "0" Then GoTo DBError
                        If InStr(KeyId, ",") > 0 Then
                            keys = Split(KeyId, ",")
                            For k = 0 To UBound(keys)
                                If Debug = "Y" Then mydebuglog.Debug(" ..... keys= " & keys(k) & ", KeyFound= " & KeyFound.ToString)
                                If InStr(Cat_Accessible, "#" & keys(k)) > 0 Then KeyFound = True
                            Next
                        Else
                            If InStr(Cat_Accessible, "#" & KeyId) > 0 Then KeyFound = True
                            If Debug = "Y" Then mydebuglog.Debug(" ..... KeyId= " & KeyId & ", KeyFound= " & KeyFound.ToString)
                        End If

                        If KeyFound Then
                            ' If new record, save
                            If DocId <> temp1 Then
                                DocRecords(i).DocId = DocId
                                DocRecords(i).DocName = SimpleString(CheckDBNull(dr(2), enumObjectType.StrType))
                                DocRecords(i).DocDesc = SimpleString(CheckDBNull(dr(3), enumObjectType.StrType))
                                DocRecords(i).DocType = CheckDBNull(dr(4), enumObjectType.StrType)
                                DocRecords(i).DocCreated = Format(CheckDBNull(dr(8), enumObjectType.DteType))
                                DocRecords(i).DocUpdated = Format(CheckDBNull(dr(9), enumObjectType.DteType))
                                DocRecords(i).DocRights = CheckDBNull(dr(11), enumObjectType.StrType)
                                DocRecords(i).DocFileName = CheckDBNull(dr(13), enumObjectType.StrType)
                                If Debug = "Y" Then mydebuglog.Debug(" ..... DocName= " & DocRecords(i).DocName &
                                        ", DocFileName= " & DocRecords(i).DocFileName & ", DocType= " & DocRecords(i).DocType &
                                        ", DocCreated= " & DocRecords(i).DocCreated & ", DocUpdated= " & DocRecords(i).DocUpdated)
                                splitcat = Split(Trim(CheckDBNull(dr(5), enumObjectType.StrType)), ",")
                                splitcatid = Split(Trim(CheckDBNull(dr(6), enumObjectType.StrType)), ",")
                                For k = 0 To UBound(splitcat)
                                    If Debug = "Y" Then mydebuglog.Debug(" ..... splitcat= " & splitcat(k) & ", splitcatid= " & splitcatid(k))
                                    If DocCategories.Length = j Then ReDim DocCategories(j + 1)
                                    DocCategories(j).DocId = DocId
                                    DocCategories(j).DocCat = splitcat(k)
                                    DocCategories(j).DocCatId = splitcatid(k)
                                    j = j + 1
                                Next
                            Else
                                splitcat = Split(Trim(CheckDBNull(dr(5), enumObjectType.StrType)), ",")
                                splitcatid = Split(Trim(CheckDBNull(dr(6), enumObjectType.StrType)), ",")
                                If Debug = "Y" Then mydebuglog.Debug(" ..... splitcat= " & splitcat(k) & ", splitcatid= " & splitcatid(k) & ", temp3= " & temp3)
                                If Trim(splitcat(0)) <> temp3 Then
                                    For k = 0 To UBound(splitcat)
                                        DocCategories(j).DocId = DocId
                                        DocCategories(j).DocCat = splitcat(k)
                                        DocCategories(j).DocCatId = splitcatid(k)
                                        j = j + 1
                                    Next
                                End If
                            End If
                            temp1 = DocId
                            temp3 = splitcatid(0)
                        End If
                    End While
                Catch ex2 As Exception
                    errmsg = errmsg & "Error getting Filtered Records: " & ex2.ToString
                    If Debug = "Y" Then mydebuglog.Debug("Error getting Filtered Records: " & ex2.ToString & vbCrLf)
                    GoTo CloseOut
                End Try
            Else
                errmsg = errmsg & "Error getting Filtered Record Count." & vbCrLf
            End If
            dr.Close()
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug("Error getting Filtered Records: " & ex.ToString & vbCrLf)
            errmsg = errmsg & "Error getting Filtered Records: " & ex.ToString
            GoTo CloseOut
        End Try
        If sSearch <> "" Then
            iTotalDisplayRecords = Str(FilterCount)
        Else
            iTotalDisplayRecords = Str(iTotalRecords)
        End If
        TotalRecs = i.ToString
        If Debug = "Y" Then
            mydebuglog.Debug("TotalRecs: " & TotalRecs.ToString & ", Categories= " & Str(j) & vbCrLf)
        End If

        ' If no records found
        If TotalRecs = 0 Then
            jdoc = "{""sEcho"":" & Trim(sEcho) & ","
            jdoc = jdoc & """iTotalRecords"":" & Trim(iTotalRecords) & ","
            jdoc = jdoc & """iTotalDisplayRecords"":" & Trim(iTotalDisplayRecords) & ","
            jdoc = jdoc & """aaData"":[] "
            jdoc = jdoc & "}"
            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "JDOC: " & vbCrLf & jdoc & vbCrLf)
            End If
            GoTo CloseOut
        End If
        If j = 0 Then
            errmsg = "Error reading document categories. " & vbCrLf
            GoTo CloseOut
        End If

        ' ============================================
        ' Prepare JSON document
PrepareResults:
        results = "Success"
        ' 	Header
        jdoc = "{""sEcho"":" & Trim(sEcho) & ","
        jdoc = jdoc & """iTotalRecords"":" & Trim(iTotalRecords) & ","
        jdoc = jdoc & """iTotalDisplayRecords"":" & Trim(iTotalDisplayRecords) & ","
        jdoc = jdoc & """aaData"":[ "
        For i = 1 To TotalRecs
            ' Get document
            DOC_ID = DocRecords(i).DocId
            DOC_TITLE = DocRecords(i).DocName
            DOC_DESC = DocRecords(i).DocDesc
            DOC_DATE = DocRecords(i).DocCreated
            DOC_TYPE = DocRecords(i).DocType
            DOC_RIGHTS = DocRecords(i).DocRights
            DOC_FILENAME = DocRecords(i).DocFileName

            ' Get related categor(ies) for display
            DOC_CATEGORY = ""
            'If CAT_FLG = "Y" And CAT_ID = "" And (CAT_NAME = "" Or CAT_NAME = "any") And j > 0 Then
            If j > 0 Then
                For k = 0 To j
                    If DocCategories.Length = k Then ReDim DocCategories(k + 1)
                    temp1 = DocCategories(k).DocId
                    temp2 = DocCategories(k).DocCat
                    If temp1 = DOC_ID And temp2 <> "" Then
                        ctr = ctr + 1
                        DOC_CATEGORY = DOC_CATEGORY & temp2 & "<br>"
                        If Debug = "Y" Then mydebuglog.Debug(" ..... category= " & temp1 & "/" & temp2)
                    End If
                Next
                If Len(DOC_CATEGORY) > 0 Then DOC_CATEGORY = Left(DOC_CATEGORY, Len(DOC_CATEGORY) - 4)
            End If

            'Create action links based on rights
            editlink = ""
            attachlink = ""
            dellink = ""
            restorelink = ""
            publishlink = ""
            assoclink = ""
            emaillink = ""
            If EDIT_FLG = "Y" Then
                editlink = "<a href=" & BasePath & "/OpenContent?OpenAgent" & edlink & "&EID=" & EMP_ID & "&ID=" & DOC_ID & " title=\u0027Edit document information\u0027>"
            End If
            If DEL_FLG = "Y" Then
                If DLT_FLG <> "Y" Then dellink = "<a href=" & BasePath & "/DelContent?OpenAgent&OI=" & CALL_ID & "&EID=" & EMP_ID & "&CIS=CNT&ID=" & DOC_ID & "&POP=" & POPUP_FLG & " title=\u0027Delete the document\u0027>"
                If ADMIN_ACCESS = "Y" And DLT_FLG = "Y" Then
                    restorelink = "<a href=" & BasePath & "/RestoreContent?OpenAgent&OI=" & CALL_ID & "&EID=" & EMP_ID & "&CIS=CNT&ID=" & DOC_ID & "&POP=" & POPUP_FLG & " title=\u0027Restore the document\u0027>"
                End If
            End If

            ' Attachment link
            If DOC_TYPE = "BitTorrent File" Then
                temp3 = DOC_TITLE
                temp3 = Replace(temp3, "<B><FONT COLOR=""RED"">NEW</FONT></B> ", "")
                temp3 = Replace(temp3, "<B><FONT COLOR=RED>NEW</FONT></B> ", "")
                temp3 = Replace(temp3, " ", "%20")
                attachlink = "<a href=\u0027https://tipsuserdownloads.s3.amazonaws.com/" & temp3 & "\u0027 target=\u0027ClassFrame\u0027 title=\u0027Open the document\u0027>"
            Else
                If InStr(DOC_FILENAME, "getti.ps") > 0 Then
                    attachlink = "<a href=JavaScript:openNewWindow(\u0027" & DOC_FILENAME & "\u0027,800,600) title=\u0027Open the document\u0027>"
                    'attachlink = "<a href=\u0027" & DOC_FILENAME & "\u0027 title=\u0027Open the document\u0027>"
                Else
                    'attachlink = "<a href=JavaScript:openNewWindow(\u0027OpenDocument.ashx?Id=" & DOC_ID & "\u0027,800,600) title=\u0027Open the document\u0027>"
                    attachlink = "<a href=\u0027OpenDocument.ashx?Id=" & DOC_ID & "\u0027 target=\u0027ClassFrame\u0027 title=\u0027Open the document\u0027  onclick=\u0027$(\u0022#Class\u0022).show();$(\u0022#ClassFrame3\u0022).show();\u0027>"
                End If
            End If

            ' If the user has editing rights to a document, they may publish it if they are a documents administrator
            If ADMIN_ACCESS = "Y" Then
                ' Setup publish link
                publishlink = "&ID=" & DOC_ID
                If OrderID <> "" Then publishlink = publishlink & "&CID=" & OrderID
                If DOMAIN <> "" Then publishlink = publishlink & "&DOM=" & DOMAIN
                If POPUP_FLG <> "" Then publishlink = publishlink & "&POP=" & POPUP_FLG
                If CST_FLG <> "" Then publishlink = publishlink & "&CST=" & CST_FLG
                If PUBLIC_FLG <> "" Then publishlink = publishlink & "&PUB=" & PUBLIC_FLG
                If EMP_ID <> "" Then publishlink = publishlink & "&EID=" & EMP_ID
                If CALL_SCREEN <> "" Then publishlink = publishlink & "&CIS=" & CALL_SCREEN
                If NOT_TYPE <> "" Then publishlink = publishlink & "&NTY=" & NOT_TYPE
                publishlink = "<a href=" & BasePath & "/PublishContent?OpenAgent" & publishlink & " title=\u0027Publish Document\u0027>"
                If Debug = "Y" Then mydebuglog.Debug(" ..... publishlink: " & publishlink)

                ' Setup Association link
                If ASN_OPT = "Y" And ASN_ID <> "" Then
                    assoclink = aslink & "&RID=" & DOC_ID & "&EID=" & EMP_ID
                    assoclink = "<a href=" & BasePath & "/UpdRecs?OpenAgent&TYP=DOC&ACT=ASN" & assoclink & " title=\u0027Associate with " & SimpleString(ASN_NAME) & "\u0027>"
                    If Debug = "Y" Then mydebuglog.Debug(" ..... assoclink: " & assoclink)
                End If

                ' Setup email link
                emaillink = "&ID=" & DOC_ID
                If EMP_ID <> "" Then emaillink = emaillink & "&EID=" & EMP_ID
                If DOMAIN <> "" Then emaillink = emaillink & "&DOM=" & DOMAIN
                If OrderID <> "" Then emaillink = emaillink & "&CID=" & OrderID
                If CALL_SCREEN <> "" Then emaillink = emaillink & "&CIS=" & CALL_SCREEN
                If NOT_TYPE <> "" Then emaillink = emaillink & "&NTY=" & NOT_TYPE
                If ASN_ID <> "" Then emaillink = emaillink & "&ACP=" & EnURL(ASN_ID)
                If ASN_KEY <> "" Then emaillink = emaillink & "&AC2=" & EnURL(ASN_KEY)
                emaillink = "<a href=" & BasePath & "/EmailContent?OpenAgent" & emaillink & " title=\u0027Email Document\u0027>"
                If Debug = "Y" Then mydebuglog.Debug(" ..... emaillink: " & emaillink)
            End If

            ' Output record row
            jdoc = jdoc & "["
            'If CAT_FLG = "Y" And CAT_ID = "" Then
            jdoc = jdoc & """" & SimpleString(DOC_CATEGORY) & ""","
            'End If
            jdoc = jdoc & """" & attachlink & "<b>" & DOC_TITLE & "</b></a>"","
            jdoc = jdoc & """<center>" & DOC_DATE & "</center>"","
            jdoc = jdoc & """<center>" & DOC_ID & "</center>"","
            'If Debug = "Y" Then mydebuglog.Debug(" ..... Rights: " & EDIT_FLG & "/" & DEL_FLG & "/" & ADMIN_ACCESS)
            If (EDIT_FLG = "Y" Or DEL_FLG = "Y" Or ADMIN_ACCESS = "Y") Then
                jdoc = jdoc & """" & DOC_DESC & ""","
                temp4 = ""
                If editlink <> "" Then temp4 = temp4 & editlink & "Edit</a>&nbsp;"
                If dellink <> "" Then temp4 = temp4 & dellink & "Delete</a>&nbsp;"
                If restorelink <> "" Then temp4 = temp4 & restorelink & "Restore</a>&nbsp;<br>"
                If publishlink <> "" Then temp4 = temp4 & publishlink & "Publish</a>&nbsp;"
                If assoclink <> "" Then temp4 = temp4 & "<br>" & assoclink & "Associate</a>&nbsp"
                If emaillink <> "" Then temp4 = temp4 & "<br>" & emaillink & "Email</a>"
                If temp4 <> "" Then
                    jdoc = jdoc & """<center>" & temp4 & "</center>"""
                Else
                    jdoc = jdoc & """"""
                End If
            Else
                jdoc = jdoc & """" & SimpleString(DOC_DESC) & """"
            End If
            jdoc = jdoc & "], "
            sc = sc + 1
        Next
        jdoc = Left(jdoc, Len(jdoc) - 2) & " "

        ' -----
        '	Close object
        jdoc = jdoc & "] }"
        jdoc = Replace(jdoc, "[TOTAL]", Trim(Str(sc)))
        If Debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "JDOC: " & vbCrLf & jdoc & vbCrLf)
        End If
        GoTo CloseOut

DBError:
        ErrLvl = "Error"
        errmsg = "There has been a database access error or no records found.  Please try again later."

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Finalize output
        'jdoc = callback & "({""ResultSet"": " & jdoc & ", "
        'jdoc = jdoc & """ErrMsg"":""" & errmsg & ""","
        'jdoc = jdoc & """Results"":""" & results & """ })"
        Try
            Dim json As JObject = JObject.Parse(jdoc)
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug("Error with JSON generated: " & ex.ToString & vbCrLf)
            errmsg = errmsg & "Error with JSON generated: " & ex.ToString
        End Try

        ' ============================================
        ' Close the log file if any
        Dim mtemp As String
        mtemp = "GetContentBE.ashx : Results: " & results & " for user id: " & EMP_ID & "  and order id: " & OrderID
        If Trim(errmsg) <> "" Then mtemp = mtemp & ", Error: " & Trim(errmsg)
        myeventlog.Info(mtemp)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                'mydebuglog.Debug("  JDOC: " & jdoc & vbCrLf)
                mydebuglog.Debug("Results: " & results & " for user id: " & EMP_ID & "  and order id: " & OrderID)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "GETCONTENTBE", LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' Send results        
        If Debug = "T" Then
            context.Response.ContentType = "text/html"
            If jdoc <> "" Then
                context.Response.Write("Success")
            Else
                context.Response.Write("Failure")
            End If
            'context.Response.Write("<h3><b>UserId:</b> " & UserId & "<br>")
            'context.Response.Write("<b>RegId:</b> " & RegId & "</h3>")
            'context.Response.Write("<br>JSON: " & jdoc)
        Else
            If jdoc = "" Then jdoc = errmsg
            'context.Response.ContentType = "application/json"
            context.Response.ContentType = "text/javascript"
            context.Response.ContentEncoding = Encoding.UTF8
            context.Response.Write(jdoc)
        End If
    End Sub

    ' =================================================d
    ' JSON FUNCTIONS
    Function DataSetToJSON(ByVal ds As DataSet) As String

        Dim json As String
        Dim dt As DataTable = ds.Tables(0)
        json = Newtonsoft.Json.JsonConvert.SerializeObject(dt)
        Return json

    End Function

    Function EscapeJSON(ByVal todo As String) As String
        If todo = "" Then
            EscapeJSON = ""
            Exit Function
        End If
        todo = Replace(todo, "\", "\\")
        todo = Replace(todo, "/", "\/")
        todo = Replace(todo, """", "\""")
        todo = Replace(todo, Chr(13), "<br>")
        todo = Replace(todo, Chr(10), "<br>")
        todo = Replace(todo, "   ", " ")
        EscapeJSON = todo
    End Function

    ' =================================================
    ' STRING FUNCTIONS
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

    Function FilterString(ByVal Instring As String) As String
        ' Remove any characters not within the ASCII 31-127 range
        Dim temp As String
        Dim outstring As String
        Dim i, j As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            FilterString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            j = Asc(Mid(temp, i, 1))
            If j > 30 And j < 128 Then
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        FilterString = outstring
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

    Public Function CheckDBNull(ByVal obj As Object,
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
        Dim t As Char
        s = str.ToCharArray
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            If Asc(s(i)) < 128 And Asc(s(i)) > 0 Then
                Try
                    b(i) = Convert.ToByte(s(i))
                Catch ex As Exception
                    b(i) = Convert.ToByte(Chr(32))
                End Try
            Else
                ' Filter out extended ASCII - convert common symbols when possible
                t = Chr(32)
                Try
                    Select Case Asc(s(i))
                        Case 147
                            t = Chr(34)
                        Case 148
                            t = Chr(34)
                        Case 145
                            t = Chr(39)
                        Case 146
                            t = Chr(39)
                        Case 150
                            t = Chr(45)
                        Case 151
                            t = Chr(45)
                        Case Else
                            t = Chr(32)
                    End Select
                Catch ex As Exception
                End Try
                b(i) = Convert.ToByte(t)
            End If
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

    Public Function SimpleString(sRawURL As String) As String
        Dim iLoop As Integer = 0
        Dim sRtn As String = ""
        Dim sTmp As String = ""
        Const sValidChars = " 1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:/.?=_-$(){}~<>"
        Try
            If Len(sRawURL) > 0 Then
                ' Loop through each char
                For iLoop = 1 To Len(sRawURL)
                    sTmp = Mid(sRawURL, iLoop, 1)
                    If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                        ' If not ValidChar, then remove
                    Else
                        sRtn = sRtn & sTmp
                    End If
                Next iLoop
            End If
            SimpleString = sRtn
        Catch ex As Exception
            SimpleString = ""
        End Try
        Return sRtn
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

    ' =================================================
    ' DEBUG FUNCTIONS
    Public Sub writeoutputfs(ByVal fs As FileStream, ByVal instring As String)
        ' This function writes a line to a previously opened filestream, and then flushes it
        ' promptly.  This assists in debugging services
        fs.Write(StringToBytes(instring), 0, Len(instring))
        fs.Write(StringToBytes(vbCrLf), 0, 2)
        fs.Flush()
    End Sub

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class