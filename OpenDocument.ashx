<%@ WebHandler Language="VB" Class="OpenDocument" %>

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
Imports System.Security.Cryptography.X509Certificates
Imports log4net

Public Class OpenDocument : Implements IHttpHandler

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

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        ' Declare variables
        Dim ErrMsg, ErrLvl As String
        Dim PublicKey As String
        Dim Debug As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Doucment declarations
        Dim d_dsize, minio_flg As String
        Dim ItemName, Domain, DocumentId, VersionId As String
        Dim Extension, LANG_CD As String
        Dim d_last_updated, v_last_updated As DateTime
        Dim retval As Long
        Dim buffer As Byte() = New Byte(4095) {}

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
        Dim PrevLink As String = Trim(context.Request.ServerVariables("HTTP_REFERER"))
        Dim BROWSER As String = Trim(context.Request.ServerVariables("HTTP_USER_AGENT"))

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' ============================================
        ' Variable setup
        Debug = "Y"
        Logging = "Y"
        Domain = ""
        PublicKey = ""
        ItemName = ""
        Extension = ""
        BROWSER = ""
        Debug = "N"
        ErrMsg = ""
        VersionId = ""
        DocumentId = ""
        LANG_CD = "ENU"
        ErrLvl = "Warning"
        minio_flg = "N"
        d_dsize = 0

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("OpenDocument_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
            AccessKey = System.Configuration.ConfigurationManager.AppSettings("minio-key")
            If AccessKey = "" Then AccessKey = "dms"
            AccessSecret = System.Configuration.ConfigurationManager.AppSettings("minio-secret")
            If AccessSecret <> "" Then AccessSecret = System.Web.HttpUtility.HtmlDecode(AccessSecret)
            If AccessSecret = "" Then AccessSecret = ""
            AccessRegion = System.Configuration.ConfigurationManager.AppSettings("minio-region")
            If AccessRegion = "" Then AccessRegion = ""
            AccessBucket = System.Configuration.ConfigurationManager.AppSettings("minio-bucket")
            If AccessBucket = "" Then AccessBucket = ""
        Catch ex As Exception
            ErrMsg = ErrMsg & vbCrLf & "Unable to get defaults from web.config: " & ex.Message
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\OpenDocument.log"
            Try
                log4net.GlobalContext.Properties("PDDLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                ErrMsg = ErrMsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If

        ' ============================================
        ' Get parameters
        If Not context.Request.QueryString("Id") Is Nothing Then
            DocumentId = context.Request.QueryString("Id")
        End If
        If Not context.Request.QueryString("VId") Is Nothing Then
            VersionId = context.Request.QueryString("VId")
        End If
        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If
        If DocumentId = "" And VersionId = "" Then
            ErrLvl = "Error"
            Select Case LANG_CD
                Case "ESN"
                    ErrMsg = "Solicitud no v&aacute;lida. No se ha especificado ning&uacute;n ID de documento o versi&oacute;n."
                Case Else
                    ErrMsg = "Invalid request. No document or version id specified."
            End Select
            GoTo DisplayErrorMsg
        End If

        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  Debug: " & Debug)
            mydebuglog.Debug("  DocumentId: " & DocumentId)
            mydebuglog.Debug("  VersionId: " & VersionId)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  DOMAIN: " & Domain)
            mydebuglog.Debug("  BROWSER: " & BROWSER & vbCrLf)
        End If

        ' ============================================
        ' Open database connection 
        ErrMsg = OpenDBConnection(ConnS, con, cmd)
        If ErrMsg <> "" Then
            GoTo DBError
        End If

        ' ============================================
        ' Process Request
        If Not cmd Is Nothing Then
            If DocumentId <> "" Then
                SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd, V.dsize, V.row_id, V.last_upd, V.minio_flg, DT.extension, V.dimage " &
                "FROM DMS.dbo.Documents D " &
                "LEFT OUTER JOIN DMS.dbo.Document_Types DT on D.data_type_id=DT.row_id " &
                "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id " &
                "WHERE D.row_id=" & DocumentId & " " &
                "ORDER BY D.last_upd DESC"
            Else
                SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd, V.dsize, V.row_id, V.last_upd, V.minio_flg, DT.extension, V.dimage " &
                "FROM DMS.dbo.Documents D " &
                "LEFT OUTER JOIN DMS.dbo.Document_Types DT on D.data_type_id=DT.row_id " &
                "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id " &
                "WHERE V.row_id=" & VersionId & " " &
                "ORDER BY D.last_upd DESC"
            End If
            If Debug = "Y" Then mydebuglog.Debug("  Get document information : " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            ItemName = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            DocumentId = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            d_last_updated = CheckDBNull(dr(2), enumObjectType.DteType)
                            d_dsize = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            VersionId = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            v_last_updated = CheckDBNull(dr(5), enumObjectType.DteType)
                            minio_flg = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                            Extension = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                            If Extension = "" And InStr(ItemName, ".") > 0 Then
                                Extension = LCase(Right(ItemName, 3))
                                If Extension = "tml" Then Extension = "html"
                                If Extension = "peg" Then Extension = "jpeg"
                                If Extension = "lsx" Then Extension = "xlsx"
                            End If
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  DocumentId=" & DocumentId & ",  ItemName=" & ItemName & ",  Extension=" & Extension & ",  VersionId=" & VersionId & ",  minio_flg=" & minio_flg & ", d_last_updated=" & d_last_updated.ToString & vbCrLf)
                            If minio_flg <> "Y" Then
                                ' Get binary from document_versions
                                If Debug = "Y" Then mydebuglog.Debug("  Getting binary from Document_Versions")
                                ReDim buffer(Val(d_dsize) - 1)
                                Try
                                    retval = dr.GetBytes(8, 0, buffer, 0, d_dsize)
                                Catch ex2 As Exception
                                    ErrMsg = ErrMsg & "Error getting item. " & ex2.ToString & vbCrLf
                                    GoTo CloseOut
                                End Try
                            End If
                        Catch ex As Exception
                            If Debug = "Y" Then mydebuglog.Debug("Error getting document: " & ex.ToString & vbCrLf)
                            GoTo DBError
                        End Try
                    End While
                Else
                    ErrMsg = ErrMsg & "Error getting document." & vbCrLf
                End If

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
                        Dim mobj2 = Minio.GetObject(AccessBucket, DocumentId & "-" & VersionId)
                        retval = mobj2.ContentLength
                        If retval > 0 Then
                            ReDim buffer(Val(retval))
                            Dim intval As Integer
                            For i = 0 To retval
                                intval = mobj2.ResponseStream.ReadByte()
                                If intval < 255 And intval > 0 Then
                                    buffer(i) = intval
                                End If
                                If intval = 255 Then buffer(i) = 255
                                If intval < 0 Then
                                    buffer(i) = 0
                                    If Debug = "Y" Then mydebuglog.Debug("   .. read " & i.ToString)
                                End If
                            Next
                        End If
                        mobj2 = Nothing
                    Catch ex2 As Exception
                        ErrMsg = ErrMsg & "Error getting object. " & ex2.ToString & vbCrLf
                        GoTo CloseOut
                    End Try

                    Try
                        Minio = Nothing
                    Catch ex2 As Exception
                        ErrMsg = ErrMsg & "Error closing Minio: " & ex2.Message & vbCrLf
                    End Try
                End If
                dr.Close()
            Catch ex As Exception
                ErrMsg = ErrMsg & "Error closing Minio: " & ex.Message & vbCrLf
            End Try
        End If

ReturnControl:
        GoTo CloseOut

DBError:
        If Debug = "Y" Then mydebuglog.Debug(">>DBError")
        ErrLvl = "Error"
        Select Case LANG_CD
            Case "ESN"
                ErrMsg = "El sistema puede no estar disponible ahora. Por favor, int&eacute;ntelo de nuevo m&aacute;s tarde"
            Case Else
                ErrMsg = "The system may be unavailable now.  Please try again later"
        End Select
        GoTo CloseOut

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
            ErrMsg = ErrMsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(ErrMsg) <> "" Then myeventlog.Error("OpenDocument.ashx: " & ErrLvl & ": " & Trim(ErrMsg))
        myeventlog.Info("OpenDocument.ashx : DocumentId : " & DocumentId & ", VersionId: " & VersionId)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(ErrMsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(ErrMsg))
                mydebuglog.Debug(vbCrLf & "Results:  DocumentId : " & DocumentId & ", VersionId: " & VersionId)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "OpenDocument", LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If
        If ErrMsg <> "" Then GoTo DisplayErrorMsg

        ' ============================================
        ' WRITE DOCUMENT
        ' Write output headers
        context.Response.Clear()
        context.Response.AddHeader("Content-Disposition", "inline; extension-token=" & Extension)
        context.Response.ContentType = GetContentType(Extension)

        ' Write stream
        Try
            'If Debug = "Y" Then mydebuglog.Debug("   .. buffer: " & buffer.Length.ToString)
            If buffer.Length = 0 Then
                ErrMsg = ErrMsg & "No data buffered: " & vbCrLf
                GoTo DisplayErrorMsg
            End If
            Dim bufferlen As Integer = 4096
            If buffer.Length < 4096 Then bufferlen = buffer.Length
            Dim strm = New MemoryStream(buffer)
            Dim byteSeq As Integer = strm.Read(buffer, 0, bufferlen)
            'If Debug = "Y" Then mydebuglog.Debug("   .. byteSeq: " & byteSeq.ToString)
            Do While byteSeq > 0
                context.Response.OutputStream.Write(buffer, 0, byteSeq)
                byteSeq = strm.Read(buffer, 0, bufferlen)
            Loop
            context.Response.OutputStream.Close()
            buffer = Nothing
            strm = Nothing
        Catch ex As Exception
            ErrMsg = ErrMsg & "Error writing results: " & ex.Message & vbCrLf
            GoTo DisplayErrorMsg
        End Try

        ' Close everthing
        Exit Sub

DisplayErrorMsg:
        context.Response.ContentType = "text/html"
        context.Response.Write("<h2><b>" & ErrMsg & "</b></h2>")
        If Debug = "Y" Then
            Using writer As New StreamWriter("C:\Logs\OpenDocument_failed.log", True)
                writer.WriteLine(Now.ToString & " - " & ErrMsg & ", ItemName: " & ItemName & ",  PublicKey: " & PublicKey & ", Domain: " & Domain & ", GetContentType: " & GetContentType(Extension) & ", BROWSER: " & BROWSER)
            End Using
        End If
    End Sub

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

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
    ' =================================================
    ' Function to translate file extension into MIME type
    Private Function GetContentType(ByVal Extension As String) As String
        If InStr(Extension, ".") > 0 Then
            Extension = Replace(Extension, ".", "")
        End If
        Select Case Extension
            Case "html"
                GetContentType = "text/html"
            Case "ai"
                GetContentType = "application/postscript"
            Case "aif"
                GetContentType = "audio/x-aiff"
            Case "aifc"
                GetContentType = "audio/x-aiff"
            Case "aiff"
                GetContentType = "audio/x-aiff"
            Case "asf"
                GetContentType = "video/x-ms-asf"
            Case "asr"
                GetContentType = "video/x-ms-asf"
            Case "asx"
                GetContentType = "video/x-ms-asf"
            Case "au"
                GetContentType = "audio/basic"
            Case "avi"
                GetContentType = "video/x-msvideo"
            Case "axs"
                GetContentType = "application/olescript"
            Case "bas"
                GetContentType = "text/plain"
            Case "bcpio"
                GetContentType = "application/x-bcpio"
            Case "bin"
                GetContentType = "application/octet-stream"
            Case "bmp"
                GetContentType = "image/bmp"
            Case "c"
                GetContentType = "text/plain"
            Case "cat"
                GetContentType = "application/vnd.ms-pkiseccat"
            Case "cdf"
                GetContentType = "application/x-cdf"
            Case "cer"
                GetContentType = "application/x-x509-ca-cert"
            Case "class"
                GetContentType = "application/octet-stream"
            Case "clp"
                GetContentType = "application/x-msclip"
            Case "cmx"
                GetContentType = "image/x-cmx"
            Case "cod"
                GetContentType = "image/cis-cod"
            Case "cpio"
                GetContentType = "application/x-cpio"
            Case "crd"
                GetContentType = "application/x-mscardfile"
            Case "crl"
                GetContentType = "application/pkix-crl"
            Case "crt"
                GetContentType = "application/x-x509-ca-cert"
            Case "csh"
                GetContentType = "application/x-csh"
            Case "css"
                GetContentType = "text/css"
            Case "dcr"
                GetContentType = "application/x-director"
            Case "der"
                GetContentType = "application/x-x509-ca-cert"
            Case "dir"
                GetContentType = "application/x-director"
            Case "dll"
                GetContentType = "application/x-msdownload"
            Case "dms"
                GetContentType = "application/octet-stream"
            Case "doc"
                GetContentType = "application/msword"
            Case "dot"
                GetContentType = "application/msword"
            Case "dvi"
                GetContentType = "application/x-dvi"
            Case "dxr"
                GetContentType = "application/x-director"
            Case "eps"
                GetContentType = "application/postscript"
            Case "etx"
                GetContentType = "text/x-setext"
            Case "evy"
                GetContentType = "application/envoy"
            Case "exe"
                GetContentType = "application/octet-stream"
            Case "fif"
                GetContentType = "application/fractals"
            Case "flr"
                GetContentType = "x-world/x-vrml"
            Case "gif"
                GetContentType = "image/gif"
            Case "gtar"
                GetContentType = "application/x-gtar"
            Case "gz"
                GetContentType = "application/x-gzip"
            Case "h"
                GetContentType = "text/plain"
            Case "hdf"
                GetContentType = "application/x-hdf"
            Case "hlp"
                GetContentType = "application/winhlp"
            Case "hqx"
                GetContentType = "application/mac-binhex40"
            Case "hta"
                GetContentType = "application/hta"
            Case "htc"
                GetContentType = "text/x-component"
            Case "htm"
                GetContentType = "text/html"
            Case "html"
                GetContentType = "text/html"
            Case "htt"
                GetContentType = "text/webviewhtml"
            Case "ico"
                GetContentType = "image/x-icon"
            Case "ief"
                GetContentType = "image/ief"
            Case "iii"
                GetContentType = "application/x-iphone"
            Case "ins"
                GetContentType = "application/x-internet-signup"
            Case "isp"
                GetContentType = "application/x-internet-signup"
            Case "jfif"
                GetContentType = "image/pipeg"
            Case "jpe"
                GetContentType = "image/jpeg"
            Case "jpeg"
                GetContentType = "image/jpeg"
            Case "jpg"
                GetContentType = "image/jpeg"
            Case "js"
                GetContentType = "application/x-javascript"
            Case "latex"
                GetContentType = "application/x-latex"
            Case "lha"
                GetContentType = "application/octet-stream"
            Case "lsf"
                GetContentType = "video/x-la-asf"
            Case "lsx"
                GetContentType = "video/x-la-asf"
            Case "lzh"
                GetContentType = "application/octet-stream"
            Case "m13"
                GetContentType = "application/x-msmediaview"
            Case "m14"
                GetContentType = "application/x-msmediaview"
            Case "m3u"
                GetContentType = "audio/x-mpegurl"
            Case "man"
                GetContentType = "application/x-troff-man"
            Case "mdb"
                GetContentType = "application/x-msaccess"
            Case "me"
                GetContentType = "application/x-troff-me"
            Case "mht"
                GetContentType = "message/rfc822"
            Case "mhtml"
                GetContentType = "message/rfc822"
            Case "mid"
                GetContentType = "audio/mid"
            Case "mny"
                GetContentType = "application/x-msmoney"
            Case "mov"
                GetContentType = "video/quicktime"
            Case "movie"
                GetContentType = "video/x-sgi-movie"
            Case "mp2"
                GetContentType = "video/mpeg"
            Case "mp3"
                GetContentType = "audio/mpeg"
            Case "mpa"
                GetContentType = "video/mpeg"
            Case "mpe"
                GetContentType = "video/mpeg"
            Case "mpeg"
                GetContentType = "video/mpeg"
            Case "mpg"
                GetContentType = "video/mpeg"
            Case "mpp"
                GetContentType = "application/vnd.ms-project"
            Case "mpv2"
                GetContentType = "video/mpeg"
            Case "ms"
                GetContentType = "application/x-troff-ms"
            Case "mvb"
                GetContentType = "application/x-msmediaview"
            Case "nws"
                GetContentType = "message/rfc822"
            Case "oda"
                GetContentType = "application/oda"
            Case "p10"
                GetContentType = "application/pkcs10"
            Case "p12"
                GetContentType = "application/x-pkcs12"
            Case "p7b"
                GetContentType = "application/x-pkcs7-certificates"
            Case "p7c"
                GetContentType = "application/x-pkcs7-mime"
            Case "p7m"
                GetContentType = "application/x-pkcs7-mime"
            Case "p7r"
                GetContentType = "application/x-pkcs7-certreqresp"
            Case "p7s"
                GetContentType = "application/x-pkcs7-signature"
            Case "pbm"
                GetContentType = "image/x-portable-bitmap"
            Case "pdf"
                GetContentType = "application/pdf"
            Case "pfx"
                GetContentType = "application/x-pkcs12"
            Case "pgm"
                GetContentType = "image/x-portable-graymap"
            Case "pko"
                GetContentType = "application/ynd.ms-pkipko"
            Case "pma"
                GetContentType = "application/x-perfmon"
            Case "pmc"
                GetContentType = "application/x-perfmon"
            Case "pml"
                GetContentType = "application/x-perfmon"
            Case "pmr"
                GetContentType = "application/x-perfmon"
            Case "pmw"
                GetContentType = "application/x-perfmon"
            Case "png"
                GetContentType = "image/png"
            Case "pnm"
                GetContentType = "image/x-portable-anymap"
            Case "po"
                GetContentType = "application/vnd.ms-powerpoint"
            Case "ppm"
                GetContentType = "image/x-portable-pixmap"
            Case "pps"
                GetContentType = "application/vnd.ms-powerpoint"
            Case "ppt"
                GetContentType = "application/vnd.ms-powerpoint"
            Case "prf"
                GetContentType = "application/pics-rules"
            Case "ps"
                GetContentType = "application/postscript"
            Case "pub"
                GetContentType = "application/x-mspublisher"
            Case "qt"
                GetContentType = "video/quicktime"
            Case "ra"
                GetContentType = "audio/x-pn-realaudio"
            Case "ram"
                GetContentType = "audio/x-pn-realaudio"
            Case "ras"
                GetContentType = "image/x-cmu-raster"
            Case "rgb"
                GetContentType = "image/x-rgb"
            Case "rmi"
                GetContentType = "audio/mid"
            Case "roff"
                GetContentType = "application/x-troff"
            Case "rtf"
                GetContentType = "application/rtf"
            Case "rtx"
                GetContentType = "text/richtext"
            Case "scd"
                GetContentType = "application/x-msschedule"
            Case "sct"
                GetContentType = "text/scriptlet"
            Case "setpay"
                GetContentType = "application/set-payment-initiation"
            Case "setreg"
                GetContentType = "application/set-registration-initiation"
            Case "sh"
                GetContentType = "application/x-sh"
            Case "shar"
                GetContentType = "application/x-shar"
            Case "sit"
                GetContentType = "application/x-stuffit"
            Case "snd"
                GetContentType = "audio/basic"
            Case "spc"
                GetContentType = "application/x-pkcs7-certificates"
            Case "spl"
                GetContentType = "application/futuresplash"
            Case "src"
                GetContentType = "application/x-wais-source"
            Case "sst"
                GetContentType = "application/vnd.ms-pkicertstore"
            Case "stl"
                GetContentType = "application/vnd.ms-pkistl"
            Case "stm"
                GetContentType = "text/html"
            Case "svg"
                GetContentType = "image/svg+xml"
            Case "sv4cpio"
                GetContentType = "application/x-sv4cpio"
            Case "sv4crc"
                GetContentType = "application/x-sv4crc"
            Case "swf"
                GetContentType = "application/x-shockwave-flash"
            Case "t"
                GetContentType = "application/x-troff"
            Case "tar"
                GetContentType = "application/x-tar"
            Case "tcl"
                GetContentType = "application/x-tcl"
            Case "tex"
                GetContentType = "application/x-tex"
            Case "texi"
                GetContentType = "application/x-texinfo"
            Case "texinfo"
                GetContentType = "application/x-texinfo"
            Case "tgz"
                GetContentType = "application/x-compressed"
            Case "tif"
                GetContentType = "image/tiff"
            Case "tiff"
                GetContentType = "image/tiff"
            Case "tr"
                GetContentType = "application/x-troff"
            Case "trm"
                GetContentType = "application/x-msterminal"
            Case "tsv"
                GetContentType = "text/tab-separated-values"
            Case "txt"
                GetContentType = "text/plain"
            Case "uls"
                GetContentType = "text/iuls"
            Case "ustar"
                GetContentType = "application/x-ustar"
            Case "vcf"
                GetContentType = "text/x-vcard"
            Case "vrml"
                GetContentType = "x-world/x-vrml"
            Case "wav"
                GetContentType = "audio/x-wav"
            Case "wcm"
                GetContentType = "application/vnd.ms-works"
            Case "wdb"
                GetContentType = "application/vnd.ms-works"
            Case "wks"
                GetContentType = "application/vnd.ms-works"
            Case "wmf"
                GetContentType = "application/x-msmetafile"
            Case "wps"
                GetContentType = "application/vnd.ms-works"
            Case "wri"
                GetContentType = "application/x-mswrite"
            Case "wrl"
                GetContentType = "x-world/x-vrml"
            Case "wrz"
                GetContentType = "x-world/x-vrml"
            Case "xaf"
                GetContentType = "x-world/x-vrml"
            Case "xbm"
                GetContentType = "image/x-xbitmap"
            Case "xla"
                GetContentType = "application/vnd.ms-excel"
            Case "xlc"
                GetContentType = "application/vnd.ms-excel"
            Case "xlm"
                GetContentType = "application/vnd.ms-excel"
            Case "xls"
                GetContentType = "application/vnd.ms-excel"
            Case "xlsx"
                GetContentType = "application/vnd.ms-excel"
            Case "xlt"
                GetContentType = "application/vnd.ms-excel"
            Case "xlw"
                GetContentType = "application/vnd.ms-excel"
            Case "xof"
                GetContentType = "x-world/x-vrml"
            Case "xpm"
                GetContentType = "image/x-xpixmap"
            Case "xwd"
                GetContentType = "image/x-xwindowdump"
            Case "z"
                GetContentType = "application/x-compress"
            Case "zip"
                GetContentType = "application/zip"
            Case Else
                GetContentType = "application/octet-stream"
        End Select
    End Function

End Class