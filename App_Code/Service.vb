Option Explicit On
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.Data.SqlClient
Imports System.IO
Imports System.Collections
Imports System.Configuration
Imports System.Net.Mail
Imports System.Math
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
'Imports System.Data.SqlServerCe
Imports System.Runtime.Caching
Imports log4net
Imports System
Imports Amazon
Imports Amazon.S3
Imports System.Net
Imports CachingWrapper.LocalCache

'<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<WebService(Namespace:="http://hciscorm.certegrity.com/svc/")>
<WebServiceBinding(ConformsTo:=WsiProfiles.None)>
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
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

    ' Logging objects
    Private myeventlog As log4net.ILog
    Private mydebuglog As log4net.ILog

    <WebMethod(Description:="Retrieves a selection of KBA questions")>
    Public Function KBALookup(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function provides questions and answers for a Knowledge-Based Authentication 
        ' system used to validate the identity of an online class attendee.

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	        - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim i As Integer
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' KBA declarations
        Dim DecodedUserId, AltUserId As String
        Dim num_questions As Integer
        Dim ques_text(100) As String
        Dim ansr_text(100) As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If RegId = "" And Debug = "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9BFM9"
            UserId = ""
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") = 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\YourServiceName.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try
        If Debug = "Y" Then
            Try
                mydebuglog.Debug(vbCrLf & "Contact In-")
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        If Not cmd Is Nothing Then
            Try
                ' Double check user
                SqlS = "SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT WHERE X_REGISTRATION_NUM='" & DecodedUserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Check to see if the contact id exists: " & SqlS)
                cmd.CommandText = SqlS
                AltUserId = cmd.ExecuteScalar()
                If Debug = "Y" Then mydebuglog.Debug("  > AltUserId: " & AltUserId & vbCrLf)
                If AltUserId = "" Then
                    SqlS = "SELECT X_REGISTRATION_NUM FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & DecodedUserId & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Locate registration number: " & SqlS)
                    cmd.CommandText = SqlS
                    DecodedUserId = cmd.ExecuteScalar()
                    If Debug = "Y" Then mydebuglog.Debug("  > DecodedUserId: " & DecodedUserId & vbCrLf)
                End If

                SqlS = "EXEC reports.dbo.OpenHCIKeys; SELECT Q.QUES_TEXT, reports.dbo.HCI_Decrypt(A.ENC_ANSR_TEXT) AS ANSR_TEXT " &
                    "FROM elearning.dbo.KBA_QUES Q " &
                    "LEFT OUTER JOIN elearning.dbo.KBA_ANSR A ON A.QUES_ID=Q.ROW_ID " &
                    "WHERE A.USER_ID='" & DecodedUserId & "' AND A.REG_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get questions/answers: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            num_questions = num_questions + 1
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ques_text(num_questions) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            ansr_text(num_questions) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            If ques_text(num_questions) = "" Then results = "Failure"
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting questions. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Questions were not found." & vbCrLf
                    dr.Close()
                    results = "Failure"
                End If
                dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' ============================================
        ' Return the KBA challenges and correct answers
        '   <kba num_questions=�#�>
        '       <question answer=�Correct Answer text�>Question text</question>
        '       <question answer=�Correct Answer text�>Question text</question>
        '       ...
        '   </kba>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        Dim resultsItem As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("kba")
        AddXMLAttribute(odoc, resultsRoot, "num_questions", num_questions.ToString)
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        Try
            ' Add result items - send what was submitted for debugging purposes 
            If Debug <> "T" Then
                For i = 1 To num_questions
                    'AddXMLChild(odoc, resultsRoot, "question", ques_text(i))
                    resultsItem = odoc.CreateElement("question")
                    resultsItem.InnerText = ques_text(i)
                    AddXMLAttribute(odoc, resultsItem, "answer", ansr_text(i))
                    resultsRoot.AppendChild(resultsItem)
                Next
            End If
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("KBALookup : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("KBALookup : Results: " & results & " for user id: " & DecodedUserId & "  and reg id:" & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for user id: " & DecodedUserId & "  and reg id:" & RegId)
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
        Return odoc
    End Function

    <WebMethod(Description:="Resets KBA answers entered for a specific user")>
    Public Function KBAReset(ByVal UserId As String, ByVal RegId As String, ByVal Debug As String) As Boolean
        ' This function resets the answers previously entered for a Knowledge-Based Authentication 

        ' The input parameters are as follows:
        '   UserId      - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   RegId	- CX_SESS_REG.ROW_ID. Registration Id
        '   Debug	- "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Output:
        '   True or False, depending on whether the answer supplied is correct or not
        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' KBA declarations
        Dim DecodedUserId, AltUserId As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim myresults As Boolean

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "False"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        myresults = False
        'Debug = "Y"

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (UserId = "" Or RegId = "") And Debug <> "T" Then
            results = "False"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            ' Fix registration id
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            ' Fix user id
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\KBAReset.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                myresults = False
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  Debug: " & Debug)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  RegId: " & RegId)
            End If
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            myresults = False
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            myresults = False
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        If Not cmd Is Nothing Then
            ' Double check user
            Try
                SqlS = "SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT WHERE X_REGISTRATION_NUM='" & DecodedUserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Check to see if the contact id exists: " & SqlS)
                cmd.CommandText = SqlS
                AltUserId = cmd.ExecuteScalar()
                If Debug = "Y" Then mydebuglog.Debug("  > AltUserId: " & AltUserId & vbCrLf)
                If AltUserId = "" Then
                    SqlS = "SELECT X_REGISTRATION_NUM FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & DecodedUserId & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Locate registration number: " & SqlS)
                    cmd.CommandText = SqlS
                    DecodedUserId = cmd.ExecuteScalar()
                    If Debug = "Y" Then mydebuglog.Debug("  > DecodedUserId: " & DecodedUserId & vbCrLf)
                End If
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
                myresults = False
                GoTo CloseOut
            End Try

            ' Reset validation answers
            Try
                SqlS = "UPDATE elearning.dbo.KBA_ANSR " &
                 "SET VAL_TEXT=NULL, ENC_VAL_TEXT=NULL, VAL_DATE=NULL " &
                 "WHERE ROW_ID IN (" &
                 "SELECT A.ROW_ID " &
                 "FROM elearning.dbo.KBA_QUES Q " &
                 "LEFT OUTER JOIN elearning.dbo.KBA_ANSR A ON A.QUES_ID=Q.ROW_ID " &
                 "WHERE A.USER_ID='" & DecodedUserId & "' AND A.REG_ID='" & RegId & "')"
                If Debug = "Y" Then mydebuglog.Debug("  Reset validation answers: " & SqlS)
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                If returnv = 0 Then
                    myresults = False
                Else
                    myresults = True
                End If
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                myresults = False
                GoTo CloseOut
            End Try

        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
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
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("KBAReset :  Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("KBAReset : Results: " & myresults & " for RegId # " & RegId & " and UserId " & DecodedUserId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & myresults & " for RegId # " & RegId & " and UserId " & DecodedUserId & " at " & Now.ToString)
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
        Return myresults
    End Function

    <WebMethod(Description:="Retrieves KBA questions and answers for a specific user")>
    Public Function KBAQLookup(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function provides questions and question ids for a Knowledge-Based Authentication 
        ' system used to validate the identity of an online class attendee.

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	        - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim i As Integer
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile, temp As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' KBA declarations
        Dim DecodedUserId, AltUserId As String
        Dim num_questions As Integer
        Dim ques_text(100) As String
        Dim ques_id(100) As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        num_questions = 0
        temp = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9BFM9"
            UserId = ""
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("KBAQLookup_debug")
            If temp = "Y" And Debug <> "T" And Debug <> "R" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\KBAQLookup.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId & vbCrLf)
            End If
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        If Not cmd Is Nothing Then
            ' Double check user
            Try
                SqlS = "SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT WHERE X_REGISTRATION_NUM='" & DecodedUserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Check to see if the contact id exists: " & SqlS)
                cmd.CommandText = SqlS
                AltUserId = cmd.ExecuteScalar()
                If Debug = "Y" Then mydebuglog.Debug("  > AltUserId: " & AltUserId & vbCrLf)
                If AltUserId = "" Then
                    SqlS = "SELECT X_REGISTRATION_NUM FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & DecodedUserId & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Locate registration number: " & SqlS)
                    cmd.CommandText = SqlS
                    DecodedUserId = cmd.ExecuteScalar()
                    If Debug = "Y" Then mydebuglog.Debug("  > DecodedUserId: " & DecodedUserId & vbCrLf)
                End If
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString & vbCrLf)
            End Try

            ' Get questions
            Try
                SqlS = "SELECT A.ROW_ID, Q.QUES_TEXT, NEWID() " &
                    "FROM elearning.dbo.KBA_ANSR A " &
                    "LEFT OUTER JOIN elearning.dbo.KBA_QUES Q ON Q.ROW_ID=A.QUES_ID " &
                    "WHERE A.REG_ID='" & RegId & "' AND A.USER_ID='" & DecodedUserId & "' AND A.VAL_DATE IS NULL " &
                    "UNION " &
                    "SELECT A.ROW_ID, Q.QUES_TEXT, NEWID() " &
                    "FROM siebeldb.dbo.CX_SESS_REG R " &
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X P ON P.ROW_ID=R.SESS_PART_ID " &
                    "LEFT OUTER JOIN elearning.dbo.KBA_ANSR A ON A.REG_ID=R.ROW_ID " &
                    "LEFT OUTER JOIN elearning.dbo.KBA_QUES Q ON Q.ROW_ID=A.QUES_ID " &
                    "WHERE P.CRSE_TSTRUN_ID='" & RegId & "' AND A.USER_ID='" & DecodedUserId & "' " &
                    "ORDER BY NEWID()"
                If Debug = "Y" Then mydebuglog.Debug("  Get questions: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            num_questions = num_questions + 1
                            ques_id(num_questions) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            ques_text(num_questions) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            If ques_text(num_questions) = "" Then results = "Failure"
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query: " & ques_id(num_questions))
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting questions. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Questions were not found." & vbCrLf
                    dr.Close()
                    results = "Failure"
                End If
                dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' ============================================
        ' Return the KBA challenges and correct answers
        '   <kba num_questions="#">
        '       <question id="Question Id">Question text</question>
        '       <question id="Question Id">Question text</question>
        '       ...
        '   </kba>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        Dim resultsItem As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("kba")
        AddXMLAttribute(odoc, resultsRoot, "num_questions", num_questions.ToString)
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        Try
            ' Add result items - send what was submitted for debugging purposes 
            If Debug <> "T" Then
                For i = 1 To num_questions
                    'AddXMLChild(odoc, resultsRoot, "question", ques_text(i))
                    resultsItem = odoc.CreateElement("question")
                    resultsItem.InnerText = ques_text(i)
                    AddXMLAttribute(odoc, resultsItem, "id", ques_id(i))
                    resultsRoot.AppendChild(resultsItem)
                Next
            End If
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("KBAQLookup : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("KBAQLookup : Results: " & results & " for RegId # " & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " at " & Now.ToString)
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
        Return odoc
    End Function

    <WebMethod(Description:="Validates an answer to a KBA question for a specific user")>
    Public Function KBAALookup(ByVal QuesId As String, ByVal Answer As String, ByVal UserId As String, ByVal Debug As String) As Boolean
        ' This function validates an answer for a Knowledge-Based Authentication 
        ' system based on a provided question id and answer, and notes that it was validated

        ' The input parameters are as follows:
        '   QuesId      - The reports.KBA_QUES.ROW_ID of the question asked
        '   Answer       - The users answer to a question
        '   UserId      - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug       - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Output:
        '   True or False, depending on whether the answer supplied is correct or not
        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging, temp As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' KBA declarations
        Dim DecodedUserId, AltUserId As String
        Dim ansr_text, Sanswer, prev_ansr As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim myresults As Boolean

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "False"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        ansr_text = ""
        myresults = False
        prev_ansr = ""
        Sanswer = ""    ' Soundex user answer
        'Debug = "Y"

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (QuesId = "" Or Answer = "") And Debug <> "T" Then
            results = "False"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            QuesId = "KBA2948OQUK-1"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
            Answer = "Sutter"
            Sanswer = Soundex1(Answer)
        Else
            ' Fix question id
            QuesId = Trim(HttpUtility.UrlEncode(QuesId))
            If InStr(QuesId, "%") > 0 Then QuesId = Trim(HttpUtility.UrlDecode(QuesId))
            If InStr(QuesId, "%") > 0 Then QuesId = Trim(QuesId)
            QuesId = EncodeParamSpaces(QuesId)
            ' Fix user registration id
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
            ' Fix answer text
            'Answer = Trim(HttpUtility.UrlEncode(Answer))
            If InStr(Answer, "%") > 0 Then Answer = Trim(HttpUtility.UrlDecode(Answer))
            If InStr(Answer, "%") > 0 Then Answer = Trim(Answer)
            Answer = Trim(UCase(Answer))
            Sanswer = Soundex1(Answer)
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("KBAALookup_debug")
            If temp = "Y" And Debug <> "T" And Debug <> "R" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            myresults = False
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\KBAALookup.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                myresults = False
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  Debug: " & Debug)
                mydebuglog.Debug("  QuesId: " & QuesId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  Answer: " & Answer)
                mydebuglog.Debug("  Sanswer: " & Sanswer)
            End If
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            myresults = False
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        If Not cmd Is Nothing Then
            ' Double check user
            Try
                SqlS = "SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT WHERE X_REGISTRATION_NUM='" & DecodedUserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Check to see if the contact id exists: " & SqlS)
                cmd.CommandText = SqlS
                AltUserId = cmd.ExecuteScalar()
                If Debug = "Y" Then mydebuglog.Debug("  > AltUserId: " & AltUserId & vbCrLf)
                If AltUserId = "" Then
                    SqlS = "SELECT X_REGISTRATION_NUM FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & DecodedUserId & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Locate registration number: " & SqlS)
                    cmd.CommandText = SqlS
                    DecodedUserId = cmd.ExecuteScalar()
                    If Debug = "Y" Then mydebuglog.Debug("  > DecodedUserId: " & DecodedUserId & vbCrLf)
                End If
            Catch ex As Exception
                errmsg = errmsg & "Error getting user. " & ex.ToString & vbCrLf
            End Try

            ' Retrieve the correct answer
            Try
                SqlS = "EXEC reports.dbo.OpenHCIKeys; SELECT reports.dbo.HCI_Decrypt(ENC_ANSR_TEXT) AS ANSR_TEXT " &
                    "FROM elearning.dbo.KBA_ANSR " &
                    "WHERE ROW_ID='" & QuesId & "' AND USER_ID='" & DecodedUserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get answer to validate: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            prev_ansr = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToUpper
                            ansr_text = Soundex1(prev_ansr)
                            If Debug = "Y" Then
                                mydebuglog.Debug("  > Found record on query.")
                                mydebuglog.Debug("     Original text: " & prev_ansr)
                                mydebuglog.Debug("     Soundex text: " & ansr_text & vbCrLf)
                            End If
                            If ansr_text = "" Then
                                myresults = False
                                errmsg = errmsg & "Error getting answer. " & vbCrLf
                            End If
                        Catch ex As Exception
                            myresults = False
                            errmsg = errmsg & "Error getting answer. " & ex.ToString & vbCrLf
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Answer was not found." & vbCrLf
                    myresults = False
                    dr.Close()
                    GoTo CloseOut
                End If
                dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                myresults = False
                GoTo CloseOut
            End Try

            ' If no correct answer found, then this means the parameters are incorrect, report failure
            If ansr_text = "" Then
                errmsg = errmsg & "Answer was not found. Parameters incorrect." & vbCrLf
                myresults = False
                GoTo CloseOut
            End If

            ' Compare the soundexed answers to validate
            If ansr_text = Sanswer Or prev_ansr = Answer Then myresults = True
            If Debug = "Y" Then
                mydebuglog.Debug("  > Match Results")
                mydebuglog.Debug("     myresults: " & myresults & vbCrLf)
            End If

            ' Write validated answer and results
            Try
                SqlS = "UPDATE elearning.dbo.KBA_ANSR " &
                "SET VAL_TEXT='" & SqlString(Answer) & "', VAL_DATE=GETDATE() " &
                "WHERE ROW_ID='" & QuesId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Updating user's answer: " & SqlS)
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                If returnv = 0 Then
                    myresults = False
                End If
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & "Error updating record. " & ex.ToString & vbCrLf
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
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
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("KBAALookup : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("KBAALookup : Results: " & myresults & " for QuesId # " & QuesId & " and answer " & Answer)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & myresults & " for QuesId # " & QuesId & " and answer " & Answer & " at " & Now.ToString)
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
        Return myresults
    End Function

    <WebMethod(Description:="Store answer to an exam question")>
    Public Function StoreExamData(ByVal STATUS As String, ByVal ASSESS_ID As String, ByVal TSTRUN_ID As String,
                                  ByVal QUES_ID As String, ByVal ANSR_ID As String, ByVal TEXT_ANSR As String,
                                  ByVal UID As String, ByVal DEBUG As String) As String

        ' This service creates or updates a record stored in elearning.ELN_TEST_ANSWER.  For use with
        ' the HTML5 assessment system

        ' The parameters are as follows:
        '   STATUS      - The current assessment status: [complete, incomplete]
        '   ASSESS_ID   - The CrseId parameter from SCORM player, which is the assessment id - S_CRSE_TST.ROW_ID 
        '   TSTRUN_ID   - The RegId parameter from SCORM player, which is the assessment attempt id - S_CRSE_TSTRUN.ROW_ID
        '   QUES_ID     - The Question Id of the assessment question - S_CRSE_TST_QUES.ROW_ID
        '   ANSR_ID     - The Answer Id of the assessment question answer - S_CRSE_TST_ANSR.ROW_ID
        '   TEXT_ANSR   - The text of the answer or interaction results
        '   UID         - The encoded user id of the assessment taker - S_CONTACT.X_REGISTRATION_NUM
        '   DEBUG       - A flag to indicate the service is to run in Debug mode or not
        '                   "Y"  - Yes for debug mode on.. logging on
        '                   "N"  - No for debug mode off.. logging off
        '                   "T"  - Test mode on.. logging off

        ' When completed, this service either returns:
        '       [Success,Failure]

        ' web.config Parameters used:
        '   siebeldb        - connection string to the database

        ' Variables
        Dim results, temp As String
        Dim errmsg As String
        Dim i, returnv As Integer

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
        mydebuglog = log4net.LogManager.GetLogger("SEDDebugLog")
        Dim logfile, Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Processing variables
        Dim ScormRegId, AssessIdent, STATUS_CD As String
        Dim DECODE_UID As String

        ' ============================================
        ' Variable setup
        errmsg = ""
        results = "Success"
        Logging = "Y"
        SqlS = ""
        returnv = 0
        ScormRegId = ""
        AssessIdent = ""
        STATUS_CD = ""
        DECODE_UID = ""

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("siebeldb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("StoreExamData_Debug")
            If DEBUG <> "T" And temp <> "" Then DEBUG = temp
        Catch ex As Exception
            errmsg = errmsg & "Unable to get defaults from web.config. " & vbCrLf
            results = ""
            GoTo CloseOut2
        End Try

        If DEBUG = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Passed Parameters-")
            mydebuglog.Debug("  TSTRUN_ID: " & TSTRUN_ID)
            mydebuglog.Debug("  STATUS: " & STATUS)
            mydebuglog.Debug("  ASSESS_ID: " & ASSESS_ID)
            mydebuglog.Debug("  QUES_ID: " & QUES_ID)
            mydebuglog.Debug("  ANSR_ID: " & ANSR_ID)
            mydebuglog.Debug("  TEXT_ANSR: " & TEXT_ANSR)
            mydebuglog.Debug("  UID: " & UID)
        End If

        ' ============================================
        ' Fix parameters
        DEBUG = UCase(Left(DEBUG, 1))
        If DEBUG = "T" Then
            TSTRUN_ID = "96T9645ZU56J"
            STATUS = "Completed"
            ASSESS_ID = "SZ4J51W"
            QUES_ID = "SZ4J51W-44"
            ANSR_ID = "SZ4J51W-44-4"
            UID = "=oENyQzRxkjW5MTN"
            DECODE_UID = "539Z91G424J"
            TEXT_ANSR = ""
        Else
            TSTRUN_ID = Trim(HttpUtility.UrlEncode(TSTRUN_ID))
            'If InStr(TSTRUN_ID, "%") > 0 Then TSTRUN_ID = Trim(HttpUtility.UrlDecode(TSTRUN_ID))
            TSTRUN_ID = Trim(HttpUtility.UrlDecode(TSTRUN_ID))
            'If InStr(TSTRUN_ID, "%") > 0 Then TSTRUN_ID = Trim(TSTRUN_ID)
            STATUS = Trim(STATUS)
            ASSESS_ID = Trim(HttpUtility.UrlEncode(ASSESS_ID))
            'If InStr(ASSESS_ID, "%") > 0 Then ASSESS_ID = Trim(HttpUtility.UrlDecode(ASSESS_ID))
            ASSESS_ID = Trim(HttpUtility.UrlDecode(ASSESS_ID))
            'If InStr(ASSESS_ID, "%") > 0 Then ASSESS_ID = Trim(ASSESS_ID)
            QUES_ID = Trim(HttpUtility.UrlEncode(QUES_ID))
            If InStr(QUES_ID, "%") > 0 Then QUES_ID = Trim(HttpUtility.UrlDecode(QUES_ID))
            QUES_ID = Trim(HttpUtility.UrlDecode(QUES_ID))
            'If InStr(QUES_ID, "%") > 0 Then QUES_ID = Trim(QUES_ID)
            ANSR_ID = Trim(HttpUtility.UrlEncode(ANSR_ID))
            If InStr(ANSR_ID, "%") > 0 Then ANSR_ID = Trim(HttpUtility.UrlDecode(ANSR_ID))
            ANSR_ID = Trim(HttpUtility.UrlDecode(ANSR_ID))
            'If InStr(ANSR_ID, "%") > 0 Then ANSR_ID = Trim(ANSR_ID)
            TEXT_ANSR = Trim(HttpUtility.UrlEncode(TEXT_ANSR))
            'If InStr(TEXT_ANSR, "%") > 0 Then TEXT_ANSR = Trim(HttpUtility.UrlDecode(TEXT_ANSR))
            TEXT_ANSR = Trim(HttpUtility.UrlDecode(TEXT_ANSR))
            'If InStr(TEXT_ANSR, "%") > 0 Then TEXT_ANSR = Trim(TEXT_ANSR)
            TEXT_ANSR = SqlString(TEXT_ANSR)
            UID = Trim(HttpUtility.UrlEncode(UID))
            'If InStr(UID, "%") > 0 Then UID = Trim(HttpUtility.UrlDecode(UID))
            UID = Trim(HttpUtility.UrlDecode(UID))
            'If InStr(UID, "%") > 0 Then UID = Trim(UID)
            DECODE_UID = FromBase64(ReverseString(UID))
        End If

        ' ============================================
        ' Open log file if applicable
        If DEBUG = "Y" Or (Logging = "Y" And DEBUG <> "T") Then
            logfile = "C:\Logs\StoreExamData.log"
            Try
                log4net.GlobalContext.Properties("SEDLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & "Error Opening Log. " & vbCrLf
                results = ""
                GoTo CloseOut2
            End Try

            If DEBUG = "Y" Then
                'mydebuglog.Debug("----------------------------------")
                'mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("")
                mydebuglog.Debug("Modified Parameters-")
                mydebuglog.Debug("  TSTRUN_ID: " & TSTRUN_ID)
                mydebuglog.Debug("  STATUS: " & STATUS)
                mydebuglog.Debug("  ASSESS_ID: " & ASSESS_ID)
                mydebuglog.Debug("  QUES_ID: " & QUES_ID)
                mydebuglog.Debug("  ANSR_ID: " & ANSR_ID)
                mydebuglog.Debug("  TEXT_ANSR: " & TEXT_ANSR)
                mydebuglog.Debug("  DECODE_UID: " & DECODE_UID)
            End If
        End If

        ' ============================================
        ' Check parameters
        If (TSTRUN_ID = "" Or QUES_ID = "" Or ASSESS_ID = "" Or UID = "") And DEBUG <> "T" Then
            results = ""
            errmsg = errmsg & "Invalid parameter(s) " & vbCrLf
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open database connections 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = ""
            GoTo CloseOut
        End If

        ' ============================================
        ' Process
        If Not cmd Is Nothing Then

            ' -----
            ' Verify Assessment Taker
            Try
                SqlS = "SELECT T.MS_IDENT, T.STATUS_CD " &
                "FROM siebeldb.dbo.S_CRSE_TSTRUN T " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST A ON A.ROW_ID=T.CRSE_TST_ID " &
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=T.PERSON_ID " &
                "WHERE C.X_REGISTRATION_NUM='" & DECODE_UID & "' AND A.ROW_ID='" & ASSESS_ID & "' AND T.ROW_ID='" & TSTRUN_ID & "'"
                If DEBUG = "Y" Then mydebuglog.Debug(vbCrLf & "  Verify assessment taker: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            AssessIdent = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                            STATUS_CD = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            If DEBUG = "Y" Then
                                mydebuglog.Debug("  > AssessIdent: " & AssessIdent)
                                mydebuglog.Debug("  > STATUS_CD: " & STATUS_CD)
                            End If
                            If AssessIdent = "" Then results = "Failure"
                        Catch ex As Exception
                            results = "Failure"
                        End Try
                    End While
                Else
                    errmsg = errmsg & "The scorm_registration_id was not found: " & TSTRUN_ID & vbCrLf
                    results = "Failure"
                End If

                ' Close reader
                Try
                    dr.Close()
                Catch ex As Exception
                End Try

            Catch oBug As Exception
                If DEBUG = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try
            If results = "Failure" Then
                errmsg = errmsg & "The user was not found to have this assessment and attempt: " & DECODE_UID & " / " & ASSESS_ID & " / " & TSTRUN_ID & vbCrLf
                GoTo CloseOut
            End If
            If STATUS_CD = "Graded" Then
                errmsg = errmsg & "This exam was already scored and cannot be taken again: " & DECODE_UID & " / " & ASSESS_ID & " / " & TSTRUN_ID & vbCrLf
                GoTo CloseOut
            End If

            ' -----
            ' Get Registration Id - this is stored in ELN_TEST_ANSWER.SESS_REG_NUM
            Try
                'SqlS = "SELECT scorm_registration_id " & _
                '    "FROM hciscorm.dbo.ScormRegistration " & _
                '    "WHERE reg_id='" & TSTRUN_ID & "'  AND crse_type='A'"
                SqlS = "SELECT player_reg_id " &
                "FROM  elearning.dbo.Elearning_Player_Data " &
                "WHERE reg_id='" & TSTRUN_ID & "'  AND crse_type='A'"
                If DEBUG = "Y" Then mydebuglog.Debug(vbCrLf & "  Get scorm_registration_id: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            ScormRegId = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                            If DEBUG = "Y" Then mydebuglog.Debug("  > ScormRegId: " & ScormRegId)
                            If ScormRegId = "" Then results = "Failure"
                        Catch ex As Exception
                            results = "Failure"
                        End Try
                    End While
                Else
                    results = "Failure"
                End If

                ' Close reader
                Try
                    dr.Close()
                Catch ex As Exception
                End Try
            Catch oBug As Exception
                If DEBUG = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try
            If results = "Failure" Then
                errmsg = errmsg & "The scorm_registration_id was not found: " & TSTRUN_ID & vbCrLf
                GoTo CloseOut
            End If

            ' -----
            ' Replace answer 
            If ScormRegId <> "" Then
                ' Delete any current record - errors are suppressed
                SqlS = "DELETE FROM elearning.dbo.ELN_TEST_ANSWER " &
                    "WHERE S_CRSE_TSTRUN_ID = '" & TSTRUN_ID & "' AND QUESTION_ID='" & QUES_ID & "'"
                'If ANSR_ID = "" Then
                'Else
                'SqlS = "DELETE FROM elearning.dbo.ELN_TEST_ANSWER " & _
                '   "WHERE S_CRSE_TSTRUN_ID = '" & TSTRUN_ID & "' AND QUESTION_ID='" & QUES_ID & "' AND ANSWER_ID='" & ANSR_ID & "'"
                'End If
                If DEBUG = "Y" Then mydebuglog.Debug(vbCrLf & "  Delete existing ELN_TEST_ANSWER: " & SqlS)
                cmd.CommandText = SqlS
                Try
                    returnv = cmd.ExecuteNonQuery()
                Catch ex As Exception
                End Try

                ' Insert the record(s)
                If ANSR_ID <> "" Then
                    If InStr(ANSR_ID, ",") > 0 Then
                        Dim ANSR_ARR() As String
                        ANSR_ARR = ANSR_ID.Split(",")
                        For i = 0 To ANSR_ARR.Length - 1
                            SqlS = "INSERT INTO elearning.dbo.ELN_TEST_ANSWER " &
                                "(EXAM_ID,QUESTION_ID,ANSWER_ID,S_CRSE_TSTRUN_ID,TEXT_ANSWER,CREATED,SESS_REG_NUM) " &
                                "VALUES ('" & ASSESS_ID & "','" & QUES_ID & "','" & ANSR_ARR(i) & "','" & TSTRUN_ID & "','" &
                                TEXT_ANSR & "',GETDATE(),'" & ScormRegId & "')"
                            If DEBUG = "Y" Then mydebuglog.Debug(vbCrLf & "  Insert ELN_TEST_ANSWER: " & SqlS)
                            cmd.CommandText = SqlS
                            Try
                                returnv = cmd.ExecuteNonQuery()
                                If returnv = 0 Then
                                    errmsg = errmsg & "Error inserting ELN_TEST_ANSWER. " & vbCrLf & "Query: " & SqlS & vbCrLf
                                    results = "Failure"
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & "Error inserting ELN_TEST_ANSWER. " & ex.ToString & vbCrLf & "Query: " & SqlS & vbCrLf
                                results = "Failure"
                            End Try
                        Next
                    Else
                        SqlS = "INSERT INTO elearning.dbo.ELN_TEST_ANSWER " &
                        "(EXAM_ID,QUESTION_ID,ANSWER_ID,S_CRSE_TSTRUN_ID,TEXT_ANSWER,CREATED,SESS_REG_NUM) " &
                        "VALUES ('" & ASSESS_ID & "','" & QUES_ID & "','" & ANSR_ID & "','" & TSTRUN_ID & "','" &
                        TEXT_ANSR & "',GETDATE(),'" & ScormRegId & "')"
                        If DEBUG = "Y" Then mydebuglog.Debug(vbCrLf & "  Insert ELN_TEST_ANSWER: " & SqlS)
                        cmd.CommandText = SqlS
                        Try
                            returnv = cmd.ExecuteNonQuery()
                            If returnv = 0 Then
                                errmsg = errmsg & "Error inserting ELN_TEST_ANSWER. " & vbCrLf & "Query: " & SqlS & vbCrLf
                                results = "Failure"
                            End If
                        Catch ex As Exception
                            errmsg = errmsg & "Error inserting ELN_TEST_ANSWER. " & ex.ToString & vbCrLf & "Query: " & SqlS & vbCrLf
                            results = "Failure"
                        End Try
                    End If
                End If
            Else
                errmsg = "Record scorm_registration_id not found: " & TSTRUN_ID & vbCrLf
                results = "Failure"
            End If
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            Try
                errmsg = errmsg & CloseDBConnection(con, cmd, dr)
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
            End Try
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If DEBUG <> "T" Then
            If DEBUG = "Y" Or Logging = "Y" Then
                mydebuglog.Debug(vbCrLf & "  " & Format(Now) & "   TSTRUN_ID/QUES_ID/ANSR_ID: " & TSTRUN_ID & " / " & QUES_ID & " / " & ANSR_ID & " > Results: " & results)
            End If
            If Trim(errmsg) <> "" Then
                myeventlog.Error("StoreExamData : Error: " & Trim(errmsg) & ", UID/TSTRUN_ID/QUES_ID/ANSR_ID: " & DECODE_UID & " / " & TSTRUN_ID & " / " & QUES_ID & " / " & ANSR_ID)
            Else
                myeventlog.Info("StoreExamData : Results: " & results & " UID/TSTRUN_ID/QUES_ID/ANSR_ID: " & DECODE_UID & " / " & TSTRUN_ID & " / " & QUES_ID & " / " & ANSR_ID)
            End If
            If DEBUG = "Y" Or Logging = "Y" Then
                Try
                    If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                    If DEBUG = "Y" Then
                        mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                        mydebuglog.Debug("----------------------------------")
                    End If
                Catch ex As Exception
                End Try
            End If
        End If

        ' Log Performance Data
        If DEBUG <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, DEBUG)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return results
    End Function

    <WebMethod(Description:="Retrieves the specified asset")> _
    Public Function GetMedia(ByVal RegId As String, ByVal UserId As String, ByVal Type As String, ByVal Debug As String, ByVal Asset As String) As Byte()

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
        '   Debug	    - "Y", "N" or "T"
        '   
        '   Asset	    - The DMS.Documents.dfilename of the asset to be retrieved, or if "default.jpg",
        '                   the first associated item in the category "Images" will be returned.

        ' web.config Parameters used:
        '   hcidb           - connection string to hcidb1.siebeldb database
        '   dms        	    - connection string to DMS.dms database
        '   cache           - connection string to cache.sdf database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String
        Dim DecodedUserId, ValidatedUserId As String

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
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("GMDebugLog")
        Dim logfile, temp As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' File handling declarations
        'Dim bfs As FileStream
        'Dim bw As BinaryWriter
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim d_dsize, d_doc_id, SaveDest As String
        Dim dLastUpd As DateTime
        Dim CRSE_ID, minio_flg, d_verid As String
        Dim killcount As Double

        Dim filecache As ObjectCache = MemoryCache.Default
        Dim fileContents(1000) As Byte
        Dim cacheItemName As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

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
        minio_flg = "N"
        d_verid = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If RegId = "" And Debug <> "T" Then
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
        If Trim(Asset) = "" Then
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
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb;ApplicationIntent=ReadOnly"
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIBDB;uid=DMS;pwd=5241200;database=DMS;ApplicationIntent=ReadOnly"
            ''c_ConnS = "Data Source=" & mypath & "cachedb\\cache.sdf;Password=s3v3n0n3;Persist Security Info=False;"
            'c_ConnS = "Server=(LocalDB)\MSSQLLocalDB;MultipleActiveResultSets=True;Integrated Security=true;AttachDbFileName=" & mypath & "cachedb\cache.mdf"  'Switched to SQL Server Express LocalDB 2014
            dms_cache_age = Trim(System.Configuration.ConfigurationManager.AppSettings("dmscacheage"))
            If dms_cache_age = "" Or Not IsNumeric(dms_cache_age) Then dms_cache_age = "30"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetMedia_debug")
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
            logfile = "C:\Logs\GetMedia.log"
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
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Type: " & Type)
                mydebuglog.Debug("  Asset: " & Asset)
                mydebuglog.Debug("  TypeTrans: " & TypeTrans)
                mydebuglog.Debug("  AccessBucket: " & AccessBucket)
                mydebuglog.Debug("  Appsetting dms_cache_age: " & dms_cache_age & vbCrLf)
            End If
        End If

        ' ============================================
        ' Open database connections
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
                    SqlS = "SELECT R.CRSE_ID, C.X_REGISTRATION_NUM " & _
                    "FROM siebeldb.dbo.CX_SESS_REG R " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
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
                If Debug = "Y" Then mydebuglog.Debug("   ... CRSE_ID: " & CRSE_ID & vbCrLf)

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
                SaveDest = mypath & Replace(Trim(Type.ToLower), " ", "_") & "_temp\" & RegId
        End Select
        Try
            Directory.CreateDirectory(SaveDest)
        Catch
        End Try
        If Debug = "Y" Then mydebuglog.Debug("  Asset caching: " & SaveDest & vbCrLf)

        ' ============================================
        ' Get the name of the asset if necessary
        If Debug = "Y" Then mydebuglog.Debug("  Looking in database for: " & LCase(Trim(System.IO.Path.GetFileNameWithoutExtension(Asset))) & vbCrLf)
        If LCase(Trim(System.IO.Path.GetFileNameWithoutExtension(Asset))) = "default" Then
            If Not d_cmd Is Nothing Then
                ' Query DMS
                SqlS = "SELECT TOP 1 D.dfilename, D.row_id, D.last_upd " & _
                    "FROM DMS.dbo.Documents D  " & _
                    "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id  " & _
                    "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id  " & _
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id  " & _
                    "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id  " & _
                    "WHERE D.row_id IS NOT NULL AND D.deleted IS NULL AND LOWER(A.name)='" & Type & "' AND C.name='Images' " & _
                    "AND DA.fkey='" & RegId & "' " & _
                    "ORDER BY D.last_upd DESC"
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
        ' Generate cache filename
        BinaryFile = SaveDest & "\" & Asset
        BinaryFile = BinaryFile.Replace(mypath, "")
        If Debug = "Y" Then mydebuglog.Debug("  Cache filename: " & BinaryFile & vbCrLf)

        ' ============================================
        ' Get DMS record containing asset 
        '  If the cache was hit, make sure the entry is current, otherwise restore it
        Dim upt_SqlStr As String = ""
        upt_SqlStr = "SELECT TOP 1 v.last_upd "
        If Not d_cmd Is Nothing Then
            ' -----
            ' Construct document search query for DMS
            If d_doc_id <> "" Then
                SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.last_upd, v.minio_flg, v.row_id, d.row_id " &
                    "FROM DMS.dbo.Documents d " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                    "WHERE d.row_id=" & d_doc_id
                upt_SqlStr = upt_SqlStr & "FROM DMS.dbo.Documents d " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                    "WHERE d.row_id=" & d_doc_id
            Else
                Select Case Type
                    Case "asset"
                        SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.last_upd, v.minio_flg, v.row_id, d.row_id " &
                            "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id "
                        upt_SqlStr = upt_SqlStr & "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id "
                        If CRSE_ID <> "" Then
                            SqlS = SqlS & "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=" &
                            TypeTrans & " and (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                            upt_SqlStr = upt_SqlStr & "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=" &
                            TypeTrans & " and (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                        Else
                            SqlS = SqlS & "WHERE (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                            upt_SqlStr = upt_SqlStr & "WHERE (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                        End If
                    Case "resource"
                        SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.last_upd, v.minio_flg, v.row_id, d.row_id " &
                            "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id " &
                            "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=" &
                            TypeTrans & " and (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                        upt_SqlStr = upt_SqlStr & "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id " &
                            "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=" &
                            TypeTrans & " and (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                    Case Else
                        SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.last_upd, v.minio_flg, v.row_id, d.row_id " &
                            "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da on da.doc_id=d.row_id " &
                            "LEFT OUTER JOIN DMS.dbo.Associations a on a.row_id=da.association_id "
                        upt_SqlStr = upt_SqlStr & "FROM DMS.dbo.Documents d " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                            "LEFT OUTER JOIN DMS.dbo.Document_Associations da on da.doc_id=d.row_id " &
                            "LEFT OUTER JOIN DMS.dbo.Associations a on a.row_id=da.association_id "
                        If CRSE_ID <> "" Then
                            SqlS = SqlS & "WHERE da.fkey='" & CRSE_ID & "' and lower(a.name)='" & Trim(Type.ToLower) & "' and " &
                            "(d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                            upt_SqlStr = upt_SqlStr & "WHERE da.fkey='" & CRSE_ID & "' and lower(a.name)='" & Trim(Type.ToLower) & "' and " &
                            "(d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                        Else
                            SqlS = SqlS & "WHERE (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                            upt_SqlStr = upt_SqlStr & "WHERE (d.dfilename='" & Asset & "' or d.name='" & Asset & "') " &
                            "AND d.deleted is null " &
                            "ORDER BY v.version DESC, d.last_upd DESC"
                        End If

                End Select
            End If

            ' **** Construct Query for last update check  ***********
            'upt_SqlStr = "SELECT TOP 1 v.last_upd " & SqlS.Substring(SqlS.IndexOf("FROM "))

            '***** Check cached object; 3/2/17; Ren Hou;  ****'
            Dim last_upt As Date = Today.AddYears(-50)
            ' Check to see if the document is in the in-memory cache
            cacheItemName = BinaryFile
            Dim docNotInDB As Boolean = 0
            If Not IsNothing(filecache(cacheItemName)) Then
                'Check if the cached item need to be renewed;
                Try
                    If Debug = "Y" Then mydebuglog.Debug("  Check last updated date for xml document: " & upt_SqlStr)
                    d_cmd.CommandText = upt_SqlStr
                    last_upt = d_cmd.ExecuteScalar()
                    If Debug = "Y" Then mydebuglog.Debug("    > last_upt: " & last_upt.ToString)
                Catch ex As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
                    results = "Failure"
                End Try

                If last_upt = Today.AddYears(-50) Then  'document no longer exists in the database
                    docNotInDB = True
                    filecache.Remove(cacheItemName)
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   " & Asset & " is not found in database! " & vbCrLf)
                    'results = "Failure"
                ElseIf last_upt > TryCast(filecache(cacheItemName), HciDMSDocument).UpdateDate Then
                    'Remove if the update_date on the cache is before the last updtaed date on DB record.
                    filecache.Remove(cacheItemName)
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Cached object " & cacheItemName & " expired.")
                End If
            Else
                If Debug = "Y" Then mydebuglog.Debug("  .." & cacheItemName & " not found in cache" & vbCrLf)
            End If
            '****  ****'

            'Load content of cached object
            If filecache(cacheItemName) Is Nothing Then
                fileContents = Nothing
            Else
                Dim tmpObj As Object
                tmpObj = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
                ReDim fileContents(tmpObj.Length)
                fileContents = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
                CacheHit = 1
                docNotInDB = False
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Retrieved Cached object " & cacheItemName & vbCrLf)
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
                                minio_flg = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                                d_verid = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(5), enumObjectType.StrType))
                                dLastUpd = d_dr(2)
                                If Debug = "Y" Then mydebuglog.Debug("  > Record found on query:  d_doc_id=" & d_doc_id & ",  d_verid=" & d_verid & ",  d_dsize=" & d_dsize & ",  minio_flg=" & minio_flg & ",  dLastUpd=" & Format(dLastUpd) & ",  cLastUpd=" & Convert.ToString(LastUpd) & ",  CacheHit=" & Format(CacheHit))

                                If minio_flg = "Y" Then
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Minio")
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
                                            ReDim outbyte(Val(retval - 1))
                                            Dim intval As Integer
                                            For i = 0 To retval - 1
                                                intval = mobj2.ResponseStream.ReadByte()
                                                If intval < 255 And intval > 0 Then
                                                    outbyte(i) = intval
                                                End If
                                                If intval = 255 Then outbyte(i) = 255
                                                If intval < 0 Then
                                                    outbyte(i) = 0
                                                    If Debug = "Y" Then mydebuglog.Debug("  >  .. " & i.ToString & "   intval: " & intval.ToString)
                                                End If
                                            Next
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
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Document_Versions")
                                    ' Get binary and attach to the object outbyte if found, not cached or updated recently
                                    '   retval will be "0" if this is not the case
                                    If d_dsize <> "" And (CacheHit = 0 Or dLastUpd > LastUpd) Then
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                    End If
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
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting asset: " & oBug.ToString)
                    results = "Failure"
                End Try

                If Debug = "Y" Then
                    mydebuglog.Debug("   ... Reported d_dsize: " & d_dsize & " : Non-cached DMS doc found: " & Str(retval) & vbCrLf)
                End If

                '***** Set cache object; 2/28/17; Ren Hou;  ****
                Dim policy As New CacheItemPolicy()
                policy.SlidingExpiration = TimeSpan.FromDays(CDbl(dms_cache_age))
                If Debug = "Y" Then mydebuglog.Debug("  Caching DMS doc to key: " & cacheItemName & vbCrLf)
                filecache.Set(cacheItemName, New HciDMSDocument(dLastUpd, outbyte), policy)
                ReDim fileContents(outbyte.Length)
                fileContents = outbyte
                '****  ****
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
        If Trim(errmsg) <> "" Then myeventlog.Error("GetMedia : Error: " & Trim(errmsg))
        If CacheHit = 1 Then
            If Debug <> "T" Then myeventlog.Info("GetMedia : Results: " & results & " for CACHED " & Type & " file: " & Asset & " by UserId # " & DecodedUserId & " with RegId " & RegId)
        Else
            If Debug <> "T" Then myeventlog.Info("GetMedia : Results: " & results & " for " & Type & " file: " & Asset & ", minio: " & minio_flg & ", doc id: " & d_doc_id & ", verid: " & d_verid & ", by UserId # " & DecodedUserId & " with RegId " & RegId)
        End If
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If CacheHit = 1 Then
                    mydebuglog.Debug("Results: " & results & " for CACHED " & Type & " file " & Asset & ", by UserId # " & DecodedUserId & " with RegId " & RegId & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
                Else
                    mydebuglog.Debug("Results: " & results & " for " & Type & " file " & Asset & ", minio: " & minio_flg & ", doc id: " & d_doc_id & ", verid: " & d_verid & ", by UserId # " & DecodedUserId & " with RegId " & RegId & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
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

    <WebMethod(Description:="Retrieves the specified web part item")> _
    Public Function GetHMedia(ByVal CrseId As String, ByVal ItemName As String, ByVal Debug As String) As Byte()

        ' This function locates the specified item, and returns it to the calling system 

        ' The input parameters are as follows:
        '
        '   CrseId       - The S_CRSE.ROW_ID of the course for which the retrieved web part name belongs.  
        '		            This maps to DMS.Document_Associations.fkey
        '   
        '   ItemName	 - The DMS.Documents.dfilename of the item to be retrieved
        '
        '   Debug	    - "Y", "N" or "T"
        '

        ' web.config Parameters used:
        '   hcidb           - connection string to hcidb1.siebeldb database
        '   dms        	    - connection string to DMS.dms database
        '   cache           - connection string to cache.sdf database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer
        Dim TypeTrans As String

        ' Cache database declarations
        'Dim c_ConnS As String
        Dim CacheHit As Integer
        Dim dAsset, dCrseId, dFileName As String
        Dim LastUpd As DateTime

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String
        Dim dms_cache_age As String

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("GHMDebugLog")
        Dim logfile, temp As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' File handling declarations
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim d_dsize, d_doc_id, SaveDest As String
        Dim dLastUpd As DateTime
        Dim CRSE_ID As String
        Dim killcount As Double

        Dim filecache As ObjectCache = MemoryCache.Default
        Dim fileContents(1000) As Byte
        Dim cacheItemName As String
        Dim minio_flg, d_verid As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        CRSE_ID = ""
        d_dsize = ""
        BinaryFile = ""
        'c_ConnS = ""
        dAsset = ""
        d_doc_id = ""
        dCrseId = ""
        dFileName = ""
        TypeTrans = ""
        'Debug = "Y"
        killcount = 0
        temp = ""
        minio_flg = ""
        d_verid = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If Trim(CrseId) = "" And Trim(ItemName) <> "" And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            CrseId = "CRSE15458AN"
            ItemName = "license1.png"
        Else
            If InStr(CrseId, "%") > 0 Then CrseId = Trim(HttpUtility.UrlDecode(CrseId))
            CrseId = EncodeParamSpaces(CrseId)
            If InStr(ItemName, "%") > 0 Then ItemName = Trim(HttpUtility.UrlDecode(ItemName))
        End If
        If Trim(CrseId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No course id specified. "
            GoTo CloseOut2
        End If
        If Trim(ItemName) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No item filename specified. "
            GoTo CloseOut2
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidbro").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb;ApplicationIntent=ReadOnly"
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIBDB;uid=DMS;pwd=5241200;database=DMS"
            'c_ConnS = "Server=(LocalDB)\MSSQLLocalDB;Integrated Security=true;AttachDbFileName=" & mypath & "cachedb\cache.mdf"  'Switched to SQL Server Express LocalDB 2014
            dms_cache_age = Trim(System.Configuration.ConfigurationManager.AppSettings("dmscacheage"))
            If dms_cache_age = "" Or Not IsNumeric(dms_cache_age) Then dms_cache_age = "30"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetHMedia_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y" Else Debug = "N"
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
            logfile = "C:\Logs\GetHMedia.log"
            Try
                log4net.GlobalContext.Properties("GHMLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  CrseId: " & CrseId)
                mydebuglog.Debug("  ItemName: " & ItemName)
                mydebuglog.Debug("  ConnS: " & ConnS)
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

        '' ============================================
        '' Open SQL Server CE database connection
        'AppDomain.CurrentDomain.SetData("SQLServerCompactEditionUnderWebHosting", True)
        'Dim c_con As New SqlConnection(c_ConnS)
        'Dim c_cmd As New SqlCommand()
        'Dim c_dr As SqlDataReader
        'If Not (My.Computer.FileSystem.FileExists(mypath & "cachedb\cache.mdf")) Then
        '    ' Create the database if necessary
        '    Try
        '        c_con.ConnectionString = c_ConnS.Substring(0, c_ConnS.IndexOf(";AttachDbFileName", 0) + 1)  'connect to the localDB instance instead of 'cache' database
        '        c_cmd.Connection = c_con
        '        c_con.Open()

        '        ' Remove old database if around
        '        Try
        '            SqlS = "DROP DATABASE [cache]"
        '            If Debug = "Y" Then mydebuglog.Debug("  Drop existing cache database query: " & SqlS & vbCrLf)
        '            c_cmd.CommandText = SqlS
        '            returnv = c_cmd.ExecuteNonQuery()
        '            con.Close()
        '        Catch ex3 As Exception
        '            If Debug = "Y" Then mydebuglog.Debug("  Error dropping cache database: " & ex3.Message & vbCrLf)
        '        End Try

        '        ' Create the new database
        '        Try
        '            c_cmd.CommandText = "CREATE DATABASE [cache] ON (NAME = cache, FILENAME = '" & mypath & "cachedb\cache.mdf')"
        '            If Debug = "Y" Then mydebuglog.Debug("  Create cache database: " & c_cmd.CommandText & vbCrLf)
        '            If c_con.State = Data.ConnectionState.Closed Then c_con.Open()
        '            returnv = c_cmd.ExecuteNonQuery()
        '            System.Threading.Thread.Sleep(5000)
        '            c_con.Close()
        '        Catch ex3 As Exception
        '            If Debug = "Y" Then mydebuglog.Debug("  Error creating cache database: " & ex3.Message & vbCrLf)
        '        End Try

        '        ' Create the cache table
        '        Try
        '            SqlS = "CREATE TABLE cache.dbo.[CACHE]( " & _
        '                "[CRSE_ID] [nvarchar](15) NOT NULL, " & _
        '                "[ASSET] [nvarchar](100) NOT NULL, " & _
        '                "[CREATED] [datetime] NOT NULL, " & _
        '                "[LAST_UPD] [datetime] NOT NULL)"
        '            If Debug = "Y" Then mydebuglog.Debug("  Create cache database table: " & SqlS & vbCrLf)
        '            If c_con.State = Data.ConnectionState.Closed Then c_con.Open()
        '            c_cmd.CommandText = SqlS
        '            returnv = c_cmd.ExecuteNonQuery()
        '        Catch ex3 As Exception
        '            If Debug = "Y" Then mydebuglog.Debug("  Error creating cache database table: " & ex3.Message & vbCrLf)
        '        End Try

        '    Catch ex2 As Exception
        '        errmsg = errmsg & vbCrLf & "Unable to create database: " & ex2.Message
        '        results = "Failure"
        '        GoTo CloseOut2
        '    End Try
        'Else
        '    ' Open cache database
        '    Try
        '        c_con = New SqlConnection(c_ConnS)
        '        c_con.Open()
        '        If Not c_con Is Nothing Then
        '            Try
        '                c_cmd = New SqlCommand(SqlS, c_con)
        '            Catch ex2 As Exception
        '                errmsg = errmsg & vbCrLf & "Unable to open cache database: " & ex2.Message
        '                results = "Failure"
        '                GoTo CloseOut2
        '            End Try
        '        Else
        '            errmsg = errmsg & vbCrLf & "Unable to open cache database"
        '            results = "Failure"
        '            GoTo CloseOut2
        '        End If

        '        ' Create the cache table
        '        Try
        '            SqlS = "IF OBJECT_ID ('cache.dbo.[CACHE]', 'U') IS NULL " & _
        '                "CREATE TABLE cache.dbo.[CACHE]( " & _
        '                "[CRSE_ID] [nvarchar](15) NOT NULL, " & _
        '                "[ASSET] [nvarchar](100) NOT NULL, " & _
        '                "[CREATED] [datetime] NOT NULL, " & _
        '                "[LAST_UPD] [datetime] NOT NULL)"
        '            If Debug = "Y" Then mydebuglog.Debug("  Create cache database table: " & SqlS & vbCrLf)
        '            If c_con.State = Data.ConnectionState.Closed Then c_con.Open()
        '            c_cmd.CommandText = SqlS
        '            returnv = c_cmd.ExecuteNonQuery()
        '        Catch ex3 As Exception
        '            If Debug = "Y" Then mydebuglog.Debug("  Error creating cache database table: " & ex3.Message & vbCrLf)
        '        End Try

        '    Catch ex As Exception
        '        errmsg = errmsg & vbCrLf & "Unable to open cache database: " & ex.Message
        '        results = "Failure"
        '        GoTo CloseOut2
        '    End Try

        'End If

        '' Open 2nd connection for cache cleaning
        'Dim c_con2 As New SqlConnection(c_ConnS)
        'Dim c_cmd2 As New SqlCommand()
        'Try
        '    c_con2 = New SqlConnection(c_ConnS)
        '    c_con2.Open()
        '    If Not c_con2 Is Nothing Then
        '        Try
        '            c_cmd2 = New SqlCommand(SqlS, c_con2)
        '        Catch ex2 As Exception
        '            errmsg = errmsg & vbCrLf & "Unable to open cache database: " & ex2.Message
        '            results = "Failure"
        '            GoTo CloseOut2
        '        End Try
        '    Else
        '        errmsg = errmsg & vbCrLf & "Unable to open cache database"
        '        results = "Failure"
        '        GoTo CloseOut2
        '    End If
        'Catch ex As Exception
        '    errmsg = errmsg & vbCrLf & "Unable to open cache database: " & ex.Message
        '    results = "Failure"
        '    GoTo CloseOut2
        'End Try

        ' ============================================
        ' Validate course id
        If Not cmd Is Nothing Then
            ' -----
            ' Query course
            Try
                SqlS = "SELECT (SELECT CASE WHEN X_FORMAT='HTML5' THEN ROW_ID ELSE '' END) AS ROW_ID " & _
                "FROM siebeldb.dbo.S_CRSE " & _
                "WHERE ROW_ID='" & CrseId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Checking course: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            CRSE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error checking course record. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "The course record was not found." & vbCrLf
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
            ' Verify the course id
            If CrseId <> CRSE_ID Then
                results = "Failure"
                CRSE_ID = ""
                errmsg = errmsg & "Course not validated: " & CrseId & " not an HTML5 course"
                GoTo CloseOut
            End If
        Else
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Create output directory for asset caching
        SaveDest = mypath & "course_temp\" & CRSE_ID
        'Try
        '    Directory.CreateDirectory(SaveDest)
        'Catch
        'End Try
        'If Debug = "Y" Then mydebuglog.Debug("   Item caching directory: " & SaveDest & vbCrLf)

        ' ============================================
        ' Get the name of the item if necessary
        If Debug = "Y" Then mydebuglog.Debug("   Looking for: " & LCase(Trim(System.IO.Path.GetFileNameWithoutExtension(ItemName))) & vbCrLf)

        '' ============================================
        '' Check to see if we've already cached this item and when
        'CacheHit = 0
        'If Not c_cmd Is Nothing Then
        '    SqlS = "SELECT LAST_UPD " & _
        '    "FROM [cache].dbo.[CACHE] " & _
        '    "WHERE CRSE_ID='" & CRSE_ID & "' AND ASSET='" & ItemName & "'"
        '    If Debug = "Y" Then mydebuglog.Debug("  Check cache: " & SqlS)
        '    Try
        '        c_cmd.CommandText = SqlS
        '        c_dr = c_cmd.ExecuteReader()
        '        If Not c_dr Is Nothing Then
        '            While c_dr.Read()
        '                Try
        '                    LastUpd = c_dr(0)
        '                    CacheHit = 1
        '                    If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  LastUpd=" & Format(LastUpd) & "  CacheHit=" & CacheHit)
        '                Catch ex As Exception
        '                    results = "Failure"
        '                    errmsg = errmsg & "Error checking cache. " & ex.ToString & vbCrLf
        '                    GoTo CloseOut
        '                End Try
        '            End While
        '        Else
        '            errmsg = errmsg & "Error checking cache." & vbCrLf
        '            c_dr.Close()
        '            results = "Failure"
        '        End If
        '        c_dr.Close()
        '    Catch oBug As Exception
        '        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error checking cache: " & oBug.ToString)
        '        results = "Failure"
        '        Try
        '            SqlS = "IF OBJECT_ID ('cache.dbo.[CACHE]', 'U') IS NULL " & _
        '                "CREATE TABLE cache.dbo.[CACHE]( " & _
        '                "[CRSE_ID] [nvarchar](15) NOT NULL, " & _
        '                "[ASSET] [nvarchar](100) NOT NULL, " & _
        '                "[CREATED] [datetime] NOT NULL, " & _
        '                "[LAST_UPD] [datetime] NOT NULL)"
        '            If Debug = "Y" Then mydebuglog.Debug("  Create cache database table: " & SqlS & vbCrLf)
        '            c_cmd.CommandText = SqlS
        '            returnv = c_cmd.ExecuteNonQuery()
        '        Catch ex3 As Exception
        '            If Debug = "Y" Then mydebuglog.Debug("  Error creating cache database table: " & ex3.Message & vbCrLf)
        '        End Try
        '    End Try
        'End If

        ' -----
        ' Generate cache filename
        BinaryFile = SaveDest & "\" & ItemName
        BinaryFile = BinaryFile.Replace(mypath, "")
        If Debug = "Y" Then mydebuglog.Debug("  Cache name: " & BinaryFile & vbCrLf)
        ' If cache hit, check to see if the file exists
        'If CacheHit = 1 Then
        '    If Not My.Computer.FileSystem.FileExists(BinaryFile) Then
        '        If Debug = "Y" Then mydebuglog.Debug("   !! marked as cached but file not found in filesystem." & vbCrLf)
        '        CacheHit = 0
        '    End If
        'End If

        ' **** Construct Query string DMS document ***********
        ' Query DMS
        SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.last_upd, v.minio_flg, v.row_id, D.row_id " &
            "FROM DMS.dbo.Documents D WITH (NOLOCK) " &
            "LEFT OUTER JOIN DMS.dbo.Document_Versions v WITH (NOLOCK) ON v.row_id=D.last_version_id " &
            "LEFT OUTER JOIN DMS.dbo.Document_Associations DA WITH (NOLOCK) on DA.doc_id=D.row_id " &
            "LEFT OUTER JOIN DMS.dbo.Document_Categories DC WITH (NOLOCK) ON DC.doc_id=D.row_id " &
            "WHERE D.row_id IS NOT NULL AND D.deleted IS NULL AND DC.cat_id=188 " &
            "AND DA.fkey='" & CRSE_ID & "' AND lower(D.dfilename)='" & ItemName.ToLower & "' " &
            "ORDER BY D.last_upd DESC"
        ' *********   ************   **********
        ' **** Construct Query string for last update check  ***********
        Dim upt_SqlStr As String = ""
        upt_SqlStr = "Select TOP 1 v.last_upd " & SqlS.Substring(SqlS.IndexOf("FROM "))
        ' ******************************************'
        '***** Check cached object; 3/2/17; Ren Hou;  ****'
        Dim last_upt As Date = Today.AddYears(-50)
        ' Check to see if the document is in the in-memory cache
        cacheItemName = BinaryFile
        Dim docNotInDB As Boolean = 0
        If Not IsNothing(filecache(cacheItemName)) Then
            'Check if the cached item need to be renewed;
            Try
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check last updated date for xml document: " & vbCrLf & upt_SqlStr)
                d_cmd.CommandText = upt_SqlStr
                last_upt = d_cmd.ExecuteScalar()
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
                results = "Failure"
            End Try

            If last_upt = Today.AddYears(-50) Then  'document no longer exists in the database
                docNotInDB = True
                filecache.Remove(cacheItemName)
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   " & ItemName & " is not found in database! " & vbCrLf)
                'results = "Failure"
            ElseIf last_upt > TryCast(filecache(cacheItemName), HciDMSDocument).UpdateDate Then
                'Remove if the update_date on the cache is before the last updtaed date on DB record.
                filecache.Remove(cacheItemName)
                mydebuglog.Debug(vbCrLf & "  Cached object " & cacheItemName & " expired.")
            End If
        End If
        '****  ********  ********  ********  ********  ****'

        'Load content of cached object
        If filecache(cacheItemName) Is Nothing Then
            fileContents = Nothing
        Else
            Dim tmpObj As Object
            tmpObj = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
            ReDim fileContents(tmpObj.Length)
            fileContents = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
            CacheHit = 1
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & " Retrieved Cached object " & cacheItemName & " !")
        End If

        If (IsNothing(fileContents) Or Debug = "R") And docNotInDB = False Then
            ' ============================================
            ' Get DMS record containing item 
            '  If the cache was hit, make sure the entry is current, otherwise restore it
            ' -----
            If Debug = "Y" Then mydebuglog.Debug("  Checking item with DMS query: " & SqlS)
            Try
                d_cmd.CommandText = SqlS
                d_dr = d_cmd.ExecuteReader()
                If Not d_dr Is Nothing Then
                    While d_dr.Read()
                        Try
                            d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                            minio_flg = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                            d_verid = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                            d_doc_id = Trim(CheckDBNull(d_dr(5), enumObjectType.StrType))
                            dLastUpd = d_dr(2)
                            If Debug = "Y" Then mydebuglog.Debug("  > Record found on query:  d_doc_id=" & d_doc_id & ",  d_verid=" & d_verid & ",  d_dsize=" & d_dsize & ",  minio_flg=" & minio_flg & ",  dLastUpd=" & Format(dLastUpd))

                            If minio_flg = "Y" Then
                                If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Minio")
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
                                        ReDim outbyte(Val(retval - 1))
                                        Dim intval As Integer
                                        For i = 0 To retval - 1
                                            intval = mobj2.ResponseStream.ReadByte()
                                            If intval < 255 And intval > 0 Then
                                                outbyte(i) = intval
                                            End If
                                            If intval = 255 Then outbyte(i) = 255
                                            If intval < 0 Then
                                                outbyte(i) = 0
                                                If Debug = "Y" Then mydebuglog.Debug("  >  .. " & i.ToString & "   intval: " & intval.ToString)
                                            End If
                                        Next
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
                                If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Document_Versions")
                                ' Get binary and attach to the object outbyte if found, not cached or updated recently
                                '   retval will be "0" if this is not the case
                                If d_dsize <> "" And (CacheHit = 0 Or dLastUpd > LastUpd) Then
                                    ReDim outbyte(Val(d_dsize) - 1)
                                    startIndex = 0
                                    retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                End If
                            End If

                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting item. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Error getting item." & vbCrLf
                    d_dr.Close()
                    results = "Failure"
                End If
                d_dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting item: " & oBug.ToString)
                results = "Failure"
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("   ... Reported d_dsize: " & d_dsize & " : Non-cached DMS doc found: " & Str(retval))
            End If

            '***** Set cache object; 2/28/17; Ren Hou;  ****
            Dim policy As New CacheItemPolicy()
            policy.SlidingExpiration = TimeSpan.FromDays(CDbl(dms_cache_age))
            filecache.Set(cacheItemName, New HciDMSDocument(dLastUpd, outbyte), policy)
            ReDim fileContents(outbyte.Length)
            fileContents = outbyte
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Caching DMS doc to key: " & cacheItemName)
            '****  ****

            '' -----            
            '' Retrieve item and store-to/replace-in the filesystem if found, 
            ''  not cached yet or updated recently
            'If retval > 0 And (CacheHit = 0 Or dLastUpd > LastUpd) Then
            '    If Debug = "Y" Then mydebuglog.Debug("  Storing to filesystem: " & BinaryFile)
            '    If (My.Computer.FileSystem.FileExists(BinaryFile)) Then Kill(BinaryFile)
            '    Try
            '        bfs = New FileStream(BinaryFile, FileMode.Create, FileAccess.Write)
            '        bw = New BinaryWriter(bfs)
            '        bw.Write(outbyte)
            '        bw.Flush()
            '        bw.Close()
            '        bfs.Close()
            '    Catch ex As Exception
            '        errmsg = errmsg & "Unable to write the file to a temp file." & ex.ToString & vbCrLf
            '        results = "Failure"
            '        retval = 0
            '    End Try
            '    bfs = Nothing
            '    bw = Nothing
            'End If

            '' ============================================
            ''Double check to make sure file exists at this point
            ''  Remove from cache if applicable
            'If Not My.Computer.FileSystem.FileExists(BinaryFile) Then
            '    If Debug = "Y" Then mydebuglog.Debug("  !! The file was not properly stored ")
            '    errmsg = errmsg & "Unable to find file: " & BinaryFile & vbCrLf
            '    results = "Failure"
            '    retval = 0

            '    ' Remove from cache
            '    If CacheHit = 1 Then
            '        SqlS = "DELETE FROM CACHE WHERE CRSE_ID='" & CRSE_ID & "' AND ASSET='" & ItemName & "'"
            '        If Debug = "Y" Then mydebuglog.Debug("  Removing cache entry: " & SqlS)
            '        Try
            '            c_cmd.CommandText = SqlS
            '            returnv = c_cmd.ExecuteNonQuery()
            '        Catch ex As Exception
            '            errmsg = errmsg & "Unable to delete from cache." & ex.ToString & vbCrLf & " with query " & SqlS & vbCrLf
            '            results = "Failure"
            '            retval = 0
            '        End Try
            '    End If
            '    GoTo CleanCache
            'End If

            '' ============================================
            '' Update cache 
            ''   If made it to this point then the file must exist
            'If CacheHit = 1 Then
            '    ' Update cache since it was already there
            '    SqlS = "UPDATE [cache].dbo.[CACHE] SET LAST_UPD=GETDATE() " & _
            '        "WHERE CRSE_ID='" & CRSE_ID & "' AND ASSET='" & ItemName & "'"
            '    If Debug = "Y" Then mydebuglog.Debug("  Updating cache entry: " & SqlS)
            'Else
            '    ' Add to cache since the file is new
            '    SqlS = "INSERT [cache].dbo.[CACHE] (CRSE_ID, ASSET, CREATED, LAST_UPD) " & _
            '    "VALUES ('" & CRSE_ID & "','" & ItemName & "', GETDATE(), GETDATE())"
            '    If Debug = "Y" Then mydebuglog.Debug("  Inserting cache entry: " & SqlS)
            'End If
            '' Execute query
            'Try
            '    c_cmd.CommandText = SqlS
            '    returnv = c_cmd.ExecuteNonQuery()
            '    If returnv = 0 Then
            '        errmsg = errmsg & "Unable to write to cache."
            '    End If
            'Catch ex As Exception
            '    errmsg = errmsg & "Unable to write to cache." & ex.ToString & vbCrLf & " with query " & SqlS & vbCrLf
            '    results = "Failure"
            '    retval = 0
            'End Try

        End If
CleanCache:
        '        ' ============================================
        '        ' Cleanup cache as required
        '        SqlS = "SELECT CRSE_ID, ASSET FROM CACHE WHERE LAST_UPD<dateadd(dy,-" & dms_cache_age & ",getdate())"
        '        If Debug = "Y" Then mydebuglog.Debug("  Locate old cache entries to remove: " & SqlS)
        '        Try
        '            c_cmd.CommandText = SqlS
        '            c_dr = c_cmd.ExecuteReader()
        '            If Not c_dr Is Nothing Then
        '                While c_dr.Read()
        '                    Try
        '                        dCrseId = Trim(CheckDBNull(c_dr(0), enumObjectType.StrType))
        '                        dAsset = Trim(CheckDBNull(c_dr(1), enumObjectType.StrType))
        '                        If dCrseId <> "" And dAsset <> "" Then
        '                            SqlS = "DELETE FROM CACHE WHERE ASSET='" & dAsset & "' AND CRSE_ID='" & dCrseId & "'"
        '                            ' Remove cache entry
        '                            If Debug = "Y" Then mydebuglog.Debug("  Remove cache entry: " & SqlS)
        '                            c_cmd2.CommandText = SqlS
        '                            returnv = c_cmd2.ExecuteNonQuery()
        '                            If returnv <> 0 Then
        '                                killcount = killcount + 1
        '                                dFileName = SaveDest & "\" & dAsset
        '                                If Debug = "Y" Then mydebuglog.Debug("   Removing file: " & dFileName)
        '                                Try
        '                                    Kill(dFileName)
        '                                Catch ex As Exception
        '                                    'errmsg = errmsg & "Unable to remove file: " & dFileName
        '                                    If Debug = "Y" Then mydebuglog.Debug("     Unable to remove file: " & dFileName)
        '                                End Try
        '                            Else
        '                                errmsg = errmsg & "Unable to delete entry: " & dCrseId & "/" & dAsset
        '                                If Debug = "Y" Then mydebuglog.Debug("     Unable to delete entry: " & dCrseId & "/" & dAsset)
        '                            End If
        '                        End If
        '                    Catch ex As Exception
        '                        results = "Failure"
        '                        errmsg = errmsg & "Error checking cache for cleanup. " & ex.ToString & vbCrLf
        '                        GoTo CloseOut
        '                    End Try
        '                End While
        '            Else
        '                errmsg = errmsg & "Error checking cache for cleanup." & vbCrLf
        '                c_dr.Close()
        '                results = "Failure"
        '            End If
        '            c_dr.Close()
        '        Catch oBug As Exception
        '            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error checking cache: " & oBug.ToString)
        '            results = "Failure"
        '        End Try
        '        If Debug = "Y" Then mydebuglog.Debug("   Cache entries removed: " & killcount.ToString)

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
        If Trim(errmsg) <> "" Then myeventlog.Error("GetHMedia : Error: " & Trim(errmsg))
        If CacheHit = 1 Then
            If Debug <> "T" Then myeventlog.Info("GetHMedia : Results: " & results & " for CACHED file " & ItemName & ", with CrseId " & CRSE_ID)
        Else
            If Debug <> "T" Then myeventlog.Info("GetHMedia : Results: " & results & " for file " & ItemName & ", minio: " & minio_flg & ", docid: " & d_doc_id & ", verid: " & d_verid & ", CrseId " & CRSE_ID)
        End If
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If CacheHit = 1 Then
                    mydebuglog.Debug("Results: " & results & " for CACHED file " & ItemName & ", with CrseId " & CRSE_ID & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
                Else
                    mydebuglog.Debug("Results: " & results & " for file " & ItemName & ", minio: " & minio_flg & ", docid: " & d_doc_id & ", verid: " & d_verid & ", with CrseId " & CRSE_ID & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
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
        ' Return item
        'If (My.Computer.FileSystem.FileExists(BinaryFile)) And Debug <> "T" Then
        Try
            'bfs = File.Open(BinaryFile, FileMode.Open, FileAccess.Read)
            'Dim lngLen As Long = bfs.Length
            'ReDim outbyte(CInt(lngLen - 1))
            'bfs.Read(outbyte, 0, CInt(lngLen))
            'Return outbyte
            Return fileContents
        Catch exp As Exception
            Return Nothing
        Finally
            'bfs.Close()
            'bfs = Nothing
            outbyte = Nothing
            fileContents = Nothing
        End Try
        'Else
        'Return Nothing
        'End If
    End Function

    <WebMethod(Description:="Retrieve any specified item in domain")> _
    Public Function GetDImage(ByVal Domain As String, ByVal PublicKey As String, ByVal ItemName As String, ByVal Debug As String) As Byte()

        ' This function locates the specified item, and returns it to the calling system as a binary

        ' The input parameters are as follows:
        '
        '   Domain      - The CX_SUBSCRIPTION.DOMAIN of the requestor. 
        '   PublicKey   - Base64 encoded, reversed "DMS.Documents.row_id"
        '   ItemName	- The "DMS.Documents.dfilename" of the item to be retrieved, or if "default.jpg",
        '                   the first associated item in the category "Images" will be returned.
        '   Debug	    - "Y", "N" or "T"
        '   

        ' web.config Parameters used:
        '   hcidb           - connection string to hcidb1.siebeldb database
        '   dms        	    - connection string to DMS.dms database
        '   cache           - connection string to cache.sdf database

        ' Variables
        Dim results, temp As String
        Dim mypath, errmsg, logging As String
        Dim DecodedPublicKey, ValidatedPublicKey As String

        ' Database declarations
        Dim SqlS As String
        Dim returnv As Integer
        Dim DomainGID As String             ' Domain's DMS.Groups.row_id

        ' Cache database declarations
        'Dim c_ConnS As String
        Dim CacheHit As Integer
        Dim cItemName As String
        Dim cLastUpd As DateTime
        Dim dDomain, dKey, dExt, dItemName, dFileName As String
        'Dim cExt, cFileName As String

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String
        Dim dms_cache_age As String

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Minio declarations
        Dim minio_flg, d_verid As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' File handling declarations
        'Dim bfs As FileStream
        'Dim bw As BinaryWriter
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim d_dsize, d_doc_id, SaveDest, d_item_name, d_access, old_row_id As String
        Dim d_last_upd As DateTime
        Dim d_ext As String
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
        DecodedPublicKey = ""
        ValidatedPublicKey = ""
        cItemName = ""
        d_ext = ""
        d_dsize = ""
        d_doc_id = ""
        d_item_name = ""
        d_access = ""
        dDomain = ""
        dKey = ""
        dExt = ""
        dItemName = ""
        dFileName = ""
        BinaryFile = ""
        DomainGID = ""
        'c_ConnS = ""
        d_ConnS = ""
        killcount = 0
        'Debug = "Y"
        minio_flg = ""
        d_verid = ""

        ' ============================================
        ' Get parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "T" Then
            Domain = "TIPS"
            PublicKey = "xkTN2ETO"
            DecodedPublicKey = "916591"
            ItemName = "2.C.5.d.jpg"
        Else
            Domain = Trim(HttpUtility.UrlEncode(Domain))
            If InStr(Domain, "%") > 0 Then Domain = Trim(HttpUtility.UrlDecode(Domain))
            If InStr(Domain, "%") > 0 Then Domain = Trim(Domain)
            Domain = EncodeParamSpaces(Domain)

            PublicKey = Trim(HttpUtility.UrlEncode(PublicKey))
            If InStr(PublicKey, "%") > 0 Then PublicKey = Trim(HttpUtility.UrlDecode(PublicKey))
            If InStr(PublicKey, "%") > 0 Then PublicKey = Trim(PublicKey)
            DecodedPublicKey = FromBase64(ReverseString(PublicKey))

            If InStr(ItemName, "%") > 0 Then ItemName = Trim(HttpUtility.UrlDecode(ItemName))
        End If

        ' ============================================
        ' Get system defaults
        Try
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIBDB;uid=DMS;pwd=5241200;database=DMS"
            'c_ConnS = "Server=(LocalDB)\MSSQLLocalDB;MultipleActiveResultSets=True;Integrated Security=true;AttachDbFileName=" & mypath & "cachedb\cache.mdf"  'Switched to SQL Server Express LocalDB 2014
            dms_cache_age = Trim(System.Configuration.ConfigurationManager.AppSettings("dmscacheage"))
            If dms_cache_age = "" Or Not IsNumeric(dms_cache_age) Then dms_cache_age = "30"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetDImage_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y" Else Debug = "N"
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
            logfile = "C:\Logs\GetDImage.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  Domain: " & Domain)
                mydebuglog.Debug("  PublicKey: " & PublicKey)
                mydebuglog.Debug("  Decoded PublicKey: " & DecodedPublicKey)
                mydebuglog.Debug("  ItemName: " & ItemName)
                mydebuglog.Debug("  Appsetting dms_cache_age: " & dms_cache_age & vbCrLf)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(ItemName) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No item specified. "
            GoTo CloseOut2
        End If
        If Trim(Domain) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No domain specified. "
            GoTo CloseOut2
        Else '1/2/2018; Ren Hou; Added validation for Domain paraneter
            Dim m As Match = Regex.Match(Domain, "^[A-Z]{1,10}$", RegexOptions.IgnoreCase)
            If Not (m.Success) Then
                results = "Failure"
                errmsg = errmsg & vbCrLf & "Incorrect domain format. "
                GoTo CloseOut2
            End If
        End If
        If Not IsNumeric(DecodedPublicKey) Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Incorrect key specified. "
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open SQL Server database connections
        errmsg = OpenDBConnection(d_ConnS, d_con, d_cmd)    ' dms
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        '' ============================================
        '' Open SQL Server LocalDB database connection for local cache
        'AppDomain.CurrentDomain.SetData("SQLServerCompactEditionUnderWebHosting", True)
        'Dim c_con As New SqlConnection(c_ConnS)
        'Dim c_cmd As New SqlCommand()
        'Dim c_dr As SqlDataReader
        'If Not (My.Computer.FileSystem.FileExists(mypath & "cachedb\cache.mdf")) Then
        '    ' Create the database if necessary
        '    Try
        '        'Dim SQLCEDB As New SqlCeEngine(c_ConnS)
        '        'SQLCEDB.CreateDatabase()
        '        'c_con = New SqlConnection(c_ConnS)
        '        c_con.ConnectionString = c_ConnS.Substring(0, c_ConnS.IndexOf(";AttachDbFileName", 0) + 1)  'connect to the localDB instance instead of 'cache' database
        '        c_cmd.Connection = c_con
        '        c_con.Open()
        '        c_cmd.CommandText = "IF DB_ID('cache') IS NULL CREATE DATABASE cache ON (NAME = cache, FILENAME = '" & mypath & "cachedb\cache.mdf')"
        '        If Debug = "Y" Then mydebuglog.Debug("  Create cache database query: " & c_cmd.CommandText & vbCrLf)
        '        returnv = c_cmd.ExecuteNonQuery()
        '        c_con.Close()
        '    Catch ex2 As Exception
        '        errmsg = errmsg & vbCrLf & "Unable to create database: " & ex2.Message
        '        results = "Failure"
        '        GoTo CloseOut2
        '    End Try
        'End If

        'Try
        '    c_con.ConnectionString = c_ConnS
        '    c_cmd.Connection = c_con
        '    c_con.Open()
        '    ' Create Tables if there are not there already
        '    SqlS = "IF OBJECT_ID('D_CACHE') IS NULL CREATE TABLE [cache].dbo.[D_CACHE]( " & _
        '        "[DOMAIN] [nvarchar](15) NOT NULL, " & _
        '        "[EXT] [nvarchar](15) NOT NULL, " & _
        '        "[KEY] [numeric](18,0) NOT NULL, " & _
        '        "[CREATED] [datetime] NOT NULL, " & _
        '        "[LAST_UPD] [datetime] NOT NULL)"
        '    If Debug = "Y" Then mydebuglog.Debug("  Create Item cache database query: " & SqlS & vbCrLf)
        '    c_cmd.CommandText = SqlS
        '    returnv = c_cmd.ExecuteNonQuery()

        '    SqlS = "IF OBJECT_ID('D_GROUP') IS NULL CREATE TABLE [cache].dbo.[D_GROUP]( " & _
        '        "[DOMAIN] [nvarchar](15) NOT NULL, " & _
        '        "[UGA] [nvarchar](15) NOT NULL)"
        '    If Debug = "Y" Then mydebuglog.Debug("  Create UGA cache database query: " & SqlS & vbCrLf)
        '    c_cmd.CommandText = SqlS
        '    returnv = c_cmd.ExecuteNonQuery()
        '    c_con.Close()
        'Catch ex3 As Exception
        '    errmsg = errmsg & vbCrLf & "Unable to create cache table: " & ex3.Message
        '    results = "Failure"
        '    GoTo CloseOut2
        'End Try

        '' Open cache database
        'Try
        '    'c_con = New SqlConnection(c_ConnS)
        '    c_con.Open()
        '    If Not c_con Is Nothing Then
        '        Try
        '            c_cmd = New SqlCommand(SqlS, c_con)
        '        Catch ex2 As Exception
        '            errmsg = errmsg & vbCrLf & "Unable to open cache database: " & ex2.Message
        '            results = "Failure"
        '            GoTo CloseOut2
        '        End Try
        '    Else
        '        errmsg = errmsg & vbCrLf & "Unable to open cache database"
        '        results = "Failure"
        '        GoTo CloseOut2
        '    End If
        'Catch ex As Exception
        '    errmsg = errmsg & vbCrLf & "Unable to open cache database: " & ex.Message
        '    results = "Failure"
        '    GoTo CloseOut2
        'End Try

        '' ============================================
        '' Locate Group Access Information for query
        '' Try local cache
        'If Not c_cmd Is Nothing Then
        '    SqlS = "SELECT [UGA] " & _
        '        "FROM [D_GROUP] " & _
        '        "WHERE [DOMAIN]='" & Domain & "'"
        '    If Debug = "Y" Then mydebuglog.Debug("  Check domain UGA cache: " & SqlS)
        '    Try
        '        c_cmd.CommandText = SqlS
        '        c_dr = c_cmd.ExecuteReader()
        '        If Not c_dr Is Nothing Then
        '            While c_dr.Read()
        '                Try
        '                    DomainGID = c_dr(0)
        '                    CacheHit = 1
        '                    If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  DomainGID=" & DomainGID)
        '                Catch ex As Exception
        '                    results = "Failure"
        '                    errmsg = errmsg & "Error checking cache. " & ex.ToString & vbCrLf
        '                    GoTo CloseOut
        '                End Try
        '            End While
        '        Else
        '            If Debug = "Y" Then mydebuglog.Debug("  > Did not find cached domain group")
        '        End If
        '        c_dr.Close()
        '    Catch oBug As Exception
        '        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error checking cache: " & oBug.ToString)
        '        results = "Failure"
        '    End Try
        'Else
        '    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Database error.")
        '    results = "Failure"
        '    GoTo CloseOut
        'End If

        ' -----
        ' Try DMS
        If DomainGID = "" Then
            If Not d_cmd Is Nothing Then
                ' -----
                ' Query DMS
                Try
                    SqlS = "SELECT UGA.row_id " & _
                       "FROM DMS.dbo.User_Group_Access UGA " & _
                        "INNER JOIN DMS.dbo.Groups G on UGA.access_id=G.row_id " & _
                       "WHERE UGA.type_id='G' AND G.name='" & Domain & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Get DOMAIN UGA: " & SqlS)
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                                DomainGID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "The record was not found." & vbCrLf
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                    results = "Failure"
                End Try
                If Debug = "Y" Then mydebuglog.Debug("   ... DomainGID: " & DomainGID)

                '' -----
                '' Save the found GID in the local cache
                'If DomainGID <> "" Then
                '    SqlS = "INSERT [D_GROUP] ([DOMAIN],[UGA]) " & _
                '        "VALUES('" & Domain & "','" & DomainGID & "')"
                '    If Debug = "Y" Then mydebuglog.Debug("  INSERT UGA cache database query: " & SqlS & vbCrLf)
                '    Try
                '        c_cmd.CommandText = SqlS
                '        returnv = c_cmd.ExecuteNonQuery()
                '    Catch ex As Exception
                '        errmsg = errmsg & "Unable to insert into cache." & ex.ToString & vbCrLf & " with query " & SqlS & vbCrLf
                '        results = "Failure"
                '        retval = 0
                '    End Try
                'End If
            Else
                results = "Failure"
                GoTo CloseOut
            End If
        End If
        If Trim(DomainGID) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No domain access found."
            GoTo CloseOut2
        End If

        ' ============================================
        ' Create output directory for ItemName caching
        SaveDest = mypath & "domain_temp\" & Domain
        'Try
        '    Directory.CreateDirectory(SaveDest)
        'Catch
        'End Try
        'If Debug = "Y" Then mydebuglog.Debug("  Caching to: " & SaveDest & vbCrLf)

        '' ============================================
        '' Check to see if and when we already cached this item
        'CacheHit = 0
        'If Not c_cmd Is Nothing Then
        '    SqlS = "SELECT [LAST_UPD],[EXT] " & _
        '        "FROM [D_CACHE] " & _
        '        "WHERE [DOMAIN]='" & Domain & "' AND [KEY]=" & DecodedPublicKey
        '    If Debug = "Y" Then mydebuglog.Debug("  Check cache for item: " & SqlS)
        '    Try
        '        c_cmd.CommandText = SqlS
        '        c_dr = c_cmd.ExecuteReader()
        '        If Not c_dr Is Nothing Then
        '            While c_dr.Read()
        '                Try
        '                    cLastUpd = c_dr(0)
        '                    cExt = Trim(CheckDBNull(c_dr(1), enumObjectType.StrType))
        '                    cFileName = SaveDest & "\" & DecodedPublicKey & "." & cExt
        '                    CacheHit = 1
        '                    If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  LastUpd=" & Format(cLastUpd) & "  CacheHit=" & CacheHit)

        '                    ' If cache hit, check to see if the file exists
        '                    If CacheHit = 1 Then
        '                        If Not My.Computer.FileSystem.FileExists(cFileName) Then
        '                            If Debug = "Y" Then mydebuglog.Debug("   !! marked as cached but file not found in filesystem." & vbCrLf)
        '                            CacheHit = 0
        '                        End If
        '                    End If
        '                Catch ex As Exception
        '                    results = "Failure"
        '                    errmsg = errmsg & "Error checking cache. " & ex.ToString & vbCrLf
        '                End Try
        '            End While
        '        Else
        '            errmsg = errmsg & "Error checking cache." & vbCrLf
        '            c_dr.Close()
        '            results = "Failure"
        '        End If
        '        c_dr.Close()
        '    Catch oBug As Exception
        '        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error checking cache: " & oBug.ToString)
        '        results = "Failure"
        '    End Try
        'End If



        ' **** Construct Query string DMS document ***********
        ' Query DMS
        SqlS = "SELECT TOP 1 D.dfilename, V.dimage, D.row_id, V.last_upd, V.dsize, DT.extension, DU.access_type, D.old_row_id, V.minio_flg, V.row_id " &
            "FROM DMS.dbo.Documents D  " &
            "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id " &
            "LEFT OUTER JOIN DMS.dbo.Document_Users DU on D.row_id=DU.doc_id " &
            "LEFT OUTER JOIN DMS.dbo.Document_Types DT on D.data_type_id=DT.row_id " &
            "WHERE DU.user_access_id=" & DomainGID & " AND (D.row_id=" & DecodedPublicKey & ") AND D.dfilename='" & ItemName & "'"
        ' *********   ************   **********
        ' **** Construct Query string for last update check  ***********
        Dim upt_SqlStr As String = ""
        upt_SqlStr = "Select TOP 1 V.last_upd " & SqlS.Substring(SqlS.IndexOf("FROM "))
        ' ******************************************'
        '***** Check cached object; 3/2/17; Ren Hou;  ****'
        Dim last_upt As Date = Today.AddYears(-50)
        ' Check to see if the document is in the in-memory cache
        BinaryFile = SaveDest & "\" & DecodedPublicKey & "-" & ItemName
        BinaryFile = BinaryFile.Replace(mypath, "")
        cacheItemName = BinaryFile
        If Not IsNothing(filecache(cacheItemName)) Then
            'Check if the cached item need to be renewed;
            Try
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check last updated date for xml document: " & vbCrLf & SqlS)
                d_cmd.CommandText = upt_SqlStr
                last_upt = d_cmd.ExecuteScalar()
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
                results = "Failure"
            End Try

            If last_upt > TryCast(filecache(cacheItemName), HciDMSDocument).UpdateDate Then
                'Remove if the update_date on the cache is before the last updtaed date on DB record.
                filecache.Remove(cacheItemName)
                mydebuglog.Debug(vbCrLf & "  Cached object " & cacheItemName & " expired.")
            End If
        End If
        '****  ********  ********  ********  ********  ****'

        'Load content of cached object
        If filecache(cacheItemName) Is Nothing Then
            fileContents = Nothing
        Else
            Dim tmpObj As Object
            tmpObj = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
            ReDim fileContents(tmpObj.Length)
            fileContents = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
            CacheHit = 1
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & " Retrieved Cached object " & cacheItemName & vbCrLf)
        End If

        If IsNothing(fileContents) Or Debug = "R" Then
            ' ============================================
            ' Get the item and meta data from the DMS
            If Debug = "Y" Then mydebuglog.Debug("  Looking in DMS for: " & ItemName & vbCrLf)
            If Not d_cmd Is Nothing Then
                If Debug = "Y" Then mydebuglog.Debug("  Get item information: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                d_item_name = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(2), enumObjectType.StrType))
                                d_last_upd = CheckDBNull(d_dr(3), enumObjectType.DteType)
                                d_dsize = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                d_ext = Trim(CheckDBNull(d_dr(5), enumObjectType.StrType))
                                d_access = Trim(CheckDBNull(d_dr(6), enumObjectType.StrType))
                                old_row_id = Trim(CheckDBNull(d_dr(7), enumObjectType.StrType))
                                minio_flg = Trim(CheckDBNull(d_dr(8), enumObjectType.StrType))
                                d_verid = Trim(CheckDBNull(d_dr(9), enumObjectType.StrType))

                                ' Verify domain read access to the document - meaning its public to that domain
                                If InStr(UCase(d_access), "R") = 0 Then
                                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  No access to this item based on domain security.")
                                    results = "Failure"
                                    GoTo CloseOut
                                End If

                                If minio_flg = "Y" Then
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Minio")
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
                                            ReDim outbyte(Val(retval - 1))
                                            Dim intval As Integer
                                            For i = 0 To retval - 1
                                                intval = mobj2.ResponseStream.ReadByte()
                                                If intval < 255 And intval > 0 Then
                                                    outbyte(i) = intval
                                                End If
                                                If intval = 255 Then outbyte(i) = 255
                                                If intval < 0 Then
                                                    outbyte(i) = 0
                                                    If Debug = "Y" Then mydebuglog.Debug("  >  .. " & i.ToString & "   intval: " & intval.ToString)
                                                End If
                                            Next
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
                                    ' Get binary and attach to the object outbyte if found, not cached or updated recently
                                    '   retval will be "0" if this is not the case
                                    If d_dsize <> "" And (CacheHit = 0 Or d_last_upd > cLastUpd) Then
                                        If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Document_Versions")
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        Try
                                            retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                        Catch ex As Exception
                                            results = "Failure"
                                            errmsg = errmsg & "Error getting item. " & ex.ToString & vbCrLf
                                            GoTo CloseOut
                                        End Try
                                    End If
                                End If

                                If Debug = "Y" Then
                                    mydebuglog.Debug("  > Found record on query.  Extension: " & d_ext)
                                    mydebuglog.Debug("                            Item Name: " & d_item_name)
                                    mydebuglog.Debug("                            Doc Id: " & d_doc_id)
                                    mydebuglog.Debug("                            Minio: " & minio_flg)
                                    mydebuglog.Debug("                            Version Id: " & d_verid)
                                    mydebuglog.Debug("                            Size: " & d_dsize)
                                    mydebuglog.Debug("                            Last Updated: " & Format(d_last_upd))
                                    mydebuglog.Debug("                            Bytes retrieved: " & retval.ToString)
                                End If

                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error getting item. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                        d_dr.Close()
                    Else
                        errmsg = errmsg & "Error getting item." & vbCrLf
                        results = "Failure"
                    End If
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting item: " & oBug.ToString)
                    results = "Failure"
                    GoTo CloseOut
                End Try

                ' Verify file is the same
                If Trim(LCase(d_item_name)) <> Trim(LCase(ItemName)) Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Unable to verify identity when comparing " & d_item_name & " to " & ItemName)
                    results = "Failure"
                    GoTo CloseOut
                End If

                ' Generate cache filename
                '  The formula for this [Document Id] + [Extension]
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Computed cache filename: " & BinaryFile & vbCrLf)

                If Debug = "Y" Then
                    mydebuglog.Debug("   ... Reported d_dsize: " & d_dsize & " : Non-cached or superceded bytes found: " & Str(retval))
                End If

            Else
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Database error.")
                results = "Failure"
                GoTo CloseOut
            End If

            '***** Set cache object; 2/28/17; Ren Hou;  ****
            Dim policy As New CacheItemPolicy()
            policy.SlidingExpiration = TimeSpan.FromDays(CDbl(dms_cache_age))
            filecache.Set(cacheItemName, New HciDMSDocument(d_last_upd, outbyte), policy)
            ReDim fileContents(outbyte.Length)
            fileContents = outbyte
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Caching DMS doc to key: " & cacheItemName & vbCrLf)
            '****  ****
        End If
        '' ============================================
        '' Retrieve item and store-to/replace-in the filesystem if found, 
        ''  not cached yet or updated recently
        'If retval > 0 And (CacheHit = 0 Or d_last_upd > cLastUpd) Then
        '    If Debug = "Y" Then mydebuglog.Debug("  Storing to filesystem: " & BinaryFile)
        '    If (My.Computer.FileSystem.FileExists(BinaryFile)) Then Kill(BinaryFile)
        '    Try
        '        bfs = New FileStream(BinaryFile, FileMode.Create, FileAccess.Write)
        '        bw = New BinaryWriter(bfs)
        '        bw.Write(outbyte)
        '        bw.Flush()
        '        bw.Close()
        '        bfs.Close()
        '    Catch ex As Exception
        '        errmsg = errmsg & "Unable to write the file to a temp file." & ex.ToString & vbCrLf
        '        results = "Failure"
        '        retval = 0
        '    End Try
        '    bfs = Nothing
        '    bw = Nothing
        'End If

        '' ============================================
        ''Double check to make sure file exists at this point
        ''  Remove from cache if applicable
        'If Not My.Computer.FileSystem.FileExists(BinaryFile) Then
        '    If Debug = "Y" Then mydebuglog.Debug("  !! The file was not properly stored ")
        '    errmsg = errmsg & "Unable to find file: " & BinaryFile & vbCrLf
        '    results = "Failure"
        '    retval = 0

        '    ' Remove from cache
        '    If CacheHit = 1 Then
        '        SqlS = "DELETE FROM [D_CACHE] WHERE [DOMAIN]='" & Domain & "' AND [KEY]=" & DecodedPublicKey
        '        If Debug = "Y" Then mydebuglog.Debug("  Removing cache entry: " & SqlS)
        '        Try
        '            c_cmd.CommandText = SqlS
        '            returnv = c_cmd.ExecuteNonQuery()
        '        Catch ex As Exception
        '            errmsg = errmsg & "Unable to delete from cache." & ex.ToString & vbCrLf & " with query " & SqlS & vbCrLf
        '            results = "Failure"
        '            retval = 0
        '        End Try
        '    End If
        '    GoTo CleanCache
        'End If

        '' ============================================
        '' Update cache 
        ''   If made it to this point then the file must exist
        'If CacheHit = 1 Then
        '    ' Update cache since it was already there
        '    SqlS = "UPDATE [D_CACHE] SET [LAST_UPD]=GETDATE() " & _
        '        "WHERE [DOMAIN]='" & Domain & "' AND [KEY]=" & DecodedPublicKey
        '    If Debug = "Y" Then mydebuglog.Debug("  Updating cache entry: " & SqlS)
        'Else
        '    ' Add to cache since the file is new
        '    SqlS = "INSERT [D_CACHE] ([DOMAIN], [KEY], [EXT], [CREATED], [LAST_UPD]) " & _
        '    "VALUES ('" & Domain & "'," & DecodedPublicKey & ", '" & d_ext & "', GETDATE(), GETDATE())"
        '    If Debug = "Y" Then mydebuglog.Debug("  Inserting cache entry: " & SqlS)
        'End If
        '' Execute query
        'Try
        '    c_cmd.CommandText = SqlS
        '    returnv = c_cmd.ExecuteNonQuery()
        '    If returnv = 0 Then
        '        errmsg = errmsg & "Unable to write to cache."
        '    End If
        'Catch ex As Exception
        '    errmsg = errmsg & "Unable to write to cache." & ex.ToString & vbCrLf & " with query " & SqlS & vbCrLf
        '    results = "Failure"
        '    retval = 0
        'End Try

CleanCache:
        '' ============================================
        '' Cleanup cache as required
        'SqlS = "DELETE FROM D_CACHE WHERE LAST_UPD<dateadd(dy,-" & dms_cache_age & ",getdate())"
        'If Debug = "Y" Then mydebuglog.Debug("  Locate old cache entries to remove: " & SqlS)
        'Try
        '    c_cmd.CommandText = SqlS
        '    c_cmd.ExecuteNonQuery()
        'Catch oBug As Exception
        '    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error deleting cache: " & oBug.ToString)
        '    results = "Failure"
        'End Try
        'If Debug = "Y" Then mydebuglog.Debug("   Cache entries removed from local CACHE table")


        'SqlS = "SELECT [DOMAIN], [KEY], [EXT] FROM [D_CACHE] WHERE [LAST_UPD]<dateadd(dy,-" & dms_cache_age & ",getdate())"
        'If Debug = "Y" Then mydebuglog.Debug("  Locate old cache entries to remove: " & SqlS)
        'Try
        '    c_cmd.CommandText = SqlS
        '    c_dr = c_cmd.ExecuteReader()
        '    If Not c_dr Is Nothing Then
        '        While c_dr.Read()
        '            Try
        '                dDomain = Trim(CheckDBNull(c_dr(0), enumObjectType.StrType))
        '                dKey = Trim(CheckDBNull(c_dr(1), enumObjectType.StrType))
        '                dExt = Trim(CheckDBNull(c_dr(2), enumObjectType.StrType))
        '                dItemName = dKey & "." & dExt
        '                If dDomain <> "" And dItemName <> "" Then
        '                    SqlS = "DELETE FROM [D_CACHE] WHERE [DOMAIN]='" & dDomain & "' AND [KEY]=" & dKey
        '                    ' Remove cache entry
        '                    If Debug = "Y" Then mydebuglog.Debug("  Remove cache entry: " & SqlS)
        '                    c_cmd.CommandText = SqlS
        '                    returnv = c_cmd.ExecuteNonQuery()
        '                    If returnv <> 0 Then
        '                        killcount = killcount + 1
        '                        dFileName = SaveDest & "\" & dItemName
        '                        If Debug = "Y" Then mydebuglog.Debug("   Removing file: " & dFileName)
        '                        Try
        '                            Kill(dFileName)
        '                        Catch ex As Exception
        '                            'errmsg = errmsg & "Unable to remove file: " & dFileName
        '                            If Debug = "Y" Then mydebuglog.Debug("     Unable to remove file: " & dFileName)
        '                        End Try
        '                    Else
        '                        errmsg = errmsg & "Unable to delete entry: " & dDomain & "/" & dItemName
        '                        If Debug = "Y" Then mydebuglog.Debug("     Unable to delete entry: " & dDomain & "/" & dItemName)
        '                    End If
        '                End If
        '            Catch ex As Exception
        '                results = "Failure"
        '                errmsg = errmsg & "Error checking cache for cleanup. " & ex.ToString & vbCrLf
        '                GoTo CloseOut
        '            End Try
        '        End While
        '    Else
        '        errmsg = errmsg & "Error checking cache for cleanup." & vbCrLf
        '        c_dr.Close()
        '        results = "Failure"
        '    End If
        '    c_dr.Close()
        'Catch oBug As Exception
        '    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error checking cache: " & oBug.ToString)
        '    results = "Failure"
        'End Try
        'If Debug = "Y" Then mydebuglog.Debug("   Cache entries removed: " & killcount.ToString)

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
        If Trim(errmsg) <> "" Then myeventlog.Error("GetDImage : Error: " & Trim(errmsg))
        If CacheHit = 1 Then
            If Debug <> "T" Then myeventlog.Info("GetDImage :Results: " & results & " for CACHED item " & ItemName & ", in " & Domain & " domain with cache filename " & DecodedPublicKey & "." & d_ext)
        Else
            If Debug <> "T" Then myeventlog.Info("GetDImage :Results: " & results & " for item " & ItemName & ", in " & Domain & " domain, minio: " & minio_flg & ", docid: " & d_doc_id & ", verid: " & d_verid & " with cache filename " & DecodedPublicKey & "." & d_ext)
        End If
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If CacheHit = 1 Then
                    mydebuglog.Debug("Results: " & results & " for CACHED item " & ItemName & ", in " & Domain & " domain with cache filename " & DecodedPublicKey & "." & d_ext & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
                Else
                    mydebuglog.Debug("Results: " & results & " for item " & ItemName & ", in " & Domain & " domain, minio: " & minio_flg & ", docid: " & d_doc_id & ", verid: " & d_verid & " with cache filename " & DecodedPublicKey & "." & d_ext & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
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
        ' Return item
        'If (My.Computer.FileSystem.FileExists(BinaryFile)) And Debug <> "T" Then
        Try
            'bfs = File.Open(BinaryFile, FileMode.Open, FileAccess.Read)
            'Dim lngLen As Long = bfs.Length
            'ReDim outbyte(CInt(lngLen - 1))
            'bfs.Read(outbyte, 0, CInt(lngLen))
            'Return outbyte
            Return fileContents
        Catch exp As Exception
            Return Nothing
        Finally
            'bfs.Close()
            'bfs = Nothing
            outbyte = Nothing
            fileContents = Nothing
        End Try
        'Else
        'Return Nothing
        'End If
    End Function

    <WebMethod(Description:="Secured retrieval of specified document in domain")> _
    Public Function GetMediaSecure(ByVal UserId As String, ByVal UserKey As String, ByVal Domain As String, ByVal PublicKey As String, ByVal Debug As String) As Byte()

        ' This function locates the specified item, and returns it to the calling system 

        ' The input parameters are as follows:
        '
        '   UserId      - The value of a cookie containing an unencoded S_CONTACT.X_REGISTRATION_NUM of the requestor.
        '
        '   UserKey     - Base64 encoded, reversed S_CONTACT.ROW_ID of the user entitled to use the document.
        '
        '   Domain      - The CX_SUBSCRIPTION.DOMAIN of the requestor. 
        '
        '   PublicKey   - Base64 encoded, reversed "DMS.Documents.row_id"
        '
        '   Debug	- "Y", "N" or "T"

        ' web.config Parameters used:
        '   hcidb           - connection string to hcidb1.siebeldb database
        '   dms        	    - connection string to DMS.dms database
        '   cache           - connection string to cache.sdf database

        ' Variables
        Dim results, temp As String
        Dim mypath, errmsg, logging As String
        Dim DecodedUserKey, DecodedPublicKey, ValidatedUserId As String
        Dim DomainGID As String             ' Domain's DMS.Groups.row_id

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

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String
        Dim dms_cache_age As String

        ' Logging declarationskk94pf
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

        ' Minio declarations
        Dim minio_flg, d_verid As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' File handling declarations
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim d_dsize, d_doc_id As String
        Dim killcount As Double
        Dim d_item_name, d_access As String
        Dim d_last_upd As DateTime
        Dim d_ext As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserKey = ""
        DecodedPublicKey = ""
        ValidatedUserId = ""
        d_dsize = ""
        BinaryFile = ""
        c_ConnS = ""
        dAsset = ""
        d_doc_id = ""
        dCrseId = ""
        dFileName = ""
        d_item_name = ""
        TypeTrans = ""
        DomainGID = ""
        killcount = 0
        temp = ""
        minio_flg = ""
        d_verid = ""
        d_ext = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If UserId = "" And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            ' This is a test document
            UserId = "27063600"
            UserKey = "TN0S20SM"
            PublicKey = "==wM5kDO4ETM"
            DecodedPublicKey = "1188993"
            Domain = "TIPS"
        Else
            UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            UserId = EncodeParamSpaces(UserId)

            If InStr(UserKey, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(Trim(UserKey)))
            If InStr(UserKey, "%") > 0 Then UserId = Trim(UserKey)
            Try
                DecodedUserKey = FromBase64(ReverseString(UserKey))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Unable to decode '" & ReverseString(UserKey) & "'"
            End Try

            Domain = Trim(HttpUtility.UrlEncode(Domain))
            If InStr(Domain, "%") > 0 Then Domain = Trim(HttpUtility.UrlDecode(Domain))
            If InStr(Domain, "%") > 0 Then Domain = Trim(Domain)
            Domain = EncodeParamSpaces(Domain)

            If InStr(PublicKey, "%") > 0 Then PublicKey = Trim(HttpUtility.UrlDecode(Trim(PublicKey)))
            If InStr(PublicKey, "%") > 0 Then PublicKey = Trim(PublicKey)
            Try
                DecodedPublicKey = FromBase64(ReverseString(PublicKey))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Unable to decode '" & ReverseString(PublicKey) & "'"
            End Try
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidbro").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb;ApplicationIntent=ReadOnly"
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIBDB;uid=DMS;pwd=5241200;database=DMS"
            'c_ConnS = "Server=(LocalDB)\MSSQLLocalDB;MultipleActiveResultSets=True;Integrated Security=true;AttachDbFileName=" & mypath & "cachedb\cache.mdf"  'Switched to SQL Server Express LocalDB 2014
            dms_cache_age = Trim(System.Configuration.ConfigurationManager.AppSettings("dmscacheage"))
            If dms_cache_age = "" Or Not IsNumeric(dms_cache_age) Then dms_cache_age = "30"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetMediaSecure_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
            If temp = "N" And Debug <> "T" Then Debug = "N"
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
            logfile = "C:\Logs\GetMediaSecure.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("GetMediaSecure Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  UserKey: " & UserKey)
                mydebuglog.Debug("  Decoded UserKey: " & DecodedUserKey)
                mydebuglog.Debug("  Domain: " & Domain)
                mydebuglog.Debug("  PublicKey: " & PublicKey)
                mydebuglog.Debug("  Decoded PublicKey: " & DecodedPublicKey)
                mydebuglog.Debug("  Appsetting dms_cache_age: " & dms_cache_age & vbCrLf)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If Trim(UserId) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No user id specified. "
            GoTo CloseOut2
        End If
        If Trim(DecodedUserKey) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No user key specified. "
            GoTo CloseOut2
        End If
        If Trim(Domain) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No domain specified. "
            GoTo CloseOut2
        End If
        If Not IsNumeric(DecodedPublicKey) Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Incorrect or no public key specified. "
            GoTo CloseOut2
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
        ' Validate identity
        '  The supplied decoded user key should be the same as the id of the contact record
        '  for the user registration id supplied
        If Not cmd Is Nothing Then
            ' -----
            ' Query user
            Try
                SqlS = "SELECT ROW_ID " & _
                "FROM siebeldb.dbo.S_CONTACT " & _
                "WHERE X_REGISTRATION_NUM='" & UserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get user contact info: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If ValidatedUserId <> DecodedUserKey Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserKey & " should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        Else
            results = "Failure"
            GoTo CloseOut
        End If
        '/*
        '        ' ============================================
        '        ' Open SQL Server Express LocalDB database connection for local cache
        '        AppDomain.CurrentDomain.SetData("SQLServerCompactEditionUnderWebHosting", True)
        '        Dim c_con As New SqlConnection(c_ConnS)
        '        Dim c_cmd As New SqlCommand()
        '        Dim c_dr As SqlDataReader

        '        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Attempting to open cache: " & mypath & "cachedb\cache.mdf")
        '        If My.Computer.FileSystem.FileExists(mypath & "cachedb\cache.mdf") Then
        '            ' Open cache database
        '            Try
        '                'c_con = New Data.SqlServerCe.SqlCeConnection(c_ConnS)
        '                c_con.Open()
        '                Try
        '                    'c_cmd = New System.Data.SqlServerCe.SqlCeCommand(SqlS, c_con)
        '                    c_cmd.Connection = c_con
        '                    c_cmd.CommandText = "IF DB_ID('cache') IS NULL CREATE DATABASE cache ON (NAME = dcache, FILENAME = '" & mypath & "cachedb\cache.mdf')"
        '                    If Debug = "Y" Then mydebuglog.Debug("Create cache database query: " & c_cmd.CommandText & vbCrLf)
        '                    If c_con.State = Data.ConnectionState.Closed Then c_con.Open()
        '                    returnv = c_cmd.ExecuteNonQuery()
        '                    c_con.Close()
        '                Catch ex2 As Exception
        '                    errmsg = errmsg & vbCrLf & "Unable to create cache database(2): " & ex2.Message
        '                    results = "Failure"
        '                    GoTo CloseOut2
        '                End Try

        '                '' Create the cache table if it is not there
        '                'Try
        '                '    SqlS = "IF OBJECT_ID (N'dbo.D_CACHE', N'U') IS NULL CREATE TABLE cache.dbo.[D_CACHE]( " & _
        '                '        "[CRSE_ID] [nvarchar](15) NOT NULL, " & _
        '                '        "[ASSET] [nvarchar](100) NOT NULL, " & _
        '                '        "[CREATED] [datetime] NOT NULL, " & _
        '                '        "[LAST_UPD] [datetime] NOT NULL)"
        '                '    If Debug = "Y" Then mydebuglog.Debug("  Create cache database table D_CACHE (2): " & SqlS & vbCrLf)
        '                '    If c_con.State = Data.ConnectionState.Closed Then c_con.Open()
        '                '    c_cmd.CommandText = SqlS
        '                '    returnv = c_cmd.ExecuteNonQuery()
        '                '    c_con.Close()
        '                'Catch ex3 As Exception
        '                '    If Debug = "Y" Then mydebuglog.Debug("  Error creating table [D_CACHE] (2): " & ex3.Message & vbCrLf)
        '                '    results = "Failure"
        '                '    GoTo CloseOut2
        '                'End Try
        '            Catch ex As Exception
        '                errmsg = errmsg & vbCrLf & "Unable to create cache database: " & ex.Message
        '                results = "Failure"
        '                GoTo CloseOut2
        '            End Try
        '        Else
        '            ' Create the database if necessary
        '            Try
        '                c_con.ConnectionString = c_ConnS.Substring(0, c_ConnS.IndexOf(";AttachDbFileName", 0) + 1)  'connect to the localDB instance instead of 'cache' database
        '                c_cmd.Connection = c_con
        '                c_con.Open()
        '                c_cmd.CommandText = "IF DB_ID('cache') IS NULL CREATE DATABASE cache ON (NAME = dcache, FILENAME = '" & mypath & "cachedb\cache.mdf')"
        '                If Debug = "Y" Then mydebuglog.Debug("  Create cache database query: " & c_cmd.CommandText & vbCrLf)
        '                returnv = c_cmd.ExecuteNonQuery()
        '                c_con.Close()

        '            Catch ex2 As Exception
        '                errmsg = errmsg & vbCrLf & "Unable to create database: " & ex2.Message
        '                results = "Failure"
        '                GoTo CloseOut2
        '            End Try
        '        End If

        '        Try
        '            c_con.ConnectionString = c_ConnS
        '            c_cmd.Connection = c_con
        '            c_con.Open()
        '            SqlS = "IF OBJECT_ID('D_CACHE') IS NULL CREATE TABLE [cache].dbo.[D_CACHE]( " & _
        '                "[DOMAIN] [nvarchar](15) NOT NULL, " & _
        '                "[EXT] [nvarchar](15) NOT NULL, " & _
        '                "[KEY] [numeric](18,0) NOT NULL, " & _
        '                "[CREATED] [datetime] NOT NULL, " & _
        '                "[LAST_UPD] [datetime] NOT NULL)"
        '            If Debug = "Y" Then mydebuglog.Debug("   > Create Item cache database query: " & SqlS)
        '            c_cmd.CommandText = SqlS
        '            returnv = c_cmd.ExecuteNonQuery()

        '            SqlS = "IF OBJECT_ID('D_GROUP') IS NULL CREATE TABLE [cache].dbo.[D_GROUP]( " & _
        '                "[DOMAIN] [nvarchar](15) NOT NULL, " & _
        '                "[UGA] [nvarchar](15) NOT NULL)"
        '            If Debug = "Y" Then mydebuglog.Debug("   > Create UGA cache database query: " & SqlS & vbCrLf)
        '            c_cmd.CommandText = SqlS
        '            returnv = c_cmd.ExecuteNonQuery()
        '        Catch ex3 As Exception
        '            errmsg = errmsg & vbCrLf & "Unable to create cache table: " & ex3.Message
        '            results = "Failure"
        '            GoTo CloseOut2
        '        End Try
        '*/

        '' ============================================
        '' Locate Group Access Information for query
        '' Try local cache
        'If Not c_cmd Is Nothing Then
        '    SqlS = "SELECT [UGA] " & _
        '        "FROM [D_GROUP] " & _
        '        "WHERE [DOMAIN]='" & Domain & "'"
        '    If Debug = "Y" Then mydebuglog.Debug("  Check domain UGA cache: " & SqlS)
        '    Try
        '        c_cmd.CommandText = SqlS
        '        c_dr = c_cmd.ExecuteReader()
        '        If Not c_dr Is Nothing Then
        '            While c_dr.Read()
        '                Try
        '                    DomainGID = c_dr(0)
        '                    CacheHit = 1
        '                    If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  DomainGID=" & DomainGID)
        '                Catch ex As Exception
        '                    results = "Failure"
        '                    errmsg = errmsg & "Error checking cache. " & ex.ToString & vbCrLf
        '                    GoTo CloseOut
        '                End Try
        '            End While
        '        Else
        '            If Debug = "Y" Then mydebuglog.Debug("  > Did not find cached domain group")
        '        End If
        '        c_dr.Close()
        '    Catch oBug As Exception
        '        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error checking cache: " & oBug.ToString)
        '        results = "Failure"
        '    End Try
        'Else
        '    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Database error.")
        '    results = "Failure"
        '    GoTo CloseOut
        'End If

        ' -----
        ' Try DMS
        If DomainGID = "" Then
            If Not d_cmd Is Nothing Then
                ' -----
                ' Query DMS
                Try
                    SqlS = "SELECT UGA.row_id " & _
                       "FROM DMS.dbo.User_Group_Access UGA " & _
                        "INNER JOIN DMS.dbo.Groups G on UGA.access_id=G.row_id " & _
                       "WHERE UGA.type_id='G' AND G.name='" & Domain & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Get DOMAIN UGA: " & SqlS)
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                                DomainGID = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "The record was not found." & vbCrLf
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                    results = "Failure"
                End Try
                If Debug = "Y" Then mydebuglog.Debug("   ... DomainGID: " & DomainGID)

                '' -----
                '' Save the found GID in the local cache
                'If DomainGID <> "" Then
                '    SqlS = "INSERT [D_GROUP] ([DOMAIN],[UGA]) " & _
                '        "VALUES('" & Domain & "','" & DomainGID & "')"
                '    If Debug = "Y" Then mydebuglog.Debug("  INSERT UGA cache database query: " & SqlS & vbCrLf)
                '    Try
                '        c_cmd.CommandText = SqlS
                '        returnv = c_cmd.ExecuteNonQuery()
                '    Catch ex As Exception
                '        errmsg = errmsg & "Unable to insert into cache." & ex.ToString & vbCrLf & " with query " & SqlS & vbCrLf
                '        results = "Failure"
                '        retval = 0
                '    End Try
                'End If
            Else
                results = "Failure"
                GoTo CloseOut
            End If
        End If
        If Trim(DomainGID) = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No domain access found."
            GoTo CloseOut2
        End If

        '' ============================================
        '' Create output directory for ItemName caching
        'SaveDest = mypath & "domain_temp\" & Domain
        'Try
        '    Directory.CreateDirectory(SaveDest)
        'Catch
        'End Try
        'If Debug = "Y" Then mydebuglog.Debug("  Caching to: " & SaveDest & vbCrLf)

        '' ============================================
        '' Check to see if and when we already cached this item
        'CacheHit = 0
        'If Not c_cmd Is Nothing Then
        '    SqlS = "SELECT [LAST_UPD],[EXT] " & _
        '        "FROM [D_CACHE] " & _
        '        "WHERE [DOMAIN]='" & Domain & "' AND [KEY]=" & DecodedPublicKey
        '    If Debug = "Y" Then mydebuglog.Debug("  Check cache for item: " & SqlS)
        '    Try
        '        c_cmd.CommandText = SqlS
        '        c_dr = c_cmd.ExecuteReader()
        '        If Not c_dr Is Nothing Then
        '            While c_dr.Read()
        '                Try
        '                    cLastUpd = c_dr(0)
        '                    cExt = Trim(CheckDBNull(c_dr(1), enumObjectType.StrType))
        '                    cFileName = SaveDest & "\" & DecodedPublicKey & "." & cExt
        '                    CacheHit = 1
        '                    If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  LastUpd=" & Format(cLastUpd) & "  CacheHit=" & CacheHit)

        '                    ' If cache hit, check to see if the file exists
        '                    If CacheHit = 1 Then
        '                        If Not My.Computer.FileSystem.FileExists(cFileName) Then
        '                            If Debug = "Y" Then mydebuglog.Debug("   !! marked as cached but file not found in filesystem." & vbCrLf)
        '                            CacheHit = 0
        '                        End If
        '                    End If
        '                Catch ex As Exception
        '                    results = "Failure"
        '                    errmsg = errmsg & "Error checking cache. " & ex.ToString & vbCrLf
        '                End Try
        '            End While
        '        Else
        '            errmsg = errmsg & "Error checking cache." & vbCrLf
        '            c_dr.Close()
        '            results = "Failure"
        '        End If
        '        c_dr.Close()
        '    Catch oBug As Exception
        '        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error checking cache: " & oBug.ToString)
        '        results = "Failure"
        '    End Try
        'End If

        '***** Check cached object; 2/28/17; Ren Hou;  ****'
        Dim filecache As ObjectCache = MemoryCache.Default
        Dim fileContents(1000) As Byte
        Dim cacheItemName As String = DomainGID & "-" & DecodedPublicKey & "-" & ValidatedUserId
        Dim last_upt As Date = Today.AddYears(-50)
        ' Check to see if the document is in the in-memory cache
        If Not IsNothing(filecache(cacheItemName)) Then
            'Check if the cached item need to be renewed
            SqlS = "SELECT TOP 1 V.last_upd " & _
                    "FROM DMS.dbo.Documents D " & _
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id " & _
                    "LEFT OUTER JOIN DMS.dbo.Document_Types DT on D.data_type_id=DT.row_id " & _
                    "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id " & _
                    "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id " & _
                    "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id " & _
                    "WHERE D.row_id=" & DecodedPublicKey & " and D.deleted IS NULL AND " & _
                    "((A.row_id=3 AND DA.fkey='" & ValidatedUserId & "' AND DA.access_flag='Y')) AND " & _
                    "DU.user_access_id"
            ' Translate DomainGID with equivalents
            If DomainGID = "11" Or DomainGID = "112" Or DomainGID = "828" Then
                SqlS = SqlS & " IN (11,112,828)"
            Else
                SqlS = SqlS & "=" & DomainGID
            End If

            Try
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check last updated date for xml document: " & vbCrLf & SqlS)
                d_cmd.CommandText = SqlS
                last_upt = d_cmd.ExecuteScalar()
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
                results = "Failure"
            End Try

            If last_upt > TryCast(filecache(cacheItemName), HciDMSDocument).UpdateDate Then
                'Remove if the update_date on the cache is before the last updtaed date on DB record.
                filecache.Remove(cacheItemName)
                mydebuglog.Debug(vbCrLf & "  Cached object " & cacheItemName & " expired." & vbCrLf)
            End If
        End If
        '****  ****'

        'Load content of cached object
        If filecache(cacheItemName) Is Nothing Then
            fileContents = Nothing
        Else
            Dim tmpObj As Object
            tmpObj = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
            ReDim fileContents(tmpObj.Length)
            fileContents = TryCast(filecache(cacheItemName), HciDMSDocument).CachedObj
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & " Retrieved Cached object " & cacheItemName & vbCrLf)
        End If

        If IsNothing(fileContents) Or Debug = "R" Then
            ' ============================================
            ' Get the item and meta data from the DMS
            If Debug = "Y" Then mydebuglog.Debug("  Looking in DMS for document # " & DecodedPublicKey & vbCrLf)
            If Not d_cmd Is Nothing Then
                ' Check to see if we have access and get meta information
                SqlS = "SELECT TOP 1 D.dfilename, V.dimage, D.row_id, V.last_upd, V.dsize as length, DT.extension, DU.access_type, V.minio_flg, V.row_id " &
                        "FROM DMS.dbo.Documents D  " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Types DT on D.data_type_id=DT.row_id " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id  " &
                        "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id  " &
                        "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id  " &
                        "WHERE D.row_id=" & DecodedPublicKey & " and D.deleted IS NULL AND  " &
                        "((A.row_id=3 AND DA.fkey='" & ValidatedUserId & "' AND DA.access_flag='Y')) AND  " &
                        "DU.user_access_id"
                ' Translate DomainGID with equivalents
                If DomainGID = "11" Or DomainGID = "112" Or DomainGID = "828" Then
                    SqlS = SqlS & " IN (11,112,828)"
                Else
                    SqlS = SqlS & "=" & DomainGID
                End If

                If Debug = "Y" Then mydebuglog.Debug("  Get item information: " & SqlS)
                Try
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                d_item_name = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(2), enumObjectType.StrType))
                                d_last_upd = CheckDBNull(d_dr(3), enumObjectType.DteType)
                                d_dsize = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                d_ext = Trim(CheckDBNull(d_dr(5), enumObjectType.StrType))
                                d_access = Trim(CheckDBNull(d_dr(6), enumObjectType.StrType))
                                minio_flg = Trim(CheckDBNull(d_dr(7), enumObjectType.StrType))
                                d_verid = Trim(CheckDBNull(d_dr(8), enumObjectType.StrType))

                                ' Verify domain read access to the document - meaning its accessible to users in that domain
                                If InStr(UCase(d_access), "R") = 0 Then
                                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  No access to this item based on domain security.")
                                    results = "Failure"
                                    GoTo CloseOut
                                End If

                                If minio_flg = "Y" Then
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Minio")
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
                                            ReDim outbyte(Val(retval - 1))
                                            Dim intval As Integer
                                            For i = 0 To retval - 1
                                                intval = mobj2.ResponseStream.ReadByte()
                                                If intval < 255 And intval > 0 Then
                                                    outbyte(i) = intval
                                                End If
                                                If intval = 255 Then outbyte(i) = 255
                                                If intval < 0 Then
                                                    outbyte(i) = 0
                                                    If Debug = "Y" Then mydebuglog.Debug("  >  .. " & i.ToString & "   intval: " & intval.ToString)
                                                End If
                                            Next
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
                                    ' Get binary and attach to the object outbyte if found, not cached or updated recently
                                    '   retval will be "0" if this is not the case
                                    'If d_dsize <> "" And (CacheHit = 0 Or d_last_upd > cLastUpd) Then
                                    If d_dsize <> "" Then
                                        If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Document_Versions")
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        Try
                                            retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                        Catch ex As Exception
                                            results = "Failure"
                                            errmsg = errmsg & "Error getting item. " & ex.ToString & vbCrLf
                                            GoTo CloseOut
                                        End Try
                                    End If
                                End If
                                If Debug = "Y" Then
                                    mydebuglog.Debug("  > Found record on query.  Extension: " & d_ext)
                                    mydebuglog.Debug("                            Item Name: " & d_item_name)
                                    mydebuglog.Debug("                            Minio: " & minio_flg)
                                    mydebuglog.Debug("                            Doc Id: " & d_doc_id)
                                    mydebuglog.Debug("                            Version Id: " & d_verid)
                                    mydebuglog.Debug("                            Size: " & d_dsize)
                                    mydebuglog.Debug("                            Last Updated: " & Format(d_last_upd))
                                    mydebuglog.Debug("                            Bytes retrieved: " & retval.ToString)
                                End If

                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error getting item. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                        d_dr.Close()
                    Else
                        errmsg = errmsg & "Error getting item." & vbCrLf
                        results = "Failure"
                    End If
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error getting item: " & oBug.ToString)
                    results = "Failure"
                    GoTo CloseOut
                End Try

                ' Verify document is the same
                If Trim(LCase(d_doc_id)) <> Trim(LCase(DecodedPublicKey)) Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Unable to verify identity when comparing " & d_doc_id & " to " & DecodedPublicKey)
                    results = "Failure"
                    GoTo CloseOut
                End If

                '***** Set cache object; 2/28/17; Ren Hou;  ****
                Dim policy As New CacheItemPolicy()
                policy.SlidingExpiration = TimeSpan.FromDays(CDbl(dms_cache_age))
                filecache.Set(cacheItemName, New HciDMSDocument(d_last_upd, outbyte), policy)
                ReDim fileContents(outbyte.Length)
                fileContents = outbyte
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Caching to key: " & cacheItemName & vbCrLf)
                '****  ****

                '' Generate cache filename
                ''  The formula for this [Document Id] + [Extension]
                'BinaryFile = SaveDest & "\" & DecodedPublicKey & "." & d_ext
                'If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Computed cache filename: " & BinaryFile & vbCrLf)

                'If Debug = "Y" Then
                '    mydebuglog.Debug("   ... Reported d_dsize: " & d_dsize & " : Non-cached or superceded bytes found: " & Str(retval))
                'End If
            Else
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Database error.")
                results = "Failure"
                GoTo CloseOut
            End If
        End If


        '' ============================================
        '' Retrieve item and store-to/replace-in the filesystem if found, 
        ''  not cached yet or updated recently
        'If retval > 0 And (CacheHit = 0 Or d_last_upd > cLastUpd) Then
        '    If Debug = "Y" Then mydebuglog.Debug("  Storing to filesystem: " & BinaryFile)
        '    If (My.Computer.FileSystem.FileExists(BinaryFile)) Then Kill(BinaryFile)
        '    Try
        '        bfs = New FileStream(BinaryFile, FileMode.Create, FileAccess.Write)
        '        bw = New BinaryWriter(bfs)
        '        bw.Write(outbyte)
        '        bw.Flush()
        '        bw.Close()
        '        bfs.Close()
        '    Catch ex As Exception
        '        errmsg = errmsg & "Unable to write the file to a temp file." & ex.ToString & vbCrLf
        '        results = "Failure"
        '        retval = 0
        '    End Try
        '    bfs = Nothing
        '    bw = Nothing
        'End If

        '' ============================================
        ''Double check to make sure file exists at this point
        ''  Remove from cache if applicable
        'If Not My.Computer.FileSystem.FileExists(BinaryFile) Then
        '    If Debug = "Y" Then mydebuglog.Debug("  !! The file was not properly stored ")
        '    errmsg = errmsg & "Unable to find file: " & BinaryFile & vbCrLf
        '    results = "Failure"
        '    retval = 0

        '    ' Remove from cache
        '    If CacheHit = 1 Then
        '        SqlS = "DELETE FROM [D_CACHE] WHERE [DOMAIN]='" & Domain & "' AND [KEY]=" & DecodedPublicKey
        '        If Debug = "Y" Then mydebuglog.Debug("  Removing cache entry: " & SqlS)
        '        Try
        '            c_cmd.CommandText = SqlS
        '            returnv = c_cmd.ExecuteNonQuery()
        '        Catch ex As Exception
        '            errmsg = errmsg & "Unable to delete from cache." & ex.ToString & vbCrLf & " with query " & SqlS & vbCrLf
        '            results = "Failure"
        '            retval = 0
        '        End Try
        '    End If
        '    GoTo CleanCache
        'End If

        '' ============================================
        '' Update cache 
        ''   If made it to this point then the file must exist
        'If CacheHit = 1 Then
        '    ' Update cache since it was already there
        '    SqlS = "UPDATE [D_CACHE] SET [LAST_UPD]=GETDATE() " & _
        '        "WHERE [DOMAIN]='" & Domain & "' AND [KEY]=" & DecodedPublicKey
        '    If Debug = "Y" Then mydebuglog.Debug("  Updating cache entry: " & SqlS)
        'Else
        '    ' Add to cache since the file is new
        '    SqlS = "INSERT [D_CACHE] ([DOMAIN], [KEY], [EXT], [CREATED], [LAST_UPD]) " & _
        '    "VALUES ('" & Domain & "'," & DecodedPublicKey & ", '" & d_ext & "', GETDATE(), GETDATE())"
        '    If Debug = "Y" Then mydebuglog.Debug("  Inserting cache entry: " & SqlS)
        'End If
        '' Execute query
        'Try
        '    c_cmd.CommandText = SqlS
        '    returnv = c_cmd.ExecuteNonQuery()
        '    If returnv = 0 Then
        '        errmsg = errmsg & "Unable to write to cache."
        '    End If
        'Catch ex As Exception
        '    errmsg = errmsg & "Unable to write to cache." & ex.ToString & vbCrLf & " with query " & SqlS & vbCrLf
        '    results = "Failure"
        '    retval = 0
        'End Try

CleanCache:
        '' ============================================
        '' Cleanup cache as required
        'SqlS = "DELETE FROM D_CACHE WHERE LAST_UPD<dateadd(dy,-" & dms_cache_age & ",getdate())"
        'If Debug = "Y" Then mydebuglog.Debug("  Locate old cache entries to remove: " & SqlS)
        'Try
        '    c_cmd.CommandText = SqlS
        '    c_cmd.ExecuteNonQuery()
        'Catch oBug As Exception
        '    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error deleting cache: " & oBug.ToString)
        '    results = "Failure"
        'End Try
        'If Debug = "Y" Then mydebuglog.Debug("   Cache entries removed: " & killcount.ToString)

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
        If Trim(errmsg) <> "" Then myeventlog.Error("GetMediaSecure : Error: " & Trim(errmsg))
        If CacheHit = 1 Then
            myeventlog.Info("GetMediaSecure : Results: " & results & " for CACHED item " & d_item_name & ", in " & Domain & " domain with cache filename " & DecodedPublicKey & "." & d_ext)
        Else
            myeventlog.Info("GetMediaSecure : Results: " & results & " for item " & d_item_name & ", in " & Domain & " domain, minio: " & minio_flg & ", docid: " & d_doc_id & ", verid: " & d_verid & " with cache filename " & DecodedPublicKey & "." & d_ext)
        End If
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If CacheHit = 1 Then
                    mydebuglog.Debug("Results: " & results & " for CACHED item " & d_item_name & ", in " & Domain & " domain with cache filename " & DecodedPublicKey & "." & d_ext & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
                Else
                    mydebuglog.Debug("Results: " & results & " for item " & d_item_name & ", in " & Domain & " domain, minio: " & minio_flg & ", docid: " & d_doc_id & ", verid: " & d_verid & " with cache filename " & DecodedPublicKey & "." & d_ext & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
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
        ' Return item
        'If (My.Computer.FileSystem.FileExists(BinaryFile)) And Debug <> "T" Then
        Try
            'bfs = File.Open(BinaryFile, FileMode.Open, FileAccess.Read)
            'Dim lngLen As Long = bfs.Length
            'ReDim outbyte(CInt(lngLen - 1))
            'bfs.Read(outbyte, 0, CInt(lngLen))
            'Return outbyte
            Return fileContents
        Catch exp As Exception
            Return Nothing
        Finally
            'bfs.Close()
            'bfs = Nothing
            outbyte = Nothing
            fileContents = Nothing
        End Try
        'Else
        'Return Nothing
        'End If
    End Function

    <WebMethod(Description:="Provides terms and definitions for a course based on a provided registration")> _
    Public Function CourseGlossary(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function provides terms and definitions for a course based on a provided registration

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	    - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim i As Integer
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' KBA declarations
        Dim DecodedUserId As String
        Dim num_terms As Integer
        Dim term(100) As String
        Dim term_id(100) As String
        Dim term_definition(100) As String
        Dim term_lang(100) As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        num_terms = 0

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\CourseGlossary.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        If Not cmd Is Nothing Then
            Try
                SqlS = "SELECT G.TERM, G.ROW_ID, G.DEFINITION, G.LANG_CD " & _
                "FROM elearning.dbo.GLOSSARY G " & _
                "LEFT OUTER JOIN elearning.dbo.GLOSSARY_CRSE GC ON GC.GLOSSARY_ID=G.ROW_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_REG R ON R.CRSE_ID=GC.CRSE_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "' AND C.X_REGISTRATION_NUM='" & DecodedUserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get terms: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            num_terms = num_terms + 1
                            term(num_terms) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            term_id(num_terms) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            term_definition(num_terms) = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            term_lang(num_terms) = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query: " & term(num_terms))
                            If term(num_terms) = "" Then results = "Failure"
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting terms. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "Terms were not found." & vbCrLf
                    dr.Close()
                    results = "Failure"
                End If
                dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try

            ' If no resyults found, then failure
            If num_terms = 0 Then results = "Failure"
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' ============================================
        ' Return the glossary terms
        '   <glossary num_terms="#">
        '       <term name="Term" id="Question Id" lang="Language Cd">Definition</term>
        '       <term name="Term" id="Question Id" lang="Language Cd">Definition</term>
        '       ...
        '   </glossary>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        Dim resultsItem As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("glossary")
        AddXMLAttribute(odoc, resultsRoot, "num_terms", num_terms.ToString)
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        Try
            ' Add result items - send what was submitted for debugging purposes 
            If Debug <> "T" Then
                For i = 1 To num_terms
                    resultsItem = odoc.CreateElement("term")
                    resultsItem.InnerText = term_definition(i)
                    AddXMLAttribute(odoc, resultsItem, "name", term(i))
                    AddXMLAttribute(odoc, resultsItem, "id", term_id(i))
                    AddXMLAttribute(odoc, resultsItem, "lang", term_lang(i))
                    resultsRoot.AppendChild(resultsItem)
                Next
            End If
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("CourseGlossary : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("CourseGlossary : Results: " & results & " for RegId # " & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " at " & Now.ToString)
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
        Return odoc
    End Function

    <WebMethod(Description:="Save course feedback to our database")> _
    Public Function SaveFeedback(ByVal RegId As String, ByVal PageId As String, ByVal BugType As String, _
       ByVal Comment As String, ByVal UserId As String, ByVal Debug As String) As Boolean
        ' This function validates an answer for a Knowledge-Based Authentication 
        ' system based on a provided question id and answer, and notes that it was validated

        ' The input parameters are as follows:
        '   RegId       - The siebeldb.CX_SESS_REG.ROW_ID of the user (learnerId)
        '   PageId      - The id of the page in the course (pageId)
        '   BugType	- The type of feedback (feedbackType)
        '   Comment     - The users comment (description)
        '   UserId      - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	- "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Output:
        '   True or False, depending on whether the feedback was saved or not
        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Feedback declarations
        Dim DecodedUserId, ValRegId, ConId As String

        ' Results declarations
        Dim myresults As Boolean

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "False"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        ValRegId = ""
        ConId = ""
        myresults = False

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If Debug = "" Then Debug = "N"
        If (RegId = "" Or Comment = "" Or UserId = "") And Debug <> "T" Then
            myresults = False
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            PageId = ""
            BugType = "other"
            Comment = "This is a test"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            ' Fix class registration id
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            ' Fix user registration id
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
            ' Fix Comment text
            Comment = Trim(HttpUtility.UrlEncode(Comment))
            If InStr(Comment, "%") > 0 Then Comment = Trim(HttpUtility.UrlDecode(Comment))
            If InStr(Comment, "%") > 0 Then Comment = Trim(Comment)
            Comment = DecodeParamSpaces(Comment)
            Comment = FilterString(Comment)
            BugType = Trim(BugType)
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\SaveFeedback.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                myresults = False
                GoTo CloseOut2
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  PageId: " & PageId)
                mydebuglog.Debug("  Comment: " & Comment)
                mydebuglog.Debug("  BugType: " & BugType)
            End If
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            myresults = False
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            myresults = False
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        If Not cmd Is Nothing Then
            ' Verify that the user exists
            Try
                SqlS = "SELECT R.ROW_ID, C.ROW_ID " & _
                    "FROM siebeldb.dbo.CX_SESS_REG R " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                    "WHERE R.ROW_ID='" & RegId & "' AND C.X_REGISTRATION_NUM='" & DecodedUserId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Validate user: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            ValRegId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            ConId = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query: " & ValRegId)
                            If ValRegId = "" Then
                                myresults = False
                                errmsg = errmsg & "Error validating user. " & vbCrLf
                            End If
                        Catch ex As Exception
                            myresults = False
                            errmsg = errmsg & "Error validating user. " & ex.ToString & vbCrLf
                        End Try
                    End While
                Else
                    errmsg = errmsg & "User not validated." & vbCrLf
                    myresults = False
                    dr.Close()
                    GoTo CloseOut
                End If
                dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                myresults = False
                GoTo CloseOut
            End Try

            ' If user not validated, report failure
            If ValRegId <> RegId Then
                errmsg = errmsg & "Error validating user." & vbCrLf
                myresults = False
                GoTo CloseOut
            End If

            ' Exit if debugging without saving
            If Debug = "T" Then
                If ValRegId = RegId Then myresults = True
                GoTo CloseOut
            End If

            ' Write comment and set results if the record does not already exist.  If it does then simply return a true result as well
            Try
                ' ELN_FEEDBACK Field Structure:
                '   USER_NAME       FK to siebeldb.dbo.S_CONTACT.ROW_ID
                '   PAGE            Page Id
                '   PAGE_STR        The breadcrumb trail to the structure
                '   COMM            The message entered by the user
                '   BUG_TYPE        Type of error
                '   INSERT_DATE     Date the error was created
                '   STATUS          Status of the feedback - "Open" by default
                '   REG_ID          FK to siebeldb.dbo.CX_SESS_REG.ROW_ID
                SqlS = "INSERT elearning.dbo.ELN_FEEDBACK " &
                    "(USER_NAME, PAGE, PAGE_STR, COMM, BUG_TYPE, INSERT_DATE, STATUS, REG_ID) " &
                    "SELECT '" & ConId & "','" & SqlString(PageId) & "','','" &
                    SqlString(Comment) & "','" & BugType & "',GETDATE(),'Open','" & RegId & "' " &
                    "WHERE NOT EXISTS (SELECT ID FROM elearning.dbo.ELN_FEEDBACK WHERE USER_NAME='" & ConId & "' AND PAGE='" & SqlString(PageId) & "' AND COMM='" & SqlString(Comment) & "')"
                If Debug = "Y" Then mydebuglog.Debug("  Inserting Feedback: " & SqlS)
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                myresults = True
                'If returnv = 0 Then
                'myresults = False
                'Else
                'myresults = True
                'End If
            Catch ex As Exception
                results = "Failure"
                errmsg = errmsg & "Error inserting feedback. " & ex.ToString & vbCrLf
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
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
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("SaveFeedback : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("SaveFeedback : Results: " & myresults & " for RegId: " & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & myresults & " for RegId: " & RegId)
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
        Return myresults
    End Function

    <WebMethod(Description:="Retrieves user-specific course configuration information")> _
    Public Function GetConfiguration(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function provides configuration and personalization information for course 
        '   attendees.

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	        - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging, scormlocal As String
        Dim DecodedUserId, ValidatedUserId As String
        Dim temp1, temp2, temp3 As String
        Dim i As Double

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Configuration Data declarations
        Dim CRSE_ID, JURISDICTION, JURIS_LVL, CITY, COUNTY, STATE, COUNTRY As String
        Dim BASE_CONTENT_URL, BASE_SERVICE_URL, MEDIA_DEST, VIDEO_DEST, FEEDBACK_EMAIL, MEDIA_X, MEDIA_Y As String
        Dim FEEDBACK_EMAIL_FROM, FEEDBACK_EMAIL_SUBJ, FEEDBACK_SUCCESS_MSG, SVC_FEEDBACK, GLOSSARY_ENABLED As String
        Dim SVC_GLOSSARY, KBA_REQD, KBA_FAIL_ALWD, KBA_MAX_TIME, KBA_FAIL_MSG, KBA_EXIT_MSG, RESOURCE_DEST As String
        Dim SVC_KBA_QUESTIONS, SVC_KBA_ANSWERS, LESSON_SKINS, LESSON_LAUNCH_LMT, LESSON_LAUNCH_MSG, LESSON_MAX_TIME As String
        Dim LESSON_MAX_TIME_MSG, LESSON_MIN_TIME, LESSON_MIN_TIME_MSG, LESSON_LINEAR, SVC_LESSON As String
        Dim QUIZ_FONT, QUIZ_FEEDBACK, QUIZ_RANDOM, QUIZ_ATTEMPTS, QUIZ_MIN_SCORE, QUIZ_TIMER_ENABLED, QUIZ_TIMER_LIMIT, SVC_QUIZ As String
        Dim REG_DATA(100), REG_QUES(100), LIC_DATA(100), LIC_ANSR(100), JURIS_ID, SHOW_VOLUME_CONTROL As String
        Dim BACKGROUND_LOADER, SHOW_PAGE_COUNT, NUMBER_BY, LOCK_CAP_SLIDE_BUTTONS, TEST_FLG As String
        Dim SHELL_PRELOAD_PCT, RTMP_FOLDER, SCENARIO_DEST, SCENARIO_ENABLED, SCENARIO_NARRATION As String
        Dim INACT_TIME_LIMIT, INACT_PROMPT_LIMIT, MIN_TIME_PER_PAGE As String
        Dim SLP_CONTINUE, SLP_MESSAGE, SLP_RESTART_MSG, SLP_RESTART_ACTION, LANG_CD As String
        Dim QUIZ_X, QUIZ_Y, QUIZ_WIDTH, QUIZ_HEIGHT As String
        Dim VIDEO_THUMB_X, VIDEO_THUMB_Y, VIDEO_X, VIDEO_Y, VIDEO_WIDTH, VIDEO_HEIGHT As String
        Dim RegCount, LicCount As Double
        Dim temp As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

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
        'Debug = "Y"

        CRSE_ID = ""
        JURISDICTION = ""
        JURIS_LVL = ""
        CITY = ""
        COUNTY = ""
        STATE = ""
        COUNTRY = ""
        BASE_CONTENT_URL = ""
        BASE_SERVICE_URL = ""
        MEDIA_DEST = ""
        MEDIA_X = ""
        MEDIA_Y = ""
        VIDEO_DEST = ""
        VIDEO_THUMB_X = ""
        VIDEO_THUMB_Y = ""
        VIDEO_X = ""
        VIDEO_Y = ""
        VIDEO_WIDTH = ""
        VIDEO_HEIGHT = ""
        RESOURCE_DEST = ""
        FEEDBACK_EMAIL = ""
        FEEDBACK_EMAIL_FROM = ""
        FEEDBACK_EMAIL_SUBJ = ""
        FEEDBACK_SUCCESS_MSG = ""
        SVC_FEEDBACK = ""
        GLOSSARY_ENABLED = ""
        SVC_GLOSSARY = ""
        KBA_REQD = ""
        KBA_FAIL_ALWD = ""
        KBA_MAX_TIME = ""
        KBA_FAIL_MSG = ""
        KBA_EXIT_MSG = ""
        SVC_KBA_QUESTIONS = ""
        SVC_KBA_ANSWERS = ""
        LESSON_SKINS = ""
        LESSON_LAUNCH_LMT = ""
        LESSON_LAUNCH_MSG = ""
        LESSON_MAX_TIME = ""
        LESSON_MAX_TIME_MSG = ""
        LESSON_MIN_TIME = ""
        LESSON_MIN_TIME_MSG = ""
        LESSON_LINEAR = ""
        SVC_LESSON = ""
        QUIZ_FONT = """"
        QUIZ_FEEDBACK = ""
        QUIZ_RANDOM = ""
        QUIZ_ATTEMPTS = ""
        QUIZ_MIN_SCORE = ""
        QUIZ_TIMER_ENABLED = ""
        QUIZ_TIMER_LIMIT = ""
        QUIZ_X = ""
        QUIZ_Y = ""
        QUIZ_WIDTH = ""
        QUIZ_HEIGHT = ""
        SHOW_VOLUME_CONTROL = "false"
        BACKGROUND_LOADER = ""
        SHOW_PAGE_COUNT = ""
        NUMBER_BY = ""
        LOCK_CAP_SLIDE_BUTTONS = ""
        INACT_TIME_LIMIT = ""
        INACT_PROMPT_LIMIT = ""
        SVC_QUIZ = ""
        JURIS_ID = ""
        TEST_FLG = ""
        SHELL_PRELOAD_PCT = ""
        RTMP_FOLDER = ""
        SCENARIO_DEST = ""
        SCENARIO_ENABLED = ""
        SCENARIO_NARRATION = "false"
        MIN_TIME_PER_PAGE = ""
        SLP_CONTINUE = ""
        SLP_MESSAGE = ""
        SLP_RESTART_MSG = ""
        SLP_RESTART_ACTION = ""
        LANG_CD = ""
        RegCount = 0
        LicCount = 0
        temp = ""
        scormlocal = "N"
        i = 0

        '8/5/15;Ren Hou; Added new valriables for HTML5 configuration
        Dim SCRN_SIZE_MIN_M = ""
        Dim SCRN_SIZE_MIN_T = ""
        Dim SCRN_SIZE_MIN_D = ""
        Dim BKG_COLOR_NORMAL = ""
        Dim BKG_COLOR_OVER = ""
        Dim BKG_COLOR_ACTIVE = ""
        Dim BKG_COLOR_DISABLED = ""
        Dim TEXT_COLOR_NORMAL = ""
        Dim TEXT_COLOR_OVER = ""
        Dim TEXT_COLOR_ACTIVE = ""
        Dim TEXT_COLOR_DISABLED = ""
        Dim SKIN_COLORS_MAIN_TEXT = ""
        Dim SKIN_COLORS_BACKGROUND = ""
        Dim SKIN_COLORS_HEADER_TEXT = ""
        Dim SKIN_COLORS_UI_BUTTON_BG = ""
        Dim SKIN_COLORS_UI_BUTTON_TEXT = ""
        Dim SKIN_COLORS_UI_BUTTON_HL = ""
        Dim SKIN_COLORS_MENU_BUTTON_TEXT = ""
        Dim SKIN_COLORS_MENU_BUTTON_HL = ""
        Dim SKIN_COLORS_PAGE_COUNTER_TEXT = ""
        Dim SKIN_COLORS_QUIZ_UI_BUTTON = ""
        Dim SKIN_COLORS_QUIZ_UI_BUTTON_TEXT = ""
        Dim SKIN_COLORS_QUIZ_SCORE_SUM_TEXT = ""
        Dim AUTO_ADV_DELAY_SEC = ""
        Dim SCORM_VERSION = ""
        Dim LSSN_STATUS_TYPE = ""
        Dim MEDIA_LOCKPAGE = ""
        Dim QUIZ_REQ_PASS_BEFORE_CONT = ""
        Dim QUIZ_ATTMPT_LIMIT = ""
        Dim QUIZ_MAX_TIME_LIMIT = ""
        Dim QUIZ_QUST_RANDOM = ""
        Dim SCORM_OGJ_BYQUIZ = ""

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=hcidb"
            scormlocal = System.Configuration.ConfigurationManager.AppSettings.Get("scormlocal")
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetConfiguration_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config"
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If RegId = "" And Debug = "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetConfiguration.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  scormlocal: " & scormlocal)
            End If
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process request
        If Not cmd Is Nothing Then
            ' -----
            ' Get registration and jurisdiction-specific information
            Try
                SqlS = "SELECT (SELECT CASE WHEN JC.KBA_REQD IS NULL THEN 'N' ELSE JC.KBA_REQD END) AS KBA_REQD,  " & _
                "(SELECT CASE WHEN JC.KBA_FAIL_ALWD IS NULL THEN '0' ELSE JC.KBA_FAIL_ALWD END) AS KBA_FAIL_ALWD,  " & _
                "(SELECT CASE WHEN JC.KBA_FAIL_MSG IS NULL THEN '' ELSE JC.KBA_FAIL_MSG END) AS KBA_FAIL_MSG,  " & _
                "(SELECT CASE WHEN JC.KBA_MAX_TIME IS NULL THEN '0' ELSE JC.KBA_MAX_TIME END) AS KBA_MAX_TIME, R.CRSE_ID, " & _
                "J.NAME, J.JURIS_LVL, J.CITY, J.COUNTY, J.STATE, J.COUNTRY, C.X_REGISTRATION_NUM, R.JURIS_ID, " & _
                "(SELECT CASE WHEN R.TEST_FLG IS NULL OR R.TEST_FLG='' THEN 'N' ELSE R.TEST_FLG END) AS TEST_FLG, " & _
                "(SELECT CASE WHEN C.X_PR_LANG_CD IS NULL OR C.X_PR_LANG_CD='' THEN S.LANG_ID ELSE C.X_PR_LANG_CD END) AS LANG_ID " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR S ON S.ROW_ID=R.TRAIN_OFFR_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_JURIS_CRSE JC ON JC.CRSE_ID=R.CRSE_ID AND JC.JURIS_ID=R.JURIS_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_JURISDICTION_X J ON J.ROW_ID=R.JURIS_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration and jurisdiction-specific information data: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            KBA_REQD = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            KBA_FAIL_ALWD = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            KBA_FAIL_MSG = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            KBA_MAX_TIME = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            CRSE_ID = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            JURISDICTION = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                            JURIS_LVL = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                            CITY = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                            COUNTY = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                            STATE = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                            COUNTRY = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                            ValidatedUserId = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                            JURIS_ID = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                            TEST_FLG = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                            LANG_CD = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                            If LANG_CD = "" Then LANG_CD = "ENU"
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
            If Debug = "Y" Then
                mydebuglog.Debug("   ... CRSE_ID: " & CRSE_ID)
                mydebuglog.Debug("   ... LANG_CD: " & LANG_CD)
                mydebuglog.Debug("   ... TEST_FLG: " & TEST_FLG)
                mydebuglog.Debug("   ... JURIS_ID: " & JURIS_ID)
            End If

            ' -----
            ' Verify the user
            If ValidatedUserId <> DecodedUserId Then
                results = "Failure"
                CRSE_ID = ""
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If

            If JURIS_ID <> "" And CRSE_ID <> "" Then
                ' -----
                ' Get Regulatory Data
                SqlS = "SELECT I.ACCESS_KEY, R.REG_TEXT " & _
                "FROM reports.dbo.DB_ISSUE I " & _
                "INNER JOIN reports.dbo.DB_ISSUE_CRSE C ON C.ISSUE_ID=I.ROW_ID " & _
                "LEFT OUTER JOIN reports.dbo.DB_REG R ON R.ISSUE_ID=I.ROW_ID " & _
                "LEFT OUTER JOIN reports.dbo.DB_JURIS J ON J.ROW_ID=R.JURIS_ID " & _
                "WHERE C.CRSE_ID='" & CRSE_ID & "' AND J.ALT_JURIS_ID='" & JURIS_ID & "' " & _
                "AND I.LANG_CD='" & LANG_CD & "' AND J.JLEVEL<>'Local'"
                If Debug = "Y" Then mydebuglog.Debug("  Get Regulatory data: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                RegCount = RegCount + 1
                                REG_DATA(RegCount) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                REG_QUES(RegCount) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                'temp = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                'REG_QUES(RegCount) = FilterString(temp)
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading Regulatory data: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                    results = "Failure"
                End Try
                dr.Close()

                ' -----
                ' Get Licensing Information from legacy system 
                'SqlS = "SELECT I.CODE, 'yes' AS STATUS " & _
                '"FROM elearning.dbo.ELN_JURISDICTION EJ " & _
                '"INNER JOIN siebeldb.dbo.CX_JURISDICTION_X J1 ON J1.JURISDICTION=EJ.CODE " & _
                '"LEFT OUTER JOIN elearning.dbo.ELN_JURISDICTION_IDENTIFICATION JI ON JI.JURISDICTION_ID=EJ.JURIS_ID " & _
                '"INNER JOIN elearning.dbo.ELN_IDENTIFICATION I ON I.IDENTIFICATION_ID=JI.IDENTIFICATION_ID " & _
                '"WHERE J1.ROW_ID='" & JURIS_ID & "' " & _
                '"WHERE J1.ROW_ID='" & JURIS_ID & "' AND I.CODE IS NOT NULL " & _
                '"UNION (SELECT CODE, 'no' AS STATUS FROM elearning.dbo.ELN_IDENTIFICATION_SCORM WHERE CODE IS NOT NULL) " & _
                '"ORDER BY I.CODE, STATUS"

                SqlS = "SELECT I.CODE, 'yes' AS STATUS " & _
                "FROM elearning.dbo.ELN_JURISDICTION EJ " & _
                "INNER JOIN siebeldb.dbo.CX_JURISDICTION_X J1 ON J1.JURISDICTION=EJ.CODE " & _
                "LEFT OUTER JOIN elearning.dbo.ELN_JURISDICTION_IDENTIFICATION JI ON JI.JURISDICTION_ID=EJ.JURIS_ID " & _
                "INNER JOIN elearning.dbo.ELN_IDENTIFICATION I ON I.IDENTIFICATION_ID=JI.IDENTIFICATION_ID " & _
                "WHERE J1.ROW_ID='" & JURIS_ID & "' AND I.CODE IS NOT NULL " & _
                "UNION (SELECT CODE, 'no' AS STATUS FROM elearning.dbo.ELN_IDENTIFICATION_SCORM WHERE CODE IS NOT NULL) " & _
                "ORDER BY I.CODE, STATUS"

                If Debug = "Y" Then mydebuglog.Debug("  Get Licensing data: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                If temp <> Trim(CheckDBNull(dr(0), enumObjectType.StrType)) Then
                                    LicCount = LicCount + 1
                                    LIC_DATA(LicCount) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                    LIC_ANSR(LicCount) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                Else
                                    LIC_ANSR(LicCount) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                End If
                                temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType))

                            Catch ex As Exception
                                errmsg = errmsg & "Error reading Licensing data: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                    results = "Failure"
                End Try
                dr.Close()

            End If
            If Debug = "Y" Then mydebuglog.Debug("  RegCount: " & RegCount.ToString)

            ' -----
            ' Get general course information if possible
            '   Replaceable keywords:
            '       [UserId]            -   S_CONTACT.ROW_ID
            '       [RegId]             -   CX_SESS_REG.ROW_ID
            '       [BASE_CONTENT_URL]  -   Base content URL
            '       [BASE_SERVICE_URL]  -   Base services URL
            If CRSE_ID <> "" Then

                Try
                    SqlS = "SELECT CRSE_ID, BASE_CONTENT_URL,BASE_SERVICE_URL,MEDIA_DEST,VIDEO_DEST,FEEDBACK_EMAIL, " & _
                    "FEEDBACK_EMAIL_FROM,FEEDBACK_EMAIL_SUBJ,FEEDBACK_SUCCESS_MSG,SVC_FEEDBACK,GLOSSARY_ENABLED, " & _
                    "SVC_GLOSSARY,KBA_REQD,KBA_FAIL_ALWD,KBA_MAX_TIME,KBA_FAIL_MSG,KBA_EXIT_MSG, " & _
                    "SVC_KBA_QUESTIONS,SVC_KBA_ANSWERS,LESSON_SKINS,LESSON_LAUNCH_LMT,LESSON_LAUNCH_MSG,LESSON_MAX_TIME, " & _
                    "LESSON_MAX_TIME_MSG,LESSON_MIN_TIME,LESSON_MIN_TIME_MSG,LESSON_LINEAR,SVC_LESSON, " & _
                    "QUIZ_FEEDBACK,QUIZ_RANDOM,QUIZ_ATTEMPTS,QUIZ_MIN_SCORE,QUIZ_TIMER_ENABLED,QUIZ_TIMER_LIMIT,SVC_QUIZ, " & _
                    "RESOURCE_DEST,NARRATION,BACKGROUND_LOADER,SHOW_PAGE_COUNT,NUMBER_BY,LOCK_CAP_SLIDE_BUTTONS, " & _
                    "INACT_TIME_LIMIT, INACT_PROMPT_LIMIT, QUIZ_FONT, SHELL_PRELOAD_PCT, RTMP_FOLDER, SCENARIO_DEST, " & _
                    "SCENARIO_ENABLED, MIN_TIME_PER_PAGE, SLP_CONTINUE, SLP_MESSAGE, SLP_RESTART_MSG, SLP_RESTART_ACTION, " & _
                    "MEDIA_X, MEDIA_Y, QUIZ_X, QUIZ_Y, QUIZ_WIDTH, QUIZ_HEIGHT, VIDEO_THUMB_X, VIDEO_THUMB_Y, VIDEO_X, " & _
                    "VIDEO_Y, VIDEO_WIDTH, VIDEO_HEIGHT " & _
                    ",SCRN_SIZE_MIN_M,SCRN_SIZE_MIN_T,SCRN_SIZE_MIN_D,BKG_COLOR_NORMAL,BKG_COLOR_OVER,BKG_COLOR_ACTIVE" & _
                    ",BKG_COLOR_DISABLED,TEXT_COLOR_NORMAL,TEXT_COLOR_OVER,TEXT_COLOR_ACTIVE,TEXT_COLOR_DISABLED,SKIN_COLORS_MAIN_TEXT" & _
                    ",SKIN_COLORS_BACKGROUND,SKIN_COLORS_HEADER_TEXT,SKIN_COLORS_UI_BUTTON_BG,SKIN_COLORS_UI_BUTTON_TEXT,SKIN_COLORS_UI_BUTTON_HL" & _
                    ",SKIN_COLORS_MENU_BUTTON_TEXT,SKIN_COLORS_MENU_BUTTON_HL,SKIN_COLORS_PAGE_COUNTER_TEXT,SKIN_COLORS_QUIZ_UI_BUTTON" & _
                    ",SKIN_COLORS_QUIZ_UI_BUTTON_TEXT,SKIN_COLORS_QUIZ_SCORE_SUM_TEXT,AUTO_ADV_DELAY_SEC,SCORM_VERSION,LSSN_STATUS_TYPE" & _
                    ",MEDIA_LOCKPAGE,QUIZ_REQ_PASS_BEFORE_CONT,QUIZ_ATTMPT_LIMIT,QUIZ_MAX_TIME_LIMIT,QUIZ_QUST_RANDOM,SCORM_OGJ_BYQUIZ,SCENARIO_NARRATION " & _
                    "FROM siebeldb.dbo.CX_CRSE_SCORM_ATTR " & _
                    "WHERE CRSE_ID='" & CRSE_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug("  Get course data: " & SqlS)
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                                If TEST_FLG = "Y" Then
                                    BASE_CONTENT_URL = "http://scorm0.certegrity.com/"
                                    BASE_SERVICE_URL = "http://scorm0.certegrity.com/svc/"
                                Else
                                    BASE_CONTENT_URL = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                    BASE_SERVICE_URL = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                    '3/12/2105; Ren Hou; replace scorm.certegrity.com to be hciscorm.certegrity.com for new SCORM Engine migration
                                    BASE_CONTENT_URL = Replace(BASE_CONTENT_URL, "//scorm.", "//hciscorm.")
                                    BASE_SERVICE_URL = Replace(BASE_SERVICE_URL, "//scorm.", "//hciscorm.")
                                End If

                                'Dim hostname As String
                                'hostname = System.Net.Dns.GetHostName()
                                'BASE_CONTENT_URL = "/"
                                'BASE_SERVICE_URL = "/svc/"

                                ' Media destination
                                MEDIA_DEST = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                temp1 = MEDIA_DEST.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_CONTENT_URL]", BASE_CONTENT_URL)
                                MEDIA_DEST = temp3.Replace("[RegId]", RegId)
                                MEDIA_X = Trim(CheckDBNull(dr(53), enumObjectType.StrType))
                                If MEDIA_X = "" Then MEDIA_X = "42"
                                MEDIA_Y = Trim(CheckDBNull(dr(54), enumObjectType.StrType))
                                If MEDIA_Y = "" Then MEDIA_Y = "125"

                                ' Video destination
                                VIDEO_DEST = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                temp1 = VIDEO_DEST.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_CONTENT_URL]", BASE_CONTENT_URL)
                                VIDEO_DEST = temp3.Replace("[RegId]", RegId)
                                RTMP_FOLDER = Trim(CheckDBNull(dr(45), enumObjectType.StrType))

                                ' Scenario destination
                                SCENARIO_DEST = Trim(CheckDBNull(dr(46), enumObjectType.StrType))
                                temp1 = SCENARIO_DEST.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_SERVICE_URL]", BASE_SERVICE_URL)
                                SCENARIO_DEST = temp3.Replace("[RegId]", RegId)
                                SCENARIO_ENABLED = Trim(CheckDBNull(dr(47), enumObjectType.StrType))
                                If SCENARIO_ENABLED = "Y" Then SCENARIO_ENABLED = "true" Else SCENARIO_ENABLED = "false"

                                ' Feedback
                                FEEDBACK_EMAIL = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                FEEDBACK_EMAIL_FROM = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                                FEEDBACK_EMAIL_SUBJ = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                FEEDBACK_SUCCESS_MSG = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                                SVC_FEEDBACK = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                                temp1 = SVC_FEEDBACK.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_SERVICE_URL]", BASE_SERVICE_URL)
                                SVC_FEEDBACK = temp3.Replace("[RegId]", RegId)

                                ' Glossary
                                GLOSSARY_ENABLED = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                                If GLOSSARY_ENABLED = "Y" Then GLOSSARY_ENABLED = "true" Else GLOSSARY_ENABLED = "false"
                                SVC_GLOSSARY = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                temp1 = SVC_GLOSSARY.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_SERVICE_URL]", BASE_SERVICE_URL)
                                SVC_GLOSSARY = temp3.Replace("[RegId]", RegId)

                                ' KBA
                                If KBA_REQD = "" Then
                                    KBA_REQD = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                                    KBA_FAIL_ALWD = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                                    KBA_MAX_TIME = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                                End If
                                If KBA_REQD = "" Then KBA_REQD = "N"
                                If KBA_REQD = "Y" Then KBA_REQD = "true"
                                If KBA_REQD = "N" Then KBA_REQD = "false"
                                KBA_FAIL_MSG = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                                KBA_EXIT_MSG = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                                SVC_KBA_QUESTIONS = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                                temp1 = SVC_KBA_QUESTIONS.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_SERVICE_URL]", BASE_SERVICE_URL)
                                SVC_KBA_QUESTIONS = temp3.Replace("[RegId]", RegId)
                                SVC_KBA_ANSWERS = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                                temp1 = SVC_KBA_ANSWERS.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_SERVICE_URL]", BASE_SERVICE_URL)
                                SVC_KBA_ANSWERS = temp3.Replace("[RegId]", RegId)

                                ' Lesson attributes
                                LESSON_SKINS = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                                LESSON_LAUNCH_LMT = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                                LESSON_LAUNCH_MSG = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                                MIN_TIME_PER_PAGE = Trim(CheckDBNull(dr(48), enumObjectType.StrType))
                                LESSON_MAX_TIME = Trim(CheckDBNull(dr(22), enumObjectType.StrType))
                                LESSON_MAX_TIME_MSG = Trim(CheckDBNull(dr(23), enumObjectType.StrType))
                                LESSON_MIN_TIME = Trim(CheckDBNull(dr(24), enumObjectType.StrType))
                                LESSON_MIN_TIME_MSG = Trim(CheckDBNull(dr(25), enumObjectType.StrType))
                                LESSON_LINEAR = Trim(CheckDBNull(dr(26), enumObjectType.StrType))
                                If LESSON_LINEAR = "Y" Then LESSON_LINEAR = "true" Else LESSON_LINEAR = "false"
                                SVC_LESSON = Trim(CheckDBNull(dr(27), enumObjectType.StrType))
                                temp1 = SVC_LESSON.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_SERVICE_URL]", BASE_SERVICE_URL)
                                SVC_LESSON = temp3.Replace("[RegId]", RegId)
                                SHELL_PRELOAD_PCT = Trim(CheckDBNull(dr(44), enumObjectType.StrType))
                                INACT_TIME_LIMIT = Trim(CheckDBNull(dr(41), enumObjectType.StrType))
                                If INACT_TIME_LIMIT = "" Then INACT_TIME_LIMIT = "0"
                                INACT_PROMPT_LIMIT = Trim(CheckDBNull(dr(42), enumObjectType.StrType))
                                If INACT_PROMPT_LIMIT = "" Then INACT_PROMPT_LIMIT = "0"
                                SLP_CONTINUE = Trim(CheckDBNull(dr(49), enumObjectType.StrType))
                                SLP_MESSAGE = Trim(CheckDBNull(dr(50), enumObjectType.StrType))
                                SLP_RESTART_MSG = Trim(CheckDBNull(dr(51), enumObjectType.StrType))
                                SLP_RESTART_ACTION = Trim(CheckDBNull(dr(52), enumObjectType.StrType))

                                ' Quiz attributes
                                QUIZ_FONT = Trim(CheckDBNull(dr(43), enumObjectType.StrType))
                                If QUIZ_FONT = "" Then QUIZ_FONT = "Myriad Pro"
                                'If Debug = "Y" Then QUIZ_FONT = "Myriad Pro"
                                QUIZ_FEEDBACK = Trim(CheckDBNull(dr(28), enumObjectType.StrType))
                                If QUIZ_FEEDBACK = "Y" Then
                                    QUIZ_FEEDBACK = "always"
                                Else
                                    If QUIZ_FEEDBACK = "X" Then
                                        QUIZ_FEEDBACK = "lastAttempt"
                                    Else
                                        QUIZ_FEEDBACK = "never"
                                    End If
                                End If
                                QUIZ_RANDOM = Trim(CheckDBNull(dr(29), enumObjectType.StrType))
                                If QUIZ_RANDOM = "Y" Then QUIZ_RANDOM = "true" Else QUIZ_RANDOM = "false"
                                QUIZ_ATTEMPTS = Trim(CheckDBNull(dr(30), enumObjectType.StrType))
                                QUIZ_MIN_SCORE = Trim(CheckDBNull(dr(31), enumObjectType.StrType))
                                QUIZ_TIMER_ENABLED = Trim(CheckDBNull(dr(32), enumObjectType.StrType))
                                If QUIZ_TIMER_ENABLED = "Y" Then QUIZ_TIMER_ENABLED = "true" Else QUIZ_TIMER_ENABLED = "false"
                                QUIZ_TIMER_LIMIT = Trim(CheckDBNull(dr(33), enumObjectType.StrType))
                                SVC_QUIZ = Trim(CheckDBNull(dr(34), enumObjectType.StrType))
                                temp1 = SVC_QUIZ.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_SERVICE_URL]", BASE_SERVICE_URL)
                                SVC_QUIZ = temp3.Replace("[RegId]", RegId)
                                QUIZ_X = Trim(CheckDBNull(dr(55), enumObjectType.StrType))
                                If QUIZ_X = "" Then QUIZ_X = "62"
                                QUIZ_Y = Trim(CheckDBNull(dr(56), enumObjectType.StrType))
                                If QUIZ_Y = "" Then QUIZ_Y = "155"
                                QUIZ_WIDTH = Trim(CheckDBNull(dr(57), enumObjectType.StrType))
                                If QUIZ_WIDTH = "" Then QUIZ_WIDTH = "813"
                                QUIZ_HEIGHT = Trim(CheckDBNull(dr(58), enumObjectType.StrType))
                                If QUIZ_HEIGHT = "" Then QUIZ_HEIGHT = "460"
                                VIDEO_THUMB_X = Trim(CheckDBNull(dr(59), enumObjectType.StrType))
                                If VIDEO_THUMB_X = "" Then VIDEO_THUMB_X = "55"
                                VIDEO_THUMB_Y = Trim(CheckDBNull(dr(60), enumObjectType.StrType))
                                If VIDEO_THUMB_Y = "" Then VIDEO_THUMB_Y = "45"
                                VIDEO_X = Trim(CheckDBNull(dr(61), enumObjectType.StrType))
                                If VIDEO_X = "" Then VIDEO_X = "210"
                                VIDEO_Y = Trim(CheckDBNull(dr(62), enumObjectType.StrType))
                                If VIDEO_Y = "" Then VIDEO_Y = "30"
                                VIDEO_WIDTH = Trim(CheckDBNull(dr(63), enumObjectType.StrType))
                                If VIDEO_WIDTH = "" Then VIDEO_WIDTH = "352"
                                VIDEO_HEIGHT = Trim(CheckDBNull(dr(64), enumObjectType.StrType))
                                If VIDEO_HEIGHT = "" Then VIDEO_HEIGHT = "288"

                                ' Resource destination
                                RESOURCE_DEST = Trim(CheckDBNull(dr(35), enumObjectType.StrType))
                                temp1 = RESOURCE_DEST.Replace("[UserId]", UserId)
                                temp2 = temp1.Replace("[CrseId]", CRSE_ID)
                                temp3 = temp2.Replace("[BASE_CONTENT_URL]", BASE_CONTENT_URL)
                                RESOURCE_DEST = temp3.Replace("[RegId]", RegId)

                                ' Misc.
                                SHOW_VOLUME_CONTROL = Trim(CheckDBNull(dr(36), enumObjectType.StrType))
                                If SHOW_VOLUME_CONTROL = "Y" Then SHOW_VOLUME_CONTROL = "true" Else SHOW_VOLUME_CONTROL = "false"
                                BACKGROUND_LOADER = Trim(CheckDBNull(dr(37), enumObjectType.StrType))
                                If BACKGROUND_LOADER = "Y" Then BACKGROUND_LOADER = "true" Else BACKGROUND_LOADER = "false"
                                SHOW_PAGE_COUNT = Trim(CheckDBNull(dr(38), enumObjectType.StrType))
                                If SHOW_PAGE_COUNT = "Y" Then SHOW_PAGE_COUNT = "true" Else SHOW_PAGE_COUNT = "false"
                                NUMBER_BY = Trim(CheckDBNull(dr(39), enumObjectType.StrType))
                                LOCK_CAP_SLIDE_BUTTONS = Trim(CheckDBNull(dr(40), enumObjectType.StrType))
                                If LOCK_CAP_SLIDE_BUTTONS = "Y" Then LOCK_CAP_SLIDE_BUTTONS = "true" Else LOCK_CAP_SLIDE_BUTTONS = "false"

                                'New variables for HTML5 config
                                SCRN_SIZE_MIN_M = Trim(CheckDBNull(dr("SCRN_SIZE_MIN_M"), enumObjectType.StrType))
                                SCRN_SIZE_MIN_T = Trim(CheckDBNull(dr("SCRN_SIZE_MIN_T"), enumObjectType.StrType))
                                SCRN_SIZE_MIN_D = Trim(CheckDBNull(dr("SCRN_SIZE_MIN_D"), enumObjectType.StrType))
                                BKG_COLOR_NORMAL = Trim(CheckDBNull(dr("BKG_COLOR_NORMAL"), enumObjectType.StrType))
                                BKG_COLOR_OVER = Trim(CheckDBNull(dr("BKG_COLOR_OVER"), enumObjectType.StrType))
                                BKG_COLOR_ACTIVE = Trim(CheckDBNull(dr("BKG_COLOR_ACTIVE"), enumObjectType.StrType))
                                BKG_COLOR_DISABLED = Trim(CheckDBNull(dr("BKG_COLOR_DISABLED"), enumObjectType.StrType))
                                TEXT_COLOR_NORMAL = Trim(CheckDBNull(dr("TEXT_COLOR_NORMAL"), enumObjectType.StrType))
                                TEXT_COLOR_OVER = Trim(CheckDBNull(dr("TEXT_COLOR_OVER"), enumObjectType.StrType))
                                TEXT_COLOR_ACTIVE = Trim(CheckDBNull(dr("TEXT_COLOR_ACTIVE"), enumObjectType.StrType))
                                TEXT_COLOR_DISABLED = Trim(CheckDBNull(dr("TEXT_COLOR_DISABLED"), enumObjectType.StrType))
                                SKIN_COLORS_MAIN_TEXT = Trim(CheckDBNull(dr("SKIN_COLORS_MAIN_TEXT"), enumObjectType.StrType))
                                SKIN_COLORS_BACKGROUND = Trim(CheckDBNull(dr("SKIN_COLORS_BACKGROUND"), enumObjectType.StrType))
                                SKIN_COLORS_HEADER_TEXT = Trim(CheckDBNull(dr("SKIN_COLORS_HEADER_TEXT"), enumObjectType.StrType))
                                SKIN_COLORS_UI_BUTTON_BG = Trim(CheckDBNull(dr("SKIN_COLORS_UI_BUTTON_BG"), enumObjectType.StrType))
                                SKIN_COLORS_UI_BUTTON_TEXT = Trim(CheckDBNull(dr("SKIN_COLORS_UI_BUTTON_TEXT"), enumObjectType.StrType))
                                SKIN_COLORS_UI_BUTTON_HL = Trim(CheckDBNull(dr("SKIN_COLORS_UI_BUTTON_HL"), enumObjectType.StrType))
                                SKIN_COLORS_MENU_BUTTON_TEXT = Trim(CheckDBNull(dr("SKIN_COLORS_MENU_BUTTON_TEXT"), enumObjectType.StrType))
                                SKIN_COLORS_MENU_BUTTON_HL = Trim(CheckDBNull(dr("SKIN_COLORS_MENU_BUTTON_HL"), enumObjectType.StrType))
                                SKIN_COLORS_PAGE_COUNTER_TEXT = Trim(CheckDBNull(dr("SKIN_COLORS_PAGE_COUNTER_TEXT"), enumObjectType.StrType))
                                SKIN_COLORS_QUIZ_UI_BUTTON = Trim(CheckDBNull(dr("SKIN_COLORS_QUIZ_UI_BUTTON"), enumObjectType.StrType))
                                SKIN_COLORS_QUIZ_UI_BUTTON_TEXT = Trim(CheckDBNull(dr("SKIN_COLORS_QUIZ_UI_BUTTON_TEXT"), enumObjectType.StrType))
                                SKIN_COLORS_QUIZ_SCORE_SUM_TEXT = Trim(CheckDBNull(dr("SKIN_COLORS_QUIZ_SCORE_SUM_TEXT"), enumObjectType.StrType))
                                AUTO_ADV_DELAY_SEC = Trim(CheckDBNull(dr("AUTO_ADV_DELAY_SEC"), enumObjectType.StrType))
                                SCORM_VERSION = Trim(CheckDBNull(dr("SCORM_VERSION"), enumObjectType.StrType))
                                LSSN_STATUS_TYPE = Trim(CheckDBNull(dr("LSSN_STATUS_TYPE"), enumObjectType.StrType))
                                MEDIA_LOCKPAGE = Trim(CheckDBNull(dr("MEDIA_LOCKPAGE"), enumObjectType.StrType))
                                QUIZ_REQ_PASS_BEFORE_CONT = Trim(CheckDBNull(dr("QUIZ_REQ_PASS_BEFORE_CONT"), enumObjectType.StrType))
                                QUIZ_ATTMPT_LIMIT = Trim(CheckDBNull(dr("QUIZ_ATTMPT_LIMIT"), enumObjectType.StrType))
                                QUIZ_MAX_TIME_LIMIT = Trim(CheckDBNull(dr("QUIZ_MAX_TIME_LIMIT"), enumObjectType.StrType))
                                QUIZ_QUST_RANDOM = Trim(CheckDBNull(dr("QUIZ_QUST_RANDOM"), enumObjectType.StrType))
                                SCORM_OGJ_BYQUIZ = Trim(CheckDBNull(dr("SCORM_OGJ_BYQUIZ"), enumObjectType.StrType))
                                SCENARIO_NARRATION = Trim(CheckDBNull(dr("SCENARIO_NARRATION"), enumObjectType.StrType))
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
            End If
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' ============================================
        ' Return the configuration data for the course
        '  If scormlocal is set to "Y" then web service calls for lesson.xml, questions.xml, resources, assets and videos are disabled
        '  All of these things will be served from the local package.  This allows the environment to be 
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        Dim resultsItem As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("config")
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        Try
            ' Add result items 
            If Debug <> "T" Then
                If CRSE_ID <> "" Then
                    ' Declarations
                    Dim XmlLevelOne As System.Xml.XmlElement
                    Dim XmlLevelTwo As System.Xml.XmlElement
                    Dim XmlLevelThree As System.Xml.XmlElement

                    ' Feedback window
                    '  <feedbackWindow successMessage="Your feedback has been received.\nPlease close this window to\nreturn to the course.">
                    '      <webservice wsdl="http://hciscorm.certegrity.com/svc/service.asmx?wsdl" operation="SaveFeedback" />
                    '  </feedbackWindow>
                    resultsItem = odoc.CreateElement("feedbackWindow")
                    AddXMLAttribute(odoc, resultsItem, "enabled", "true")
                    AddXMLAttribute(odoc, resultsItem, "successMessage", FEEDBACK_SUCCESS_MSG)
                    XmlLevelOne = odoc.CreateElement("webservice")
                    AddXMLAttribute(odoc, XmlLevelOne, "wsdl", SVC_FEEDBACK & "?wsdl")
                    AddXMLAttribute(odoc, XmlLevelOne, "operation", "SaveFeedback")
                    resultsItem.AppendChild(XmlLevelOne)
                    resultsRoot.AppendChild(resultsItem)

                    ' Background Loader
                    '   <backgroundLoader enabled="true" />
                    resultsItem = odoc.CreateElement("backgroundLoader")
                    AddXMLAttribute(odoc, resultsItem, "enabled", BACKGROUND_LOADER)
                    resultsRoot.AppendChild(resultsItem)

                    ' Navigator
                    '   <bottomNav showPageCount="true" numberBy="allButLastChapter" lockCaptivateSlideButtons="true" showVolumeControl="false" />
                    resultsItem = odoc.CreateElement("bottomNav")
                    AddXMLAttribute(odoc, resultsItem, "showPageCount", SHOW_PAGE_COUNT)
                    AddXMLAttribute(odoc, resultsItem, "numberBy", NUMBER_BY)
                    AddXMLAttribute(odoc, resultsItem, "lockCaptivateSlideButtons", LOCK_CAP_SLIDE_BUTTONS)
                    AddXMLAttribute(odoc, resultsItem, "showVolumeControl", SHOW_VOLUME_CONTROL)
                    resultsRoot.AppendChild(resultsItem)

                    ' Scenario
                    '   <scenario>
                    '       <webservice enabled=�true|false� wsdl=�http://hciscorm.certegrity.com/svc/service.asmx?wsdl� operation=�GetScenarioData� />
                    '   </scenario>
                    If SCENARIO_NARRATION = "" Then SCENARIO_NARRATION = "false"
                    resultsItem = odoc.CreateElement("scenario")

                    XmlLevelOne = odoc.CreateElement("webservice")
                    AddXMLAttribute(odoc, XmlLevelOne, "enabled", SCENARIO_ENABLED)
                    AddXMLAttribute(odoc, XmlLevelOne, "wsdl", SCENARIO_DEST & "?wsdl")
                    AddXMLAttribute(odoc, XmlLevelOne, "operation", "GetScenarioData")
                    resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("scenarioNarration")
                    AddXMLAttribute(odoc, XmlLevelOne, "enabled", SCENARIO_NARRATION)
                    resultsItem.AppendChild(XmlLevelOne)

                    resultsRoot.AppendChild(resultsItem)

                    ' Glossary
                    '  <glossary>
                    '      <webservice enabled="true" wsdl = "http://hciscorm.certegrity.com/svc/service.asmx?wsdl" operation="CourseGlossary" />
                    '  </glossary>
                    resultsItem = odoc.CreateElement("glossary")
                    AddXMLAttribute(odoc, resultsItem, "enabled", GLOSSARY_ENABLED)
                    XmlLevelOne = odoc.CreateElement("webservice")
                    AddXMLAttribute(odoc, XmlLevelOne, "enabled", GLOSSARY_ENABLED)
                    AddXMLAttribute(odoc, XmlLevelOne, "wsdl", SVC_GLOSSARY & "?wsdl")
                    AddXMLAttribute(odoc, XmlLevelOne, "operation", "CourseGlossary")
                    resultsItem.AppendChild(XmlLevelOne)
                    XmlLevelOne = odoc.CreateElement("letterBarTextColor")
                    AddXMLAttribute(odoc, XmlLevelOne, "active", "0xffffff")
                    AddXMLAttribute(odoc, XmlLevelOne, "used", "0x000000")
                    AddXMLAttribute(odoc, XmlLevelOne, "unused", "0x00441d")
                    resultsItem.AppendChild(XmlLevelOne)
                    resultsRoot.AppendChild(resultsItem)

                    ' KBA
                    '  <kba enabled="true" failureLimit="3" failureMessage="You have failed. Please try again." exitMessage="You have reached your failure limit. The lesson is exiting."> 
                    '      <questions>
                    '         <webservice wsdl="http://hciscorm.certegrity.com/svc/service.asmx?wsdl" operation="KBAQLookup" />
                    '      </questions>
                    '      <answer>
                    '         <webservice wsdl="http://hciscorm.certegrity.com/svc/service.asmx?wsdl" operation="KBAALookup" />
                    '      </answer>
                    '  </kba>
                    resultsItem = odoc.CreateElement("kba")
                    AddXMLAttribute(odoc, resultsItem, "enabled", KBA_REQD)
                    AddXMLAttribute(odoc, resultsItem, "testMode", "false")
                    AddXMLAttribute(odoc, resultsItem, "failureLimit", KBA_FAIL_ALWD)
                    AddXMLAttribute(odoc, resultsItem, "timeLimit", KBA_MAX_TIME)
                    AddXMLAttribute(odoc, resultsItem, "failureMessage", KBA_FAIL_MSG)
                    AddXMLAttribute(odoc, resultsItem, "timeoutExitMessage", "The time limit to answer a personal verification question has expired, and so must exit you from this course.")
                    AddXMLAttribute(odoc, resultsItem, "exitMessage", KBA_EXIT_MSG)

                    XmlLevelOne = odoc.CreateElement("questions")
                    XmlLevelTwo = odoc.CreateElement("webservice")
                    AddXMLAttribute(odoc, XmlLevelTwo, "wsdl", SVC_KBA_QUESTIONS & "?wsdl")
                    AddXMLAttribute(odoc, XmlLevelTwo, "operation", "KBAQLookup")
                    XmlLevelOne.AppendChild(XmlLevelTwo)
                    resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("answer")
                    XmlLevelTwo = odoc.CreateElement("webservice")
                    AddXMLAttribute(odoc, XmlLevelTwo, "wsdl", SVC_KBA_ANSWERS & "?wsdl")
                    AddXMLAttribute(odoc, XmlLevelTwo, "operation", "KBAALookup")
                    XmlLevelOne.AppendChild(XmlLevelTwo)
                    resultsItem.AppendChild(XmlLevelOne)
                    resultsRoot.AppendChild(resultsItem)

                    ' Lesson
                    '  <lesson skinFolder="skins/default/" launchLimit="2" launchLimitMessage="You have reached the launch limit for this lesson. Click OK to close the window." maximumTimeLimit="10800" maximumTimeLimitMessage="Your 3 hour time limit time has expired. The lesson will now exit." minimumTimeLimit="7200" minimumTimeLimitMessage="You must spend at least 2 hours in this lesson before exiting." requireLinearProgression="true">  
                    '	    <webservice enabled="true" wsdl="http://hciscorm.certegrity.com/svc/service.asmx?wsdl" operation="GetLessonData" />
                    '  </lesson>
                    resultsItem = odoc.CreateElement("lesson")
                    AddXMLAttribute(odoc, resultsItem, "skinFolder", LESSON_SKINS)
                    AddXMLAttribute(odoc, resultsItem, "launchLimit", LESSON_LAUNCH_LMT)
                    AddXMLAttribute(odoc, resultsItem, "launchLimitMessage", LESSON_LAUNCH_MSG)
                    AddXMLAttribute(odoc, resultsItem, "minimumTimeLimitPerPage", MIN_TIME_PER_PAGE)
                    AddXMLAttribute(odoc, resultsItem, "maximumTimeLimit", LESSON_MAX_TIME)
                    AddXMLAttribute(odoc, resultsItem, "maximumTimeLimitMessage", LESSON_MAX_TIME_MSG)
                    AddXMLAttribute(odoc, resultsItem, "minimumTimeLimit", LESSON_MIN_TIME)
                    AddXMLAttribute(odoc, resultsItem, "minimumTimeLimitMessage", LESSON_MIN_TIME_MSG)
                    AddXMLAttribute(odoc, resultsItem, "inactivityTimeLimit", INACT_TIME_LIMIT)
                    AddXMLAttribute(odoc, resultsItem, "inactivityPromptTimeLimit", INACT_PROMPT_LIMIT)
                    AddXMLAttribute(odoc, resultsItem, "requireLinearProgression", LESSON_LINEAR)
                    AddXMLAttribute(odoc, resultsItem, "shellPreloadPercent", SHELL_PRELOAD_PCT)

                    XmlLevelOne = odoc.CreateElement("webservice")
                    If scormlocal = "Y" Then
                        AddXMLAttribute(odoc, XmlLevelOne, "enabled", "false")
                    Else
                        AddXMLAttribute(odoc, XmlLevelOne, "enabled", "true")
                    End If
                    AddXMLAttribute(odoc, XmlLevelOne, "wsdl", SVC_LESSON & "?wsdl")
                    AddXMLAttribute(odoc, XmlLevelOne, "operation", "GetLessonData")
                    resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("subsequentLaunchPrompt")
                    AddXMLAttribute(odoc, XmlLevelOne, "message", SLP_MESSAGE)
                    AddXMLAttribute(odoc, XmlLevelOne, "continueButtonLabel", SLP_CONTINUE)
                    AddXMLAttribute(odoc, XmlLevelOne, "restartButtonLabel", SLP_RESTART_MSG)
                    AddXMLAttribute(odoc, XmlLevelOne, "restartButtonAction", SLP_RESTART_ACTION)
                    resultsItem.AppendChild(XmlLevelOne)
                    resultsRoot.AppendChild(resultsItem)

                    ' Media
                    '  <media folder="media/[lang]/" />
                    '  Fix folder name
                    If scormlocal <> "Y" Then
                        resultsItem = odoc.CreateElement("media")
                        AddXMLAttribute(odoc, resultsItem, "folder", MEDIA_DEST)
                        AddXMLAttribute(odoc, resultsItem, "x", MEDIA_X)
                        AddXMLAttribute(odoc, resultsItem, "y", MEDIA_Y)
                        AddXMLAttribute(odoc, resultsItem, "lockRoot", "false")
                        AddXMLAttribute(odoc, resultsItem, "preloadPercent", "50")
                        '8/4/15;Ren Hou; Added for new HTML5 config
                        AddXMLAttribute(odoc, resultsItem, "lockPages", MEDIA_LOCKPAGE)
                        resultsRoot.AppendChild(resultsItem)
                    End If

                    ' Quiz
                    '  <quiz>
                    '       <feedback showOn="always"/>
                    '       <questions randomize="false" attemptLimit="2" minPassingScore="80.000000"/>
                    '       <questionTimer enabled="false" timeLimit="0"/>
                    '       <webservice enabled="true" wsdl="http://hciscorm.certegrity.com/svc/service.asmx?wsdl" operation="QuizLibrary"/>
                    '  </quiz>
                    resultsItem = odoc.CreateElement("quiz")
                    AddXMLAttribute(odoc, resultsItem, "x", QUIZ_X)
                    AddXMLAttribute(odoc, resultsItem, "y", QUIZ_Y)
                    AddXMLAttribute(odoc, resultsItem, "width", QUIZ_WIDTH)
                    AddXMLAttribute(odoc, resultsItem, "height", QUIZ_HEIGHT)
                    '8/4/15;Ren Hou; Added for new HTML5 config
                    AddXMLAttribute(odoc, resultsItem, "requirePassBeforeContinue", QUIZ_ATTMPT_LIMIT)
                    AddXMLAttribute(odoc, resultsItem, "quizAttemptLimit", QUIZ_ATTMPT_LIMIT)
                    AddXMLAttribute(odoc, resultsItem, "quizMaximumTimeLimit", QUIZ_MAX_TIME_LIMIT)

                    XmlLevelOne = odoc.CreateElement("audio")
                    AddXMLAttribute(odoc, XmlLevelOne, "question", "false")
                    AddXMLAttribute(odoc, XmlLevelOne, "feedback", "false")
                    resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("autoAdvance")
                    AddXMLAttribute(odoc, XmlLevelOne, "after", "none")
                    'AddXMLAttribute(odoc, XmlLevelOne, "delaySeconds", "1")
                    '8/4/15;Ren Hou; Modified for new HTML5 config
                    AUTO_ADV_DELAY_SEC = IIf(AUTO_ADV_DELAY_SEC.Length() > 0, AUTO_ADV_DELAY_SEC, "1")
                    AddXMLAttribute(odoc, XmlLevelOne, "delaySeconds", AUTO_ADV_DELAY_SEC)
                    resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("feedback")
                    AddXMLAttribute(odoc, XmlLevelOne, "showOn", QUIZ_FEEDBACK)
                    AddXMLAttribute(odoc, XmlLevelOne, "position", "underButtons")
                    resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("questions")
                    AddXMLAttribute(odoc, XmlLevelOne, "randomize", QUIZ_RANDOM)
                    AddXMLAttribute(odoc, XmlLevelOne, "attemptLimit", QUIZ_ATTEMPTS)
                    AddXMLAttribute(odoc, XmlLevelOne, "minPassingScore", QUIZ_MIN_SCORE)
                    AddXMLAttribute(odoc, XmlLevelOne, "hideQuestionNumbers", "true")
                    AddXMLAttribute(odoc, XmlLevelOne, "onNextIfPrevIsIncorrectReplaceItAndGoBack", "false")
                    '8/4/15;Ren Hou; Added for new HTML5 config
                    AddXMLAttribute(odoc, XmlLevelOne, "randomizeQuestions", QUIZ_QUST_RANDOM)

                    resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("questionTimer")
                    AddXMLAttribute(odoc, XmlLevelOne, "enabled", QUIZ_TIMER_ENABLED)
                    AddXMLAttribute(odoc, XmlLevelOne, "timeLimit", QUIZ_TIMER_LIMIT)
                    resultsItem.AppendChild(XmlLevelOne)
                    'XmlLevelOne = odoc.CreateElement("summary")
                    'AddXMLAttribute(odoc, XmlLevelOne, "enabled", "true")
                    'AddXMLAttribute(odoc, XmlLevelOne, "title", "Quiz Results")
                    'AddXMLAttribute(odoc, XmlLevelOne, "passMessage", "Click NEXT to go on.")
                    'AddXMLAttribute(odoc, XmlLevelOne, "failMessage", "")
                    'AddXMLAttribute(odoc, XmlLevelOne, "allowResetLessonOnFail", "false")
                    'AddXMLAttribute(odoc, XmlLevelOne, "allowRetryQuizOnFail", "true")
                    'AddXMLAttribute(odoc, XmlLevelOne, "resetLessonButtonLabel", "REVIEW SECTIONS")
                    'AddXMLAttribute(odoc, XmlLevelOne, "retryQuizButtonLabel", "RETAKE QUIZ")
                    'AddXMLAttribute(odoc, XmlLevelOne, "titleOfNextLesson", "")
                    'AddXMLAttribute(odoc, XmlLevelOne, "titleOfCourse", "")
                    'AddXMLAttribute(odoc, XmlLevelOne, "enableNextButton", "whenPassed")
                    'resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("text")
                    AddXMLAttribute(odoc, XmlLevelOne, "color", "0x333333")
                    AddXMLAttribute(odoc, XmlLevelOne, "size", "18")
                    AddXMLAttribute(odoc, XmlLevelOne, "font", QUIZ_FONT)
                    AddXMLAttribute(odoc, XmlLevelOne, "dropshadow", "false")
                    resultsItem.AppendChild(XmlLevelOne)

                    XmlLevelOne = odoc.CreateElement("videoQuestion")
                    AddXMLAttribute(odoc, XmlLevelOne, "thumbnailX", VIDEO_THUMB_X)
                    AddXMLAttribute(odoc, XmlLevelOne, "thumbnailY", VIDEO_THUMB_Y)
                    AddXMLAttribute(odoc, XmlLevelOne, "videoX", VIDEO_X)
                    AddXMLAttribute(odoc, XmlLevelOne, "videoY", VIDEO_Y)
                    AddXMLAttribute(odoc, XmlLevelOne, "videoWidth", VIDEO_WIDTH)
                    AddXMLAttribute(odoc, XmlLevelOne, "videoHeight", VIDEO_HEIGHT)
                    resultsItem.AppendChild(XmlLevelOne)
                    XmlLevelOne = odoc.CreateElement("webservice")
                    If scormlocal = "Y" Then
                        AddXMLAttribute(odoc, XmlLevelOne, "enabled", "false")
                    Else
                        AddXMLAttribute(odoc, XmlLevelOne, "enabled", "true")
                    End If
                    AddXMLAttribute(odoc, XmlLevelOne, "wsdl", SVC_QUIZ & "?wsdl")
                    AddXMLAttribute(odoc, XmlLevelOne, "operation", "QuizLibrary")
                    resultsItem.AppendChild(XmlLevelOne)
                    resultsRoot.AppendChild(resultsItem)

                    ' Video
                    ' If the local config.xml rtmpFolder is set then it overrides this value
                    '  <video folder="video/[lang]/" />
                    If scormlocal <> "Y" Then
                        resultsItem = odoc.CreateElement("video")
                        AddXMLAttribute(odoc, resultsItem, "folder", VIDEO_DEST)
                        If RTMP_FOLDER <> "" Then AddXMLAttribute(odoc, resultsItem, "rtmpFolder", RTMP_FOLDER)
                        resultsRoot.AppendChild(resultsItem)
                    End If

                    ' Resources
                    '  <resources folder="video/[lang]/" />
                    If scormlocal <> "Y" Then
                        resultsItem = odoc.CreateElement("resources")
                        AddXMLAttribute(odoc, resultsItem, "folder", RESOURCE_DEST)
                        resultsRoot.AppendChild(resultsItem)
                    End If

                    ' Jurisdiction
                    ' <jurisdiction name="Texas"
                    '	city=""/
                    '	county=""/
                    '	state="TX"/
                    '	country="TX"/>
                    resultsItem = odoc.CreateElement("jurisdiction")
                    AddXMLAttribute(odoc, resultsItem, "name", JURISDICTION)
                    AddXMLAttribute(odoc, resultsItem, "city", CITY)
                    AddXMLAttribute(odoc, resultsItem, "county", COUNTY)
                    AddXMLAttribute(odoc, resultsItem, "state", STATE)
                    AddXMLAttribute(odoc, resultsItem, "country", COUNTRY)
                    ' <acceptableIds anyStateId="" ... />
                    If LicCount > 0 Then
                        XmlLevelOne = odoc.CreateElement("acceptableIds")
                        For i = 1 To LicCount
                            AddXMLAttribute(odoc, XmlLevelOne, LIC_DATA(i), LIC_ANSR(i))
                        Next
                        'XmlLevelOne.AppendChild(XmlLevelTwo)
                        resultsItem.AppendChild(XmlLevelOne)
                    End If
                    ' </acceptableIds>
                    ' <alcoholLaws legalAgeToSell="" ...
                    'If RegCount > 0 Then
                    XmlLevelOne = odoc.CreateElement("alcoholLaws")
                    For i = 1 To RegCount
                        AddXMLAttribute(odoc, XmlLevelOne, REG_DATA(i), REG_QUES(i))
                    Next
                    resultsItem.AppendChild(XmlLevelOne)
                    '                    End If
                    ' </alcoholLaws>
                    resultsRoot.AppendChild(resultsItem)
                    ' </jurisdiction>

                    '8/4/15;Ren Hou; Added new Tags for new HTML5 config
                    '<responsiveDesign>
                    resultsItem = odoc.CreateElement("responsiveDesign")
                    XmlLevelOne = odoc.CreateElement("screenSize")
                    AddXMLAttribute(odoc, XmlLevelOne, "minWidthMobile", SCRN_SIZE_MIN_M)
                    AddXMLAttribute(odoc, XmlLevelOne, "minWidthTablet", SCRN_SIZE_MIN_T)
                    AddXMLAttribute(odoc, XmlLevelOne, "minWidthDesktop", SCRN_SIZE_MIN_D)
                    resultsItem.AppendChild(XmlLevelOne)
                    resultsRoot.AppendChild(resultsItem)

                    '8/4/15;Ren Hou; Added new Tags for new HTML5 config
                    '<leftMenu>
                    resultsItem = odoc.CreateElement("leftMenu")
                    XmlLevelOne = odoc.CreateElement("item")
                    XmlLevelTwo = odoc.CreateElement("bkgColor")
                    AddXMLAttribute(odoc, XmlLevelTwo, "normal", BKG_COLOR_NORMAL)
                    AddXMLAttribute(odoc, XmlLevelTwo, "over", BKG_COLOR_OVER)
                    AddXMLAttribute(odoc, XmlLevelTwo, "active", BKG_COLOR_ACTIVE)
                    AddXMLAttribute(odoc, XmlLevelTwo, "disabled", BKG_COLOR_DISABLED)
                    XmlLevelOne.AppendChild(XmlLevelTwo)
                    XmlLevelTwo = odoc.CreateElement("textColor")
                    AddXMLAttribute(odoc, XmlLevelTwo, "normal", TEXT_COLOR_NORMAL)
                    AddXMLAttribute(odoc, XmlLevelTwo, "over", TEXT_COLOR_OVER)
                    AddXMLAttribute(odoc, XmlLevelTwo, "active", TEXT_COLOR_ACTIVE)
                    AddXMLAttribute(odoc, XmlLevelTwo, "disabled", TEXT_COLOR_DISABLED)
                    XmlLevelOne.AppendChild(XmlLevelTwo)
                    resultsItem.AppendChild(XmlLevelOne)
                    resultsRoot.AppendChild(resultsItem)

                    '8/4/15;Ren Hou; Added new Tags for new HTML5 config
                    '<skinColors>
                    resultsItem = odoc.CreateElement("skinColors")
                    AddXMLAttribute(odoc, resultsItem, "mainText", SKIN_COLORS_MAIN_TEXT)
                    AddXMLAttribute(odoc, resultsItem, "background", SKIN_COLORS_BACKGROUND)
                    AddXMLAttribute(odoc, resultsItem, "headerText", SKIN_COLORS_HEADER_TEXT)
                    AddXMLAttribute(odoc, resultsItem, "uiButtonBackground", SKIN_COLORS_UI_BUTTON_BG)
                    AddXMLAttribute(odoc, resultsItem, "uiButtonText", SKIN_COLORS_UI_BUTTON_TEXT)
                    AddXMLAttribute(odoc, resultsItem, "uiButtonHighlight", SKIN_COLORS_UI_BUTTON_HL)
                    AddXMLAttribute(odoc, resultsItem, "menuButtonText", SKIN_COLORS_MENU_BUTTON_TEXT)
                    AddXMLAttribute(odoc, resultsItem, "menuButtonTextHighlight", SKIN_COLORS_MENU_BUTTON_HL)
                    AddXMLAttribute(odoc, resultsItem, "pageCounterText", SKIN_COLORS_PAGE_COUNTER_TEXT)
                    AddXMLAttribute(odoc, resultsItem, "quizUiButton", SKIN_COLORS_QUIZ_UI_BUTTON)
                    AddXMLAttribute(odoc, resultsItem, "quizUiButtonText", SKIN_COLORS_QUIZ_UI_BUTTON_TEXT)
                    AddXMLAttribute(odoc, resultsItem, "quizScoreSummaryText", SKIN_COLORS_QUIZ_SCORE_SUM_TEXT)
                    resultsRoot.AppendChild(resultsItem)

                End If
            End If
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

        ' Debug output configuration document
        If Debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "Results: " & vbCrLf & odoc.InnerXml)
        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetConfiguration : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GetConfiguration : Results: " & results & " by UserId # " & DecodedUserId & " with RegId " & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & results & " by UserId # " & DecodedUserId & " with RegId " & RegId & " at " & Now.ToString)
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
        Return odoc
    End Function

    <WebMethod(Description:="Retrieves an XML representation of course content")> _
    Public Function GetLessonData(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function locates the XML data for a course and returns it to the invoker 

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	        - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results, temp As String
        Dim mypath, errmsg, logging As String
        Dim DecodedUserId, ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' File handling declarations
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim abyBuffer(1000) As Byte
        Dim d_dsize As String
        Dim CRSE_ID, TEST_FLG As String
        Dim filecache As ObjectCache = MemoryCache.Default
        Dim fileContents As String
        Dim cached As Boolean

        ' Minio declarations
        Dim minio_flg, d_verid, d_doc_id As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

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
        TEST_FLG = ""
        d_dsize = ""
        BinaryFile = ""
        temp = ""
        fileContents = ""
        minio_flg = ""
        d_verid = ""
        d_doc_id = ""
        cached = False

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetLessonData_debug")
            If temp = "Y" And Debug <> "T" And Debug <> "R" Then Debug = "Y"
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
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        'Debug = "Y"
        If RegId = "" And Debug = "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetLessonData.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Open database connections
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
        ' Get course information
        If Not cmd Is Nothing Then
            ' -----
            ' Query registration
            Try
                SqlS = "SELECT R.CRSE_ID, C.X_REGISTRATION_NUM, R.TEST_FLG " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Get registration: " & vbCrLf & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            CRSE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            ValidatedUserId = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            TEST_FLG = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
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
            If Debug = "Y" Then
                mydebuglog.Debug("   ... CRSE_ID: " & CRSE_ID)
                mydebuglog.Debug("   ... TEST_FLG: " & TEST_FLG)
            End If

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

        ' ============================================
        '' Create output directory
        'SaveDest = mypath & "course_temp"
        'Try
        '    Directory.CreateDirectory(SaveDest)
        'Catch
        'End Try

        ' ============================================
        Dim last_upt As Date = Today.AddYears(-50)
        ' Check to see if the document is in the in-memory cache
        'If Not IsNothing(filecache("lesson-" & CRSE_ID)) Then
        'Check if the cached item need to be renewed; v.last_upd; 2/27/17; Ren Hou;
        'SqlS = "SELECT TOP 1 v.last_upd " & _
        '    "FROM DMS.dbo.Documents d " & _
        '    "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " & _
        '    "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " & _
        '    "WHERE d.ext_id='" & CRSE_ID & "' and dc.cat_id=6 ORDER BY v.version DESC, d.last_upd DESC"
        'Try
        '    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check last updated date for xml document: " & vbCrLf & SqlS)
        '    d_cmd.CommandText = SqlS
        '    last_upt = d_cmd.ExecuteScalar()
        'Catch ex As Exception
        '    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
        '    results = "Failure"
        'End Try
        '
        'If last_upt > TryCast(filecache("lesson-" & CRSE_ID), HciDMSDocument).UpdateDate Then
        '    'Remove if the update_date on the cache is before the last updtaed date on DB record.
        '    filecache.Remove("lesson-" & CRSE_ID)
        '    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Removing Cached object " & "lesson-" & CRSE_ID)
        'End If
        'End If
        '  *********** ''
        'Load content of cached object
        If filecache("lesson-" & Trim(CRSE_ID)) Is Nothing Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Cannot retrieve cached object " & "lesson-" & Trim(CRSE_ID))
            fileContents = ""
        Else
            cached = True
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Retrieved Cached object " & "lesson-" & Trim(CRSE_ID))
            fileContents = TryCast(filecache("lesson-" & Trim(CRSE_ID)), HciDMSDocument).CachedObj
            If Debug = "Y" Then mydebuglog.Debug("  ... length: " & fileContents.Length.ToString)
        End If

        If fileContents Is Nothing Or Debug = "R" Or fileContents = "" Then
            ' ============================================
            ' Get document containing XML and store in a local temp file                
            If Not d_cmd Is Nothing Then

                ' If in test mode, then look for a resource with the name "test-lesson.xml"
                If TEST_FLG = "Y" Then
                    SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.minio_flg, v.row_id, d.row_id " &
                        "FROM DMS.dbo.Documents d " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id " &
                        "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=10 " &
                        " and (d.dfilename='test-lesson.xml' or d.name='test-lesson.xml') " &
                        "AND d.deleted is null " &
                        "ORDER BY v.version DESC, d.last_upd DESC"
                    Try
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get test xml document: " & vbCrLf & SqlS)
                        d_cmd.CommandText = SqlS
                        d_dr = d_cmd.ExecuteReader()
                        If Not d_dr Is Nothing Then
                            While d_dr.Read()
                                Try
                                    d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                    minio_flg = Trim(CheckDBNull(d_dr(2), enumObjectType.StrType))
                                    d_verid = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                                    d_doc_id = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                    If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_doc_id=" & d_doc_id & ",  d_verid=" & d_verid & ",  d_dsize=" & d_dsize & ",  minio_flg=" & minio_flg)

                                    If minio_flg = "Y" Then
                                        If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Minio")
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
                                                ReDim outbyte(Val(retval - 1))
                                                Dim intval As Integer
                                                For i = 0 To retval - 1
                                                    intval = mobj2.ResponseStream.ReadByte()
                                                    If intval < 255 And intval > 0 Then
                                                        outbyte(i) = intval
                                                    End If
                                                    If intval = 255 Then outbyte(i) = 255
                                                    If intval < 0 Then
                                                        outbyte(i) = 0
                                                        If Debug = "Y" Then mydebuglog.Debug("  >  .. " & i.ToString & "   intval: " & intval.ToString)
                                                    End If
                                                Next
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
                                        If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Document_Versions")
                                        ' Get binary and attach to the object outbyte
                                        If d_dsize <> "" Then
                                            ReDim outbyte(Val(d_dsize) - 1)
                                            startIndex = 0
                                            retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                        End If
                                    End If

                                Catch ex As Exception
                                    results = "Failure"
                                    errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                    GoTo CloseOut
                                End Try
                            End While
                        Else
                            errmsg = errmsg & "The record was not found." & vbCrLf
                            d_dr.Close()
                            results = "Failure"
                        End If
                        d_dr.Close()

                        ' If located the test, then go ahead and load it
                        If retval > 0 Then
                            GoTo RetrieveBinary
                        End If

                    Catch oBug As Exception
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                        results = "Failure"
                    End Try
                End If

                ' Retrieve normal document 
                SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.last_upd, v.minio_flg, v.row_id, d.row_id " &
                    "FROM DMS.dbo.Documents d " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                    "WHERE d.ext_id='" & CRSE_ID & "' and dc.cat_id=6 ORDER BY v.version DESC, d.last_upd DESC"
                Try
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get lesson.xml document: " & SqlS)
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                last_upt = d_dr.GetDateTime(2)
                                minio_flg = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                                d_verid = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(5), enumObjectType.StrType))
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_dsize=" & d_dsize & ", d_doc_id=" & d_doc_id & ", d_verid=" & d_verid & ", d_dsize=" & d_dsize & ", minio_flg=" & minio_flg)

                                If minio_flg = "Y" Then
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Minio")
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
                                            ReDim outbyte(Val(retval - 1))
                                            Dim intval As Integer
                                            For i = 0 To retval - 1
                                                intval = mobj2.ResponseStream.ReadByte()
                                                If intval < 255 And intval > 0 Then
                                                    outbyte(i) = intval
                                                End If
                                                If intval = 255 Then outbyte(i) = 255
                                                If intval < 0 Then
                                                    outbyte(i) = 0
                                                    If Debug = "Y" Then mydebuglog.Debug("  >  .. " & i.ToString & "   intval: " & intval.ToString)
                                                End If
                                            Next
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
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Document_Versions")
                                    ' Get binary and attach to the object outbyte
                                    If d_dsize <> "" Then
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                        last_upt = d_dr.GetDateTime(2)
                                    End If
                                End If

                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "The record was not found." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                    results = "Failure"
                End Try
                If Debug = "Y" Then
                    mydebuglog.Debug("   ... Reported d_dsize: " & d_dsize & " : Found: " & Str(retval))
                End If

                ' -----            
                ' Retrieve document and store to the temp file
RetrieveBinary:
                If retval > 0 Then
                    fileContents = Encoding.UTF8.GetString(outbyte, 0, outbyte.Length)
                End If
            Else
                results = "Failure"
                GoTo CloseOut
            End If

            ' Store to cache and text string
            Dim policy As New CacheItemPolicy()
            'policy.SlidingExpiration = TimeSpan.FromDays(CDbl(dms_cache_age))
            'fileContents = File.ReadAllText(BinaryFile)
            'filecache.Set("lesson-" & CRSE_ID, fileContents, policy)

            BinaryFile = "lesson-" & Trim(CRSE_ID)
            filecache.Set(BinaryFile, New HciDMSDocument(last_upt, fileContents), policy)
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Caching DMS doc to key: " & BinaryFile & vbCrLf)
            If Debug = "R" Then Debug = "T"
            'If (My.Computer.FileSystem.FileExists(BinaryFile)) And Debug <> "Y" Then Kill(BinaryFile)
        Else
            ' Mark as retrieved from cache
            retval = fileContents.Length
            BinaryFile = "lesson-" & CRSE_ID
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Retrieved from cache key: " & BinaryFile & vbCrLf)
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

        ' ============================================
        ' Load the temp file into an XML document
        'If retval > 0 And BinaryFile <> "" And Debug <> "T" Then
        If retval > 0 And Debug <> "T" Then
            ' Return the XML stored as a text string
            Try
                odoc.LoadXml(fileContents.Substring(fileContents.IndexOf(Environment.NewLine)))
            Catch ex As Exception
                errmsg = errmsg & "Unable to load XML temp file: " & ex.ToString & vbCrLf
                results = "Failure"
            End Try
        Else
            ' Return the status of the search/retrieval
            Dim resultsDeclare As System.Xml.XmlDeclaration
            Dim resultsRoot As System.Xml.XmlElement

            ' Create container with results
            resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
            odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

            ' Create root node
            resultsRoot = odoc.CreateElement("results")
            odoc.InsertAfter(resultsRoot, resultsDeclare)
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
            resultsDeclare = Nothing
            resultsRoot = Nothing
        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetLessonData : Error: " & Trim(errmsg))
        If cached Then
            If Debug <> "T" Then myeventlog.Info("GetLessonData : CACHED Results: " & results & " from lesson-" & CRSE_ID & " for RegId # " & RegId)
        Else
            If Debug <> "T" Then myeventlog.Info("GetLessonData : Results: " & results & ", docid: " & d_doc_id & ", minio: " & minio_flg & ", verid: " & d_verid & ", for RegId # " & RegId)
        End If
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If cached Then
                    mydebuglog.Debug("Results: CACHED " & results & " from lesson-" & CRSE_ID & " for RegId # " & RegId & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
                Else
                    mydebuglog.Debug("Results: " & results & ", docid: " & d_doc_id & ", minio: " & minio_flg & ", verid: " & d_verid & ", for RegId # " & RegId & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
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
        ' Return results
        Return odoc
    End Function

    <WebMethod(Description:="Retrieves questions and answers to in-course assessments")> _
    Public Function QuizLibrary(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function locates the XML quiz data for a course and returns it to the invoker 

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	        - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results, temp As String
        Dim mypath, errmsg, logging As String
        Dim DecodedUserId, ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Minio declarations
        Dim minio_flg, d_verid, d_doc_id As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' File handling declarations
        'Dim bfs As FileStream
        'Dim bw As BinaryWriter
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile As String
        Dim abyBuffer(1000) As Byte
        Dim d_dsize, SaveDest As String
        Dim CRSE_ID, TEST_FLG As String
        Dim filecache As ObjectCache = MemoryCache.Default
        Dim questionsCache As New CachingWrapper.LocalCache
        Dim fileContents, cacheName As String
        Dim cached As Boolean = False

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

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
        TEST_FLG = "N"
        d_dsize = ""
        BinaryFile = ""
        temp = ""
        fileContents = ""
        minio_flg = ""
        d_verid = ""
        d_doc_id = ""
        cacheName = ""

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("QuizLibrary_debug")
            If temp = "Y" And Debug <> "T" And Debug <> "R" Then Debug = "Y"
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
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If RegId = "" And Debug = "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\QuizLibrary.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Open database connections
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
        ' Get course information
        If Not cmd Is Nothing Then
            ' -----
            ' Query registration
            Try
                SqlS = "SELECT R.CRSE_ID, C.X_REGISTRATION_NUM, R.TEST_FLG " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            CRSE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            ValidatedUserId = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            TEST_FLG = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            If TEST_FLG = "" Then TEST_FLG = "N"
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
            If Debug = "Y" Then
                mydebuglog.Debug("   ... CRSE_ID: " & CRSE_ID)
                mydebuglog.Debug("   ... TEST_FLG: " & TEST_FLG)
            End If

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

        ' ============================================
        ' Create output directory
        SaveDest = mypath & "course_temp"
        Try
            Directory.CreateDirectory(SaveDest)
        Catch
        End Try

        ' ============================================
        Dim last_upt As Date = Today.AddYears(-50)
        ' Check to see if the document is in the in-memory cache
        If Not IsNothing(filecache("questions-" & CRSE_ID)) Then
            'Check if the cached item need to be renewed; v.last_upd; 2/27/17; Ren Hou;
            SqlS = "SELECT TOP 1 v.last_upd " & _
                "FROM DMS.dbo.Documents d " & _
                "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " & _
                "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " & _
                "WHERE d.ext_id='" & CRSE_ID & "' and dc.cat_id=7 ORDER BY v.version DESC, d.last_upd DESC"
            Try
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check last updated date for xml document: " & vbCrLf & SqlS)
                d_cmd.CommandText = SqlS
                last_upt = d_cmd.ExecuteScalar()
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
                results = "Failure"
            End Try

            If last_upt > TryCast(filecache("questions-" & CRSE_ID), HciDMSDocument).UpdateDate Then
                'Remove if the update_date on the cache is before the last updtaed date on DB record.
                filecache.Remove("questions-" & CRSE_ID)
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Cached object " & "questions-" & CRSE_ID & " expired.")
            End If
        End If
        '  *********** ''

        'Load content of cached object
        cacheName = "questions-" & CRSE_ID
        If filecache(cacheName) Is Nothing Then
            fileContents = ""
        Else
            cached = True
            fileContents = TryCast(filecache(cacheName), HciDMSDocument).CachedObj
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Retrieved filecache Cached object " & cacheName)
        End If

        ' Try from LocalCache.dll
        Try
            If fileContents = "" Then
                If Not questionsCache.GetCachedItem(cacheName) Is Nothing Then
                    fileContents = questionsCache.GetCachedItem(cacheName)
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Retrieved LocalCache Cached object " & cacheName)
                End If
            End If
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  LocalCache error: " & ex.Message)
        End Try

        If fileContents Is Nothing Or Debug = "R" Or fileContents = "" Then

            ' ============================================
            ' Get document containing XML and store in a local temp file      
            If Not d_cmd Is Nothing Then

                ' If in test mode, then look for a resource with the name "test-questions.xml"
                If TEST_FLG = "Y" Then
                    SqlS = "SELECT TOP 1 v.dsize, v.dimage " &
                        "FROM DMS.dbo.Documents d " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id " &
                        "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=10 " &
                        " and (d.dfilename='test-questions.xml' or d.name='test-questions.xml') " &
                        "AND d.deleted is null " &
                        "ORDER BY v.version DESC, d.last_upd DESC"
                    Try
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get test xml document: " & vbCrLf & SqlS)
                        d_cmd.CommandText = SqlS
                        d_dr = d_cmd.ExecuteReader()
                        If Not d_dr Is Nothing Then
                            While d_dr.Read()
                                Try
                                    d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                    If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_dsize=" & d_dsize)

                                    ' Get binary and attach to the object outbyte
                                    If d_dsize <> "" Then
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                    End If

                                Catch ex As Exception
                                    results = "Failure"
                                    errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                    GoTo CloseOut
                                End Try
                            End While
                        Else
                            errmsg = errmsg & "The record was not found." & vbCrLf
                            d_dr.Close()
                            results = "Failure"
                        End If
                        d_dr.Close()

                        ' If located the test, then go ahead and load it
                        If retval > 0 Then
                            GoTo RetrieveBinary
                        End If

                    Catch oBug As Exception
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                        results = "Failure"
                    End Try
                End If

                ' Locate standard document
                SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.last_upd, v.minio_flg, v.row_id, d.row_id " &
                    "FROM DMS.dbo.Documents d " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                    "WHERE d.ext_id='" & CRSE_ID & "' and dc.cat_id=7 ORDER BY v.version DESC, d.last_upd DESC"
                Try
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get xml document: " & vbCrLf & SqlS)
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                minio_flg = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                                d_verid = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(5), enumObjectType.StrType))
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_dsize=" & d_dsize & ",  d_doc_id=" & d_doc_id & ",  d_verid=" & d_verid & ",  minio_flg=" & minio_flg)

                                If minio_flg = "Y" Then
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Minio")
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
                                            ReDim outbyte(Val(retval - 1))
                                            Dim intval As Integer
                                            For i = 0 To retval - 1
                                                intval = mobj2.ResponseStream.ReadByte()
                                                If intval < 255 And intval > 0 Then
                                                    outbyte(i) = intval
                                                End If
                                                If intval = 255 Then outbyte(i) = 255
                                                If intval < 0 Then
                                                    outbyte(i) = 0
                                                    If Debug = "Y" Then mydebuglog.Debug("  >  .. " & i.ToString & "   intval: " & intval.ToString)
                                                End If
                                            Next
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
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Document_Versions")
                                    ' Get binary and attach to the object outbyte
                                    If d_dsize <> "" Then
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                        last_upt = d_dr.GetDateTime(2)
                                    End If
                                End If

                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "The record was not found." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                    results = "Failure"
                End Try
                If Debug = "Y" Then
                    mydebuglog.Debug("   ... Reported d_dsize: " & d_dsize & " : Found: " & Str(retval) & vbCrLf)
                End If

                ' -----            
                ' Retrieve document and store to the temp file
RetrieveBinary:
                If retval > 0 Then
                    fileContents = Encoding.UTF8.GetString(outbyte, 0, outbyte.Length)
                End If

            Else
                results = "Failure"
                GoTo CloseOut
            End If

            ' Store to localcache
            Try
                If fileContents.Length > 0 Then
                    questionsCache.AddToCache(cacheName, fileContents, CachingWrapper.CachePriority.NotRemovable)
                End If
            Catch ex As Exception
            End Try

            ' Store to filecache and text string
            Dim policy As New CacheItemPolicy()
            'policy.SlidingExpiration = TimeSpan.FromDays(CDbl(dms_cache_age))
            'fileContents = File.ReadAllText(BinaryFile)
            filecache.Set(cacheName, New HciDMSDocument(last_upt, fileContents), policy)
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Caching DMS doc to key: " & cacheName & vbCrLf)
            If Debug = "R" Then Debug = "T"
        Else
            ' Mark as retrieved from cache
            retval = fileContents.Length
            'BinaryFile = "questions-" & CRSE_ID
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Bytes retrieved from cache key: " & retval.ToString & vbCrLf)
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

        ' ============================================
        ' Return an XML document
        'If retval > 0 And BinaryFile <> "" And Debug <> "T" Then
        If retval > 0 And Debug <> "T" Then
            ' Return the XML stored as a text string
            Try
                odoc.LoadXml(fileContents.Substring(fileContents.IndexOf(Environment.NewLine)))
            Catch ex As Exception
                errmsg = errmsg & "Unable to load XML temp file: " & ex.ToString & vbCrLf
                results = "Failure"
            End Try
        Else
            ' Return the status of the search/retrieval
            Dim resultsDeclare As System.Xml.XmlDeclaration
            Dim resultsRoot As System.Xml.XmlElement

            ' Create container with results
            resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
            odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

            ' Create root node
            resultsRoot = odoc.CreateElement("results")
            odoc.InsertAfter(resultsRoot, resultsDeclare)
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
            resultsDeclare = Nothing
            resultsRoot = Nothing
        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("QuizLibrary : Error: " & Trim(errmsg))
        If cached Then
            If Debug <> "T" Then myeventlog.Info("QuizLibrary : CACHED Results: " & results & " for RegId # " & RegId)
        Else
            If Debug <> "T" Then myeventlog.Info("QuizLibrary : Results: " & results & ", minio: " & minio_flg & ", docid: " & d_doc_id & ", verid: " & d_verid & ", for RegId # " & RegId)
        End If
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If cached Then
                    mydebuglog.Debug("Results: CACHED " & results & " for RegId # " & RegId & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
                Else
                    mydebuglog.Debug("Results: " & results & ", minio: " & minio_flg & ", docid: " & d_doc_id & ", verid: " & d_verid & ", for RegId # " & RegId & ". Started " & LogStartTime & " .. Finished " & Now.ToString)
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
        ' Return results
        Return odoc
    End Function

    <WebMethod(Description:="Retrieves an XML representation of course scenario data")> _
    Public Function GetScenarioData(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function locates the XML scenario data file for a course and returns it to the invoker 

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	        - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results, temp As String
        Dim mypath, errmsg, logging As String
        Dim DecodedUserId, ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' DMS Database declarations
        Dim d_con As SqlConnection
        Dim d_cmd As SqlCommand
        Dim d_dr As SqlDataReader
        Dim d_ConnS As String

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Minio declarations
        Dim minio_flg, d_verid, d_doc_id As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' File handling declarations
        Dim bfs As FileStream
        Dim bw As BinaryWriter
        Dim outbyte(1000) As Byte
        Dim retval As Long
        Dim startIndex As Long = 0
        Dim BinaryFile, BFileName As String
        Dim abyBuffer(1000) As Byte
        Dim d_dsize, SaveDest As String
        Dim CRSE_ID, TEST_FLG As String
        Dim filecache As ObjectCache = MemoryCache.Default
        Dim fileContents As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

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
        TEST_FLG = ""
        d_dsize = ""
        BinaryFile = ""
        temp = ""
        fileContents = ""
        minio_flg = ""
        d_verid = ""
        d_doc_id = ""

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            d_ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            If d_ConnS = "" Then d_ConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetScenarioData_debug")
            If temp = "Y" And Debug <> "T" And Debug <> "R" Then Debug = "Y"
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
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If RegId = "" And Debug = "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-E95M1"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetScenarioData.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Open database connections
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
        ' Get course information
        If Not cmd Is Nothing Then
            ' -----
            ' Query registration
            Try
                SqlS = "SELECT R.CRSE_ID, C.X_REGISTRATION_NUM, R.TEST_FLG " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Get registration: " & vbCrLf & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            CRSE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            ValidatedUserId = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            TEST_FLG = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
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
            If Debug = "Y" Then
                mydebuglog.Debug("   ... CRSE_ID: " & CRSE_ID)
                mydebuglog.Debug("   ... TEST_FLG: " & TEST_FLG)
            End If

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

        ' ============================================
        ' Create output directory
        SaveDest = mypath & "course_temp\" & CRSE_ID
        Try
            Directory.CreateDirectory(SaveDest)
        Catch
        End Try

        ' ============================================
        ' Check to see if the document is in the in-memory cache
        fileContents = TryCast(filecache("scenarios-" & CRSE_ID), String)
        If fileContents Is Nothing Or Debug = "R" Or fileContents = "" Then

            ' ============================================
            ' Get document containing XML and store in a local temp file      
            '  Look for the document scenarios.xml stored as a course resource 
            If Not d_cmd Is Nothing Then

                ' If in test mode, then look for a resource with the name "test-scenario.xml"
                If TEST_FLG = "Y" Then
                    SqlS = "SELECT TOP 1 v.dsize, v.dimage " & _
                        "FROM DMS.dbo.Documents d " & _
                        "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " & _
                        "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " & _
                        "LEFT OUTER JOIN DMS.dbo.Document_Associations da ON da.doc_id=d.row_id " & _
                        "WHERE da.fkey='" & CRSE_ID & "' and da.association_id=15 and dc.cat_id=10 " & _
                        " and (d.dfilename='test-scenarios.xml' or d.name='test-scenarios.xml') " & _
                        "AND d.deleted is null " & _
                        "ORDER BY v.version DESC, d.last_upd DESC"
                    Try
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Get test xml document: " & vbCrLf & SqlS)
                        d_cmd.CommandText = SqlS
                        d_dr = d_cmd.ExecuteReader()
                        If Not d_dr Is Nothing Then
                            While d_dr.Read()
                                Try
                                    d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                    If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_dsize=" & d_dsize)

                                    ' Get binary and attach to the object outbyte
                                    If d_dsize <> "" Then
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                    End If

                                Catch ex As Exception
                                    results = "Failure"
                                    errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                    GoTo CloseOut
                                End Try
                            End While
                        Else
                            errmsg = errmsg & "The record was not found." & vbCrLf
                            d_dr.Close()
                            results = "Failure"
                        End If
                        d_dr.Close()

                        ' If located the test, then go ahead and load it
                        If retval > 0 Then
                            GoTo RetrieveBinary
                        End If

                    Catch oBug As Exception
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                        results = "Failure"
                    End Try
                End If

                ' Query DMS
                SqlS = "SELECT TOP 1 v.dsize, v.dimage, v.minio_flg, v.row_id, d.row_id  " &
                    "FROM DMS.dbo.Documents d " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Versions v ON v.row_id=d.last_version_id " &
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories dc ON dc.doc_id=d.row_id " &
                    "WHERE d.ext_id='" & CRSE_ID & "' and dc.cat_id=10 and d.dfilename='scenarios.xml' " &
                    "ORDER BY v.version DESC, d.last_upd DESC"
                Try
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Get xml document: " & vbCrLf & SqlS)
                    d_cmd.CommandText = SqlS
                    d_dr = d_cmd.ExecuteReader()
                    If Not d_dr Is Nothing Then
                        While d_dr.Read()
                            Try
                                d_dsize = Trim(CheckDBNull(d_dr(0), enumObjectType.StrType))
                                minio_flg = Trim(CheckDBNull(d_dr(2), enumObjectType.StrType))
                                d_verid = Trim(CheckDBNull(d_dr(3), enumObjectType.StrType))
                                d_doc_id = Trim(CheckDBNull(d_dr(4), enumObjectType.StrType))
                                If Debug = "Y" Then mydebuglog.Debug("  > Found record on query.  d_doc_id=" & d_doc_id & ",  d_verid=" & d_verid & ",  d_dsize=" & d_dsize & ",  minio_flg=" & minio_flg)

                                If minio_flg = "Y" Then
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Minio")
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
                                            ReDim outbyte(Val(retval - 1))
                                            Dim intval As Integer
                                            For i = 0 To retval - 1
                                                intval = mobj2.ResponseStream.ReadByte()
                                                If intval < 255 And intval > 0 Then
                                                    outbyte(i) = intval
                                                End If
                                                If intval = 255 Then outbyte(i) = 255
                                                If intval < 0 Then
                                                    'outbyte(i) = 0
                                                    'If Debug = "Y" Then mydebuglog.Debug("  >  .. " & i.ToString & "   intval: " & intval.ToString)
                                                End If
                                            Next
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
                                    If Debug = "Y" Then mydebuglog.Debug("  > Getting binary from Document_Versions")
                                    ' Get binary and attach to the object outbyte
                                    If d_dsize <> "" Then
                                        ReDim outbyte(Val(d_dsize) - 1)
                                        startIndex = 0
                                        retval = d_dr.GetBytes(1, 0, outbyte, 0, d_dsize)
                                    End If
                                End If

                            Catch ex As Exception
                                results = "Failure"
                                errmsg = errmsg & "Error checking record. " & ex.ToString & vbCrLf
                                GoTo CloseOut
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "The record was not found." & vbCrLf
                        d_dr.Close()
                        results = "Failure"
                    End If
                    d_dr.Close()
                Catch oBug As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                    results = "Failure"
                End Try
                If Debug = "Y" Then
                    mydebuglog.Debug("   ... Reported d_dsize: " & d_dsize & " : Found: " & Str(retval))
                End If

                ' -----            
                ' Retrieve document and store to the temp file
RetrieveBinary:
                If retval > 0 Then
                    ' Check for existance of the temporary directory - create if needed
                    SaveDest = mypath & "course_temp\" & CRSE_ID
                    Try
                        Directory.CreateDirectory(SaveDest)
                    Catch
                    End Try

                    ' Retrieve and store the document
                    BFileName = "scenarios.xml"
                    BinaryFile = SaveDest & "\" & BFileName
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Saving to " & BinaryFile)
                    If (My.Computer.FileSystem.FileExists(BinaryFile)) Then Kill(BinaryFile)
                    Try
                        bfs = New FileStream(BinaryFile, FileMode.Create, FileAccess.Write)
                        bw = New BinaryWriter(bfs)
                        bw.Write(outbyte)
                        bw.Flush()
                        bw.Close()
                        bfs.Close()
                    Catch ex As Exception
                        errmsg = errmsg & "Unable to write the xml document to a temp file." & ex.ToString & vbCrLf
                        results = "Failure"
                        retval = 0
                    End Try
                    bfs = Nothing
                    bw = Nothing
                End If
            Else
                results = "Failure"
                GoTo CloseOut
            End If


            ' Store to cache and text string
            Dim policy As New CacheItemPolicy()
            fileContents = File.ReadAllText(BinaryFile)
            filecache.Set("scenarios-" & CRSE_ID, fileContents, policy)
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Caching to key: scenarios-" & CRSE_ID & vbCrLf)
            If Debug = "R" Then Debug = "T"
            If (My.Computer.FileSystem.FileExists(BinaryFile)) And Debug <> "Y" Then Kill(BinaryFile)
        Else
            ' Mark as retrieved from cache
            retval = fileContents.Length
            BinaryFile = "scenarios-" & CRSE_ID
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Retrieved from cache key: " & BinaryFile & vbCrLf)
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

        ' ============================================
        ' Load the temp file into an XML document
        If retval > 0 And BinaryFile <> "" And Debug <> "T" Then
            Try
                odoc.LoadXml(fileContents)
            Catch ex As Exception
                errmsg = errmsg & "Unable to load XML temp file: " & ex.ToString & vbCrLf
                results = "Failure"
            End Try
        Else
            Dim resultsDeclare As System.Xml.XmlDeclaration
            Dim resultsRoot As System.Xml.XmlElement

            ' Create container with results
            resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
            odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

            ' Create root node
            resultsRoot = odoc.CreateElement("results")
            odoc.InsertAfter(resultsRoot, resultsDeclare)
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
            resultsDeclare = Nothing
            resultsRoot = Nothing
        End If
        If (My.Computer.FileSystem.FileExists(BinaryFile)) And Debug <> "Y" Then Kill(BinaryFile)

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetScenarioData : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GetScenarioData : Results: " & results & " for RegId # " & RegId & ", docid: " & d_doc_id & ",  verid: " & d_verid & " at " & Now.ToString)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " at " & Now.ToString)
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
        Return odoc
    End Function

    <WebMethod(Description:="Retrieves an XML representation on an assessment")> _
    Public Function GetExamXML(ByVal AssessmentId As String, ByVal Debug As String) As XmlDocument
        ' This function generates extracts the specified assessment (S_CRSE_TST records)
        ' into an XML document for the purpose of generating an online assessment

        ' The input parameters are as follows:
        '   AssessmentId    - The S_CRSE_TST.ROW_ID of the assessment
        '   Debug	        - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Assessment declarations
        Dim temp, database, Mode As String
        Dim RecCount As Integer
        Dim MAX_POINTS, TIME_ALLOWED, PASSING_SCORE, QUES_ID, POINTS, QUES_TEXT, QUES_TYPE_CD As String
        Dim QUES_SEQ_NUM, ANSR_ID, CORRECT_ANSR_FLG, ANSR_TEXT, ANSR_SEQ_NUM, ANSR_CD, ASSESS_NAME As String
        Dim Q_MULTILINE_FLG, Q_ROWS, Q_COLS, Q_SIZE_LIMIT As String
        Dim SCORE_PCT As Double
        Dim sQUES_ID As String  ' Used to store last answer id
        Dim InQues As Boolean
        Dim randomizeQuestionSequence As String
        Dim NUM_POINTS, NUM_QUESTIONS, QUES_CNT As Integer
        Dim SURVEY_FLG As String
        Dim QUES_BANK As Boolean

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        temp = ""
        database = ""
        SqlS = ""
        returnv = 0
        RecCount = 0
        sQUES_ID = ""
        InQues = False
        Q_MULTILINE_FLG = ""
        Q_ROWS = ""
        Q_COLS = ""
        Q_SIZE_LIMIT = ""
        NUM_POINTS = 0
        NUM_QUESTIONS = 0
        QUES_BANK = False
        SURVEY_FLG = "N"
        QUES_CNT = 0

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If AssessmentId = "" And Debug = "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            AssessmentId = "1-9BFM9"
        Else
            AssessmentId = Trim(AssessmentId)
        End If

        ' ============================================
        ' Fix spaces if found
        If InStr(AssessmentId, " ") > 0 Then
            temp = ""
            For i = 1 To Len(AssessmentId)
                If Mid(AssessmentId, i, 1) = " " Then
                    temp = temp & "+"
                Else
                    temp = temp & Mid(AssessmentId, i, 1)
                End If
            Next
            AssessmentId = temp
            temp = ""
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetExamXML.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  AssessmentId:" & AssessmentId)
            End If
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try
        If Debug = "Y" Then
            Try
                mydebuglog.Debug(vbCrLf & "Contact In-")
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Open output streams
        Dim memory_stream As New MemoryStream
        Dim assessXML As New XmlTextWriter(memory_stream, System.Text.Encoding.UTF8)

        ' ============================================
        ' Determine whether this is a question bank exam
        SqlS = "SELECT T.MAX_POINTS, COUNT(*) AS NUM_QUES, T.X_SURVEY_FLG " & _
        "FROM siebeldb.dbo.S_CRSE_TST T " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_QUES Q ON Q.CRSE_TST_ID=T.ROW_ID " & _
        "WHERE T.ROW_ID='" & AssessmentId & "' " & _
        "GROUP BY T.MAX_POINTS, T.X_SURVEY_FLG "
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check for question bank assessment: " & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        NUM_POINTS = CheckDBNull(dr(0), enumObjectType.IntType)
                        NUM_QUESTIONS = CheckDBNull(dr(1), enumObjectType.IntType)
                        SURVEY_FLG = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                        If NUM_POINTS <> NUM_QUESTIONS And SURVEY_FLG <> "Y" Then QUES_BANK = True
                        If Debug = "Y" Then mydebuglog.Debug("   > SURVEY_FLG: " & SURVEY_FLG & ", NUM_POINTS: " & NUM_POINTS.ToString & ",  NUM_QUESTIONS: " & NUM_QUESTIONS.ToString & ", QUES_BANK: " & QUES_BANK.ToString)
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading assessment: " & ex.ToString & vbCrLf
                    End Try
                End While
            End If
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Error locating assessment records. " & ex.ToString
            results = "Failure"
            GoTo CloseOut
        End Try
        dr.Close()

        If NUM_POINTS = 0 Then
            errmsg = errmsg & vbCrLf & "Error locating assessment records. NUM_POINTS = 0"
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        If QUES_BANK Then
            SqlS = "SELECT T.MAX_POINTS, T.X_TIME_ALLOWED, T.PASSING_SCORE, " & _
            "Q.ROW_ID, Q.POINTS, Q.QUES_TEXT, Q.QUES_TYPE_CD, Q.QUES_SEQ_NUM, " & _
            "A.ROW_ID, A.CORRECT_ANSR_FLG, A.ANSR_TEXT, A.ANSR_SEQ_NUM, A.X_ANSR_CD, " & _
            "T.X_SURVEY_FLG, T.NAME, CAST(A.MS_IDENT AS VARCHAR) AS NEW_ID, " & _
            "Q.X_MULTILINE_FLG, Q.X_ROWS, Q.X_COLS, Q.X_SIZE_LIMIT " & _
            "FROM siebeldb.dbo.S_CRSE_TST T " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_QUES Q ON Q.CRSE_TST_ID=T.ROW_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_ANSR A ON A.CRSE_TST_QUES_ID=Q.ROW_ID " & _
            "WHERE Q.ROW_ID IN (SELECT TOP " & NUM_POINTS.ToString.Trim & " ROW_ID  " & _
            "FROM siebeldb.dbo.S_CRSE_TST_QUES  " & _
            "WHERE CRSE_TST_ID='" & AssessmentId & "' " & _
            "order by newid()) " & _
            "ORDER BY Q.QUES_SEQ_NUM, A.ANSR_SEQ_NUM"
        Else
            SqlS = "SELECT T.MAX_POINTS, T.X_TIME_ALLOWED, T.PASSING_SCORE, " & _
            "Q.ROW_ID, Q.POINTS, Q.QUES_TEXT, Q.QUES_TYPE_CD, Q.QUES_SEQ_NUM, " & _
            "A.ROW_ID, A.CORRECT_ANSR_FLG, A.ANSR_TEXT, A.ANSR_SEQ_NUM, A.X_ANSR_CD, " & _
            "T.X_SURVEY_FLG, T.NAME, CAST(A.MS_IDENT AS VARCHAR) AS NEW_ID, " & _
            "Q.X_MULTILINE_FLG, Q.X_ROWS, Q.X_COLS, Q.X_SIZE_LIMIT " & _
            "FROM siebeldb.dbo.S_CRSE_TST T " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_QUES Q ON Q.CRSE_TST_ID=T.ROW_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_ANSR A ON A.CRSE_TST_QUES_ID=Q.ROW_ID " & _
            "WHERE T.ROW_ID='" & AssessmentId & "' " & _
            "ORDER BY Q.QUES_SEQ_NUM, A.ANSR_SEQ_NUM"
        End If
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get assessment: " & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        RecCount = RecCount + 1
                        ' ----
                        ' Get raw data
                        MAX_POINTS = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                        TIME_ALLOWED = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                        PASSING_SCORE = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                        QUES_ID = UCase(Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString)
                        POINTS = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                        QUES_TEXT = Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString
                        QUES_TYPE_CD = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                        QUES_SEQ_NUM = Trim(CheckDBNull(dr(7), enumObjectType.StrType)).ToString
                        ANSR_ID = UCase(Trim(CheckDBNull(dr(15), enumObjectType.StrType)).ToString)
                        If ANSR_ID = "" And QUES_TYPE_CD = "Text" Then ANSR_ID = QUES_ID
                        CORRECT_ANSR_FLG = Trim(CheckDBNull(dr(9), enumObjectType.StrType)).ToString
                        ANSR_TEXT = Trim(CheckDBNull(dr(10), enumObjectType.StrType)).ToString
                        If ANSR_TEXT = "" Then ANSR_TEXT = "Enter answer here"
                        ANSR_SEQ_NUM = Trim(CheckDBNull(dr(11), enumObjectType.StrType)).ToString
                        If ANSR_SEQ_NUM = "" Then ANSR_SEQ_NUM = "1"
                        ANSR_CD = UCase(Trim(CheckDBNull(dr(12), enumObjectType.StrType)).ToString)
                        If ANSR_CD = "" And IsNumeric(ANSR_SEQ_NUM) Then
                            ANSR_CD = Chr(64 + Val(ANSR_SEQ_NUM))
                        End If
                        Q_MULTILINE_FLG = UCase(Trim(CheckDBNull(dr(16), enumObjectType.StrType)).ToString)
                        Q_ROWS = UCase(Trim(CheckDBNull(dr(17), enumObjectType.StrType)).ToString)
                        Q_COLS = UCase(Trim(CheckDBNull(dr(18), enumObjectType.StrType)).ToString)
                        Q_SIZE_LIMIT = UCase(Trim(CheckDBNull(dr(19), enumObjectType.StrType)).ToString)

                        ' ----
                        ' Translate raw data for output
                        If Trim(CheckDBNull(dr(13), enumObjectType.StrType)) = "Y" Then
                            Mode = "survey"
                            randomizeQuestionSequence = "false"
                        Else
                            Mode = "test"
                            randomizeQuestionSequence = "true"
                        End If
                        ASSESS_NAME = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                        If Val(MAX_POINTS) > 0 And Val(PASSING_SCORE) > 0 Then
                            SCORE_PCT = (Val(PASSING_SCORE) / Val(MAX_POINTS)) * 100
                        End If
                        Select Case QUES_TYPE_CD
                            Case "Multiple Choice"
                                QUES_TYPE_CD = "multiple choice multiple answer"
                            Case "Text Match"
                                QUES_TYPE_CD = "matching"
                            Case "Screen"
                                QUES_TYPE_CD = "slide"
                            Case "Text"
                                QUES_TYPE_CD = "fill in the blank"
                            Case "Single Choice"
                                QUES_TYPE_CD = "multiple choice single answer"
                        End Select

                        ' ----
                        ' If first record, write document header
                        If RecCount = 1 Then
                            assessXML.Formatting = Formatting.Indented
                            assessXML.Indentation = 2
                            assessXML.WriteStartDocument(False)

                            '<assessment feedbackTime="end state" feedbackLevel="nothing" randomizeAnswerSequence="false" mode="test"
                            '	randomizeQuestionSequence="true" recordInteractions="true" displayTitle="true" timeLimit="6000"
                            '	showAllQuestionsOnSinglePage="false" allowRetakeWhenFailed="true"  maxNumberOfAttempts="2" showBackButton="true"
                            '	showExplanations="incorrect only" explanationLevel="all" xmlns="http://www.rusticisoftware.com/assessment.xsd">

                            assessXML.WriteStartElement("assessment")
                            assessXML.WriteAttributeString("feedbackTime", "never")
                            assessXML.WriteAttributeString("feedbackLevel", "nothing")
                            assessXML.WriteAttributeString("randomizeAnswerSequence", "false")  ' Need to make this table driven
                            assessXML.WriteAttributeString("mode", Mode)
                            assessXML.WriteAttributeString("randomizeQuestionSequence", randomizeQuestionSequence)  ' Need to make this table driven
                            assessXML.WriteAttributeString("recordInteractions", "true")
                            assessXML.WriteAttributeString("displayTitle", "false")
                            assessXML.WriteAttributeString("timeLimit", TIME_ALLOWED)
                            assessXML.WriteAttributeString("showAllQuestionsOnSinglePage", "false")
                            assessXML.WriteAttributeString("allowRetakeWhenFailed", "true")     ' Need to make this table driven
                            assessXML.WriteAttributeString("maxNumberOfAttempts", "1000")       ' Need to make this table driven
                            assessXML.WriteAttributeString("showBackButton", "true")
                            assessXML.WriteAttributeString("showExplanations", "incorrect only")
                            assessXML.WriteAttributeString("explanationLevel", "all")
                            assessXML.WriteAttributeString("xmlns", "http://hciscorm.certegrity.com/assessment.xsd")

                            '	<title>Demonstration Assessment</title>
                            assessXML.WriteStartElement("title")
                            assessXML.WriteString(ASSESS_NAME)
                            assessXML.WriteEndElement()

                            '	<stylesheetUrl>sample_assessment.css</stylesheetUrl>
                            assessXML.WriteStartElement("stylesheetUrl")
                            assessXML.WriteString("assessment.css")
                            assessXML.WriteEndElement()

                            '	<finalReportTemplate>
                            '		<![CDATA[
                            '			<h3>Congratulations $STUDENT_NAME, you have finished the demonstration quiz!</h3>
                            '			<div>You have <i>$SUCCESS_STATUS</i> this quiz with a score of <strong>$SCORE%</strong>.</div>
                            '			<div style='color: blue'>Please exit the quiz by clicking the "exit" button above.</div>
                            '			<br /><br />
                            '		]]>
                            '	</finalReportTemplate>
                            assessXML.WriteStartElement("finalReportTemplate")
                            If LCase(Mode) = "survey" Then
                                temp = "<h2>Thank you $STUDENT_NAME for completing this survey</h2>" & _
                                    "<div id='ExitMsg'>To proceed click the ""Exit"" button.</div>" & _
                                    "<br /><br />"
                            Else
                                temp = "<h2>Congratulations $STUDENT_NAME, you have finished this " & Mode & "</h2>" & _
                                    "<div id='ExitMsg'>To proceed click the ""Exit"" button.</div>" & _
                                    "<br /><br />"
                            End If
                            assessXML.WriteCData(temp)
                            assessXML.WriteEndElement()
                            '	<!--numberOfQuestionsToRandomlySelect>2</numberOfQuestionsToRandomlySelect-->

                            '	<passingScore>65.0</passingScore>
                            assessXML.WriteStartElement("passingScore")
                            assessXML.WriteString(Format(SCORE_PCT))
                            assessXML.WriteEndElement()
                        End If

                        ' ----
                        ' Question
                        '<question id="1-9AUSC" type="multiple choice single answer">
                        '	<text>What is the address of Rustici Software's home on the web?</text>
                        '</question>
                        If QUES_ID <> sQUES_ID Then
                            QUES_CNT = QUES_CNT + 1
                            Try
                                If InQues = True Then assessXML.WriteEndElement() ' Close out prior question if necessary
                                assessXML.WriteStartElement("question")
                                assessXML.WriteAttributeString("id", QUES_ID)
                                assessXML.WriteAttributeString("type", QUES_TYPE_CD)
                                If QUES_TYPE_CD = "fill in the blank" Then
                                    If Q_MULTILINE_FLG = "Y" Then
                                        If Q_ROWS = "" Or Not IsNumeric(Q_ROWS) Then Q_ROWS = "5"
                                        If Q_COLS = "" Or Not IsNumeric(Q_COLS) Then Q_COLS = "60"
                                        assessXML.WriteAttributeString("textBoxSize", "60")
                                        assessXML.WriteAttributeString("multiline", "true")
                                        assessXML.WriteAttributeString("rows", Q_ROWS)
                                        assessXML.WriteAttributeString("cols", Q_COLS)
                                    Else
                                        If Q_SIZE_LIMIT = "" Or Not IsNumeric(Q_SIZE_LIMIT) Then Q_SIZE_LIMIT = "60"
                                        assessXML.WriteAttributeString("textBoxSize", Q_SIZE_LIMIT)
                                        assessXML.WriteAttributeString("multiline", "false")
                                        assessXML.WriteAttributeString("rows", "")
                                        assessXML.WriteAttributeString("cols", "")
                                    End If
                                End If
                                assessXML.WriteStartElement("text")
                                assessXML.WriteString(QUES_TEXT)
                                assessXML.WriteEndElement()
                            Catch ex As Exception
                                errmsg = errmsg & vbCrLf & "Error writing question. " & ex.ToString
                                results = "Failure"
                            End Try
                        End If
                        If QUES_ID <> "" Then InQues = True

                        ' ----
                        ' Answer
                        If Debug = "Y" Then mydebuglog.Debug("QUES # " & QUES_CNT.ToString & ",  ANSWER: " & RecCount.ToString & "  ... " & QUES_SEQ_NUM & "/" & ANSR_SEQ_NUM & "   ID: " & ANSR_ID & " - " & ANSR_TEXT)
                        Try
                            If ANSR_ID <> "" Then
                                Select Case QUES_TYPE_CD
                                    Case "slide"
                                        ' Do nothing
                                    Case "matching"
                                        ' ???
                                    Case "fill in the blank"
                                        '<answer letter="a"  id="q4a1" isCorrect="true" fillInTheBlankEvaluationMethod="regular expression">
                                        '	<text>www.adlnet.org</text>
                                        '	<regex>(adlnet.org)$|(ADLNET.ORG)$</regex>
                                        '</answer>
                                        assessXML.WriteStartElement("answer")
                                        assessXML.WriteAttributeString("letter", ANSR_CD)
                                        assessXML.WriteAttributeString("id", ANSR_ID)
                                        If CORRECT_ANSR_FLG = "Y" Then
                                            assessXML.WriteAttributeString("isCorrect", "true")
                                        Else
                                            assessXML.WriteAttributeString("isCorrect", "false")
                                        End If
                                        assessXML.WriteAttributeString("fillInTheBlankEvaluationMethod", "regular expression")
                                        assessXML.WriteStartElement("text")
                                        assessXML.WriteString(ANSR_TEXT)
                                        assessXML.WriteEndElement()
                                        assessXML.WriteStartElement("regex")
                                        assessXML.WriteString("(" & LCase(ANSR_TEXT) & ")$|(" & UCase(ANSR_TEXT) & ")$")
                                        assessXML.WriteEndElement()
                                        assessXML.WriteEndElement()
                                    Case Else
                                        '	<answer letter="a" id="q1a1" isCorrect="true">
                                        '		<text>www.scorm.com</text>
                                        '		<explanation>www.scorm.com is Rustici Software's web address.</explanation>
                                        '	</answer>
                                        assessXML.WriteStartElement("answer")
                                        assessXML.WriteAttributeString("letter", ANSR_CD)
                                        assessXML.WriteAttributeString("id", ANSR_ID)
                                        assessXML.WriteAttributeString("isCorrect", CORRECT_ANSR_FLG)
                                        assessXML.WriteStartElement("text")
                                        assessXML.WriteString(ANSR_TEXT)
                                        assessXML.WriteEndElement()
                                        assessXML.WriteStartElement("explanation")
                                        assessXML.WriteString("")
                                        assessXML.WriteEndElement()
                                        assessXML.WriteEndElement()
                                End Select
                            Else
                                ' No answer was found
                                If QUES_TYPE_CD = "fill in the blank" Then
                                    assessXML.WriteStartElement("answer")
                                    assessXML.WriteAttributeString("letter", ANSR_CD)
                                    assessXML.WriteAttributeString("id", ANSR_ID)
                                    assessXML.WriteAttributeString("isCorrect", CORRECT_ANSR_FLG)
                                    assessXML.WriteAttributeString("fillInTheBlankEvaluationMethod", "regular expression")
                                    assessXML.WriteStartElement("text")
                                    assessXML.WriteString(ANSR_TEXT)
                                    assessXML.WriteEndElement()
                                    assessXML.WriteStartElement("regex")
                                    assessXML.WriteString("(" & LCase(ANSR_TEXT) & ")$|(" & UCase(ANSR_TEXT) & ")$")
                                    assessXML.WriteEndElement()
                                    assessXML.WriteEndElement()
                                End If
                            End If
                        Catch ex As Exception
                            errmsg = errmsg & vbCrLf & "Error writing answer. " & ex.ToString
                            results = "Failure"
                        End Try
                        sQUES_ID = QUES_ID

                    Catch ex As Exception
                        errmsg = errmsg & "Error reading assessment: " & ex.ToString & vbCrLf
                    End Try
                End While

                ' Close document
                If RecCount > 0 Then
                    If InQues Then
                        assessXML.WriteEndElement()
                    End If
                    assessXML.WriteEndElement()
                    assessXML.Flush()
                End If
            End If
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Error locating assessment records. " & ex.ToString
            results = "Failure"
            GoTo CloseOut
        End Try
        dr.Close()

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
        ' Turn results into an XML document
        Dim stream_reader As New StreamReader(memory_stream)
        Try
            memory_stream.Seek(0, SeekOrigin.Begin)
            If memory_stream.Length > 0 And Debug = "Y" Then
                mydebuglog.Debug("Results: " & vbCrLf & stream_reader.ReadToEnd())
                mydebuglog.Debug("Size of results: " & memory_stream.Length.ToString)
            End If
            If memory_stream.Length > 0 And Debug <> "T" Then
                'odoc = New XmlDocument
                'odoc.LoadXml(stream_reader.ReadToEnd())
                If File.Exists("c:\temp\temp.xml") Then File.Delete("c:\temp\temp.xml")
                Dim tempxml As FileStream
                tempxml = New FileStream("c:\temp\temp.xml", FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                memory_stream.WriteTo(tempxml)
                tempxml.Close()
                tempxml = Nothing
                odoc = New XmlDocument
                odoc.Load("c:\temp\temp.xml")
                If Debug <> "Y" And File.Exists("c:\temp\temp.xml") Then File.Delete("c:\temp\temp.xml")
            Else
                Dim resultsDeclare As System.Xml.XmlDeclaration
                Dim resultsRoot As System.Xml.XmlElement
                resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
                odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
                resultsRoot = odoc.CreateElement("assessment")
                odoc.InsertAfter(resultsRoot, resultsDeclare)
                If Debug = "T" Then
                    If memory_stream.Length > 0 Then
                        AddXMLChild(odoc, resultsRoot, "results", "Success")
                    Else
                        AddXMLChild(odoc, resultsRoot, "results", "Failure")
                    End If
                Else
                    AddXMLChild(odoc, resultsRoot, "results", "Failure")
                    AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
                End If
                resultsDeclare = Nothing
                resultsRoot = Nothing
            End If

            ' Close some memory objects
            stream_reader.Close()
            memory_stream.Close()
            assessXML.Close()
            stream_reader = Nothing
            memory_stream = Nothing
            assessXML = Nothing
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug("Error converting results " & ex.ToString)
        End Try

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetExamXML : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GetExamXML : Results: " & results & " for AssessmentId: " & AssessmentId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug(vbCrLf & "  Results: " & results & " for AssessmentId: " & AssessmentId & " at " & Now.ToString)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
                mydebuglog.Debug("Error outputing results " & ex.ToString)
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

        ' Close other objects
        Try
            iDoc = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc

        ' Close memory objects
        odoc = Nothing

    End Function

    <WebMethod(Description:="Retrieves an XML representation on an HTML5 assessment")> _
    Public Function GetExamHTML5(ByVal TstRunId As String, ByVal Debug As String) As XmlDocument
        ' This function generates extracts the specified assessment (S_CRSE_TST records)
        ' into an XML document for the purpose of generating an online assessment

        ' The input parameters are as follows:
        '   TstRunId    	- The S_CRSE_TSTRUN.ROW_ID of the assessment taker
        '   Debug	        - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Assessment declarations
        Dim temp, database, Mode, AssessmentId As String
        Dim RecCount As Integer
        Dim MAX_POINTS, TIME_ALLOWED, PASSING_SCORE, QUES_ID, POINTS, QUES_TEXT, QUES_TYPE_CD As String
        Dim QUES_SEQ_NUM, ANSR_ID, CORRECT_ANSR_FLG, ANSR_TEXT, ANSR_SEQ_NUM, ANSR_CD, ASSESS_NAME As String
        Dim Q_MULTILINE_FLG, Q_ROWS, Q_COLS, Q_SIZE_LIMIT As String
        Dim SCORE_PCT As Double
        Dim sQUES_ID As String  ' Used to store last answer id
        Dim InQues As Boolean
        Dim randomizeQuestionSequence As String
        Dim XFORMAT, XFEED_TIME, XFEED_LVL, XRND_ANSR, XRND_QUES, XRECORD, XDISP_TITLE, XDISP_SINGLE, XALLOW_RETAKE As String
        Dim XSHOW_BACK, XSHOW_EXPL, XEXPL_LVL, XKBA_ATMPTS, XINACT_TIME, XQUEST_TIME, XSCORE, X_TRIES_ALLOWED As String
        Dim USER_NAME, T_LANGCD, KBA_Q_NUM, KBA_FLG, FST_NAME As String
        Dim NUM_POINTS, NUM_QUESTIONS As Integer
        Dim SURVEY_FLG As String
        Dim QUES_BANK As Boolean

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        temp = ""
        database = ""
        SqlS = ""
        returnv = 0
        RecCount = 0
        sQUES_ID = ""
        InQues = False
        Q_MULTILINE_FLG = ""
        Q_ROWS = ""
        Q_COLS = ""
        Q_SIZE_LIMIT = ""
        USER_NAME = ""
        XFORMAT = ""
        XFEED_TIME = ""
        XFEED_LVL = ""
        XRND_ANSR = ""
        XRND_QUES = ""
        XRECORD = ""
        XDISP_TITLE = ""
        XDISP_SINGLE = ""
        XALLOW_RETAKE = ""
        XSHOW_BACK = ""
        XSHOW_EXPL = ""
        XEXPL_LVL = ""
        XKBA_ATMPTS = ""
        XINACT_TIME = ""
        XQUEST_TIME = ""
        XSCORE = ""
        X_TRIES_ALLOWED = ""
        FST_NAME = ""
        NUM_POINTS = 0
        NUM_QUESTIONS = 0
        QUES_BANK = False
        SURVEY_FLG = "N"
        AssessmentId = ""

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetExamHTML5_debug")
            If temp = "Y" And Debug <> "T" And Debug <> "R" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If TstRunId = "" And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            TstRunId = "PT+732632"
        Else
            TstRunId = Trim(HttpUtility.UrlEncode(TstRunId))
            If InStr(TstRunId, "%") > 0 Then TstRunId = Trim(HttpUtility.UrlDecode(TstRunId))
            If InStr(TstRunId, " ") > 0 Then TstRunId = EncodeParamSpaces(TstRunId)
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetExamHTML5.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  TstRunId:" & TstRunId)
            End If
        End If

        ' Validate parameters
        If TstRunId = "undefined" Or TstRunId = "" Then
            errmsg = errmsg & vbCrLf & "TstRunId undefined. "
            results = "Failure"
            GoTo CloseOut2
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Open output streams
        Dim memory_stream As New MemoryStream
        Dim assessXML As New XmlTextWriter(memory_stream, System.Text.Encoding.UTF8)

        ' ============================================
        ' Determine whether this is a question bank exam
        SqlS = "SELECT T.MAX_POINTS, COUNT(*) AS NUM_QUES, T.X_SURVEY_FLG, T.ROW_ID " & _
        "FROM siebeldb.dbo.S_CRSE_TSTRUN R " & _
        "INNER JOIN siebeldb.dbo.S_CRSE_TST T ON T.ROW_ID=R.CRSE_TST_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_QUES Q ON Q.CRSE_TST_ID=T.ROW_ID " & _
        "WHERE R.ROW_ID='" & TstRunId & "' " & _
        "GROUP BY T.MAX_POINTS, T.X_SURVEY_FLG, T.ROW_ID "
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check for question bank assessment: " & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        NUM_POINTS = CheckDBNull(dr(0), enumObjectType.IntType)
                        NUM_QUESTIONS = CheckDBNull(dr(1), enumObjectType.IntType)
                        SURVEY_FLG = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                        AssessmentId = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                        If NUM_POINTS <> NUM_QUESTIONS And SURVEY_FLG <> "Y" Then QUES_BANK = True
                        If Debug = "Y" Then mydebuglog.Debug("   > AssessmentId: " & AssessmentId & ", SURVEY_FLG: " & SURVEY_FLG & ", NUM_POINTS: " & NUM_POINTS.ToString & ",  NUM_QUESTIONS: " & NUM_QUESTIONS.ToString & ", QUES_BANK: " & QUES_BANK.ToString)
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading assessment: " & ex.ToString & vbCrLf
                    End Try
                End While
            End If
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Error locating assessment records. " & ex.ToString
            results = "Failure"
            GoTo CloseOut
        End Try
        dr.Close()

        If NUM_POINTS = 0 Or AssessmentId = "" Then
            errmsg = errmsg & vbCrLf & "Error locating assessment records. NUM_POINTS = 0. "
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        If QUES_BANK Then
            SqlS = "SELECT T.MAX_POINTS, T.X_TIME_ALLOWED, T.PASSING_SCORE, " & _
            "Q.ROW_ID, Q.POINTS, Q.QUES_TEXT, Q.QUES_TYPE_CD, Q.QUES_SEQ_NUM, " & _
            "A.ROW_ID, A.CORRECT_ANSR_FLG, A.ANSR_TEXT, A.ANSR_SEQ_NUM, A.X_ANSR_CD, " & _
            "T.X_SURVEY_FLG, T.NAME, CAST(A.MS_IDENT AS VARCHAR) AS NEW_ID, " & _
            "Q.X_MULTILINE_FLG, Q.X_ROWS, Q.X_COLS, Q.X_SIZE_LIMIT, T.X_LANG_ID, T.X_KBA_QUES_NUM, " & _
            "T.X_FORMAT, T.X_FEED_TIME, T.X_FEED_LVL, T.X_RND_ANSR, T.X_RND_QUES, T.X_RECORD, T.X_DISP_TITLE, T.X_DISP_SINGLE, T.X_ALLOW_RETAKE, " & _
            "T.X_SHOW_BACK, T.X_SHOW_EXPL, T.X_EXPL_LVL,  T.X_KBA_ATMPTS,  T.X_INACT_TIME,  T.X_QUEST_TIME, T.X_SCORE, " & _
            "T.X_TRIES_ALLOWED, T.ROW_ID, C.FST_NAME + ' ' + C.LAST_NAME AS USER_NAME, C.FST_NAME " & _
            "FROM siebeldb.dbo.S_CRSE_TSTRUN R " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.PERSON_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON T.ROW_ID=R.CRSE_TST_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_QUES Q ON Q.CRSE_TST_ID=T.ROW_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_ANSR A ON A.CRSE_TST_QUES_ID=Q.ROW_ID " & _
            "WHERE R.ROW_ID='" & TstRunId & "' AND Q.ROW_ID IN (SELECT TOP " & NUM_POINTS.ToString.Trim & " ROW_ID " & _
            "FROM siebeldb.dbo.S_CRSE_TST_QUES WHERE CRSE_TST_ID='" & AssessmentId & "' ORDER BY newid()) " & _
            "ORDER BY Q.QUES_SEQ_NUM, A.ANSR_SEQ_NUM"
        Else
            SqlS = "SELECT T.MAX_POINTS, T.X_TIME_ALLOWED, T.PASSING_SCORE, " & _
            "Q.ROW_ID, Q.POINTS, Q.QUES_TEXT, Q.QUES_TYPE_CD, Q.QUES_SEQ_NUM, " & _
            "A.ROW_ID, A.CORRECT_ANSR_FLG, A.ANSR_TEXT, A.ANSR_SEQ_NUM, A.X_ANSR_CD, " & _
            "T.X_SURVEY_FLG, T.NAME, CAST(A.MS_IDENT AS VARCHAR) AS NEW_ID, " & _
            "Q.X_MULTILINE_FLG, Q.X_ROWS, Q.X_COLS, Q.X_SIZE_LIMIT, T.X_LANG_ID, T.X_KBA_QUES_NUM, " & _
            "T.X_FORMAT, T.X_FEED_TIME, T.X_FEED_LVL, T.X_RND_ANSR, T.X_RND_QUES, T.X_RECORD, T.X_DISP_TITLE, T.X_DISP_SINGLE, T.X_ALLOW_RETAKE, " & _
            "T.X_SHOW_BACK, T.X_SHOW_EXPL, T.X_EXPL_LVL,  T.X_KBA_ATMPTS,  T.X_INACT_TIME,  T.X_QUEST_TIME, T.X_SCORE, " & _
            "T.X_TRIES_ALLOWED, T.ROW_ID, C.FST_NAME + ' ' + C.LAST_NAME AS USER_NAME, C.FST_NAME " & _
            "FROM siebeldb.dbo.S_CRSE_TSTRUN R " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.PERSON_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON T.ROW_ID=R.CRSE_TST_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_QUES Q ON Q.CRSE_TST_ID=T.ROW_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST_ANSR A ON A.CRSE_TST_QUES_ID=Q.ROW_ID " & _
            "WHERE R.ROW_ID='" & TstRunId & "' " & _
            "ORDER BY Q.QUES_SEQ_NUM, A.ANSR_SEQ_NUM"
        End If
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get assessment attempt: " & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        RecCount = RecCount + 1
                        ' ----
                        ' Get raw data
                        MAX_POINTS = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                        TIME_ALLOWED = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                        PASSING_SCORE = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                        QUES_ID = UCase(Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString)
                        POINTS = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                        QUES_TEXT = Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString
                        QUES_TYPE_CD = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                        QUES_SEQ_NUM = Trim(CheckDBNull(dr(7), enumObjectType.StrType)).ToString
                        'ANSR_ID = UCase(Trim(CheckDBNull(dr(15), enumObjectType.StrType)).ToString)
                        ANSR_ID = UCase(Trim(CheckDBNull(dr(8), enumObjectType.StrType)).ToString)
                        If ANSR_ID = "" And QUES_TYPE_CD = "Text" Then ANSR_ID = QUES_ID
                        CORRECT_ANSR_FLG = Trim(CheckDBNull(dr(9), enumObjectType.StrType)).ToString
                        ANSR_TEXT = Trim(CheckDBNull(dr(10), enumObjectType.StrType)).ToString
                        If ANSR_TEXT = "" Then ANSR_TEXT = "Enter answer here"
                        ANSR_SEQ_NUM = Trim(CheckDBNull(dr(11), enumObjectType.StrType)).ToString
                        If ANSR_SEQ_NUM = "" Then ANSR_SEQ_NUM = "1"
                        ANSR_CD = UCase(Trim(CheckDBNull(dr(12), enumObjectType.StrType)).ToString)
                        If ANSR_CD = "" And IsNumeric(ANSR_SEQ_NUM) Then
                            ANSR_CD = Chr(64 + Val(ANSR_SEQ_NUM))
                        End If
                        Q_MULTILINE_FLG = UCase(Trim(CheckDBNull(dr(16), enumObjectType.StrType)).ToString)
                        Q_ROWS = UCase(Trim(CheckDBNull(dr(17), enumObjectType.StrType)).ToString)
                        Q_COLS = UCase(Trim(CheckDBNull(dr(18), enumObjectType.StrType)).ToString)
                        Q_SIZE_LIMIT = UCase(Trim(CheckDBNull(dr(19), enumObjectType.StrType)).ToString)
                        T_LANGCD = UCase(Trim(CheckDBNull(dr(20), enumObjectType.StrType)).ToString)
                        KBA_Q_NUM = CheckDBNull(dr(21), enumObjectType.IntType)
                        KBA_FLG = IIf(KBA_Q_NUM > 0, "true", "false")
                        XFORMAT = UCase(Trim(CheckDBNull(dr(22), enumObjectType.StrType)).ToString)
                        XFEED_TIME = UCase(Trim(CheckDBNull(dr(23), enumObjectType.StrType)).ToString)
                        XFEED_LVL = UCase(Trim(CheckDBNull(dr(24), enumObjectType.StrType)).ToString)
                        XRND_ANSR = UCase(Trim(CheckDBNull(dr(25), enumObjectType.StrType)).ToString)
                        XRND_QUES = UCase(Trim(CheckDBNull(dr(26), enumObjectType.StrType)).ToString)
                        XRECORD = UCase(Trim(CheckDBNull(dr(27), enumObjectType.StrType)).ToString)
                        XDISP_TITLE = UCase(Trim(CheckDBNull(dr(28), enumObjectType.StrType)).ToString)
                        XDISP_SINGLE = UCase(Trim(CheckDBNull(dr(29), enumObjectType.StrType)).ToString)
                        XALLOW_RETAKE = UCase(Trim(CheckDBNull(dr(30), enumObjectType.StrType)).ToString)
                        XSHOW_BACK = UCase(Trim(CheckDBNull(dr(31), enumObjectType.StrType)).ToString)
                        XSHOW_EXPL = UCase(Trim(CheckDBNull(dr(32), enumObjectType.StrType)).ToString)
                        XEXPL_LVL = UCase(Trim(CheckDBNull(dr(33), enumObjectType.StrType)).ToString)
                        XKBA_ATMPTS = UCase(Trim(CheckDBNull(dr(34), enumObjectType.StrType)).ToString)
                        XINACT_TIME = UCase(Trim(CheckDBNull(dr(35), enumObjectType.StrType)).ToString)
                        XQUEST_TIME = UCase(Trim(CheckDBNull(dr(36), enumObjectType.StrType)).ToString)
                        XSCORE = UCase(Trim(CheckDBNull(dr(37), enumObjectType.StrType)).ToString)
                        X_TRIES_ALLOWED = UCase(Trim(CheckDBNull(dr(38), enumObjectType.StrType)).ToString)
                        AssessmentId = UCase(Trim(CheckDBNull(dr(39), enumObjectType.StrType)).ToString)
                        USER_NAME = Trim(CheckDBNull(dr(40), enumObjectType.StrType)).ToString
                        FST_NAME = Trim(CheckDBNull(dr(41), enumObjectType.StrType)).ToString

                        ' ----
                        ' Translate raw data for output
                        If Trim(CheckDBNull(dr(13), enumObjectType.StrType)) = "Y" Then
                            Mode = "survey"
                            randomizeQuestionSequence = "false"
                        Else
                            Mode = "test"
                            randomizeQuestionSequence = "true"
                        End If
                        ASSESS_NAME = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                        If Val(MAX_POINTS) > 0 And Val(PASSING_SCORE) > 0 Then
                            SCORE_PCT = (Val(PASSING_SCORE) / Val(MAX_POINTS)) * 100
                        End If
                        Select Case QUES_TYPE_CD
                            Case "Multiple Choice"
                                QUES_TYPE_CD = "multiple choice multiple answer"
                            Case "Text Match"
                                QUES_TYPE_CD = "matching"
                            Case "Screen"
                                QUES_TYPE_CD = "slide"
                            Case "Text"
                                QUES_TYPE_CD = "fill in the blank"
                            Case "Single Choice"
                                QUES_TYPE_CD = "multiple choice single answer"
                        End Select

                        ' ----
                        ' If first record, write document header
                        If RecCount = 1 Then
                            assessXML.Formatting = Formatting.Indented
                            assessXML.Indentation = 2
                            assessXML.WriteStartDocument(False)

                            '<assessment feedbackTime="end state" feedbackLevel="nothing" randomizeAnswerSequence="false" mode="test"
                            '	randomizeQuestionSequence="true" recordInteractions="true" displayTitle="true" timeLimit="6000"
                            '	showAllQuestionsOnSinglePage="false" allowRetakeWhenFailed="true"  maxNumberOfAttempts="2" showBackButton="true"
                            '	showExplanations="incorrect only" explanationLevel="all" xmlns="http://www.rusticisoftware.com/assessment.xsd">

                            assessXML.WriteStartElement("assessment")
                            assessXML.WriteAttributeString("xmlns", "http://hciscorm.certegrity.com/assessment.xsd")
                            assessXML.WriteAttributeString("feedbackTime", XFEED_TIME)
                            assessXML.WriteAttributeString("feedbackLevel", XFEED_LVL)
                            assessXML.WriteAttributeString("randomizeAnswerSequence", IIf(XRND_ANSR = "Y", "true", "false"))
                            assessXML.WriteAttributeString("mode", Mode)
                            assessXML.WriteAttributeString("randomizeQuestionSequence", IIf(XRND_QUES = "Y", "true", "false"))
                            assessXML.WriteAttributeString("recordInteractions", IIf(XRECORD = "Y", "true", "false"))
                            assessXML.WriteAttributeString("displayTitle", IIf(XDISP_TITLE = "Y", "true", "false"))
                            assessXML.WriteAttributeString("timeLimit", TIME_ALLOWED)
                            assessXML.WriteAttributeString("showAllQuestionsOnSinglePage", IIf(XDISP_SINGLE = "Y", "true", "false"))
                            assessXML.WriteAttributeString("allowRetakeWhenFailed", IIf(XALLOW_RETAKE = "Y", "true", "false"))
                            assessXML.WriteAttributeString("maxNumberOfAttempts", X_TRIES_ALLOWED)
                            assessXML.WriteAttributeString("showBackButton", IIf(XSHOW_BACK = "Y", "true", "false"))
                            assessXML.WriteAttributeString("showExplanations", XSHOW_EXPL)
                            assessXML.WriteAttributeString("explanationLevel", XEXPL_LVL)
                            assessXML.WriteAttributeString("langCd", T_LANGCD)
                            assessXML.WriteAttributeString("kba", KBA_FLG)
                            assessXML.WriteAttributeString("allowedKbaAttempts", XKBA_ATMPTS)
                            assessXML.WriteAttributeString("inactivityTimer", XINACT_TIME)
                            assessXML.WriteAttributeString("questionTimer", XQUEST_TIME)
                            assessXML.WriteAttributeString("scoreAssessment", IIf(XSCORE = "Y", "true", "false"))
                            assessXML.WriteAttributeString("assessmentId", AssessmentId)


                            '	<title>Demonstration Assessment</title>
                            assessXML.WriteStartElement("title")
                            assessXML.WriteString(ASSESS_NAME)
                            assessXML.WriteEndElement()

                            '   <studentName tstRunID="1-HCYUI">Ren Hou</studentName>
                            assessXML.WriteStartElement("studentName")
                            assessXML.WriteAttributeString("tstRunID", TstRunId)
                            assessXML.WriteString(USER_NAME)
                            assessXML.WriteEndElement()

                            '	<stylesheetUrl>sample_assessment.css</stylesheetUrl>
                            assessXML.WriteStartElement("stylesheetUrl")
                            assessXML.WriteString("assessment.css")
                            assessXML.WriteEndElement()

                            '	<finalReportTemplate>
                            '		<![CDATA[
                            '			<h3>Congratulations $STUDENT_NAME, you have finished the demonstration quiz!</h3>
                            '			<div>You have <i>$SUCCESS_STATUS</i> this quiz with a score of <strong>$SCORE%</strong>.</div>
                            '			<div style='color: blue'>Please exit the quiz by clicking the "exit" button above.</div>
                            '			<br /><br />
                            '		]]>
                            '	</finalReportTemplate>
                            assessXML.WriteStartElement("finalReportTemplate")
                            If LCase(Mode) = "survey" Then
                                Select T_LANGCD
                                    Case "ENU"
                                        temp = "<h2>Thank you " & FST_NAME & " for completing this survey.</h2>" & _
                                            "<div id='ExitMsg'>Click the ""Submit Survey"" button to submit your survey answers.</div>" & _
                                            "<br /><br />"
                                    Case "ESN", "ESP"
                                        temp = "<h2>Gracias " & FST_NAME & " por completar la encuesta.</h2>" & _
                                            "<div id='ExitMsg'>Haga clic en el bot�n ""Enviar Encuesta"" para enviar las respuestas de una encuesta.</div>" & _
                                            "<br /><br />"
                                        'Case "KOR"
                                        '    temp = "<h2>Thank you " & FST_NAME & " for completing this survey.</h2>" & _
                                        '        "<div id='ExitMsg'>Click the ""Submit Survey"" button to submit your survey answers.</div>" & _
                                        '        "<br /><br />"
                                    Case Else
                                        temp = "<h2>Thank you " & FST_NAME & " for completing this survey.</h2>" & _
                                           "<div id='ExitMsg'>Click the ""Submit Survey"" button to submit your survey answers.</div>" & _
                                           "<br /><br />"
                                End Select
                            Else
                                Select T_LANGCD
                                    Case "ENU"
                                        temp = "<h2>Congratulations, " & FST_NAME & ", you have finished this exam.</h2>" & _
                                           "<div id='ExitMsg'>Click the ""Submit Exam"" button to submit the exam for scoring.</div>" & _
                                           "<br /><br />"
                                    Case "ESN", "ESP"
                                        temp = "<h2>Congratualations, " & FST_NAME & ", Usted ha terminado este examen.</h2>" & _
                                            "<div id='ExitMsg'>Haga clic en el bot�n ""Enviar el examen"" para que su examen pueda obtener puntaje.</div>" & _
                                            "<br /><br />"
                                    Case Else
                                        temp = "<h2>Congratulations, " & FST_NAME & ", you have finished this exam.</h2>" & _
                                           "<div id='ExitMsg'>Click the ""Submit Exam"" button to submit the exam for scoring.</div>" & _
                                           "<br /><br />"
                                End Select
                            End If
                            assessXML.WriteCData(temp)
                            assessXML.WriteEndElement()
                            '	<!--numberOfQuestionsToRandomlySelect>2</numberOfQuestionsToRandomlySelect-->

                            '	<passingScore>65.0</passingScore>
                            assessXML.WriteStartElement("passingScore")
                            assessXML.WriteString(Format(SCORE_PCT))
                            assessXML.WriteEndElement()
                        End If

                        ' ----
                        ' Question
                        '<question id="1-9AUSC" type="multiple choice single answer">
                        '	<text>What is the address of Rustici Software's home on the web?</text>
                        '</question>
                        If QUES_ID <> sQUES_ID Then
                            Try
                                If InQues = True Then assessXML.WriteEndElement() ' Close out prior question if necessary
                                assessXML.WriteStartElement("question")
                                assessXML.WriteAttributeString("id", QUES_ID)
                                assessXML.WriteAttributeString("type", QUES_TYPE_CD)
                                If QUES_TYPE_CD = "fill in the blank" Then
                                    If Q_MULTILINE_FLG = "Y" Then
                                        If Q_ROWS = "" Or Not IsNumeric(Q_ROWS) Then Q_ROWS = "5"
                                        If Q_COLS = "" Or Not IsNumeric(Q_COLS) Then Q_COLS = "60"
                                        assessXML.WriteAttributeString("textBoxSize", "60")
                                        assessXML.WriteAttributeString("multiline", "true")
                                        assessXML.WriteAttributeString("rows", Q_ROWS)
                                        assessXML.WriteAttributeString("cols", Q_COLS)
                                    Else
                                        If Q_SIZE_LIMIT = "" Or Not IsNumeric(Q_SIZE_LIMIT) Then Q_SIZE_LIMIT = "60"
                                        assessXML.WriteAttributeString("textBoxSize", Q_SIZE_LIMIT)
                                        assessXML.WriteAttributeString("multiline", "false")
                                        assessXML.WriteAttributeString("rows", "")
                                        assessXML.WriteAttributeString("cols", "")
                                    End If
                                End If
                                assessXML.WriteStartElement("text")
                                assessXML.WriteString(QUES_TEXT)
                                assessXML.WriteEndElement()
                            Catch ex As Exception
                                errmsg = errmsg & vbCrLf & "Error writing question. " & ex.ToString
                                results = "Failure"
                            End Try
                        End If
                        If QUES_ID <> "" Then InQues = True

                        ' ----
                        ' Answer
                        If Debug = "Y" Then mydebuglog.Debug("  ANSWER: " & RecCount.ToString & "  ... " & QUES_SEQ_NUM & "/" & ANSR_SEQ_NUM & "   ID: " & ANSR_ID & " - " & ANSR_TEXT)
                        Try
                            If ANSR_ID <> "" Then
                                Select Case QUES_TYPE_CD
                                    Case "slide"
                                        ' Do nothing
                                    Case "matching"
                                        ' ???
                                    Case "fill in the blank"
                                        '<answer letter="a"  id="q4a1" isCorrect="true" fillInTheBlankEvaluationMethod="regular expression">
                                        '	<text>www.adlnet.org</text>
                                        '	<regex>(adlnet.org)$|(ADLNET.ORG)$</regex>
                                        '</answer>
                                        assessXML.WriteStartElement("answer")
                                        assessXML.WriteAttributeString("letter", ANSR_CD)
                                        assessXML.WriteAttributeString("id", ANSR_ID)
                                        If CORRECT_ANSR_FLG = "Y" Then
                                            assessXML.WriteAttributeString("isCorrect", "true")
                                        Else
                                            assessXML.WriteAttributeString("isCorrect", "false")
                                        End If
                                        assessXML.WriteAttributeString("fillInTheBlankEvaluationMethod", "regular expression")
                                        assessXML.WriteStartElement("text")
                                        assessXML.WriteString(ANSR_TEXT)
                                        assessXML.WriteEndElement()
                                        assessXML.WriteStartElement("regex")
                                        assessXML.WriteString("(" & LCase(ANSR_TEXT) & ")$|(" & UCase(ANSR_TEXT) & ")$")
                                        assessXML.WriteEndElement()
                                        assessXML.WriteEndElement()
                                    Case Else
                                        '	<answer letter="a" id="q1a1" isCorrect="true">
                                        '		<text>www.scorm.com</text>
                                        '		<explanation>www.scorm.com is Rustici Software's web address.</explanation>
                                        '	</answer>
                                        assessXML.WriteStartElement("answer")
                                        assessXML.WriteAttributeString("letter", ANSR_CD)
                                        assessXML.WriteAttributeString("id", ANSR_ID)
                                        assessXML.WriteAttributeString("isCorrect", CORRECT_ANSR_FLG)
                                        assessXML.WriteStartElement("text")
                                        assessXML.WriteString(ANSR_TEXT)
                                        assessXML.WriteEndElement()
                                        assessXML.WriteStartElement("explanation")
                                        assessXML.WriteString("")
                                        assessXML.WriteEndElement()
                                        assessXML.WriteEndElement()
                                End Select
                            Else
                                ' No answer was found
                                If QUES_TYPE_CD = "fill in the blank" Then
                                    assessXML.WriteStartElement("answer")
                                    assessXML.WriteAttributeString("letter", ANSR_CD)
                                    assessXML.WriteAttributeString("id", ANSR_ID)
                                    assessXML.WriteAttributeString("isCorrect", CORRECT_ANSR_FLG)
                                    assessXML.WriteAttributeString("fillInTheBlankEvaluationMethod", "regular expression")
                                    assessXML.WriteStartElement("text")
                                    assessXML.WriteString(ANSR_TEXT)
                                    assessXML.WriteEndElement()
                                    assessXML.WriteStartElement("regex")
                                    assessXML.WriteString("(" & LCase(ANSR_TEXT) & ")$|(" & UCase(ANSR_TEXT) & ")$")
                                    assessXML.WriteEndElement()
                                    assessXML.WriteEndElement()
                                End If
                            End If
                        Catch ex As Exception
                            errmsg = errmsg & vbCrLf & "Error writing answer. " & ex.ToString
                            results = "Failure"
                        End Try
                        sQUES_ID = QUES_ID

                    Catch ex As Exception
                        errmsg = errmsg & "Error reading assessment: " & ex.ToString & vbCrLf
                    End Try
                End While

                ' Close document
                If RecCount > 0 Then
                    If InQues Then
                        assessXML.WriteEndElement()
                    End If
                    assessXML.WriteEndElement()
                    assessXML.Flush()
                End If
            End If
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Error locating assessment records. " & ex.ToString
            results = "Failure"
            GoTo CloseOut
        End Try
        dr.Close()

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
        ' Turn results into an XML document
        Dim stream_reader As New StreamReader(memory_stream)
        Try
            memory_stream.Seek(0, SeekOrigin.Begin)
            If memory_stream.Length > 0 And Debug = "Y" Then
                mydebuglog.Debug("Results: " & vbCrLf & stream_reader.ReadToEnd())
                mydebuglog.Debug("Size of results: " & memory_stream.Length.ToString)
            End If
            If memory_stream.Length > 0 And Debug <> "T" Then
                'odoc = New XmlDocument
                'odoc.LoadXml(stream_reader.ReadToEnd())
                If File.Exists("c:\temp\temp.xml") Then File.Delete("c:\temp\temp.xml")
                Dim tempxml As FileStream
                tempxml = New FileStream("c:\temp\temp.xml", FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                memory_stream.WriteTo(tempxml)
                tempxml.Close()
                tempxml = Nothing
                odoc = New XmlDocument
                odoc.Load("c:\temp\temp.xml")
                If Debug <> "Y" And File.Exists("c:\temp\temp.xml") Then File.Delete("c:\temp\temp.xml")
            Else
                Dim resultsDeclare As System.Xml.XmlDeclaration
                Dim resultsRoot As System.Xml.XmlElement
                resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
                odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
                resultsRoot = odoc.CreateElement("assessment")
                odoc.InsertAfter(resultsRoot, resultsDeclare)
                If Debug = "T" Then
                    If memory_stream.Length > 0 Then
                        AddXMLChild(odoc, resultsRoot, "results", "Success")
                    Else
                        AddXMLChild(odoc, resultsRoot, "results", "Failure")
                    End If
                Else
                    AddXMLChild(odoc, resultsRoot, "results", "Failure")
                    AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
                End If
                resultsDeclare = Nothing
                resultsRoot = Nothing
            End If

            ' Close some memory objects
            stream_reader.Close()
            memory_stream.Close()
            assessXML.Close()
            stream_reader = Nothing
            memory_stream = Nothing
            assessXML = Nothing
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug("Error converting results " & ex.ToString)
        End Try

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetExamHTML5 : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GetExamHTML5 : Results: " & results & " for TstRunId: " & TstRunId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug(vbCrLf & "  Results: " & results & " for TstRunId: " & TstRunId & " at " & Now.ToString)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
                mydebuglog.Debug("Error outputing results " & ex.ToString)
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

        ' Close other objects
        Try
            iDoc = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc

        ' Close memory objects
        odoc = Nothing

    End Function

    <WebMethod(Description:="Provides a list of states and their BAC levels for use in course interaction(s)")> _
    Public Function GetStateBACLevels(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function provides a list of states and their BAC levels

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	    - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim i As Integer
        Dim mypath, errmsg, logging As String
        Dim ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' BAC
        Dim DecodedUserId As String
        Dim num_states As Integer
        Dim state(100) As String
        Dim bacLimit(100) As String
        Dim LangCd As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        num_states = 0
        ValidatedUserId = ""
        LangCd = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetStateBACLevels.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Verify Identity
        If Not cmd Is Nothing Then
            Try
                SqlS = "SELECT C.X_REGISTRATION_NUM, C.X_PR_LANG_CD " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            LangCd = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If (ValidatedUserId <> DecodedUserId Or DecodedUserId = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        End If
        If LangCd = "" Then LangCd = "ENU"

        ' ============================================
        ' Process request
        If Not cmd Is Nothing Then
            Try
                SqlS = "SELECT J.NAME, R.REG_TEXT, J.JLEVEL " & _
                "FROM reports.dbo.DB_JURIS J " & _
                "INNER JOIN reports.dbo.DB_REG R ON R.JURIS_ID=J.ROW_ID " & _
                "INNER JOIN reports.dbo.DB_ISSUE I ON I.ROW_ID=R.ISSUE_ID " & _
                "WHERE I.ACCESS_KEY='duiLimit' AND J.COUNTRY_CD<>'ZZZ' AND R.REG_TEXT<>'Information not available' AND J.JLEVEL='State' AND J.LANG_CD='" & LangCd & "' " & _
                "ORDER BY J.SEQ, J.SEQ_S, J.SEQ_L"
                If Debug = "Y" Then mydebuglog.Debug("  Get BACs: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            num_states = num_states + 1
                            state(num_states) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            bacLimit(num_states) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query: " & state(num_states))
                            If state(num_states) = "" Then results = "Failure"
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting BAC. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                Else
                    errmsg = errmsg & "BACs were not found." & vbCrLf
                    dr.Close()
                    results = "Failure"
                End If
                dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try

            ' If no results found, then failure
            If num_states = 0 Then results = "Failure"
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' ============================================
        ' Return the BAC information
        '   <States numStates="#">
        '       <State name="[name]" bacLimit="[limit] />
        '       ...
        '   </States>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        Dim resultsItem As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("States")
        AddXMLAttribute(odoc, resultsRoot, "numStates", num_states.ToString)
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        Try
            ' Add result items - send what was submitted for debugging purposes 
            If Debug <> "T" Then
                For i = 1 To num_states
                    resultsItem = odoc.CreateElement("State")
                    AddXMLAttribute(odoc, resultsItem, "name", state(i))
                    AddXMLAttribute(odoc, resultsItem, "bacLimit", bacLimit(i))
                    resultsRoot.AppendChild(resultsItem)
                Next
            End If
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetStateBACLevels : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GetStateBACLevels : Results: " & results & " for RegId # " & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " at " & Now.ToString)
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
        Return odoc
    End Function

    <WebMethod(Description:="Provides a list of all jurisdictions with BAC levels for use in course interaction(s)")> _
    Public Function GetAllBACLevels(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        ' This function provides a list of states and their BAC levels

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	    - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String
        Dim ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim reader As XmlReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' BAC
        Dim DecodedUserId As String
        Dim num_states As Integer
        Dim state(100) As String
        Dim bacLimit(100) As String
        Dim LangCd As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        num_states = 1
        ValidatedUserId = ""
        LangCd = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetAllBACLevels.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Verify Identity
        If Not cmd Is Nothing Then
            Try
                SqlS = "SELECT C.X_REGISTRATION_NUM, C.X_PR_LANG_CD " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            LangCd = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If (ValidatedUserId <> DecodedUserId Or DecodedUserId = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        End If
        If LangCd = "" Then LangCd = "ENU"

        ' ============================================
        ' Process request
        If Not cmd Is Nothing Then
            Try
                SqlS = "siebeldb.dbo.sp_GenerateGetAllBACLevelsXML"
                If Debug = "Y" Then mydebuglog.Debug("  Get BACs: " & SqlS)
                cmd.CommandText = SqlS
                'cmd.Parameters.AddWithValue("@lang_cd", LangCd)
                cmd.Parameters.AddWithValue("@lang_cd", LangCd)
                cmd.CommandType = Data.CommandType.StoredProcedure

                reader = cmd.ExecuteXmlReader()
                reader.Read()
                'Dim doc As New XmlDocument()
                odoc.Load(reader)
                'doc.Save("test.xml")
                reader.Close()

                If odoc Is Nothing Or odoc.DocumentElement Is Nothing Then
                    errmsg = errmsg & "BACs were not found." & vbCrLf
                    num_states = 0
                    results = "Failure"
                End If

            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try

            ' If no results found, then failure
            If num_states = 0 Then results = "Failure"
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
            reader = Nothing
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' ============================================
        ' Return the BAC information

        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        'Dim resultsItem As System.Xml.XmlElement

        If num_states = 0 Then
            resultsRoot = odoc.CreateElement("glossary")
            AddXMLAttribute(odoc, resultsRoot, "num_terms", "0")
            odoc.InsertAfter(resultsRoot, Nothing)
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
        Else
            ' Create container with results
            resultsRoot = odoc.GetElementsByTagName("Countries")(0)
            resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
            odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
            'resultsItem = odoc.CreateElement("results")
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))

        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetStateBACLevels : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GetStateBACLevels : Results: " & results & " for RegId # " & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " at " & Now.ToString)
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
        Return odoc
    End Function
    '   <WebMethod(MessageName:="GetExercise", Description:="Provides a list of answers and answers of a ID checking exercise ")> _
    Public Function GetIDExercise(ByVal RegId As String, ByVal UserId As String, ByVal ExerciseId As String, ByVal Debug As String) As XmlDocument
        ' This function provides a list of states and their BAC levels

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	    - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging As String
        Dim ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim reader As XmlReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' BAC
        Dim DecodedUserId As String
        Dim num_states As Integer
        Dim state(100) As String
        Dim bacLimit(100) As String
        Dim LangCd As String
        Dim JurisId As String
        Dim CrseId As String
        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        num_states = 1
        ValidatedUserId = ""
        LangCd = ""
        JurisId = "All" 'default to "All"
        CrseId = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut
        End If
        If Debug = "T" Then
            RegId = "TIV8232804P8"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetIDExercise.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Verify Identity
        If Not cmd Is Nothing Then
            Try
                SqlS = "SELECT C.X_REGISTRATION_NUM, C.X_PR_LANG_CD, R.JURIS_ID, R.CRSE_ID " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            LangCd = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            JurisId = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            CrseId = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If (ValidatedUserId <> DecodedUserId Or DecodedUserId = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        End If
        If LangCd = "" Then LangCd = "ENU"

        ' ============================================
        ' Process request
        If Not cmd Is Nothing Then
            Try
                SqlS = "siebeldb.dbo.sp_GenerateGetIDExerciseXML"
                If Debug = "Y" Then mydebuglog.Debug("  Get BACs: " & SqlS)
                cmd.CommandText = SqlS
                cmd.Parameters.AddWithValue("@lang_cd", LangCd)
                cmd.Parameters.AddWithValue("@juris_id", JurisId)
                cmd.Parameters.AddWithValue("@crse_id", CrseId)
                cmd.Parameters.AddWithValue("@exer_id", ExerciseId)

                cmd.CommandType = Data.CommandType.StoredProcedure

                reader = cmd.ExecuteXmlReader()
                reader.Read()
                odoc.Load(reader)
                reader.Close()

                If odoc.DocumentElement Is Nothing Then
                    errmsg = errmsg & "Exercise items were not found." & vbCrLf
                    num_states = 0
                    results = "Failure"
                End If

            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try

            ' If no results found, then failure
            If num_states = 0 Then results = "Failure"
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
            reader = Nothing
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' ============================================
        ' Return the Exercise information

        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        'Dim resultsItem As System.Xml.XmlElement

        If odoc.DocumentElement Is Nothing Then
            resultsRoot = odoc.CreateElement("exercises")
            'AddXMLAttribute(odoc, resultsRoot, "num_terms", "0")
            odoc.InsertAfter(resultsRoot, Nothing)
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
        Else
            ' Create container with results
            resultsRoot = odoc.GetElementsByTagName("exercises")(0)
            resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
            odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))

        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetStateBACLevels : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GetStateBACLevels : Results: " & results & " for RegId # " & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " at " & Now.ToString)
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
        Return odoc
    End Function

    <WebMethod(MessageName:="GetIDExercise", Description:="Provides a list of answers and answers of a ID checking exercise ")> _
    Public Function GetIDExercise(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument
        Return GetIDExercise(RegId, UserId, "E1", Debug)  'Default the ExerciseId to "E1" (ID Exercise)
    End Function

    <WebMethod(Description:="Retrieves a list of local government regulations in a state for use in course interaction(s)")> _
    Public Function GetLocalRegs(ByVal RegId As String, ByVal UserId As String, ByVal Debug As String) As XmlDocument

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   Debug	    - "Y", "N" or "T"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results, temp As String
        Dim i As Integer
        Dim mypath, errmsg, logging As String
        Dim ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Regulatory
        Dim JURIS_ID, LANG_CD As String
        Dim DecodedUserId As String
        Dim num_regs As Integer
        Dim num_locals As Integer
        Dim jname(1000) As String
        Dim reg(1000) As String
        Dim accesskey(1000) As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        num_regs = 0
        num_locals = 0
        ValidatedUserId = ""
        temp = ""
        JURIS_ID = ""
        LANG_CD = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-EUP5Q"
            UserId = "==QQPZzMwMjMxEzMPRlU"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GetLocalRegs_debug")
            If temp = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetLocalRegs.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            End If
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Verify Registrants Identity
        If Not cmd Is Nothing Then
            Try
                SqlS = "SELECT C.X_REGISTRATION_NUM, R.JURIS_ID, " & _
                "(SELECT CASE WHEN C.X_PR_LANG_CD IS NULL OR C.X_PR_LANG_CD='' THEN S.LANG_ID ELSE C.X_PR_LANG_CD END) AS LANG_ID " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR S ON S.ROW_ID=R.TRAIN_OFFR_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            JURIS_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            LANG_CD = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If (ValidatedUserId <> DecodedUserId Or DecodedUserId = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        End If

        ' ============================================
        ' Process request
        If Not cmd Is Nothing Then
            ' Retrieve count of regulations
            Try
                SqlS = "SELECT COUNT(*) " & _
                "FROM reports.dbo.DB_JURIS " & _
                "WHERE JLEVEL='Local' AND ALT_JURIS_ID='" & JURIS_ID & "' AND LANG_CD='" & LANG_CD & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get local regs: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            num_locals = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting regulations. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                End If
                dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try

            ' Retrieve regulations
            Try
                temp = ""
                SqlS = "SELECT J.NAME, I.ACCESS_KEY, R.REG_TEXT  " & _
                "FROM reports.dbo.DB_JURIS J " & _
                "INNER JOIN reports.dbo.DB_REG R ON R.JURIS_ID=J.ROW_ID " & _
                "INNER JOIN reports.dbo.DB_ISSUE I ON I.ROW_ID=R.ISSUE_ID " & _
                "WHERE J.JLEVEL='Local' AND J.ALT_JURIS_ID='" & JURIS_ID & "' AND J.LANG_CD='" & LANG_CD & "' AND R.REG_TEXT IS NOT NULL " & _
                "ORDER BY J.NAME, I.SEQ"
                If Debug = "Y" Then mydebuglog.Debug("  Get local regs: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            num_regs = num_regs + 1
                            If num_regs > 1000 Then
                                ReDim jname(num_regs)
                                ReDim reg(num_regs)
                                ReDim accesskey(num_regs)
                            End If
                            jname(num_regs) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            accesskey(num_regs) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            reg(num_regs) = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            temp = jname(num_regs)
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error getting regulations. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                    If Debug = "Y" Then
                        mydebuglog.Debug("  > Number regulations: " & Str(num_regs))
                        mydebuglog.Debug("  > Number localities: " & Str(num_locals))
                    End If
                Else
                    errmsg = errmsg & "BACs were not found." & vbCrLf
                    dr.Close()
                    results = "Failure"
                End If
                dr.Close()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try

            ' If no results found, then failure
            If num_regs = 0 Then results = "Failure"
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

        ' ============================================
        ' Return the local government regulatory information
        '   <Localities numLocalities="#">
        '       <Locality name="[name]" ... />
        '       ...
        '   </Localities>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        Dim resultsItem As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("Localities")
        If Debug <> "T" Then AddXMLAttribute(odoc, resultsRoot, "numLocalities", num_locals.ToString)
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        Try
            ' Add result items - send what was submitted for debugging purposes 
            If Debug <> "T" And results <> "Failure" Then
                temp = ""
                resultsItem = odoc.CreateElement("Locality")
                For i = 1 To num_regs
                    If temp <> jname(i) Then
                        If i <> 1 Then
                            resultsRoot.AppendChild(resultsItem)
                            resultsItem = odoc.CreateElement("Locality")
                        End If
                        AddXMLAttribute(odoc, resultsItem, "name", jname(i))
                    End If
                    AddXMLAttribute(odoc, resultsItem, accesskey(i), reg(i))
                    temp = jname(i)
                Next
                resultsRoot.AppendChild(resultsItem)
            End If
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GetLocalRegs :  Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GetLocalRegs : Results: " & results & " for RegId # " & RegId & " and JurisId " & JURIS_ID)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & results & " for RegId # " & RegId & " and JurisId " & JURIS_ID & " at " & Now.ToString)
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
        Return odoc
    End Function

    <WebMethod(Description:="Saves the supplied course note to a database table")> _
    Public Function SaveCourseNote(ByVal RegId As String, ByVal UserId As String, ByVal NoteText As String, _
            ByVal ScreenId As String, ByVal Debug As String) As XmlDocument
        ' This function saves the supplied course note to the database table elearning.ELN_NOTES.
        ' Before saving the record it removes any existing one.

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the user
        '   NoteText	    - The text of the note   
        '   ScreenId        - The id of the screen
        '   Debug           - "Y", "N" or "T"

        ' Returns:  "success" or "failure"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging, DecodedUserId, ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Failure"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        ValidatedUserId = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            DecodedUserId = "RTO31123036OA"
            ScreenId = "001"
            NoteText = "Test"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
            NoteText = Trim(HttpUtility.UrlDecode(NoteText))
            If InStr(NoteText, "%") > 0 Then NoteText = Trim(HttpUtility.UrlDecode(NoteText))
            If InStr(NoteText, "%") > 0 Then NoteText = Trim(NoteText)
            ScreenId = Trim(HttpUtility.UrlDecode(ScreenId))
            If InStr(ScreenId, "%") > 0 Then ScreenId = Trim(HttpUtility.UrlDecode(ScreenId))
            If InStr(ScreenId, "%") > 0 Then ScreenId = Trim(ScreenId)
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\SaveCourseNote.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  ScreenId: " & ScreenId)
                mydebuglog.Debug("  NoteText: " & Left(NoteText, 10) & "...")
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If RegId = "" Then
            errmsg = errmsg & vbCrLf & "Registration Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If
        If UserId = "" And Debug <> "T" Then
            errmsg = errmsg & vbCrLf & "User Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If
        If ScreenId = "" Then
            errmsg = errmsg & vbCrLf & "Screen Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If

        ' Commentted out the Note check to allow empty note; 5/24/17; Ren Hou; Per Chris Bobbitt;
        'If NoteText = "" Then
        '    errmsg = errmsg & vbCrLf & "Note not supplied. "
        '    results = "Failure"
        '    GoTo CloseOut2
        'End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Verify User Identity
        If Not cmd Is Nothing And Debug <> "T" Then
            Try
                SqlS = "SELECT C.X_REGISTRATION_NUM " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If (ValidatedUserId <> DecodedUserId Or DecodedUserId = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        End If

        ' ============================================
        ' Process data

        ' ELN_NOTES Field Structure:
        '   REG_ID          FK to siebeldb.dbo.CX_SESS_REG.ROW_ID
        '   USER_ID         FK to siebeldb.dbo.S_CONTACT.X_REGISTRATION_NUM
        '   SCREEN          Screen Id
        '   CREATED         Datetime when the note was saved
        '   NOTE_TEXT       Text of the note          

        If Not cmd Is Nothing Then
            ' Remove any existing note for this screen/registration
            Try
                SqlS = "DELETE FROM elearning.dbo.ELN_NOTES " & _
                    "WHERE REG_ID='" & RegId & "' " & _
                    "AND SCREEN='" & ScreenId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Delete existing note: " & SqlS)
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
            End Try

            ' Add the note
            Try
                SqlS = "INSERT elearning.dbo.ELN_NOTES " & _
                    "(REG_ID, USER_ID, SCREEN, CREATED, NOTE_TEXT)" & _
                    "VALUES ('" & RegId & "','" & DecodedUserId & "','" & ScreenId & "'," & _
                    "GETDATE(),'" & SqlString(NoteText) & "')"
                If Debug = "Y" Then mydebuglog.Debug("  Inserting note: " & SqlS)
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                If returnv > 0 Then
                    results = "Success"
                End If
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
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
        ' Return success or failure
        '   <notes results="" error=""/>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("notes")
        Try
            AddXMLAttribute(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLAttribute(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLAttribute(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("SaveCourseNote : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("SaveCourseNote : Results: " & results & " for RegId # " & RegId & " and screen " & ScreenId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " and screen " & ScreenId & " at " & Now.ToString)
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
        Return odoc
    End Function

    <WebMethod(Description:="Saves CardMsg info")> _
    Public Function SaveCardMsgInfo(ByVal RegId As String, ByVal UserId As String, ByVal CardMsg As String, _
            ByVal Debug As String) As XmlDocument
        ' This function saves the supplied course note to the database table elearning.ELN_NOTES.
        ' Before saving the record it removes any existing one.

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the user
        '   CardMsg	        - The value for CardMsg (saved in siebeldb.dbo.CX_SESS_REG.CARD_MSG)  
        '   Debug           - "Y", "N" or "T"

        ' Returns:  "success" or "failure"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging, DecodedUserId, ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Failure"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        ValidatedUserId = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            DecodedUserId = "RTO31123036OA"
            CardMsg = "Test"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
            CardMsg = Trim(HttpUtility.UrlDecode(CardMsg))
            If InStr(CardMsg, "%") > 0 Then CardMsg = Trim(HttpUtility.UrlDecode(CardMsg))
            If InStr(CardMsg, "%") > 0 Then CardMsg = Trim(CardMsg)
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\SaveCardMsgInfo.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  CardMsg: " & Left(CardMsg, 30) & "...")
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If RegId = "" Then
            errmsg = errmsg & vbCrLf & "Registration Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If
        If UserId = "" And Debug <> "T" Then
            errmsg = errmsg & vbCrLf & "User Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If
        ' Commentted out the Note check to allow empty note; 5/24/17; Ren Hou; Per Chris Bobbitt;
        'If NoteText = "" Then
        '    errmsg = errmsg & vbCrLf & "Note not supplied. "
        '    results = "Failure"
        '    GoTo CloseOut2
        'End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Verify User Identity
        If Not cmd Is Nothing And Debug <> "T" Then
            Try
                SqlS = "SELECT C.X_REGISTRATION_NUM " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If (ValidatedUserId <> DecodedUserId Or DecodedUserId = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        End If

        ' ============================================
        ' Process data

        If Not cmd Is Nothing Then
            ' Save Crad Message value
            Try
                SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                    " SET CARD_MSG = '" & CardMsg & "' " & _
                    " WHERE ROW_ID='" & RegId & "' " '& _
                '" AND CONTACT_ID='" & DecodedUserId & "' " 'UserId
                If Debug = "Y" Then mydebuglog.Debug("  Delete existing note: " & SqlS)
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                If returnv > 0 Then
                    results = "Success"
                End If

            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
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
        ' Return success or failure
        '   <notes results="" error=""/>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("notes")
        Try
            AddXMLAttribute(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLAttribute(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            AddXMLAttribute(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("SaveCardMsgInfo : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("SaveCardMsgInfo : Results: " & results & " for RegId # " & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " at " & Now.ToString)
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
        Return odoc
    End Function

    ' =================================================
    ' EMAIL
    Public Function PrepareMail(ByVal FromEmail As String, ByVal ToEmail As String, ByVal Subject As String, _
        ByVal Body As String, ByVal Debug As String, ByRef fs As FileStream) As Boolean
        ' This function wraps message info into the XML necessary to call the SendMail web service function.
        ' This is used by other services executing from this application.
        ' Assumptions:  Create a record in MESSAGES and IDs are unknown 
        Dim wp As String

        ' Web service declarations
        Dim EmailService As New com.certegrity.cloudsvc.basic.Service

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
        wp = wp & "<ReplyTo></ReplyTo>"
        wp = wp & "<Subject>" & Subject & "</Subject>"
        wp = wp & "<Body>" & Body & "</Body>"
        wp = wp & "<Format></Format>"
        wp = wp & "</EMailMessage></EMailMessageList>"
        If Debug = "Y" Then mydebuglog.Debug("Email XML: " & wp)

        PrepareMail = EmailService.SendMail(wp)

    End Function

    <WebMethod(Description:="Gets the student note for the specified screen")> _
    Public Function GetCourseNote(ByVal RegId As String, ByVal UserId As String, _
                ByVal ScreenId As String, ByVal Debug As String) As XmlDocument
        ' This function gets the course note from the database table elearning.ELN_NOTES
        ' for the supplied screen id.

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the user
        '   ScreenId        - The id of the screen
        '   Debug           - "Y", "N" or "T"

        ' Returns:  "success" or "failure"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim mypath, errmsg, logging, DecodedUserId, ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Notes declarations
        Dim NoteText, NoteCreated As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Failure"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        NoteText = ""
        NoteCreated = ""
        ValidatedUserId = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            DecodedUserId = "RTO31123036OA"
            ScreenId = "001"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
            ScreenId = Trim(HttpUtility.UrlDecode(ScreenId))
            If InStr(ScreenId, "%") > 0 Then ScreenId = Trim(HttpUtility.UrlDecode(ScreenId))
            If InStr(ScreenId, "%") > 0 Then ScreenId = Trim(ScreenId)
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetCourseNote.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  ScreenId: " & ScreenId)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If RegId = "" Then
            errmsg = errmsg & vbCrLf & "Registration Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If
        If UserId = "" And Debug <> "T" Then
            errmsg = errmsg & vbCrLf & "User Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If
        If ScreenId = "" Then
            errmsg = errmsg & vbCrLf & "Screen Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Verify User Identity
        If Not cmd Is Nothing And Debug <> "T" Then
            Try
                SqlS = "SELECT C.X_REGISTRATION_NUM " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If (ValidatedUserId <> DecodedUserId Or DecodedUserId = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        End If

        ' ============================================
        ' Process data

        ' ELN_NOTES Field Structure:
        '   REG_ID          FK to siebeldb.dbo.CX_SESS_REG.ROW_ID
        '   USER_ID         FK to siebeldb.dbo.S_CONTACT.X_REGISTRATION_NUM
        '   SCREEN          Screen Id
        '   CREATED         Datetime when the note was saved
        '   NOTE_TEXT       Text of the note          

        If Not cmd Is Nothing Then
            ' Locate the note for the user and screen
            Try
                SqlS = "SELECT NOTE_TEXT, CONVERT(VARCHAR,CREATED,21) " & _
                "FROM elearning.dbo.ELN_NOTES " & _
                "WHERE REG_ID='" & RegId & "' AND SCREEN='" & ScreenId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get note: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            NoteText = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            NoteCreated = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            results = "Success"
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error locating record. " & ex.ToString & vbCrLf
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
            If Debug = "Y" Then mydebuglog.Debug("   ... Found note created on: " & NoteCreated)
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
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
        ' Return success or failure
        '   <notes results="" error=""/>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("note")
        Try
            If Debug = "T" Then
                If NoteText = "Test" Then results = "Success"
            End If
            AddXMLAttribute(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLAttribute(odoc, resultsRoot, "error", Trim(errmsg))
            If Debug <> "T" Then
                resultsRoot.InnerText = "<![CDATA[" & Trim(NoteText) & "]]>"
                AddXMLChild(odoc, resultsRoot, "datetime", Trim(NoteCreated))
            End If
        Catch ex As Exception
            AddXMLAttribute(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        ' ============================================
        ' Close the log file if any
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " and screen " & ScreenId & " at " & Now.ToString)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
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
        ' Return results
        Return odoc
    End Function

    <WebMethod(Description:="Gets the student note for the specified screen")> _
    Public Function GetAllCourseNotes(ByVal RegId As String, ByVal UserId As String, _
                ByVal SortOrder As String, ByVal SortDirection As String, ByVal Debug As String) As XmlDocument
        ' This function gets all of the course notes from the database table elearning.ELN_NOTES
        ' for the supplied registration id.

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the user
        '   SortOrder       - Sort order of the list (either "date" or "screen")
        '   SortDirection   - Sort direction of the list (either "asc" or "desc")
        '   Debug           - "Y", "N" or "T"

        ' Returns:  "success" or "failure"

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim i As Integer
        Dim mypath, errmsg, logging, DecodedUserId, ValidatedUserId As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Notes declarations
        Dim NumNotes As Integer
        Dim NoteText(1000) As String
        Dim NoteCreated(1000) As String
        Dim NoteScreen(1000) As String

        ' Results declarations
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Failure"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        NumNotes = 0
        ValidatedUserId = ""

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If (RegId = "" Or UserId = "") And Debug <> "T" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter(s) missing. "
            GoTo CloseOut2
        End If
        If Debug = "T" Then
            RegId = "1-9ONMD"
            DecodedUserId = "RTO31123036OA"
        Else
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
            SortOrder = LCase(Trim(HttpUtility.UrlDecode(SortOrder)))
            If InStr(SortOrder, "%") > 0 Then SortOrder = Trim(HttpUtility.UrlDecode(SortOrder))
            If InStr(SortOrder, "%") > 0 Then SortOrder = Trim(SortOrder)
            SortDirection = LCase(Trim(HttpUtility.UrlDecode(SortDirection)))
            If InStr(SortDirection, "%") > 0 Then SortDirection = Trim(HttpUtility.UrlDecode(SortDirection))
            If InStr(SortDirection, "%") > 0 Then SortDirection = Trim(SortOrder)
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GetAllCourseNotes.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
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
                mydebuglog.Debug("  RegId: " & RegId)
                mydebuglog.Debug("  UserId: " & UserId)
                mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
                mydebuglog.Debug("  SortOrder: " & SortOrder)
                mydebuglog.Debug("  SortDirection: " & SortDirection)
            End If
        End If

        ' ============================================
        ' Validate Parameters
        If RegId = "" Then
            errmsg = errmsg & vbCrLf & "Registration Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If
        If UserId = "" And Debug <> "T" Then
            errmsg = errmsg & vbCrLf & "User Id not supplied. "
            results = "Failure"
            GoTo CloseOut2
        End If
        If SortOrder <> "" Then
            If SortOrder <> "date" And SortOrder <> "screen" Then
                errmsg = errmsg & vbCrLf & "Sort order invalid. "
                results = "Failure"
                GoTo CloseOut2
            End If
        Else
            SortOrder = "screen"
        End If
        If SortDirection <> "" Then
            If SortDirection <> "asc" And SortDirection <> "desc" Then
                errmsg = errmsg & vbCrLf & "Sort direction invalid. "
                results = "Failure"
                GoTo CloseOut2
            End If
        Else
            SortDirection = "asc"
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIBDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Verify User Identity
        If Not cmd Is Nothing And Debug <> "T" Then
            Try
                SqlS = "SELECT C.X_REGISTRATION_NUM " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "WHERE R.ROW_ID='" & RegId & "'"
                If Debug = "Y" Then mydebuglog.Debug("  Get registration: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            If Debug = "Y" Then mydebuglog.Debug("  > Found record on query")
                            ValidatedUserId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
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
            If Debug = "Y" Then mydebuglog.Debug("   ... ValidatedUserId: " & ValidatedUserId)

            ' -----
            ' Verify the user
            If (ValidatedUserId <> DecodedUserId Or DecodedUserId = "") And Debug <> "T" Then
                results = "Failure"
                errmsg = errmsg & "User not validated: " & DecodedUserId & ". Should have been: " & ValidatedUserId
                GoTo CloseOut
            End If
        End If

        ' ============================================
        ' Process data

        ' ELN_NOTES Field Structure:
        '   REG_ID          FK to siebeldb.dbo.CX_SESS_REG.ROW_ID
        '   USER_ID         FK to siebeldb.dbo.S_CONTACT.X_REGISTRATION_NUM
        '   SCREEN          Screen Id
        '   CREATED         Datetime when the note was saved
        '   NOTE_TEXT       Text of the note          

        If Not cmd Is Nothing Then
            ' Locate the note for the user and screen
            Try
                SqlS = "SELECT NOTE_TEXT, CONVERT(VARCHAR,CREATED,21), SCREEN " & _
                "FROM elearning.dbo.ELN_NOTES " & _
                "WHERE REG_ID='" & RegId & "' "
                If SortOrder = "screen" Then SqlS = SqlS & "ORDER BY SCREEN "
                If SortOrder = "date" Then SqlS = SqlS & "ORDER BY CREATED "
                If SortDirection <> "" Then SqlS = SqlS & UCase(SortDirection)
                If Debug = "Y" Then mydebuglog.Debug("  Get note: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            NumNotes = NumNotes + 1
                            NoteText(NumNotes) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            NoteCreated(NumNotes) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            NoteScreen(NumNotes) = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        Catch ex As Exception
                            results = "Failure"
                            errmsg = errmsg & "Error locating record. " & ex.ToString & vbCrLf
                            GoTo CloseOut
                        End Try
                    End While
                    If NumNotes > 0 Then results = "Success"
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
            If Debug = "Y" Then mydebuglog.Debug("   ... # Notes Found: " & NumNotes.ToString)
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            'dr.Close()
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
        ' Return success or failure
        '   <notes number="" results="" error="">
        '	<note text="" datetime="" screen="" />
        '   </notes>
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        Dim resultsItem As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("notes")
        Try
            ' <notes>
            If Debug = "T" Then
                If NumNotes > 0 Then results = "Success"
            Else
                AddXMLAttribute(odoc, resultsRoot, "number", NumNotes.ToString)
            End If
            AddXMLAttribute(odoc, resultsRoot, "results", Trim(results))
            If errmsg <> "" Then AddXMLAttribute(odoc, resultsRoot, "error", Trim(errmsg))
            odoc.InsertAfter(resultsRoot, resultsDeclare)
            ' <note>
            If Debug <> "T" Then
                For i = 1 To NumNotes
                    resultsItem = odoc.CreateElement("note")
                    'AddXMLAttribute(odoc, resultsItem, "datetime", Trim(NoteCreated(i)))
                    'AddXMLAttribute(odoc, resultsItem, "screen", Trim(NoteScreen(i)))
                    If Debug <> "T" Then resultsItem.InnerText = "<![CDATA[" & Trim(NoteText(i)) & "]]>"
                    AddXMLChild(odoc, resultsItem, "datetime", Trim(NoteCreated(i)))
                    AddXMLChild(odoc, resultsItem, "screen", Trim(NoteScreen(i)))
                    resultsRoot.AppendChild(resultsItem)
                Next
            End If
        Catch ex As Exception
            AddXMLAttribute(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try


        ' ============================================
        ' Close the log file if any
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for RegId # " & RegId & " and number found: " & NumNotes.ToString & " at " & Now.ToString)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
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
        ' Return results
        Return odoc
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
    Public Shared Function Soundex1(ByVal s As String) As String
        Const CodeTab = " 123 12  22455 12623 1 2 2"
        '                abcdefghijklnmopqrstuvwxyz
        Dim c As Integer
        Dim p As Integer : p = 1
        Do
            If p > Len(s) Then Soundex1 = s : Exit Function
            c = Asc(Mid(s, p, 1))
            p = p + 1
            If c >= 65 And c <= 90 Then Exit Do
            If c >= 97 And c <= 122 Then c = c - 32 : Exit Do
        Loop
        Dim ss As String, PrevCode As String
        ss = Chr(c)
        PrevCode = Mid$(CodeTab, c - 64, 1)
        Do While Len(ss) < 4 And p <= Len(s)
            c = Asc(Mid(s, p))
            If c >= 65 And c <= 90 Then
                ' nop
            ElseIf c >= 97 And c <= 122 Then
                c = c - 32
            Else
                c = 0
            End If
            Dim Code As String : Code = "?"
            If c <> 0 Then
                Code = Mid(CodeTab, c - 64, 1)
                If Code <> " " And Code <> PrevCode Then ss = ss & Code
            End If
            PrevCode = Code
            p = p + 1
        Loop
        If Len(ss) < 4 Then ss = ss & RepeatS(4 - Len(ss), "0")
        Soundex1 = ss
    End Function

    Public Shared Function RepeatS(ByVal instr As String, ByVal n As Integer) As String
        Dim result = String.Empty
        Dim i As Integer
        For i = 0 To n - 1
            result += instr
        Next
        Return result
    End Function

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


    Private Class HciDMSDocument
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

End Class
