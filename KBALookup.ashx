<%@ WebHandler Language="VB" Class="KBALookup" %>

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

Public Class KBALookup : Implements IHttpHandler

    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        ' This function provides questions and answers for a Knowledge-Based Authentication 
        ' system used to validate the identity of an online class attendee and returns the
        ' data in a JSON format

        ' The input parameters are as follows:
        '   RegId           - The CX_SESS_REG.ROW_ID of the attendee
        '   UserId          - Base64 encoded, reversed S_CONTACT.X_REGISTRATION_NUM of the
        '                       user.
        '   CrseType        - Indicates whether this is "C"ourse or "A"ssessment questions
        '   CrseId          - The id of the course or assessment
        '   callback        - The name of the Javascript callback function in which to wrap the resulting JSON 
        '   Debug           - "Y", "N" or "T"

        ' ============================================
        ' Declare variables

        ' General
        Dim results As String
        Dim errmsg, logging As String
        Dim RegId As String
        Dim UserId As String
        Dim Debug As String
        Dim Extension As String
        Dim CrseType, CrseId, JurisNum As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer
        Dim ktable As New DataTable
        Dim kdataset As New DataSet

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' KBA declarations
        Dim DecodedUserId As String
        Dim num_questions As Integer
        Dim ques_text(100) As String
        Dim ansr_text(100) As String
        Dim jdoc, callback As String

        ' ============================================
        ' Variable setup
        Debug = ""
        RegId = ""
        JurisNum = "0"
        UserId = ""
        Extension = ""
        errmsg = ""
        jdoc = ""
        Logging = "Y"
        ErrMsg = ""
        results = "Success"
        SqlS = ""
        returnv = 0
        DecodedUserId = ""
        CrseType = ""
        CrseId = ""
        callback = "KBAData"
        'Debug = "Y"
        'RegId = "1-EZI59"
        'UserId = "==QQPZzMwMjMxEzMPRlU"
        'CrseType = "A"
        'CrseId = "A2EHOYR"

        ' ============================================
        ' Open log file if applicable
        If logging = "Y" Then
            logfile = "C:\Logs\YourServiceName.log"
            Try
                log4net.GlobalContext.Properties("LogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try
        End If

        ' ============================================
        ' Get parameters
        If Not context.Request.QueryString("Debug") Is Nothing Then
            Debug = context.Request.QueryString("Debug")
            If Debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            End If
        End If

        If Not context.Request.QueryString("callback") Is Nothing Then
            callback = context.Request.QueryString("callback")
        End If

        If Not context.Request.QueryString("RegId") Is Nothing Then
            RegId = context.Request.QueryString("RegId")
        Else
            If Debug <> "T" Then
                If RegId = "" Then
                    errmsg = "No registration specified"
                    GoTo DisplayErrorMsg
                End If
            Else
                RegId = "1-EZI59"
            End If
        End If

        If Not context.Request.QueryString("UserId") Is Nothing Then
            UserId = context.Request.QueryString("UserId")
        Else
            If Debug <> "T" Then
                If UserId = "" Then
                    errmsg = "No user id specified"
                    GoTo DisplayErrorMsg
                End If
            Else
                UserId = "==QQPZzMwMjMxEzMPRlU"
                DecodedUserId = FromBase64(ReverseString(UserId))
            End If
        End If

        If Not context.Request.QueryString("CrseType") Is Nothing Then
            CrseType = context.Request.QueryString("CrseType")
        Else
            If Debug <> "T" Then
                If CrseType = "" Then
                    errmsg = "No course type specified"
                    GoTo DisplayErrorMsg
                End If
            Else
                CrseType = "A"
            End If
        End If

        If Not context.Request.QueryString("CrseId") Is Nothing Then
            CrseId = context.Request.QueryString("CrseId")
        Else
            If Debug <> "T" Then
                If UserId = "" Then
                    errmsg = "No course id specified"
                    GoTo DisplayErrorMsg
                End If
            Else
                CrseId = "A2EHOYR"
            End If
        End If

        ' ============================================
        ' Check parameters
        Debug = UCase(Left(Debug, 1))
        If RegId = "" Or UserId = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "Parameter()s missing"
            GoTo CloseOut2
        End If
        If Debug <> "T" Then
            RegId = Trim(HttpUtility.UrlEncode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(HttpUtility.UrlDecode(RegId))
            If InStr(RegId, "%") > 0 Then RegId = Trim(RegId)
            RegId = EncodeParamSpaces(RegId)
            UserId = Trim(HttpUtility.UrlEncode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(HttpUtility.UrlDecode(UserId))
            If InStr(UserId, "%") > 0 Then UserId = Trim(UserId)
            DecodedUserId = FromBase64(ReverseString(UserId))
            CrseId = Trim(HttpUtility.UrlEncode(CrseId))
            If InStr(CrseId, "%") > 0 Then CrseId = Trim(HttpUtility.UrlDecode(CrseId))
            If InStr(CrseId, "%") > 0 Then CrseId = Trim(CrseId)
            CrseType = UCase(Left(CrseType, 1))
            If CrseType <> "A" And CrseType <> "C" Then CrseType = "C"
        End If
        If Debug = "Y" Then
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  Debug: " & Debug)
            mydebuglog.Debug("  RegId: " & RegId)
            mydebuglog.Debug("  UserId: " & UserId)
            mydebuglog.Debug("  CrseType: " & CrseType)
            mydebuglog.Debug("  CrseId: " & CrseId)
            mydebuglog.Debug("  Decoded UserId: " & DecodedUserId)
            mydebuglog.Debug("  callback: " & callback & vbCrLf)
        End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
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
        ' Retrieve data
        If Not cmd Is Nothing Then
            ' Retrieve the number of questions required
            Try
                If CrseType = "A" Then
                    SqlS = "SELECT X_KBA_QUES_NUM " & _
                        "FROM siebeldb.dbo.S_CRSE_TST " & _
                        "WHERE ROW_ID='" & CrseId & "'"
                Else
                    SqlS = "SELECT JC.KBA_QUESTIONS " & _
                        "FROM siebeldb.dbo.CX_SESS_REG S  " & _
                        "LEFT OUTER JOIN siebeldb.dbo.CX_JURIS_CRSE JC ON JC.JURIS_ID=S.JURIS_ID AND JC.CRSE_ID=S.CRSE_ID " & _
                        "WHERE S.ROW_ID='" & RegId & "'"
                End If
                If Debug = "Y" Then mydebuglog.Debug("  Get number questions required: " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                While dr.Read()
                    Try
                        JurisNum = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        If JurisNum = "" Then JurisNum = "0"
                    Catch ex As Exception
                        results = "Failure"
                        errmsg = errmsg & "Error getting questions. " & ex.ToString & vbCrLf
                        GoTo CloseOut
                    End Try
                End While
                Try
                    dr.Close()
                Catch ex As Exception
                End Try
                If Debug = "Y" Then mydebuglog.Debug("  > JurisNum: " & JurisNum & vbCrLf)
            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try

            ' Exit if no questions required
            If JurisNum = "0" Then
                jdoc = callback & "([]);"
                GoTo CloseOut
            End If

            ' Retrieve KBA questions
            Try
                SqlS = "EXEC reports.dbo.OpenHCIKeys; SELECT TOP " & JurisNum & " Q.QUES_TEXT, reports.dbo.HCI_Decrypt(A.ENC_ANSR_TEXT) AS ANSR_TEXT " &
                    "FROM elearning.dbo.KBA_QUES Q " &
                    "LEFT OUTER JOIN elearning.dbo.KBA_ANSR A ON A.QUES_ID=Q.ROW_ID " &
                    "WHERE A.USER_ID='" & DecodedUserId & "' AND A.REG_ID='" & RegId & "' " &
                    "ORDER BY NEWID()"
                If Debug = "Y" Then mydebuglog.Debug("  Get questions/answers: " & SqlS & vbCrLf)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    ' Process into dataset and convert to JSON
                    ktable.Load(dr)
                    ktable.TableName = "KBA"
                    num_questions = ktable.Rows.Count
                    kdataset.Tables.Add(ktable)
                    jdoc = DataSetToJSON(kdataset)
                Else
                    errmsg = errmsg & "Questions were not found." & vbCrLf
                    dr.Close()
                    results = "Failure"
                End If

                ' Fix null at end of string
                If jdoc.IndexOf("null") > 0 Then
                    'jdoc = jdoc.Replace(",null]", "," & num_questions.ToString & "]")
                    jdoc = jdoc.Replace(",null]", "]")
                End If
                jdoc = jdoc.Replace("""Key"":", "")
                jdoc = jdoc.Replace(",""Value""", "")
                jdoc = callback & "(" & jdoc & ");"

                ' Close datareader
                Try
                    dr.Close()
                Catch ex As Exception
                End Try

            Catch oBug As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
                results = "Failure"
            End Try
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
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("KBALookup.ashx : Error: " & Trim(errmsg))
        myeventlog.Info("KBALookup.ashx : Results: " & results & " for user id: " & DecodedUserId & "  and reg id:" & RegId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  JDOC: " & jdoc & vbCrLf)
                mydebuglog.Debug("Results: " & results & " for user id: " & DecodedUserId & "  and reg id:" & RegId)
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


DisplayErrorMsg:
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
            context.Response.ContentType = "application/json"
            'context.Response.ContentEncoding = Encoding.Unicode
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


    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class