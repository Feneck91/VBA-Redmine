VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Redmine Endpoint
Public BaseUri As String
'Redmine API Key
Public ApiKey As String
'Redmine Proxy
Public Proxy As String
'Redmine Proxy's login
Public ProxyLogin As String
'Redmine Proxy's password
Public ProxyPassword As String
' Limit the informations bring back to the server. It prevent application to hang if there is too many informations to retrieve
Public LimitRequest As Long
' Internal json instance
Private json As Dictionary
' Current loaded projects cache. Use to don't requets many times.
Private projects As Collection
' Current loaded custom fields cache. Use to don't requets many times.
Private customFields As Collection
' Current loaded custom sttus fields cache. Use to don't requets many times.
Private statuses As Collection

'Constructor
Private Sub Class_Initialize()
  Me.BaseUri = vbNullString
  Me.ApiKey = vbNullString
  Me.LimitRequest = 2000
  Set projects = Nothing
  Set customFields = Nothing
  Set statuses = Nothing
End Sub

'Destructor
Private Sub Class_Terminate()
  Set json = Nothing
  Set projects = Nothing
  Set customFields = Nothing
  Set statuses = Nothing
End Sub

Public Function CreateUser(ByRef user As RedmineUser, _
                           Optional ByVal password As String = "", _
                           Optional ByVal auth_source_id As String = "", _
                           Optional ByVal mail_notification As String = "", _
                           Optional ByVal mulst_change_password As Boolean = False, _
                           Optional ByVal generate_password As Boolean = False, _
                           Optional ByVal send_information As Boolean = False) As Boolean
    Dim json As String
    json = user.generateJson(password, auth_source_id, mail_notification, must_change_password, generate_password, send_information)
    CreateUser = SendApiRequest("POST", "users.json", requestBody:=json)
End Function

'Get user information
Public Function GetUsers(Optional ByVal userId As Integer = -1, _
                         Optional ByVal incMemberships As Boolean = True, _
                         Optional ByVal incGroups As Boolean = True, _
                         Optional ByVal status As Integer = 1, _
                         Optional ByVal name As String = "", _
                         Optional ByVal group_id As Integer = -1, _
                         Optional ByVal numPage As Integer = -1, _
                         Optional ByVal limit As Integer = -1 _
                        ) As RedmineUsers
    Dim ret As RedmineUsers
    Dim reqPath, incStr As String
    Dim params As Object
    Dim bIsDone As Boolean
    Dim lOffset As Long
    
    reqPath = "users"
    If userId >= 0 Then
        reqPath = reqPath & "/" & userId
    End If

    incStr = ""
    If incMemberships Then incStr = incStr & ",memberships"
    If incGroups Then incStr = incStr & ",groups"
    Set ret = New RedmineUsers
  
    If limit = -1 Then
        limit = 500 ' Max number by request
    End If
    lOffset = 0
    
    Do
        Set params = CreateObject("Scripting.Dictionary")
        If incStr <> "" Then Call params.Add("include", Mid(incStr, 2))
        If status <> 1 Then Call params.Add("status", status)
        If name <> "" Then Call params.Add("name", name)
        If group_id <> -1 Then Call params.Add("group_id", group_id)
        If lOffset <> 0 Then
            Call params.Add("offset", lOffset)
        ElseIf numPage <> -1 Then
            Call params.Add("offset", numPage)
        End If
        If limit <> -1 Then
            Call params.Add("limit", limit)
        End If
    
        bIsDone = True
        If SendApiRequest("GET", reqPath & ".json", requestParam:=params) Then
            Dim r As RedmineUser
        
            If json.Exists("user") Then
                'Single result
                Set r = Nothing
                Set r = New RedmineUser
                Call r.load(json("user"))
                Call ret.users.Add(r)
            Else
                'Multiple result
                Dim u As Object
                For Each u In json("users")
                    Set r = Nothing
                    Set r = New RedmineUser
                    Call r.load(u)
                    Call ret.users.Add(r)
                    If ret.users.Count = LimitRequest Then
                        Exit For
                    End If
                Next
            End If
            
            If json.Exists("total_count") And ret.usersNumber = -1 Then
                ret.usersNumber = CLng(json("total_count"))
            End If
            If json.Exists("limit") And ret.limit = -1 Then
                ret.limit = CLng(json("limit"))
            End If
            If lOffset = 0 And json.Exists("offset") And ret.numPage = -1 Then
                ret.numPage = CLng(json("offset"))
            End If
            
            ' If numPage = -1, read only this page, not the others
            If numPage = -1 And json.Exists("total_count") And json.Exists("offset") And json.Exists("limit") Then
                lOffset = CLng(json("limit")) + lOffset
                If lOffset < CLng(json("total_count")) And (LimitRequest = -1 Or ret.users.Count < LimitRequest) Then
                    ' Continue the read
                    bIsDone = False
                End If
            End If
            
            Set GetUsers = ret
        Else
            Set GetUsers = Nothing
        End If
    Loop While Not bIsDone
End Function


' Get All groups
Public Function GetGroups(Optional ByVal groupId As Integer = -1, _
                          Optional ByVal incUsers As Boolean = False, _
                          Optional ByVal incMemberships As Boolean = False) As Collection
    Dim reqPath, incStr As String
    Dim params As Object

    reqPath = "groups"
    If groupId >= 0 Then
        reqPath = reqPath & "/" & groupId
    End If

    incStr = ""
    If incMemberships Then incStr = incStr & ",memberships"
    If incUsers Then incStr = incStr & ",users"

    Set params = CreateObject("Scripting.Dictionary")
    If incStr <> "" Then Call params.Add("include", Mid(incStr, 2))
  
    If SendApiRequest("GET", reqPath & ".json", requestParam:=params) Then
        Dim ret As Collection
        Dim r As RedmineGroup
        Set ret = New Collection

        If json.Exists("group") Then
            'Single result
            Set r = Nothing
            Set r = New RedmineGroup
            Call r.load(json("group"))
            Call ret.Add(r)
        Else
            'Multiple result
            Dim u As Object
            For Each u In json("groups")
                Set r = Nothing
                Set r = New RedmineGroup
                Call r.load(u)
                Call ret.Add(r)
            Next
        End If
        Set GetGroups = ret
    Else
        Set GetGroups = Nothing
    End If
End Function

'Get projects informations.
'
'@param projectId: Project ID
'@param incTrackers: Also returns a list of project trackers
'@param incCategories: Also returns a list of project categories
'@param incModules: Also returns a list of modules in the project
'@param incActivities: also returns the most recent activity of the project
'@return: A collection of instance of RedmineProject, Nothing if not found.
Public Function GetProjects(Optional ByVal projectId As String = "", _
                            Optional ByVal incTrackers As Boolean = False, _
                            Optional ByVal incCategories As Boolean = False, _
                            Optional ByVal incModules As Boolean = False, _
                            Optional ByVal incActivities As Boolean = False) As Collection
    Dim reqPath, incStr As String
    Dim params As Object
    Dim bIsDone As Boolean
    Dim lOffset As Long
    Dim ret As Collection

    reqPath = "projects"
    lOffset = 0
    If projectId <> "" Then
        reqPath = reqPath & "/" & projectId
    End If
  
    If incTrackers Then incStr = incStr & ",trackers"
    If incCategories Then incStr = incStr & ",issue_categories"
    If incModules Then incStr = incStr & ",enabled_modules"
    If incActivities Then incStr = incStr & ",time_entry_activities"
    Set ret = New Collection
  
    If projects Is Nothing Or incStr <> "" Then
        Do
            bIsDone = True
            If incStr <> "" And lOffset = 0 Then
                Set params = Nothing
            Else
                Set params = CreateObject("Scripting.Dictionary")
                If incStr <> "" Then
                    params.Add "include", Mid(incStr, 2)
                End If
                If lOffset <> 0 Then
                    params.Add "offset", lOffset
                End If
            End If
    
            If SendApiRequest("GET", reqPath & ".json", requestParam:=params) Then
                Dim r As RedmineProject

                If json.Exists("project") Then
                    'Single result
                    Set r = Nothing
                    Set r = New RedmineProject
                    Call r.load(json("project"))
                    Call ret.Add(r)
                Else
                    'Multiple result
                    Dim u As Object
            
                    For Each u In json("projects")
                        Set r = Nothing
                        Set r = New RedmineProject
                        Call r.load(u)
                        Call ret.Add(r)
                    Next
                    If json.Exists("total_count") And json.Exists("offset") And json.Exists("limit") Then
                        lOffset = CLng(json("limit")) + lOffset
                        If lOffset < CLng(json("total_count")) Then
                            ' Continue the read
                            bIsDone = False
                        End If
                    End If
                End If
                ' Assign results
                Set projects = ret
                Set GetProjects = ret
            Else
                ' No result
                Set GetProjects = Nothing
            End If
        Loop While Not bIsDone
    Else
        Set GetProjects = projects
    End If
End Function

'Get project informations by project name.
'
' No request is done if the internal list of projects has been already done.
'
'@param projectName: Project's name
'@return: An instance of RedmineProject, Nothing if not found.
Public Function GetProjectByName(ByVal projectName As String) As RedmineProject
    Dim project As RedmineProject
  
    Set GetProjectByName = Nothing
    For Each project In GetProjects()
        If project.name = projectName Then
            Set GetProjectByName = project
            Exit For
        End If
    Next
End Function

'Get custom fields informations from a specific url.
'
' This can be used if specific URL is used to retrieve custom fields (without having administrator privileges)
'
'@return: A collection of instance of RedmineCustomField, Nothing if not found.
Public Function GetCustomFields() As Collection
    Set GetCustomFields = GetCustomFieldsFromURL("custom_fields.json", "custom_fields")
End Function

'Get custom fields informations from a specific url.
'
' This can be used if specific URL is used to retrieve custom fields (without having administrator privileges)
' This request are used to call specific api created by Redmine plugin (in server side), because
' some fields need administrator authorization so this specific request get some informations without needed the administrator rights
'
'@param reqPath Request path
'@param jsonEntry list of json entrie's names (separate by '/') where to decode custom fields.
'@return: A collection of instance of RedmineCustomField, Nothing if not found.
Public Function GetCustomFieldsFromURL(ByVal reqPath As String, ByVal jsonEntry As String) As Collection
    If customFields Is Nothing Then
        Dim ret As Collection

        Set ret = New Collection
        Set GetCustomFieldsFromURL = Nothing
  
        If SendApiRequest("GET", reqPath) Then
            Dim cfVal As RedmineCustomField
            
            If jsonEntry <> vbNullString Then
                Dim arrSplitStrings() As String
                Dim index As Integer
                
                arrSplitStrings = Split(jsonEntry, "/")
                
                For index = 0 To UBound(arrSplitStrings) - 1
                    If json.Exists(arrSplitStrings(index)) Then
                        Set json = json(arrSplitStrings(index))
                    Else
                        Set json = Nothing
                        Exit For
                    End If
                Next index
                jsonEntry = arrSplitStrings(UBound(arrSplitStrings))
            End If

            If Not json Is Nothing Then
                If json.Exists(jsonEntry) Then
                    Dim cf As Object
        
                    For Each cf In json(jsonEntry)
                        Set cfVal = New RedmineCustomField
                        Call cfVal.load(cf)
                        Call ret.Add(cfVal)
                    Next
                  
                    Set GetCustomFieldsFromURL = ret
                    Set customFields = ret
                End If
            End If
        End If
    Else
        Set GetCustomFieldsFromURL = customFields
    End If
End Function

'Get custom field informations by name.
'
' No request is done if the internal list of custom field has been already done.
'
'@param customFieldName: Custom Field's name
'@return: An instance of RedmineCustomField, Nothing if not found.
Public Function GetCustomFieldByName(ByVal customFieldName As String) As RedmineCustomField
    Dim customfield As RedmineCustomField
  
    Set GetCustomFieldByName = Nothing
    For Each customfield In GetCustomFields
        If customfield.name = customFieldName Then
            Set GetCustomFieldByName = customfield
            Exit For
        End If
    Next
End Function

'Get custom status informations.
'
'@return: A collection of instance of RedmineStatus, Nothing if not found.
Public Function GetStatuses() As Collection
    If statuses Is Nothing Then
        Dim reqPath As String
        Dim ret As Collection

        reqPath = "issue_statuses"
        Set ret = New Collection
  
        If SendApiRequest("GET", reqPath & ".json") Then
            Dim cfVal As RedmineStatus
              
            If json.Exists("issue_statuses") Then
                Dim cf As Object
      
                For Each cf In json("issue_statuses")
                    Set cfVal = New RedmineStatus
                    Call cfVal.load(cf)
                    Call ret.Add(cfVal)
                Next
            End If
            Set GetStatuses = ret
            Set statuses = ret
        Else
            Set GetStatuses = Nothing
        End If
    Else
        Set GetStatuses = statuses
    End If
End Function

'Get status informations by status name.
'
' No request is done if the internal list of status has been already done.
'
'@param statusName: Status's name
'@return: An instance of RedmineStatus, Nothing if not found.
Public Function GetStatusByName(ByVal statusName As String) As RedmineStatus
    Dim status As RedmineStatus
  
    Set GetStatusByName = Nothing
    For Each status In GetStatuses
        If status.name = statusName Then
            Set GetStatusByName = status
            Exit For
        End If
    Next
End Function

Public Function GetIssues(ByVal project As RedmineProject, _
                          Optional ByVal numPage As Integer = -1, _
                          Optional ByVal limit As Integer = -1, _
                          Optional ByVal filters As RedmineFilters _
                         ) As RedmineIssues
    Dim reqPath, optionsStr As String
    Dim params As Object
    Dim bIsDone As Boolean
    Dim lOffset As Long
    Dim ret As RedmineIssues

    lOffset = 0
 
    Set ret = New RedmineIssues
    
    ' Filter by project
    If project Is Nothing Then
        reqPath = "issues"
    Else
        reqPath = "projects/" & project.identifier & "/issues"
    End If
    
    If limit = -1 Then
        limit = 500 ' Max number by request
    End If
  
    Do
        bIsDone = True
        If filters Is Nothing And numPage = -1 And limit = -1 Then
            Set params = Nothing
        Else
            If filters Is Nothing Then
                Set params = CreateObject("Scripting.Dictionary")
            Else
                Set params = filters.GetFilters()
            End If
            If lOffset <> 0 Then
                Call params.Add("offset", lOffset)
            ElseIf numPage <> -1 Then
                Call params.Add("offset", numPage)
            End If
            If limit <> -1 Then
                Call params.Add("limit", limit)
            End If
        End If
    
        If SendApiRequest("GET", reqPath & ".json", requestParam:=params) Then
            Dim r As RedmineIssue
            
            If json.Exists("issue") Then
                'Single result
                Set r = Nothing
                Set r = New RedmineIssue
                Call r.load(json("issue"))
                ret.issuesNumber = 1
                ret.numPage = 0
                Call ret.issues.Add(r)
            Else
                'Multiple result
                Dim u As Object
        
                For Each u In json("issues")
                    Set r = Nothing
                    Set r = New RedmineIssue
                    Call r.load(u)
                    Call ret.issues.Add(r)
                    If ret.issues.Count = LimitRequest Then
                        Exit For
                    End If
                Next
                
                If json.Exists("total_count") And ret.issuesNumber = -1 Then
                    ret.issuesNumber = CLng(json("total_count"))
                End If
                If json.Exists("limit") And ret.limit = -1 Then
                    ret.limit = CLng(json("limit"))
                End If
                If lOffset = 0 And json.Exists("offset") And ret.numPage = -1 Then
                    ret.numPage = CLng(json("offset"))
                End If
                
                ' If numPage = -1, read only this page, not the others
                If numPage = -1 And json.Exists("total_count") And json.Exists("offset") And json.Exists("limit") Then
                    lOffset = CLng(json("limit")) + lOffset
                    If lOffset < CLng(json("total_count")) And (LimitRequest = -1 Or ret.issues.Count < LimitRequest) Then
                        ' Continue the read
                        bIsDone = False
                        ret.numPage = -1 ' More than one page, don't fill this info, it has no sens
                    End If
                End If
            End If
            Set GetIssues = ret
        Else
            Set GetIssues = Nothing
        End If
    Loop While Not bIsDone
End Function

'Send API request to Redmine
'@param method: methods (GET, POST, PUT, DELETE)
'@param requestPath: Request destination resource
'@param requestBody: Request body
'@return: Success or failure of request
Public Function SendApiRequest(ByVal method As String, _
                               ByVal requestPath As String, _
                               Optional ByVal requestParam As Object = Nothing, _
                               Optional ByVal requestBody As String = "") As Boolean
    Const HTTPREQUEST_PROXYSETTING_PROXY = 2
    Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
    Dim key As Variant
    Dim postUri As String
    Dim req As Object
  
    SendApiRequest = False
  
    If BaseUri = vbNullString Or ApiKey = vbNullString Then
        'No required parameters
        Exit Function
    End If
  
    'Create HttpRequest object
    Set req = CreateHttpRequest()
  
    'URL assembly
    postUri = Me.BaseUri & requestPath
    'Add parameters
    If Not requestParam Is Nothing And TypeName(requestParam) = "Dictionary" Then
        If LBound(requestParam.Keys) >= 0 Then
            postUri = postUri & "?"
            For Each key In requestParam
                If TypeName(requestParam.item(key)) = "Collection" Then
                    Dim item As Variant
                    
                    For Each item In requestParam.item(key)
                        postUri = postUri & UrlEncodeUtf8(key) & "=" & UrlEncodeUtf8(item) & "&"
                    Next
                Else
                    postUri = postUri & UrlEncodeUtf8(key) & "=" & UrlEncodeUtf8(requestParam.item(key)) & "&"
                End If
            Next
        End If
    End If

    'Send request
    Call req.Open(method, postUri, False)
    If Proxy <> vbNullString Then
        Call req.SetProxy(HTTPREQUEST_PROXYSETTING_PROXY, Proxy)
        If ProxyLogin <> vbNullString Then
            Call req.SetCredentials(ProxyLogin, ProxyPassword, HTTPREQUEST_SETCREDENTIALS_FOR_PROXY)
        End If
    End If
    
    Call req.setRequestHeader("Content-Type", "application/json")
    Call req.setRequestHeader("X-Redmine-API-Key", Me.ApiKey)
    Call req.send(requestBody)
  
    If Not req.status < 300 Then
        Debug.Print "StatusCode: " & req.status
        Debug.Print "Response: " & req.responseText
        If req.status = 403 Then
            ' Access denied
            Call Err.Raise(Number:=vbObjectError + 1, description:="Request '" & Me.BaseUri & requestPath & "': access denied, may be need specific authorization?")
        End If
        Exit Function
    End If
  
    'Response perspective
    Set json = Nothing
    Set json = JsonConverter.ParseJson(req.responseText)
  
    SendApiRequest = True
End Function

'Text character code conversion
'@param text: Byte string to be converted
'@param fromCharset: Character code before conversion
'@param toCharset: Character code of conversion destination (default unicode)
'@return: Converted string
Private Function ConvTextEncoding(ByVal text, _
                                  ByVal fromCharset As String, _
                                  Optional ByVal toCharset As String = "unicode") As String
    Dim convText As String
    With CreateObject("ADODB.Stream")
        .Open
        .Type = adTypeText
        .Charset = toCharset
        .WriteText text
        .Position = 0
        .Type = adTypeText
        .Charset = fromCharset
    
        On Error GoTo myLabel
        convText = .ReadText()
        ConvTextEncoding = Mid(convText, 3, Len(convText))

        On Error GoTo 0
    End With
Exit Function

myLabel:
    ConvTextEncoding = StrConv(text, vbUnicode, 1041)
End Function

Private Function SaveAsFile(ByVal datas, ByVal path As String) As Boolean
    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        .Write datas
        .SaveToFile path, adSaveCreateOverWrite
        .Close
    End With
End Function

Private Function CreateHttpRequest() As Object
    'Create WinHttpRequest / XMLHTTPRequest object
    Dim progIDs As Variant
    Dim ret As Object
    Dim I As Long
   
    Set ret = Nothing ' Initialize
    progIDs = Array("WinHttp.WinHttpRequest.5.1", _
                    "WinHttp.WinHttpRequest.5", _
                    "WinHttp.WinHttpRequest", _
                    "Msxml2.ServerXMLHTTP.6.0", _
                    "Msxml2.ServerXMLHTTP.5.0", _
                    "Msxml2.ServerXMLHTTP.4.0", _
                    "Msxml2.ServerXMLHTTP.3.0", _
                    "Msxml2.ServerXMLHTTP", _
                    "Microsoft.ServerXMLHTTP", _
                    "Msxml2.XMLHTTP.6.0", _
                    "Msxml2.XMLHTTP.5.0", _
                    "Msxml2.XMLHTTP.4.0", _
                    "Msxml2.XMLHTTP.3.0", _
                    "Msxml2.XMLHTTP", _
                    "Microsoft.XMLHTTP")
    On Error Resume Next
    For I = LBound(progIDs) To UBound(progIDs)
        Set ret = CreateObject(progIDs(I))
        If Not ret Is Nothing Then Exit For
    Next
    On Error GoTo 0
    Set CreateHttpRequest = ret
End Function

'Return a URL-encoded string
Private Function UrlEncodeUtf8(ByVal str As String) As String
    UrlEncodeUtf8 = URLEncode(str)
End Function

