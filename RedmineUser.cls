VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id, status As Long

' Used when retrieve user from groups request, in this case only id is filled, the others are not filled.
Public name As String

' Used when retrieve users only
Public login, firstname, lastname, mail As String

Private pMemberships, pGroups As Collection

' Get memberships
Public Property Get memberships() As Collection
  Set memberships = pMemberships
End Property

' Get groups
Public Property Get groups() As Collection
  Set groups = pGroups
End Property

'Constructor
Private Sub Class_Initialize()
    Set pMemberships = New Collection
    Set pGroups = New Collection
End Sub

'Destructor
Private Sub Class_Terminate()
    Set pMemberships = Nothing
    Set pGroups = Nothing
End Sub

Public Function load(ByVal json As Dictionary)
    Me.id = CLng(json("id"))
    
    ' Only when retrieve groups with users
    If json.Exists("name") Then
        Me.name = json("name")
    End If
    If json.Exists("firstname") Then
        Me.firstname = json("firstname")
    End If
    If json.Exists("lastname") Then
        Me.lastname = json("lastname")
    End If
    If json.Exists("mail") Then
        Me.mail = json("mail")
    End If
    If json.Exists("status") Then
        Me.status = json("status")
    End If
  
    Dim val As Object
    If json.Exists("memberships") Then
        Dim p As RedmineProject
        For Each val In json("memberships")
            Set p = New RedmineProject
            p.id = CLng(val("project")("id"))
            p.name = val("project")("name")
            Call pMemberships.Add(p)
        Next
    End If
  
    If json.Exists("groups") Then
        Dim g As RedmineGroup
        For Each val In json("groups")
            Set g = New RedmineGroup
            Call g.load(val)
            Call pGroups.Add(g)
        Next
    End If
End Function

Public Function generateJson(Optional ByVal password As String = "", _
                             Optional ByVal auth_source_id As String = "", _
                             Optional ByVal mail_notification As String = "", _
                             Optional ByVal must_change_password As Boolean = False, _
                             Optional ByVal generate_password As Boolean = False, _
                             Optional ByVal send_information As Boolean = False) As String
    generateJson = ""
    Dim root, user As Dictionary
  
    Set root = New Dictionary
    Set user = New Dictionary
  
    Call user.Add("login", Me.login)
    Call user.Add("firstname", Me.firstname)
    Call user.Add("lastname", Me.lastname)
    Call user.Add("mail", Me.mail)
  
    If password <> "" Then Call user.Add("password", password)
    If auth_source_id <> "" Then Call user.Add("auth_source_id", auth_source_id)
    If mail_notification <> "" Then Call user.Add("mail_notification", mail_notification)
    If must_change_password Then Call user.Add("must_change_password", must_change_password)
    If generate_password Then Call user.Add("generate_password", generate_password)
  
    Call root.Add("user", user)
    If send_information Then Call root.Add("send_information", send_information)
  
    generateJson = JsonConverter.ConvertToJson(root)
End Function

