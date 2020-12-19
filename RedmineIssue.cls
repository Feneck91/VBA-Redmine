VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public subject, description  As String
' Contains 'id' and 'name' values fields (Dictionary)
Public project, tracker, status, priority, author, assignedTo As Dictionary
' Contains 'id' and 'name', 'value' values fields (Dictionary)
Public customFields As Collection
Public done_ratio As Double
Public is_private As Boolean
Public created_on, update_on As Date

'Constructor
Private Sub Class_Initialize()
    Set project = CreateObject("Scripting.Dictionary")
    Set tracker = CreateObject("Scripting.Dictionary")
    Set status = CreateObject("Scripting.Dictionary")
    Set priority = CreateObject("Scripting.Dictionary")
    Set author = CreateObject("Scripting.Dictionary")
    Set assignedTo = CreateObject("Scripting.Dictionary")
    Set customFields = New Collection
End Sub

'Destructor
Private Sub Class_Terminate()
    Set project = Nothing
    Set tracker = Nothing
    Set status = Nothing
    Set priority = Nothing
    Set author = Nothing
    Set assignedTo = Nothing
    Set customFields = Nothing
End Sub

Private Sub FillIdName(ByVal json As Dictionary, _
                       ByRef rValue As Dictionary, _
                       ByVal strKey As String _
                      )
    If json.Exists(strKey) Then
        Call rValue.Add("id", CLng(json(strKey)("id")))
        Call rValue.Add("name", json(strKey)("name"))
    End If
End Sub

Private Function FillIdNameValue(ByVal json As Dictionary) As Dictionary
    Dim d As Dictionary
    Set d = CreateObject("Scripting.Dictionary")
                           
    Call d.Add("id", CLng(json("id")))
    Call d.Add("name", json("name"))
    Call d.Add("value", json("value"))
   
    Set FillIdNameValue = d
End Function

Public Function load(ByVal json As Dictionary)
    Me.id = CLng(json("id"))
    
    Call FillIdName(json, Me.project, "project")
    Call FillIdName(json, Me.tracker, "tracker")
    Call FillIdName(json, Me.status, "status")
    Call FillIdName(json, Me.priority, "priority")
    Call FillIdName(json, Me.author, "author")
    Call FillIdName(json, Me.assignedTo, "assigned_to")
    If json.Exists("subject") Then
        Me.subject = json("subject")
    End If
    If json.Exists("description") Then
        Me.description = json("description")
    End If
    If json.Exists("is_private") Then
        Me.is_private = json("is_private")
    End If
    
    Dim val As Object

    If json.Exists("custom_fields") Then
        For Each val In json("custom_fields")
            Call customFields.Add(FillIdNameValue(val))
        Next
    End If
End Function

