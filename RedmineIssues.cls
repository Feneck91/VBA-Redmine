VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineIssues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Number of total issues in Remine database with filters (not int values that could be less, depending of issues page number)
Public issuesNumber As Long
' Current issue page
Public numPage As Integer
' List of issues for this page
Public issues As Collection

'Constructor
Private Sub Class_Initialize()
    issuesNumber = -1
    numPage = -1
    Set issues = New Collection
End Sub

'Destructor
Private Sub Class_Terminate()
    Set issues = Nothing
End Sub

