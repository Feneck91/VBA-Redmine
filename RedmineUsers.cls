VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Number of total user in Remine database with filters (not int values that could be less, depending of users page number)
Public usersNumber As Long
' Current user page
Public numPage As Long
' Current users per page
Public limit As Long
' List of users for this page
Public users As Collection

'Constructor
Private Sub Class_Initialize()
    usersNumber = -1
    numPage = -1
    limit = -1
    Set users = New Collection
End Sub

'Destructor
Private Sub Class_Terminate()
    Set users = Nothing
End Sub

