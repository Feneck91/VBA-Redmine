VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public name As String
Public is_closed As Boolean

'Constructor
Private Sub Class_Initialize()
End Sub

'Destructor
Private Sub Class_Terminate()
End Sub

' Load a dictionnary of status values
Public Function load(ByVal json As Dictionary)
  Me.id = CLng(json("id"))
  Me.name = json("name")
  If json.Exists("is_closed") Then
    Me.is_closed = json("is_closed")
  Else
    Me.is_closed = False
  End If
End Function

