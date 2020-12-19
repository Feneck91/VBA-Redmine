VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public name, identifier, description, homepage As String
Public is_public As Boolean

'Constructor
Private Sub Class_Initialize()
End Sub

'Destructor
Private Sub Class_Terminate()
End Sub

' Because sometime
Public Function GetDescriptionWithoutImage() As String
    Dim arrSplitStrings() As String
    
    arrSplitStrings = Split(description, "*")

    If UBound(arrSplitStrings, 1) >= 1 Then
        GetDescriptionWithoutImage = arrSplitStrings(1)
    Else
        GetDescriptionWithoutImage = description
    End If
End Function

Public Function load(ByVal json As Dictionary)
    Me.id = CLng(json("id"))
    Me.name = json("name")
    Me.identifier = json("identifier")
    Me.description = json("description")
    Me.homepage = json("homepage")
    Me.is_public = json("is_public")
End Function

