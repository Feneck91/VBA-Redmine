VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedminePossibleValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public value, label As String

'Constructor
Private Sub Class_Initialize()
End Sub

'Destructor
Private Sub Class_Terminate()
End Sub

' Load a dictionnary of possible values
Public Function load(ByVal json As Dictionary)
  Me.value = json("value")
  Me.label = json("label")
End Function

' Set value and label with same value
Public Function setValueAndLabel(ByVal value As String)
  Me.value = value
  Me.label = value
End Function

