VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineCustomField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public name, customized_type, field_format, default_value  As String
Public possible_values As Collection

'Constructor
Private Sub Class_Initialize()
    Set possible_values = Nothing
End Sub

'Destructor
Private Sub Class_Terminate()
    Set possible_values = Nothing
End Sub

' Load a dictionnary of custom fields
Public Function load(ByVal json As Dictionary)
    Me.id = json("id")
    Me.name = json("name")
    Me.customized_type = json("customized_type")
    Me.field_format = json("field_format")
    Me.default_value = json("default_value")
  
    If json.Exists("possible_values") Then
        Dim p As RedminePossibleValue
    
        Set possible_values = New Collection
        
        If TypeName(json("possible_values")) = "Collection" Then
            Dim val As Variant
    
            For Each val In json("possible_values")
                Set p = New RedminePossibleValue
                Call p.setValueAndLabel(val)
                Call possible_values.Add(p)
            Next
        Else
            Dim valObj As Object
    
            For Each valObj In json("possible_values")
                Set p = New RedminePossibleValue
                Call p.load(valObj)
                Call possible_values.Add(p)
            Next
        End If
    End If
End Function

' Used to manually initialize the class.
'
' This can be use into a software to make request without having administrator privilege (needed to get custom fields list).
'
'@param id Identifier, this is this field that is used to make request filter.
'@param name Displayed name of this field
'@param field_format format of this field. Can be 'bool', 'string', 'text', 'date','int'.
'@param customized_type Custum type
'@param default_value Default value
'@param possible_values Collection of RedminePossibleValue, used only when field_format is 'list', else pass Nothing
Public Sub Init(ByVal id As Long, ByVal name As String, ByVal field_format As String, ByVal customized_type As String, ByVal default_value As String, ByVal possible_values As Collection)
    Me.id = id
    Me.name = name
    Me.customized_type = customized_type
    Me.field_format = field_format
    Me.default_value = default_value
    Set Me.possible_values = possible_values
End Sub


' Search label into possible values an return the possible value if found, Nothing else
Public Function getPossibleValueFromLabel(ByVal val As String) As RedminePossibleValue
    Set getPossibleValueFromLabel = Nothing
    
    If Not possible_values Is Nothing Then
        Dim valFor As RedminePossibleValue
        
        For Each valFor In possible_values
            If val = valFor.label Then
                Set getPossibleValueFromLabel = valFor
                Exit For
            End If
        Next
    End If
End Function

' Search value into possible values an return the possible value if found, Nothing else
Public Function getPossibleValueFromValue(ByVal val As String) As RedminePossibleValue
    Set getPossibleValueFromlValue = Nothing
    
    If Not possible_values Is Nothing Then
        Dim valFor As RedminePossibleValue
        
        For Each valFor In possible_values
            If val = valFor.value Then
                Set getPossibleValueFromlValue = valFor
                Exit For
            End If
        Next
    End If
End Function


