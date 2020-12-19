VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private filters As Dictionary

'Constructor
Private Sub Class_Initialize()
    Set filters = CreateObject("Scripting.Dictionary")
End Sub

'Destructor
Private Sub Class_Terminate()
    Set filters = Nothing
End Sub

' Get filters
Public Function GetFilters() As Dictionary
    Set GetFilters = CloneDictionary(filters)
End Function

Private Sub Addfilter(ByVal key As String, _
                      ByVal value As String _
                     )
    If filters.Exists(key) Then
        If TypeName(filters("key")) <> "Collection" Then
            ' Remove value from dictionary and add it into a list of values
            Dim lst As Collection
            Set lst = New Collection
            Call lst.Add(filters(key))
            Call filters.Remove(key)
            Call filters.Add(key, lst)
        End If
        Call filters(key).Add(value)
    Else
        Call filters.Add(key, value)
    End If
End Sub

' Set the filter status
'
' By default, all opened issues or retrieve from the server if not filter are set.
'
' @param operator:
'        'all' or '*' to take all, in this case values is not needed
'        'opened' or 'o' to take only opened issues, in this case values is not needed
'        'closed' or 'c' to take only closed issues, in this case values is not needed
'        '=' to get a list of status id to retrieve, in this case values must be filled with at least one correct value
'        '!=' to get a list of status id to ignore, in this case values must be filled with at least one correct value
' @param values is a collection of RedmineStatus to get / remove depending of the filter, some values could be Nothing (ignored) but at least one must be correctly filled.
' @throw Exception if parameters are not valid
Public Sub SetFilterStatus(ByVal operator As String, Optional ByVal values As Collection)
    Dim val As RedmineStatus
    
    If operator = "all" Or operator = "*" Then
        ' Get All without status filter
        Call Addfilter("f[]", "status_id")
        Call Addfilter("op[status_id]", "*")
    ElseIf operator = "opened" Or operator = "o" Then
        ' Get All without status filter
        Call Addfilter("f[]", "status_id")
        Call Addfilter("op[status_id]", "o")
    ElseIf operator = "closed" Or operator = "c" Then
        ' Get All without status filter
        Call Addfilter("f[]", "status_id")
        Call Addfilter("op[status_id]", "c")
    ElseIf operator = "=" Or operator = "!=" Then
        Dim op As String
        
        op = operator
        If op = "!=" Then
            op = "!"
        End If
        Call Addfilter("f[]", "status_id")
        Call Addfilter("op[status_id]", op)
        
        If values Is Nothing Then
            ' Unsupported operator
            Call Err.Raise(Number:=vbObjectError + 1, description:="Need collection of values in filter status with '" & operator & "' operator!")
        Else
            Dim iCountAdded As Integer
            iCountAdded = 0
            If values.Count > 0 Then
                For Each val In values
                    ' Add all values
                    If Not val Is Nothing Then
                        Call Addfilter("v[status_id][]", val.id)
                        iCountAdded = iCountAdded + 1
                    End If
                Next
                If iCountAdded = 0 Then
                    Call Err.Raise(Number:=vbObjectError + 1, description:="Need at least one value in collection of values in filter status with '" & operator & "' operator!")
                End If
            Else
                Call Err.Raise(Number:=vbObjectError + 1, description:="Need not empty collection of values in filter status with '" & operator & "' operator!")
            End If
        End If
    Else
        ' Unsupported operator
        Call Err.Raise(Number:=vbObjectError + 1, description:="Unknown '" & operator & "' operator in filter status!")
    End If
End Sub





