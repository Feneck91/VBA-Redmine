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
        If TypeName(filters(key)) <> "Collection" Then
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
'        '!' to get a list of status id to ignore, in this case values must be filled with at least one correct value
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
    ElseIf operator = "=" Or operator = "!" Then
        Call Addfilter("f[]", "status_id")
        Call Addfilter("op[status_id]", operator)
        
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

' Set the filter custom field
'
' This is used to make possible filters with redmine custom fields.
' The operator are depending of the current type of custom fields.
'
' @param operator:
'       For bool custom fields:
'           '=' Equal, only one value (int = 0 / 1) or boolean is mandatory.
'           '!' Different, only one value (int = 0 / 1) or boolean is mandatory.
'       For string custom fields:
'           '=' Equal, only one value.
'           '!' Different, only one value.
'           '~' Contains, only one value.
'           '!~' Don't contains, only one value.
'           '*' All fields that are not empty, no value needed.
'           '!*' All fields that are empty, no value needed.
'       For text custom fields:
'           '~' Contains, only one value.
'           '!~' Don't contains, only one value.
'           '*' All fields that are not empty, no value needed.
'           '!*' All fields that are empty, no value needed.
'       For int custom fields:
'           '=' Equal, only one integer value is mandatory.
'           '>=' Upper or equal, only one integer value is mandatory.
'           '<=' Lower or equal, only one integer value is mandatory.
'           '><' Between, only two integer values are mandatory.
'           '*' All fields that are not empty, no value needed.
'           '!*' All fields that are empty, no value needed.
'       For date custom fields:
'           '=' Equal, only one date as String or Date value is mandatory.
'           '>=' Upper or equal, only one as String or Date value is mandatory.
'           '<=' Lower or equal, only one as String or Date value is mandatory.
'           '><' Between, only two as String or Date values are mandatory.
'           '*' All fields that are not empty, no value needed.
'           '!*' All fields that are empty, no value needed.
'       For list custom fields:
'           '=' Equal, only one value (into the possible values list).
'           '!' Different, only one value (into the possible values list).
' @param values is a collection of value used by the filter.
' @throw Exception if parameters are not valid
Public Sub SetFilterCustomField(ByVal operator As String, ByVal customfield As RedmineCustomField, Optional ByVal values As Collection)
    Dim cfId As String
    
    cfId = "cf_" & customfield.id
    If customfield.field_format = "bool" Then
        If operator = "=" Or operator = "!" Then
            Call Addfilter("f[]", cfId)
            Call Addfilter("op[" & cfId & "]", operator)

            If Not values Is Nothing Then
                If values.Count = 1 Then
                    If TypeName(values.item(0)) = "Byte" Or TypeName(values.item(0)) = "Integer" Or TypeName(values.item(0)) = "Long" Then
                        Call Addfilter("v[" & cfId & "]", IIf(values.item(0) = 0, "1", "0"))
                    ElseIf TypeName(values.item(0)) = "Boolean" Then
                        Call Addfilter("v[" & cfId & "]", IIf(values.item(0) = True, "1", "0"))
                    Else
                        ' Unsupported values
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have 'values' parameter as Integer or Boolean type!")
                    End If
                Else
                    ' Unsupported number of values
                    Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have one and only one value for 'values' parameter!")
                End If
            Else
                ' Unsupported values = Nothing
                Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have Nothing for 'values' parameter!")
            End If
        Else
            ' Unsupported operator
            Call Err.Raise(Number:=vbObjectError + 1, description:="Unknown '" & operator & "' operator in filter custom field for " & customfield.field_format & " type!")
        End If
    ElseIf customfield.field_format = "string" Or customfield.field_format = "text" Then
        If operator = "~" Or operator = "!~" Or operator = "*" Or operator = "!*" Or (customfield.field_format = "string" And (operator = "=" Or operator = "!")) Then
            Call Addfilter("f[]", cfId)
            Call Addfilter("op[" & cfId & "]", operator)

            If operator = "*" Or operator = "!*" Then
                ' Empty / Not empty
                If Not values Is Nothing Then
                    If values.Count <> 0 Then
                        ' Unsupported operation
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have 'values' parameter!")
                    End If
                End If
            ElseIf Not values Is Nothing Then
                If values.Count = 1 Then
                    If TypeName(values.item(1)) = "String" Then
                        Call Addfilter("v[" & cfId & "][]", values.item(1))
                    Else
                        ' Unsupported values
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have 'values' parameter as String type!")
                    End If
                Else
                    ' Unsupported number of values
                    Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have one and only one value for 'values' parameter!")
                End If
            Else
                ' Unsupported values = Nothing
                Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have Nothing for 'values' parameter!")
            End If
        Else
            ' Unsupported operator
            Call Err.Raise(Number:=vbObjectError + 1, description:="Unknown '" & operator & "' operator in filter custom field for " & customfield.field_format & " type!")
        End If
    ElseIf customfield.field_format = "string" Then
        If operator = "=" Or operator = "!" Or operator = "~" Or operator = "!~" Or operator = "*" Or operator = "!*" Then
            Call Addfilter("f[]", cfId)
            Call Addfilter("op[" & cfId & "]", operator)

            If operator = "*" Or operator = "!*" Then
                ' Empty / Not empty
                If Not values Is Nothing Then
                    If values.Count <> 0 Then
                        ' Unsupported operation
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have 'values' parameter!")
                    End If
                End If
            ElseIf Not values Is Nothing Then
                If values.Count = 1 Then
                    If TypeName(values.item(1)) = "String" Then
                        Call Addfilter("v[" & cfId & "][]", values.item(1))
                    Else
                        ' Unsupported values
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have 'values' parameter as String type!")
                    End If
                Else
                    ' Unsupported number of values
                    Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have one and only one value for 'values' parameter!")
                End If
            Else
                ' Unsupported values = Nothing
                Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have Nothing for 'values' parameter!")
            End If
        Else
            ' Unsupported operator
            Call Err.Raise(Number:=vbObjectError + 1, description:="Unknown '" & operator & "' operator in filter custom field for " & customfield.field_format & " type!")
        End If
    ElseIf customfield.field_format = "int" Then
        If operator = "=" Or operator = ">=" Or operator = "<=" Or operator = "><" Or operator = "*" Or operator = "!*" Then
            Call Addfilter("f[]", cfId)
            Call Addfilter("op[" & cfId & "]", operator)

            If operator = "*" Or operator = "!*" Then
                ' Empty / Not empty
                If Not values Is Nothing Then
                    If values.Count <> 0 Then
                        ' Unsupported operation
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have 'values' parameter!")
                    End If
                End If
            ElseIf Not values Is Nothing Then
                If operator = "=" Or operator = ">=" Or operator = "<=" Then
                    If values.Count = 1 Then
                        If TypeName(values.item(1)) = "String" Or TypeName(values.item(1)) = "Byte" Or TypeName(values.item(1)) = "Integer" Or TypeName(values.item(1)) = "Long" Then
                            Call Addfilter("v[" & cfId & "][]", CLng(values.item(1)))
                        Else
                            ' Unsupported values
                            Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have 'values' parameter as String or Number type!")
                        End If
                    Else
                        ' Unsupported number of values
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have one and only one value for 'values' parameter!")
                    End If
                ElseIf operator = "><" Then
                    If values.Count = 2 Then
                        If (TypeName(values.item(1)) = "String" Or TypeName(values.item(1)) = "Byte" Or TypeName(values.item(1)) = "Integer" Or TypeName(values.item(1)) = "Long") And (TypeName(values.item(2)) = "String" Or TypeName(values.item(2)) = "Byte" Or TypeName(values.item(2)) = "Integer" Or TypeName(values.item(2)) = "Long") Then
                            Call Addfilter("v[" & cfId & "][]", CLng(values.item(1)))
                            Call Addfilter("v[" & cfId & "][]", CLng(values.item(2)))
                        Else
                            ' Unsupported values
                            Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have 'values' parameter as String or Number type!")
                        End If
                    Else
                        ' Unsupported number of values
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have two and only two values for 'values' parameter!")
                    End If
                End If
            Else
                ' Unsupported values = Nothing
                Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have Nothing for 'values' parameter!")
            End If
        Else
            ' Unsupported operator
            Call Err.Raise(Number:=vbObjectError + 1, description:="Unknown '" & operator & "' operator in filter custom field for " & customfield.field_format & " type!")
        End If
    ElseIf customfield.field_format = "list" Then
        If operator = "=" Or operator = "!" Then
            Dim iCountAdded As Integer
            Dim val As Variant
            
            Call Addfilter("f[]", cfId)
            Call Addfilter("op[" & cfId & "]", operator)
            iCountAdded = 0
            If values.Count > 0 Then
                For Each val In values
                    ' Add all values
                    If TypeName(val) = "String" Then
                        Dim possibleValue As RedminePossibleValue
                        
                        Set possibleValue = customfield.getPossibleValueFromLabel(val)
                        If Not possibleValue Is Nothing Then
                            Call Addfilter("v[" & cfId & "][]", possibleValue.value)
                            iCountAdded = iCountAdded + 1
                        Else
                            Call Err.Raise(Number:=vbObjectError + 1, description:="The value '" & val & "' is not found into the possible values of the custom field with '" & operator & "' operator" & " for " & customfield.field_format & " type!")
                        End If
                    End If
                Next
                If iCountAdded = 0 Then
                    Call Err.Raise(Number:=vbObjectError + 1, description:="Need at least one value in collection of values in custom field filter with '" & operator & "' operator for " & customfield.field_format & " type!")
                End If
            Else
                Call Err.Raise(Number:=vbObjectError + 1, description:="Need not empty collection of values in filter custom field with '" & operator & "' operator for " & customfield.field_format & " type!")
            End If
        Else
            ' Unsupported operator
            Call Err.Raise(Number:=vbObjectError + 1, description:="Unknown '" & operator & "' operator in filter custom field for " & customfield.field_format & " type!")
        End If
    ElseIf customfield.field_format = "date" Then
        If operator = "=" Or operator = ">=" Or operator = "<=" Or operator = "><" Or operator = "*" Or operator = "!*" Then
            Call Addfilter("f[]", cfId)
            Call Addfilter("op[" & cfId & "]", operator)

            If operator = "*" Or operator = "!*" Then
                ' Empty / Not empty
                If Not values Is Nothing Then
                    If values.Count <> 0 Then
                        ' Unsupported operation
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have 'values' parameter!")
                    End If
                End If
            ElseIf Not values Is Nothing Then
                If operator = "=" Or operator = ">=" Or operator = "<=" Then
                    If values.Count = 1 Then
                        If TypeName(values.item(1)) = "String" Or TypeName(values.item(1)) = "Date" Then
                            Dim dt As Date
                            
                            dt = CDate(values.item(1))
                            Call Addfilter("v[" & cfId & "][]", Format(dt, "yyyy-MM-dd"))
                        Else
                            ' Unsupported values
                            Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have 'values' parameter as String or Number type!")
                        End If
                    Else
                        ' Unsupported number of values
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have one and only one value for 'values' parameter!")
                    End If
                ElseIf operator = "><" Then
                    If values.Count = 2 Then
                        If (TypeName(values.item(1)) = "String" Or TypeName(values.item(1)) = "Date") And (TypeName(values.item(2)) = "String" Or TypeName(values.item(2)) = "Date") Then
                            Dim dt1, dt2 As Date
                            
                            dt1 = CDate(values.item(1))
                            dt2 = CDate(values.item(2))
                            
                            Call Addfilter("v[" & cfId & "][]", Format(dt1, "yyyy-MM-dd"))
                            Call Addfilter("v[" & cfId & "][]", Format(dt2, "yyyy-MM-dd"))
                        Else
                            ' Unsupported values
                            Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have 'values' parameter as String or Number type!")
                        End If
                    Else
                        ' Unsupported number of values
                        Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must have two and only two values for 'values' parameter!")
                    End If
                End If
            Else
                ' Unsupported values = Nothing
                Call Err.Raise(Number:=vbObjectError + 1, description:="Operator '" & operator & "' for " & customfield.field_format & " custom field must don't have Nothing for 'values' parameter!")
            End If
        Else
            ' Unsupported operator
            Call Err.Raise(Number:=vbObjectError + 1, description:="Unknown '" & operator & "' operator in filter custom field for " & customfield.field_format & " type!")
        End If
    Else
        ' Unsupported type
        Call Err.Raise(Number:=vbObjectError + 1, description:="Field type " & customfield.field_format & " is not yet supported!")
    End If
End Sub
