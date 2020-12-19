Attribute VB_Name = "Tests"
Sub TestRedmine()
    Dim redmine As RedmineApi
    Dim users As Collection
    Dim user As RedmineUser
    Dim projects As Collection
    Dim project As RedmineProject
    Dim customFields As Collection
    Dim customField As RedmineCustomField
    Dim possibleValue As RedminePossibleValue
    Dim issues As RedmineIssues
    Dim issue As RedmineIssue
    Dim status As RedmineStatus
    Dim strOut As String
    Dim n As Integer
        
    Set redmine = New RedmineApi
    n = FreeFile()

    redmine.BaseUri = "http://example.com/"
    redmine.ApiKey = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    ' redmine.Proxy = "http://my_proxy_url.xxx:8000"
    Open "D:\path where write output\Output.txt" For Output As #n

    Set users = redmine.GetUsers()
    For Each user In users
      strOut = "User = " & user.login & ":" & user.firstname & user.lastname
      Debug.Print strOut
      Print #n, strOut ' write to file
    Next
    
    Debug.Print
    Print #n, "" ' write to file
    Debug.Print "----------------------------------------------"
    Print #n, "----------------------------------------------" ' write to file
    Debug.Print
    Print #n, "" ' write to file
    
    Set projects = redmine.GetProjects ') ' https://geomaps.buzinet.fr/redmine/projects.json?key=327cdb31ad7ad15eb2d0b354ff6094ca1ee2085c&limit=50&offset=0
    For Each project In projects
      strOut = "Project = " & project.name & ":" & project.identifier & "/" & project.id
      Debug.Print strOut
      Print #n, strOut ' write to file
    Next
    
    Debug.Print
    Print #n, "" ' write to file
    Debug.Print "----------------------------------------------"
    Print #n, "----------------------------------------------" ' write to file
    Debug.Print
    Print #n, "" ' write to file
    
    Dim strProjectName As String
    
    strProjectName = "Base des défauts"
    ' strProjectName = "Base des Actions"
    Set project = redmine.GetProjectByName(strProjectName)
    If project Is Nothing Then
      strOut = "Project '" & strProjectName & "' not found!"
      Debug.Print strOut
      Print #n, strOut ' write to file
    Else
      strOut = "Project = " & project.name & ":" & project.identifier & "/" & project.id
      Debug.Print strOut
      Print #n, strOut ' write to file
    End If

    Debug.Print
    Print #n, "" ' write to file
    Debug.Print "----------------------------------------------"
    Print #n, "----------------------------------------------" ' write to file
    Debug.Print
    Print #n, "" ' write to file
    
    For Each customField In redmine.GetCustomFields()
      strOut = "Custom Field (" & customField.name & ") : id=" & customField.id & " / type = " & customField.customized_type & " / format = " & customField.field_format & " / default = '" & customField.default_value & "'"
      Debug.Print strOut
      Print #n, strOut ' write to file
      
      If Not customField.possible_values Is Nothing Then
        For Each possibleValue In customField.possible_values
          strOut = "  - Value = " & possibleValue.value & " / Label = " & possibleValue.label
          Debug.Print strOut
          Print #n, strOut ' write to file
      Next
      End If
    Next
  
    Debug.Print
    Print #n, "" ' write to file
    Debug.Print "----------------------------------------------"
    Print #n, "----------------------------------------------" ' write to file
    Debug.Print
    Print #n, "" ' write to file
  
    For Each status In redmine.GetStatuses()
        strOut = "- Status (" & status.id & ") : " & vbCrLf & _
                 "   - name = " & status.name & vbCrLf & _
                 "   - is_closed = " & status.is_closed & vbCrLf
        Debug.Print strOut
        Print #n, strOut ' write to file
    Next
  
    Debug.Print
    Print #n, "" ' write to file
    Debug.Print "----------------------------------------------"
    Print #n, "----------------------------------------------" ' write to file
    Debug.Print
    Print #n, "" ' write to file
    
    Set issues = redmine.GetIssues(redmine.GetProjectByName("FT Réf. Produit"))
    ' By default, with no filter, only opened issues are retrieve
    strOut = "Number of issues found = " & issues.issuesNumber
    Debug.Print strOut
    Print #n, strOut ' write to file
    
    For Each issue In issues.issues
        strOut = "- Issue (" & issue.id & ") : " & vbCrLf & _
                 "   - Subject = " & issue.subject & vbCrLf & _
                 "   - Status = " & issue.status("name") & vbCrLf & _
                 "   - Priority = " & issue.priority("name") & vbCrLf & _
                 "   - Author = " & issue.author("name") & vbCrLf & _
                 "   - Assigned to = " & issue.assignedTo("name")
                 ' "   - Description = " & issue.description & vbCrLf & _
        Debug.Print strOut
        Print #n, strOut ' write to file
    Next
  
    Debug.Print
    Print #n, "" ' write to file
    Debug.Print "----------------------------------------------"
    Print #n, "----------------------------------------------" ' write to file
    Debug.Print
    Print #n, "" ' write to file
    
    Dim collectionStatus As Collection
    Dim filters As RedmineFilters
    
    Set collectionStatus = New Collection
    Set filters = New RedmineFilters

    Call collectionStatus.Add(redmine.GetStatusByName("Nouveau"))
    Call collectionStatus.Add(redmine.GetStatusByName("Approuvé"))
    Call filters.SetFilterStatus("=", collectionStatus)
    
    Set issues = redmine.GetIssues(redmine.GetProjectByName("FT Réf. Produit"), -1, -1, filters)
    strOut = "Number of issues found = " & issues.issuesNumber
    Debug.Print strOut
    Print #n, strOut ' write to file
    
    For Each issue In issues.issues
        strOut = "- Issue (" & issue.id & ") : " & vbCrLf & _
                 "   - Subject = " & issue.subject & vbCrLf & _
                 "   - Status = " & issue.status("name") & vbCrLf & _
                 "   - Priority = " & issue.priority("name") & vbCrLf & _
                 "   - Author = " & issue.author("name") & vbCrLf & _
                 "   - Assigned to = " & issue.assignedTo("name")
                 ' "   - Description = " & issue.description & vbCrLf & _
        Debug.Print strOut
        Print #n, strOut ' write to file
    Next
    
    Close #n
End Sub
