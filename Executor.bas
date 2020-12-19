Attribute VB_Name = "Executor"
Private Sub FillTab(ByRef sheet As Worksheet, ByRef usersInformations() As String, ByVal title As String)
    Dim rangeValues, rangeHeader, rangeTitle As range
    Dim oTableStyle As TableStyle
    Dim widthCells As Integer
    
    Const INDEX_COLUMN_START = 1
    
    widthCells = -1
    On Error Resume Next
    widthCells = UBound(usersInformations, 2)
    On Error GoTo 0
    
    Set rangeTitle = sheet.range(sheet.Cells(1, INDEX_COLUMN_START), sheet.Cells(1, INDEX_COLUMN_START + IIf(widthCells = -1, 10, widthCells)))
    
    ' Assign title
    sheet.Cells(1, INDEX_COLUMN_START).value = title
    rangeTitle.Merge
    rangeTitle.Cells.VerticalAlignment = xlCenter
    rangeTitle.Cells.HorizontalAlignment = xlCenter
    rangeTitle.Characters.Font.Bold = True
    rangeTitle.Interior.Color = RGB(250, 250, 68)

    If widthCells = -1 Then
        ' Error
        rangeTitle.Characters.Font.Color = RGB(255, 0, 0)
    Else
        ' Compute array range
        Set rangeValues = sheet.range(sheet.Cells(3, INDEX_COLUMN_START), sheet.Cells(3 + UBound(usersInformations, 1), INDEX_COLUMN_START + widthCells))
        Set rangeHeader = sheet.range(sheet.Cells(3, INDEX_COLUMN_START), sheet.Cells(3, INDEX_COLUMN_START + widthCells))
        
        ' Assign values
        rangeValues.value = usersInformations
        
        ' Title no error
        rangeTitle.Characters.Font.Color = RGB(32, 157, 35)
        
        ' Set tab style
        With sheet.ListObjects.Add(xlSrcRange, rangeValues.CurrentRegion, xlYes, xlYes)
            .name = title
            .Comment = title
            .TableStyle = "TableStyleMedium9"
        End With
        
        rangeValues.Borders.LineStyle = xlContinuous
        rangeValues.Borders.Weight = xlHairline
        rangeValues.Borders(xlEdgeTop).LineStyle = xlContinuous
        rangeValues.Borders.Color = RGB(127, 127, 127)
        rangeValues.Borders(xlEdgeTop).Color = RGB(0, 0, 0)
        rangeValues.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        rangeValues.Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
        rangeValues.Borders(xlEdgeRight).Color = RGB(0, 0, 0)
        rangeValues.Borders(xlEdgeTop).Weight = xlMedium
        rangeValues.Borders(xlEdgeBottom).Weight = xlMedium
        rangeValues.Borders(xlEdgeLeft).Weight = xlMedium
        rangeValues.Borders(xlEdgeRight).Weight = xlMedium
        rangeValues.Cells.HorizontalAlignment = xlLeft
        rangeValues.Cells.VerticalAlignment = xlTop
        rangeHeader.Cells.Borders(xlEdgeBottom).Weight = xlMedium
        rangeHeader.Cells.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        rangeHeader.Cells.HorizontalAlignment = xlCenter
        rangeHeader.Cells.VerticalAlignment = xlCenter
    
        ' Auto fit width
        ' rangeValues.WrapText = False
        rangeHeader.EntireColumn.AutoFit
        rangeValues.EntireRow.AutoFit
        
        ' Activate it
        sheet.Activate
    End If
End Sub

Sub OnButtonExecuteClick()
    Dim redmine As RedmineApi
    Dim project As RedmineProject
    Dim status As RedmineStatus
    Dim index As Integer
    Dim sheetName As String
    Dim rowNumber As Integer
    
    Const ColumnCheck = "G"
    Const ColumnIssueDisplayDescription = "H"
    Const ColumnIssueProjectFilter = "D"
    Const ColumnIssueStatusFilter = "E"
    
    
    ' Delete all pages
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Application.DisplayAlerts = False
        If ws.name <> "Main" Then
            Call ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    '
    '
    ' Now make Initialize Redmine
    '
    '
    Set redmine = New RedmineApi
    redmine.BaseUri = ThisWorkbook.Sheets("Main").Cells(4, "B").value
    redmine.ApiKey = ThisWorkbook.Sheets("Main").Cells(6, "B").value
    redmine.Proxy = ThisWorkbook.Sheets("Main").Cells(5, "B").value
    
    '
    '
    ' Get Users list
    '
    '
    rowNumber = 9
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim users As Collection
        Dim user As RedmineUser
        
        sheetName = "Users"
        Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetName
        Call Worksheets("Main").Activate
        
        Set users = redmine.GetUsers()
        
        Dim usersInformations() As String
        
        ReDim usersInformations(users.Count, 4)
        usersInformations(0, 0) = "Id"
        usersInformations(0, 1) = "Login"
        usersInformations(0, 2) = "First name"
        usersInformations(0, 3) = "Last Name"
        usersInformations(0, 4) = "Email"
        For index = 1 To users.Count
            Set user = users.item(index)
            usersInformations(index, 0) = user.id
            usersInformations(index, 1) = user.login
            usersInformations(index, 2) = user.firstname
            usersInformations(index, 3) = user.lastname
            usersInformations(index, 4) = user.mail
        Next index
        Call FillTab(Worksheets(sheetName), usersInformations, "List of all Redmine users ")
    End If
    
    '
    '
    ' Get Projects list
    '
    '
    rowNumber = 10
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim projects As Collection
        
        sheetName = "Projects"
        Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetName
        Call Worksheets("Main").Activate
        
        Set projects = redmine.GetProjects()
        
        Dim projectsInformations() As String
        ReDim projectsInformations(projects.Count, 4)
        
        projectsInformations(0, 0) = "Id"
        projectsInformations(0, 1) = "Name"
        projectsInformations(0, 2) = "Identifier"
        projectsInformations(0, 3) = "Description"
        projectsInformations(0, 4) = "Public"
        For index = 1 To projects.Count
            Set project = projects.item(index)
            projectsInformations(index, 0) = project.id
            projectsInformations(index, 1) = project.name
            projectsInformations(index, 2) = project.identifier
            projectsInformations(index, 3) = project.GetDescriptionWithoutImage()
            projectsInformations(index, 4) = project.is_public
        Next index
        Call FillTab(Worksheets(sheetName), projectsInformations, "List of all Redmine projects")
    End If
    
    '
    '
    ' Get Custom Fields list
    '
    '
    rowNumber = 11
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim customFields As Collection
        Dim customField As RedmineCustomField
        
        sheetName = "Custom Fields"
        Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetName
        Call Worksheets("Main").Activate
        
        Set customFields = redmine.GetCustomFields()
        
        Dim customFieldsInformations() As String
        ReDim customFieldsInformations(customFields.Count, 6)
        
        customFieldsInformations(0, 0) = "Id"
        customFieldsInformations(0, 1) = "Name"
        customFieldsInformations(0, 2) = "Type"
        customFieldsInformations(0, 3) = "Format"
        customFieldsInformations(0, 4) = "Default"
        customFieldsInformations(0, 5) = "Value"
        customFieldsInformations(0, 6) = "Label"
        
        For index = 1 To customFields.Count
            Set customField = customFields.item(index)
            customFieldsInformations(index, 0) = customField.id
            customFieldsInformations(index, 1) = customField.name
            customFieldsInformations(index, 2) = customField.customized_type
            customFieldsInformations(index, 3) = customField.field_format
            customFieldsInformations(index, 4) = customField.default_value
            If Not customField.possible_values Is Nothing Then
                Dim possibleValue As RedminePossibleValue
                Dim allPossibleValues, allPossibleLabels, separator As String
                
                separator = ""
                allPossibleValues = ""
                allPossibleLabels = ""
                
                For Each possibleValue In customField.possible_values
                    allPossibleValues = allPossibleValues & separator & possibleValue.value
                    allPossibleLabels = allPossibleLabels & separator & possibleValue.label
                    separator = vbCrLf
                Next
                customFieldsInformations(index, 5) = allPossibleValues
                customFieldsInformations(index, 6) = allPossibleLabels
            Else
                customFieldsInformations(index, 5) = ""
                customFieldsInformations(index, 6) = ""
            End If
        Next index
        Call FillTab(Worksheets(sheetName), customFieldsInformations, "List of all Custom Fields")
    End If
    
    '
    '
    ' Get Status list
    '
    '
    rowNumber = 12
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim statuses As Collection
        
        sheetName = "Statuses"
        Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetName
        Call Worksheets("Main").Activate
        
        Set statuses = redmine.GetStatuses()
        
        Dim statusedInformations() As String
        ReDim statusedInformations(statuses.Count, 2)
        
        statusedInformations(0, 0) = "Id"
        statusedInformations(0, 1) = "Name"
        statusedInformations(0, 2) = "Is Closes"
        For index = 1 To statuses.Count
            Set status = statuses.item(index)
            statusedInformations(index, 0) = status.id
            statusedInformations(index, 1) = status.name
            statusedInformations(index, 2) = status.is_closed
        Next index
        Call FillTab(Worksheets(sheetName), statusedInformations, "List of all statuses")
    End If
    
    '
    '
    ' Get Issues list
    '
    '
    rowNumber = 13
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim collectionStatus As Collection
        Dim filters As RedmineFilters
        Dim issues As RedmineIssues
        Dim issue As RedmineIssue
        Dim valueStatusFilter, valueProjectFilter
        Dim errorMessage, errorTitleMessage As String
        Dim displayDescription As Boolean
        Dim issuesFieldsInformations() As String
        
        displayDescription = Worksheets("Main").Cells(rowNumber, ColumnIssueDisplayDescription).value
        sheetName = "Issues"
        Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetName
        Call Worksheets("Main").Activate
        
        valueProjectFilter = Worksheets("Main").Cells(rowNumber, ColumnIssueProjectFilter).value
        
        If valueProjectFilter <> vbNullString Then
            Set project = redmine.GetProjectByName(valueProjectFilter)
        Else
            ' Project not filled
            errorTitleMessage = "No project filled"
            errorMessage = "The project's name is mandatory,  please fill it in " & ColumnIssueProjectFilter & rowNumber & " cell!"
        End If
       
        If Not project Is Nothing Then
            ' Ok project found
            
            ' Manage filters
            Set filters = New RedmineFilters
            valueStatusFilter = Worksheets("Main").Cells(rowNumber, ColumnIssueStatusFilter).value
            If valueStatusFilter <> vbEmpty Then
                ' Check Issue Filters
                Dim arrSplitStrings() As String
                arrSplitStrings = Split(valueStatusFilter, "|")
                If UBound(arrSplitStrings, 1) = 0 Or UBound(arrSplitStrings, 1) = 1 Then
                    If UBound(arrSplitStrings, 1) = 0 Then
                        ' Nothing to do here but cannot put And between 2 conditions, if first is false, the second is evaluate too and crash!
                    ElseIf Trim(arrSplitStrings(1)) <> vbNullString Then
                        Dim arrSplitStatusFilter() As String
                        
                        Set collectionStatus = New Collection
                        arrSplitStatusFilter = Split(arrSplitStrings(1), ",")
                        For index = 0 To UBound(arrSplitStatusFilter)
                            Dim statusText As String

                            statusText = Trim(arrSplitStatusFilter(index))
                            Set status = redmine.GetStatusByName(statusText)
                            If status Is Nothing Then
                                errorTitleMessage = "Unknown status"
                                errorMessage = "Unknown status '" & Trim(statusText) & "'!" & vbCrLf & vbCrLf & "Please have a look into column's comment."
                                Exit For
                            End If
                            Call collectionStatus.Add(status)
                        Next index
                    End If
                    If errorMessage = vbEmpty Then
                        ' If no error
                        On Error GoTo ManageException:
                        Call filters.SetFilterStatus(Trim(arrSplitStrings(0)), collectionStatus)
                        On Error GoTo 0
                    End If
                Else
                    errorTitleMessage = "Bad status format"
                    errorMessage = "Bad format for status filter!" & vbCrLf & vbCrLf & "Please have a look into column's comment."
                End If
            End If
            
            ' Request Issues only if no errors!
            If errorMessage = vbEmpty Then
                On Error GoTo ManageException:
                Set issues = redmine.GetIssues(project, -1, -1, filters)
                On Error GoTo 0
                If Not issues Is Nothing Then
                    ReDim issuesFieldsInformations(issues.issuesNumber, 5 + IIf(displayDescription, 1, 0))
                
                    issuesFieldsInformations(0, 0) = "Id"
                    issuesFieldsInformations(0, 1) = "Subject"
                    issuesFieldsInformations(0, 2) = "Status"
                    issuesFieldsInformations(0, 3) = "Priority"
                    issuesFieldsInformations(0, 4) = "Author"
                    issuesFieldsInformations(0, 5) = "Assign To"
                    If displayDescription Then
                        issuesFieldsInformations(0, 6) = "Description"
                    End If
            
                    For index = 1 To issues.issuesNumber
                        Set issue = issues.issues.item(index)
                        issuesFieldsInformations(index, 0) = issue.id
                        issuesFieldsInformations(index, 1) = issue.subject
                        issuesFieldsInformations(index, 2) = issue.status("name")
                        issuesFieldsInformations(index, 3) = issue.priority("name")
                        issuesFieldsInformations(index, 4) = issue.author("name")
                        issuesFieldsInformations(index, 5) = issue.assignedTo("name")
                        If displayDescription Then
                            issuesFieldsInformations(index, 6) = issue.description
                        End If
                    Next index
                    Call FillTab(Worksheets(sheetName), issuesFieldsInformations, "List of Issues, " & issues.issuesNumber & " found.")
                ElseIf errorMessage = vbEmpty Then
                    ' Cannot retrieve issues ?
                    errorTitleMessage = "Cannot retrieve Issues"
                    errorMessage = "An error occurs, cannot retrieve Issues in Redmine database!"
                End If
            End If
        ElseIf errorMessage = vbEmpty Then
            ' Project not found
            errorTitleMessage = "Bad project's name"
            errorMessage = "Project '" & valueProjectFilter & "' not found in Redmine database!"
        End If
        
        ' If error display error message box
        If errorMessage <> vbEmpty Then
            ReDim issuesFieldsInformations(0)
            Call FillTab(Worksheets(sheetName), issuesFieldsInformations, "Cannot fill this sheet: " & errorMessage)
            Call MsgBox(errorMessage, vbOKOnly + vbCritical, errorTitleMessage)
        End If
    End If
    GoTo EndFunction:
ManageException:
    errorTitleMessage = "Exception!"
    errorMessage = Err.description
    Resume Next  ' Go back
EndFunction:
End Sub

