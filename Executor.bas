Attribute VB_Name = "Executor"
Private Sub FillTab(ByRef sheet As Worksheet, ByRef usersInformations() As String, ByVal title As String, Optional ByVal numberItems As Long = -1)
    Dim rangeValues, rangeHeader, rangeTitle As range
    Dim oTableStyle As TableStyle
    Dim widthCells As Integer
    
    Const INDEX_COLUMN_START = 1
    Const INDEX_ROW_TITLE = 1
    Const INDEX_ROW_LIMIT_MESSAGE = 2
    Const INDEX_ROW_TABLE = 4
    
    widthCells = -1
    On Error Resume Next
    widthCells = UBound(usersInformations, 2)
    On Error GoTo 0
    
    Set rangeTitle = sheet.range(sheet.Cells(INDEX_ROW_TITLE, INDEX_COLUMN_START), sheet.Cells(INDEX_ROW_TITLE, INDEX_COLUMN_START + IIf(widthCells = -1, 10, widthCells)))
    
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
        If numberItems <> -1 And numberItems <> UBound(usersInformations, 1) Then
            ' All item has not been retrieve from server du to RedmineAPI.LimitRequest limit
            Dim rangeLimit
            Set rangeLimit = sheet.range(sheet.Cells(INDEX_ROW_LIMIT_MESSAGE, INDEX_COLUMN_START), sheet.Cells(INDEX_ROW_LIMIT_MESSAGE, INDEX_COLUMN_START + IIf(widthCells = -1, 10, widthCells)))
            
            ' Assign Limit
            sheet.Cells(2, INDEX_COLUMN_START).value = "Be careful, there is only " & UBound(usersInformations, 1) & " items displayed, database contains " & numberItems & " items!"
            rangeLimit.Merge
            rangeLimit.Cells.VerticalAlignment = xlCenter
            rangeLimit.Cells.HorizontalAlignment = xlCenter
            rangeLimit.Characters.Font.Bold = True
            rangeLimit.Interior.Color = RGB(255, 174, 0)
            rangeLimit.Characters.Font.Color = RGB(255, 0, 0)
        End If
    
        ' Compute array range
        Set rangeValues = sheet.range(sheet.Cells(INDEX_ROW_TABLE, INDEX_COLUMN_START), sheet.Cells(INDEX_ROW_TABLE + UBound(usersInformations, 1), INDEX_COLUMN_START + widthCells))
        Set rangeHeader = sheet.range(sheet.Cells(INDEX_ROW_TABLE, INDEX_COLUMN_START), sheet.Cells(INDEX_ROW_TABLE, INDEX_COLUMN_START + widthCells))
        
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
    Dim customfield As RedmineCustomField
    Dim arrSplitStrings() As String
    Dim user As RedmineUser
    Dim index As Integer
    Dim sheetName As String
    Dim rowNumber As Integer
    Dim collectionCustomField As Collection
    Const ROW_BEGIN_GENERATION = 10
    
    Const ColumnCheck = "H"
    Const ColumnIssueDisplayDescription = "I"
    Const ColumnIssueDisplayCustomFields = "D"
    Const ColumnIssueProjectFilter = "E"
    Const ColumnIssueStatusFilter = "F"
    Const ColumnIssueCustomFieldFilter = "G"
    
    
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
    If ThisWorkbook.Sheets("Main").Cells(7, "B").value <> vbNullString Then
        redmine.LimitRequest = ThisWorkbook.Sheets("Main").Cells(7, "B").value
    End If
    
    '
    '
    ' Get Users list
    '
    '
    rowNumber = ROW_BEGIN_GENERATION + 0
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim users As RedmineUsers
        
        sheetName = "Users"
        Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetName
        Call Worksheets("Main").Activate
        
        Set users = redmine.GetUsers(numPage:=20)
        
        Dim usersInformations() As String
        
        ReDim usersInformations(users.users.Count, 4)
        usersInformations(0, 0) = "Id"
        usersInformations(0, 1) = "Login"
        usersInformations(0, 2) = "First name"
        usersInformations(0, 3) = "Last Name"
        usersInformations(0, 4) = "Email"
        For index = 1 To users.users.Count
            Set user = users.users.item(index)
            usersInformations(index, 0) = user.id
            usersInformations(index, 1) = user.login
            usersInformations(index, 2) = user.firstname
            usersInformations(index, 3) = user.lastname
            usersInformations(index, 4) = user.mail
        Next index
        Call FillTab(Worksheets(sheetName), usersInformations, "List of Redmine users, " & users.usersNumber & " found.", users.usersNumber)
    End If
    
    '
    '
    ' Get Projects list
    '
    '
    rowNumber = ROW_BEGIN_GENERATION + 1
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
        Call FillTab(Worksheets(sheetName), projectsInformations, "List of all Redmine projects, " & projects.Count & " found.")
    End If
    
    '
    '
    ' Get Custom Fields list
    '
    '
    rowNumber = ROW_BEGIN_GENERATION + 2
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim customFields As Collection
        
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
            Set customfield = customFields.item(index)
            customFieldsInformations(index, 0) = customfield.id
            customFieldsInformations(index, 1) = customfield.name
            customFieldsInformations(index, 2) = customfield.customized_type
            customFieldsInformations(index, 3) = customfield.field_format
            customFieldsInformations(index, 4) = customfield.default_value
            If Not customfield.possible_values Is Nothing Then
                Dim possibleValue As RedminePossibleValue
                Dim allPossibleValues, allPossibleLabels, separator As String
                
                separator = ""
                allPossibleValues = ""
                allPossibleLabels = ""
                
                For Each possibleValue In customfield.possible_values
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
    ' Get Group list
    '
    '
    rowNumber = ROW_BEGIN_GENERATION + 3
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim groups As Collection
        Dim group As RedmineGroup
        
        sheetName = "Groups"
        Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetName
        Call Worksheets("Main").Activate
        
        Set groups = redmine.GetGroups(-1, True, True)
        
        Dim groupsInformations() As String
        Dim infosText As String
        ReDim groupsInformations(groups.Count, 3)
        
        groupsInformations(0, 0) = "Id"
        groupsInformations(0, 1) = "Name"
        groupsInformations(0, 2) = "Users"
        groupsInformations(0, 3) = "Memberships"
        For index = 1 To groups.Count
            Set group = groups.item(index)
            groupsInformations(index, 0) = group.id
            groupsInformations(index, 1) = group.name
            infosText = vbNullString
            For Each user In group.users()
                infosText = infosText & IIf(infosText = vbNullString, "", vbCrLf) & user.name & " (" & user.id & ")"
                
            Next
            groupsInformations(index, 2) = infosText
            infosText = vbNullString
            For Each project In group.memberships()
                infosText = infosText & IIf(infosText = vbNullString, "", vbCrLf) & project.name
            Next
            groupsInformations(index, 3) = infosText
        Next index
        Call FillTab(Worksheets(sheetName), groupsInformations, "List of all groups, " & groups.Count & " found.")
    End If
    
    '
    '
    ' Get Status list
    '
    '
    rowNumber = ROW_BEGIN_GENERATION + 4
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
        Call FillTab(Worksheets(sheetName), statusedInformations, "List of all statuses, " & statuses.Count & " found.")
    End If
    
    '
    '
    ' Get Issues list
    '
    '
    rowNumber = ROW_BEGIN_GENERATION + 5
    If Worksheets("Main").Cells(rowNumber, ColumnCheck).value = True Then
        Dim collectionStatus As Collection
        Dim filters As RedmineFilters
        Dim issues As RedmineIssues
        Dim issue As RedmineIssue
        Dim valueFilter, valueProjectFilter
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
        End If
       
        If (Not project Is Nothing) Or valueProjectFilter = vbNullString Then
            ' Ok project found
            
            ' Manage filters
            Set filters = New RedmineFilters
            
            ' Check Status Filters
            valueFilter = Worksheets("Main").Cells(rowNumber, ColumnIssueStatusFilter).value
            If valueFilter <> vbEmpty Then
                ' Check Issue Filters
                arrSplitStrings = Split(valueFilter, "|")
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
            
            ' Check Custom Field Filters
            valueFilter = Worksheets("Main").Cells(rowNumber, ColumnIssueCustomFieldFilter).value
            If errorMessage = vbEmpty And valueFilter <> vbEmpty Then
                ' Check Custom Field Filters
                Dim arrSplitLinesStrings() As String
                Dim indexLine As Integer
                
                arrSplitLinesStrings = Split(valueFilter, vbLf)
                For indexLine = 0 To UBound(arrSplitLinesStrings)
                    Dim customFieldText As String
                    
                    arrSplitStrings = Split(Trim(arrSplitLinesStrings(indexLine)), "|")
                    
                    If UBound(arrSplitStrings, 1) = 2 Then
                        customFieldText = Trim(arrSplitStrings(0))
                        Set customfield = redmine.GetCustomFieldByName(customFieldText)
                        If customfield Is Nothing Then
                            errorTitleMessage = "Unknown custom field"
                            errorMessage = "Unknown custom field '" & Trim(customFieldText) & "'!" & vbCrLf & vbCrLf & "Please have a look into column's comment."
                            Exit For
                        Else
                            Set collectionCustomField = New Collection
                        
                            If Trim(arrSplitStrings(0)) <> vbNullString Then
                                Dim arrSplitCustomFieldFilter() As String
                                arrSplitCustomFieldFilter = Split(arrSplitStrings(2), ",")
                                For index = 0 To UBound(arrSplitCustomFieldFilter)
                                    Call collectionCustomField.Add(Trim(arrSplitCustomFieldFilter(index)))
                                Next index
                            End If
                            If errorMessage = vbEmpty Then
                                ' If no error
                                On Error GoTo ManageException:
                                Call filters.SetFilterCustomField(Trim(arrSplitStrings(1)), customfield, collectionCustomField)
                                On Error GoTo 0
                            End If
                        End If
                    Else
                        errorTitleMessage = "Bad custom field format"
                        errorMessage = "Bad format for custom field filter!" & vbCrLf & vbCrLf & "Please have a look into column's comment."
                    End If
                Next indexLine
            End If
            
            ' Request Issues only if no errors!
            If errorMessage = vbEmpty Then
                On Error GoTo ManageException:
                Set issues = redmine.GetIssues(project, -1, -1, filters)
                On Error GoTo 0
                If Not issues Is Nothing Then
                    Dim valueAddFields As String
                    Dim indexLineAddField As Integer
                    
                    Set collectionCustomField = New Collection
                    
                    valueAddFields = Worksheets("Main").Cells(rowNumber, ColumnIssueDisplayCustomFields).value
                    If valueAddFields <> vbNullString Then
                        ' Check Custom Field Filters
                        Dim arrSplitAddFieldsString() As String
                        
                        arrSplitAddFieldsString = Split(valueAddFields, vbLf)
                        For indexLineAddField = 0 To UBound(arrSplitAddFieldsString)
                            Set customfield = redmine.GetCustomFieldByName(Trim(arrSplitAddFieldsString(indexLineAddField)))
                            If customfield Is Nothing Then
                                errorTitleMessage = "Unknown custom field"
                                errorMessage = "Unknown added display custom field '" & Trim(customFieldText) & "'!" & vbCrLf & vbCrLf & "Please have a look into column's comment."
                                Exit For
                            Else
                                Call collectionCustomField.Add(customfield)
                            End If
                        Next indexLineAddField
                    End If
                    
                    If errorMessage = vbEmpty Then
                        ReDim issuesFieldsInformations(issues.issues.Count, 5 + IIf(displayDescription, 1, 0) + collectionCustomField.Count)
                    
                        issuesFieldsInformations(0, 0) = "Id"
                        issuesFieldsInformations(0, 1) = "Subject"
                        issuesFieldsInformations(0, 2) = "Status"
                        issuesFieldsInformations(0, 3) = "Priority"
                        issuesFieldsInformations(0, 4) = "Author"
                        issuesFieldsInformations(0, 5) = "Assign To"
                        If displayDescription Then
                            issuesFieldsInformations(0, 6) = "Description"
                        End If
                        
                        indexLineAddField = 6 + IIf(displayDescription, 1, 0)
                        For Each customfield In collectionCustomField
                            issuesFieldsInformations(0, indexLineAddField) = customfield.name
                            indexLineAddField = indexLineAddField + 1
                        Next
                
                        For index = 1 To issues.issues.Count
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
                            
                            indexLineAddField = 6 + IIf(displayDescription, 1, 0)
                            For Each customfield In collectionCustomField
                                issuesFieldsInformations(index, indexLineAddField) = issue.getCustomFieldValue(customfield)
                                indexLineAddField = indexLineAddField + 1
                            Next
                        Next index
                        Call FillTab(Worksheets(sheetName), issuesFieldsInformations, "List of Issues, " & issues.issuesNumber & " found.", issues.issuesNumber)
                    End If
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

