Option Explicit
Attribute VB_Name = "PivotSlicerCombinedModule"

' ==================================================================================
' CONSTANTS AND TYPE DECLARATIONS
' ==================================================================================

' Layout Constants
Private Const PIVOT_START_ROW As Long = 23
Private Const SLICER_LEFT_OFFSET As Double = 150
Private Const SLICER_GROUP_SPACING As Double = 10
Private Const SLICER_COLUMNS_PER_GROUP As Integer = 3
Private Const GROUP_TOP_ROW As Long = 20
Private Const PIVOT_ROW_SPACING As Long = 2

' Color Constants
Private Const COLOR_M_GROUP As Long = 15921906  ' RGB(242, 220, 219)
Private Const COLOR_Q_GROUP As Long = 14349306  ' RGB(226, 239, 218)
Private Const COLOR_SQ_GROUP As Long = 16244215 ' RGB(222, 235, 247)

' Worksheet Names
Private Const DATA_SHEET_NAME As String = "Tidied Data"
Private Const PIVOT_SHEET_NAME As String = "PivotTable"

' Progress Constants
Private Const SPINNER_CHARS As String = "|/-\"

' Error Constants
Private Const ERR_DATA_SHEET_NOT_FOUND As Long = vbObjectError + 1001
Private Const ERR_NO_DATA_FOUND As Long = vbObjectError + 1002
Private Const ERR_PIVOT_CREATION_FAILED As Long = vbObjectError + 1003

' Type Definitions
Private Type SlicerGroupConfig
    Name As String
    Prefix As String
    Color As Long
    Slicers As Collection
End Type

Private Type OverallProgress
    CurrentStep As Long
    TotalSteps As Long
    StartTime As Double
    Phase As String
    SpinnerIndex As Integer
End Type

Private Enum ProcessPhase
    InitializingData = 1
    CreatingPivotTables = 2
    CreatingSlicers = 3
    OrganizingSlicers = 4
    ConnectingSlicers = 5
    Finalizing = 6
End Enum

Private Enum SlicerGroupType
    MGroup = 1
    QGroup = 2
    SQGroup = 3
End Enum

' ==================================================================================
' MAIN COMBINED ENTRY POINT
' ==================================================================================

Public Sub CreateAndConnectPivotTablesAndSlicers()
    On Error GoTo ErrorHandler
    
    Dim progress As OverallProgress
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim dataRange As Range, pivotCache As PivotCache
    Dim allSlicers As Collection
    
    ' Initialize
    InitializeCombinedProcess progress
    
    ' Phase 1: Initialize and validate data
    UpdateCombinedProgress progress, InitializingData, "Initializing and validating data..."
    Set wsData = GetDataWorksheet()
    Set wsPivot = GetOrCreatePivotWorksheet()
    Set dataRange = GetValidatedDataRange(wsData)
    Set pivotCache = CreatePivotCache(dataRange)
    
    ' Calculate total steps for accurate progress
    CalculateTotalSteps progress, dataRange.Columns.Count
    
    ' Phase 2: Create pivot tables and slicers
    UpdateCombinedProgress progress, CreatingPivotTables, "Creating pivot tables and slicers..."
    Set allSlicers = CreatePivotTablesWithSlicers(wsPivot, pivotCache, dataRange, progress)
    
    ' Phase 3: Organize slicers
    UpdateCombinedProgress progress, OrganizingSlicers, "Organizing slicers into groups..."
    If allSlicers.Count > 0 Then
        OrganizeSlicersByGroups allSlicers, wsPivot, progress
    End If
    
    ' Phase 4: Connect all slicers to all pivot tables
    UpdateCombinedProgress progress, ConnectingSlicers, "Connecting slicers to pivot tables..."
    ConnectAllSlicersToAllPivotTables wsPivot, progress
    
    ' Phase 5: Finalize
    UpdateCombinedProgress progress, Finalizing, "Finalizing..."
    FinalizeCombinedProcess progress, allSlicers.Count
    
    Exit Sub
    
ErrorHandler:
    CleanupCombinedProcess
    HandleCombinedError Err.Number, Err.Description, "CreateAndConnectPivotTablesAndSlicers"
End Sub

' ==================================================================================
' INDIVIDUAL ENTRY POINTS (for flexibility)
' ==================================================================================

Public Sub CreatePivotTablesAndSlicersOnly()
    On Error GoTo ErrorHandler
    
    OptimizeExcelPerformance True
    
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim dataRange As Range, pivotCache As PivotCache
    Dim allSlicers As Collection
    
    Set wsData = GetDataWorksheet()
    Set wsPivot = GetOrCreatePivotWorksheet()
    Set dataRange = GetValidatedDataRange(wsData)
    Set pivotCache = CreatePivotCache(dataRange)
    
    Dim dummyProgress As OverallProgress
    Set allSlicers = CreatePivotTablesWithSlicers(wsPivot, pivotCache, dataRange, dummyProgress)
    
    If allSlicers.Count > 0 Then
        OrganizeSlicersByGroups allSlicers, wsPivot, dummyProgress
    End If
    
    OptimizeExcelPerformance False
    MsgBox "Created " & allSlicers.Count & " pivot tables and slicers!", vbInformation
    Exit Sub
    
ErrorHandler:
    OptimizeExcelPerformance False
    MsgBox "Error creating pivot tables: " & Err.Description, vbCritical
End Sub

Public Sub ConnectSlicersOnly()
    On Error GoTo ErrorHandler
    
    OptimizeExcelPerformance True
    
    Dim wsPivot As Worksheet
    Set wsPivot = GetOrCreatePivotWorksheet()
    
    Dim dummyProgress As OverallProgress
    ConnectAllSlicersToAllPivotTables wsPivot, dummyProgress
    
    OptimizeExcelPerformance False
    MsgBox "Slicer connections completed!", vbInformation
    Exit Sub
    
ErrorHandler:
    OptimizeExcelPerformance False
    MsgBox "Error connecting slicers: " & Err.Description, vbCritical
End Sub

' ==================================================================================
' COMBINED PROCESS MANAGEMENT
' ==================================================================================

Private Sub InitializeCombinedProcess(ByRef progress As OverallProgress)
    OptimizeExcelPerformance True
    progress.StartTime = Timer
    progress.CurrentStep = 0
    progress.SpinnerIndex = 0
    Application.DisplayStatusBar = True
End Sub

Private Sub CalculateTotalSteps(ByRef progress As OverallProgress, columnCount As Long)
    Dim pivotTableCount As Long
    Dim slicerCount As Long
    
    pivotTableCount = columnCount
    slicerCount = columnCount
    
    progress.TotalSteps = 5 + _
                         pivotTableCount + _
                         slicerCount + _
                         slicerCount + _
                         (slicerCount * pivotTableCount) + _
                         2
End Sub

Private Sub UpdateCombinedProgress(ByRef progress As OverallProgress, phase As ProcessPhase, message As String)
    progress.CurrentStep = progress.CurrentStep + 1
    progress.Phase = message
    
    Dim percentDone As Integer
    If progress.TotalSteps > 0 Then
        percentDone = Int((progress.CurrentStep / progress.TotalSteps) * 100)
    End If
    
    Dim spinnerChar As String
    progress.SpinnerIndex = (progress.SpinnerIndex + 1) Mod 4
    spinnerChar = Mid(SPINNER_CHARS, progress.SpinnerIndex + 1, 1)
    
    Dim elapsedTime As Double
    elapsedTime = Timer - progress.StartTime
    
    Application.StatusBar = spinnerChar & " " & percentDone & "% | " & message & " | " & _
                           "Step " & progress.CurrentStep & " of " & progress.TotalSteps & " | " & _
                           "Time: " & FormatTime(elapsedTime)
End Sub

Private Sub FinalizeCombinedProcess(progress As OverallProgress, slicerCount As Long)
    Dim elapsedTime As Double
    elapsedTime = Timer - progress.StartTime
    
    Application.StatusBar = "✓ Complete! Created and connected " & slicerCount & " slicers | Time: " & FormatTime(elapsedTime)
    
    OptimizeExcelPerformance False
    
    Dim message As String
    message = "✅ Process completed successfully!" & vbCrLf & vbCrLf & _
              "Created: " & slicerCount & " pivot tables and slicers" & vbCrLf & _
              "Connected: All slicers to all pivot tables" & vbCrLf & _
              "Total time: " & FormatTime(elapsedTime)
    
    MsgBox message, vbInformation, "Success"
    
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
End Sub

Private Sub CleanupCombinedProcess()
    OptimizeExcelPerformance False
    Application.StatusBar = False
End Sub

' ==================================================================================
' WORKSHEET AND DATA MANAGEMENT
' ==================================================================================

Private Function GetDataWorksheet() As Worksheet
    On Error GoTo ErrorHandler
    Set GetDataWorksheet = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Exit Function
ErrorHandler:
    Err.Raise ERR_DATA_SHEET_NOT_FOUND, "GetDataWorksheet", "'" & DATA_SHEET_NAME & "' worksheet not found."
End Function

Private Function GetOrCreatePivotWorksheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(PIVOT_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = PIVOT_SHEET_NAME
    Else
        ClearWorksheetContent ws
    End If
    
    Set GetOrCreatePivotWorksheet = ws
End Function

Private Sub ClearWorksheetContent(ws As Worksheet)
    On Error Resume Next
    Dim pt As PivotTable, sc As SlicerCache
    
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    
    For Each sc In ThisWorkbook.SlicerCaches
        sc.Delete
    Next sc
    
    ws.Shapes.SelectAll
    Selection.Delete
    ws.Cells.Clear
    On Error GoTo 0
End Sub

Private Function GetValidatedDataRange(wsData As Worksheet) As Range
    Dim lastRow As Long, lastCol As Long
    
    lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Or lastCol < 1 Then
        Err.Raise ERR_NO_DATA_FOUND, "GetValidatedDataRange", "No valid data found."
    End If
    
    Set GetValidatedDataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
End Function

Private Function CreatePivotCache(dataRange As Range) As PivotCache
    Set CreatePivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
End Function

' ==================================================================================
' PIVOT TABLE CREATION
' ==================================================================================

Private Function CreatePivotTablesWithSlicers(wsPivot As Worksheet, pc As PivotCache, dataRange As Range, ByRef progress As OverallProgress) As Collection
    Dim allSlicers As New Collection
    Dim colCount As Long, currentRow As Long
    Dim colIndex As Long
    
    colCount = dataRange.Columns.Count
    currentRow = PIVOT_START_ROW
    
    For colIndex = 1 To colCount
        progress.CurrentStep = progress.CurrentStep + 1
        
        Dim fieldName As String
        fieldName = dataRange.Cells(1, colIndex).Value
        
        UpdateCombinedProgress progress, CreatingPivotTables, "Creating pivot table for: " & fieldName
        
        Dim pt As PivotTable
        Set pt = CreateSinglePivotTable(wsPivot, pc, fieldName, currentRow)
        
        If Not pt Is Nothing Then
            Dim slicer As slicer
            Set slicer = CreateSlicerForPivotTable(wsPivot, pt, fieldName)
            
            If Not slicer Is Nothing Then
                allSlicers.Add slicer
            End If
            
            currentRow = currentRow + pt.TableRange2.Rows.Count + PIVOT_ROW_SPACING
        End If
        
        Set pt = Nothing
        Set slicer = Nothing
    Next colIndex
    
    Set CreatePivotTablesWithSlicers = allSlicers
End Function

Private Function CreateSinglePivotTable(wsPivot As Worksheet, pc As PivotCache, fieldName As String, startRow As Long) As PivotTable
    On Error GoTo ErrorHandler
    
    Dim pt As PivotTable
    Set pt = wsPivot.PivotTables.Add( _
        PivotCache:=pc, _
        TableDestination:=wsPivot.Cells(startRow, 1))
    
    With pt
        .PivotFields(fieldName).Orientation = xlRowField
        .AddDataField .PivotFields(fieldName), "Count", xlCount
        .AddDataField .PivotFields(fieldName), "% of Total", xlCount
        .PivotFields("% of Total").Calculation = xlPercentOfTotal
    End With
    
    With wsPivot.Cells(startRow - 1, 1)
        .Value = fieldName
        .Font.Bold = True
    End With
    
    Set CreateSinglePivotTable = pt
    Exit Function
    
ErrorHandler:
    Set CreateSinglePivotTable = Nothing
End Function

' ==================================================================================
' SLICER CREATION AND MANAGEMENT
' ==================================================================================

Private Function CreateSlicerForPivotTable(wsPivot As Worksheet, pt As PivotTable, fieldName As String) As slicer
    On Error GoTo ErrorHandler
    
    Dim sc As SlicerCache
    Set sc = CreateSlicerCache_Compatible(pt, fieldName)
    
    If Not sc Is Nothing Then
        Set CreateSlicerForPivotTable = sc.Slicers.Add(wsPivot)
    End If
    Exit Function
    
ErrorHandler:
    Set CreateSlicerForPivotTable = Nothing
End Function

Private Function CreateSlicerCache_Compatible(pt As PivotTable, fieldName As String) As SlicerCache
    Dim sc As SlicerCache
    
    On Error Resume Next
    
    #If Mac Then
        Set sc = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
    #Else
        Set sc = ThisWorkbook.SlicerCaches.Add2(pt, fieldName)
        If sc Is Nothing Then
            Set sc = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
        End If
    #End If
    
    If sc Is Nothing Then
        Set sc = ThisWorkbook.SlicerCaches.Add(pt, pt.PivotFields(fieldName))
    End If
    
    On Error GoTo 0
    Set CreateSlicerCache_Compatible = sc
End Function

' ==================================================================================
' SLICER ORGANIZATION AND GROUPING
' ==================================================================================

Private Sub OrganizeSlicersByGroups(allSlicers As Collection, wsPivot As Worksheet, ByRef progress As OverallProgress)
    Dim sortedSlicers() As slicer
    Dim groups(1 To 3) As SlicerGroupConfig
    
    sortedSlicers = SortSlicersAlphabetically(allSlicers)
    InitializeSlicerGroups groups
    CategorizeSlicers sortedSlicers, groups
    PositionAndStyleGroups groups, wsPivot
End Sub

Private Function SortSlicersAlphabetically(slicers As Collection) As slicer()
    Dim slicerArray() As slicer
    Dim i As Long, j As Long
    Dim tempSlicer As slicer
    
    If slicers.Count = 0 Then Exit Function
    
    ReDim slicerArray(1 To slicers.Count)
    
    For i = 1 To slicers.Count
        Set slicerArray(i) = slicers(i)
    Next i
    
    For i = 1 To UBound(slicerArray) - 1
        For j = i + 1 To UBound(slicerArray)
            If slicerArray(i).Caption > slicerArray(j).Caption Then
                Set tempSlicer = slicerArray(i)
                Set slicerArray(i) = slicerArray(j)
                Set slicerArray(j) = tempSlicer
            End If
        Next j
    Next i
    
    SortSlicersAlphabetically = slicerArray
End Function

Private Sub InitializeSlicerGroups(ByRef groups() As SlicerGroupConfig)
    groups(MGroup).Name = "M_Group"
    groups(MGroup).Prefix = "M -"
    groups(MGroup).Color = COLOR_M_GROUP
    Set groups(MGroup).Slicers = New Collection
    
    groups(QGroup).Name = "Q_Group"
    groups(QGroup).Prefix = "Q -"
    groups(QGroup).Color = COLOR_Q_GROUP
    Set groups(QGroup).Slicers = New Collection
    
    groups(SQGroup).Name = "SQ_Group"
    groups(SQGroup).Prefix = "SQ -"
    groups(SQGroup).Color = COLOR_SQ_GROUP
    Set groups(SQGroup).Slicers = New Collection
End Sub

Private Sub CategorizeSlicers(slicers() As slicer, ByRef groups() As SlicerGroupConfig)
    Dim i As Long
    Dim slicerCaption As String
    
    For i = LBound(slicers) To UBound(slicers)
        slicerCaption = slicers(i).Caption
        
        If Left(slicerCaption, 3) = "M -" Then
            groups(MGroup).Slicers.Add slicers(i)
        ElseIf Left(slicerCaption, 3) = "Q -" Then
            groups(QGroup).Slicers.Add slicers(i)
        ElseIf Left(slicerCaption, 4) = "SQ -" Then
            groups(SQGroup).Slicers.Add slicers(i)
        End If
    Next i
End Sub

Private Sub PositionAndStyleGroups(groups() As SlicerGroupConfig, wsPivot As Worksheet)
    Dim groupLeft As Double, groupTop As Double
    Dim groupIndex As Integer
    
    groupTop = wsPivot.Rows(GROUP_TOP_ROW).Top
    groupLeft = wsPivot.Columns("E").Left
    
    For groupIndex = LBound(groups) To UBound(groups)
        If groups(groupIndex).Slicers.Count > 0 Then
            PositionSlicersInGrid groups(groupIndex).Slicers, groupLeft, groupTop
            ApplySlicerStyling_MacCompatible groups(groupIndex).Slicers, groups(groupIndex).Color
            
            If groups(groupIndex).Slicers.Count > 1 Then
                GroupSlicerShapes groups(groupIndex).Slicers, wsPivot, groups(groupIndex).Name
            End If
            
            Dim groupWidth As Double
            groupWidth = WorksheetFunction.Min(groups(groupIndex).Slicers.Count, SLICER_COLUMNS_PER_GROUP) * SLICER_LEFT_OFFSET
            groupLeft = groupLeft + groupWidth + SLICER_GROUP_SPACING
        End If
    Next groupIndex
End Sub

Private Sub PositionSlicersInGrid(slicers As Collection, startLeft As Double, startTop As Double)
    Dim slicer As slicer
    Dim slicerIndex As Long, row As Long, col As Long
    Dim slicerHeight As Double
    
    If slicers.Count = 0 Then Exit Sub
    
    slicerHeight = slicers(1).Shape.Height
    
    slicerIndex = 1
    For Each slicer In slicers
        row = (slicerIndex - 1) \ SLICER_COLUMNS_PER_GROUP
        col = (slicerIndex - 1) Mod SLICER_COLUMNS_PER_GROUP
        
        With slicer.Shape
            .Left = startLeft + col * SLICER_LEFT_OFFSET
            .Top = startTop + row * slicerHeight
        End With
        
        slicerIndex = slicerIndex + 1
    Next slicer
End Sub

' ==================================================================================
' MAC-COMPATIBLE STYLING
' ==================================================================================

Private Sub ApplySlicerStyling_MacCompatible(slicers As Collection, groupColor As Long)
    #If Mac Then
        ApplyBasicSlicerStyling slicers, groupColor
    #Else
        ApplyAdvancedSlicerStyling slicers, groupColor
    #End If
End Sub

#If Mac Then
Private Sub ApplyBasicSlicerStyling(slicers As Collection, groupColor As Long)
    Dim slicer As slicer
    
    On Error Resume Next
    For Each slicer In slicers
        With slicer.Shape
            If .Fill.Visible Then
                .Fill.ForeColor.RGB = groupColor
            End If
        End With
    Next slicer
    On Error GoTo 0
End Sub
#Else
Private Sub ApplyAdvancedSlicerStyling(slicers As Collection, groupColor As Long)
    Dim styleName As String
    Dim sty As SlicerStyle
    Dim slicer As slicer
    
    styleName = "CustomStyle_" & groupColor
    
    On Error Resume Next
    Set sty = ThisWorkbook.SlicerStyles(styleName)
    On Error GoTo 0
    
    If sty Is Nothing Then
        On Error Resume Next
        Set sty = ThisWorkbook.SlicerStyles.Add(styleName, "SlicerStyleLight1")
        If Not sty Is Nothing Then
            With sty
                .SlicerStyleElements(xlSlicerSelectedItemWithData).Interior.Color = groupColor
                .SlicerStyleElements(xlSlicerSelectedItemWithNoData).Interior.Color = groupColor
                .SlicerStyleElements(xlSlicerUnselectedItemWithData).Interior.Color = groupColor
                .SlicerStyleElements(xlSlicerUnselectedItemWithNoData).Interior.Color = groupColor
            End With
        End If
        On Error GoTo 0
    End If
    
    If Not sty Is Nothing Then
        For Each slicer In slicers
            On Error Resume Next
            slicer.Style = sty.Name
            On Error GoTo 0
        Next slicer
    End If
End Sub
#End If

Private Sub GroupSlicerShapes(slicers As Collection, wsPivot As Worksheet, groupName As String)
    On Error Resume Next
    
    Dim shapeNames() As String
    Dim i As Long
    Dim slicer As slicer
    
    ReDim shapeNames(1 To slicers.Count)
    
    i = 1
    For Each slicer In slicers
        shapeNames(i) = slicer.Name
        i = i + 1
    Next slicer
    
    Dim grpShape As Shape
    Set grpShape = wsPivot.Shapes.Range(shapeNames).Group
    If Not grpShape Is Nothing Then
        grpShape.Name = groupName & "_Slicers"
    End If
    
    On Error GoTo 0
End Sub

' ==================================================================================
' SLICER CONNECTION LOGIC
' ==================================================================================

Private Sub ConnectAllSlicersToAllPivotTables(wsPivot As Worksheet, ByRef progress As OverallProgress)
    Dim pivotTables As Collection
    Dim slicerCache As SlicerCache
    Dim newConnections As Long, existingConnections As Long
    
    Set pivotTables = CachePivotTables(wsPivot)
    
    For Each slicerCache In ThisWorkbook.SlicerCaches
        ConnectSlicerCacheToAllPivotTables slicerCache, pivotTables, progress, newConnections, existingConnections
    Next slicerCache
    
    UpdateCombinedProgress progress, ConnectingSlicers, _
        "Connected " & newConnections & " new, " & existingConnections & " existing"
End Sub

Private Function CachePivotTables(wsPivot As Worksheet) As Collection
    Dim pivotTables As New Collection
    Dim pt As PivotTable
    
    For Each pt In wsPivot.PivotTables
        pivotTables.Add pt
    Next pt
    
    Set CachePivotTables = pivotTables
End Function

Private Sub ConnectSlicerCacheToAllPivotTables(slicerCache As SlicerCache, pivotTables As Collection, ByRef progress As OverallProgress, ByRef newConnections As Long, ByRef existingConnections As Long)
    Dim connectedPivotNames As Collection
    Dim pt As PivotTable
    
    Set connectedPivotNames = GetConnectedPivotTableNames(slicerCache)
    
    For Each pt In pivotTables
        progress.CurrentStep = progress.CurrentStep + 1
        
        If Not IsPivotTableAlreadyConnected(pt.Name, connectedPivotNames) Then
            If ConnectSlicerToPivotTable(slicerCache, pt) Then
                newConnections = newConnections + 1
            End If
        Else
            existingConnections = existingConnections + 1
        End If
        
        If progress.CurrentStep Mod 5 = 0 Then
            UpdateCombinedProgress progress, ConnectingSlicers, "Connecting slicers... (" & newConnections & " new)"
        End If
    Next pt
End Sub

Private Function GetConnectedPivotTableNames(slicerCache As SlicerCache) As Collection
    Dim connectedNames As New Collection
    Dim pt As PivotTable
    
    On Error Resume Next
    For Each pt In slicerCache.PivotTables
        connectedNames.Add pt.Name
    Next pt
    On Error GoTo 0
    
    Set GetConnectedPivotTableNames = connectedNames
End Function

Private Function IsPivotTableAlreadyConnected(pivotTableName As String, connectedNames As Collection) As Boolean
    Dim connectedName As Variant
    
    For Each connectedName In connectedNames
        If CStr(connectedName) = pivotTableName Then
            IsPivotTableAlreadyConnected = True
            Exit Function
        End If
    Next connectedName
    
    IsPivotTableAlreadyConnected = False
End Function

Private Function ConnectSlicerToPivotTable(slicerCache As SlicerCache, pt As PivotTable) As Boolean
    On Error GoTo ErrorHandler
    
    slicerCache.PivotTables.AddPivotTable pt
    ConnectSlicerToPivotTable = True
    Exit Function
    
ErrorHandler:
    ConnectSlicerToPivotTable = False
    Debug.Print "Failed to connect slicer to pivot table: " & pt.Name & " - " & Err.Description
End Function

' ==================================================================================
' UTILITY FUNCTIONS
' ==================================================================================

Private Sub OptimizeExcelPerformance(optimize As Boolean)
    Static originalScreenUpdating As Boolean, originalCalculation As XlCalculation, originalEnableEvents As Boolean
    
    If optimize Then
        originalScreenUpdating = Application.ScreenUpdating
        originalCalculation = Application.Calculation
        originalEnableEvents = Application.EnableEvents
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        Application.ScreenUpdating = originalScreenUpdating
        Application.Calculation = originalCalculation
        Application.EnableEvents = originalEnableEvents
    End If
End Sub

Private Function FormatTime(seconds As Double) As String
    Dim minutes As Integer
    minutes = Int(seconds / 60)
    FormatTime = minutes & ":" & Format(Int(seconds Mod 60), "00")
End Function

Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

Private Sub HandleCombinedError(errNumber As Long, errDescription As String, procedureName As String)
    Dim errorMsg As String
    Select Case errNumber
        Case ERR_DATA_SHEET_NOT_FOUND
            errorMsg = "Data sheet '" & DATA_SHEET_NAME & "' not found."
        Case ERR_NO_DATA_FOUND
            errorMsg = "No valid data found in the data sheet."
        Case Else
            errorMsg = "Error in " & procedureName & ": " & errDescription
    End Select
    
    MsgBox "❌ " & errorMsg, vbCritical, "Process Error"
End Sub

Private Function IsMac() As Boolean
    #If Mac Then
        IsMac = True
    #Else
        IsMac = False
    #End If
End Function
