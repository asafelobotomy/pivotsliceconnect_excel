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
    ' Allows running just the creation part
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
    ' Allows running just the connection part
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
    ' Estimate total steps for better progress tracking
    Dim pivotTableCount As Long
    Dim slicerCount As Long
    
    pivotTableCount = columnCount
    slicerCount = columnCount
    
    progress.TotalSteps = 5 + ' Initialization phases
                         pivotTableCount + ' Creating pivot tables
                         slicerCount + ' Creating slicers  
                         slicerCount + ' Organizing slicers
                         (slicerCount * pivotTableCount) + ' Connecting slicers
                         2 ' Finalization
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
    
    ' Clear status bar after delay
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
End Sub

Private Sub CleanupCombinedProcess()
    OptimizeExcelPerformance False
    Application.StatusBar = False
End Sub

' ==================================================================================
' PIVOT TABLE CREATION (Enhanced from first macro)
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
        
        ' Create pivot table
        Dim pt As PivotTable
        Set pt = CreateSinglePivotTable(wsPivot, pc, fieldName, currentRow)
        
        If Not pt Is Nothing Then
            ' Create slicer for this pivot table
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
' SLICER CONNECTION (Enhanced from second macro)
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
        
        If progress.CurrentStep Mod 5 = 0 Then ' Update every 5 steps
            UpdateCombinedProgress progress, ConnectingSlicers, "Connecting slicers... (" & newConnections & " new)"
        End If
    Next pt
End Sub

' ==================================================================================
' UTILITY FUNCTIONS (Shared between both functionalities)
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

' ==================================================================================
' PLACEHOLDER FUNCTIONS (Include all other functions from previous macros)
' ==================================================================================

' Include all the other functions from both previous macros here:
' - CreateSlicerForPivotTable
' - OrganizeSlicersByGroups
' - CachePivotTables
' - GetConnectedPivotTableNames
' - IsPivotTableAlreadyConnected
' - ConnectSlicerToPivotTable
' - All slicer styling functions
' - All slicer grouping functions
' etc.

