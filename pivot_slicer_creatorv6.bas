Option Explicit
Attribute VB_Name = "PivotSlicerSafeModule"

' ==================================================================================
' SAFE PERFORMANCE OPTIMIZED VERSION
' ==================================================================================

' Layout Constants
Private Const PIVOT_START_ROW As Long = 23
Private Const SLICER_LEFT_OFFSET As Double = 150
Private Const SLICER_GROUP_SPACING As Double = 10
Private Const SLICER_COLUMNS_PER_GROUP As Integer = 3
Private Const GROUP_TOP_ROW As Long = 20
Private Const PIVOT_ROW_SPACING As Long = 2

' Performance Constants
Private Const PROGRESS_UPDATE_FREQUENCY As Long = 5
Private Const BATCH_SIZE As Long = 3

' Color Constants
Private Const COLOR_M_GROUP As Long = 15921906
Private Const COLOR_Q_GROUP As Long = 14349306
Private Const COLOR_SQ_GROUP As Long = 16244215

' Worksheet Names
Private Const DATA_SHEET_NAME As String = "Tidied Data"
Private Const PIVOT_SHEET_NAME As String = "PivotTable"

' Debug flag
Private Const DEBUG_MODE As Boolean = True

' Type Definitions
Private Type SlicerInfo
    SlicerObject As Slicer
    Caption As String
    GroupType As Integer
    Position As Integer
End Type

Private Type PerformanceMetrics
    StartTime As Double
    PivotCreationTime As Double
    SlicerCreationTime As Double
    ConnectionTime As Double
    TotalOperations As Long
End Type

' ==================================================================================
' MAIN SAFE ENTRY POINT
' ==================================================================================

Public Sub CreateAndConnectPivotTablesAndSlicers_Safe()
    Dim metrics As PerformanceMetrics
    metrics.StartTime = Timer
    
    On Error GoTo ErrorHandler
    
    ' Conservative Excel optimization
    Call OptimizeExcelPerformance_Conservative(True)
    
    ' Step 1: Validate workbook
    DebugPrint "Step 1: Validating workbook..."
    If Not ValidateWorkbook_Safe() Then
        MsgBox "❌ Workbook validation failed", vbCritical
        GoTo CleanupAndExit
    End If
    
    ' Step 2: Get worksheets
    DebugPrint "Step 2: Getting worksheets..."
    Dim wsData As Worksheet, wsPivot As Worksheet
    If Not GetWorksheets_Safe(wsData, wsPivot) Then
        MsgBox "❌ Failed to get worksheets", vbCritical
        GoTo CleanupAndExit
    End If
    
    ' Step 3: Get data range
    DebugPrint "Step 3: Getting data range..."
    Dim dataRange As Range
    If Not GetDataRange_Safe(wsData, dataRange) Then
        MsgBox "❌ Failed to get data range", vbCritical
        GoTo CleanupAndExit
    End If
    
    ' Step 4: Create pivot cache
    DebugPrint "Step 4: Creating pivot cache..."
    Dim pivotCache As PivotCache
    If Not CreatePivotCache_Safe(dataRange, pivotCache) Then
        MsgBox "❌ Failed to create pivot cache", vbCritical
        GoTo CleanupAndExit
    End If
    
    ' Step 5: Extract field names
    DebugPrint "Step 5: Extracting field names..."
    Dim fieldNames() As String
    fieldNames = ExtractFieldNames_Safe(dataRange)
    
    Application.StatusBar = "Creating " & UBound(fieldNames) + 1 & " pivot tables..."
    
    ' Step 6: Create pivot tables
    DebugPrint "Step 6: Creating pivot tables..."
    metrics.PivotCreationTime = Timer
    Dim pivotTables() As PivotTable
    If Not CreatePivotTables_Safe(wsPivot, pivotCache, fieldNames, pivotTables) Then
        MsgBox "❌ Failed to create pivot tables", vbCritical
        GoTo CleanupAndExit
    End If
    metrics.PivotCreationTime = Timer - metrics.PivotCreationTime
    
    ' Step 7: Create slicers
    DebugPrint "Step 7: Creating slicers..."
    Application.StatusBar = "Creating slicers..."
    metrics.SlicerCreationTime = Timer
    Dim slicerInfos() As SlicerInfo
    If Not CreateSlicers_Safe(wsPivot, pivotTables, fieldNames, slicerInfos) Then
        MsgBox "❌ Failed to create slicers", vbCritical
        GoTo CleanupAndExit
    End If
    metrics.SlicerCreationTime = Timer - metrics.SlicerCreationTime
    
    ' Step 8: Organize slicers
    DebugPrint "Step 8: Organizing slicers..."
    Application.StatusBar = "Organizing slicers..."
    OrganizeSlicers_Safe wsPivot, slicerInfos
    
    ' Step 9: Connect slicers
    DebugPrint "Step 9: Connecting slicers..."
    Application.StatusBar = "Connecting slicers..."
    metrics.ConnectionTime = Timer
    Dim connectionsCount As Long
    connectionsCount = ConnectSlicers_Safe(pivotTables, slicerInfos)
    metrics.ConnectionTime = Timer - metrics.ConnectionTime
    
    ' Step 10: Show results
    ShowResults_Safe metrics, UBound(slicerInfos) + 1, connectionsCount
    
CleanupAndExit:
    Call OptimizeExcelPerformance_Conservative(False)
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Call OptimizeExcelPerformance_Conservative(False)
    Application.StatusBar = False
    
    Dim errorMsg As String
    errorMsg = "Error in CreateAndConnectPivotTablesAndSlicers_Safe:" & vbCrLf & _
               "Number: " & Err.Number & vbCrLf & _
               "Description: " & Err.Description & vbCrLf & _
               "Source: " & Err.Source
    
    MsgBox "❌ " & errorMsg, vbCritical, "Process Error"
    DebugPrint "ERROR: " & errorMsg
End Sub

' ==================================================================================
' CONSERVATIVE PERFORMANCE OPTIMIZATION
' ==================================================================================

Private Sub OptimizeExcelPerformance_Conservative(enable As Boolean)
    Static originalScreenUpdating As Boolean
    Static originalCalculation As XlCalculation
    Static originalEnableEvents As Boolean
    Static stored As Boolean
    
    On Error Resume Next
    
    If enable Then
        If Not stored Then
            originalScreenUpdating = Application.ScreenUpdating
            originalCalculation = Application.Calculation
            originalEnableEvents = Application.EnableEvents
            stored = True
        End If
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        If stored Then
            Application.ScreenUpdating = originalScreenUpdating
            Application.Calculation = originalCalculation
            Application.EnableEvents = originalEnableEvents
        End If
    End If
    
    On Error GoTo 0
End Sub

' ==================================================================================
' SAFE VALIDATION AND SETUP
' ==================================================================================

Private Function ValidateWorkbook_Safe() As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if data sheet exists
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    
    ' Check if it has data
    If ws.UsedRange.Rows.Count < 2 Then
        DebugPrint "ERROR: No data found in " & DATA_SHEET_NAME
        ValidateWorkbook_Safe = False
        Exit Function
    End If
    
    DebugPrint "SUCCESS: Workbook validation passed"
    ValidateWorkbook_Safe = True
    Exit Function
    
ErrorHandler:
    DebugPrint "ERROR: Workbook validation failed - " & Err.Description
    ValidateWorkbook_Safe = False
End Function

Private Function GetWorksheets_Safe(ByRef wsData As Worksheet, ByRef wsPivot As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ' Get data worksheet
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    DebugPrint "SUCCESS: Got data worksheet"
    
    ' Get or create pivot worksheet
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets(PIVOT_SHEET_NAME)
    On Error GoTo 0
    
    If wsPivot Is Nothing Then
        DebugPrint "Creating new pivot worksheet..."
        Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsPivot.Name = PIVOT_SHEET_NAME
        DebugPrint "SUCCESS: Created pivot worksheet"
    Else
        DebugPrint "Clearing existing pivot worksheet..."
        If Not ClearWorksheet_Safe(wsPivot) Then
            DebugPrint "WARNING: Worksheet clearing had issues"
        End If
    End If
    
    GetWorksheets_Safe = True
    Exit Function
    
ErrorHandler:
    DebugPrint "ERROR: Failed to get worksheets - " & Err.Description
    GetWorksheets_Safe = False
End Function

Private Function ClearWorksheet_Safe(ws As Worksheet) As Boolean
    On Error Resume Next
    
    DebugPrint "Clearing worksheet content..."
    
    ' Simple approach - just clear cells
    ws.Cells.Clear
    
    ' Clear slicer caches carefully
    Dim i As Long
    For i = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        ThisWorkbook.SlicerCaches(i).Delete
    Next i
    
    ' Clear shapes if they exist
    If ws.Shapes.Count > 0 Then
        For i = ws.Shapes.Count To 1 Step -1
            ws.Shapes(i).Delete
        Next i
    End If
    
    ClearWorksheet_Safe = True
    DebugPrint "SUCCESS: Cleared worksheet"
    On Error GoTo 0
End Function

Private Function GetDataRange_Safe(wsData As Worksheet, ByRef dataRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long, lastCol As Long
    
    With wsData
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set dataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
    End With
    
    DebugPrint "SUCCESS: Data range - " & dataRange.Address & " (" & dataRange.Rows.Count & " rows, " & dataRange.Columns.Count & " cols)"
    GetDataRange_Safe = True
    Exit Function
    
ErrorHandler:
    DebugPrint "ERROR: Failed to get data range - " & Err.Description
    GetDataRange_Safe = False
End Function

Private Function CreatePivotCache_Safe(dataRange As Range, ByRef pivotCache As PivotCache) As Boolean
    On Error GoTo ErrorHandler
    
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    DebugPrint "SUCCESS: Created pivot cache"
    CreatePivotCache_Safe = True
    Exit Function
    
ErrorHandler:
    DebugPrint "ERROR: Failed to create pivot cache - " & Err.Description
    CreatePivotCache_Safe = False
End Function

Private Function ExtractFieldNames_Safe(dataRange As Range) As String()
    Dim fieldNames() As String
    Dim i As Long
    
    ReDim fieldNames(0 To dataRange.Columns.Count - 1)
    
    For i = 1 To dataRange.Columns.Count
        fieldNames(i - 1) = CStr(dataRange.Cells(1, i).Value)
    Next i
    
    DebugPrint "SUCCESS: Extracted " & UBound(fieldNames) + 1 & " field names"
    ExtractFieldNames_Safe = fieldNames
End Function

' ==================================================================================
' SAFE PIVOT TABLE CREATION
' ==================================================================================

Private Function CreatePivotTables_Safe(wsPivot As Worksheet, pc As PivotCache, fieldNames() As String, ByRef pivotTables() As PivotTable) As Boolean
    On Error GoTo ErrorHandler
    
    ReDim pivotTables(0 To UBound(fieldNames))
    Dim currentRow As Long
    Dim i As Long, successCount As Long
    
    currentRow = PIVOT_START_ROW
    
    For i = 0 To UBound(fieldNames)
        DebugPrint "Creating pivot table " & i + 1 & " for field: " & fieldNames(i)
        
        Dim pt As PivotTable
        If CreateSinglePivotTable_Safe(wsPivot, pc, fieldNames(i), currentRow, pt) Then
            Set pivotTables(i) = pt
            successCount = successCount + 1
            
            ' Add title
            With wsPivot.Cells(currentRow - 1, 1)
                .Value = fieldNames(i)
                .Font.Bold = True
            End With
            
            currentRow = currentRow + pt.TableRange2.Rows.Count + PIVOT_ROW_SPACING
        End If
        
        If i Mod PROGRESS_UPDATE_FREQUENCY = 0 Then
            Application.StatusBar = "Creating pivot tables... (" & i + 1 & "/" & UBound(fieldNames) + 1 & ")"
        End If
    Next i
    
    DebugPrint "SUCCESS: Created " & successCount & " pivot tables"
    CreatePivotTables_Safe = (successCount > 0)
    Exit Function
    
ErrorHandler:
    DebugPrint "ERROR: Failed to create pivot tables - " & Err.Description
    CreatePivotTables_Safe = False
End Function

Private Function CreateSinglePivotTable_Safe(wsPivot As Worksheet, pc As PivotCache, fieldName As String, startRow As Long, ByRef pt As PivotTable) As Boolean
    On Error GoTo ErrorHandler
    
    Set pt = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Cells(startRow, 1))
    
    With pt
        .PivotFields(fieldName).Orientation = xlRowField
        .AddDataField .PivotFields(fieldName), "Count", xlCount
        .AddDataField .PivotFields(fieldName), "% of Total", xlCount
        .PivotFields("% of Total").Calculation = xlPercentOfTotal
    End With
    
    CreateSinglePivotTable_Safe = True
    Exit Function
    
ErrorHandler:
    DebugPrint "ERROR: Failed to create pivot table for " & fieldName & " - " & Err.Description
    CreateSinglePivotTable_Safe = False
End Function

' ==================================================================================
' SAFE SLICER CREATION
' ==================================================================================

Private Function CreateSlicers_Safe(wsPivot As Worksheet, pivotTables() As PivotTable, fieldNames() As String, ByRef slicerInfos() As SlicerInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ReDim slicerInfos(0 To UBound(pivotTables))
    Dim i As Long, successCount As Long
    
    For i = 0 To UBound(pivotTables)
        If Not pivotTables(i) Is Nothing Then
            DebugPrint "Creating slicer " & i + 1 & " for field: " & fieldNames(i)
            
            Dim slicer As slicer
            If CreateSingleSlicer_Safe(wsPivot, pivotTables(i), fieldNames(i), slicer) Then
                Set slicerInfos(i).SlicerObject = slicer
                slicerInfos(i).Caption = slicer.Caption
                slicerInfos(i).GroupType = DetermineGroupType_Safe(slicer.Caption)
                successCount = successCount + 1
            End If
        End If
        
        If i Mod PROGRESS_UPDATE_FREQUENCY = 0 Then
            Application.StatusBar = "Creating slicers... (" & i + 1 & "/" & UBound(pivotTables) + 1 & ")"
        End If
    Next i
    
    DebugPrint "SUCCESS: Created " & successCount & " slicers"
    CreateSlicers_Safe = (successCount > 0)
    Exit Function
    
ErrorHandler:
    DebugPrint "ERROR: Failed to create slicers - " & Err.Description
    CreateSlicers_Safe = False
End Function

Private Function CreateSingleSlicer_Safe(wsPivot As Worksheet, pt As PivotTable, fieldName As String, ByRef slicer As slicer) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sc As SlicerCache
    
    ' Mac-compatible approach
    #If Mac Then
        Set sc = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
    #Else
        Set sc = ThisWorkbook.SlicerCaches.Add2(pt, fieldName)
        If sc Is Nothing Then Set sc = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
    #End If
    
    If Not sc Is Nothing Then
        Set slicer = sc.Slicers.Add(wsPivot)
        CreateSingleSlicer_Safe = True
    End If
    Exit Function
    
ErrorHandler:
    DebugPrint "ERROR: Failed to create slicer for " & fieldName & " - " & Err.Description
    CreateSingleSlicer_Safe = False
End Function

Private Function DetermineGroupType_Safe(Caption As String) As Integer
    Select Case Left(Caption, 4)
        Case "M - ", "M -"
            DetermineGroupType_Safe = 1
        Case "Q - ", "Q -"
            DetermineGroupType_Safe = 2
        Case "SQ -"
            DetermineGroupType_Safe = 3
        Case Else
            DetermineGroupType_Safe = 0
    End Select
End Function

' ==================================================================================
' SAFE SLICER ORGANIZATION
' ==================================================================================

Private Sub OrganizeSlicers_Safe(wsPivot As Worksheet, ByRef slicerInfos() As SlicerInfo)
    On Error Resume Next
    
    DebugPrint "Organizing slicers..."
    
    ' Simple positioning - no complex sorting for now
    Dim groupLeft As Double, groupTop As Double
    Dim groupCounters(0 To 3) As Integer
    Dim i As Long
    
    groupTop = wsPivot.Rows(GROUP_TOP_ROW).Top
    groupLeft = wsPivot.Columns("E").Left
    
    For i = 0 To UBound(slicerInfos)
        If Not slicerInfos(i).SlicerObject Is Nothing Then
            Dim groupType As Integer
            groupType = slicerInfos(i).GroupType
            
            If groupType > 0 Then
                Dim Row As Integer, col As Integer
                Row = groupCounters(groupType) \ SLICER_COLUMNS_PER_GROUP
                col = groupCounters(groupType) Mod SLICER_COLUMNS_PER_GROUP
                
                With slicerInfos(i).SlicerObject.Shape
                    .Left = groupLeft + (groupType - 1) * (3 * SLICER_LEFT_OFFSET + SLICER_GROUP_SPACING) + col * SLICER_LEFT_OFFSET
                    .Top = groupTop + Row * 100
                End With
                
                groupCounters(groupType) = groupCounters(groupType) + 1
            End If
        End If
    Next i
    
    DebugPrint "SUCCESS: Organized slicers"
    On Error GoTo 0
End Sub

' ==================================================================================
' SAFE SLICER CONNECTION
' ==================================================================================

Private Function ConnectSlicers_Safe(pivotTables() As PivotTable, slicerInfos() As SlicerInfo) As Long
    On Error Resume Next
    
    DebugPrint "Connecting slicers..."
    
    Dim connectionsCount As Long
    Dim i As Long, j As Long
    
    For i = 0 To UBound(slicerInfos)
        If Not slicerInfos(i).SlicerObject Is Nothing Then
            Dim sc As SlicerCache
            Set sc = slicerInfos(i).SlicerObject.SlicerCache
            
            For j = 0 To UBound(pivotTables)
                If Not pivotTables(j) Is Nothing Then
                    sc.PivotTables.AddPivotTable pivotTables(j)
                    If Err.Number = 0 Then
                        connectionsCount = connectionsCount + 1
                    End If
                    Err.Clear
                End If
            Next j
        End If
        
        If i Mod BATCH_SIZE = 0 Then
            Application.StatusBar = "Connected " & connectionsCount & " slicer-pivot pairs..."
        End If
    Next i
    
    DebugPrint "SUCCESS: Made " & connectionsCount & " connections"
    ConnectSlicers_Safe = connectionsCount
    On Error GoTo 0
End Function

' ==================================================================================
' RESULTS AND UTILITIES
' ==================================================================================

Private Sub ShowResults_Safe(metrics As PerformanceMetrics, slicerCount As Long, connectionsCount As Long)
    Dim totalTime As Double
    totalTime = Timer - metrics.StartTime
    
    Dim message As String
    message = "✅ Process completed successfully!" & vbCrLf & vbCrLf & _
              "Created: " & slicerCount & " pivot tables & slicers" & vbCrLf & _
              "Connections: " & connectionsCount & " slicer-pivot pairs" & vbCrLf & _
              "Total time: " & FormatTime_Safe(totalTime) & vbCrLf & vbCrLf & _
              "Breakdown:" & vbCrLf & _
              "• Pivot creation: " & FormatTime_Safe(metrics.PivotCreationTime) & vbCrLf & _
              "• Slicer creation: " & FormatTime_Safe(metrics.SlicerCreationTime) & vbCrLf & _
              "• Connections: " & FormatTime_Safe(metrics.ConnectionTime)
    
    MsgBox message, vbInformation, "Success"
    DebugPrint "COMPLETED: " & message
End Sub

Private Function FormatTime_Safe(seconds As Double) As String
    If seconds < 1 Then
        FormatTime_Safe = Format(seconds * 1000, "0") & "ms"
    Else
        FormatTime_Safe = Format(seconds, "0.0") & "s"
    End If
End Function

Private Sub DebugPrint(message As String)
    If DEBUG_MODE Then
        Debug.Print Now & ": " & message
    End If
End Sub
