Option Explicit
Attribute VB_Name = "PivotSlicerOptimizedModule"

' ==================================================================================
' PERFORMANCE OPTIMIZED VERSION
' ==================================================================================

' Layout Constants
Private Const PIVOT_START_ROW As Long = 23
Private Const SLICER_LEFT_OFFSET As Double = 150
Private Const SLICER_GROUP_SPACING As Double = 10
Private Const SLICER_COLUMNS_PER_GROUP As Integer = 3
Private Const GROUP_TOP_ROW As Long = 20
Private Const PIVOT_ROW_SPACING As Long = 2

' Performance Constants
Private Const PROGRESS_UPDATE_FREQUENCY As Long = 10  ' Update every 10 operations
Private Const BATCH_SIZE As Long = 5  ' Process items in batches

' Color Constants
Private Const COLOR_M_GROUP As Long = 15921906
Private Const COLOR_Q_GROUP As Long = 14349306
Private Const COLOR_SQ_GROUP As Long = 16244215

' Worksheet Names
Private Const DATA_SHEET_NAME As String = "Tidied Data"
Private Const PIVOT_SHEET_NAME As String = "PivotTable"

' Type Definitions for Better Performance
Private Type SlicerInfo
    SlicerObject As slicer
    Caption As String
    GroupType As Integer  ' 1=M, 2=Q, 3=SQ, 0=Other
    Position As Integer   ' Position within group
End Type

Private Type ConnectionMatrix
    SlicerCacheIndex As Long
    PivotTableIndex As Long
    IsConnected As Boolean
End Type

Private Type PerformanceMetrics
    StartTime As Double
    PivotCreationTime As Double
    SlicerCreationTime As Double
    ConnectionTime As Double
    TotalOperations As Long
End Type

' ==================================================================================
' MAIN OPTIMIZED ENTRY POINT
' ==================================================================================

Public Sub CreateAndConnectPivotTablesAndSlicers_Optimized()
    On Error GoTo ErrorHandler
    
    Dim metrics As PerformanceMetrics
    metrics.StartTime = Timer
    
    ' Ultra-aggressive performance optimization
    OptimizeExcelPerformance_Aggressive True
    
    ' Pre-validate everything to avoid mid-process failures
    If Not PreValidateWorkbook() Then Exit Sub
    
    ' Get and cache all required objects upfront
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim dataRange As Range, pivotCache As PivotCache
    Dim fieldNames() As String
    
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set wsPivot = GetOrCreatePivotWorksheet_Fast()
    Set dataRange = GetDataRange_Fast(wsData)
    Set pivotCache = CreatePivotCache(dataRange)
    fieldNames = ExtractFieldNames_Fast(dataRange)
    
    Application.StatusBar = "🚀 Creating " & UBound(fieldNames) + 1 & " pivot tables..."
    
    ' Create all pivot tables in batch
    metrics.PivotCreationTime = Timer
    Dim pivotTables() As PivotTable
    pivotTables = CreateAllPivotTables_Batch(wsPivot, pivotCache, fieldNames)
    metrics.PivotCreationTime = Timer - metrics.PivotCreationTime
    
    ' Create all slicers in batch
    metrics.SlicerCreationTime = Timer
    Dim slicerInfos() As SlicerInfo
    slicerInfos = CreateAllSlicers_Batch(wsPivot, pivotTables, fieldNames)
    metrics.SlicerCreationTime = Timer - metrics.SlicerCreationTime
    
    Application.StatusBar = "🎯 Organizing " & UBound(slicerInfos) + 1 & " slicers..."
    
    ' Organize and position slicers (optimized)
    OrganizeSlicers_Optimized wsPivot, slicerInfos
    
    Application.StatusBar = "🔗 Connecting slicers to pivot tables..."
    
    ' Connect all slicers to all pivot tables (matrix approach)
    metrics.ConnectionTime = Timer
    ConnectSlicers_Matrix pivotTables, slicerInfos
    metrics.ConnectionTime = Timer - metrics.ConnectionTime
    
    ' Finalize
    ShowPerformanceResults metrics, UBound(slicerInfos) + 1
    OptimizeExcelPerformance_Aggressive False
    Exit Sub
    
ErrorHandler:
    OptimizeExcelPerformance_Aggressive False
    MsgBox "❌ Error: " & Err.Description, vbCritical
End Sub

' ==================================================================================
' ULTRA-AGGRESSIVE PERFORMANCE OPTIMIZATION
' ==================================================================================

Private Sub OptimizeExcelPerformance_Aggressive(enable As Boolean)
    Static originalSettings As Variant
    
    If enable Then
        ' Store original settings
        originalSettings = Array( _
            Application.ScreenUpdating, _
            Application.Calculation, _
            Application.EnableEvents, _
            Application.DisplayAlerts, _
            Application.DisplayStatusBar, _
            Application.PrintCommunication, _
            Application.Interactive)
        
        ' Maximize performance
        With Application
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
            .DisplayAlerts = False
            .DisplayStatusBar = True  ' Keep for progress
            .PrintCommunication = False
            .Interactive = False
        End With
    Else
        ' Restore original settings
        If Not IsEmpty(originalSettings) Then
            With Application
                .ScreenUpdating = originalSettings(0)
                .Calculation = originalSettings(1)
                .EnableEvents = originalSettings(2)
                .DisplayAlerts = originalSettings(3)
                .DisplayStatusBar = originalSettings(4)
                .PrintCommunication = originalSettings(5)
                .Interactive = originalSettings(6)
            End With
        End If
    End If
End Sub

' ==================================================================================
' FAST DATA EXTRACTION AND VALIDATION
' ==================================================================================

Private Function PreValidateWorkbook() As Boolean
    On Error GoTo ErrorHandler
    
    ' Quick validation without object creation
    If ThisWorkbook.Sheets(DATA_SHEET_NAME).UsedRange.Rows.Count < 2 Then
        MsgBox "❌ No data found in " & DATA_SHEET_NAME & " sheet", vbCritical
        PreValidateWorkbook = False
        Exit Function
    End If
    
    PreValidateWorkbook = True
    Exit Function
    
ErrorHandler:
    MsgBox "❌ Validation failed: " & Err.Description, vbCritical
    PreValidateWorkbook = False
End Function

Private Function GetOrCreatePivotWorksheet_Fast() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(PIVOT_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = PIVOT_SHEET_NAME
    Else
        ' Comprehensive clear
        ClearWorksheetContent_Fast ws
    End If
    
    Set GetOrCreatePivotWorksheet_Fast = ws
End Function

Private Sub ClearWorksheetContent_Fast(ws As Worksheet)
    On Error Resume Next
    
    ' Clear pivot tables first
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    
    ' Clear slicer caches (backward loop to avoid collection modification issues)
    Dim i As Long
    For i = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        ThisWorkbook.SlicerCaches(i).Delete
    Next i
    
    ' Clear shapes and remaining content
    ws.Shapes.SelectAll
    If Selection.Count > 0 Then Selection.Delete
    ws.Cells.Clear
    
    On Error GoTo 0
End Sub

Private Function GetDataRange_Fast(wsData As Worksheet) As Range
    ' Use more efficient range detection
    Dim lastRow As Long, lastCol As Long
    
    With wsData
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set GetDataRange_Fast = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
    End With
End Function

Private Function ExtractFieldNames_Fast(dataRange As Range) As String()
    ' Extract field names into array for faster access
    Dim fieldNames() As String
    Dim i As Long
    
    ReDim fieldNames(0 To dataRange.Columns.Count - 1)
    
    For i = 1 To dataRange.Columns.Count
        fieldNames(i - 1) = dataRange.Cells(1, i).Value
    Next i
    
    ExtractFieldNames_Fast = fieldNames
End Function

' ==================================================================================
' BATCH PIVOT TABLE CREATION
' ==================================================================================

Private Function CreateAllPivotTables_Batch(wsPivot As Worksheet, pc As PivotCache, fieldNames() As String) As PivotTable()
    Dim pivotTables() As PivotTable
    Dim currentRow As Long
    Dim i As Long
    
    ReDim pivotTables(0 To UBound(fieldNames))
    currentRow = PIVOT_START_ROW
    
    ' Create all pivot tables without individual formatting
    For i = 0 To UBound(fieldNames)
        Set pivotTables(i) = CreatePivotTable_Fast(wsPivot, pc, fieldNames(i), currentRow)
        
        If Not pivotTables(i) Is Nothing Then
            currentRow = currentRow + pivotTables(i).TableRange2.Rows.Count + PIVOT_ROW_SPACING
        End If
        
        ' Batch progress update
        If i Mod PROGRESS_UPDATE_FREQUENCY = 0 Then
            Application.StatusBar = "📊 Creating pivot tables... (" & i + 1 & "/" & UBound(fieldNames) + 1 & ")"
        End If
    Next i
    
    ' Batch format all titles at once
    FormatPivotTitles_Batch wsPivot, fieldNames
    
    CreateAllPivotTables_Batch = pivotTables
End Function

Private Function CreatePivotTable_Fast(wsPivot As Worksheet, pc As PivotCache, fieldName As String, startRow As Long) As PivotTable
    On Error GoTo ErrorHandler
    
    Dim pt As PivotTable
    Set pt = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Cells(startRow, 1))
    
    ' Configure fields in one go
    With pt
        .PivotFields(fieldName).Orientation = xlRowField
        .AddDataField .PivotFields(fieldName), "Count", xlCount
        .AddDataField .PivotFields(fieldName), "% of Total", xlCount
        .PivotFields("% of Total").Calculation = xlPercentOfTotal
    End With
    
    Set CreatePivotTable_Fast = pt
    Exit Function
    
ErrorHandler:
    Set CreatePivotTable_Fast = Nothing
End Function

Private Sub FormatPivotTitles_Batch(wsPivot As Worksheet, fieldNames() As String)
    ' Format all titles in batch using ranges
    Dim currentRow As Long, i As Long
    currentRow = PIVOT_START_ROW
    
    For i = 0 To UBound(fieldNames)
        With wsPivot.Cells(currentRow - 1, 1)
            .Value = fieldNames(i)
            .Font.Bold = True
        End With
        currentRow = currentRow + EstimatePivotTableRows() + PIVOT_ROW_SPACING
    Next i
End Sub

Private Function EstimatePivotTableRows() As Long
    ' Estimate average pivot table rows for spacing
    EstimatePivotTableRows = 10  ' Conservative estimate
End Function

' ==================================================================================
' OPTIMIZED SLICER CREATION
' ==================================================================================

Private Function CreateAllSlicers_Batch(wsPivot As Worksheet, pivotTables() As PivotTable, fieldNames() As String) As SlicerInfo()
    Dim slicerInfos() As SlicerInfo
    Dim i As Long
    
    ReDim slicerInfos(0 To UBound(pivotTables))
    
    For i = 0 To UBound(pivotTables)
        Set slicerInfos(i).SlicerObject = CreateSlicer_Fast(wsPivot, pivotTables(i), fieldNames(i))
        
        If Not slicerInfos(i).SlicerObject Is Nothing Then
            slicerInfos(i).Caption = slicerInfos(i).SlicerObject.Caption
            slicerInfos(i).GroupType = DetermineGroupType_Fast(slicerInfos(i).Caption)
        End If
        
        If i Mod PROGRESS_UPDATE_FREQUENCY = 0 Then
            Application.StatusBar = "🎛️ Creating slicers... (" & i + 1 & "/" & UBound(pivotTables) + 1 & ")"
        End If
    Next i
    
    CreateAllSlicers_Batch = slicerInfos
End Function

Private Function CreateSlicer_Fast(wsPivot As Worksheet, pt As PivotTable, fieldName As String) As slicer
    On Error GoTo ErrorHandler
    
    Dim sc As SlicerCache
    
    ' Mac-compatible slicer cache creation
    #If Mac Then
        Set sc = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
    #Else
        Set sc = ThisWorkbook.SlicerCaches.Add2(pt, fieldName)
        If sc Is Nothing Then Set sc = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
    #End If
    
    If Not sc Is Nothing Then
        Set CreateSlicer_Fast = sc.Slicers.Add(wsPivot)
    End If
    Exit Function
    
ErrorHandler:
    Set CreateSlicer_Fast = Nothing
End Function

Private Function DetermineGroupType_Fast(caption As String) As Integer
    ' Fast group type determination using Select Case
    Select Case Left(caption, 4)
        Case "M - ", "M -"
            DetermineGroupType_Fast = 1
        Case "Q - ", "Q -"
            DetermineGroupType_Fast = 2
        Case "SQ -"
            DetermineGroupType_Fast = 3
        Case Else
            DetermineGroupType_Fast = 0
    End Select
End Function

' ==================================================================================
' OPTIMIZED SLICER ORGANIZATION
' ==================================================================================

Private Sub OrganizeSlicers_Optimized(wsPivot As Worksheet, ByRef slicerInfos() As SlicerInfo)
    ' Quick sort using more efficient algorithm
    QuickSortSlicers slicerInfos, 0, UBound(slicerInfos)
    
    ' Pre-calculate all positions
    Dim positions() As Variant
    positions = CalculateAllPositions(slicerInfos, wsPivot)
    
    ' Apply positions and styling in batch
    ApplyPositionsAndStyling_Batch slicerInfos, positions
End Sub

Private Sub QuickSortSlicers(ByRef arr() As SlicerInfo, low As Long, high As Long)
    ' Much faster than bubble sort for large datasets
    If low < high Then
        Dim pivotIndex As Long
        pivotIndex = PartitionSlicers(arr, low, high)
        QuickSortSlicers arr, low, pivotIndex - 1
        QuickSortSlicers arr, pivotIndex + 1, high
    End If
End Sub

Private Function PartitionSlicers(ByRef arr() As SlicerInfo, low As Long, high As Long) As Long
    Dim pivot As String
    Dim i As Long, j As Long
    
    pivot = arr(high).Caption
    i = low - 1
    
    For j = low To high - 1
        If arr(j).Caption <= pivot Then
            i = i + 1
            SwapSlicerInfo arr(i), arr(j)
        End If
    Next j
    
    SwapSlicerInfo arr(i + 1), arr(high)
    PartitionSlicers = i + 1
End Function

Private Sub SwapSlicerInfo(ByRef a As SlicerInfo, ByRef b As SlicerInfo)
    Dim temp As SlicerInfo
    temp = a
    a = b
    b = temp
End Sub

Private Function CalculateAllPositions(slicerInfos() As SlicerInfo, wsPivot As Worksheet) As Variant()
    ' Pre-calculate all positions to avoid repeated calculations
    Dim positions() As Variant
    Dim groupCounters(0 To 3) As Integer
    Dim groupStarts(0 To 3) As Double
    Dim i As Long
    
    ReDim positions(0 To UBound(slicerInfos), 0 To 1)  ' Left, Top
    
    ' Calculate group starting positions
    groupStarts(1) = wsPivot.Columns("E").Left  ' M group
    groupStarts(2) = groupStarts(1) + (3 * SLICER_LEFT_OFFSET) + SLICER_GROUP_SPACING  ' Q group
    groupStarts(3) = groupStarts(2) + (3 * SLICER_LEFT_OFFSET) + SLICER_GROUP_SPACING  ' SQ group
    
    For i = 0 To UBound(slicerInfos)
        Dim groupType As Integer
        groupType = slicerInfos(i).GroupType
        
        If groupType > 0 Then
            Dim row As Integer, col As Integer
            row = groupCounters(groupType) \ SLICER_COLUMNS_PER_GROUP
            col = groupCounters(groupType) Mod SLICER_COLUMNS_PER_GROUP
            
            positions(i, 0) = groupStarts(groupType) + col * SLICER_LEFT_OFFSET  ' Left
            positions(i, 1) = wsPivot.Rows(GROUP_TOP_ROW).Top + row * 100  ' Top (estimated height)
            
            groupCounters(groupType) = groupCounters(groupType) + 1
        End If
    Next i
    
    CalculateAllPositions = positions
End Function

Private Sub ApplyPositionsAndStyling_Batch(slicerInfos() As SlicerInfo, positions() As Variant)
    ' Apply all positions and styling in batch
    Dim i As Long
    
    For i = 0 To UBound(slicerInfos)
        If Not slicerInfos(i).SlicerObject Is Nothing Then
            ' Position
            With slicerInfos(i).SlicerObject.Shape
                .Left = positions(i, 0)
                .Top = positions(i, 1)
            End With
            
            ' Basic styling (Mac-compatible)
            ApplyStyling_Fast slicerInfos(i).SlicerObject, slicerInfos(i).GroupType
        End If
    Next i
End Sub

Private Sub ApplyStyling_Fast(slicer As slicer, groupType As Integer)
    ' Fast styling without custom SlicerStyles
    On Error Resume Next
    
    Dim groupColor As Long
    Select Case groupType
        Case 1: groupColor = COLOR_M_GROUP
        Case 2: groupColor = COLOR_Q_GROUP
        Case 3: groupColor = COLOR_SQ_GROUP
        Case Else: Exit Sub
    End Select
    
    #If Mac Then
        ' Mac version - basic shape styling
        With slicer.Shape
            If .Fill.Visible Then .Fill.ForeColor.RGB = groupColor
        End With
    #Else
        ' Windows version - could use SlicerStyles if needed
        With slicer.Shape
            If .Fill.Visible Then .Fill.ForeColor.RGB = groupColor
        End With
    #End If
    
    On Error GoTo 0
End Sub

' ==================================================================================
' MATRIX-BASED CONNECTION (FASTEST METHOD)
' ==================================================================================

Private Sub ConnectSlicers_Matrix(pivotTables() As PivotTable, slicerInfos() As SlicerInfo)
    ' Use matrix approach for maximum efficiency
    Dim slicerCaches As Collection
    Dim connectionMatrix() As Boolean
    Dim i As Long, j As Long
    Dim connectionsEnabled As Long
    
    ' Build slicer cache collection
    Set slicerCaches = New Collection
    For i = 0 To UBound(slicerInfos)
        If Not slicerInfos(i).SlicerObject Is Nothing Then
            slicerCaches.Add slicerInfos(i).SlicerObject.SlicerCache
        End If
    Next i
    
    ' Create connection matrix
    ReDim connectionMatrix(1 To slicerCaches.Count, 0 To UBound(pivotTables))
    
    ' Batch connect all slicers to all pivot tables
    For i = 1 To slicerCaches.Count
        For j = 0 To UBound(pivotTables)
            On Error Resume Next
            slicerCaches(i).PivotTables.AddPivotTable pivotTables(j)
            If Err.Number = 0 Then
                connectionMatrix(i, j) = True
                connectionsEnabled = connectionsEnabled + 1
            End If
            On Error GoTo 0
        Next j
        
        ' Batch progress update
        If i Mod BATCH_SIZE = 0 Then
            Application.StatusBar = "🔗 Connected " & connectionsEnabled & " slicer-pivot pairs..."
        End If
    Next i
    
    Application.StatusBar = "✅ Connected " & connectionsEnabled & " total slicer-pivot pairs"
End Sub

' ==================================================================================
' PERFORMANCE REPORTING
' ==================================================================================

Private Sub ShowPerformanceResults(metrics As PerformanceMetrics, slicerCount As Long)
    Dim totalTime As Double
    totalTime = Timer - metrics.StartTime
    
    Dim message As String
    message = "🚀 Performance Report" & vbCrLf & vbCrLf & _
              "Created: " & slicerCount & " pivot tables & slicers" & vbCrLf & _
              "Total time: " & FormatTime(totalTime) & vbCrLf & vbCrLf & _
              "Breakdown:" & vbCrLf & _
              "• Pivot creation: " & FormatTime(metrics.PivotCreationTime) & vbCrLf & _
              "• Slicer creation: " & FormatTime(metrics.SlicerCreationTime) & vbCrLf & _
              "• Connections: " & FormatTime(metrics.ConnectionTime) & vbCrLf & vbCrLf & _
              "Speed: " & Round(slicerCount / totalTime, 1) & " items/second"
    
    Application.StatusBar = "✅ Complete! " & slicerCount & " items in " & FormatTime(totalTime)
    MsgBox message, vbInformation, "Performance Complete"
    
    ' Clear status bar after delay
    Application.OnTime Now + TimeValue("00:00:05"), "ClearStatusBar"
End Sub

Private Function FormatTime(seconds As Double) As String
    If seconds < 1 Then
        FormatTime = Format(seconds * 1000, "0") & "ms"
    Else
        FormatTime = Format(seconds, "0.0") & "s"
    End If
End Function

Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

' ==================================================================================
' UTILITY FUNCTIONS
' ==================================================================================

Private Function CreatePivotCache(dataRange As Range) As PivotCache
    Set CreatePivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
End Function
