Option Explicit
Attribute VB_Name = "PivotSlicerModule"

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

' Color Constants (using Long values for better Mac compatibility)
Private Const COLOR_M_GROUP As Long = 15921906  ' RGB(242, 220, 219) - Light Red
Private Const COLOR_Q_GROUP As Long = 14349306  ' RGB(226, 239, 218) - Light Green
Private Const COLOR_SQ_GROUP As Long = 16244215 ' RGB(222, 235, 247) - Light Blue

' Worksheet Names
Private Const DATA_SHEET_NAME As String = "Tidied Data"
Private Const PIVOT_SHEET_NAME As String = "PivotTable"

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

Private Type PivotConfig
    StartRow As Long
    SlicerOffset As Double
    GroupSpacing As Double
    ColumnsPerGroup As Integer
End Type

Private Enum SlicerGroupType
    MGroup = 1
    QGroup = 2
    SQGroup = 3
End Enum

' ==================================================================================
' MAIN ENTRY POINT
' ==================================================================================

Public Sub CreatePivotTablesAndSlicers()
    On Error GoTo ErrorHandler
    
    ' Optimize Excel performance
    OptimizeExcelPerformance True
    
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim dataRange As Range, pivotCache As PivotCache
    Dim allSlicers As Collection
    
    ' Initialize and validate data
    Set wsData = GetDataWorksheet()
    Set wsPivot = GetOrCreatePivotWorksheet()
    Set dataRange = GetValidatedDataRange(wsData)
    Set pivotCache = CreatePivotCache(dataRange)
    
    ' Create pivot tables and slicers
    Set allSlicers = CreatePivotTablesWithSlicers(wsPivot, pivotCache, dataRange)
    
    ' Organize slicers by groups
    If allSlicers.Count > 0 Then
        OrganizeSlicersByGroups allSlicers, wsPivot
    End If
    
    ' Cleanup and restore Excel settings
    CleanupObjects wsData, wsPivot, dataRange, pivotCache
    OptimizeExcelPerformance False
    Application.StatusBar = False
    
    MsgBox "Successfully created " & allSlicers.Count & " slicers and pivot tables!", vbInformation, "Success"
    Exit Sub
    
ErrorHandler:
    OptimizeExcelPerformance False
    Application.StatusBar = False
    HandleError Err.Number, Err.Description, "CreatePivotTablesAndSlicers"
End Sub

' ==================================================================================
' WORKSHEET AND DATA MANAGEMENT
' ==================================================================================

Private Function GetDataWorksheet() As Worksheet
    On Error GoTo ErrorHandler
    
    Set GetDataWorksheet = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Exit Function
    
ErrorHandler:
    Err.Raise ERR_DATA_SHEET_NOT_FOUND, "GetDataWorksheet", _
        "'" & DATA_SHEET_NAME & "' worksheet not found. Please ensure the data sheet exists."
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
        ' Clear existing pivot tables and slicers
        ClearWorksheetContent ws
    End If
    
    Set GetOrCreatePivotWorksheet = ws
End Function

Private Sub ClearWorksheetContent(ws As Worksheet)
    On Error Resume Next
    
    ' Clear pivot tables
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    
    ' Clear slicers
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        sc.Delete
    Next sc
    
    ' Clear shapes and remaining content
    ws.Shapes.SelectAll
    Selection.Delete
    ws.Cells.Clear
    
    On Error GoTo 0
End Sub

Private Function GetValidatedDataRange(wsData As Worksheet) As Range
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    
    ' Find actual data boundaries
    lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Validate data exists
    If lastRow < 2 Or lastCol < 1 Then
        Err.Raise ERR_NO_DATA_FOUND, "GetValidatedDataRange", _
            "No valid data found. Data must have headers and at least one data row."
    End If
    
    Set dataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
    
    ' Additional validation
    If Not ValidateDataRange(dataRange) Then
        Err.Raise ERR_NO_DATA_FOUND, "GetValidatedDataRange", _
            "Data validation failed. Please check your data format."
    End If
    
    Set GetValidatedDataRange = dataRange
End Function

Private Function ValidateDataRange(dataRange As Range) As Boolean
    ValidateDataRange = False
    
    If dataRange Is Nothing Then Exit Function
    If dataRange.Rows.Count < 2 Then Exit Function
    If dataRange.Columns.Count < 1 Then Exit Function
    
    ' Check for empty headers
    Dim col As Long
    For col = 1 To dataRange.Columns.Count
        If Trim(dataRange.Cells(1, col).Value) = "" Then Exit Function
    Next col
    
    ValidateDataRange = True
End Function

' ==================================================================================
' PIVOT TABLE AND CACHE MANAGEMENT
' ==================================================================================

Private Function CreatePivotCache(dataRange As Range) As PivotCache
    On Error GoTo ErrorHandler
    
    Set CreatePivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Exit Function
    
ErrorHandler:
    Err.Raise ERR_PIVOT_CREATION_FAILED, "CreatePivotCache", _
        "Failed to create pivot cache: " & Err.Description
End Function

Private Function CreatePivotTablesWithSlicers(wsPivot As Worksheet, pc As PivotCache, dataRange As Range) As Collection
    Dim allSlicers As New Collection
    Dim colCount As Long, currentRow As Long
    Dim colIndex As Long
    
    colCount = dataRange.Columns.Count
    currentRow = PIVOT_START_ROW
    
    For colIndex = 1 To colCount
        ShowProgress colIndex, colCount, "Creating Pivot Table"
        
        Dim fieldName As String
        fieldName = dataRange.Cells(1, colIndex).Value
        
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
            
            ' Update position for next pivot table
            currentRow = currentRow + pt.TableRange2.Rows.Count + PIVOT_ROW_SPACING
        End If
        
        ' Clean up object references
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
    
    ' Configure pivot table fields
    With pt
        .PivotFields(fieldName).Orientation = xlRowField
        .AddDataField .PivotFields(fieldName), "Count", xlCount
        .AddDataField .PivotFields(fieldName), "% of Total", xlCount
        .PivotFields("% of Total").Calculation = xlPercentOfTotal
    End With
    
    ' Add title above pivot table
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
    
    ' Try modern method first (Excel 2013+)
    #If Mac Then
        ' Mac version may not support Add2, try Add method
        Set sc = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
    #Else
        ' Try Add2 first for Windows
        Set sc = ThisWorkbook.SlicerCaches.Add2(pt, fieldName)
        If sc Is Nothing Then
            Set sc = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
        End If
    #End If
    
    ' Final fallback
    If sc Is Nothing Then
        Set sc = ThisWorkbook.SlicerCaches.Add(pt, pt.PivotFields(fieldName))
    End If
    
    On Error GoTo 0
    Set CreateSlicerCache_Compatible = sc
End Function

' ==================================================================================
' SLICER ORGANIZATION AND GROUPING
' ==================================================================================

Private Sub OrganizeSlicersByGroups(allSlicers As Collection, wsPivot As Worksheet)
    Dim sortedSlicers() As slicer
    Dim groups(1 To 3) As SlicerGroupConfig
    
    ' Sort slicers alphabetically
    sortedSlicers = SortSlicersAlphabetically(allSlicers)
    
    ' Initialize group configurations
    InitializeSlicerGroups groups
    
    ' Categorize slicers into groups
    CategorizeSlicers sortedSlicers, groups
    
    ' Position and style each group
    PositionAndStyleGroups groups, wsPivot
End Sub

Private Function SortSlicersAlphabetically(slicers As Collection) As slicer()
    Dim slicerArray() As slicer
    Dim i As Long, j As Long
    Dim tempSlicer As slicer
    
    If slicers.Count = 0 Then Exit Function
    
    ReDim slicerArray(1 To slicers.Count)
    
    ' Copy to array
    For i = 1 To slicers.Count
        Set slicerArray(i) = slicers(i)
    Next i
    
    ' Bubble sort by caption
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
    ' M Group
    groups(MGroup).Name = "M_Group"
    groups(MGroup).Prefix = "M -"
    groups(MGroup).Color = COLOR_M_GROUP
    Set groups(MGroup).Slicers = New Collection
    
    ' Q Group
    groups(QGroup).Name = "Q_Group"
    groups(QGroup).Prefix = "Q -"
    groups(QGroup).Color = COLOR_Q_GROUP
    Set groups(QGroup).Slicers = New Collection
    
    ' SQ Group
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
            ' Position slicers in grid
            PositionSlicersInGrid groups(groupIndex).Slicers, groupLeft, groupTop
            
            ' Apply styling (Mac-compatible)
            ApplySlicerStyling_MacCompatible groups(groupIndex).Slicers, groups(groupIndex).Color
            
            ' Group shapes if more than one slicer
            If groups(groupIndex).Slicers.Count > 1 Then
                GroupSlicerShapes groups(groupIndex).Slicers, wsPivot, groups(groupIndex).Name
            End If
            
            ' Update left position for next group
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
    
    ' Get height from first slicer
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
        ' Mac version - simplified styling
        ApplyBasicSlicerStyling slicers, groupColor
    #Else
        ' Windows version - full SlicerStyles support
        ApplyAdvancedSlicerStyling slicers, groupColor
    #End If
End Sub

#If Mac Then
Private Sub ApplyBasicSlicerStyling(slicers As Collection, groupColor As Long)
    Dim slicer As slicer
    
    On Error Resume Next
    For Each slicer In slicers
        ' Apply basic shape formatting for Mac
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
' UTILITY FUNCTIONS
' ==================================================================================

Private Sub OptimizeExcelPerformance(optimize As Boolean)
    If optimize Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If
End Sub

Private Sub ShowProgress(current As Long, total As Long, operation As String)
    Dim progressText As String
    progressText = operation & " (" & current & " of " & total & ")..."
    Application.StatusBar = progressText
End Sub

Private Sub CleanupObjects(ParamArray objects() As Variant)
    Dim i As Integer
    On Error Resume Next
    For i = 0 To UBound(objects)
        Set objects(i) = Nothing
    Next i
    On Error GoTo 0
End Sub

Private Sub HandleError(errNumber As Long, errDescription As String, procedureName As String)
    Dim errorMsg As String
    
    Select Case errNumber
        Case ERR_DATA_SHEET_NOT_FOUND
            errorMsg = "Data sheet not found. Please ensure '" & DATA_SHEET_NAME & "' exists."
        Case ERR_NO_DATA_FOUND
            errorMsg = "No valid data found. Please check your data format."
        Case ERR_PIVOT_CREATION_FAILED
            errorMsg = "Failed to create pivot tables. " & errDescription
        Case Else
            errorMsg = "An unexpected error occurred in " & procedureName & ": " & errDescription
    End Select
    
    MsgBox errorMsg, vbCritical, "Error"
End Sub

' ==================================================================================
' MAC DETECTION UTILITY
' ==================================================================================

Private Function IsMac() As Boolean
    #If Mac Then
        IsMac = True
    #Else
        IsMac = False
    #End If
End Function
