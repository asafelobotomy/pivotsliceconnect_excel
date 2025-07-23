Option Explicit
Attribute VB_Name = "ConnectionsModule"

' ==================================================================================
' CONSTANTS AND TYPE DECLARATIONS
' ==================================================================================

' Worksheet Names
Private Const PIVOT_SHEET_NAME As String = "PivotTable"

' Error Constants
Private Const ERR_PIVOT_SHEET_NOT_FOUND As Long = vbObjectError + 2001
Private Const ERR_NO_PIVOT_TABLES As Long = vbObjectError + 2002
Private Const ERR_NO_SLICER_CACHES As Long = vbObjectError + 2003
Private Const ERR_CONNECTION_FAILED As Long = vbObjectError + 2004

' Progress Constants
Private Const PROGRESS_UPDATE_INTERVAL As Long = 5  ' Update every 5%
Private Const SPINNER_CHARS As String = "|/-\"

' Type Definitions
Private Type ConnectionStats
    TotalSteps As Long
    ProcessedSteps As Long
    NewConnections As Long
    AlreadyLinked As Long
    StartTime As Double
End Type

Private Type ProgressConfig
    LastPercent As Integer
    SpinnerIndex As Integer
    UpdateInterval As Integer
End Type

' ==================================================================================
' MAIN ENTRY POINT
' ==================================================================================

Public Sub ConnectSlicers_StatusBar_Final()
    On Error GoTo ErrorHandler
    
    ' Optimize Excel performance
    OptimizeExcelPerformance True
    
    Dim wsPivot As Worksheet
    Dim stats As ConnectionStats
    Dim progress As ProgressConfig
    
    ' Initialize and validate
    Set wsPivot = GetOrCreatePivotWorksheet()
    ValidateWorksheetForConnections wsPivot
    
    ' Initialize progress tracking
    InitializeProgressTracking progress
    stats.StartTime = Timer
    
    ' Perform the connections
    ConnectAllSlicersToPivotTables wsPivot, stats, progress
    
    ' Cleanup and show results
    CleanupAndShowResults stats
    Exit Sub
    
ErrorHandler:
    CleanupAndShowResults stats
    HandleConnectionError Err.Number, Err.Description, "ConnectSlicers_StatusBar_Final"
End Sub

' ==================================================================================
' WORKSHEET MANAGEMENT
' ==================================================================================

Private Function GetOrCreatePivotWorksheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(PIVOT_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = PIVOT_SHEET_NAME
    End If
    
    Set GetOrCreatePivotWorksheet = ws
End Function

Private Sub ValidateWorksheetForConnections(wsPivot As Worksheet)
    ' Check if pivot tables exist
    If wsPivot.PivotTables.Count = 0 Then
        Err.Raise ERR_NO_PIVOT_TABLES, "ValidateWorksheetForConnections", _
            "No pivot tables found in '" & PIVOT_SHEET_NAME & "' worksheet. Please create pivot tables first."
    End If
    
    ' Check if slicer caches exist
    If ThisWorkbook.SlicerCaches.Count = 0 Then
        Err.Raise ERR_NO_SLICER_CACHES, "ValidateWorksheetForConnections", _
            "No slicer caches found. Please create slicers first."
    End If
End Sub

' ==================================================================================
' CONNECTION LOGIC
' ==================================================================================

Private Sub ConnectAllSlicersToPivotTables(wsPivot As Worksheet, ByRef stats As ConnectionStats, ByRef progress As ProgressConfig)
    Dim pivotTables As Collection
    Dim slicerCache As SlicerCache
    
    ' Cache all pivot tables for better performance
    Set pivotTables = CachePivotTables(wsPivot)
    
    ' Calculate total steps for progress tracking
    stats.TotalSteps = pivotTables.Count * ThisWorkbook.SlicerCaches.Count
    
    ' Connect each slicer to all pivot tables
    For Each slicerCache In ThisWorkbook.SlicerCaches
        ConnectSlicerCacheToAllPivotTables slicerCache, pivotTables, stats, progress
    Next slicerCache
End Sub

Private Function CachePivotTables(wsPivot As Worksheet) As Collection
    Dim pivotTables As New Collection
    Dim pt As PivotTable
    
    For Each pt In wsPivot.PivotTables
        pivotTables.Add pt
    Next pt
    
    Set CachePivotTables = pivotTables
End Function

Private Sub ConnectSlicerCacheToAllPivotTables(slicerCache As SlicerCache, pivotTables As Collection, ByRef stats As ConnectionStats, ByRef progress As ProgressConfig)
    Dim connectedPivotNames As Collection
    Dim pt As PivotTable
    Dim isAlreadyConnected As Boolean
    
    ' Get list of already connected pivot tables
    Set connectedPivotNames = GetConnectedPivotTableNames(slicerCache)
    
    ' Connect to each pivot table if not already connected
    For Each pt In pivotTables
        stats.ProcessedSteps = stats.ProcessedSteps + 1
        
        isAlreadyConnected = IsPivotTableAlreadyConnected(pt.Name, connectedPivotNames)
        
        If Not isAlreadyConnected Then
            If ConnectSlicerToPivotTable(slicerCache, pt) Then
                stats.NewConnections = stats.NewConnections + 1
            End If
        Else
            stats.AlreadyLinked = stats.AlreadyLinked + 1
        End If
        
        ' Update progress
        UpdateProgressStatus stats, progress
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
    ' Log error but continue with other connections
    Debug.Print "Failed to connect slicer to pivot table: " & pt.Name & " - " & Err.Description
End Function

' ==================================================================================
' PROGRESS TRACKING
' ==================================================================================

Private Sub InitializeProgressTracking(ByRef progress As ProgressConfig)
    progress.LastPercent = -1
    progress.SpinnerIndex = 0
    progress.UpdateInterval = PROGRESS_UPDATE_INTERVAL
    
    ' Ensure status bar is visible
    Application.DisplayStatusBar = True
End Sub

Private Sub UpdateProgressStatus(stats As ConnectionStats, ByRef progress As ProgressConfig)
    Dim percentDone As Integer
    Dim shouldUpdate As Boolean
    
    If stats.TotalSteps = 0 Then Exit Sub
    
    percentDone = Int((stats.ProcessedSteps / stats.TotalSteps) * 100)
    shouldUpdate = (percentDone <> progress.LastPercent) And (percentDone Mod progress.UpdateInterval = 0)
    
    If shouldUpdate Or percentDone = 100 Then
        DisplayProgressMessage stats, progress, percentDone
        progress.LastPercent = percentDone
    End If
End Sub

Private Sub DisplayProgressMessage(stats As ConnectionStats, ByRef progress As ProgressConfig, percentDone As Integer)
    Dim spinnerChar As String
    Dim elapsedTime As Double
    Dim estimatedRemaining As Double
    Dim statusMessage As String
    
    ' Update spinner
    progress.SpinnerIndex = (progress.SpinnerIndex + 1) Mod 4
    spinnerChar = Mid(SPINNER_CHARS, progress.SpinnerIndex + 1, 1)
    
    ' Calculate timing
    elapsedTime = Timer - stats.StartTime
    If stats.ProcessedSteps > 0 Then
        estimatedRemaining = (stats.TotalSteps - stats.ProcessedSteps) * (elapsedTime / stats.ProcessedSteps)
    Else
        estimatedRemaining = 0
    End If
    
    ' Build status message
    statusMessage = BuildStatusMessage(spinnerChar, percentDone, stats, elapsedTime, estimatedRemaining)
    
    ' Update status bar
    Application.StatusBar = statusMessage
End Sub

Private Function BuildStatusMessage(spinnerChar As String, percentDone As Integer, stats As ConnectionStats, elapsedTime As Double, estimatedRemaining As Double) As String
    Dim message As String
    
    message = spinnerChar & " " & percentDone & "% complete | " & _
              "Processed: " & stats.ProcessedSteps & " / " & stats.TotalSteps & " | " & _
              "New: " & stats.NewConnections & " | " & _
              "Already: " & stats.AlreadyLinked & " | " & _
              "Elapsed: " & FormatTime(elapsedTime) & " | " & _
              "Remaining: ~" & FormatTime(estimatedRemaining)
    
    BuildStatusMessage = message
End Function

Private Function FormatTime(seconds As Double) As String
    Dim minutes As Integer
    Dim remainingSeconds As Integer
    
    minutes = Int(seconds / 60)
    remainingSeconds = Int(seconds Mod 60)
    
    FormatTime = minutes & ":" & Format(remainingSeconds, "00")
End Function

' ==================================================================================
' CLEANUP AND RESULTS
' ==================================================================================

Private Sub CleanupAndShowResults(stats As ConnectionStats)
    Dim finalMessage As String
    Dim elapsedTime As Double
    
    ' Calculate final elapsed time
    If stats.StartTime > 0 Then
        elapsedTime = Timer - stats.StartTime
    End If
    
    ' Final status bar update
    finalMessage = "✓ 100% complete | " & _
                  "Processed: " & stats.TotalSteps & " / " & stats.TotalSteps & " | " & _
                  "New: " & stats.NewConnections & " | " & _
                  "Already: " & stats.AlreadyLinked & " | " & _
                  "Total time: " & FormatTime(elapsedTime)
    
    Application.StatusBar = finalMessage
    
    ' Restore Excel settings
    OptimizeExcelPerformance False
    
    ' Show completion message
    ShowCompletionMessage stats, elapsedTime
    
    ' Clear status bar after a moment
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

Private Sub ShowCompletionMessage(stats As ConnectionStats, elapsedTime As Double)
    Dim message As String
    
    message = "✓ Slicer linking complete!" & vbCrLf & vbCrLf & _
              "New connections made: " & stats.NewConnections & vbCrLf & _
              "Already linked: " & stats.AlreadyLinked & vbCrLf & _
              "Total processed: " & stats.TotalSteps & vbCrLf & _
              "Time elapsed: " & FormatTime(elapsedTime)
    
    MsgBox message, vbInformation, "Connection Complete"
End Sub

Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

' ==================================================================================
' UTILITY FUNCTIONS
' ==================================================================================

Private Sub OptimizeExcelPerformance(optimize As Boolean)
    Static originalScreenUpdating As Boolean
    Static originalCalculation As XlCalculation
    Static originalEnableEvents As Boolean
    Static originalDisplayStatusBar As Boolean
    
    If optimize Then
        ' Save original settings
        originalScreenUpdating = Application.ScreenUpdating
        originalCalculation = Application.Calculation
        originalEnableEvents = Application.EnableEvents
        originalDisplayStatusBar = Application.DisplayStatusBar
        
        ' Optimize
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayStatusBar = True
    Else
        ' Restore original settings
        Application.ScreenUpdating = originalScreenUpdating
        Application.Calculation = originalCalculation
        Application.EnableEvents = originalEnableEvents
        Application.DisplayStatusBar = originalDisplayStatusBar
    End If
End Sub

Private Sub HandleConnectionError(errNumber As Long, errDescription As String, procedureName As String)
    Dim errorMsg As String
    
    Select Case errNumber
        Case ERR_PIVOT_SHEET_NOT_FOUND
            errorMsg = "PivotTable sheet not found. Please create pivot tables first."
        Case ERR_NO_PIVOT_TABLES
            errorMsg = "No pivot tables found. Please create pivot tables first."
        Case ERR_NO_SLICER_CACHES
            errorMsg = "No slicers found. Please create slicers first."
        Case ERR_CONNECTION_FAILED
            errorMsg = "Failed to connect slicers to pivot tables. " & errDescription
        Case Else
            errorMsg = "An unexpected error occurred in " & procedureName & ": " & errDescription
    End Select
    
    MsgBox "❌ " & errorMsg, vbCritical, "Connection Error"
End Sub

' ==================================================================================
' MAC COMPATIBILITY FUNCTIONS
' ==================================================================================

Private Function IsMac() As Boolean
    #If Mac Then
        IsMac = True
    #Else
        IsMac = False
    #End If
End Function

' ==================================================================================
' ADDITIONAL UTILITY FUNCTIONS
' ==================================================================================

Public Sub DisconnectAllSlicers()
    ' Utility function to disconnect all slicers (useful for testing)
    On Error Resume Next
    
    Dim slicerCache As SlicerCache
    Dim pt As PivotTable
    Dim disconnectedCount As Long
    
    Application.ScreenUpdating = False
    
    For Each slicerCache In ThisWorkbook.SlicerCaches
        For Each pt In slicerCache.PivotTables
            slicerCache.PivotTables.RemovePivotTable pt
            disconnectedCount = disconnectedCount + 1
        Next pt
    Next slicerCache
    
    Application.ScreenUpdating = True
    
    MsgBox "Disconnected " & disconnectedCount & " slicer connections.", vbInformation, "Disconnect Complete"
    On Error GoTo 0
End Sub

Public Sub ShowSlicerConnectionStatus()
    ' Utility function to show current connection status
    Dim slicerCache As SlicerCache
    Dim statusMessage As String
    Dim totalConnections As Long
    
    statusMessage = "Current Slicer Connections:" & vbCrLf & vbCrLf
    
    For Each slicerCache In ThisWorkbook.SlicerCaches
        statusMessage = statusMessage & "Slicer: " & slicerCache.Name & vbCrLf
        statusMessage = statusMessage & "Connected to " & slicerCache.PivotTables.Count & " pivot tables" & vbCrLf & vbCrLf
        totalConnections = totalConnections + slicerCache.PivotTables.Count
    Next slicerCache
    
    statusMessage = statusMessage & "Total connections: " & totalConnections
    
    MsgBox statusMessage, vbInformation, "Connection Status"
End Sub
