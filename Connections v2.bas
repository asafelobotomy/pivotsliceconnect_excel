Option Explicit
Attribute VB_Name = "ConnectionsModule"
Sub ConnectSlicers_StatusBar_Final()
    Dim wsPivot As Worksheet
    Dim slicerCache As SlicerCache
    Dim pt As PivotTable
    Dim scPT As PivotTable
    Dim ptList As Collection
    Dim connectedCount As Long
    Dim alreadyLinkedCount As Long
    Dim stepIndex As Long
    Dim totalSteps As Long
    Dim startTime As Double
    Dim elapsedTime As Double
    Dim estRemaining As Double
    Dim isConnected As Boolean
    Dim scNames As Collection
    Dim ptName As String
    Dim linkedName As Variant
    Dim spinnerChars As Variant
    Dim spinnerChar As String
    Dim percentDone As Integer
    Dim lastPercent As Integer
    Dim spinnerIndex As Integer
    Dim prevStatusBar As Boolean

    On Error GoTo SafeExit

    ' Setup
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets("PivotTable")
    On Error GoTo 0
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsPivot.Name = "PivotTable"
    End If
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    prevStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

    ' Spinner config
    spinnerChars = Array("|", "/", "-", "\")
    spinnerIndex = 0
    lastPercent = -1

    ' Cache all PivotTables
    Set ptList = New Collection
    For Each pt In wsPivot.PivotTables
        ptList.Add pt
    Next pt

    totalSteps = ptList.Count * ThisWorkbook.SlicerCaches.Count
    connectedCount = 0
    alreadyLinkedCount = 0
    stepIndex = 0
    startTime = Timer

    ' Main loop
    For Each slicerCache In ThisWorkbook.SlicerCaches
        Set scNames = New Collection
        For Each scPT In slicerCache.PivotTables
            scNames.Add scPT.Name
        Next scPT

        For Each pt In ptList
            stepIndex = stepIndex + 1
            isConnected = False
            ptName = pt.Name

            For Each linkedName In scNames
                If linkedName = ptName Then
                    isConnected = True
                    Exit For
                End If
            Next linkedName

            If Not isConnected Then
                slicerCache.PivotTables.AddPivotTable pt
                connectedCount = connectedCount + 1
            Else
                alreadyLinkedCount = alreadyLinkedCount + 1
            End If

            ' Status bar update only when % changes
            percentDone = Int((stepIndex / totalSteps) * 100)
            If percentDone <> lastPercent Then
                spinnerChar = spinnerChars(spinnerIndex Mod 4)
                spinnerIndex = spinnerIndex + 1
                lastPercent = percentDone
                elapsedTime = Timer - startTime
                estRemaining = (totalSteps - stepIndex) * (elapsedTime / stepIndex)

                Application.StatusBar = spinnerChar & " " & percentDone & "% complete | Connected: " & connectedCount & " / " & totalSteps & _
                    " | Elapsed: " & Round(elapsedTime / 60, 0) & " min | Remaining: ~" & Round(estRemaining / 60, 0) & " min"
            End If
        Next pt
    Next slicerCache

SafeExit:
    elapsedTime = Timer - startTime
    Application.StatusBar = "? 100% complete | Connected: " & connectedCount & " / " & totalSteps & _
        " | Elapsed: " & Round(elapsedTime / 60, 0) & " min | Remaining: ~0 min"

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = prevStatusBar

    If Err.Number <> 0 Then
        MsgBox "? Error: " & Err.Description, vbExclamation
    Else
        MsgBox "? Slicer linking complete. " & connectedCount & " connections made.", vbInformation
    End If
End Sub


