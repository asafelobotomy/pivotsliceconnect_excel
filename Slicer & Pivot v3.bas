Option Explicit
Attribute VB_Name = "PivotSlicerModule"
Sub CreatePivotTablesAndSlicers()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim lastRow As Long, lastCol As Long
    Dim colIndex As Integer, pivotRow As Long
    Dim rng As Range
    Dim slicer As Slicer
    Dim sc As SlicerCache
    Dim fieldName As String
    Dim slicerLeftBase As Double
    Dim slicerLeftOffset As Double
    Dim slicerTop(1 To 3) As Double
    Dim currentColumn As Integer
    Dim slicerCollection As Collection
    Dim slicersM As Collection, slicersQ As Collection, slicersSQ As Collection
    Dim grpShape As Shape

    ' Set worksheets
    Set wsData = ThisWorkbook.Sheets("Tidied Data")

    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets("PivotTable")
    On Error GoTo 0
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsPivot.Name = "PivotTable"
    End If

    ' Find last row and last column in Tidied Data
    lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Set data range
    Set rng = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    ' Create Pivot Cache
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rng)

    ' Start placing Pivot Tables from row 23
    pivotRow = 23

    ' Initialize slicer positions
    slicerLeftBase = wsPivot.Columns("E").Left
    slicerLeftOffset = 150
    slicerTop(1) = wsPivot.Rows(1).Top
    slicerTop(2) = wsPivot.Rows(1).Top
    slicerTop(3) = wsPivot.Rows(1).Top
    currentColumn = 1

    ' Collections to store slicers
    Set slicerCollection = New Collection
    Set slicersM = New Collection
    Set slicersQ = New Collection
    Set slicersSQ = New Collection

   ' Loop through each column to create Pivot Tables and Slicers
    For colIndex = 1 To lastCol
        ' Define the data range for the current column

        ' Create a Pivot Table
        Set pt = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Cells(pivotRow, 1))

        ' Set Pivot Table fields
        With pt
            .PivotFields(wsData.Cells(1, colIndex).Value).Orientation = xlRowField
            .AddDataField .PivotFields(wsData.Cells(1, colIndex).Value), "Count", xlCount
            .AddDataField .PivotFields(wsData.Cells(1, colIndex).Value), "% of Total", xlCount
            .PivotFields("% of Total").Calculation = xlPercentOfTotal
            .PivotFields(wsData.Cells(1, colIndex).Value).Orientation = xlRowField ' Add Data Column
        End With

        ' Add a title above the Pivot Table
        wsPivot.Cells(pivotRow - 1, 1).Value = wsData.Cells(1, colIndex).Value
        wsPivot.Cells(pivotRow - 1, 1).Font.Bold = True

        ' Move to the next position, leaving two-row space
        pivotRow = pivotRow + pt.TableRange2.Rows.Count + 2

        ' Determine the field for slicer (RowFields or ColumnFields)
        If pt.RowFields.Count > 0 Then
            fieldName = pt.RowFields(1).Name
        ElseIf pt.ColumnFields.Count > 0 Then
            fieldName = pt.ColumnFields(1).Name
        Else
            MsgBox "PivotTable '" & pt.Name & "' has no Row or Column fields to create a slicer.", vbExclamation
            GoTo NextPivot
        End If

        ' Create a slicer cache for the PivotTable and field
        On Error Resume Next
        Set sc = ThisWorkbook.SlicerCaches.Add2(pt, fieldName)
        On Error GoTo 0

        ' Add and position the slicer
        If Not sc Is Nothing Then
            Set slicer = sc.Slicers.Add(wsPivot)
            With slicer
                .Left = slicerLeftBase + (currentColumn - 1) * slicerLeftOffset ' Align based on first column and offset
                .Top = slicerTop(currentColumn)
                slicerTop(currentColumn) = slicerTop(currentColumn) + .Height ' Move to the next position in the column
            End With

            ' Store slicer object for later sorting
            slicerCollection.Add slicer

            ' Move to the next column, loop back to the first after the third
            currentColumn = currentColumn + 1
            If currentColumn > 3 Then currentColumn = 1
        End If

NextPivot:
    Next colIndex

    ' Convert collection to array for sorting
    Dim slicers() As Slicer
    Dim tempSlicer As Slicer
    Dim i As Integer, j As Integer

    ReDim slicers(1 To slicerCollection.Count)
    For i = 1 To slicerCollection.Count
        Set slicers(i) = slicerCollection(i)
    Next i

    ' Sort slicers alphabetically by caption
    For i = 1 To UBound(slicers) - 1
        For j = i + 1 To UBound(slicers)
            If slicers(i).Caption > slicers(j).Caption Then
                Set tempSlicer = slicers(i)
                Set slicers(i) = slicers(j)
                Set slicers(j) = tempSlicer
            End If
        Next j
    Next i

    ' Categorize slicers
    For i = 1 To UBound(slicers)
        If Left(slicers(i).Caption, 3) = "M -" Then
            slicersM.Add slicers(i)
        ElseIf Left(slicers(i).Caption, 3) = "Q -" Then
            slicersQ.Add slicers(i)
        ElseIf Left(slicers(i).Caption, 4) = "SQ -" Then
            slicersSQ.Add slicers(i)
        End If
    Next i

    Dim groupIndex As Integer
    Dim groupColl As Collection
    Dim shapeNames() As String
    Dim idx As Integer
    Dim prefix As String
    Dim groupColor As Long
    Dim groupTop As Double
    Dim groupLeft As Double
    Dim groupSpacing As Double

    ' Start placing groups from the left of the sheet
    groupTop = wsPivot.Rows(1).Top
    groupLeft = slicerLeftBase
    groupSpacing = 10

    For groupIndex = 1 To 3
        Select Case groupIndex
            Case 1
                Set groupColl = slicersM
                prefix = "M"
                groupColor = RGB(242, 220, 219) 'light red
            Case 2
                Set groupColl = slicersQ
                prefix = "Q"
                groupColor = RGB(226, 239, 218) 'light green
            Case 3
                Set groupColl = slicersSQ
                prefix = "SQ"
                groupColor = RGB(222, 235, 247) 'light blue
        End Select

        If Not groupColl Is Nothing And groupColl.Count > 0 Then
            ReDim shapeNames(1 To groupColl.Count)
            idx = 1

            Dim colPos As Integer
            Dim rowPos As Integer
            Dim slicerHeight As Double

            colPos = 0
            rowPos = 0
            slicerHeight = groupColl(1).Shape.Height

            For Each slicer In groupColl
                With slicer.Shape
                    .Left = groupLeft + colPos * slicerLeftOffset
                    .Top = groupTop + rowPos * slicerHeight
                    ' Ensure the color is limited to 24-bit range to avoid
                    ' "value out of range" errors on some Excel versions
                    .Fill.ForeColor.RGB = groupColor And &HFFFFFF
                End With
                shapeNames(idx) = slicer.Name
                idx = idx + 1

                colPos = colPos + 1
                If colPos >= 3 Then
                    colPos = 0
                    rowPos = rowPos + 1
                End If
            Next slicer

            ' Advance the starting left for the next group
            groupLeft = groupLeft + (WorksheetFunction.Min(groupColl.Count, 3) * slicerLeftOffset) + groupSpacing

            If groupColl.Count > 1 Then
                Set grpShape = wsPivot.Shapes.Range(shapeNames).Group
                grpShape.Name = "Group_" & prefix & "_Slicers"
            End If
        End If
    Next groupIndex

    MsgBox "Pivot Tables and Slicers Created Successfully!", vbInformation
End Sub

