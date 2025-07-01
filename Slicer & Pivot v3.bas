Attribute VB_Name = "Module1"
Sub CreatePivotTablesAndSlicers()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As pivotTable
    Dim lastRow As Long, lastCol As Long
    Dim colIndex As Integer, pivotRow As Long
    Dim rng As Range, pRange As Range
    Dim slicer As slicer
    Dim sc As slicerCache
    Dim fieldName As String
    Dim slicerLeftBase As Double
    Dim slicerLeftOffset As Double
    Dim slicerTop(1 To 3) As Double
    Dim currentColumn As Integer
    Dim slicerCollection As Collection
    Dim grpShape As Shape

    ' Set worksheets
    Set wsData = ThisWorkbook.Sheets("Tidied Data")
    Set wsPivot = ThisWorkbook.Sheets("PivotTable")

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

    ' Collection to store slicer names and captions for sorting
    Set slicerCollection = New Collection

   ' Loop through each column to create Pivot Tables and Slicers
    For colIndex = 1 To lastCol
        ' Define the data range for the current column
        Set pRange = wsData.Range(wsData.Cells(1, colIndex), wsData.Cells(lastRow, colIndex))

        ' Create a Pivot Table
        Set pt = wsPivot.pivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Cells(pivotRow, 1))

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

            ' Add slicer name and caption to the collection
            slicerCollection.Add slicer.Name & ":" & slicer.Caption

            ' Move to the next column, loop back to the first after the third
            currentColumn = currentColumn + 1
            If currentColumn > 3 Then currentColumn = 1
        End If

NextPivot:
    Next colIndex

     ' Sort slicers alphabetically by caption
    Dim sortedSlicers() As String
    Dim temp As String
    Dim i As Integer, j As Integer

    ReDim sortedSlicers(1 To slicerCollection.Count)

    For i = 1 To slicerCollection.Count
        sortedSlicers(i) = slicerCollection(i)
    Next i

    For i = 1 To UBound(sortedSlicers) - 1
        For j = i + 1 To UBound(sortedSlicers)
            If Split(sortedSlicers(i), ":")(1) > Split(sortedSlicers(j), ":")(1) Then
                temp = sortedSlicers(i)
                sortedSlicers(i) = sortedSlicers(j)
                sortedSlicers(j) = temp
            End If
        Next j
    Next i

    ' Reposition slicers and group by prefix
    Dim slicerName As String, captionText As String
    Dim shapeArrayM() As String, shapeArrayQ() As String, shapeArraySQ() As String
    Dim mCount As Integer, qCount As Integer, sqCount As Integer
    Dim sortedTop(1 To 3) As Double

    mCount = 0: qCount = 0: sqCount = 0

    For i = LBound(sortedSlicers) To UBound(sortedSlicers)
        captionText = Split(sortedSlicers(i), ":")(1)
        If Left(captionText, 3) = "M -" Then mCount = mCount + 1
        If Left(captionText, 3) = "Q -" Then qCount = qCount + 1
        If Left(captionText, 4) = "SQ -" Then sqCount = sqCount + 1
    Next i

    ReDim shapeArrayM(1 To mCount)
    ReDim shapeArrayQ(1 To qCount)
    ReDim shapeArraySQ(1 To sqCount)

    mCount = 1: qCount = 1: sqCount = 1
    sortedColumn = 1
    sortedTop(1) = wsPivot.Rows(1).Top
    sortedTop(2) = wsPivot.Rows(1).Top
    sortedTop(3) = wsPivot.Rows(1).Top

    For i = LBound(sortedSlicers) To UBound(sortedSlicers)
        slicerName = Split(sortedSlicers(i), ":")(0)
        captionText = Split(sortedSlicers(i), ":")(1)

        With wsPivot.Shapes(slicerName)
            .Left = slicerLeftBase + (sortedColumn - 1) * slicerLeftOffset
            .Top = sortedTop(sortedColumn)
            sortedTop(sortedColumn) = sortedTop(sortedColumn) + .Height
        End With

        If Left(captionText, 3) = "M -" Then
            shapeArrayM(mCount) = slicerName: mCount = mCount + 1
        ElseIf Left(captionText, 3) = "Q -" Then
            shapeArrayQ(qCount) = slicerName: qCount = qCount + 1
        ElseIf Left(captionText, 4) = "SQ -" Then
            shapeArraySQ(sqCount) = slicerName: sqCount = sqCount + 1
        End If

        sortedColumn = sortedColumn + 1
        If sortedColumn > 3 Then sortedColumn = 1
    Next i

    If mCount > 1 Then
        Set grpShape = wsPivot.Shapes.Range(shapeArrayM).Group
        grpShape.Name = "Group_M_Slicers"
    End If
    If qCount > 1 Then
        Set grpShape = wsPivot.Shapes.Range(shapeArrayQ).Group
        grpShape.Name = "Group_Q_Slicers"
    End If
    If sqCount > 1 Then
        Set grpShape = wsPivot.Shapes.Range(shapeArraySQ).Group
        grpShape.Name = "Group_SQ_Slicers"
    End If

    MsgBox "Pivot Tables and Slicers Created Successfully!", vbInformation
End Sub

