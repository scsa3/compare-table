Attribute VB_Name = "NewMacros"
Sub CompareTableCreater()
    Dim newDocument As Document, macroDocument As Document
    Set macroDocument = ActiveDocument
    Set newDocument = Documents.Add
    
    CreateTable newDocument
    macroDocument.Activate
    CopyPasteRevisions newDocument
    ' CleanTable newDocument
End Sub
Private Sub CreateTable(newDocument As Document)
    Dim myRange As Range
    Set myRange = newDocument.Range(Start:=0, End:=0)
    Dim myTable As Table
    Set myTable = newDocument.Tables.Add(Range:=myRange, NumRows:=1, NumColumns:=4)
    With myTable.Borders
        .InsideColor = wdColorAutomatic
        .InsideLineStyle = wdLineStyleSingle
        .OutsideColor = wdColorAutomatic
        .OutsideLineStyle = wdLineStyleSingle
    End With
    myTable.Cell(1, 1).Range.InsertAfter "Page Number"
    myTable.Cell(1, 2).Range.InsertAfter "Page Number/Page Number"
    myTable.Cell(1, 3).Range.InsertAfter "Before"
    myTable.Cell(1, 4).Range.InsertAfter "After"
End Sub
Private Sub CopyPasteRevisions(newDocument As Document)
    Dim compateTable As Table
    Set compateTable = newDocument.Tables(1)
    Dim latestRow As Row
    Dim copyRange As Range
    
    For Each p In ActiveDocument.Paragraphs
        Set myRange = p.Range
        myRange.End = myRange.End - 1
        
        If myRange.Revisions.Count < 1 Then GoTo NextIteration
        If myRange.End - myRange.Start <= 0 Then GoTo NextIteration
        If myRange.InlineShapes.Count > 0 Then GoTo NextIteration
        If myRange.ShapeRange.Count > 0 Then GoTo NextIteration
        myRange.Copy

        compateTable.Rows.Add
        Set latestRow = compateTable.Rows.Last
        latestRow.Cells(1).Range.InsertAfter myRange.Information(wdActiveEndAdjustedPageNumber)
        latestRow.Cells(2).Range.InsertAfter myRange.Information(wdActiveEndAdjustedPageNumber)
        latestRow.Cells(2).Range.InsertAfter "/"
        latestRow.Cells(2).Range.InsertAfter myRange.Information(wdActiveEndAdjustedPageNumber)
        latestRow.Cells(3).Range.Paste
        latestRow.Cells(3).Range.Revisions.RejectAll
        latestRow.Cells(4).Range.Paste
NextIteration:
    Next
End Sub
' Private Sub CleanTable(newDocument As Document)
'     Dim compareTable As Table
'     Dim compareRow As Row
'     Dim rejectRevision As Revision
'     Set compareTable = newDocument.Tables(1)
'     For Each compareRow In compareTable.Rows
'         For Each rejectRevision In compareRow.Cells(3).Range.Revisions
'             rejectRevision.Reject
'         Next
'     Next
' End Sub

