Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngCheck As Range
    Dim cel As Range
    Dim LastRow As Long
    Dim ws As Worksheet
    
    Set rngCheck = Range("A1:C1")
    
    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For Each cel In ws.Range("A1:C" & LastRow)
            If Not Intersect(cel, rngCheck) Is Nothing Then
                If WorksheetFunction.CountIf(rngCheck, cel.Value) > 1 Then
                    MsgBox "Dikkat! Mükerrer kayıt: " & cel.Value, vbInformation, "Uyarı"
                End If
            End If
        Next cel
    Next ws
End Sub
