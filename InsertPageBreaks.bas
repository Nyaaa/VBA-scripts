Attribute VB_Name = "InsertPageBreaks"
Sub InsertPageBreaks()
    Dim i As Long, J As Long
    J = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    For i = J To 2 Step -1
        If Range("A" & i).Value = "" Then
            ActiveSheet.HPageBreaks.Add Before:=Range("A" & i + 1)
        End If
    Next i
End Sub

