Attribute VB_Name = "CreateHyperLinks"
Sub CreateHyperLinks()

    Dim iRow, iCol As Integer

    lRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

    iRow = 2
    iCol = 8
        Do While iRow <= lRow
            If ActiveSheet.Cells(iRow, iCol).Value <> "" Then ' skipping empty cells
                ActiveSheet.Cells(iRow, iCol) = Range("folderPath").Value & ActiveSheet.Cells(iRow, iCol).Value
                MsgBox ActiveSheet.Cells(iRow, iCol).Value
                ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(iRow, iCol), Address:=ActiveSheet.Cells(iRow, iCol).Value, _
                TextToDisplay:=ActiveSheet.Cells(iRow, iCol).Value
            End If
            iRow = iRow + 1
        Loop
    
End Sub

