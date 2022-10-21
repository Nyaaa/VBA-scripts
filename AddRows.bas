Attribute VB_Name = "AddRows"
Sub resizeTable()

    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("Table1")
    n = Application.InputBox("How many rows to add?")
    'add new rows at the end of the table
    For i = 1 To n
        Set newRow = tbl.ListRows.Add
        With newRow
            .Range(5) = 1
            .Range(6) = "PCS"
        End With
    Next i
    
    'Autofill row numbers
    Set SourceRange = ActiveSheet.Range("A15:A16") 'table index column location
    Set fillRange = ActiveSheet.ListObjects("Table1").ListColumns(1).DataBodyRange
    SourceRange.AutoFill Destination:=fillRange

End Sub

