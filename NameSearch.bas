Attribute VB_Name = "NameSearch"
' VLOOKUP macro

Sub Search()
    Dim i As Integer
    Dim WB As Workbook
    Dim lRow As Long
    Dim fso As New FileSystemObject
    
    file = "PATH\TO\FILE\names.xlsx"
    
    sfile = ThisWorkbook.Name
    ofile = fso.GetFileName(file)
    Set WB = Workbooks.Open(file)
    Workbooks(sfile).Activate
        
    If Not WB Is Nothing Then
        
        'Find last row
        lRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
            
        For i = 1 To lRow
            Application.StatusBar = "Progress: " & i & " of " & lRow & ": " & Format(i / lRow, "0%")
            DoEvents
            curCell = Workbooks(sfile).ActiveSheet.Cells(CStr(i), "B") ' go through this column
            'MsgBox curCell
            If curCell <> "" Then
                getName = Application.VLookup(curCell, Workbooks(ofile).Sheets(1).Range("A:B"), 2, False) ' Data stored here
                If Not IsError(getName) Then
                    Workbooks(sfile).ActiveSheet.Cells(CStr(i), "C") = getName ' print results in this column
                End If
            End If
        Next i
        Call WB.Close(file)
    End If
    Application.StatusBar = "Done"
    
End Sub

