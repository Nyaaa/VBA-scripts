Attribute VB_Name = "copySheets"
Sub Copier()

    Dim x As Integer
    x = InputBox("How many copies do you want?")
    For numtimes = 1 To x
        ActiveWorkbook.Sheets(1).Copy _
        After:=ActiveWorkbook.Sheets(1)
    Next
    
End Sub

