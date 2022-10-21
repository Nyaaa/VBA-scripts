Attribute VB_Name = "FileList"
Dim folderPath As String

Private Sub cmd_button_BROWSEforFolder_Click()

    On Error GoTo err
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFolderPicker)

    'To allow or disable to multi select
    fileExplorer.AllowMultiSelect = False

    With fileExplorer
        If .Show = -1 Then 'Any folder is selected
            folderPath = .SelectedItems.Item(1)

        Else ' else dialog is cancelled
            MsgBox "You have cancelled the dialogue"
            folderPath = "NONE" ' when cancelled set blank as file path.
            End
        End If
    End With
err:

End Sub

Sub ListFiles()
    cmd_button_BROWSEforFolder_Click
    ThisWorkbook.ActiveSheet.Cells.ClearContents
    Cells(1, 1) = "#"
    Cells(1, 2) = "Filename"
    Cells(1, 3) = "Date created"
    Cells(1, 4) = "Folder"
    subFolderLoop
    ThisWorkbook.ActiveSheet.Columns("A:D").AutoFit
End Sub

Private Sub subFolderLoop()
    Dim fso, oFolder, oSubfolder, oFile, queue As Collection

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
    queue.Add fso.GetFolder(folderPath)

    Do While queue.Count > 0
        Set oFolder = queue(1)
        queue.Remove 1 'dequeue
        '...insert any folder processing code here...
        For Each oSubfolder In oFolder.SubFolders
            queue.Add oSubfolder 'enqueue
        Next oSubfolder
        For Each oFile In oFolder.Files
            sArray = Split(oFile.Path, "\")
            lastindex = UBound(sArray) - 1
            Cells(i + 2, 1) = i + 1
            Cells(i + 2, 2) = oFile.Name
            Cells(i + 2, 3) = oFile.DateCreated
            Cells(i + 2, 4) = sArray(lastindex)
            i = i + 1
        Next oFile
    Loop

End Sub
