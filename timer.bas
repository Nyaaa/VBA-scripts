Attribute VB_Name = "timer"

Sub update()
    Dim Conn As WorkbookConnection
    Dim StartingTime As Single

    StartingTime = timer

    For Each Conn In ThisWorkbook.Connections
        Conn.Refresh
    Next Conn

    Application.StatusBar = Format((timer - StartingTime) / 86400, "hh:mm:ss")

End Sub
