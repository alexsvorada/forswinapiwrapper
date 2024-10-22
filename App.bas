Attribute VB_Name = "App"
Option Explicit

Public Sub Runnable()
Attribute Runnable.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim wrapper As t_ForsWinAPIWrapper
    initForsWinAPIWrapper wrapper
    
    With ThisWorkbook.Worksheets("Main")
        If Not IsEmpty(.Range("B" & 2)) Then
            processCommands wrapper, .Range("B" & 2).value
            Exit Sub
        End If
    End With
    
    With ThisWorkbook.Worksheets("Data")
        'Code starts here
        processCommands wrapper, ""
    End With
End Sub
