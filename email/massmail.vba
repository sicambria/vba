Sub massmail()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range

    Application.ScreenUpdating = False
    Set OutApp = CreateObject("Outlook.Application")

    For Each cell In Worksheets("Sheet1").Columns("B").Cells
        Set OutMail = OutApp.CreateItem(0)
        If cell.Value Like "?*@?*.?*" Then      'try with less conditions first
            With OutMail
                .To = Cells(cell.Row, "B").Value
                .Subject = "Your CX-ray result"
                .Body = "Hi " + Cells(cell.Row, "A") + "," + vbCrLf + vbCrLf + " Your message... " + vbCrLf + Cells(cell.Row, "C").Value + vbCrLf + vbCrLf + "Br,"
                .display
                Stop                            'wait here for the stop
            End With
            Cells(cell.Row, "D").Value = "sent"
            Set OutMail = Nothing
        End If
    Next cell

    'Set OutApp = Nothing                        'it will be Nothing after End Sub
    Application.ScreenUpdating = True

End Sub
