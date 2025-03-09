Attribute VB_Name = "modStopper"
Public Sub StopEmailSending()
    ContinueSending = False
    MsgBox "Email sending process will stop after the current batch.", vbInformation, "Stopping"
End Sub
