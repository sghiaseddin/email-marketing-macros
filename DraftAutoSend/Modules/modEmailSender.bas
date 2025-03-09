Attribute VB_Name = "modEmailSender"
Public DraftsFolder As Object
Public MailItems As Object
Public CurrentIndex As Integer
Public Interval As Integer
Public Batch As Integer
Public SelectedAccount As String
Public SendingInProgress As Boolean
Public ContinueSending As Boolean


Public Sub StartSending(SelectedDrafts As Object, MailList As Object, MailInterval As Integer, MailBatch As Integer)
    If SendingInProgress Then
        MsgBox "Sending process is already running!", vbExclamation, "Warning"
        Exit Sub
    End If
    
    ContinueSending = True ' Allow sending process to continue
    SendingInProgress = True
    Set DraftsFolder = SelectedDrafts
    Set MailItems = MailList
    CurrentIndex = MailItems.Count
    Interval = MailInterval
    Batch = MailBatch
    SelectedAccount = MailSelectedAccount

    ' Start background sending
    SendNextEmail Interval, Batch
End Sub

Public Sub SendNextEmail(MailInterval As Integer, MailBatch As Integer)
    Dim OutlookNamespace As Object
    Dim OutlookApp As Object
    Dim RootFolder As Object
    Dim OutboxFolder As Object
    Dim MailItem As Object
    Dim i As Integer
    
    If Not SendingInProgress Or CurrentIndex = 0 Then
        SendingInProgress = False
        MsgBox "All emails sent successfully!", vbInformation, "Completed"
        Exit Sub
    End If
    
    Interval = MailInterval
    Batch = MailBatch
    Do While i <= Batch And CurrentIndex > 0
        Set MailItem = MailItems(CurrentIndex)
        
        ' Send the email
        MailItem.Send
        
        ' Move to the next email
        CurrentIndex = CurrentIndex - 1
        i = i + 1
    Loop
    
    ' Trigger "Send All" in Outlook (like clicking the button)
    Set OutlookApp = GetObject(, "Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    OutlookNamespace.SendAndReceive True
    
    ' Schedule next send using a background timer
    ScheduleNextEmail Interval, Batch
End Sub

Private Sub ScheduleNextEmail(MailInterval As Integer, MailBatch As Integer)
    Dim StartTime As Double
    StartTime = Timer
    Interval = MailInterval
    Batch = MailBatch
    
    If Not ContinueSending Then
        MsgBox "Email sending has been stopped.", vbInformation, "Process Stopped"
        Exit Sub
    End If
    
    ' Run in background without freezing UI
    Do While Timer < StartTime + Interval
        DoEvents  ' Allow Outlook to remain responsive
    Loop
    
    ' Call SendNextEmail again after the interval
    SendNextEmail Interval, Batch
End Sub

Public Sub StopSending()
    SendingInProgress = False
    MsgBox "Sending process stopped!", vbExclamation, "Stopped"
End Sub
