VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmailSenderForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5173
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7616
   OleObjectBlob   =   "EmailSenderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmailSenderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: EmailSenderForm
' Controls:
' - cmbAccount (ComboBox): Dropdown for email accounts
' - txtInterval (TextBox): Interval in seconds
' - txtBatch (TextBox): Batch size as count
' - lblETA (Label): Estimated time display
' - btnStart (Button): Starts the process

Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub txtBatch_Change()

End Sub

Private Sub UserForm_Initialize()
    Dim OutlookNamespace As Object
    Dim OutlookAccount As Object
    
    Set OutlookNamespace = GetObject(, "Outlook.Application").GetNamespace("MAPI")
    
    ' Populate the dropdown with email accounts
    For Each OutlookAccount In OutlookNamespace.Accounts
        cmbAccount.AddItem OutlookAccount.SmtpAddress
    Next OutlookAccount
    
    ' Set default values
    If cmbAccount.ListCount > 0 Then cmbAccount.ListIndex = 0
    txtInterval.Value = 60
    txtBatch.Value = 10
    lblETA.Caption = ""
End Sub

Private Sub cmbAccount_Change()
    Call CalculateETA
End Sub

Private Sub txtInterval_Change()
    Call CalculateETA
End Sub

Private Sub CalculateETA()
    Dim OutlookNamespace As Object
    Dim DraftsFolder As Object
    Dim RootFolder As Object
    Dim SelectedAccount As String
    Dim ItemCount As Integer
    Dim Interval As Integer
    Dim Batch As Integer
    
    Set OutlookNamespace = GetObject(, "Outlook.Application").GetNamespace("MAPI")
    SelectedAccount = cmbAccount.Value
    
    ' Find the selected account's Drafts folder
    For Each RootFolder In OutlookNamespace.Folders
        If LCase(RootFolder.Name) = LCase(SelectedAccount) Then
            Set DraftsFolder = RootFolder.Folders("Drafts")
            Exit For
        End If
    Next RootFolder
    
    If Not DraftsFolder Is Nothing Then
        ItemCount = DraftsFolder.Items.Count
        Interval = Val(txtInterval.Value)
        Batch = Val(txtBatch.Value)
        If Batch = 0 Then
            Batch = 1
        End If
        
        If ItemCount > 0 And Interval > 0 Then
            lblETA.Caption = Format((ItemCount * Interval) / (Batch * 60), "0.0") & " minutes"
        Else
            lblETA.Caption = "N/A"
        End If
    End If
End Sub

Private Sub btnStart_Click()
    Dim OutlookNamespace As Object
    Dim OutlookApp As Object
    Dim RootFolder As Object
    Dim SelectedAccount As String
    Dim i As Integer
    Dim TempItems As Object
    
    Set OutlookApp = GetObject(, "Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    SelectedAccount = cmbAccount.Value
    Interval = Val(txtInterval.Value)
    Batch = Val(txtBatch.Value)
    
    ' Find the correct Drafts folder
    For Each RootFolder In OutlookNamespace.Folders
        If LCase(RootFolder.Name) = LCase(SelectedAccount) Then
            Set DraftsFolder = RootFolder.Folders("Drafts")
            Exit For
        End If
    Next RootFolder
    
    ' Get all mail items in Drafts
    Set TempItems = CreateObject("Scripting.Dictionary") ' Temporary storage for valid emails
    
    ' Loop through all draft emails and check if they have a recipient
    For i = 1 To DraftsFolder.Items.Count
        If DraftsFolder.Items(i).Class = 43 Then ' Ensure it's a MailItem (43 = MailItem class)
            If Trim(DraftsFolder.Items(i).To) <> "" Then
                TempItems.Add TempItems.Count + 1, DraftsFolder.Items(i)
            End If
        End If
    Next i
    
    ' If no valid emails found, stop execution
    If TempItems.Count = 0 Then
        MsgBox "No valid drafts found with recipient addresses!", vbExclamation, "Error"
        SendingInProgress = False
        Exit Sub
    End If
    
    Set MailItems = TempItems
    CurrentIndex = MailItems.Count
    
    ' Check if there are drafts to send
    If MailItems.Count = 0 Then
        MsgBox "No drafts found in this folder!", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Confirm before starting
    If MsgBox("Start sending " & MailItems.Count & " emails at " & Interval & "s interval and " & Batch & " batch size?", vbYesNo + vbQuestion, "Confirm") = vbNo Then Exit Sub
    
    ' Minimize form
    Me.Hide
    
    ' Start sending emails in background
    modEmailSender.StartSending DraftsFolder, MailItems, Interval, Batch
End Sub



