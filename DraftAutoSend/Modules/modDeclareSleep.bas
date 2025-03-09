Attribute VB_Name = "modDeclareSleep"
#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If
