VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LOGIN 
   Caption         =   "UserForm1"
   ClientHeight    =   8172
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9288
   OleObjectBlob   =   "LOGIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbLogin_Click()
    
End Sub

Private Sub cmdExit_Click()
    Dim response As VbMsgBoxResult
    
    'Displays a confirmation dialog before existing
    response = MsgBox("Do you really want to quit the application ?", vbYesNo + vbQuestion, "Quit")
    'If the user clicks "Yes", close the application
    If response = vbYes Then
        Application.Quit
    End If
End Sub

Private Sub cmdLOGIN_Click()
   DASHBOARD_USER.Show
End Sub

Private Sub CommandButton3_Click()
SIGN_UP.Show
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub login_Click()

End Sub

Private Sub TextBox2_Change()

End Sub
