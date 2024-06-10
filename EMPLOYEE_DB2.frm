VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EMPLOYEE_DB2 
   Caption         =   "UserForm7"
   ClientHeight    =   9024
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   14616
   OleObjectBlob   =   "EMPLOYEE_DB2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EMPLOYEE_DB2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton5_Click()
Dim response As VbMsgBoxResult
    
    'Displays a confirmation dialog before existing
    response = MsgBox("Do you really want to quit the application ?", vbYesNo + vbQuestion, "Quit")
    'If the user clicks "Yes", close the application
    If response = vbYes Then
        Application.Quit
    End If
End Sub

Private Sub Label17_Click()

End Sub

Private Sub UserForm_Click()

End Sub
