VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NEW_PRODUCT 
   Caption         =   "UserForm1"
   ClientHeight    =   4932
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10944
   OleObjectBlob   =   "NEW_PRODUCT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NEW_PRODUCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()
Dim response As VbMsgBoxResult
    
    'Displays a confirmation dialog before existing
    response = MsgBox("Do you really want to quit the application ?", vbYesNo + vbQuestion, "Quit")
    'If the user clicks "Yes", close the application
    If response = vbYes Then
        Application.Quit
    End If
End Sub

Private Sub Label6_Click()

End Sub
