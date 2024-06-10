VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ORDER 
   Caption         =   "UserForm6"
   ClientHeight    =   9492
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   15528
   OleObjectBlob   =   "ORDER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton5_Click()
Dim response As VbMsgBoxResult
    
    'Displays a confirmation dialog before closing
    response = MsgBox("Do you ant to save the changes efore closing ?", vbYesNoCancel + vbQuestion, "Close workbook")
    'Perform the appropriate action based on the user's response
    Select Case response
        Case vbYes
           'Save the workbook and close it
           ActiveWorkbook.Save
           Application.Quit
        Case vbNo
           'Close the workbook without saving
           Application.Quit
        Case vbCancel
           'Cancel the close operation
           Exit Sub
    End Select
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image4_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub TextBox7_Change()

End Sub
