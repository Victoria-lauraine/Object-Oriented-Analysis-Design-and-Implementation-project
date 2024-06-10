VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DASHBOARD 
   Caption         =   "UserForm5"
   ClientHeight    =   7248
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   12816
   OleObjectBlob   =   "DASHBOARD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DASHBOARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
ORDER.Show
End Sub

Private Sub CommandButton10_Click()
EXPENSES.Show
End Sub

Private Sub CommandButton12_Click()
NEW_PRODUCT.Show
End Sub

Private Sub CommandButton2_Click()
PROCUREMENT_DATA.Show
End Sub

Private Sub CommandButton3_Click()
INVOICES.Show
End Sub

Private Sub CommandButton5_Click()
CUSTOMER.Show
End Sub

Private Sub CommandButton7_Click()
SUPPLIER_DATABASE.Show
End Sub

Private Sub CommandButton8_Click()
Dim response As VbMsgBoxResult
    
    'Displays a confirmation dialog before existing
    response = MsgBox("Do you really want to quit the application ?", vbYesNo + vbQuestion, "Quit")
    'If the user clicks "Yes", close the application
    If response = vbYes Then
        Application.Quit
    End If
End Sub

Private Sub CommandButton9_Click()
EMPLOYEE_DB2.Show
End Sub

Private Sub Label1_Click()

End Sub
