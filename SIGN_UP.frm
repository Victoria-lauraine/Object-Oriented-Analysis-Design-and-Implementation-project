VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SIGN_UP 
   Caption         =   "UserForm2"
   ClientHeight    =   7500
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9096
   OleObjectBlob   =   "SIGN_UP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SIGN_UP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim lr As Long
lr = Sheets("Data").Range("A" & Rows.Count).End(xlUp).Row + l

If MsgBox("Do you want to save data?", vbYesNo + vbQuestion, "Question") = vbNo Then
    Exit Sub
End If

Sheets("Data").Cells(lr, "A").Value = Me.TextBox1.Value
Sheets("Data").Cells(lr, "B").Value = Me.TextBox3.Value
Sheets("Data").Cells(lr, "C").Value = Me.TextBox2.Value

Me.TextBox1 = ""
Me.TextBox3 = ""
Me.TextBox2 = ""

SIGN_UP.Hide
LOGIN.Show
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Frame1_Click()

End Sub
