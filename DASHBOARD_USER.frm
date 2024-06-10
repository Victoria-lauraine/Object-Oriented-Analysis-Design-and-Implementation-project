VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DASHBOARD_USER 
   Caption         =   "UserForm1"
   ClientHeight    =   6204
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10260
   OleObjectBlob   =   "DASHBOARD_USER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DASHBOARD_USER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
ORDER.Show
End Sub

Private Sub CommandButton3_Click()
Unload Me
LOGIN.Show
End Sub

