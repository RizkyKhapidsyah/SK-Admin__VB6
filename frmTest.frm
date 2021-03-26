VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Press Me!"
      Height          =   1395
      Left            =   465
      TabIndex        =   0
      Top             =   720
      Width           =   3765
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim objLogInTool        As InteractiveLogIn
    Set objLogInTool = New InteractiveLogIn
    If objLogInTool.LogInAs("Admin", "your momma - combat boots - shower", "Arlington") Then
        MsgBox "Successfully Impersonating!"
    End If
    objLogInTool.LogOut
End Sub
