VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Login : Directory Protector"
   ClientHeight    =   2805
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1657.287
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   2
      Top             =   2280
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   2325
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   1014.062
      X2              =   2366.144
      Y1              =   1205.3
      Y2              =   1205.3
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   1014.062
      X2              =   2366.144
      Y1              =   1205.3
      Y2              =   1205.3
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   112.674
      X2              =   3380.205
      Y1              =   496.3
      Y2              =   496.3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   112.674
      X2              =   3380.205
      Y1              =   496.3
      Y2              =   496.3
   End
   Begin VB.Label Label1 
      Caption         =   $"frmLogin.frx":0442
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-=-=-=--=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
' Directory Protector
'-----------------------------------------------
' Author: Mahatab-ur-Rashid
' E-Mail: mahatabur@yahoo.com
' Web Site: www15.brinkster.com/mahatabur
'-=-=-=--=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
Dim PassWord As String

On Error GoTo ErrHnd:

'This file save the password
Open "c:\pw.dat" For Input As #1
Input #1, PassWord

If txtPassword.Text = PassWord Then
Close #1
Load frmProtect
frmProtect.Show
frmProtect.Refresh
Unload frmLogin
Else
MsgBox "Your Password is not correct!", vbCritical, "Directory Protector"
txtPassword.Text = ""
End If

Close #1

ErrHnd:
Select Case Err.Number
Case 53:
Open "c:\pw.dat" For Append As #1
Write #1, txtPassword.Text
Close #1
MsgBox "Your new Password accepted!", vbInformation, "Directory Protector"
End Select
End Sub

