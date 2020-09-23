VERSION 5.00
Begin VB.Form frmProtect 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory Protector : Add a Directory"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmProtect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3240
      TabIndex        =   12
      Top             =   1320
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1650
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Width           =   5775
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1770
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " [ Select a Directory ] "
      ForeColor       =   &H00000000&
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Add Directory"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Remove Protection"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   4
         Left            =   5520
         Picture         =   "frmProtect.frx":0442
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   3
         Left            =   4920
         Picture         =   "frmProtect.frx":0884
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   2
         Left            =   4320
         Picture         =   "frmProtect.frx":1672
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   1
         Left            =   3720
         Picture         =   "frmProtect.frx":1AB4
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   0
         Left            =   3120
         Picture         =   "frmProtect.frx":1EF6
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   465
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   3120
         X2              =   6000
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   3120
         X2              =   6000
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   6000
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   240
         X2              =   6000
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Protected Directory List:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Directory's New View as:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   7
         Top             =   960
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Directory Name [ With Path ]:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   2565
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   2880
         X2              =   2880
         Y1              =   480
         Y2              =   3120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   2880
         X2              =   2880
         Y1              =   480
         Y2              =   3120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Directory List:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Drive List:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmProtect"
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

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'Following's are some special Extensions we add with the folder name...
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'SUPPOSE we have a Folder name "New Folder". Now we
'add ".{21EC2020-3AEA-1069-A2DD-08002B30309D}" as extension
'with that folder, that means we Rename "New Folder"
'to "New Folder.{21EC2020-3AEA-1069-A2DD-08002B30309D}".
'You will see that it becomes "Control Panel"!!
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Const Control_P = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Const My_COMP = ".{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
Const Desk_TOP = ".{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}"
Const NetWork = ".{208D2C60-3AEA-1069-A2D7-08002B30309D}"
Const IE = ".{FBF23B42-E3F0-101B-8488-00AA003E56F8}"
Const RBin = ".{645FF040-5081-101B-9F08-00AA002F954E}"
Const Printer = ".{2227A280-3AEA-1069-A2DE-08002B30309D}"
Const HTMLDoc = ".{25336920-03F9-11CF-8FD0-00AA00686F13}"
Const TaskS = ".{255b3f60-829e-11cf-8d8b-00aa0060f5bf}"
Const WaveFile = ".{0003000D-0000-0000-C000-000000000046}"
Const MovClip = ".{00022602-0000-0000-C000-000000000046}"
Const WinIcon = ".{00021401-0000-0000-C000-000000000046}"


Private Sub Form_Load()
On Error GoTo errnum
Dim Str, Attrib, Icon As String
Combo1.Clear
List1.Clear

Combo1.AddItem "My Network Places"
Combo1.AddItem "Control Panel"
Combo1.AddItem "Recyclebin"
Combo1.AddItem "HTML document"
Combo1.AddItem "Printers"
Combo1.ListIndex = 0

txtPath.Text = Dir1.List(Dir1.ListIndex)

'Following file has changed Folders information
Open "c:\flist.dat" For Input As #1

While Not EOF(1)
Input #1, Str, Attrib, Icon
List1.AddItem Str & " [ Icon: " & Attrib & " ]"
Wend

Close #1
If List1.ListIndex <> -1 Then
List1.ListIndex = 0
End If

errnum:

If Err.Number = 53 Then
Exit Sub
End If
If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, Err.Number
End If

End Sub


Private Sub Combo1_Click()
Dim i As Integer
For i = 0 To 4
Image1(i).BorderStyle = 0
Next i
Image1(Combo1.ListIndex).BorderStyle = 1
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Str, Attrib, Icon As String
Dim i, count As Integer

On Error GoTo ErrHnd:
Close #1

'Initialize Icon variable
Select Case Combo1.ListIndex
Case 0: Icon = NetWork
Case 1: Icon = Control_P
Case 2: Icon = RBin
Case 3: Icon = HTMLDoc
Case 4: Icon = Printer
End Select

Select Case Index

'If "Add Directory" button clicked
Case 0:

'Save Folder path, It's look and Extension
Open "c:\flist.dat" For Append As #1
Write #1, txtPath.Text, Combo1.List(Combo1.ListIndex), Icon
Close #1

'Rename Folder
Name txtPath.Text As txtPath.Text & Icon

List1.Clear

'Call "Form_Load" Event
Form_Load

'If "Remove Protection" button clicked
Case 1:

'Get the selected item's Index form the List Box
count = List1.ListIndex

'Open "flist.dat" for information
Open "c:\flist.dat" For Input As #1

i = 0
While Not EOF(1)
Input #1, Str, Attrib, Icon

'"flist.dat" and ListBox maintain same serial of the item
If i = count Then
Name Str & Icon As Str
Close #1
GoTo next_s
End If
i = i + 1
Wend

next_s:
Close #1

List1.Clear
Form_Load

'Save new List
Open "c:\flist.dat" For Input As #1
Open "c:\flist.tmp" For Append As #2
i = 0
While Not EOF(1)

If i = count Then
Input #1, Str, Attrib, Icon
GoTo 2:
End If
Input #1, Str, Attrib, Icon
Write #2, Str, Attrib, Icon
i = i + 1
Wend

2:
Close #1
Close #2

Kill "C:\flist.dat"
Name "C:\flist.tmp" As "C:\flist.dat"
List1.Clear
Form_Load
End Select

ErrHnd:
If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, Err.Number
End If
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtPath.Text = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

