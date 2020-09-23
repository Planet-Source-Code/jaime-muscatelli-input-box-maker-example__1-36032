VERSION 5.00
Begin VB.Form FRMBOX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Box Maker v1.0 ©2002 Jaime Muscatelli"
   ClientHeight    =   4680
   ClientLeft      =   3000
   ClientTop       =   2640
   ClientWidth     =   5070
   Icon            =   "FRMBOX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5070
   Begin VB.PictureBox PicCONT 
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4995
      TabIndex        =   7
      Top             =   3360
      Width           =   5055
      Begin VB.TextBox txtSOURCE 
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "FRMBOX.frx":0442
         Top             =   120
         Width           =   4695
      End
   End
   Begin VB.Frame FRAMAIN 
      Caption         =   "Input Box Syntax"
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton CMDABOUT 
         Caption         =   "&About"
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton CMDGenerate 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton CMDPreview 
         Caption         =   "&Preview"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtDefault 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Text            =   "Default"
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox txtPROMPT 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "FRMBOX.frx":044E
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtTITLE 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Text            =   "Title"
         Top             =   480
         Width           =   3255
      End
      Begin VB.Image IMGJ 
         DragIcon        =   "FRMBOX.frx":0455
         DragMode        =   1  'Automatic
         Height          =   1200
         Left            =   3600
         Picture         =   "FRMBOX.frx":0897
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1395
      End
   End
   Begin VB.Timer TMRCaption 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label LBLINFO 
      AutoSize        =   -1  'True
      Caption         =   "Press Ctrl E to copy. Double click to select text"
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   4440
      Width           =   3300
   End
   Begin VB.Menu MNUCOPY 
      Caption         =   "MNUCOPY"
      Visible         =   0   'False
      Begin VB.Menu MNUCOPYCOPY 
         Caption         =   "MNUCOPYCOPY"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "FRMBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDABOUT_Click()
Dim InputValue As String
Dim TXTMESSAGE As String
InputValue = InputBox("What is your name?", "User Input: Name", "")
TXTMESSAGE = "Thank you " & InputValue & " " & "for using Input Box Maker. This was made in VB6, in about 4 hours. This is ©2002 Jaime Muscatelli."
If Len(InputValue) = 0 Then
Exit Sub
ElseIf InputValue = " " Then
Exit Sub
End If
MsgBox TXTMESSAGE, vbInformation + vbOKOnly + vbSystemModal, FULLCAPTION
End Sub

Private Sub CMDGenerate_Click()
Dim TXT, TXT2 As String
TXT = "Dim InputValue as string"
TXT2 = "InputValue = InputBox("
txtSOURCE.Text = TXT & vbCrLf & _
                 TXT2 & Chr(34) & txtPROMPT.Text & Chr(34) & "," & Chr(34) & txtTITLE.Text & Chr(34) & "," & Chr(34) & txtDefault.Text & Chr(34) & ")"
txtSOURCE.SetFocus
End Sub

Private Sub CMDPreview_Click()
InputBox txtPROMPT.Text, txtTITLE.Text, txtDefault.Text
End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim NRESPONSE As Variant
Cancel = True
NRESPONSE = MsgBox("Do you wish to close " & TITLE & "?", vbQuestion + vbYesNo + vbSystemModal, FULLCAPTION)
If NRESPONSE = vbNo Then
Exit Sub
Else
End
End If
End Sub

Private Sub MNUCOPYCOPY_Click()
Clipboard.Clear
Clipboard.SetText txtSOURCE
End Sub

Private Sub TMRCaption_Timer()
Me.Caption = FULLCAPTION
End Sub
Private Sub txtDefault_DblClick()
Dim TXTLENGTH As Variant
TXTLENGTH = Len(txtDefault.Text)
txtDefault.SelStart = 0
txtDefault.SelLength = TXTLENGTH
End Sub
Private Sub txtPROMPT_DblClick()
Dim TXTLENGTH As Variant
TXTLENGTH = Len(txtPROMPT.Text)
txtPROMPT.SelStart = 0
txtPROMPT.SelLength = TXTLENGTH
End Sub

Private Sub txtSOURCE_DblClick()
Dim TXTLENGTH As Variant
TXTLENGTH = Len(txtSOURCE.Text)
txtSOURCE.SelStart = 0
txtSOURCE.SelLength = TXTLENGTH
End Sub

Private Sub txtSOURCE_GotFocus()
Dim TXTLENGTH As Variant
TXTLENGTH = Len(txtSOURCE.Text)
txtSOURCE.SelStart = 0
txtSOURCE.SelLength = TXTLENGTH
End Sub

Private Sub txtTITLE_DblClick()
Dim TXTLENGTH As Variant
TXTLENGTH = Len(txtTITLE.Text)
txtTITLE.SelStart = 0
txtTITLE.SelLength = TXTLENGTH
End Sub
