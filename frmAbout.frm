VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‰Ÿ«„ «·—”«∆· «·’Ê Ì…"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "≈€·«ﬁ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ã„Ì⁄ «·ÕﬁÊﬁ „Õ›ÊŸ… ··„»—„Ã Ê·« ÌÕﬁ ·ﬂ «· ’—› «·« »⁄œ «·„Ê«›ﬁ… «·—”„Ì… „‰ «·„»—„Ã"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   1
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4560
      Width           =   4185
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   4800
      Picture         =   "frmAbout.frx":0000
      Top             =   4560
      Width           =   240
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   7560
      X2              =   9120
      Y1              =   2760
      Y2              =   2400
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   7320
      X2              =   10440
      Y1              =   3360
      Y2              =   2880
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ÿ«„ «·—”«∆· «·’Ê Ì… ⁄»— «·Â« ›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   3225
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   0
      Left            =   4080
      Picture         =   "frmAbout.frx":0288
      Top             =   240
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   660
      Index           =   1
      Left            =   120
      Picture         =   "frmAbout.frx":3A27
      Top             =   840
      Width           =   660
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":6B5F
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1515
      Index           =   0
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   4305
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Index           =   8
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
Sub OptemizeTitle()
    Icon = Nothing
    lblT(8).Top = 0
    lblT(8).Left = 0
    lblT(8).Width = Me.Width
    lblT(8).ZOrder 1

    For i = 0 To lin.Count - 1
        lin(i).X1 = 0
        lin(i).X2 = Me.Width
        lin(i).Y1 = lblT(8).Height + 2
        lin(i).Y2 = lblT(8).Height + 2
    Next i
    lin(0).ZOrder 1
'    lblInfo(2).Left = Me.Width - lblInfo(2).Width - 1100
    'lblInfo(0).Left = Me.Width - Me.lblInfo(0).Width - 500

End Sub

Private Sub Form_Load()
    OptemizeTitle
End Sub
