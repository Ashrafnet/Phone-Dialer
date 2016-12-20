VERSION 5.00
Begin VB.Form frmAddG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«÷«›… „Ã„Ê⁄… ÃœÌœ…"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOrder 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
      Height          =   375
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "«÷«›…"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddG.frx":0000
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
      Height          =   795
      Index           =   0
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   4185
   End
   Begin VB.Image Image1 
      Height          =   660
      Index           =   1
      Left            =   240
      Picture         =   "frmAddG.frx":00A5
      Top             =   840
      Width           =   660
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   0
      Left            =   3960
      Picture         =   "frmAddG.frx":31DD
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«ﬂ » «”„ «·„Ã„Ê⁄… «· Ì  —Ìœ «÷«› Â« Ê„‰ À„ «‰ﬁ— ⁄·Ï “— «÷«›…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   675
      Index           =   3
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   2985
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«÷«›… „Ã„Ê⁄… ÃœÌœ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   2
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1980
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   8
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6495
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
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   7560
      X2              =   9120
      Y1              =   2760
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„Ã„Ê⁄…"
      Height          =   195
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   915
   End
End
Attribute VB_Name = "frmAddG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Gid As Long
Private Sub cmdOrder_Click(Index As Integer)
    Select Case Index
        Case 0
            addG Text1.Text
            Unload Me
        Case 1
            Unload Me
    End Select
End Sub
Sub addG(gname As String)
    On Error GoTo er:
       Dim rs As New Recordset
       rs.Open "select id from groups where id=" & Gid, db, adOpenStatic
       If rs.RecordCount > 0 Then
            rs.Close
            rs.Open "update groups set name='" & gname & "' where id=" & Gid, db
            MsgBox " „   ⁄œÌ· «·„Ã„Ê⁄… " & gname & " »‰Ã«Õ ", vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation
        Else
            rs.Close
            rs.Open "insert into groups(name)values ('" & gname & "')", db
            MsgBox " „  «÷«›… «·„Ã„Ê⁄… " & gname & " »‰Ã«Õ ", vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation
        End If
       
    Exit Sub
er:
    MsgBox Err.Description
Resume
End Sub

Private Sub Form_Load()
    OptemizeTitle
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
    lblInfo(2).Left = Me.Width - lblInfo(2).Width - 1100
    'lblInfo(0).Left = Me.Width - Me.lblInfo(0).Width - 500

End Sub

Private Sub Text1_Change()
    cmdOrder(0).Enabled = Len(Trim(Text1.Text))
End Sub
