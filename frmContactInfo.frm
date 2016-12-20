VERSION 5.00
Begin VB.Form frmContactInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÕÊ· ÃÂ…"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   0
      Left            =   2880
      RightToLeft     =   -1  'True
      ScaleHeight     =   3225
      ScaleWidth      =   7065
      TabIndex        =   4
      Top             =   3600
      Width           =   7095
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Height          =   380
         Index           =   0
         Left            =   840
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Height          =   975
         Index           =   2
         Left            =   840
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Height          =   380
         Index           =   1
         Left            =   840
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblTit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„·\„ﬂ«‰ «·⁄„·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1920
         Width           =   1710
      End
      Begin VB.Label lblTit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblTit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·Â« ›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4740
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   1
      Left            =   3000
      RightToLeft     =   -1  'True
      ScaleHeight     =   3225
      ScaleWidth      =   4545
      TabIndex        =   15
      Top             =   3720
      Width           =   4575
      Begin VB.ListBox List2 
         Height          =   1620
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "«‰ﬁ— ‰ﬁ— Ì‰ ·„⁄«Ì‰… «·„·› «·’Ê Ì"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄»«—… ⁄‰ „Ã„Ê⁄… «·—”«∆· «·’Ê Ì… «· Ì  „ «—”«·Â« «·Ï Â–Â «·ÃÂ… ›Ì «·„ﬂ«·„… «·„Õœœ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   0
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   120
         Width           =   4065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ«∆„… «·„·›«  «·’Ê Ì…"
         Height          =   255
         Index           =   1
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   2
      Left            =   120
      RightToLeft     =   -1  'True
      ScaleHeight     =   3225
      ScaleWidth      =   4545
      TabIndex        =   22
      Top             =   2040
      Width           =   4575
      Begin VB.ListBox List3 
         Height          =   1620
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "«‰ﬁ— ‰ﬁ— Ì‰ ·„⁄«Ì‰… «·„·› «·’Ê Ì"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ«∆„… «·„·›«  «·’Ê Ì…"
         Height          =   255
         Index           =   2
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄»«—… ⁄‰ „Ã„Ê⁄… «·—”«∆· «·’Ê Ì… «· Ì  „  ”ÃÌ·Â« „‰ ÃÂ… «·« ’«· «·„Õœœ… ›Ì «·„ﬂ«·„… «·„Õœœ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   1
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   120
         Width           =   4065
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   3
      Left            =   120
      RightToLeft     =   -1  'True
      ScaleHeight     =   3225
      ScaleWidth      =   4545
      TabIndex        =   26
      Top             =   2040
      Width           =   4575
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   1215
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄»«—… ⁄‰ „Ã„Ê⁄… «·«—ﬁ«„ Ê«·«Õ—Ê› Ê«·«“—«— «· Ì ﬁ«„ ’«Õ» ÃÂ… «·« ’«· »‰ﬁ—Â« « À«¡ «·„ﬂ«·„… «·„Õœœ……"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   4
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   120
         Width           =   4065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄… «·Õ—Ê› Ê«·«—ﬁ«„ «· Ì ﬁ«„ »‰ﬁ—Â«"
         Height          =   255
         Index           =   3
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   4680
      RightToLeft     =   -1  'True
      ScaleHeight     =   2595
      ScaleWidth      =   2355
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "«Œ — «·„ﬂ«„·… „‰ «·ﬁ«∆„… · —Ï  ›«’Ì·Â«"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ Ê ÊﬁÌ  «·„ﬂ«·„…"
         Height          =   255
         Index           =   0
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "⁄«„"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Õ”‰«"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "«·—”«∆· «·„” ·„…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "«” Œœ„ Â–« «·“— ·⁄—÷ «·„·›«  «·’Ê Ì… «· Ì «—”·  «·Ï Â–Â «·ÃÂ…"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "«·„›« ÌÕ «· Ì ‰ﬁ—Â«"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "·⁄—÷ Ã„Ì⁄ «·«“—«— Ê«·„›« ÌÕ «· Ì ‰ﬁ—Â« ’«Õ» Â–Â «·ÃÂ… «À‰«¡ «·„ﬂ«·„…"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "—”«·… „”Ã·…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "·⁄—÷ „·›«  «·’Ê  «·„”Ã·… „‰ ﬁ»· Â–Â «·ÃÂ…"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   120
      Picture         =   "frmContactInfo.frx":0000
      Top             =   120
      Width           =   1860
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Â‰« Ì „ ”—œ »⁄÷ «·„⁄·Ê„«  «·Œ«’… »ÃÂ…  «·« ’«· «·„Õœœ…,ﬂ„«  —Ï Â‰«ﬂ „⁄·Ê„«  ÕÊ· «·‘Œ’ , ÊÂ‰«ﬂ „⁄·Ê„«  ÕÊ· ‰‘«ÿ«  «·‘Œ’"
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
      Height          =   795
      Index           =   3
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4065
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»ÿ«ﬁ…  ⁄—Ì›"
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
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1140
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   8
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9015
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   10440
      X2              =   13560
      Y1              =   3360
      Y2              =   2880
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   10680
      X2              =   12240
      Y1              =   2760
      Y2              =   2400
   End
End
Attribute VB_Name = "frmContactInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Log_ID As Long
Dim C_ID As Long
Private Sub cmdOrder_Click(Index As Integer)
    For i = 0 To cmdOrder.Count - 1
        cmdOrder(i).Enabled = True
        pic(i).Visible = False
    Next i
    pic(Index).Visible = 1
    cmdOrder(Index).Enabled = False
    
    pic(Index).ZOrder
    If Index = 0 Then Picture1.Visible = False: Exit Sub
    If Log_ID = -1 Then
        Picture1.Visible = True
        Picture1.ZOrder
    End If
End Sub

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
    lblInfo(2).Left = Me.Width - lblInfo(2).Width - 300
    

End Sub

Private Sub Form_Load()
'    isExplorerBar1.Align = 1
'    isExplorerBar1.Left = 6360
'    isExplorerBar1.Top = 1440
    OptemizeTitle
'    BuildMenu
    pic(0).ZOrder
    For i = 0 To pic.Count - 1
        pic(i).Left = 120
        pic(i).Top = 2040
        pic(i).BorderStyle = 0
    Next i
    Picture1.BorderStyle = 0
End Sub
Public Sub LoadInfo(ContactID As Long, LogID As Long)
    Log_ID = LogID
    LoadCallsDates ContactID
    LoadCallskeys LogID
    LoadCallsSendedFiles LogID
    LoadCallsRecordedFiles LogID
End Sub

Sub LoadCallsDates(UID As Long)
    Dim rs As New Recordset
    On Error GoTo er:
        rs.Open "select thedate,thetime,ID from log where uid=" & UID, db
        List1.Clear
        While Not rs.EOF
            List1.AddItem rs!Thedate & " " & rs!thetime
            List1.ItemData(i) = rs!ID
            i = i + 1
            rs.MoveNext
        Wend
    Exit Sub
er:
    
End Sub
Sub LoadCallskeys(LogID As Long)
    Dim rs As New Recordset
    On Error GoTo er:
        rs.Open "select keys from log where id=" & LogID, db
        Text1 = ""
        Text1 = rs!keys & ""
    Exit Sub
er:
    'MsgBox Err.Description
End Sub
Sub LoadCallsSendedFiles(LogID As Long)
    Dim rs As New Recordset
    On Error GoTo er:
        rs.Open "select wavsend from log where id=" & LogID, db
        List2.Clear
        Dim recorded() As String
        recorded = Split(rs.Fields("wavsend") & "", ",", , vbTextCompare)
        List2.Clear
        For Each xx In recorded
            List2.AddItem xx
        Next

    Exit Sub
er:
    
End Sub
Sub LoadCallsRecordedFiles(LogID As Long)
    Dim rs As New Recordset
    On Error GoTo er:
        rs.Open "select wavrecord from log where id=" & LogID, db
        List3.Clear
        Dim recorded() As String
        recorded = Split(rs.Fields("wavrecord") & "", ",", , vbTextCompare)
        List3.Clear
        For Each xx In recorded
            List3.AddItem xx
        Next

    Exit Sub
er:
    
End Sub
Sub BuildMenu()
    
    isExplorerBar1.AddSpecialGroup "«œ«—… ÃÂ«  «·« ’«·"
    isExplorerBar1.AddItem -1, "0", "«÷«›… „Ã„Ê⁄… ÃœÌœ…"
    isExplorerBar1.AddItem -1, "1", "«÷«›… ÃÂ… « ’«· ÃœÌœ…"
    isExplorerBar1.AddItem -1, "2", "»ÕÀ ⁄‰ ÃÂ… « ’«·"
    isExplorerBar1.AddItem -1, "7", "Õ–› „Ã„Ê⁄… «Ê ÃÂ… « ’«·"
    isExplorerBar1.AddItem -1, "3", "«œ«—… ÃÂ«  «·« ’«·"
End Sub
Sub LoadUserInfo(cID As Long)
        On Error GoTo er:
    C_ID = cID
    Dim rs As New Recordset
    Dim i As Integer
    rs.Open "select * from contact where id=" & cID, db
    txtInfo(0) = rs!Name
    txtInfo(1) = rs!phone
    txtInfo(2) = rs!myWork
        
    
    Exit Sub
er:
    MsgBox Err.Description

End Sub

Private Sub List1_Click()
    LoadInfo C_ID, List1.ItemData(List1.ListIndex)
    Log_ID = -1
End Sub

Private Sub List2_DblClick()
    If List2.Text <> "" Then
        sndPlaySound List2.Text, 1
    Else
        MsgBox "«Œ — «·„·› „‰ «·ﬁ«∆„… «Ê·«", vbCritical
    End If
End Sub

Private Sub List3_DblClick()
    If List3.Text <> "" Then
        sndPlaySound List3.Text, 1
    Else
        MsgBox "«Œ — «·„·› „‰ «·ﬁ«∆„… «Ê·«", vbCritical
    End If

End Sub
