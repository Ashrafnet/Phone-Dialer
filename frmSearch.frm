VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSearch 
   Caption         =   "‰ «∆Ã «·»ÕÀ"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   10215
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   380
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "«ﬂ » Â‰« Ã“¡ „‰ «·«”„ «·–Ì  »ÕÀ ⁄‰Â"
      Top             =   1320
      Width           =   5895
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   6855
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12091
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "„"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«·«”„"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«·—ﬁ„"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«·⁄„·"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   120
      Picture         =   "frmSearch.frx":0000
      Top             =   0
      Width           =   1050
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ Â‰« Ì „ ⁄—÷ ‰ «∆Ã «·»ÕÀ"
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
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   5025
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "⁄—÷ ‰ «∆Ã «·»ÕÀ"
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
      Left            =   5505
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   3120
      Y1              =   960
      Y2              =   480
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   240
      X2              =   1800
      Y1              =   360
      Y2              =   0
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   8
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7575
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSearch.frx":33CD
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This flag is set when the user chooses Cancel.
Dim CancelFlag


Private Sub Form_Load()
IsfrmMainLoaded = True
    
    LVFullRowSelect Me.ListView1
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ListView1.Width = Me.ScaleWidth - 250
    Text1.Width = ListView1.Width
    ListView1.Height = Height - 2400

    OptemizeTitle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IsfrmMainLoaded = False
End Sub
Sub OptemizeTitle()
        Dim i As Integer
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
    lblInfo(3).Left = Me.Width - Me.lblInfo(3).Width - 500

End Sub

Private Sub helps_Click(Index As Integer)
    Select Case Index
        Case 0
            MsgBox " „  »—„Ã… «·»—‰«„Ã ⁄‰ ÿ—Ìﬁ:" & vbNewLine & "   „.«‘—› ﬂ„«· «·ﬁ’«’ " & vbNewLine & "AshrafNet4u@HotMail.Com" & vbNewLine & "Aqssass@Ccast.Edu.Ps", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading
            
        Case 1
        
    End Select
End Sub



Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count < 1 Then Exit Sub
    frmContactInfo.Caption = ListView1.SelectedItem.SubItems(1)
    frmContactInfo.LoadUserInfo Mid(ListView1.SelectedItem.key, 2)
    frmContactInfo.LoadInfo Mid(ListView1.SelectedItem.key, 2), -1
    frmContactInfo.Show 1
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If ListView1.ListItems.Count < 1 Then
       'pub2s(1).Enabled = False
       'pub2s(3).Enabled = False
    Else
       ' pub2s(1).Enabled = 1
       'pub2s(3).Enabled = 1

    End If
    If Button = 2 Then
       ' Me.PopupMenu Me.pub2
    End If
End Sub

Private Sub mnuoptions_Click(Index As Integer)
        Select Case Index
        Case 0
            BringWindowToTop frmCall.Hwnd
            frmCall.Show
        Case 1
            
        Case 2
            frmOptions.Show 1
    End Select

End Sub

Private Sub MSComm1_OnComm()
    MsgBox ""
End Sub


Private Sub mnus_Click(Index As Integer)
    Select Case Index
        Case 0
        
        Case 5
            End
    End Select
End Sub

Private Sub mnus1_Click(Index As Integer)
    Select Case Index
        Case 0
            pop1s_Click 0
        Case 1
            pub2s_Click 0
    End Select
End Sub

Private Sub pop1s_Click(Index As Integer)
''    Select Case Index
''        Case 0
''            frmAddG.Show 1
''            frmAddG.Gid = 0
''
''        Case 1
''            frmAddG.cmdOrder(0).Caption = " ⁄œÌ·"
''            frmAddG.Gid = List1.ItemData(List1.ListIndex)
''            frmAddG.Text1 = List1.Text
''            frmAddG.Show 1
''        Case 3
''            DelGroup List1.ItemData(List1.ListIndex)
''    End Select
End Sub
Sub DelGroup(Gid As Long)
    On Error GoTo er:
    Dim rs As New Recordset
    If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ Õ–› «·„Ã„Ê⁄… »ﬂ«„·Â«" & vbNewLine & "„·«ÕŸ… " & vbNewLine & "           ”Ì „ Õ–› Ã„Ì⁄ ÃÂ«  «·« ’«· «·„‰ ”»… «·Ï Â–Â «·„Ã„Ê⁄…", vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation) = vbNo Then Exit Sub
    rs.Open "delete * from Groups where id=" & Gid, db
    MsgBox " „  ⁄„·Ì… Õ–› «·„Ã„Ê⁄… »‰Ã«Õ", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading
    Exit Sub
er:
    MsgBox Err.Description
End Sub
Sub DelContact(cID As Long)
    On Error GoTo er:
    Dim rs As New Recordset
    If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ Õ–› Â–Â «·ÃÂ…" & vbNewLine & " " & vbNewLine & "           ", vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation) = vbNo Then Exit Sub
    rs.Open "delete * from contact where id=" & cID, db
    MsgBox " „  ⁄„·Ì… «·Õ–›  »‰Ã«Õ", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading
    Exit Sub
er:
    MsgBox Err.Description
End Sub

Sub LoadContacts(Optional strToSearch As String = "")
    On Error GoTo er:
    Dim rs As New Recordset
    Dim i As Integer
    rs.Open "select * from contact where  name like'%" & strToSearch & "%'", db
    ListView1.ListItems.Clear
    i = 1
    While Not rs.EOF
        Dim xx As ListItem
        Set xx = ListView1.ListItems.Add(, "A" & rs!ID, i, 1, 1)
        xx.SubItems(1) = rs!Name
        xx.SubItems(2) = rs!phone
        xx.SubItems(3) = rs!myWork
        rs.MoveNext
        i = i + 1
    Wend
    Exit Sub
er:
    MsgBox Err.Description
End Sub

Private Sub pub2s_Click(Index As Integer)
''On Error Resume Next
''    Select Case Index
''        Case 0
''            frmAddContact.Combo1.Text = List1.Text
''            frmAddContact.Show 1
''            LoadContacts List1.ItemData(List1.ListIndex)
''        Case 1
''            frmAddContact.cmdOrder(0).Caption = " ⁄œÌ·"
''            frmAddContact.Text1(0) = ListView1.SelectedItem.SubItems(2) ' no
''            frmAddContact.Text1(1) = ListView1.SelectedItem.SubItems(1)  ' name
''            frmAddContact.Text1(2) = ListView1.SelectedItem.SubItems(3)  ' work
''            frmAddContact.Combo1.Text = List1.Text
''            frmAddContact.cID = Mid(ListView1.SelectedItem.key, 2)
''            frmAddContact.Show 1
''            LoadContacts List1.ItemData(List1.ListIndex)
''        Case 3
''            DelContact Mid(ListView1.SelectedItem.key, 2)
''            LoadContacts List1.ItemData(List1.ListIndex)
''    End Select

End Sub



Private Sub Text1_Change()
    LoadContacts Text1
End Sub

