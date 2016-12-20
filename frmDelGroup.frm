VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDelGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Õ–› „Ã„Ê⁄… «Ê ÃÂ… „⁄Ì‰…"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Õ–› „Ã„Ê⁄…"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4305
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Cancel          =   -1  'True
      Caption         =   "≈€·«ﬁ"
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4305
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Õ–› ÃÂ…"
      Enabled         =   0   'False
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
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4305
      Width           =   1215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4471
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDelGroup.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDelGroup.frx":0552
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "frmDelGroup.frx":0AA4
      Top             =   120
      Width           =   960
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬁ„ »«Œ Ì«— «·„Ã„Ê⁄… «Ê ÃÂ… «·« ’«· «· Ì  —Ìœ Õ–›Â« Ê„‰ À„ «‰ﬁ— ⁄·Ï «·“— «·„‰«”» ﬂ„«  —Ï »«·«”›·"
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
      Height          =   435
      Index           =   3
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   4305
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Õ–› „Ã„Ê⁄… «Ê ÃÂ… « ’«·"
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
      Left            =   3270
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   2565
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Õ–Ì— : Õ–› «·„Ã„Ê⁄… ÌƒœÌ «·Ï Õ–› Ã„Ì⁄ «·„‰ ”»Ì‰ «·Ï Â–Â «·„Ã„Ê⁄… "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   0
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄« "
      Height          =   195
      Index           =   16
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Index           =   8
      Left            =   -120
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
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÃÂ«  «·„‰ ”»… «·Ï Â–Â «·„Ã„Ê⁄…"
      Height          =   195
      Index           =   1
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   2355
   End
End
Attribute VB_Name = "frmDelGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOrder_Click(Index As Integer)
    Select Case Index
    Case 0
        If IsAdmin = False Then MsgBox "·«  „ ·ﬂ «·’·«ÕÌ« ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
        If List1.Text <> "" Then
            DelGroup (List1.ItemData(List1.ListIndex))
        Else
            MsgBox "«Œ — «Ê·« «·„Ã„Ê⁄… „‰ ﬁ«∆„… «·„Ã„Ê⁄« ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
        End If
    Case 1
        Unload Me
    Case 2
        If IsAdmin = False Then MsgBox "·«  „ ·ﬂ «·’·«ÕÌ« ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
        If ListView1.ListItems.Count < 1 Then MsgBox "·« ÌÊÃœ ⁄‰«’— ›Ì «·ﬁ«∆„…", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
        If ListView1.SelectedItem.Text <> "" Then
            DelContact Mid(ListView1.SelectedItem.key, 2)
            LoadContacts List1.ItemData(List1.ListIndex)
        Else
            MsgBox "«Œ — «Ê·« «·ÃÂ… „‰ ﬁ«∆„… «·ÃÂ« ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
        End If

End Select
    
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

Function delUser(UserName As String, GroupName As String) As Boolean
On Error GoTo er:
    If MsgBox(" Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ Õ–› «·ÃÂ… " & UserName, vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading) = vbYes Then
        Dim rs As New ADODB.Recordset
        Set rs = New ADODB.Recordset
        Dim strQ As String
        strQ = "delete user_sms WHERE uname = '" & UserName & "' and gname='" & GroupName & "'"
        rs.Open strQ, db
        delUser = True
    End If
Exit Function
er:
    delUser = False
    MsgBox Err.Description, vbCritical
End Function

Function delType(Typename As String) As Boolean
On Error GoTo er:
    If MsgBox(" Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ Õ–› «·„Ã„Ê⁄… " & Typename & vbNewLine & " Õ–Ì— : " & vbNewLine & "          »„Ã—œ «·÷€ÿ ⁄·Ï „Ê«›ﬁ ”Ì „ Õ–› «·„Ã„Ê⁄… ÊÃ„Ì⁄ «·„‰ ”»Ì‰ ·Â« ﬂ„« ÌŸÂ— ›Ì «·ﬁ«∆„… «·„Ã«Ê—… ·ﬁ«∆„… «·„Ã„Ê⁄« ", vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading) = vbYes Then
        Dim rs As New ADODB.Recordset
        Set rs = New ADODB.Recordset
        Dim strQ As String
        strQ = "delete * from group_sms WHERE gname = '" & Typename & "'"
        rs.Open strQ, db
        strQ = "delete * from user_sms WHERE gname = '" & Typename & "'"
        rs.Open strQ, db
        delType = True
    End If
Exit Function
er:
    delType = False
    MsgBox Err.Description, vbCritical
End Function

Private Sub Form_Load()
    ListView1.ColumnHeaders. _
    Add , , "«·«”„", ListView1.Width / 3
    ListView1.ColumnHeaders. _
    Add , , "«·Â« ›", ListView1.Width / 3, _
    lvwColumnCenter
    ' Set View property to Report.
    ListView1.View = lvwReport
    OptemizeTitle
On Error GoTo er:
    LoadGroups
    LVFullRowSelect Me.ListView1
    'Option1(0).Value = True
    Exit Sub
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
End Sub
Sub LoadGroups()
    On Error GoTo er:
    Dim rs As New Recordset
    Dim i As Integer
    rs.Open "select * from groups ", db
    List1.Clear
    i = 0
    While Not rs.EOF
        List1.AddItem rs!Name
        List1.ItemData(i) = rs!ID
        rs.MoveNext
        i = i + 1
    Wend
    On Error Resume Next
    List1.Text = List1.List(0)
    Exit Sub
er:
    MsgBox Err.Description
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

Function ShowAllÚGroups()
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "Select * from " & "group_sms ", db, adOpenDynamic, adLockBatchOptimistic
    List1.Clear
    While Not rs.EOF
        List1.AddItem rs.Fields("Gname")
        rs.MoveNext
    Wend
    Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    'Resume
End Function

Private Sub List1_Click()
    If List1.Text <> "" Then
        LoadContacts List1.ItemData(List1.ListIndex)
        cmdOrder(0).Enabled = True
    Else
        cmdOrder(0).Enabled = 0
    End If

End Sub
Sub LoadContacts(Gid As Long)
    On Error GoTo er:
    Dim rs As New Recordset
    
    rs.Open "select * from contact where gid=" & Gid, db
    ListView1.ListItems.Clear
    While Not rs.EOF
        Dim xx As ListItem
        Set xx = ListView1.ListItems.Add(, "A" & rs!ID, rs!Name, 1, 1)
        xx.SubItems(1) = rs!phone
        rs.MoveNext
    Wend
    Exit Sub
er:
    MsgBox Err.Description
End Sub

Sub DelGroup(Gid As Long)
    On Error GoTo er:
    Dim rs As New Recordset
    If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ Õ–› «·„Ã„Ê⁄… »ﬂ«„·Â«" & vbNewLine & "„·«ÕŸ… " & vbNewLine & "           ”Ì „ Õ–› Ã„Ì⁄ ÃÂ«  «·« ’«· «·„‰ ”»… «·Ï Â–Â «·„Ã„Ê⁄…", vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation) = vbNo Then Exit Sub
    rs.Open "delete * from Groups where id=" & Gid, db
                List1.RemoveItem List1.ListIndex
            ListView1.ListItems.Clear

    MsgBox " „  ⁄„·Ì… Õ–› «·„Ã„Ê⁄… »‰Ã«Õ", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading
    Exit Sub
er:
    MsgBox Err.Description
End Sub

Function LoadUsers(GroupName As String)
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "Select * from " & "user_sms where gname='" & GroupName & "'", db, adOpenDynamic, adLockBatchOptimistic
    ListView1.ListItems.Clear
    
    Dim itmX As ListItem

    While Not rs.EOF
        Set itmX = ListView1.ListItems. _
            Add(, , CStr(rs.Fields("uname")), 1, 1)
        If Not IsNull(umobile) Then
            itmX.SubItems(1) = CStr(rs.Fields("umobile") & "")
        End If
        rs.MoveNext
    Wend
    lblInfo(1) = "«·ÃÂ«  «·„‰ ”»… «·Ï Â–Â «·„Ã„Ê⁄… = " & ListView1.ListItems.Count & " ÃÂ… "
        Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading

    Resume
End Function


Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    If Item.Text <> "" Then
        cmdOrder(2).Enabled = True
    Else
        cmdOrder(2).Enabled = 0
    End If
End Sub


