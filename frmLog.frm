VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLog 
   Caption         =   "⁄—÷  ⁄ﬁ» «·—”«∆·"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   12930
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Technic"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   1
      Left            =   720
      MaxLength       =   12
      MousePointer    =   1  'Arrow
      TabIndex        =   35
      ToolTipText     =   "⁄—÷ «·«“—«— «· Ì ﬁ«„ »‰ﬁ—Â« «À‰«¡ «·„ﬂ«·„… «·Â« ›Ì…"
      Top             =   8520
      Width           =   3495
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   34
      ToolTipText     =   "„⁄«Ì‰… «·„·› «·’Ê Ì"
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   33
      ToolTipText     =   "„⁄«Ì‰… «·„·› «·’Ê Ì"
      Top             =   7560
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   31
      ToolTipText     =   "«·„·›«  «·’Ê Ì… «· Ì  „  ”ÃÌ·Â« „‰ ÃÂ… «·« ’«· «·„Õœœ…"
      Top             =   7560
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   8400
      RightToLeft     =   -1  'True
      ScaleHeight     =   5835
      ScaleWidth      =   3915
      TabIndex        =   12
      Top             =   1560
      Width           =   3975
      Begin VB.CommandButton Command2 
         Caption         =   "»ÕÀ"
         Height          =   330
         Index           =   0
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   5400
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "«ﬂ » Ã“¡ „‰ «”„ «·„—”· «·ÌÂ"
         Top             =   5400
         Width           =   2175
      End
      Begin VB.OptionButton opnshow 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ Õ”» ‰ ÌÃ… «·»ÕÀ ›Ì «”„ «·„—”· «·ÌÂ"
         Height          =   255
         Index           =   7
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Õœœ Â–« «·ŒÌ«— · »ÕÀ ⁄‰ «·„—”· «·ÌÂ"
         Top             =   5040
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "⁄—÷"
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   4560
         Width           =   495
      End
      Begin VB.OptionButton opnshow 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ Õ”» «· «—ÌŒ ﬂ„« Ì·Ì:"
         Height          =   255
         Index           =   6
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   3480
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1920
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1080
         Width           =   2895
      End
      Begin VB.OptionButton opnshow 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ —”«∆· «·‘Â— «·”«»ﬁ ›ﬁÿ"
         Height          =   255
         Index           =   5
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   3120
         Width           =   2535
      End
      Begin VB.OptionButton opnshow 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ —”«∆· «·‘Â— «·Õ«·Ì ›ﬁÿ"
         Height          =   255
         Index           =   4
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2760
         Width           =   2535
      End
      Begin VB.OptionButton opnshow 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ —”«∆· «·ÌÊ„ ›ﬁÿ"
         Height          =   255
         Index           =   3
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton opnshow 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ Õ”» ÃÂ«“ «·ﬂ„»ÌÊ — «·–Ì «—”· „‰ Œ·«·Â"
         Height          =   255
         Index           =   2
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Â‰« Ì „ ⁄—÷ Ã„Ì⁄ «·—”«∆· «· Ì «—”·  „‰ Œ·«· ÃÂ«“ ﬂ„»ÌÊ — „⁄Ì‰"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.OptionButton opnshow 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ Õ”» «·„—”·"
         Height          =   255
         Index           =   1
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton opnshow 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ «·Ã„Ì⁄"
         Height          =   255
         Index           =   0
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   38880
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   4200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   38880
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   195
         Index           =   1
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‰"
         Height          =   195
         Index           =   0
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   3840
         Width           =   180
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Œ — ÿ—Ìﬁ… «·⁄—÷"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   55
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   120
         Width           =   1425
      End
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6600
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   380
      Index           =   2
      Left            =   720
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   6120
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   380
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5640
      Width           =   3495
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7011
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·«“—«— «· Ì ﬁ«„ »‰ﬁ—Â«"
      Height          =   195
      Index           =   4
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   9000
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„·›«  «·’Ê Ì… «· Ì «” ﬁ»· "
      Height          =   195
      Index           =   1
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   8160
      Width           =   1950
   End
   Begin VB.Label lblNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·—”«∆· ›Ì «·ﬁ«∆„…"
      Height          =   255
      Index           =   1
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·ÃÂ«  ›Ì «·ﬁ«∆„…"
      Height          =   255
      Index           =   0
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   2535
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   7920
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
            Picture         =   "frmLog.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„·›«  «·’Ê Ì… «· Ì «—”· "
      Height          =   195
      Index           =   3
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   7200
      Width           =   1830
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ Ê ÊﬁÌ  «·«—”«·"
      Height          =   195
      Index           =   2
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   6240
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„—”·"
      Height          =   195
      Index           =   0
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "⁄—÷  ⁄ﬁ» «·—”«∆·"
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
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   1830
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLog.frx":0352
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
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   6105
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "frmLog.frx":03EE
      Top             =   120
      Width           =   960
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   10680
      X2              =   12240
      Y1              =   720
      Y2              =   360
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   10320
      X2              =   13440
      Y1              =   960
      Y2              =   480
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   8
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnus 
         Caption         =   "⁄—÷ »ÿ«ﬁ…  ⁄—Ì›Ì…"
         Index           =   0
      End
      Begin VB.Menu mnus 
         Caption         =   "«⁄«œ… « ’«·..."
         Index           =   1
      End
   End
   Begin VB.Menu file 
      Caption         =   "„·›"
      Begin VB.Menu files 
         Caption         =   "≈€·«ﬁ"
         Index           =   0
      End
   End
   Begin VB.Menu help 
      Caption         =   "„”«⁄œ…"
      Begin VB.Menu helps 
         Caption         =   "«· ⁄·Ì„«  Ê«·œ⁄„"
         Index           =   0
      End
      Begin VB.Menu helps 
         Caption         =   "ÕÊ· «·‰Ÿ«„"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPreview_Click(Index As Integer)
Select Case Index
   Case 0
    If List2.Text <> "" Then
        sndPlaySound List2.Text, 1
    Else
        MsgBox "«Œ — «·„·› „‰ «·ﬁ«∆„… «Ê·«", vbCritical
    End If
Case 1
        If List1.Text <> "" Then
        sndPlaySound List1.Text, 1
    Else
        MsgBox "«Œ — «·„·› „‰ «·ﬁ«∆„… «Ê·«", vbCritical
    End If

End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
    Select Case Index
        Case 0
            opnshow(1).Value = True
            ShowˆLogsAsSender Combo1(0).Text
        Case 1
            opnshow(2).Value = True
            ShowˆLogsAsComputer Combo1(1).Text
    End Select

End Sub

Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass
    opnshow(6).Value = True
    ShowˆLogsAsDate DTPicker1(0).Value, DTPicker1(1).Value
    Screen.MousePointer = vbDefault
End Sub



Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            Screen.MousePointer = vbHourglass
            ShowˆLogsAsÚSearch Text2
            opnshow(7).Value = True
            Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub DTPicker1_CallbackKeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    opnshow(6).Value = True
End Sub

Private Sub DTPicker1_Change(Index As Integer)
opnshow(6).Value = True
End Sub

Private Sub DTPicker1_Click(Index As Integer)
opnshow(6).Value = True
End Sub

Private Sub files_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            
    End Select
End Sub

Private Sub Form_Load()
    OptemizeTitle
    ListView1.ColumnHeaders. _
    Add , , "«·„—”·"
    ListView1.ColumnHeaders. _
    Add , , "—ﬁ„ «·„—”· «·ÌÂ"
    ListView1.ColumnHeaders. _
    Add , , "«”„ «·„—”· «·ÌÂ", ListView1.Width / 3
    ListView1.ColumnHeaders. _
    Add , , " «—ÌŒ Ê ÊﬁÌ  «·«—”«·", ListView1.Width / 3
'    ListView1.ColumnHeaders. _
    Add , , "«· ÊﬁÌ ", ListView1.Width / 3
    ListView1.View = lvwReport
    LVFullRowSelect Me.ListView1
    ShowLogs
    opnshow(3).Value = True
    DTPicker1(0).Value = Date
    DTPicker1(1).Value = Date
    Picture1.BorderStyle = 0
    ShowAllComputers
    ShowAllSender

End Sub
Function ShowAllSender()
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select distinct log_user  from log", db, adOpenDynamic, adLockBatchOptimistic
    Combo1(0).Clear
    While Not rs.EOF
        Combo1(0).AddItem rs.Fields("log_user")
        rs.MoveNext
    Wend
    Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    'Resume
End Function
Function ShowAllComputers()
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim x As String
    
    rs.Open "select distinct compname  from log ", db, adOpenDynamic, adLockBatchOptimistic
    Combo1(1).Clear
    While Not rs.EOF
        
            x = rs.Fields("compname")
            Combo1(1).AddItem x
        
        rs.MoveNext
    Wend
    Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    'Resume
End Function

Function ShowLogInfo(UserName As String, ID As String)
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "Select * from " & "log  where log_user='" & UserName & "' and id=" & CInt(ID) & "", db, adOpenDynamic, adLockBatchOptimistic
    
    Text1(0) = rs.Fields("log_user")
    Text1(2) = rs.Fields("thedate")
    Dim recorded() As String
    recorded = Split(rs.Fields("wavrecord") & "", ",", , vbTextCompare)
    List1.Clear
    For Each xx In recorded
        List1.AddItem xx
    Next
    recorded = Split(rs.Fields("wavsend") & "", ",", , vbTextCompare)
    List2.Clear
    For Each xx In recorded
        If Trim(xx) <> "" Then
            List2.AddItem xx
        End If
    Next
    Text1(1) = rs.Fields("keys") & ""
    rs.Close
    Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    'Resume
End Function

Function ShowLogs()
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rs1 As New Recordset
    
    rs.Open "Select * from " & "log ", db, adOpenDynamic, adLockBatchOptimistic
    ListView1.ListItems.Clear
    Dim xx As Long
    Dim itmX As ListItem
    While Not rs.EOF
        
        Set itmX = ListView1.ListItems. _
        Add(, , CStr(rs.Fields("log_user")), 1, 1)
        If InStr(1, rs.Fields("thephone"), ",") Then
            Dim No() As String
            No = Split(rs.Fields("thephone"), ",", , vbTextCompare)
            For Each y In No()
                xx = xx + 1
            Next

            itmX.SubItems(1) = "„Ã„Ê⁄…"
            itmX.SubItems(2) = "„Ã„Ê⁄… «‘Œ«’"
        Else
           itmX.SubItems(1) = CStr(rs.Fields("thephone"))
           rs1.Open "select * from contact  where phone='" & rs.Fields("thephone") & "'", db, adOpenStatic
           On Error Resume Next
           If rs1.RecordCount = 0 Then
                itmX.SubItems(2) = CStr("„ÃÂÊ·" & "")
            Else
                itmX.SubItems(2) = CStr(rs1.Fields("name") & "")
            End If
           xx = xx + 1
           rs1.Close
        End If
        itmX.SubItems(3) = rs!Thedate
        itmX.key = "A" & CStr(rs.Fields("id"))
        rs.MoveNext
    Wend
    lblNo(0) = "⁄œœ «·ÃÂ«  ›Ì «·ﬁ«∆„… = " & ListView1.ListItems.Count & " ÃÂ… "
    lblNo(1) = "⁄œœ «·—”«∆· ›Ì «·ﬁ«∆„… = " & xx & " —”«·… "
        Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    Resume
End Function

Sub OptemizeTitle()
        
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

Function FindMobile(MobileNo As String) As Boolean
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim x As String
    MobileNo = Trim(MobileNo)
    MobileNo = Replace(MobileNo, "-", "")
    MobileNo = Left(MobileNo, InStr(1, MobileNo, ";") - 2)
    rs.Open "Select * from " & "user_sms  where umobile='" & MobileNo & "'", db, adOpenStatic
    While Not rs.EOF
        x = x & rs.Fields("gname") & vbNewLine
        rs.MoveNext
    Wend
    rs.MoveFirst
    If rs.RecordCount > 0 Then
        MsgBox "«·„Ã‹‹‹„‹‹‹‹‹Ê⁄‹‹‹‹«  «· Ì Ì‰ „Ì ·Â« : " & vbNewLine & x & vbNewLine & "«·«”‹‹‹‹‹‹‹‹„ : " & "" & rs.Fields("uname") & vbNewLine & " «·„Õ‹‹„‹‹‹‹‹Ê· : " & rs.Fields("umobile"), vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading
    Else
        MsgBox "Â–« «·—ﬁ„ €Ì— „ÊÃÊœ ⁄‰œ‰« ÷„‰ «·„Ã„Ê⁄« ", vbCritical
    End If
    Exit Function
er:
    'MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    'Resume
End Function

Private Sub Form_Resize()
    If Me.Width < 10380 Then Me.Width = 10380: Exit Sub
    OptemizeTitle
    ListView1.Width = Me.Width - Picture1.Width - 220
    Picture1.Left = ListView1.Width + 120
    ListView1.Left = 120
'    List1.Left = 120
'    List2.Left = 120
'    Text1(0).Left = 120
'    Text1(2).Left = 120
'    lblNo(0).Top = ListView1.Height + ListView1.Top + 120
'    lblNo(1).Top = ListView1.Height + ListView1.Top + 120
End Sub

Private Sub helps_Click(Index As Integer)
    Select Case Index
        Case 0
            ShowHelp
        Case 1
            AboutSystem
    End Select
End Sub

Private Sub List1_DblClick()
'    Dim x As String
'    x = Mid(List1.Text, 5)
'    FindMobile x
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    With ListView1
        .SortKey = ColumnHeader.Index - 1
        .Sorted = True
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With

End Sub

Private Sub ListView1_DblClick()
    mnus_Click 0
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    ShowLogInfo Item.Text, Mid(Item.key, 2)
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo er:
    If Button = 2 Then
        If ListView1.SelectedItem.Selected = True Then
            Me.PopupMenu mnu, , , , mnus(0)
        End If
    End If
    Exit Sub
er:
End Sub

Private Sub mnus_Click(Index As Integer)
On Error GoTo er:
Dim rs1 As New Recordset
    Select Case Index
        Case 0
            rs1.Open "select * from contact  where phone='" & ListView1.SelectedItem.SubItems(1) & "'", db, adOpenStatic
            If rs1.RecordCount = 0 Then
                MsgBox "Â–Â «·ÃÂ…  €Ì— „ÊÃÊœ… ÷„‰ ﬁ«⁄œ… «·»Ì«‰« ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
                Exit Sub
            End If
            frmContactInfo.LoadInfo rs1.Fields("id"), Mid(ListView1.SelectedItem.key, 2)
            frmContactInfo.Caption = ListView1.SelectedItem.SubItems(2)
            frmContactInfo.LoadUserInfo rs1.Fields("id")
            frmContactInfo.Show 1
        Case 1
            frmSendOne.Text1 = ListView1.SelectedItem.SubItems(1) ' no
            frmSendOne.Show 1
    End Select
    Exit Sub
er:
End Sub

Private Sub opnshow_Click(Index As Integer)
Screen.MousePointer = vbHourglass
    Select Case Index
        Case 0
            ShowˆLogsAsDate
        Case 1
            If Combo1(0).Text <> "" Then
                Combo1_Click 0
            Else
                Combo1(0).Text = Combo1(0).List(0)
            End If
        Case 2
            If Combo1(1).Text <> "" Then
                Combo1_Click 1
            Else
                Combo1(1).Text = Combo1(1).List(0)
            End If

        Case 3 ' today show
            ShowˆLogsAsDate Date, Date
        Case 4 ' Current Month
            Dim date1 As String
            Dim date2 As String
            date1 = "1-" & Month(Date) & "-" & Year(Date)
            date2 = MaxDayInMonth(Month(Date), Year(Date)) & "-" & Month(Date) & "-" & Year(Date)
            ShowˆLogsAsDate CDate(date1), CDate(date2)
        Case 5 ' Last Month
            If Month(Date) <> 1 Then
                date1 = "1-" & Month(Date) - 1 & "-" & Year(Date)
                date2 = MaxDayInMonth(Month(Date) - 1) & "-" & Month(Date) - 1 & "-" & Year(Date)
                ShowˆLogsAsDate CDate(date1), CDate(date2)
            Else
                date1 = "1-" & 12 & "-" & Year(Date) - 1
                date2 = MaxDayInMonth(Month(Date), Year(Date)) & "-" & 12 & "-" & Year(Date) - 1
                ShowˆLogsAsDate CDate(date1), CDate(date2)
            End If
        Case 7
            Text2.SetFocus
    End Select
Screen.MousePointer = vbDefault
End Sub

Function ShowˆLogsAsSender(SenderName As String)
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rs1 As New Recordset
    rs.Open "Select * from " & "log where (log_user) ='" & (SenderName) & "'", db, adOpenDynamic, adLockBatchOptimistic

    ListView1.ListItems.Clear
    Dim xx As Long
    Dim itmX As ListItem
    While Not rs.EOF
        
        Set itmX = ListView1.ListItems. _
        Add(, , CStr(rs.Fields("log_user")), 1, 1)
        If InStr(1, rs.Fields("thephone"), ",") Then
            Dim No() As String
            No = Split(rs.Fields("thephone"), ",", , vbTextCompare)
            For Each y In No()
                xx = xx + 1
            Next

            itmX.SubItems(1) = "„Ã„Ê⁄…"
            itmX.SubItems(2) = "„Ã„Ê⁄… «‘Œ«’"
        Else
           itmX.SubItems(1) = CStr(rs.Fields("thephone"))
           rs1.Open "select * from contact  where phone='" & rs.Fields("thephone") & "'", db, adOpenStatic
           On Error Resume Next
           If rs1.RecordCount = 0 Then
                itmX.SubItems(2) = CStr("„ÃÂÊ·" & "")
            Else
                itmX.SubItems(2) = CStr(rs1.Fields("name") & "")
            End If
           xx = xx + 1
           rs1.Close
        End If
        itmX.SubItems(3) = rs!Thedate
        itmX.key = "A" & CStr(rs.Fields("id"))
        rs.MoveNext
    Wend
    lblNo(0) = "⁄œœ «·ÃÂ«  ›Ì «·ﬁ«∆„… = " & ListView1.ListItems.Count & " ÃÂ… "
    lblNo(1) = "⁄œœ «·—”«∆· ›Ì «·ﬁ«∆„… = " & xx & " —”«·… "
        Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
End Function
Function ShowˆLogsAsComputer(ComputerName As String)
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rs1 As New Recordset
    rs.Open "Select * from " & "log where compname ='" & ComputerName & "'", db, adOpenDynamic, adLockBatchOptimistic
    ListView1.ListItems.Clear
    Dim xx As Long
    Dim itmX As ListItem
    While Not rs.EOF
        
        Set itmX = ListView1.ListItems. _
        Add(, , CStr(rs.Fields("log_user")), 1, 1)
        If InStr(1, rs.Fields("thephone"), ",") Then
            Dim No() As String
            No = Split(rs.Fields("thephone"), ",", , vbTextCompare)
            For Each y In No()
                xx = xx + 1
            Next

            itmX.SubItems(1) = "„Ã„Ê⁄…"
            itmX.SubItems(2) = "„Ã„Ê⁄… «‘Œ«’"
        Else
           itmX.SubItems(1) = CStr(rs.Fields("thephone"))
           rs1.Open "select * from contact  where phone='" & rs.Fields("thephone") & "'", db, adOpenStatic
           On Error Resume Next
           If rs1.RecordCount = 0 Then
                itmX.SubItems(2) = CStr("„ÃÂÊ·" & "")
            Else
                itmX.SubItems(2) = CStr(rs1.Fields("name") & "")
            End If
           xx = xx + 1
           rs1.Close
        End If
        itmX.SubItems(3) = rs!Thedate
        itmX.key = "A" & CStr(rs.Fields("id"))
        rs.MoveNext
    Wend
    lblNo(0) = "⁄œœ «·ÃÂ«  ›Ì «·ﬁ«∆„… = " & ListView1.ListItems.Count & " ÃÂ… "
    lblNo(1) = "⁄œœ «·—”«∆· ›Ì «·ﬁ«∆„… = " & xx & " —”«·… "
        Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    Resume
End Function
Function ShowˆLogsAsÚSearch(str As String)
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rs1 As New Recordset
    rs.Open "Select * from " & "log where (thename) like '%" & (str) & "%'", db, adOpenDynamic, adLockBatchOptimistic

    ListView1.ListItems.Clear
    Dim xx As Long
    Dim itmX As ListItem
    While Not rs.EOF
        
        Set itmX = ListView1.ListItems. _
        Add(, , CStr(rs.Fields("log_user")), 1, 1)
        If InStr(1, rs.Fields("thephone"), ",") Then
            Dim No() As String
            No = Split(rs.Fields("thephone"), ",", , vbTextCompare)
            For Each y In No()
                xx = xx + 1
            Next

            itmX.SubItems(1) = "„Ã„Ê⁄…"
            itmX.SubItems(2) = "„Ã„Ê⁄… «‘Œ«’"
        Else
           itmX.SubItems(1) = CStr(rs.Fields("thephone"))
           rs1.Open "select * from contact  where phone='" & rs.Fields("thephone") & "'", db, adOpenStatic
           On Error Resume Next
           If rs1.RecordCount = 0 Then
                itmX.SubItems(2) = CStr("„ÃÂÊ·" & "")
            Else
                itmX.SubItems(2) = CStr(rs1.Fields("name") & "")
            End If
           xx = xx + 1
           rs1.Close
        End If
        itmX.SubItems(3) = rs!Thedate
        itmX.key = "A" & CStr(rs.Fields("id"))
        rs.MoveNext
    Wend
    lblNo(0) = "⁄œœ «·ÃÂ«  ›Ì «·ﬁ«∆„… = " & ListView1.ListItems.Count & " ÃÂ… "
    lblNo(1) = "⁄œœ «·—”«∆· ›Ì «·ﬁ«∆„… = " & xx & " —”«·… "
    Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    Resume
End Function

Function ShowˆLogsAsDate(Optional StartDate As Date, Optional EndDate As Date)
On Error GoTo er:
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rs1 As New Recordset
    If StartDate <> Null And EndDate <> Null Then
        StartDate = Format(StartDate, "dd/mm/yyyy")
        EndDate = Format(EndDate, "dd/mm/yyyy")
        rs.Open "Select * from " & "log where (thedate) between  (" & StartDate & ") and  (" & EndDate & ")", db, adOpenDynamic, adLockBatchOptimistic
    ElseIf StartDate <> Null Then
        StartDate = Format(StartDate, "dd/mm/yyyy")
        rs.Open "Select * from " & "log where thedate= (" & StartDate & ")", db, adOpenDynamic, adLockBatchOptimistic
    ElseIf EndDate <> Null Then
        EndDate = Format(EndDate, "dd/mm/yyyy")
        rs.Open "Select * from " & "log where (thedate) ('" & EndDate & "','dd/mm/yyyy')", db, adOpenDynamic, adLockBatchOptimistic
    Else
        rs.Open "Select * from " & "log ", db, adOpenDynamic, adLockBatchOptimistic
    End If

    ListView1.ListItems.Clear
    Dim xx As Long
    Dim itmX As ListItem
    While Not rs.EOF
        
        Set itmX = ListView1.ListItems. _
        Add(, , CStr(rs.Fields("log_user")), 1, 1)
        If InStr(1, rs.Fields("thephone"), ",") Then
            Dim No() As String
            No = Split(rs.Fields("thephone"), ",", , vbTextCompare)
            For Each y In No()
                xx = xx + 1
            Next

            itmX.SubItems(1) = "„Ã„Ê⁄…"
            itmX.SubItems(2) = "„Ã„Ê⁄… «‘Œ«’"
        Else
           itmX.SubItems(1) = CStr(rs.Fields("thephone"))
           rs1.Open "select * from contact  where phone='" & rs.Fields("thephone") & "'", db, adOpenStatic
           On Error Resume Next
           If rs1.RecordCount = 0 Then
                itmX.SubItems(2) = CStr("„ÃÂÊ·" & "")
            Else
                itmX.SubItems(2) = CStr(rs1.Fields("name") & "")
            End If
           xx = xx + 1
           rs1.Close
        End If
        itmX.SubItems(3) = rs!Thedate
        itmX.key = "A" & CStr(rs.Fields("id"))
        rs.MoveNext
    Wend
    lblNo(0) = "⁄œœ «·ÃÂ«  ›Ì «·ﬁ«∆„… = " & ListView1.ListItems.Count & " ÃÂ… "
    lblNo(1) = "⁄œœ «·—”«∆· ›Ì «·ﬁ«∆„… = " & xx & " —”«·… "
        Exit Function
er:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    Resume
End Function

Private Sub Text2_GotFocus()
    Command2(0).Default = True
End Sub

Private Sub Text2_LostFocus()
Command2(0).Default = 0
End Sub
