VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ’œÌ— ÃÂ«  «·« ’«·"
   ClientHeight    =   7740
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   323
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "«” ⁄—«÷"
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3960
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "«—Ìœ  ‘›Ì— «·„·› »ﬂ·„… „—Ê—"
      Height          =   255
      Index           =   4
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·„Ã„Ê⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄„·\„ﬂ«‰ «·⁄„·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·—ﬁ„"
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
      Height          =   255
      Index           =   1
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·«”„"
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
      Height          =   255
      Index           =   0
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2880
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ’œÌ—"
      Default         =   -1  'True
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
      Height          =   380
      Index           =   0
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4048
      View            =   3
      Arrange         =   1
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
      NumItems        =   0
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄«Ì‰… „« ”Ì „  ’œÌ—Â"
      Height          =   195
      Index           =   4
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4440
      Width           =   1515
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Õœœ «·«‰ „”«— «·„·› «·–Ì  —Ìœ «· ’œÌ— «·ÌÂ"
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
      Index           =   1
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3720
      Width           =   3480
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   5520
      Picture         =   "frmExport.frx":0000
      Top             =   480
      Width           =   720
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmExport.frx":0ECA
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
      Height          =   1395
      Index           =   6
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   5745
   End
   Begin VB.Image Image1 
      Height          =   660
      Index           =   1
      Left            =   120
      Picture         =   "frmExport.frx":1085
      Top             =   720
      Width           =   660
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ŒÌ«—«  «· ’œÌ—"
      Height          =   195
      Index           =   0
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   1005
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ — «·„Ã„Ê⁄…"
      Height          =   195
      Index           =   16
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ’œÌ— ÃÂ«  « ’«·"
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
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«” Œœ„ Â–Â «·«œ«… · ’œÌ— «·ÃÂ«  «·Ï „·› «·«ﬂ”· «Ê „‰ „·›«  «·”Ì ›Ì «” "
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
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   3585
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
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Index           =   8
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6495
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnus 
         Caption         =   "»ÿ«ﬁ…  ⁄—Ì›"
         Index           =   0
      End
      Begin VB.Menu mnus 
         Caption         =   "≈“«·… «·ÃÂ«  «·„Õœœ…"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
'    lblInfo(2).Left = Me.Width - lblInfo(2).Width - 300
    

End Sub
Sub LoadGroups()
    On Error GoTo er:
    Dim rs As New Recordset
    Dim i As Integer
    rs.Open "select * from groups ", db, adOpenStatic
    Combo1.Clear
    i = 1
    If rs.RecordCount < 1 Then Exit Sub
    Combo1.AddItem "Ã„Ì⁄ «·„Ã„Ê⁄« "
    Combo1.ItemData(0) = 0
    While Not rs.EOF
        Combo1.AddItem rs!Name
        Combo1.ItemData(i) = rs!ID
        rs.MoveNext
        i = i + 1
    Wend
    On Error Resume Next
    Combo1.Text = Combo1.List(0)
    Exit Sub
er:
    MsgBox Err.Description
End Sub
Sub LoadContacts(Gid As Long)
    On Error GoTo er:
    Dim rs As New Recordset
    Dim rsG As New Recordset
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    If Gid = 0 Then
        rs.Open "select * from contact", db
    Else
        rs.Open "select * from contact where gid=" & Gid, db
    End If

    ListView1.ColumnHeaders.Add , , "«·«”„", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "«·Â« ›", ListView1.Width / 3, lvwColumnCenter

    

    If Check1(2).Value = 1 And Check1(3).Value = 1 Then
        ListView1.ColumnHeaders.Add , , "«·⁄„·", ListView1.Width / 3, lvwColumnCenter
        ListView1.ColumnHeaders.Add , , "«”„ «·„Ã„Ê⁄…", ListView1.Width / 3, lvwColumnCenter
    ElseIf Check1(2).Value = 1 Then
        ListView1.ColumnHeaders.Add , , "«·⁄„·", ListView1.Width / 3, lvwColumnCenter
    ElseIf Check1(3).Value = 1 Then
        ListView1.ColumnHeaders.Add , , "«”„ «·„Ã„Ê⁄…", ListView1.Width / 3, lvwColumnCenter
    End If


On Error Resume Next
    ListView1.View = lvwReport
    While Not rs.EOF
        rsG.Open "select name from groups where id=" & rs!Gid, db
        Dim xx As ListItem
        Set xx = ListView1.ListItems.Add(, "A" & rs!ID, rs!Name, 1, 1)
        xx.SubItems(1) = rs!phone
        If Check1(2).Value = 1 And Check1(3).Value = 1 Then
            xx.SubItems(2) = rs!myWork
            xx.SubItems(3) = rsG!Name
        ElseIf Check1(2).Value = 1 Then
            xx.SubItems(2) = rs!myWork
        ElseIf Check1(3).Value = 1 Then
            xx.SubItems(2) = rsG!Name
        End If

        
        rs.MoveNext
        rsG.Close
    Wend
    Exit Sub
er:
    MsgBox Err.Description
    Resume
End Sub

Private Sub Check1_Click(Index As Integer)
    If Index = 4 Then
        If Check1(4).Value Then
            txtPass(0).Enabled = True
            txtPass(0).BackColor = vbWhite
        Else
            txtPass(0).Enabled = 0
            txtPass(0).BackColor = vbButtonFace
        End If
        Exit Sub
    End If
    Combo1_Click
End Sub

Private Sub Combo1_Click()
    If Combo1.Text <> "" Then
        LoadContacts Combo1.ItemData(Combo1.ListIndex)
        Command1(0).Enabled = True
    Else
        Command1(0).Enabled = 0
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            If ListView1.ListItems.Count < 1 Then MsgBox "·« ÌÊÃœ ⁄‰«’— ›Ì «·ﬁ«∆„…", vbCritical: Exit Sub
            If Trim(Text1) = "" Then MsgBox "Õœœ „”«— «·Õ›Ÿ «Ê·«", vbCritical: Exit Sub
            Export
        Case 1
           Unload Me
        Case 2
            
    End Select
End Sub
Sub Export()
    On Error GoTo er:

        MousePointer = vbHourglass
        Dim xlApp As Excel.Application
        Dim xlWB As Excel.Workbook
        Dim xlWS As Excel.Worksheet
    
        Set xlApp = New Excel.Application
        Set xlWB = xlApp.Workbooks.Add
        Set xlWS = xlWB.Worksheets.Add
    
        
        If Check1(0).Value Then
            For i = 1 To ListView1.ListItems.Count
                xlWS.Cells(i, 1).Value = ListView1.ListItems(i).Text
                
            Next i
        End If
        If Check1(1).Value Then
            For i = 1 To ListView1.ListItems.Count
                xlWS.Cells(i, 2).Value = ListView1.ListItems(i).SubItems(1)
            Next i
        End If
        If Check1(2).Value Then
            For i = 1 To ListView1.ListItems.Count
                xlWS.Cells(i, 3).Value = ListView1.ListItems(i).SubItems(2)
            Next i
        End If
        If Check1(3).Value Then
            For i = 1 To ListView1.ListItems.Count
                xlWS.Cells(i, 4).Value = ListView1.ListItems(i).SubItems(3)
            Next i
        End If

        If Check1(4).Value Then
            xlWS.SaveAs Text1, , txtPass(0)
        Else
            xlWS.SaveAs Text1
        End If
        xlApp.Quit
    ' Free memory
        Set xlWS = Nothing
        Set xlWB = Nothing
        Set xlApp = Nothing
        MousePointer = vbDefault
        MsgBox " „  ⁄„·Ì…  ’œÌ— «·»Ì«‰«  »‰Ã«Õ ", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading
Exit Sub
er:
    MousePointer = vbDefault
    If Err.Number = 1004 Then Exit Sub
    
    MsgBox "Ì»œÊ «‰ »—‰«„Ã «·«ﬂ”· €Ì— „ Ê›— ⁄·Ï ÃÂ«“ﬂ" & vbNewLine & "„⁄·Ê„«  ›‰Ì…:" & vbNewLine & Err.Description, vbCritical + vbMsgBoxRtlReading + vbMsgBoxRight


End Sub

Private Sub Command2_Click()
    CommonDialog1.Filter = "„·› «ﬂ”· (*.xls)|*.xls"
    CommonDialog1.ShowSave
    
    If CommonDialog1.Filename <> "" Then
        On Error Resume Next
        
        Text1.Text = CommonDialog1.Filename
    End If

End Sub

Private Sub Form_Load()
    OptemizeTitle
    LoadGroups
    LVFullRowSelect Me.ListView1

End Sub

Private Sub ListView1_DblClick()
mnus_Click 0
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

    Select Case Index
        Case 0
            frmContactInfo.Caption = ListView1.SelectedItem.SubItems(2)
            frmContactInfo.LoadUserInfo Mid(ListView1.SelectedItem.key, 2)
            frmContactInfo.Show 1
        Case 1
            
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i - j).Selected = True Then
                    ListView1.ListItems.Remove ListView1.ListItems(i - j).Index
                    j = j + 1
                End If
            Next i
'            LoadContacts Combo1.ItemData(Combo1.ListIndex)
    End Select
        Exit Sub
er:

End Sub

