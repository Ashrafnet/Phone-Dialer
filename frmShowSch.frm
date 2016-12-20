VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmShowSch 
   Caption         =   "≈œ«—… «·„Â«„ «·„ÃœÊ·…"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   10065
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "„"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«· «—ÌŒ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«· ÊﬁÌ "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«·Õ«·…"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«”„ «·„Â„…"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShowSch.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShowSch.frx":3052
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShowSch.frx":33A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShowSch.frx":38F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShowSch.frx":3C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShowSch.frx":3F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShowSch.frx":42EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   120
      Picture         =   "frmShowSch.frx":463E
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ì „ Â‰« ”—œ Ã„Ì⁄ «·„Â«„ «·„ÃœÊ·… «· Ì ”Ì „ «·⁄„· ⁄·ÌÂ« ›Ì Õ«· Õ«‰ «·«Ê«‰ · ‰›Ì–Â«"
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
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3585
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”—œ ·Ã„Ì⁄ «·„Â«„ «·„ÃœÊ·…"
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
      Left            =   3300
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Index           =   8
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6495
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   7800
      X2              =   10920
      Y1              =   3360
      Y2              =   2880
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   8040
      X2              =   9600
      Y1              =   2760
      Y2              =   2400
   End
   Begin VB.Menu mnu 
      Caption         =   "„‰ÌÊ"
      Visible         =   0   'False
      Begin VB.Menu mnus 
         Caption         =   " ⁄ÿÌ·"
         Index           =   0
      End
      Begin VB.Menu mnus 
         Caption         =   "Õ–›"
         Index           =   1
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnus 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnus 
         Caption         =   "«⁄«œ…  ”„Ì….."
         Index           =   3
      End
      Begin VB.Menu mnus 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnus 
         Caption         =   "Œ’«∆’"
         Index           =   5
      End
   End
   Begin VB.Menu mnu1 
      Caption         =   "mnu1"
      Visible         =   0   'False
      Begin VB.Menu mnus1 
         Caption         =   "„Â„… ÃœÌœ…..."
         Index           =   0
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "„Â„…"
      Begin VB.Menu mnufiles 
         Caption         =   "ÃœÌœ…..."
         Index           =   0
      End
      Begin VB.Menu mnufiles 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnufiles 
         Caption         =   "≈€·«ﬁ"
         Index           =   2
      End
   End
   Begin VB.Menu option 
      Caption         =   "«œ«Ê "
      Begin VB.Menu options 
         Caption         =   "„ﬂ«·„… Ã„«⁄Ì…..."
         Index           =   0
      End
      Begin VB.Menu options 
         Caption         =   "„ﬂ«·„… ›—œÌ…..."
         Index           =   1
      End
      Begin VB.Menu options 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu options 
         Caption         =   "ŒÌ«—«  «·‰Ÿ«„..."
         Index           =   3
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
Attribute VB_Name = "frmShowSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub LoadSchdulers()
    On Error GoTo er:
    Dim rs As New Recordset
    
    rs.Open "select * from schmaster ", db
    ListView1.ListItems.Clear
    i = 1
    While Not rs.EOF
        Dim xx As ListItem
        Set xx = ListView1.ListItems.Add(, "A" & rs!SchID, i, 6, 6)
        If rs!schdate = "1" Then
            xx.SubItems(1) = "ÌÊ„Ì«"
            
        Else
            xx.SubItems(1) = rs!schdate
        End If
        xx.SubItems(2) = rs!schtime
        xx.SubItems(4) = rs!schname & ""
        xx.SubItems(3) = "‰‘ÿ…"
        rs.MoveNext
        i = i + 1
    Wend
    Exit Sub
er:
    MsgBox Err.Description
End Sub

Sub OptemizeTitle()
    LVFullRowSelect Me.ListView1
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
    lblInfo(3).Left = Me.Width - lblInfo(2).Width - 1500
    

End Sub




Private Sub Form_Load()
    OptemizeTitle
    LoadSchdulers

    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    OptemizeTitle
    ListView1.Width = Width - 350
    ListView1.Height = Height - 2133
    For i = 1 To ListView1.ColumnHeaders.Count
        ListView1.ColumnHeaders(i).Width = ListView1.Width \ 5
    Next i
    ListView1.ColumnHeaders(1).Width = ListView1.ColumnHeaders(1).Width - 1482
End Sub



Private Sub helps_Click(Index As Integer)
    Select Case Index
        Case 0
            ShowHelp
        Case 1
            AboutSystem
    End Select
End Sub

Private Sub ListView1_DblClick()
    frmEditSch.LaodSchInfo Mid(ListView1.SelectedItem.key, 2)
    frmEditSch.Show 1
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
On Error GoTo er:
    Dim rs As New Recordset
    rs.Open "select active from schmaster where schid=" & Mid(ListView1.SelectedItem.key, 2), db
    If rs!active = 1 Then
        rs.Close
        mnus(0).Caption = "  ⁄ÿÌ·"
    Else
        rs.Close
        mnus(0).Caption = " „ﬂÌ‰ "
    End If
Exit Sub
er:
    MsgBox Err.Description
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If ListView1.SelectedItem.Selected = True Then
            Me.PopupMenu mnu
        Else
            Me.PopupMenu mnu1
        End If
    End If
End Sub

Private Sub mnufiles_Click(Index As Integer)
    Select Case Index
        Case 0
            mnus1_Click 0
        Case 2
            Unload Me
    End Select
End Sub

Private Sub mnus_Click(Index As Integer)
    Dim rs As New Recordset
    Select Case Index
        Case 0
            rs.Open "select active from schmaster where schid=" & Mid(ListView1.SelectedItem.key, 2), db
            If rs!active = 1 Then
                rs.Close
                If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ  ⁄ÿÌ· «·Â„… " & ListView1.SelectedItem.SubItems(3) & " ?", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbNo Then Exit Sub
                rs.Open "update schmaster set active='" & "0" & "' where schid=" & Mid(ListView1.SelectedItem.key, 2), db
                mnus(0).Caption = " „ﬂÌ‰ "
            Else
                rs.Close
                If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ  „ﬂÌ‰ «·Â„… " & ListView1.SelectedItem.SubItems(3) & " ?", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbNo Then Exit Sub
                rs.Open "update schmaster set active='" & "1" & "' where schid=" & Mid(ListView1.SelectedItem.key, 2), db
                mnus(0).Caption = " ⁄ÿÌ· "
            End If

        Case 1
            If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ Õ–› «·Â„…  " & ListView1.SelectedItem.SubItems(3) & " ?" & vbNewLine & "      ·« Ì„ﬂ‰ «· —«Ã⁄ ⁄‰ ⁄„·Ì… «·Õ–› ", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbNo Then Exit Sub
            rs.Open "delete * from schmaster where schid=" & Mid(ListView1.SelectedItem.key, 2), db
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
            MsgBox " „ Õ–› «·„Â„… »‰Ã«Õ", vbInformation

        Case 3 ' rename jop
            Dim x As String
            rs.Open "select schname from schmaster where schid=" & Mid(ListView1.SelectedItem.key, 2), db
            x = rs!schname
            x = Trim(InputBox("«ﬂ » «·«”„ «·ÃœÌœ ··„Â„… «·Õ«·Ì… Ê„‰ À„ «‰ﬁ— ⁄·Ï „Ê«›ﬁ", "«⁄«œ…  ”„Ì… „Â„…", x))
            rs.Close
            If x = "" Then Exit Sub
            rs.Open "update schmaster set schname='" & x & "' where schid=" & Mid(ListView1.SelectedItem.key, 2), db
            ListView1.SelectedItem.SubItems(4) = x

        Case 5
            ListView1_DblClick
    End Select
End Sub

Private Sub mnus1_Click(Index As Integer)
    Select Case Index
        Case 0
            frmSchoduler.Show 1
            LoadSchdulers
        Case 2
'            Unload Me
    End Select
End Sub

Private Sub options_Click(Index As Integer)
    Select Case Index
        Case 0
            BringWindowToTop frmCall.Hwnd
            frmCall.Show
        Case 1
            frmSendOne.Show 1
        Case 3
            frmOptions.Show 1
    End Select
End Sub
