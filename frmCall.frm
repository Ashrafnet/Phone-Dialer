VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4E0DBD51-5B94-4955-B721-E2F079EF9999}#23.0#0"; "Wave.ocx"
Begin VB.Form frmCall 
   Caption         =   "« ’«· »ÃÂ«  ⁄œÌœ…"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   14010
   WindowState     =   2  'Maximized
   Begin PhoneBazookaWaveCntl.objAudio objAudio 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1085
      _ExtentY        =   979
   End
   Begin VB.CommandButton Command3 
      Caption         =   "«»œ¡ «·« ’«·"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   15
      ToolTipText     =   "«“«·… «·„·› «·„Õœœ"
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<--"
      Height          =   330
      Index           =   4
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Õ–› Ã„Ì⁄ «·⁄‰«’—"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-->"
      Height          =   330
      Index           =   3
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "«÷«›… «·⁄‰«’— «·„Õœœ… «·Ï ﬁ«∆„… «·«—”«·"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<--"
      Height          =   330
      Index           =   5
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Õ–›  «·⁄‰«’— «·„Õœœ…"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-->>"
      Height          =   330
      Index           =   6
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   11
      ToolTipText     =   "«÷«›… ﬂ· «·ﬁ«∆„…"
      Top             =   4440
      Width           =   495
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   5655
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9975
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
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«·—ﬁ„"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«·⁄„·"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   6840
      TabIndex        =   0
      Top             =   2280
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   11245
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
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«·—ﬁ„"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "«·⁄„·"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Õœœ «·„·› «·’Ê Ì"
      Filter          =   "*.wav"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "«—›«ﬁ „·› ’Ê Ì"
      Top             =   1680
      Width           =   375
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Õœœ „”«— «·„·› «·’Ê Ì"
      Top             =   1680
      Width           =   5175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6840
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   6255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   1275
      ScaleWidth      =   6075
      TabIndex        =   16
      Top             =   7920
      Visible         =   0   'False
      Width           =   6135
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ì „ Õ«·Ì« «·« ’«· »‹‹"
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
         Height          =   195
         Index           =   1
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   1470
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ì „ Õ«·Ì« «·« ’«· »‹‹"
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
         Height          =   195
         Index           =   0
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1470
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   2160
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   240
      Picture         =   "frmCall.frx":0000
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCall.frx":44C5
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
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   6225
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "—”«·… ’Ê Ì… «·Ï „Ã„Ê⁄… «‘Œ«’"
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
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   3210
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   10080
      X2              =   11640
      Y1              =   2760
      Y2              =   2400
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   9840
      X2              =   12960
      Y1              =   3360
      Y2              =   2880
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Index           =   8
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   14055
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCall.frx":456C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCall.frx":4ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCall.frx":5010
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCall.frx":5362
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCall.frx":56B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCall.frx":5A06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÃÂ«  «· Ì ” „ «·« ’«· »Â«"
      Height          =   195
      Index           =   2
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÃÂ«  «·„ Ê›—…"
      Height          =   195
      Index           =   1
      Left            =   12000
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ — „Ã„Ê⁄…"
      Height          =   195
      Index           =   0
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.Menu file 
      Caption         =   "„·›"
      Begin VB.Menu files 
         Caption         =   "≈€·«ﬁ"
         Index           =   0
      End
   End
   Begin VB.Menu option 
      Caption         =   "«œÊ« "
      Begin VB.Menu options 
         Caption         =   "„ﬂ«·„… ›—œÌ…..."
         Index           =   0
      End
      Begin VB.Menu options 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu options 
         Caption         =   "ŒÌ«—« ..."
         Index           =   2
      End
   End
   Begin VB.Menu help 
      Caption         =   " ⁄·Ì„« "
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
Attribute VB_Name = "frmCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
    LoadContacts Combo1.ItemData(Combo1.ListIndex)
End Sub


Private Sub Command1_Click(Index As Integer)
    Dim j As Integer
    Select Case Index
        Case 0
            frmAddUserToGroup.Show 1
            Combo1_Click
        Case 1
            
        Case 3
            If ListView1.ListItems.Count = 0 Then MsgBox "·« ÌÊÃœ ⁄‰«’— ›Ì «·ﬁ«∆„…", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            If ListView1.SelectedItem.Text = "" Then MsgBox "«Œ — ⁄‰’— ⁄·Ï «·«ﬁ· „‰ «·ﬁ«∆„…  «Ê·«", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            'If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ «÷«›… Ã„Ì⁄ «·⁄‰«’— «·„Õœœ… " & "" & "" & "" & vbNewLine & " ", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbYes Then
                For j = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(j).Selected = True Then
                        If Not IsIN(j) Then
                            'Dim itmX As ListItem
                            Set itmX = ListView2.ListItems. _
                            Add(, ListView1.ListItems(j).key, CStr(ListView1.ListItems(j).Text), 6, 6)
                            itmX.SubItems(1) = CStr(ListView1.ListItems(j).SubItems(1))
                            itmX.SubItems(2) = CStr(ListView1.ListItems(j).SubItems(2))
                            itmX.SubItems(3) = CStr(ListView1.ListItems(j).SubItems(3))
                            
                        End If
                    End If
                Next j
            'End If

            
        Case 4
            If ListView2.ListItems.Count = 0 Then MsgBox "·« ÌÊÃœ ⁄‰«’— ›Ì «·ﬁ«∆„…", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            If ListView2.SelectedItem.Text = "" Then MsgBox "«Œ — ⁄‰’— ⁄·Ï «·«ﬁ· „‰ «·ﬁ«∆„…  «Ê·«", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ «“«·… Ã„Ì⁄ «·⁄‰«’—  " & "" & vbNewLine & "", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbYes Then
                ListView2.ListItems.Clear
            End If
            
        Case 5
            If ListView2.ListItems.Count = 0 Then MsgBox "·« ÌÊÃœ ⁄‰«’— ›Ì «·ﬁ«∆„…", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            If ListView2.SelectedItem.Text = "" Then MsgBox "«Œ — ⁄‰’— ⁄·Ï «·«ﬁ· „‰ «·ﬁ«∆„…  «Ê·«", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            'If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ «“«·… Ã„Ì⁄ «·⁄‰«’— «·„Õœœ… " & vbNewLine & "", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbYes Then
                For j = 1 To ListView2.ListItems.Count
                    If ListView2.ListItems(j - i).Selected = True Then
                        ListView2.ListItems.Remove j - i
                        i = i + 1
                    End If
                Next j
            'End If
        Case 6
            If ListView1.ListItems.Count = 0 Then MsgBox "·« ÌÊÃœ ⁄‰«’— ›Ì «·ﬁ«∆„…", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            If ListView1.SelectedItem.Text = "" Then MsgBox "«Œ — ⁄‰’— ⁄·Ï «·«ﬁ· „‰ «·ﬁ«∆„…  «Ê·«", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ «÷«›… Ã„Ì⁄ «·⁄‰«’—  " & "" & " «·Ï ﬁ«∆„… «·ÃÂ«  «· Ì ”Ì „ «·«—”«· ·Â« " & "" & vbNewLine & " „⁄ «·⁄·„ «‰Â ·« Ì„ﬂ‰ «· —«Ã⁄ ⁄‰ Â–« «·«Ã—«¡", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbYes Then
                For j = 1 To ListView1.ListItems.Count
                    If Not IsIN(j) Then
                        Set itmX = ListView2.ListItems. _
                        Add(, ListView1.ListItems(j).key, CStr(ListView1.ListItems(j).Text), 6, 6)
                        itmX.SubItems(1) = CStr(ListView1.ListItems(j).SubItems(1))
                        itmX.SubItems(2) = CStr(ListView1.ListItems(j).SubItems(2))
                            itmX.SubItems(3) = CStr(ListView1.ListItems(j).SubItems(3))
                    End If
                Next j
            End If


    End Select
    'lblInfo(5) = "ﬁ«∆„… «·ÃÂ«  «· Ì ”Ì „ «·«—”«· ·Â« = " & ListView2.ListItems.Count & " ÃÂ… "

End Sub
Function IsIN(Index As Integer) As Boolean
On Error Resume Next
    For i = 1 To ListView2.ListItems.Count
        'If Index = i Then IsIN = 1: Exit Function
        If ListView2.ListItems(i).SubItems(1) = ListView1.ListItems(Index).SubItems(1) And ListView2.ListItems(i).SubItems(2) = ListView1.ListItems(Index).SubItems(2) Then
            IsIN = True
            Exit Function
        Else
            IsIN = False
        End If
    Next
End Function


Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            CommonDialog1.Filter = "Wave Files (*.wav)|*.wav"
            CommonDialog1.ShowOpen
            
            If CommonDialog1.Filename <> "" Then
                On Error Resume Next
                If FileLen(CommonDialog1.Filename) < 1 Then Exit Sub
                If IsINList(CommonDialog1.Filename) Then Exit Sub
                If Not IsPhoneWave(CommonDialog1.Filename) Then Exit Sub
                Combo2.AddItem CommonDialog1.Filename
                Combo2.Text = CommonDialog1.Filename
            End If

        Case 1
            On Error Resume Next
            If Combo2.ListCount < 1 Then Exit Sub
            Combo2.RemoveItem Combo2.ListIndex
            Combo2.Text = Combo2.List(0)
    End Select
End Sub
Function IsPhoneWave(WavePath As String, Optional NoMSG As Boolean = False) As Boolean
    Dim xx As String
    If objAudio.OpenWaveFile(WavePath) Then
        If (objAudio.GetWaveFormat.wFormatTag = 1) And (objAudio.GetWaveFormat.wBitsPerSample = 16) And (objAudio.GetWaveFormat.nSamplesPerSec = 8000) Then
            
            objAudio.SetWaveFormat PCM_8kHz16bit_Voice_Modems
            IsPhoneWave = True
        Else
            
        xx = " ‰”Ìﬁ «·„·› «·Õ«·Ì :" & vbNewLine & "           "
'        xx = xx & "" & objAudio.GetWaveFormat.wFormatTag & vbNewLine & "           "
        xx = xx & "Bits per sample = " & objAudio.GetWaveFormat.wBitsPerSample & vbNewLine & "           "
        xx = xx & "Samples per second = " & objAudio.GetWaveFormat.nSamplesPerSec
 
            
            IsPhoneWave = False
            If NoMSG Then Exit Function
            MsgBox "Œÿ√: ÌÃ» «‰  Œ «— „·› ’Ê Ì „ Ê«›ﬁ „⁄ «·„·›«  «· Ì Ì„ﬂ‰ ·Œÿ «·Â« › «‰ ÌÕ„·Â« ÊÂÌ „·›«  «·’Ê  «· Ì  ﬂÊ‰ »«·’Ì€… «· «·Ì… :  " & vbCr & vbLf & vbLf & "PCM 16-bit 8Khz" & vbNewLine & vbNewLine & " " & xx, _
                vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
        End If
    Else
        MsgBox "Error: Couldn't find wave file." & vbCr & vbLf & vbLf, _
            vbOKOnly + vbCritical
    End If

End Function

Function IsINList(str As String) As Boolean
    For i = 0 To Combo2.ListCount - 1
        If Combo2.List(i) = str Then IsINList = True: Exit Function
    Next i
End Function

Private Sub Command3_Click()
    If ListView2.ListItems.Count < 1 Then MsgBox "·« ÌÊÃœ ⁄‰«’— ›Ì «·ﬁ«∆„…", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
    If Combo2.ListCount < 1 Then MsgBox "ÌÃ» «‰  Õœœ „·› ’Ê Ì Ê«Õœ ⁄·Ï «·«ﬁ· ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
    If MsgBox("Â· «‰  „ «ﬂœ „‰ «‰ﬂ  —Ìœ «‰   ’· »ﬂ· «·ÃÂ«  «·„ÊÃÊœ… ›Ì «·ﬁ«∆„… «⁄·«Â ø" & vbNewLine & vbNewLine & "        Ì»·€ ⁄œœ «·ÃÂ«  «· Ì ”Ì „ «·« ’«· »Â« " & ListView1.ListItems.Count & " ÃÂ… " & vbNewLine & "        Ì»·€ ⁄œœ «·„·›«  «·’Ê Ì… «· Ì ”Ì „ «—”«·Â« ⁄»— «·„Êœ„  " & Combo2.ListCount & " ", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbNo Then Exit Sub
    Dim wavfiles As String
    For i = 0 To Combo2.ListCount - 1
        wavfiles = wavfiles & Combo2.List(i) & ","
    Next i
    Picture1.Visible = True
    DoEvents
    ProgressBar1.Visible = True
    ProgressBar1.Value = 0
    ProgressBar1.Max = ListView2.ListItems.Count
    For i = 1 To ListView2.ListItems.Count
        ProgressBar1.Value = ProgressBar1.Value + 1
        If i = 2 Then
            LogCall ListView2.ListItems(i).SubItems(2), ListView2.ListItems(i).SubItems(1), 1, wavfiles, Mid(ListView2.ListItems(i).key, 2)
        Else
            LogCall ListView2.ListItems(i).SubItems(2), ListView2.ListItems(i).SubItems(1), 0, wavfiles, Mid(ListView2.ListItems(i).key, 2)
        End If
    Next i
    ProgressBar1.Value = ProgressBar1.Max
    'ProgressBar1.Visible = False
End Sub

Private Sub files_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            
        Case 2
            
    End Select
End Sub

Private Sub Form_Load()
    LoadGroups
    LVFullRowSelect Me.ListView1
    LVFullRowSelect Me.ListView2
    OptemizeTitle
End Sub
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
    Picture1.BorderStyle = 0
End Sub

Sub LoadGroups()
    On Error GoTo er:
    Dim rs As New Recordset
    Dim i As Integer
    rs.Open "select * from groups ", db
    Combo1.Clear
    i = 0
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
    Dim i As Integer
    rs.Open "select * from contact where gid=" & Gid, db
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

Private Sub Form_Resize()
On Error Resume Next
    OptemizeTitle
    ListView2.Width = (Me.Width \ 2) - 600
    ListView1.Width = ListView2.Width
    ListView1.Height = Me.Height - 2900
    ListView2.Height = Me.Height - 4300
    Command3.Top = ListView2.Top + ListView2.Height + 120
    Picture1.Top = ListView2.Top + ListView2.Height + 10
    ListView1.Left = ListView2.Width + 950
    'ProgressBar1.Top = Command3.Top
'    ProgressBar1.Width = ListView2.Width - Command3.Width - 240
    For i = 3 To 6
        Command1(i).Left = ListView1.Width + 300
    Next i
    Combo1.Width = ListView1.Width
    Combo1.Left = ListView1.Left
    Combo2.Width = ListView2.Width - (Command2(1).Left + Command2(1).Width + 0)
    Combo2.Left = ListView2.Left + (Command2(1).Left + Command2(1).Width + 0)

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
    Command1_Click 3
End Sub

Private Sub ListView2_DblClick()
    Command1_Click 5
End Sub

Private Sub options_Click(Index As Integer)
    Select Case Index
        Case 0
            frmSendOne.Show 1
        Case 2
            frmOptions.Show 1
    End Select
End Sub
