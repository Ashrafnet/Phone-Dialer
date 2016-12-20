VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ŒÌ«—«  «·‰Ÿ«„"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   3
      Left            =   240
      RightToLeft     =   -1  'True
      ScaleHeight     =   3555
      ScaleWidth      =   7155
      TabIndex        =   19
      Top             =   2280
      Width           =   7215
      Begin VB.CommandButton Command3 
         Caption         =   "ŒÌ«—«  „ ﬁœ„…"
         Height          =   495
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtMaxMessageLength 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3960
         TabIndex        =   29
         Text            =   "60"
         ToolTipText     =   " ÕœÌœ «ﬁ’Ï “„‰ · ”ÃÌ· «·„ÊÃ… «·’Ê Ì… «À‰«¡ «·„ﬂ«·„…"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtMaxAttempts 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3960
         TabIndex        =   21
         Text            =   "2"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRetryWait 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3960
         TabIndex        =   20
         Text            =   "30"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«ﬁ’Ï “„‰ · ”ÃÌ· «·„ÊÃ… «·’Ê Ì… "
         Height          =   195
         Left            =   4680
         TabIndex        =   30
         Top             =   1080
         Width           =   2355
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ „Õ«Ê·«  «·« ’«·"
         Height          =   195
         Index           =   2
         Left            =   5640
         TabIndex        =   23
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·Êﬁ  «·÷«∆⁄ »Ì‰ «·„Õ«Ê·« "
         Height          =   195
         Left            =   5160
         TabIndex        =   22
         Top             =   600
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   2
      Left            =   240
      RightToLeft     =   -1  'True
      ScaleHeight     =   3555
      ScaleWidth      =   7155
      TabIndex        =   15
      Top             =   2280
      Width           =   7215
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtModem 
         Height          =   380
         Left            =   1800
         TabIndex        =   26
         Top             =   480
         Width           =   5175
      End
      Begin VB.CheckBox chkEnableVoice 
         Alignment       =   1  'Right Justify
         Caption         =   " „ﬂÌ‰ «·’Ê "
         Height          =   495
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chkInterruptWave 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÿ⁄ «·„ÊÃ… «·’Ê Ì…"
         Height          =   400
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "«” Œœ„ Â–« «·ŒÌ«— ›Ì Õ«· «‰ﬂ ﬂ‰   —Ìœ «‰ ·« Ì „ ﬁ—«¡… „ÊÃ… ’Ê Ì… ÃœÌœ… ÿ«·„« «‰ «·„ÊÃ… «·’Ê Ì… «·Õ«·Ì…  ⁄„·"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CheckBox chkUseSoundCard 
         Alignment       =   1  'Right Justify
         Caption         =   "«” Œœ„ ﬂ—  «·’Ê  ·«Œ—«Ã «·’Ê "
         Height          =   495
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Õœœ Â–« «·ŒÌ«— «–« ﬂ‰   —Ìœ «‰  ” „⁄ «·Ï «·„·› «·’Ê Ì «À‰«¡ «·ﬁ—«¡… «Ê «· ”ÃÌ·"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÃÂ«“ «·„Êœ„"
         Height          =   195
         Left            =   6315
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   1
      Left            =   240
      RightToLeft     =   -1  'True
      ScaleHeight     =   3555
      ScaleWidth      =   7155
      TabIndex        =   8
      Top             =   2280
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   " €ÌÌ—"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "«‰ﬁ·« Â‰« · €ÌÌ— ﬂ·„… «·„—Ê—"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3960
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   3960
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " √ﬂÌœ«·ﬂ·„… "
         Height          =   195
         Index           =   1
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂ·„… «·ÃœÌœ…"
         Height          =   195
         Index           =   0
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰€ÌÌ— ﬂ·„… «·„—Ê—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5505
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   1470
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   0
      Left            =   240
      RightToLeft     =   -1  'True
      ScaleHeight     =   3555
      ScaleWidth      =   7155
      TabIndex        =   7
      Top             =   2280
      Width           =   7215
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "·«  ⁄—÷ «·‘«‘… «·—∆Ì”Ì… ··»—‰«„Ã ›Ì »œ¡  ‘€Ì· «·‰Ÿ«„"
         Height          =   255
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   600
         Value           =   -1  'True
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄„· „⁄ »œ¡  ‘€Ì· «·‰Ÿ«„"
         Height          =   375
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7223
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "⁄«„"
            Key             =   "genral"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "«·Õ„«Ì… Ê«·«„«‰"
            Key             =   "security"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "«·« ’«·"
            Key             =   "calling"
            Object.Tag             =   ""
            Object.ToolTipText     =   "]Â‰« Ì „ ⁄—÷ ŒÌ«—«  «·« ’«· «À‰«¡ «·„ﬂ«·„…"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Œ’«∆’ ÃÂ«“ «·„Êœ„"
            Key             =   "modem"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOrder 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
      Height          =   380
      Index           =   2
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   " ÿ»Ìﬁ"
      Enabled         =   0   'False
      Height          =   380
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "„Ê«›ﬁ"
      Default         =   -1  'True
      Height          =   380
      Index           =   0
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1590
      Left            =   240
      Picture         =   "frmOptions.frx":0000
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmOptions.frx":57D0
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
      Height          =   915
      Index           =   3
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   5505
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ŒÌ«—«  ÊÀÊ«»  «·‰Ÿ«„"
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
      Left            =   5730
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1875
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   8
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7935
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   8280
      X2              =   11400
      Y1              =   3720
      Y2              =   3240
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   8520
      X2              =   10080
      Y1              =   3120
      Y2              =   2760
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    cmdOrder(1).Enabled = True
End Sub

Private Sub cmdOrder_Click(Index As Integer)
    Select Case Index
        Case 0
            SaveSettings
            Unload Me
        Case 1
            SaveSettings
             
        Case 2
            Unload Me
    End Select
End Sub

Private Sub Command1_Click()
On Error GoTo er:
    Command1.Enabled = False
    If Text1(0) = Text1(1) Then
        Dim rs As New Recordset
        rs.Open "update admin set pass='" & Text1(0) & "'", db
        MsgBox " „  ⁄„·Ì…  €ÌÌ— ﬂ·„… «·„—Ê— »‰Ã«Õ", vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation
    Else
       MsgBox "ÌÃ» «‰  ﬂÊ‰ «·ﬂ·„… «·ÃœÌœ… „ÿ«»ﬁ… ·· √ﬂÌœ!", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    End If
    Exit Sub
er:
    MsgBox "Œÿ√ «À‰«¡  ⁄œÌ· ﬂ·„… «·„—Ê— , ﬁœ Ì—Ã⁄ «·”»» ·«‰ﬂ ·«  „ ·ﬂ «·’·«ÕÌ«  «·ﬂ«›Ì… ··ﬁÌ«„ »Â–Â «·„Â„…", vbCritical
End Sub

Private Sub Command2_Click()
On Error GoTo er:
    If mdiMain.Apax1.TapiSelectDevice Then
        txtModem = mdiMain.Apax1.SelectedDevice
    End If
    Exit Sub
er:
    MsgBox "Ì»œÊ «‰ ﬂ—  «·„Êœ ·« Ìœ⁄„ «·’Ê " & vbNewLine & "„⁄·Ê„«  ›‰Ì… " & vbNewLine & vbNewLine & Err.Description, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
End Sub

Private Sub Command3_Click()
    MsgBox "UnderConstruction"
End Sub

Private Sub Form_Load()
    OptemizeTitle
    LoadSettings
    cmdOrder(1).Enabled = False
    Picture1(0).ZOrder
End Sub
Sub LoadOptions()
On Error Resume Next
  txtMaxAttempts = GetSetting("VLS", "Apax", "txtMaxAttempts", txtMaxAttempts)
  txtRetryWait = GetSetting("VLS", "Apax", "txtRetryWait", txtRetryWait)
  txtMaxMessageLength = GetSetting("VLS", "Apax", "txtMaxMessageLength", txtMaxMessageLength)
  chkEnableVoice = GetSetting("VLS", "Apax", "chkEnableVoice", chkEnableVoice.Value)
  chkInterruptWave = GetSetting("VLS", "Apax", "chkInterruptWave", chkInterruptWave.Value)
  chkUseSoundCard = GetSetting("VLS", "Apax", "chkUseSoundCard", chkUseSoundCard.Value)
  txtModem = GetSetting("VLS", "Apax", "txtModem", txtModem)
  
  mdiMain.Apax1.MaxAttempts = txtMaxAttempts.Text
  mdiMain.Apax1.TapiRetryWait = txtRetryWait.Text
  mdiMain.Apax1.MaxMessageLength = txtMaxMessageLength.Text
  mdiMain.Apax1.EnableVoice = chkEnableVoice.Value
  mdiMain.Apax1.InterruptWave = chkInterruptWave.Value
  mdiMain.Apax1.UseSoundCard = chkUseSoundCard.Value
  
End Sub
Sub SetOptions()
  On Error Resume Next
  SaveSetting "VLS", "Apax", "txtMaxAttempts", txtMaxAttempts
  SaveSetting "VLS", "Apax", "txtRetryWait", txtRetryWait
  SaveSetting "VLS", "Apax", "txtMaxMessageLength", txtMaxMessageLength
  SaveSetting "VLS", "Apax", "chkEnableVoice", chkEnableVoice.Value
  SaveSetting "VLS", "Apax", "chkInterruptWave", chkInterruptWave.Value
  SaveSetting "VLS", "Apax", "chkUseSoundCard", chkUseSoundCard.Value
  SaveSetting "VLS", "Apax", "txtModem", txtModem
  
  mdiMain.Apax1.MaxAttempts = txtMaxAttempts.Text
  mdiMain.Apax1.TapiRetryWait = txtRetryWait.Text
  mdiMain.Apax1.MaxMessageLength = txtMaxMessageLength.Text
  mdiMain.Apax1.EnableVoice = chkEnableVoice.Value
  mdiMain.Apax1.InterruptWave = chkInterruptWave.Value
  mdiMain.Apax1.UseSoundCard = chkUseSoundCard.Value

End Sub

Sub OptemizeTitle()
    For i = 0 To Picture1.Count - 1
        Picture1(i).BorderStyle = 0
        Picture1(i).Left = 240
        Picture1(i).Top = 2280
    Next i
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

Sub LoadSettings()
 Const WinInfo = "Software\Microsoft\Windows\CurrentVersion"
 Const WinInfoNT = "Software\Microsoft\Windows NT\CurrentVersion"
    If GetWindows Then
        x = GetString(HKEY_LOCAL_MACHINE, WinInfo & "\Run", "VLS")
     Else
        x = GetString(HKEY_LOCAL_MACHINE, WinInfoNT & "\Run", "VLS")
    End If
    If x <> "" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    LoadOptions
End Sub
Sub SaveSettings()
    StartWithWindows Check1.Value
   SetOptions
    'SaveSetting "VLS", "Start", "Startup", Check1.Value
    
    cmdOrder(1).Enabled = False
    
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.key
        Case "genral"
            Picture1(0).ZOrder
        Case "security"
            Picture1(1).ZOrder
        Case "calling"
            Picture1(3).ZOrder
        Case "modem"
            Picture1(2).ZOrder
    End Select
End Sub

Private Sub Text1_Change(Index As Integer)
    Command1.Enabled = True
End Sub
