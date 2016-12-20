VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ”ÃÌ· «·œŒÊ·"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   360
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "sendtous"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox txtInfo 
      Height          =   330
      Index           =   0
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "adminsms"
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "œŒÊ·"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Œ—ÊÃ"
      Height          =   330
      Index           =   1
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOrder 
      Cancel          =   -1  'True
      Caption         =   "«Œ›«¡ «·»—‰«„Ã"
      Height          =   330
      Index           =   2
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "«” Œœ„ Â–« «·ŒÌ«— »œ·« „‰ «€·«ﬁ «·»—‰«„Ã ·«‰Â —»„« ÌÊÃœ „Â«„ „ÃœÊ·… ⁄·Ï «·»—‰«„Ã «·ﬁÌ«„ »Â«"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   " –ﬂ— ﬂ·„… «·„—Ê— ⁄·Ï Â–« «·ÃÂ«“"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   3615
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmLogin.frx":0000
      Left            =   360
      List            =   "frmLogin.frx":0007
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ”ÃÌ· «·œŒÊ· «·Ï ‰Ÿ«„ «·—”«∆· «·’Ê Ì…"
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
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   240
      Width           =   3765
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogin.frx":0016
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
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   4305
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   120
      Picture         =   "frmLogin.frx":00CA
      Top             =   120
      Width           =   1440
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
      Height          =   1575
      Index           =   8
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„… «·„—Ê—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   4785
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   780
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„” Œœ„"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   1110
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «· ”ÃÌ·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2880
      Width           =   885
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim canUnload As Boolean


Private Sub cmdOrder_Click(Index As Integer)
    Select Case Index
        Case 0
            IsAdmin = True
            If cmbType.ListIndex = 0 Then
                If Login(txtInfo(0), txtInfo(1)) Then
                    If Check1.Value = 1 Then
                       SaveSetting "VMS", "info", "User", txtInfo(0)
                       SaveSetting "VMS", "info", "pass", txtInfo(1)
                       SaveSetting "VMS", "info", "Type", cmbType.ListIndex
                       SaveSetting "SMSSender", "Info", "save", 1
                    Else
                        SaveSetting "VMS", "info", "pass", ""
                        SaveSetting "SMSSender", "Info", "save", 0
                        SaveSetting "VMS", "info", "Type", cmbType.ListIndex
                    End If
                    canUnload = False
                    mdiMain.StatusBar1.Panels(2).Text = " „  ”ÃÌ· «·œŒÊ·"
                    CurrentUserName = txtInfo(0)
                    
                    Unload Me
                Else
                    canUnload = 1
                    mdiMain.StatusBar1.Panels(2).Text = "Œÿ√ «À‰«¡  ”ÃÌ· «·œŒÊ·"
                    MsgBox "Œÿ√ «À‰«¡  ”ÃÌ· «·œŒÊ·" & vbNewLine & " Ì—ÃÏ „‰ ﬂ «»… «”„ «·„” Œœ„ «Ê ﬂ·„… «·„—Ê— »‘ﬂ· ’ÕÌÕ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
                End If
            Else
                ConnectByEmployee
            End If
        Case 1
            On Error Resume Next
            If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ «·Œ—ÊÃ „‰ «·»—‰«„Ã " & vbNewLine & "„⁄ «·⁄·„ «‰Â „„ﬂ‰ «‰  ﬂÊ‰ Â‰«ﬂ „Â«„ „ÃœÊ·… ⁄·Ï «·»—‰«„Ã «‰ ÌﬁÊ„ »Â« ›Ì  ÊﬁÌ Â« «·„Õœœ", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbYes Then
                Unload mdiMain.m_frmSysTray
                Set mdiMain.m_frmSysTray = Nothing
                End
            End If
        Case 2
            canUnload = False
            Unload Me
            mdiMain.Hide
    End Select
End Sub
Sub ConnectByEmployee()
    canUnload = False
    Unload Me
End Sub

Private Sub Form_Load()
    txtInfo(0) = GetSetting("VMS", "info", "user", "adminsms")
    txtInfo(1) = GetSetting("VMS", "info", "pass", "")

    cmbType.Text = cmbType.List(CInt(GetSetting("VMS", "info", "Type", 0)))
    
    Check1.Value = GetSetting("SMSSender", "Info", "save", 0)
    canUnload = True
    RemoveSysMenuX Me
    OptemizeTitle
    txtInfo(0).SelStart = 0
    txtInfo(0).SelLength = Len(txtInfo(0))
    txtInfo_Change 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = canUnload
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
    lblInfo(0).Left = Me.Width - Me.lblInfo(0).Width - 500

End Sub


Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)

End Sub

Private Sub txtInfo_Change(Index As Integer)
    If txtInfo(0) <> "" And txtInfo(1) <> "" Then
        cmdOrder(0).Enabled = 1
    Else
        cmdOrder(0).Enabled = 0
    End If

End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = Len(txtInfo(Index))

End Sub
Function Login(User As String, Pass As String) As Boolean
    On Error GoTo er:
    Dim rs As New Recordset
    rs.Open "select * from admin where username='" & User & "' and pass ='" & Pass & "'", db, adOpenStatic
    If rs.RecordCount > 0 Then Login = True
    Exit Function
er:
    Login = False
    MsgBox Err.Description
End Function

