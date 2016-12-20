VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4E0DBD51-5B94-4955-B721-E2F079EF9999}#23.0#0"; "Wave.ocx"
Begin VB.Form frmRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ”ÃÌ· ’Ê "
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·ﬂ «»… ›Êﬁ «·„·› «–« ﬂ«‰ „ÊÃÊœ"
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   5400
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "‰”Œ «·„”«—"
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
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "«‰”Œ „”«— «·„› «·–Ì  „  ”ÃÌ·Â"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   3720
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ŒÌ«—« "
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
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   11
      ToolTipText     =   "ŒÌ«—«  ÃÂ«“ «·’Ê "
      Top             =   4440
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Õœœ „”«— «·„·› «·’Ê Ì"
      Top             =   2280
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "«—›«ﬁ „·› ’Ê Ì"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4440
      Width           =   855
   End
   Begin VB.ComboBox cmbWaveDevice 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5040
      Width           =   4575
   End
   Begin PhoneBazookaWaveCntl.objAudio objAudio 
      Left            =   0
      Top             =   1560
      _ExtentX        =   1085
      _ExtentY        =   979
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "⁄—÷ «· ”ÃÌ·"
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
      Index           =   1
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "«‰ﬁ— Â‰« ·⁄—÷ „«  „  ”ÃÌ·Â"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "»œ¡ «· ”ÃÌ·"
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
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "«‰ﬁ— Â‰« ·»œ¡  ”ÃÌ· „·› ’Ê Ì ÃœÌœ"
      Top             =   2640
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cdWaveFile 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Õœœ «·„·› «·’Ê Ì"
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«” Œœ„ Â–Â «·«œ«… · ”ÃÌ· «·„·›«  «·’Ê Ì… «·„ Ê«›ﬁ… „⁄ «·„Êœ„«  , ÕÌÀ ·« Ì„ﬂ‰ «” Œœ«„ «·« ’Ì€… „⁄Ì‰… „ Ê«›ﬁ… „⁄ «ÃÂ“… «·„Êœ„"
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
      Height          =   795
      Index           =   0
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   4305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÂ«“ «·’Ê "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   7
      Top             =   5160
      Width           =   915
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
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„”Ã· «·’Ê "
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
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬁ„ » ÕœÌœ „”«— «·„·› «·–Ì ”Ì „ Õ›Ÿ «·„·› ⁄·ÌÂ Ê„‰ À„ «‰ﬁ— ⁄·Ï “— »œ¡ «· ”ÃÌ·"
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
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   0
      Left            =   4680
      Picture         =   "frmRecord.frx":0000
      Top             =   240
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   660
      Index           =   1
      Left            =   360
      Picture         =   "frmRecord.frx":18CE
      Top             =   720
      Width           =   660
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   8
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Recording As Boolean
Dim Playing As Boolean
Private Sub cmdOrder_Click(Index As Integer)
'Label1.Visible = True
    Select Case Index
        Case 0
            
            StartRecord
        Case 1
            PlayBackWave
        Case 2
            Clipboard.SetText Combo2.Text
    End Select
    
End Sub

Private Sub Combo2_Click()
    If Trim(Combo2.Text) <> "" Then
        cmdOrder(0).Enabled = True
    Else
        cmdOrder(0).Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
On Error Resume Next
        
    ' tell the common dialog that we want to know if the user presses cancel
    cdWaveFile.CancelError = True
        
    ' we want the common dialog to filter in only Wave files
    ' filters are very useful when using the common dialog control
    cdWaveFile.Filter = "Wave Files (*.wav)|*.wav"
    
    cdWaveFile.InitDir = GetPathString
    
    ' tell the common dialog to show itself in the FileOpen mode
    cdWaveFile.ShowSave
        
    ' code stops executing and waits for a result from the common dialog box
        
    ' did you press cancel?
    If Err = cdlCancel Then
        'MsgBox "No file selected", vbOKOnly
        Exit Sub
    End If
        
    ' did the user choose a respectable file name?
    If (cdWaveFile.Filename = vbNullString) Then
        'MsgBox "No file selected", vbOKOnly
        Exit Sub
    End If
        Combo2.AddItem cdWaveFile.Filename
        Combo2.Text = cdWaveFile.Filename
'    If objAudio.OpenWaveFile(cdWaveFile.Filename) Then
'
'    Else
'        MsgBox "Error : Couldn't find wave file", vbOKOnly
'    End If

End Sub

Private Sub Command3_Click()
    If Me.Height = 5415 Then
        Me.Height = 5415 + 1222
    Else
        Me.Height = 5415
    End If
End Sub

Private Sub Form_Load()
OptemizeTitle
    Recording = False
    ' string to hold our wave devices
    Dim strWaveDevList()        As String
    
    ' load the available wave devices in a combo box
    objAudio.LoadWaveOutDevices strWaveDevList
    For i = 0 To UBound(strWaveDevList)
        cmbWaveDevice.AddItem strWaveDevList(i)
    Next i
    cmbWaveDevice.ListIndex = 0

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
'    lblInfo(2).Left = Me.Width - lblInfo(2).Width - 300
'    lblInfo(0).Left = Me.Width - Me.lblInfo(0).Width - 500

End Sub

Private Sub Form_Resize()
    Command1.Top = Me.Height - 620 - Command3.Height
End Sub

Private Sub objAudio1_FinishedPlaying()
    cmdOrder(1).Caption = "⁄—÷ «· ”ÃÌ·"
End Sub
Sub StartRecord()
   On Error GoTo er:
   cmdOrder(0).Caption = "«Ìﬁ«› «· ”ÃÌ·"
   If Recording = False Then
        
        If cmbWaveDevice.ListIndex <> -1 Then
            cmdOrder(0).Caption = "«Ìﬁ«› «· ”ÃÌ·"
            Label1 = 0: Label1.Visible = 1: Timer1.Enabled = 1
            Recording = True
            ' let set the audio format
            objAudio.SetWaveFormat PCM_8kHz16bit_Voice_Modems
            
            ' start the Mic Input
            objAudio.StartMicInput cmbWaveDevice.ListIndex
            
            ' record audio to file
            objAudio.OpenOutputFileToRecordTo Combo2.Text
        Else
            Recording = False
            MsgBox "·„  ﬁ„ »«Œ Ì«— ÃÂ«“ «·’Ê " & vbNewLine & vbNewLine & "    ·«Œ Ì«— ÃÂ«“ ’Ê Ì «‰ﬁ— ⁄·Ï “— ŒÌ«—«  Ê„‰ À„ Õœœ ÃÂ«“ «·’Ê  ", vbOKOnly
        End If
    Else
        cmdOrder(1).Enabled = True
        cmdOrder(2).Enabled = True
        Recording = False
        ' need to stop recording
        cmdOrder(0).Caption = "»œ¡ «· ”ÃÌ·"
        Label1 = 0: Label1.Visible = 0: Timer1.Enabled = 0
        objAudio.CloseOutputFile
        objAudio.StopMicInput
    End If

   
   Exit Sub
er:
    MsgBox Err.Description, vbCritical
End Sub

Sub PlayBackWave()
On Error GoTo er:
    ' lets see if we have valid wave devices to play to
    If cmbWaveDevice.ListIndex <> -1 Then
        If Not Playing Then
            
            ' lets open the wave file, this also sets the audio format for us
            If objAudio.OpenWaveFile(Combo2.Text) Then
                Playing = True
                Label1 = 0: Label1.Visible = 1: Timer1.Enabled = 1
                cmdOrder(1).Caption = "«Ìﬁ«› «·⁄—÷"
                cmdOrder(1).ToolTipText = "«‰ﬁ— Â‰« ·«Ìﬁ«› «·⁄—÷"

                ' start the speaker output
                objAudio.StartSpeakerOutput cmbWaveDevice.ListIndex
                
                ' play the wave file to the speakers
                objAudio.PlayFile Combo2.Text, objAudio.hSpeakerWaveOut
            End If
        Else
            objAudio.StopSpeakerOutput
            Playing = False
            Label1 = 0: Label1.Visible = 0: Timer1.Enabled = 0
            cmdOrder(1).Caption = "⁄—÷ «· ”ÃÌ·"
            cmdOrder(1).ToolTipText = "«‰ﬁ— Â‰« ·⁄—÷ „«  „  ”ÃÌ·Â"
        End If
    End If
    Exit Sub
er:
    Playing = False
    MsgBox Err.Description, vbCritical
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    objAudio.StopWaveProcessing
End Sub

Private Sub objAudio_FinishedPlaying()
    Playing = False
    objAudio.StopSpeakerOutput
    cmdOrder(1).Caption = "⁄—÷ «· ”ÃÌ·"
    cmdOrder(1).ToolTipText = "«‰ﬁ— Â‰« ·⁄—÷ „«  „  ”ÃÌ·Â"
    Label1 = 0: Label1.Visible = 0: Timer1.Enabled = 0
End Sub

Private Sub Timer1_Timer()
    Label1 = Val(Label1) + 1
End Sub
