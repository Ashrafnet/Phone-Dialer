VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4E0DBD51-5B94-4955-B721-E2F079EF9999}#23.0#0"; "Wave.ocx"
Object = "{797E7185-0DB7-4E3A-939B-234871F7FAC9}#1.11#0"; "Apax1.ocx"
Begin VB.Form frmSendOne 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«—”«· —”«·… ’Ê Ì… «·Ï ‘Œ’"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin APAX1.Apax Apax1 
      Height          =   2775
      Left            =   6360
      TabIndex        =   28
      Top             =   1440
      Width           =   3855
      Baud            =   19200
      ComNumber       =   0
      DeviceType      =   1
      DataBits        =   8
      DTR             =   -1  'True
      HWFlowUseDTR    =   0   'False
      HWFlowUseRTS    =   0   'False
      HWFlowRequireDSR=   0   'False
      HWFlowRequireCTS=   0   'False
      LogAllHex       =   0   'False
      Logging         =   0
      LogHex          =   -1  'True
      LogName         =   "APRO.LOG"
      LogSize         =   10000
      Parity          =   0
      PromptForPort   =   -1  'True
      RS485Mode       =   0   'False
      RTS             =   -1  'True
      StopBits        =   1
      SWFlowOptions   =   0
      XOffChar        =   19
      XOnChar         =   17
      WinsockMode     =   0
      WinsockAddress  =   ""
      WinsockPort     =   "telnet"
      WsTelnet        =   -1  'True
      AnswerOnRing    =   2
      EnableVoice     =   0   'False
      MaxAttempts     =   3
      InterruptWave   =   -1  'True
      MaxMessageLength=   60
      SelectedDevice  =   "Agere Systems AC'97 Modem"
      SilenceThreshold=   50
      TapiNumber      =   ""
      TapiRetryWait   =   30
      TrimSeconds     =   2
      UseSoundCard    =   -1  'True
      CaptureFile     =   "APROTERM.CAP"
      CaptureMode     =   0
      Color           =   8388608
      Columns         =   80
      Emulation       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   24
      ScrollbackEnabled=   0   'False
      ScrollbackRows  =   200
      TerminalActive  =   -1  'True
      TerminalBlinkTime=   500
      TerminalHalfDuplex=   0   'False
      TerminalLazyByteDelay=   200
      TerminalLazyTimeDelay=   100
      TerminalUseLazyDisplay=   -1  'True
      TerminalWantAllKeys=   -1  'True
      Version         =   "1.12"
      Object.Visible         =   -1  'True
      DataTriggerString=   ""
      ProtocolStatusDisplay=   -1  'True
      Protocol        =   7
      AbortNoCarrier  =   0   'False
      AsciiCharDelay  =   0
      AsciiCRTranslation=   0
      AsciiEOFTimeout =   364
      AsciiEOLChar    =   13
      AsciiLFTranslation=   0
      AsciiLineDelay  =   0
      AsciiSuppressCtrlZ=   0   'False
      BlockCheckMethod=   4
      FinishWait      =   364
      HandshakeRetry  =   10
      HandshakeWait   =   1092
      HonorDirectory  =   0   'False
      IncludeDirectory=   0   'False
      KermitCtlPrefix =   35
      KermitHighbitPrefix=   89
      KermitMaxLen    =   80
      KermitMaxWindows=   0
      KermitPadCharacter=   0
      KermitPadCount  =   0
      KermitRepeatPrefix=   126
      KermitSWCTurnDelay=   0
      KermitTerminator=   13
      KermitTimeoutSecs=   5
      ReceiveDirectory=   ""
      ReceiveFileName =   ""
      RTSLowForWrite  =   0   'False
      SendFileName    =   "*.*"
      StatusInterval  =   10
      TransmitTimeout =   1092
      UpcaseFileNames =   -1  'True
      WriteFailAction =   2
      XYmodemBlockWait=   91
      Zmodem8K        =   0   'False
      ZmodemFileOptions=   5
      ZmodemFinishRetry=   0
      ZmodemOptionOverride=   0   'False
      ZmodemRecover   =   0   'False
      ZmodemSkipNoFile=   0   'False
      Caption         =   "APAX v1.14"
      CaptionAlignment=   2
      CaptionWidth    =   100
      LightWidth      =   40
      LightsLitColor  =   255
      LightsNotLitColor=   8421376
      ShowLightCaptions=   -1  'True
      ShowLights      =   -1  'True
      ShowStatusBar   =   -1  'True
      ShowToolBar     =   -1  'True
      ShowDeviceSelButton=   -1  'True
      ShowConnectButtons=   -1  'True
      ShowProtocolButtons=   -1  'True
      ShowTerminalButtons=   -1  'True
      DoubleBuffered  =   0   'False
      Enabled         =   -1  'True
      Cursor          =   0
      TapiStatusDisplay=   0   'False
      CommPort        =   0
      DTREnable       =   -1  'True
      Handshaking     =   0
      InBufferSize    =   1024
      OutBufferSize   =   512
      RTSEnable       =   -1  'True
      Settings        =   "19200,N,8,1"
      InputMode       =   0
      InputLen        =   0
      MSCommCompatible=   0   'False
      RTThreshold     =   0
      SThreshold      =   0
      FTPAccount      =   ""
      FTPConnectTimeout=   0
      FTPFileType     =   1
      FTPPassword     =   ""
      FTPRestartAt    =   0
      FTPServerAddress=   ""
      FTPTransferTimeout=   1092
      FTPUserName     =   ""
      FilterTapiDevices=   -1  'True
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "ﬁÿ⁄ «·« ’«·"
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
      Index           =   2
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   5235
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7435
            Text            =   "Ã«Â“"
            TextSave        =   "Ã«Â“"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin PhoneBazookaWaveCntl.objAudio objAudio 
      Left            =   5640
      Top             =   1560
      _ExtentX        =   1085
      _ExtentY        =   979
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   26
      ToolTipText     =   "«“«·… «·„·› «·„Õœœ"
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   27
      ToolTipText     =   "«—›«ﬁ „·› ’Ê Ì"
      Top             =   1680
      Width           =   375
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   24
      ToolTipText     =   "„”Õ «·—ﬁ„"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "#"
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
      Index           =   11
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "0"
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
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "9"
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
      Index           =   9
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "8"
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
      Index           =   8
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "7"
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
      Index           =   7
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "6"
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
      Index           =   6
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "5"
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
      Index           =   5
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "4"
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
      Index           =   4
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "3"
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
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "2"
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
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "1"
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
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cdWaveFile 
      Left            =   5640
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Wave Files (*.wav)|*.wav"
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
      Height          =   380
      Index           =   1
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "«»œ√ «·« ’«·"
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
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "Õœœ „”«— «·„·› «·’Ê Ì"
      Top             =   1680
      Width           =   4455
   End
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
      Height          =   380
      Left            =   600
      MaxLength       =   12
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3- «‰ﬁ— ⁄·Ï “— «»œ√ «·« ’«· ·Ì „ «·« ’«· »’«Õ» «·—ﬁ„ «·„œŒ·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   5
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3960
      Width           =   2385
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ŒÿÊ«  «·« ’«·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Index           =   4
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2280
      Width           =   1290
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2- «‰ﬁ— ⁄·Ï «·«—ﬁ«„ «·„Ã«Ê—… · ÕœÌœ —ﬁ„ «·Â« › «·–Ì  —Ìœ «·« ’«· »Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3240
      Width           =   2385
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1- Õœœ «·„·›«  «·’Ê Ì… «· Ì  —Ìœ «‰ Ì „ ‰ﬁ·Â« «·Ï «·„—”· «·ÌÂ „‰ Œ·«· «·„Êœ„"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2520
      Width           =   2385
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„·› «·’Ê Ì"
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
      Index           =   1
      Left            =   4425
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "—”«·… ’Ê Ì… ·‘Œ’"
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
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1950
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«ﬂ » —ﬁ„  «·‘Œ’ «·–Ì  —Ìœ «‰  —”· ·Â —”«·… ’Ê Ì… ⁄«Ã·… Ê„‰ À„ «Œ — «·„·› «·’Ê Ì À„ «‰ﬁ— ⁄·Ï “— «—”«·"
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
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3825
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   120
      Picture         =   "frmSendOne.frx":0000
      Top             =   120
      Width           =   1155
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   7200
      X2              =   8760
      Y1              =   480
      Y2              =   120
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   6960
      X2              =   10080
      Y1              =   1080
      Y2              =   600
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   2760
      Top             =   2160
      Width           =   2655
   End
End
Attribute VB_Name = "frmSendOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Apax1_OnTapiConnect()
    If Apax1.EnableVoice Then
        StatusBar1.Panels(2).Text = " „ «‰‘«¡ «·« ’«·"
        Apax1.TapiPlayWaveFile Combo2.Text
    Else
        StatusBar1.Panels(2).Text = " „ «‰‘«¡ «·« ’«· ·ﬂ‰ «·„Êœ„ ·« Ìœ⁄„ «·’Ê "
    End If
    
End Sub

Private Sub Apax1_OnTapiDTMF(ByVal Digit As Byte, ByVal ErrorCode As Long)
    StatusBar1.Panels(2).Text = "·ﬁœ ‰ﬁ— «·„ ’· «·»⁄Ìœ ⁄·Ï “— " & CStr(Digit - 48)
End Sub

Private Sub Apax1_OnTapiFail()
    If Apax1.TapiCancelled Then
      StatusBar1.Panels(2).Text = " „ «·€«¡ «·« ’«·"
    Else
      StatusBar1.Panels(2).Text = ("›‘· «·« ’«·")
    End If
    cmdOrder(0).Visible = 1
    cmdOrder(2).Visible = 0
End Sub

Private Sub Apax1_OnTapiStatus(ByVal First As Boolean, ByVal Last As Boolean, ByVal Device As Long, ByVal Message As Long, ByVal Param1 As Long, ByVal Param2 As Long, ByVal Param3 As Long)
    StatusBar1.Panels(2).Text = TapiStateStr(Apax1.TapiState)
    StatusBar1.Panels(1).Text = Apax1.TapiStatusMsg(Message, Param1, Param2)
    If InStr(1, StatusBar1.Panels(1).Text, "retry") Then
        StatusBar1.Panels(2).Text = "«‰ Ÿ«— «⁄«œ… „Õ«Ê·… «·« ’«·"
    End If
End Sub

Private Sub Apax1_OnTapiWaveNotify(ByVal MSG As Apax1.TxWaveMessage)
  Dim S As String
  Select Case MSG
    Case waPlayOpen: S = "Ì „ «·«‰  ‘€Ì· «·„·› «·’Ê Ì"
    Case waPlayDone: S = "«‰ ÂÏ «·„·› «·’Ê Ì „‰ «·ﬁ—«¡…"
    Case waPlayClose: S = " „ «€·«ﬁ «·„·› «·’Ê Ì"
    Case waRecordOpen: S = "»œ¡ «· ”ÃÌ·"
    Case waDataReady: S = "waDataReady"
    Case waRecordClose: S = "«‰ Â  ⁄„·Ì… «· ”ÃÌ·"
  End Select
  StatusBar1.Panels(2).Text = S
End Sub

Private Sub Apax1_OnTapiWaveSilence(StopRecording As Boolean, Hangup As Boolean)
'Â–« «·ÕœÀ Ì‰‘«¡ «–« ﬂ«‰ «·„ ’· «·»⁄Ìœ ﬁœ  Êﬁ› ⁄‰ «·ﬂ·«„ , ›Ì„ﬂ‰ﬂ «‰  ﬁÊ„ »⁄„·  Êﬁ› ·⁄„·Ì…  ”ÃÌ· «·’Ê  , «Ê ﬁÿ⁄ «·„ﬂ«·„…
    MsgBox "ÕœÀ Wave Silence" & "       StopRecording Value : " & StopRecording & "       Hangup Value : " & Hangup
End Sub

Private Sub cmdNo_Click(Index As Integer)
    Select Case Index
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
            Text1 = Text1 & Index
        Case 10
            Text1 = Text1 & "*"
        Case 11
            Text1 = Text1 & "#"
    End Select
End Sub

Private Sub cmdOrder_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim wavfiles As String
            For i = 0 To Combo2.ListCount
                wavfiles = wavfiles & Combo2.List(i) & ","
            Next i
            wavfiles = Combo2.Text
            If Not IsPhoneWave(Combo2.Text) Then Exit Sub
            If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ «·« ’«· »Â–Â «·ÃÂ…ø", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbNo Then Exit Sub
            Apax1.TapiNumber = Text1
            StatusBar1.Panels(2).Text = "Ì „ «·«‰ «·« ’«·"
            MsgBox Apax1.TapiDial
            cmdOrder(0).Visible = 0
            cmdOrder(2).Visible = 1
'            Apax1.TapiPlayWaveFile "c:\x.wav"
            LogCall Text1, "„ÃÂÊ·", 1, wavfiles, 0
        Case 1
            Unload Me
        Case 2 ' disconnect
            Apax1.TapiCancelCall
    End Select
End Sub

Private Sub Combo2_Change()
    Text1_Change
End Sub

Private Sub Combo2_Click()
    Text1_Change
End Sub

Private Sub Command1_Click()
    Text1 = ""
End Sub

'''Private Sub Command2_Click()
'''On Error GoTo er:
'''    CommonDialog1.ShowOpen
'''
'''    If CommonDialog1.Filename <> "" Then
'''        Combo2.AddItem CommonDialog1.Filename
'''        Combo2.Text = CommonDialog1.Filename
'''    End If
'''Exit Sub
'''er:
''''    MsgBox "Cancel", vbCritical
'''End Sub
Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
        On Error Resume Next
            cdWaveFile.Filter = "Wave Files (*.wav)|*.wav"
            cdWaveFile.ShowOpen
            
            If cdWaveFile.Filename <> "" Then
                
                If FileLen(cdWaveFile.Filename) < 1 Then Exit Sub
                If IsINList(cdWaveFile.Filename) Then Exit Sub
                If Not IsPhoneWave(cdWaveFile.Filename) Then Exit Sub
                Combo2.AddItem cdWaveFile.Filename
                Combo2.Text = cdWaveFile.Filename
            End If

        Case 1
            On Error Resume Next
            If Combo2.ListCount < 1 Then Exit Sub
            Combo2.RemoveItem Combo2.ListIndex
            Combo2.Text = Combo2.List(0)
    End Select
End Sub
Function IsINList(str As String) As Boolean
    For i = 0 To Combo2.ListCount - 1
        If Combo2.List(i) = str Then IsINList = True: Exit Function
    Next i
End Function
Private Sub Form_Load()
    OptemizeTitle
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
    'lblInfo(0).Left = Me.Width - Me.lblInfo(0).Width - 500

End Sub

Private Sub List1_Click()
    MsgBox List1.Text
End Sub

Private Sub Text1_Change()
    If Text1 <> "" And Combo2.Text <> "" Then
        cmdOrder(0).Enabled = 1
    Else
        cmdOrder(0).Enabled = False
    End If
End Sub

Private Sub VTapi1_OnAnswered()
    VTapi1.PlaybackFile Combo2.Text
End Sub

Private Sub VTapi1_OnDebug(ByVal Message As String)
    List1.AddItem Message
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then Exit Sub
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
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
