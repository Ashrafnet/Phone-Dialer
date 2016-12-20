VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{4E0DBD51-5B94-4955-B721-E2F079EF9999}#23.0#0"; "Wave.ocx"
Object = "{797E7185-0DB7-4E3A-939B-234871F7FAC9}#1.11#0"; "Apax1.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "äÙÇã ÇáÑÓÇÆá ÇáÕæÊíÉ ÚÈÑ ÇáåÇÊÝ"
   ClientHeight    =   9030
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   RightToLeft     =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin dailer.isExplorerBar isExplorerBar1 
      Align           =   3  'Align Left
      Height          =   8235
      Left            =   3600
      TabIndex        =   1
      Top             =   420
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   14526
      FontName        =   "MS Sans Serif"
      FontCharset     =   178
      UxThemeText     =   0   'False
      Begin APAX1.Apax Apax1 
         Height          =   2055
         Left            =   0
         TabIndex        =   12
         Top             =   5280
         Visible         =   0   'False
         Width           =   3615
         Baud            =   19200
         ComNumber       =   0
         DeviceType      =   0
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
         SelectedDevice  =   ""
         SilenceThreshold=   50
         TapiNumber      =   ""
         TapiRetryWait   =   60
         TrimSeconds     =   2
         UseSoundCard    =   0   'False
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
         Caption         =   "Ashraf Net 4 Programming"
         CaptionAlignment=   0
         CaptionWidth    =   132
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
         ShowTerminalButtons=   0   'False
         DoubleBuffered  =   0   'False
         Enabled         =   -1  'True
         Cursor          =   0
         TapiStatusDisplay=   -1  'True
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
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1680
         Top             =   4560
      End
      Begin PhoneBazookaWaveCntl.objAudio objAudio 
         Left            =   840
         Top             =   6600
         _ExtentX        =   1085
         _ExtentY        =   979
      End
      Begin ComctlLib.ImageList ImageList2 
         Left            =   720
         Top             =   3840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8655
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "09/10/1427"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   1
            Object.Width           =   12726
            TextSave        =   "07:33 ã"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin dailer.isExplorerBar isExplorerBar2 
      Align           =   3  'Align Left
      Height          =   8235
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   14526
      FontName        =   "MS Sans Serif"
      FontCharset     =   178
      UxThemeText     =   0   'False
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   120
         RightToLeft     =   -1  'True
         ScaleHeight     =   4215
         ScaleWidth      =   3375
         TabIndex        =   8
         Top             =   3720
         Width           =   3375
         Begin ComctlLib.ListView ListView1 
            Height          =   3495
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   6165
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
               Text            =   "ã"
               Object.Width           =   411
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "ÇáÇÓã"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "ÇáÑÞã"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "ÇáÚãá"
               Object.Width           =   0
            EndProperty
         End
         Begin ComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   -120
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
                  Picture         =   "mdiMain.frx":0000
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "mdiMain.frx":0552
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÚÑÖ ÇáäÊÇÆÌ ßÇãáÉ"
            Height          =   195
            Left            =   1890
            MouseIcon       =   "mdiMain.frx":0AA4
            MousePointer    =   99  'Custom
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   3720
            Width           =   1245
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         RightToLeft     =   -1  'True
         ScaleHeight     =   1575
         ScaleWidth      =   3375
         TabIndex        =   3
         Top             =   1800
         Width           =   3375
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÈÍË ÈÏáÇáÉ ÇáÑÞã"
            Height          =   255
            Index           =   1
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÈÍË ÈÏáÇáÉ ÇáÇÓã"
            Height          =   255
            Index           =   0
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   720
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Caption         =   "ÈÍË"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1
Public CurrWaveFiles As String
Public CurrPhone As String ' the current phone no is calling
Public CurrActiveSch As String ' the current Active Schdulaer

Private Sub Apax1_OnTapiConnect()
    If Apax1.EnableVoice Then
        StatusBar1.Panels(2).Text = "Êã ÇäÔÇÁ ÇáÇÊÕÇá"
        Apax1.TapiPlayWaveFile CurrWaveFile
    Else
        StatusBar1.Panels(2).Text = "Êã ÇäÔÇÁ ÇáÇÊÕÇá áßä ÇáãæÏã áÇ íÏÚã ÇáÕæÊ"
    End If
    
End Sub

Private Sub Apax1_OnTapiDTMF(ByVal Digit As Byte, ByVal ErrorCode As Long)
    StatusBar1.Panels(2).Text = "áÞÏ äÞÑ ÇáãÊÕá ÇáÈÚíÏ Úáì ÒÑ " & CStr(Digit - 48)
End Sub

Private Sub Apax1_OnTapiFail()
    If Apax1.TapiCancelled Then
      StatusBar1.Panels(2).Text = "Êã ÇáÛÇÁ ÇáÇÊÕÇá"
    Else
      StatusBar1.Panels(2).Text = ("ÝÔá ÇáÇÊÕÇá")
    End If
End Sub

Private Sub Apax1_OnTapiStatus(ByVal First As Boolean, ByVal Last As Boolean, ByVal Device As Long, ByVal message As Long, ByVal Param1 As Long, ByVal Param2 As Long, ByVal Param3 As Long)
    StatusBar1.Panels(2).Text = TapiStateStr(Apax1.TapiState)
    StatusBar1.Panels(1).Text = Apax1.TapiStatusMsg(message, Param1, Param2)
    If InStr(1, StatusBar1.Panels(1).Text, "retry") Then
        StatusBar1.Panels(2).Text = "ÇäÊÙÇÑ ÇÚÇÏÉ ãÍÇæáÉ ÇáÇÊÕÇá"
    End If
End Sub

Private Sub Apax1_OnTapiWaveNotify(ByVal MSG As Apax1.TxWaveMessage)
  Dim s As String
  Select Case MSG
    Case waPlayOpen: s = "íÊã ÇáÇä ÊÔÛíá ÇáãáÝ ÇáÕæÊí"
    Case waPlayDone: s = "ÇäÊåì ÇáãáÝ ÇáÕæÊí ãä ÇáÞÑÇÁÉ"
    Case waPlayClose: s = "Êã ÇÛáÇÞ ÇáãáÝ ÇáÕæÊí"
    Case waRecordOpen: s = "ÈÏÁ ÇáÊÓÌíá"
    Case waDataReady: s = "waDataReady"
    Case waRecordClose: s = "ÇäÊåÊ ÚãáíÉ ÇáÊÓÌíá"
  End Select
  StatusBar1.Panels(2).Text = s
End Sub

Private Sub Apax1_OnTapiWaveSilence(StopRecording As Boolean, Hangup As Boolean)
'åÐÇ ÇáÍÏË íäÔÇÁ ÇÐÇ ßÇä ÇáãÊÕá ÇáÈÚíÏ ÞÏ ÊæÞÝ Úä ÇáßáÇã , Ýíãßäß Çä ÊÞæã ÈÚãá ÊæÞÝ áÚãáíÉ ÊÓÌíá ÇáÕæÊ , Çæ ÞØÚ ÇáãßÇáãÉ
    StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & "  ÍÏË Wave Silence" & "       StopRecording Value : " & StopRecording & "       Hangup Value : " & Hangup
End Sub
Private Sub Command1_Click()
    SearchContacts Text1
End Sub
Sub SearchContacts(strToSearch As String)
    On Error GoTo er:
    Dim rs As New Recordset
    Dim i As Integer
    If Option1(0).Value Then
        rs.Open "select * from contact where name like'%" & strToSearch & "%'", db
    Else
        rs.Open "select * from contact where phone like'%" & strToSearch & "%'", db
    End If
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
    Resume
End Sub


Private Sub isExplorerBar2_ItemClick(sGroup As String, sItemKey As String)
    Select Case sItemKey
        Case 0
        
        Case 12
            isExplorerBar1.Visible = 1
            isExplorerBar2.Visible = 0

    End Select
End Sub

Private Sub Label1_Click()
    BringWindowToTop frmSearch.Hwnd
    frmSearch.Show
    frmSearch.ListView1.ListItems.Clear
    For i = 1 To ListView1.ListItems.Count
        Dim xx As ListItem
        Set xx = frmSearch.ListView1.ListItems.Add(, ListView1.ListItems(i).key, i, 1, 1)
        xx.SubItems(1) = ListView1.ListItems(i).SubItems(1)
        xx.SubItems(2) = ListView1.ListItems(i).SubItems(2)
        xx.SubItems(3) = ListView1.ListItems(i).SubItems(3)
    Next i

    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.ForeColor = vbRed
End Sub

Private Sub MDIForm_Load()
    isExplorerBar1.Align = 4
    isExplorerBar2.Align = 4
    isExplorerBar2.Visible = False
    BuildMenu
    ConnectToDB "db.mdb"
    
    If Command = "/min" Then
        Hide
    Else
        Show
        frmLogin.Show 1
        'frmMian.Show
        
    End If
    SysTry
    LoadOptions
End Sub
Sub SysTry()
        Set m_frmSysTray = New frmSysTray
        With m_frmSysTray
            .AddMenuItem "ÚÑÖ ÇáÈÑäÇãÌ  ", "open", True
            .AddMenuItem "-"
            .AddMenuItem "ÎÑæÌ", "close"
            .ToolTip = "äÙÇã ÇáÑÓÇÆá ÇáÕæÊíÉ"
        End With

End Sub

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
   Select Case sKey
   Case "open"
      Me.Show
      Me.ZOrder
        frmLogin.Show 1
   Case "close"
        Unload m_frmSysTray
        Set m_frmSysTray = Nothing
      End
   Case Else
      MsgBox "Clicked item with key " & sKey, vbInformation
   End Select

End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
On Error Resume Next
      Me.Show
      Me.ZOrder
    frmLogin.Show 1
End Sub


Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then
        m_frmSysTray.ShowMenu
    End If

End Sub
Sub BuildMenu()
    
    isExplorerBar1.AddSpecialGroup "ÇÏÇÑÉ ÌåÇÊ ÇáÇÊÕÇá"
    isExplorerBar1.AddItem -1, "0", "ÇÖÇÝÉ ãÌãæÚÉ ÌÏíÏÉ"
    isExplorerBar1.AddItem -1, "1", "ÇÖÇÝÉ ÌåÉ ÇÊÕÇá ÌÏíÏÉ"
    isExplorerBar1.AddItem -1, "2", "ÈÍË Úä ÌåÉ ÇÊÕÇá"
    isExplorerBar1.AddItem -1, "7", "ÍÐÝ ãÌãæÚÉ Çæ ÌåÉ ÇÊÕÇá"
    isExplorerBar1.AddItem -1, "3", "ÇÏÇÑÉ ÌåÇÊ ÇáÇÊÕÇá"
    
    isExplorerBar1.AddGroup "G2", "ÎÏãÇÊ ÇáÑÓÇÆá ÇáÕæÊíÉ"
    isExplorerBar1.AddItem "G2", "4", "ÇÑÓÇá ÑÓÇáÉ ÕæÊíÉ áÔÎÕ"
    isExplorerBar1.AddItem "G2", "5", "ÇÑÓÇá ÑÓÇáÉ ÕæÊíÉ áÚÏÉ ÇÔÎÇÕ"
    isExplorerBar1.AddItem "G2", "6", "ÇÍÕÇÆíÇÊ ÇáÑÓÇÆá "
    isExplorerBar1.AddItem "G2", "8", "ÊÚÞÈ ÇáÑÓÇÆá ÇáÕæÊíÉ "
    isExplorerBar1.AddItem "G2", "11", "ãÓÌá ÇáÕæÊ "
    'isExplorerBar1.AddItem "G2", "13", "ÇäÕÇÊ ááãßÇáãÇÊ ÇáæÇÑÏÉ "
    
    isExplorerBar1.AddGroup "G3", "ÎÏãÇÊ ÌÏæáÉ ÇáãåÇã"
    isExplorerBar1.AddItem "G3", "9", "ãÚÇáÌ ÌÏæáÉ ãåÇã ÇáÑÓÇÆá ÇáÕæÊíÉ "
    isExplorerBar1.AddItem "G3", "10", "ÇÏÇÑÉ ÌÏæáÉ ÇáãåÇã "
    
'    isExplorerBar1.AddItem "G3", "13", "ÍÐÝ ãåãÉ ãÌÏæáÉ "
    
    isExplorerBar1.AddDetailsGroup "ÊÝÇÕíá", "  ÎÏãÇÊ ÇáÑÓÇÆá ÇáÕæÊíÉ", "ÚÈÇÑÉ Úä äÙÇã ßÇãá áÇÑÓÇá æÇÓÊÞÈÇá ÇáãßÇáãÇÊ ÇáÕæÊíÉ ÚÈÑ ÇáãæÏã ÇáãÊÕá ÈÌåÇÒ ÇáÍÇÓæÈ, íãßäß åÐÇ ÇáäÙÇã ãä ÇÑÓÇá ãÞÇØÚ ÕæÊíÉ áÌåÇÊ ÊÞæã ÇäÊ ÈÊÍÏíÏåÇ Ýí ÇáæÞÊ ÇáÐí ÊÑíÏå ßãÇ íãßäß ãä Úãá ÇÑÔíÝ áßá ÇáãßÇáãÇÊ ÇáÕÇÏÑÉ æÇáæÇÑÏÉ ÛíÑåÇ ãä ÇáããíÒÇÊ ."

    isExplorerBar2.AddGroup "SearchParameters", "ÈÍË Úä ÌåÉ ÇÊÕÇá", 1
        
    '5: Now, Attach the picture Box to the group.
    isExplorerBar2.SetGroupChild "SearchParameters", Picture1
    

    isExplorerBar2.AddGroup "searchresult", "äÊÇÆÌ ÇáÈÍË", 1
    isExplorerBar2.SetGroupChild "searchresult", Picture2
    
    isExplorerBar2.AddGroup "links", "ãæÇÖÚ ÇÎÑì"
    isExplorerBar2.AddItem "links", "12", "ÚæÏÉ Çáì ÇáÞÇÆãÉ ÇáÑÆíÓíÉ"
    
End Sub

Private Sub isExplorerBar1_ItemClick(sGroup As String, sItemKey As String)
On Error Resume Next
    Select Case sItemKey
        Case 0
             frmAddG.Show 1
             If IsfrmMainLoaded Then frmMian.LoadGroups
        Case 1
            frmAddContact.Show 1
            If IsfrmMainLoaded Then frmMian.LoadContacts frmMian.List1.ItemData(frmMian.List1.ListIndex)
            
            
        Case 2
            isExplorerBar1.Visible = False
            isExplorerBar2.Visible = True
            Text1.SetFocus
        Case 3 ' Contacts Manager
            BringWindowToTop frmMian.Hwnd
            frmMian.Show
        Case 4 'send one
            frmSendOne.Show 1
        Case 5 ' send many
            BringWindowToTop frmCall.Hwnd
            frmCall.Show
        Case 6 ' Statics
            frmStatics.Show 1
        Case 7 '
            frmDelGroup.Show 1
            If IsfrmMainLoaded Then frmMian.LoadGroups: frmMian.LoadContacts frmMian.List1.ItemData(frmMian.List1.ListIndex)
        Case 8 ' Log
            BringWindowToTop frmLog.Hwnd
            frmLog.Show
            frmLog.ShowLogs
            
    Case 9
        frmSchoduler.Show 1
    Case 10
        BringWindowToTop frmShowSch.Hwnd
        frmShowSch.Show
    Case 11
        frmRecord.Show 1
        
    Case 13 ' del Jop
        BringWindowToTop frmShowSch.Hwnd
        frmShowSch.Show
    End Select
End Sub

Private Sub MDIForm_Initialize()
    If App.PrevInstance = True Then MsgBox "áÇ íãßä ÊÔÛíá ÇßËÑ ãä ÈÑäÇãÌ", vbCritical: End
    InitCommonControlsXP
    
End Sub
Sub LoadOptions()
On Error Resume Next
  mdiMain.Apax1.MaxAttempts = GetSetting("VLS", "Apax", "txtMaxAttempts", txtMaxAttempts)
  mdiMain.Apax1.TapiRetryWait = GetSetting("VLS", "Apax", "txtRetryWait", txtRetryWait)
  mdiMain.Apax1.MaxMessageLength = GetSetting("VLS", "Apax", "txtMaxMessageLength", txtMaxMessageLength)
  mdiMain.Apax1.EnableVoice = GetSetting("VLS", "Apax", "chkEnableVoice", chkEnableVoice.Value)
  mdiMain.Apax1.InterruptWave = GetSetting("VLS", "Apax", "chkInterruptWave", chkInterruptWave.Value)
  mdiMain.Apax1.UseSoundCard = GetSetting("VLS", "Apax", "chkUseSoundCard", chkUseSoundCard.Value)
  mdiMain.Apax1.SelectedDevice = GetSetting("VLS", "Apax", "txtModem", txtModem)
  
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Cancel = True
    Hide
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.ForeColor = vbBlue
End Sub

Private Sub Text1_Change()
    Command1.Enabled = Len(Trim(Text1)) > 0
    Text1_GotFocus
End Sub

Private Sub Text1_GotFocus()
    Command1.Default = Len(Trim(Text1)) > 0
End Sub

Private Sub Text1_LostFocus()
    Command1.Default = 0
End Sub
Function IsPhoneWave(WavePath As String, Optional NoMSG As Boolean = False) As Boolean
    Dim xx As String
    If objAudio.OpenWaveFile(WavePath) Then
        If (objAudio.GetWaveFormat.wFormatTag = 1) And (objAudio.GetWaveFormat.wBitsPerSample = 16) And (objAudio.GetWaveFormat.nSamplesPerSec = 8000) Then
            
            objAudio.SetWaveFormat PCM_8kHz16bit_Voice_Modems
            IsPhoneWave = True
        Else
            
        xx = "ÊäÓíÞ ÇáãáÝ ÇáÍÇáí :" & vbNewLine & "           "
'        xx = xx & "" & objAudio.GetWaveFormat.wFormatTag & vbNewLine & "           "
        xx = xx & "Bits per sample = " & objAudio.GetWaveFormat.wBitsPerSample & vbNewLine & "           "
        xx = xx & "Samples per second = " & objAudio.GetWaveFormat.nSamplesPerSec
 
            
            IsPhoneWave = False
            If NoMSG Then Exit Function
            MsgBox "ÎØÃ: íÌÈ Çä ÊÎÊÇÑ ãáÝ ÕæÊí ãÊæÇÝÞ ãÚ ÇáãáÝÇÊ ÇáÊí íãßä áÎØ ÇáåÇÊÝ Çä íÍãáåÇ æåí ãáÝÇÊ ÇáÕæÊ ÇáÊí Êßæä ÈÇáÕíÛÉ ÇáÊÇáíÉ :  " & vbCr & vbLf & vbLf & "PCM 16-bit 8Khz" & vbNewLine & vbNewLine & " " & xx, _
                vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
        End If
    Else
        MsgBox "Error: Couldn't find wave file." & vbCr & vbLf & vbLf, _
            vbOKOnly + vbCritical
    End If

End Function
Public Sub MakeCalls(phonesNo As String, WaveFilesToPlay As String)
    Dim cno() As String
    cno = Split(phonesNo, ",")
    For Each No In cno
        MakeCall No, WaveFilesToPlay
    Next
    
End Sub
Public Sub MakeCall(phoneNo As String, WaveFilesToPlay As String)
    CurrWaveFiles = WaveFilesToPlay
    Apax1.TapiNumber = phoneNo
    StatusBar1.Panels(2).Text = "íÊã ÇáÇä ãÍÇæáÉÇáÇÊÕÇá"
    Apax1.TapiDial
End Sub

Private Sub Timer1_Timer()
    CurrActiveSch = GetCurrSchID
    If CurrActiveSch > 0 Then
        Timer1.Enabled = False
    End If
    
End Sub
Sub MakeSchCall(SchID As Long)
    On Error GoTo er:
    Dim rs As New Recordset
    Dim rsx As New Recordset
    rs.Open "select * from schmaster where schid =" & SchID, db, adOpenStatic
    rs.Open "select * from schslave where schid =" & SchID, db, adOpenStatic
    If rs.RecordCount = 0 Then Exit Function
    If rsx.RecordCount = 0 Then Exit Function
    
    While Not rsx.EOF
        xx = xx & "," & rsx!cno
        rsx.MoveNext
    Wend
    
    MakeCalls xx, rs!schsnd
Exit Function
er:
'    MsgBox Err.Description
    Resume Next

End Sub
Function GetCurrSchID() As Long
On Error GoTo er:
    Dim rs As New Recordset
    rs.Open "select * from schmaster where active ='1' ", db, adOpenStatic
    If rs.RecordCount = 0 Then GetCurrSchID = -1: Exit Function
    If Format(rs!schdate, "dd-mm-yyyy") = Format(Date, "dd-mm-yyyy") Then
        If Format(rs!schtime, "hh:mm") = Format(time, "hh:mm") Then
            GetCurrSchID = rs!SchID
        End If
        
    End If
    
Exit Function
er:
'    MsgBox Err.Description
'    Resume
End Function
