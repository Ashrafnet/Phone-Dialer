VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«” Ì—«œ ÃÂ«  « ’«·"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "Õœœ „”«— «·„·› «·’Ê Ì"
      Top             =   2760
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Index           =   0
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      ToolTipText     =   "«” ⁄—«÷..."
      Top             =   2760
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   3480
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   3480
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
      Height          =   380
      Index           =   1
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«” Ì—«œ"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   380
      Index           =   0
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Õœœ «·„·› «·’Ê Ì"
      FileName        =   "*.wav,*.mp3"
      Filter          =   "*.wav"
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Õœœ „·› «·«ﬂ”·"
      Height          =   195
      Index           =   7
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   660
      Index           =   1
      Left            =   120
      Picture         =   "frmImport.frx":0000
      Top             =   720
      Width           =   660
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmImport.frx":3138
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
      Height          =   1275
      Index           =   6
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1200
      Width           =   5955
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«” Œœ„ Â–Â «·«œ«… ·«” Ì—«œ «·ÃÂ«  „‰ „·› «·«ﬂ”· «Ê „‰ „·›«  «·”Ì ›Ì «” "
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
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3585
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«” Ì—«œ ÃÂ«  « ’«·"
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
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
   Begin VB.Image Image1 
      Height          =   1155
      Index           =   0
      Left            =   5640
      Picture         =   "frmImport.frx":32D3
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄…"
      Height          =   195
      Index           =   5
      Left            =   2730
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„·"
      Height          =   195
      Index           =   4
      Left            =   2985
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   360
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„"
      Height          =   195
      Index           =   1
      Left            =   6255
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Width           =   330
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„"
      Height          =   195
      Index           =   16
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄«Ì‰… „« ”Ì „ «” Ì—«œÂ"
      Height          =   195
      Index           =   0
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4080
      Width           =   1605
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Index           =   8
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6975
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
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnus 
         Caption         =   "≈“«·… «·ÃÂ«  «·„Õœœ…"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmImport"
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

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            
        Case 1
           Unload Me
        Case 2
            
    End Select
End Sub

Private Sub Command2_Click(Index As Integer)
    CommonDialog1.Filter = "„·› «ﬂ”· (*.xls)|*.xls"
            CommonDialog1.ShowOpen
            
            If CommonDialog1.Filename <> "" Then
                On Error Resume Next
                If FileLen(CommonDialog1.Filename) < 1 Then Exit Sub
                Combo2.AddItem CommonDialog1.Filename
                Combo2.Text = CommonDialog1.Filename
            End If
End Sub

Private Sub Form_Load()
    OptemizeTitle
    For i = 0 To Combo1.Count - 1
        For j = 0 To Combo1.Count - 1
            Combo1(j).AddItem "«·⁄„Êœ —ﬁ„ " & i + 1
        Next j
        Combo1(i).Text = Combo1(0).List(i)
    Next i
End Sub

Private Sub ListView1_Click()
preview
End Sub

Private Sub mnus_Click(Index As Integer)
     On Error GoTo er:

    Select Case Index

        Case 0
            
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i - j).Selected = True Then
                    ListView1.ListItems.Remove ListView1.ListItems(i - j).Index
                    j = j + 1
                End If
            Next i
            
    End Select
        Exit Sub
er:

End Sub
Sub preview()
    MsgBox getExcel(1, "1", Combo2.Text)
End Sub
Function getExcel(rowval As Integer, columnval As String, excelfile As String)
'    Dim excelSheet As Object 'Excel Sheet object
    'Create an instance of Excel by file nam
    '     e
    Dim excelSheet As Excel.Worksheet
    
    Set excelSheet = CreateObject(excelfile)
    mycell$ = columnval & rowval
    getExcel = excelSheet.Range(mycell$).Value
    'Retrieve the result using the cell by r
    '     ow and column
    Set excelSheet = Nothing 'release object
End Function
Sub Export()
    On Error GoTo er:

        MousePointer = vbHourglass
        Dim xlApp As Excel.Application
        Dim xlWB As Excel.Workbook
        Dim xlWS As Excel.Worksheet
        
        Set xlApp = New Excel.Application
        Set xlWB = xlApp.Workbooks.Add
        Set xlWS = xlWB.Worksheets.Add
    
'''        If Check1(0).Value Then
'''            For i = 1 To ListView1.ListItems.Count
'''                xlWS.Cells(i, 1).Value = ListView1.ListItems(i).Text
'''
'''            Next i
'''        End If
'''        If Check1(1).Value Then
'''            For i = 1 To ListView1.ListItems.Count
'''                xlWS.Cells(i, 2).Value = ListView1.ListItems(i).SubItems(1)
'''            Next i
'''        End If
'''        If Check1(2).Value Then
'''            For i = 1 To ListView1.ListItems.Count
'''                xlWS.Cells(i, 3).Value = ListView1.ListItems(i).SubItems(2)
'''            Next i
'''        End If
'''        If Check1(3).Value Then
'''            For i = 1 To ListView1.ListItems.Count
'''                xlWS.Cells(i, 4).Value = ListView1.ListItems(i).SubItems(3)
'''            Next i
'''        End If
'''
'''        If Check1(4).Value Then
'''            xlWS.SaveAs Text1, , txtPass(0)
'''        Else
'''            xlWS.SaveAs Text1
'''        End If
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
