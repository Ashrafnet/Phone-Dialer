VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmStatics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«Õ’«∆Ì«  «·—”«∆·(ﬁÌœ «·«‰‘«¡)"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "«Õ’«∆Ì«  —ﬁ„Ì…"
      Height          =   375
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   8040
      Width           =   1455
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5295
      Left            =   120
      OleObjectBlob   =   "frmStatics.frx":0000
      TabIndex        =   5
      Top             =   2640
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Õ”‰«"
      Height          =   380
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStatics.frx":287F
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
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4305
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStatics.frx":293D
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
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   5745
   End
   Begin VB.Image Image1 
      Height          =   660
      Index           =   1
      Left            =   120
      Picture         =   "frmStatics.frx":2A37
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«Õ’«∆Ì«  «·—”«∆·"
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
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   1155
      Index           =   0
      Left            =   5160
      Picture         =   "frmStatics.frx":5B6F
      Top             =   120
      Width           =   1155
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   8160
      X2              =   14400
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   8160
      X2              =   14400
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Index           =   8
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmStatics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
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
    

End Sub

Private Sub Form_Load()
    OptemizeTitle
    LoadGraphData
    MSChart1.chartType = VtChSeriesType3dBar
    MSChart1.Plot.Projection = VtProjectionTypeOblique
    AddLegend MSChart1, True
End Sub

Sub LoadGraph()
'    MSChart1.ChartData
End Sub

Sub LoadGraphData()
On Error Resume Next

'  This subroutine shows you how to setup and pass data to the chart object
'  imax is the maximum of columns that you want to display as demo data
'  datascale will adjust the actual values to suit a different axis scale

   imax = 6
   Dim x() As Variant
   Dim iRow As Integer
  
   ReDim x(1 To 8, 1 To imax + 1)
  
    
  dataScale = 1
  
  iRow = 1
  x(1, iRow) = "Period"
  x(2, iRow) = "June 1998"
  x(3, iRow) = "June 1999"
  x(4, iRow) = "June 2000"
  x(5, iRow) = "June 2001"
  x(6, iRow) = "June 2002"
  x(7, iRow) = "June 2003"
  x(8, iRow) = "June 2004"
  
  For iRow = 5 To imax + 1
  
  x(1, iRow) = "Ann Depr%" & iRow
  x(2, iRow) = (12.2 + (iRow * Rnd) * 20) * dataScale
  x(3, iRow) = (45 + (iRow * Rnd) * 20) * dataScale
  x(4, iRow) = (36 + (iRow * Rnd) * 20) * dataScale
  x(5, iRow) = (28 + (iRow * Rnd) * 20) * dataScale
  x(6, iRow) = (38 + (iRow * Rnd) * 20) * dataScale
  x(7, iRow) = (25 + (iRow * Rnd) * 20) * dataScale
  x(8, iRow) = (16 + (iRow * Rnd) * 20) * dataScale
  
  Next iRow
  

addChartData:

' Reset the chart back to default to avoid any surprises
  MSChart1.ToDefaults
    
 
  Call addDataArray(MSChart1, x(), True)
End Sub
