VERSION 5.00
Begin VB.Form frmAddContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«÷«›… ÃÂ… « ’«·"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "«÷«›… „Ã„Ê⁄…"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3000
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton cmdOrder 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
      Height          =   375
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "«÷›"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "frmAddContact.frx":0000
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ — «·„Ã„Ê⁄… Ê„‰ À„ «ﬂ » «·«”„ À„ «·—ﬁ„ «·Â« ›Ì «Ê —ﬁ„ «·ÃÊ«· Ê„‰ À„ «ﬂ » ⁄‰Ê«‰ «·⁄„· «–« ﬂ«‰ „ Ê›—« À„ «‰ﬁ— ⁄·Ï “— «÷›"
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
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   480
      Width           =   4305
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«÷«›… ÃÂ… « ’«· ÃœÌœ…"
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
      TabIndex        =   11
      Top             =   240
      Width           =   2190
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Index           =   8
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6495
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„·"
      Height          =   195
      Index           =   3
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„"
      Height          =   195
      Index           =   2
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„"
      Height          =   195
      Index           =   1
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄…"
      Height          =   195
      Index           =   0
      Left            =   4380
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "frmAddContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cID As Long
Private Sub cmdOrder_Click(Index As Integer)
    Select Case Index
        Case 0
            addContact Text1(1), Combo1.ItemData(Combo1.ListIndex), Text1(0), Text1(2)
        Case 1
            Unload Me
        
    End Select
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

Sub addContact(cName As String, Gid As Long, No As String, Work As String)
    On Error GoTo er:
       Dim rs As New Recordset
       Dim rsUp As New Recordset
       rs.Open "select id from contact where id=" & cID, db, adOpenStatic
       If rs.RecordCount > 0 Then
            rsUp.Open "update contact set gid='" & Gid & "',name='" & cName & "',phone  ='" & No & "',mywork='" & Work & "' where id=" & rs!ID, db
            MsgBox " „   ⁄œÌ· «·ÃÂ… " & cName & " »‰Ã«Õ ", vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation
            Unload Me
        Else
            
            rs.Close
            rs.Open "insert into contact(gid,name,phone,mywork) Values (" & Gid & ",'" & cName & "','" & No & "','" & Work & "')", db
            Text1(0) = "": Text1(1) = "": Text1(2) = "": cmdOrder(0).Enabled = False: Text1(1).SetFocus
            MsgBox " „  «÷«›… «·ÃÂ… " & cName & " »‰Ã«Õ ", vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation
        End If
       
    Exit Sub
er:
    MsgBox Err.Description
    Resume
End Sub

Private Sub Command1_Click()
    frmAddG.Show 1
    LoadGroups
End Sub

Private Sub Form_Load()
    OptemizeTitle
    LoadGroups
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

Private Sub Text1_Change(Index As Integer)
    If Combo1.Text = "" Then cmdOrder(0).Enabled = False: Exit Sub
    If Trim(Text1(1).Text) <> "" And Trim(Text1(0).Text) <> "" Then
        cmdOrder(0).Enabled = True
    Else
        cmdOrder(0).Enabled = False
    End If
End Sub
