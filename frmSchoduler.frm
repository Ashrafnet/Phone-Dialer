VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4E0DBD51-5B94-4955-B721-E2F079EF9999}#23.0#0"; "Wave.ocx"
Begin VB.Form frmSchoduler 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„⁄«·Ã ÃœÊ·… «·„Â«„"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PhoneBazookaWaveCntl.objAudio objAudio 
      Left            =   5160
      Top             =   7560
      _ExtentX        =   1085
      _ExtentY        =   979
   End
   Begin VB.PictureBox pic 
      Height          =   5415
      Index           =   4
      Left            =   -5520
      RightToLeft     =   -1  'True
      ScaleHeight     =   5355
      ScaleWidth      =   6555
      TabIndex        =   28
      Top             =   4680
      Width           =   6615
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "«‰ ÂÏ «·„⁄«·Ã „‰ ⁄„·Ì…  Ã„Ì⁄ «·„⁄·Ê„«  «··«“„… ··„Â„… «·„ÃœÊ·… ﬁ„ »«·‰ﬁ— ⁄·Ï “— Õ›Ÿ Õ Ï  ﬂ„· «·⁄„·Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1920
         Width           =   3975
      End
   End
   Begin VB.PictureBox pic 
      Height          =   5415
      Index           =   3
      Left            =   2520
      RightToLeft     =   -1  'True
      ScaleHeight     =   5355
      ScaleWidth      =   6555
      TabIndex        =   23
      Top             =   1680
      Width           =   6615
      Begin VB.CommandButton Command2 
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   39
         ToolTipText     =   "„⁄«Ì‰… «·„·› «·’Ê Ì"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "down"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "«“«Õ… «·„·› «·„Õœœ «·Ï «”›·"
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "«“«Õ… «·„·› «·„Õœœ «·Ï «⁄·Ï"
         Top             =   2640
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6360
         Top             =   7680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Õœœ «·„·› «·’Ê Ì"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   33
         ToolTipText     =   "«“«·… «·„·› «·„Õœœ"
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         ToolTipText     =   "«—›«ﬁ „·› ’Ê Ì"
         Top             =   1920
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   840
         TabIndex        =   32
         ToolTipText     =   "«·„·›«  «·’Ê ÌÏ «· Ì ”Ì „ ‰ﬁ·Â« «·Ï ÃÂ«  «·« ’«·"
         Top             =   1560
         Width           =   5175
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   5640
         TabIndex        =   27
         Top             =   4920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text1"
         BuddyDispid     =   196613
         OrigLeft        =   3240
         OrigTop         =   3000
         OrigRight       =   3735
         OrigBottom      =   3255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Text            =   "1"
         Top             =   4920
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSchoduler.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌœ «·„·› «·’Ê Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   4
         Left            =   4530
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„—«  „⁄«Êœ… «·« ’«·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   3
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   3840
         Width           =   1875
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ„ » ÕœÌœ ⁄œœ „—«  „⁄«Êœ… «·« ’«· ›Ì Õ«· ›‘·  ⁄„·Ì… «·« ’«· »«Ì ÃÂ… „‰ «·ÃÂ« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   4200
         Width           =   5655
      End
   End
   Begin VB.PictureBox pic 
      Height          =   5415
      Index           =   0
      Left            =   1920
      RightToLeft     =   -1  'True
      ScaleHeight     =   5355
      ScaleWidth      =   6555
      TabIndex        =   6
      Top             =   1440
      Width           =   6615
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Text            =   "„Â„… ÃœÌœ…"
         ToolTipText     =   "÷⁄ Â‰« «”„  Ê÷ÌÕÌ ·ÿ»Ì⁄… Â–Â «·„Â„…"
         Top             =   4320
         Width           =   4455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«”„  Ê÷ÌÕÌ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   3960
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«Â·« »ﬂ„ ›Ì „⁄«·Ã ÃœÊ·… «·„Â«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.PictureBox pic 
      Height          =   5415
      Index           =   1
      Left            =   -1800
      RightToLeft     =   -1  'True
      ScaleHeight     =   5355
      ScaleWidth      =   6555
      TabIndex        =   8
      Top             =   1560
      Width           =   6615
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "‰›– Â–Â «·„Â„… ÌÊ„Ì«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   3240
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   63832065
         CurrentDate     =   38995
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   2760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   63832066
         CurrentDate     =   38995.8541666667
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”Ì „  ‰›Ì– Â–Â «·„Â„… » «—ÌŒ 5\5\2006 ›Ì  „«„ «·”«⁄… 12:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   4
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   4080
         Width           =   5295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   855
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   3960
         Width           =   5535
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ„ » ÕœÌœ “„‰  ‰›Ì– «·„Â„… «·Õ«·Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ„ » ÕœÌœ  «—ÌŒ  ‰›Ì– «·„Â„…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌœ «·Êﬁ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌœ «· «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOrder 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
      Height          =   380
      Index           =   2
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "«· «·Ì>>"
      Height          =   380
      Index           =   1
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "<<«·”«»ﬁ"
      Enabled         =   0   'False
      Height          =   380
      Index           =   0
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox pic 
      Height          =   5415
      Index           =   2
      Left            =   3120
      RightToLeft     =   -1  'True
      ScaleHeight     =   5355
      ScaleWidth      =   6555
      TabIndex        =   17
      Top             =   2040
      Width           =   6615
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Õ–›"
         Height          =   375
         Index           =   1
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "«“«·… Ã„Ì⁄ «·ÃÂ«  «·„Õœœ… „‰ «·ﬁ«∆„…"
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "«÷«›….."
         Height          =   375
         Index           =   0
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   4920
         Width           =   735
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "«·«”„"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "«·—ﬁ„"
            Object.Width           =   2540
         EndProperty
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   3
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmSchoduler.frx":0095
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmSchoduler.frx":30E7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmSchoduler.frx":3439
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSchoduler.frx":398B
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌœ «·ÃÂ«  «· Ì ”Ì „ «·« ’«· »Â«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   3135
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   120
         Width           =   3240
      End
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   240
      Picture         =   "frmSchoduler.frx":3A3B
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSchoduler.frx":8BDD
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
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   7785
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÃœÊ·… «·„Â«„"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   8
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10815
   End
   Begin VB.Line lin 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   11400
      X2              =   14520
      Y1              =   3360
      Y2              =   2880
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   11640
      X2              =   13200
      Y1              =   2760
      Y2              =   2400
   End
End
Attribute VB_Name = "frmSchoduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Currpic As Integer
Private Sub Check1_Click()
    If Check1.Value Then
        DTPicker1.Enabled = 0
        lblInfo(4) = "”Ì „  ‰›Ì– Â–Â «·„Â„… ﬂ· ÌÊ„ ›Ì  „«„ «·”«⁄… " & Format(DTPicker2.Value, "HH:MM:SS")
    Else
        lblInfo(4) = "”Ì „  ‰›Ì– Â–Â «·„Â„… » «—ÌŒ " & DTPicker1.Value & " ›Ì  „«„ «·”«⁄… " & Format(DTPicker2.Value, "HH:MM:SS")
        DTPicker1.Enabled = 1
    End If
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    Select Case Index
        Case 0
            frmContacs.IsEdit = 0
            frmContacs.Show 1
        Case 1
            On Error Resume Next
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i - j).Selected Then
                    ListView1.ListItems.Remove i - j
                    j = j + 1
                End If
            Next i
            ListView1.SetFocus
            ListView1.ListItems(1).Selected = True
    End Select
End Sub

Private Sub cmdOrder_Click(Index As Integer)
    Select Case Index
        Case 2
            Unload Me
        Case 1
            If Currpic = 2 Then
                If ListView1.ListItems.Count < 1 Then
                    MsgBox "ÌÃ» «‰  ÷Ì› ÃÂ… « ’«· ⁄·Ï «·«ﬁ·", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
                    Exit Sub
                End If
            End If
            If Currpic = 3 Then
                If List1.ListCount < 1 Then
                    MsgBox "ÌÃ» «‰  ÷Ì› „·› ’Ê Ì ⁄·Ï «·«ﬁ·", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
                    Exit Sub
                End If
            End If

            If Currpic = 3 Then cmdOrder(1).Caption = "Õ›Ÿ"
            If Currpic = 4 Then SetSchduler: Exit Sub
            cmdOrder(0).Enabled = True
            Currpic = Currpic + 1
            pic(Currpic).ZOrder
        Case 0
            cmdOrder(1).Caption = "«· «·Ì>>"
            If Currpic = 1 Then cmdOrder(0).Enabled = 0
            
            cmdOrder(1).Enabled = True
            Currpic = Currpic - 1
            pic(Currpic).ZOrder
    End Select
End Sub
Sub SetSchduler()
    On Error GoTo er:
    Dim rs As New Recordset
    Dim rsX As New Recordset
    recall = Text1
    schtime = Format(DTPicker2.Value, "HH:MM:SS")
    schdate = Format(DTPicker1.Value, "dd-mm-yyyy")

    rsX.Open "select schid from schmaster where schdate='" & schdate & "' and schtime='" & schtime & "'", db, adOpenStatic
    If rsX.RecordCount > 0 Then MsgBox "ÌÊÃœ „Â„… „ÃœÊ·… »‰›” «· «—ÌŒ Ê«· ÊﬁÌ  «·„Õœœ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
    For i = 0 To List1.ListCount - 1
        schsnd = schsnd & List1.List(i) & ","
    Next i
    schsnd = Replace(schsnd, "'", "''")
    If Check1.Value Then
        schdate = "1"
    End If
    rs.Open "insert into schmaster(schtime,schdate,schsnd,recall,schname) Values ('" & schtime & "','" & schdate & "','" & schsnd & "','" & recall & "','" & txtName & "')", db
    rsX.Close
    rsX.Open "select schid from schmaster where schdate='" & schdate & "' and schtime='" & schtime & "'", db
    SchID = rsX!SchID
    For i = 1 To ListView1.ListItems.Count
        rs.Open "insert into schslave(schid,cname,cno) Values ('" & SchID & "','" & ListView1.ListItems(i).Text & "','" & ListView1.ListItems(i).SubItems(1) & "')", db
    Next i
    Unload Me
    Exit Sub
er:
    MsgBox "ﬁ„ » €ÌÌ— «”„ «·„Â„…", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
    pic(0).ZOrder
    Currpic = 0
    cmdOrder(0).Enabled = False
    cmdOrder(1).Caption = "«· «·Ì>>"
    'Resume
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
                List1.AddItem CommonDialog1.Filename
                List1.Text = CommonDialog1.Filename
            End If

        Case 1
            On Error Resume Next
            If List1.ListCount < 1 Then Exit Sub
            List1.RemoveItem List1.ListIndex
            List1.Text = List1.List(0)
            
        Case 2
            On Error Resume Next
            xx = List1.ListIndex
            If xx = 0 Then Exit Sub
            List1.AddItem List1.Text, List1.ListIndex - 1
            List1.RemoveItem List1.ListIndex
            List1.SetFocus
            List1.Selected(xx - 1) = True
            
        Case 3
            'On Error Resume Next
            xx = List1.ListIndex
            If xx = List1.ListCount - 1 Then Exit Sub
            List1.AddItem List1.Text, List1.ListIndex + 2
            List1.RemoveItem List1.ListIndex
            List1.SetFocus
            List1.Selected(xx + 1) = True
        Case 4
            If List1.Text <> "" Then
                sndPlaySound List1.Text, 1
            Else
                MsgBox "«Œ — „·› ’Ê  „‰ «·ﬁ«∆„… «Ê·«", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading
            End If

    End Select
End Sub
Function IsINList(str As String) As Boolean
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = str Then IsINList = True: Exit Function
    Next i
End Function
Private Sub DTPicker1_Change()
    lblInfo(4) = "”Ì „  ‰›Ì– Â–Â «·„Â„… » «—ÌŒ " & DTPicker1.Value & " ›Ì  „«„ «·”«⁄… " & Format(DTPicker2.Value, "HH:MM:SS")
End Sub

Private Sub DTPicker2_Change()
Check1_Click
End Sub

Private Sub Form_Load()
    
    Currpic = 0
    pic(Currpic).ZOrder
    OptemizeTitle
    DTPicker1.Value = Date
End Sub
Sub OptemizeTitle()
    
    Icon = Nothing
    lblT(8).Top = 0
    lblT(8).Left = 0
    lblT(8).Width = Me.Width
    lblT(8).ZOrder 1
    For i = 0 To pic.Count - 1
        pic(i).Top = 1680
        pic(i).Left = 120
        pic(i).BorderStyle = 0
    Next i
    For i = 0 To lin.Count - 1
        lin(i).X1 = 0
        lin(i).X2 = Me.Width
        lin(i).Y1 = lblT(8).Height + 2
        lin(i).Y2 = lblT(8).Height + 2
    Next i
    lin(0).ZOrder 1
    lblInfo(2).Left = Me.Width - lblInfo(2).Width - 300
    LVFullRowSelect Me.ListView1

End Sub

