VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4E0DBD51-5B94-4955-B721-E2F079EF9999}#23.0#0"; "Wave.ocx"
Begin VB.Form frmEditSch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Õ—Ì— „Â„… „ÃœÊ·…"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      Height          =   5465
      Index           =   5
      Left            =   5760
      RightToLeft     =   -1  'True
      ScaleHeight     =   5400
      ScaleWidth      =   6720
      TabIndex        =   40
      Top             =   3960
      Width           =   6785
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- «·„Â„… ‰‘ÿ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Â„… „⁄ÿ·…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   10
         Left            =   5550
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- ·„ Ì „ «—”«· Â–Â «·„Â„… »⁄œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   600
         Width           =   2340
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- ·„ Ì „ «—”«· Â–Â «·„Â„… »⁄œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   840
         Width           =   2340
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Õ«·… «·«—”«·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   7
         Left            =   5610
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.PictureBox pic 
      Height          =   5465
      Index           =   4
      Left            =   5160
      RightToLeft     =   -1  'True
      ScaleHeight     =   5400
      ScaleWidth      =   6720
      TabIndex        =   34
      Top             =   4080
      Width           =   6785
      Begin VB.CommandButton Command1 
         Caption         =   "«⁄«œ…  ”„Ì… «·„Â„… «·Õ«·Ì…"
         Height          =   375
         Index           =   2
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   " ⁄ÿÌ· «·„Â„… «·Õ«·Ì…"
         Height          =   375
         Index           =   1
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Õ–› «·„Â„… «·Õ«·Ì…"
         Height          =   375
         Index           =   0
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ŒÌ«—«  «·„Â„… «·Õ«·Ì…"
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
         Index           =   5
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Â–Â „Ã„Ê⁄… „‰ «·ŒÌ«—«  «·Œ«’… »«·„Â„… «·Õ«·Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   3765
      End
   End
   Begin VB.PictureBox pic 
      Height          =   5465
      Index           =   1
      Left            =   4680
      RightToLeft     =   -1  'True
      ScaleHeight     =   5400
      ScaleWidth      =   6720
      TabIndex        =   24
      Top             =   4080
      Width           =   6785
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
         TabIndex        =   25
         Top             =   3240
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3480
         TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   28
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5895
      Left            =   6840
      TabIndex        =   23
      Top             =   1920
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10398
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   " ÊﬁÌ  «·„Â„…"
            Key             =   "dat"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ÃÂ«  «·« ’«·"
            Key             =   "contact"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "„·›«  «·’Ê "
            Key             =   "snd"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ŒÌ«—«  «·„Â„…"
            Object.Tag             =   "options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Õ«·… «·„Â„…"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "„Ê«›ﬁ"
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
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Cancel          =   -1  'True
      Caption         =   "«·€«¡ «·«„—"
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
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   " ÿ»Ìﬁ"
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
      Index           =   2
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   7680
      Width           =   1215
   End
   Begin VB.PictureBox pic 
      Height          =   5415
      Index           =   2
      Left            =   1320
      RightToLeft     =   -1  'True
      ScaleHeight     =   5355
      ScaleWidth      =   6555
      TabIndex        =   11
      Top             =   1800
      Width           =   6615
      Begin VB.CommandButton cmdAdd 
         Caption         =   "«÷«›….."
         Height          =   375
         Index           =   0
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Õ–›"
         Height          =   375
         Index           =   1
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "«“«·… Ã„Ì⁄ «·ÃÂ«  «·„Õœœ… „‰ «·ﬁ«∆„…"
         Top             =   4920
         Width           =   735
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   360
         TabIndex        =   14
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
               Picture         =   "frmEditSch.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmEditSch.frx":3052
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmEditSch.frx":33A4
               Key             =   ""
            EndProperty
         EndProperty
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
         TabIndex        =   16
         Top             =   120
         Width           =   3240
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditSch.frx":38F6
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
         TabIndex        =   15
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.PictureBox pic 
      Height          =   5415
      Index           =   3
      Left            =   120
      RightToLeft     =   -1  'True
      ScaleHeight     =   5355
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   1920
      Width           =   6615
      Begin PhoneBazookaWaveCntl.objAudio objAudio 
         Left            =   120
         Top             =   3840
         _ExtentX        =   1085
         _ExtentY        =   979
      End
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
         TabIndex        =   46
         ToolTipText     =   "„⁄«Ì‰… «·„·› «·’Ê Ì"
         Top             =   1560
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   5400
         TabIndex        =   33
         Top             =   4920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "Text1"
         BuddyDispid     =   196618
         OrigLeft        =   5760
         OrigTop         =   4920
         OrigRight       =   6015
         OrigBottom      =   5295
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Text            =   "1"
         Top             =   4920
         Width           =   270
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   840
         TabIndex        =   5
         ToolTipText     =   "«·„·›«  «·’Ê ÌÏ «· Ì ”Ì „ ‰ﬁ·Â« «·Ï ÃÂ«  «·« ’«·"
         Top             =   1560
         Width           =   5430
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         TabIndex        =   4
         ToolTipText     =   "«—›«ﬁ „·› ’Ê Ì"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "«“«·… «·„·› «·„Õœœ"
         Top             =   2280
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
         TabIndex        =   2
         Top             =   2640
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
         TabIndex        =   1
         Top             =   3000
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6360
         Top             =   7680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Õœœ «·„·› «·’Ê Ì"
         Filter          =   "*.wav"
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
         TabIndex        =   10
         Top             =   4200
         Width           =   5655
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
         TabIndex        =   9
         Top             =   3840
         Width           =   1875
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
         TabIndex        =   8
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditSch.frx":39A6
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
         TabIndex        =   7
         Top             =   600
         Width           =   5655
      End
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
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEditSch.frx":3A3B
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
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   480
      Width           =   5385
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   240
      Picture         =   "frmEditSch.frx":3AF2
      Top             =   240
      Width           =   960
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   11640
      X2              =   13200
      Y1              =   2760
      Y2              =   2400
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
   Begin VB.Label lblT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   8
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmEditSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sch_ID As Integer
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
            frmContacs.IsEdit = True
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
    cmdOrder(2).Enabled = True
End Sub

Private Sub cmdOrder_Click(Index As Integer)

    Select Case Index
        Case 2
            If List1.ListCount < 1 Then MsgBox "ÌÃ» «‰  ÷Ì› „·› ’Ê Ì ⁄·Ï «·«ﬁ·", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            If ListView1.ListItems.Count < 1 Then MsgBox "ÌÃ» «‰  ÷Ì› ÃÂ… « ’«· ⁄·Ï «·«ﬁ·", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub

            cmdOrder(2).Enabled = False
            SetSchduler
        Case 1
            Unload Me
        Case 0
            If List1.ListCount < 1 Then MsgBox "ÌÃ» «‰  ÷Ì› „·› ’Ê Ì ⁄·Ï «·«ﬁ·", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
            If ListView1.ListItems.Count < 1 Then MsgBox "ÌÃ» «‰  ÷Ì› ÃÂ… « ’«· ⁄·Ï «·«ﬁ·", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub

            SetSchduler
            Unload Me
    End Select
End Sub

Public Sub LaodSchInfo(SchID As Integer)
    On Error GoTo er:
    Dim rs As New Recordset
    Dim x() As String
    Sch_ID = SchID
    rs.Open "select * from schmaster where schid=" & SchID, db
    
    If rs!schdate = "1" Then Check1.Value = 1 Else Check1.Value = 0: DTPicker1.Value = rs!schdate
    DTPicker2.Value = rs!schtime
    x = Split(rs!schsnd, ",")
    List1.Clear
    For Each y In x
        If y <> "" Then List1.AddItem y
        
    Next
    Text1 = rs!recall
    rs.Close
    rs.Open "select * from schslave where schid=" & SchID, db
    While Not rs.EOF
        Dim xx As ListItem
        Set xx = ListView1.ListItems.Add(, , , 2, 2)
        xx.Text = rs!cName
        xx.SubItems(1) = rs!cno
        rs.MoveNext
    Wend
    Exit Sub
er:
    'MsgBox Err.Description
    Resume Next
End Sub

Sub SetSchduler()
    On Error GoTo er:
    Dim rs As New Recordset
    Dim rsX As New Recordset
    Dim rsUp As New Recordset
    recall = Text1
    schtime = Format(DTPicker2.Value, "HH:MM:SS")
    schdate = Format(DTPicker1.Value, "dd-mm-yyyy")

    rsX.Open "select schid from schmaster where schdate='" & schdate & "' and schtime='" & schtime & "'", db, adOpenStatic
    'If rsX.RecordCount > 0 Then MsgBox "ÌÊÃœ „Â„… „ÃœÊ·… »‰›” «· «—ÌŒ Ê«· ÊﬁÌ  «·„Õœœ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading: Exit Sub
    For i = 0 To List1.ListCount - 1
        schsnd = schsnd & List1.List(i) & ","
    Next i
    schsnd = Replace(schsnd, "'", "''")
    If Check1.Value Then
        schdate = "1"
    End If
    rsUp.Open "update schmaster set schtime='" & schtime & "',schdate  ='" & schdate & "',schsnd='" & schsnd & "',recall='" & recall & "' where schid=" & Sch_ID, db
    rs.Open "delete * from schslave where schid=" & Sch_ID, db
    For i = 1 To ListView1.ListItems.Count
        rs.Open "insert into schslave(schid,cname,cno) Values ('" & Sch_ID & "','" & ListView1.ListItems(i).Text & "','" & ListView1.ListItems(i).SubItems(1) & "')", db
    Next i
    
    Exit Sub
er:
    MsgBox Err.Description
    Resume
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim rs As New Recordset
    
    Select Case Index
        Case 0
            If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ Õ–› «·Â„… «·Õ«·Ì… ø" & vbNewLine & "      ·« Ì„ﬂ‰ «· —«Ã⁄ ⁄‰ ⁄„·Ì… «·Õ–› ", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbNo Then Exit Sub
            rs.Open "delete * from schmaster where schid=" & Sch_ID, db
            frmShowSch.ListView1.ListItems.Remove frmShowSch.ListView1.SelectedItem.Index
            Unload Me
            MsgBox " „ Õ–› «·„Â„… »‰Ã«Õ", vbInformation
        Case 1
            rs.Open "select active from schmaster where schid=" & Sch_ID, db
            If rs!active = 1 Then
                rs.Close
                If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ  ⁄ÿÌ· «·Â„… «·Õ«·Ì… ø", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbNo Then Exit Sub
                rs.Open "update schmaster set active='" & "0" & "' where schid=" & Sch_ID, db
                Command1(1).Caption = " „ﬂÌ‰ «·„Â„… «·Õ«·Ì…"
            Else
                rs.Close
                If MsgBox("Â· «‰  „ √ﬂœ „‰ «‰ﬂ  —Ìœ  „ﬂÌ‰ «·Â„… «·Õ«·Ì… ø", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo) = vbNo Then Exit Sub
                rs.Open "update schmaster set active='" & "1" & "' where schid=" & Sch_ID, db
                Command1(1).Caption = " ⁄ÿÌ· «·„Â„… «·Õ«·Ì…"
            End If
        Case 2
            Dim x As String
            rs.Open "select schname from schmaster where schid=" & Sch_ID, db
            x = rs!schname
            x = Trim(InputBox("«ﬂ » «·«”„ «·ÃœÌœ ··„Â„… «·Õ«·Ì… Ê„‰ À„ «‰ﬁ— ⁄·Ï „Ê«›ﬁ", "«⁄«œ…  ”„Ì… „Â„…", x))
            rs.Close
            If x = "" Then Exit Sub
            rs.Open "update schmaster set schname='" & x & "' where schid=" & Sch_ID, db
            Me.Caption = " Õ—Ì— «·„Â„… " & x
            frmShowSch.ListView1.SelectedItem.SubItems(4) = x
    End Select
End Sub

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
            
        Case 4 ' preview paly file
            If List1.Text <> "" Then
                sndPlaySound List1.Text, 1
                ' lets open the wave file, this also sets the audio format for us
'                If objAudio.OpenWaveFile(List1.Text) Then
'
'                    ' start the speaker output
'                    objAudio.StartSpeakerOutput 0
'
'                    ' play the wave file to the speakers
'                    objAudio.PlayFile List1.Text, objAudio.hSpeakerWaveOut
'                End If
            Else
                MsgBox "«·—Ã«¡ «Œ — «·„·› «·’Ê Ì „‰ «·ﬁ«∆„…", vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading
            End If


    End Select
    cmdOrder(2).Enabled = True
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
Dim rs As New Recordset
    pic(1).ZOrder
    OptemizeTitle
    DTPicker1.Value = Date
    rs.Open "select active from schmaster where schid=" & Sch_ID, db
    If rs!active = 1 Then
        Command1(1).Caption = " ⁄ÿÌ· «·„Â„… «·Õ«·Ì…"
    Else
        Command1(1).Caption = " „ﬂÌ‰ «·„Â„… «·Õ«·Ì…"
    End If
    rs.Close
    rs.Open "select schname from schmaster  where schid=" & Sch_ID, db
    x = rs!schname
    Me.Caption = " Õ—Ì— «·„Â„… " & x
End Sub
Sub OptemizeTitle()
    Icon = Nothing
    lblT(8).Top = 0
    lblT(8).Left = 0
    lblT(8).Width = Me.Width
    lblT(8).ZOrder 1
    For i = 1 To pic.Count
        pic(i).Top = 1880
        pic(i).Left = 145
        pic(i).Height = 5535
        pic(i).Width = 6785
        pic(i).BorderStyle = 0
    Next i
    Me.TabStrip1.Top = 1560
    TabStrip1.Left = 120
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


Private Sub lblInfo_Change(Index As Integer)
    cmdOrder(2).Enabled = True
End Sub


Private Sub TabStrip1_Click()
    pic(TabStrip1.SelectedItem.Index).ZOrder
End Sub

Private Sub Text1_Change()
On Error GoTo er:
     UpDown1.Value = Text1
     cmdOrder(2).Enabled = True
    Exit Sub
er:
    UpDown1.Value = 0

End Sub

Private Sub UpDown1_Change()
    On Error GoTo er:
    Text1 = UpDown1.Value
    Exit Sub
er:
    Text1 = 0
End Sub
