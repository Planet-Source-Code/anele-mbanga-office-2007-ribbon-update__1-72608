VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "Get Date"
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Set Date"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Timer AnimateLogo 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   4560
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Animate Logo"
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Update Fonts"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set Windows"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Read Windows"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Read Names"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rename Saver"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdRenameTab 
      Caption         =   "Rename Tab 2"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Progress"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Timer Downloading 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3840
      Top             =   4560
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   7815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Ribbon"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin TheRibbon.ACPRibbon ACPRibbon1 
      Height          =   2130
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   3757
      ImageSize       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UsePermissions  =   0   'False
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   230
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":059A
            Key             =   "commentw"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A99
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1033
            Key             =   "find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15CD
            Key             =   "opened"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B67
            Key             =   "report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2101
            Key             =   "npo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":269B
            Key             =   "empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C35
            Key             =   "full"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31CF
            Key             =   "restore"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3521
            Key             =   "isazi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3ABB
            Key             =   "inbox"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4055
            Key             =   "experts"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45EF
            Key             =   "runsql2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A41
            Key             =   "survey"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4FDB
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5135
            Key             =   "xx"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5D07
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6159
            Key             =   "excel"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BD7B
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BED5
            Key             =   "table"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C46F
            Key             =   "ie"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CA09
            Key             =   "sum"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CB6B
            Key             =   "key1"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CFBD
            Key             =   "module"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D117
            Key             =   "stats"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D431
            Key             =   "new"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D7DB
            Key             =   "print"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DBD9
            Key             =   "taskt"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13B73
            Key             =   "attacht"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13FC5
            Key             =   "verify"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":142DF
            Key             =   "defer"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":146F6
            Key             =   "discuss"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14B0C
            Key             =   "maybe"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14F23
            Key             =   "move"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1533E
            Key             =   "risk"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15751
            Key             =   "yes"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15B67
            Key             =   "high"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15F37
            Key             =   "normal"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16325
            Key             =   "low"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16714
            Key             =   "furious"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16B38
            Key             =   "happy"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16F6A
            Key             =   "neutral"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1739B
            Key             =   "upsat"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":177C4
            Key             =   "sad"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17BF0
            Key             =   "task25"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17FC6
            Key             =   "task50"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1837A
            Key             =   "task75"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1870E
            Key             =   "task100"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18B0B
            Key             =   "task0"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18EFB
            Key             =   "email"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19495
            Key             =   "hight"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1986D
            Key             =   "lowt"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19C64
            Key             =   "normalt"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A05A
            Key             =   "furioust"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A486
            Key             =   "happyt"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A8C0
            Key             =   "neutralt"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1ACF9
            Key             =   "upsatt"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B12A
            Key             =   "sadt"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B55B
            Key             =   "defert"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B97A
            Key             =   "discusst"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BD98
            Key             =   "maybet"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C1B7
            Key             =   "movet"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C5DA
            Key             =   "riskt"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C9F5
            Key             =   "yest"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CE13
            Key             =   "task25t"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D1F1
            Key             =   "task50t"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D5AD
            Key             =   "task75t"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D949
            Key             =   "task100t"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DD4E
            Key             =   "task0t"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E146
            Key             =   "green"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E598
            Key             =   "red"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E9EA
            Key             =   "organization"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EE3C
            Key             =   "region"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F28E
            Key             =   "department"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":222B0
            Key             =   "owner"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23102
            Key             =   "resources"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2939C
            Key             =   "target1"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29936
            Key             =   "date"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29D59
            Key             =   "perspective"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A1AB
            Key             =   "duedate"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A745
            Key             =   "complete"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2ACDF
            Key             =   "expected"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30901
            Key             =   "taborder"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30E9B
            Key             =   "link"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31075
            Key             =   "column"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31B3F
            Key             =   "runsql"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32068
            Key             =   "taskx"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":323F8
            Key             =   "attach"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":327D6
            Key             =   "info"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32C28
            Key             =   "develop"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":41A73
            Key             =   "mindmanager"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4377D
            Key             =   "suite"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45007
            Key             =   "star"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":46289
            Key             =   "sync"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4C4DF
            Key             =   "offline"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D269
            Key             =   "highr"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4DFF3
            Key             =   "lowr"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4ED7D
            Key             =   "mediumr"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4FB07
            Key             =   "wss"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":51E89
            Key             =   "wssdoc"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":529D3
            Key             =   "toolicon"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55185
            Key             =   "useraccount"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5551E
            Key             =   "calender"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55838
            Key             =   "chart"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55A9F
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55E06
            Key             =   "list"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5606B
            Key             =   "newsomething"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":56378
            Key             =   "iconopen"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":566B4
            Key             =   "profile"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":569E2
            Key             =   "project"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":56D44
            Key             =   "resources1"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":57079
            Key             =   "reports"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":575D6
            Key             =   "info1"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":57D28
            Key             =   "warn"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":58042
            Key             =   "traffic"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":58494
            Key             =   "target"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":588E6
            Key             =   "doclibrary"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":59430
            Key             =   "live1"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":63D72
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":641C4
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":69DE6
            Key             =   "exportproject"
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":73A34
            Key             =   "importmpp"
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":817D5
            Key             =   "x"
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":818E7
            Key             =   "calendar2"
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":87881
            Key             =   "decrement"
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":87CD3
            Key             =   "increment"
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":88125
            Key             =   "collaborate"
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8DD47
            Key             =   "review2"
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":95249
            Key             =   "progress"
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":95AEC
            Key             =   "yellowr"
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":95ED4
            Key             =   "greenr"
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":962C6
            Key             =   "projectplan"
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":96D18
            Key             =   "redr"
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":971C7
            Key             =   "people"
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":98D9E
            Key             =   "bundle"
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":99169
            Key             =   "running"
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9E95B
            Key             =   "stopped"
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F36D
            Key             =   "right"
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F907
            Key             =   "left"
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9FEA1
            Key             =   "deletex"
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A096B
            Key             =   "editx"
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A1435
            Key             =   "check"
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A174F
            Key             =   "group1"
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A2321
            Key             =   "none"
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A2738
            Key             =   "bluer"
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A3397
            Key             =   "purpler"
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A40F8
            Key             =   "task"
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A4DB4
            Key             =   "note"
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A5194
            Key             =   "money"
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A55FC
            Key             =   "warn1"
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A62E3
            Key             =   "question"
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A6C0A
            Key             =   "change2"
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A7120
            Key             =   "excel2"
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A86A8
            Key             =   "chart1"
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A9E12
            Key             =   "pdf1"
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AA5B3
            Key             =   "robot1"
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AAAB0
            Key             =   "wssw1"
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AB470
            Key             =   "resource"
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ABCC5
            Key             =   "day"
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AC27E
            Key             =   "wssw"
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ACF53
            Key             =   "group"
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ADCE6
            Key             =   "robot"
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AEF05
            Key             =   "calendar1"
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AFCD5
            Key             =   "actionw"
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B0143
            Key             =   "action1"
         EndProperty
         BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B05B4
            Key             =   "action"
         EndProperty
         BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B0A0E
            Key             =   "powerpoint"
         EndProperty
         BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B1D1B
            Key             =   "pie"
         EndProperty
         BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B625F
            Key             =   "shake"
         EndProperty
         BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B6A0F
            Key             =   "newx"
         EndProperty
         BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B7480
            Key             =   "refreshmeeting"
         EndProperty
         BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B8162
            Key             =   "discuss1"
         EndProperty
         BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B9269
            Key             =   "write"
         EndProperty
         BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B982E
            Key             =   "action2"
         EndProperty
         BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B9C80
            Key             =   "company"
         EndProperty
         BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BA1BC
            Key             =   "redfolder"
         EndProperty
         BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BB0CC
            Key             =   "greenfolder"
         EndProperty
         BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BBF67
            Key             =   "construct"
         EndProperty
         BeginProperty ListImage180 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BC695
            Key             =   "camera"
         EndProperty
         BeginProperty ListImage181 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BD0FF
            Key             =   "expand"
         EndProperty
         BeginProperty ListImage182 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BDF55
            Key             =   "live"
         EndProperty
         BeginProperty ListImage183 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BE461
            Key             =   "change"
         EndProperty
         BeginProperty ListImage184 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BE9AF
            Key             =   "documents"
         EndProperty
         BeginProperty ListImage185 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BEDB8
            Key             =   "docs"
         EndProperty
         BeginProperty ListImage186 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BF1C9
            Key             =   "docsfolder"
         EndProperty
         BeginProperty ListImage187 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BF5C6
            Key             =   "sitevisit"
         EndProperty
         BeginProperty ListImage188 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BFD02
            Key             =   "photo"
         EndProperty
         BeginProperty ListImage189 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C0777
            Key             =   "tracking"
         EndProperty
         BeginProperty ListImage190 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C0C4C
            Key             =   "report1"
         EndProperty
         BeginProperty ListImage191 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C1132
            Key             =   "recommendationw"
         EndProperty
         BeginProperty ListImage192 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C1647
            Key             =   "recommendationt"
         EndProperty
         BeginProperty ListImage193 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C1B65
            Key             =   "commentt"
         EndProperty
         BeginProperty ListImage194 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C2073
            Key             =   "wizard1"
         EndProperty
         BeginProperty ListImage195 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C3E09
            Key             =   "milestone"
         EndProperty
         BeginProperty ListImage196 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C41C2
            Key             =   "view"
         EndProperty
         BeginProperty ListImage197 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C45AF
            Key             =   "wizard"
         EndProperty
         BeginProperty ListImage198 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C4C64
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage199 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C51CB
            Key             =   "checkmark"
         EndProperty
         BeginProperty ListImage200 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C5773
            Key             =   "xmark"
         EndProperty
         BeginProperty ListImage201 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C5DEE
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage202 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C676A
            Key             =   "project.show"
         EndProperty
         BeginProperty ListImage203 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CCFCC
            Key             =   "executivet"
         EndProperty
         BeginProperty ListImage204 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CE3DF
            Key             =   "executivew"
         EndProperty
         BeginProperty ListImage205 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CF84D
            Key             =   "key"
         EndProperty
         BeginProperty ListImage206 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CFC61
            Key             =   "keyt"
         EndProperty
         BeginProperty ListImage207 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D0074
            Key             =   "reviewt"
         EndProperty
         BeginProperty ListImage208 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D060E
            Key             =   "revieww"
         EndProperty
         BeginProperty ListImage209 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D0989
            Key             =   "increaseform"
         EndProperty
         BeginProperty ListImage210 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D1004
            Key             =   "notenlarge"
         EndProperty
         BeginProperty ListImage211 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D144B
            Key             =   "ts"
         EndProperty
         BeginProperty ListImage212 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D2C21
            Key             =   "blue"
         EndProperty
         BeginProperty ListImage213 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D3978
            Key             =   "brown"
         EndProperty
         BeginProperty ListImage214 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D485A
            Key             =   "bluet"
         EndProperty
         BeginProperty ListImage215 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D55BF
            Key             =   "ambert"
         EndProperty
         BeginProperty ListImage216 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D59D1
            Key             =   "greent"
         EndProperty
         BeginProperty ListImage217 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D5DED
            Key             =   "redt"
         EndProperty
         BeginProperty ListImage218 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D61FB
            Key             =   "offlinef"
         EndProperty
         BeginProperty ListImage219 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D6608
            Key             =   "onlinef"
         EndProperty
         BeginProperty ListImage220 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D6A2F
            Key             =   "closedw"
         EndProperty
         BeginProperty ListImage221 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D6DE4
            Key             =   "openedw"
         EndProperty
         BeginProperty ListImage222 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D7197
            Key             =   "synchronize"
         EndProperty
         BeginProperty ListImage223 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D85CE
            Key             =   "moon1"
         EndProperty
         BeginProperty ListImage224 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D8A20
            Key             =   "moon2"
         EndProperty
         BeginProperty ListImage225 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D8E72
            Key             =   "moon3"
         EndProperty
         BeginProperty ListImage226 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D92C4
            Key             =   "moon4"
         EndProperty
         BeginProperty ListImage227 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D9716
            Key             =   "moon5"
         EndProperty
         BeginProperty ListImage228 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D9B68
            Key             =   "moon6"
         EndProperty
         BeginProperty ListImage229 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D9FBA
            Key             =   "moon7"
         EndProperty
         BeginProperty ListImage230 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DA40C
            Key             =   "moon8"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngCounter As Long
Private animation As Integer
Private Sub ACPRibbon1_ButtonClick(ByVal Id As String, ByVal Caption As String)
    On Error Resume Next
    Text1.Text = Id & ", " & Caption
    Select Case Id
    Case "mnuExit"
        MsgBox "Are you sure that you want to exit this application?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Exit"
    Case "3"
    Case "godisgood"
        MsgBox "God is good all the time"
    End Select
    Err.Clear
End Sub
Private Sub ACPRibbon1_ComboClick(ByVal ComboName As String, ByVal Text As String)
    On Error Resume Next
    Text1.Text = ComboName & ", " & Text
    Err.Clear
End Sub
Private Sub ACPRibbon1_DatePickClick(ByVal DatePickName As String, ByVal DatePicked As String)
    On Error Resume Next
    Text1.Text = DatePickName & ", " & DatePicked
    Err.Clear
End Sub

Private Sub ACPRibbon1_MainMenuClick(ByVal Id As String)
    Text1.Text = Id
End Sub

Private Sub ACPRibbon1_MenuClick(ByVal Id As String, ByVal Caption As String)
    On Error Resume Next
    Text1.Text = Id
    Select Case Id
    Case "offline"
        ACPRibbon1.EditTopButton "offline", "online", "Work online", "onlinef", "Work online"
        MsgBox "You are now working online."
    Case "online"
        ACPRibbon1.EditTopButton "online", "offline", "Work offline", "offlinef", "Work offline"
        MsgBox "You are now working offline."
    End Select
    Err.Clear
End Sub
Private Sub AnimateLogo_Timer()
    On Error Resume Next
    animation = animation + 1
    If animation > 8 Then animation = 1
    ACPRibbon1.Icon = "moon" & CStr(animation)
    Err.Clear
End Sub
Private Sub cmdRenameTab_Click()
    On Error Resume Next
    ACPRibbon1.RenameTab "tab2", "Anele Mbanga"
    Err.Clear
End Sub
Public Sub Command1_Click()
    On Error Resume Next
    animation = 0
    Command8.Caption = "Animate Logo"
    AnimateLogo.Enabled = False
    With ACPRibbon1
    .UsePermissions = False
    .Clear
    .ImageList = imgIcons
    .ImageSize = Size320
    .Top = 0
    .Left = 0
    .Icon = "ie"
    .ResizeLogo 480
    .AddTopButton "print", "Printer", "print", "Print"
    .AddTopButtonMenu "print", "mnuPrintAll", "Print All"
    .AddTopButtonMenu "print", "mnuPrintThis", "Print This"
    .AddTopButton "offline", "Offline", "offlinef", "Work Offline"
    .AddTab "tab1", "Example", True
    .AddCat "cat1", "tab1", "Cat 1", True, ""
    .AddButton "but1", "cat1", "Search", "find", True, "", False
    .AddComboBox "but2", "cat1", "Names", "", "cboNames", 2000
    .AddComboBoxItem "cboNames", "Anele Mbanga"
    .AddComboBoxItem "cboNames", "Sikelela Mbanga"
    .AddComboBoxItem "cboNames", "Usibabale Mbanga"
    .AddComboBoxItem "cboNames", "Olothando Mbanga"
    .AddComboBox "but3", "cat1", "Fonts", "", "cboFonts", 3000
    .AddComboBoxItem "cboFonts", "Arial"
    .AddComboBoxItem "cboFonts", "Tahoma"
    .AddComboBoxItem "cboFonts", "Gothica"
    .AddTextBox "but4", "cat1", "Windows", "Testing text boxes", "txtWindows", 1500
    .AddDatePicker "but5", "cat1", "My Date", "Select date of birth", "DOB", 1355
    .AddProgressBar "but6", "cat1", "Progress", "Show my progress bar", "progNote", 2000, 0, 500
    .AddLabel "but7", "cat1", "So Far", "0", False, "How far we are at", False
    .AddTab "tab2", "Tab 2", True
    .AddCat "cat2", "tab2", "Group 1", False, ""
    .AddButton "but8", "cat2", "Search", "save", False, "", False
    .AddButton "saver", "cat1", "Saver", "Save", True, "Saving data", False
    ''''''''''''''''''
    .AddButtonMenu "but1", "mnuFile", "File", True
    .AddButtonMenu "but1", "mnuFile\mnuOpen", "Open", True
    .AddButtonMenu "but1", "mnuFile\mnuSave", "Save"
    .AddButtonMenu "but1", "mnuFile\mnuDelete", "Delete"
    ''''
    .AddButtonMenu "but1", "mnuSearch", "Search Database", True
    .AddButtonMenu "but1", "mnuSearch\mnuNames", "Names"
    .AddButtonMenu "but1", "mnuSearch\mnuDates", "Dates"
    .AddButtonMenu "but1", "mnuCompress", "Compress Database", True
    .AddButtonMenu "but1", "mnuSearch\mnuTables", "Tables"
    .AddButtonMenu "but1", "-", "-"
    .AddButtonMenu "but1", "mnuExit", "Exit"
    .AddButtonMenu "but1", "mnuCompress\mnuMakeMDE", "Make MDE"
    .AddButtonMenu "but1", "mnuCompress\mnuRepair", "Repair"
    
    ' add circle menu
    .AddCircleMenu "mnuNew", "New"
    .AddCircleMenu "mnuOpen", "Open"
    .AddCircleMenu "mnuDelete", "Delete"
    .AddCircleMenu "-", "-"
    .AddCircleMenu "mnuExit", "Exit"
    .Refresh
    End With
    Err.Clear
End Sub
Private Sub Command10_Click()
    On Error Resume Next
    MsgBox ACPRibbon1.DatePickerGetDate("dob")
    Err.Clear
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    ACPRibbon1.ProgressBarReset "progNote", 100
    lngCounter = 0
    Downloading.Enabled = True
    Err.Clear
End Sub
Private Sub Command3_Click()
    On Error Resume Next
    ACPRibbon1.EditButton "saver", "God is Good", "synchronize", False, "God is good all the time.", False, "", "godisgood"
    Err.Clear
End Sub
Private Sub Command4_Click()
    On Error Resume Next
    MsgBox ACPRibbon1.ComboBoxGetText("cboNames")
    Err.Clear
End Sub
Private Sub Command5_Click()
    On Error Resume Next
    MsgBox ACPRibbon1.TextBoxGetText("txtWindows")
    Err.Clear
End Sub
Private Sub Command6_Click()
    On Error Resume Next
    ACPRibbon1.TextBoxSetText "txtWindows", "My Window"
    Err.Clear
End Sub
Private Sub Command7_Click()
    On Error Resume Next
    Dim fontTot As Long
    Dim fontCnt As Long
    Dim myF As Variant
    ACPRibbon1.ComboBoxClear "cboFonts"
    For fontCnt = 0 To Screen.FontCount - 1
        ACPRibbon1.AddComboBoxItem "cboFonts", Screen.Fonts(fontCnt)
    Next
    ACPRibbon1.ComboBoxRefresh
    Err.Clear
End Sub
Private Sub Command8_Click()
    On Error Resume Next
    Select Case Command8.Caption
    Case "Animate Logo"
        AnimateLogo.Enabled = True
        Command8.Caption = "Stop Animation"
    Case Else
        animation = 0
        Command8.Caption = "Animate Logo"
        AnimateLogo.Enabled = False
        ACPRibbon1.Icon = "ie"
    End Select
    Err.Clear
End Sub
Private Sub Command9_Click()
    On Error Resume Next
    ACPRibbon1.DatePickerSetDate "DOB", "01/01/2009"
    Err.Clear
End Sub
Private Sub Downloading_Timer()
    On Error Resume Next
    Downloading.Enabled = False
    lngCounter = lngCounter + 1
    ACPRibbon1.ProgressBarUpdate "progNote", lngCounter
    'ACPRibbon1.EditLabel "but7", "So far", CStr(lngCounter)
    ACPRibbon1.LabelUpdate "but7", CStr(lngCounter)
    Text1.Text = lngCounter
    Downloading.Enabled = True
    If lngCounter > 100 Then
        lngCounter = 0
        ACPRibbon1.LabelUpdate "but7", "0"
        Downloading.Enabled = False
    End If
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Command1_Click
    Text1.Text = ""
    Set ACPRibbon1.ParentForm = Me
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Set ACPRibbon1.ParentForm = Me
    Err.Clear
End Sub
