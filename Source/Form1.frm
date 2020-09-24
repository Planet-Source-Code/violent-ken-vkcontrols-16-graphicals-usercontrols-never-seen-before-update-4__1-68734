VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "*\AOCX\vkUserControlsXP.vbp"
Begin VB.Form Form1 
   Caption         =   "Tests"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   150
   ClientWidth     =   13260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   13260
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkBar vkBar21 
      Height          =   255
      Left            =   3150
      TabIndex        =   76
      Top             =   2325
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      Decimals        =   0
      Max             =   250
      Value           =   1
      BackPicture     =   "Form1.frx":0000
      FrontPicture    =   "Form1.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkLabel vkLabel2 
      Height          =   1815
      Left            =   3150
      TabIndex        =   77
      Top             =   690
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3201
      BorderStyle     =   1
      BackColor       =   16777215
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin vkUserContolsXP.vkFrame vkFrame9 
      Height          =   3975
      Left            =   1440
      TabIndex        =   78
      Top             =   3000
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7011
      Caption         =   "Unicode is supported (try to move me !)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleHeight     =   320
      Picture         =   "Form1.frx":0038
      PictureAlignment=   1
      UseUnicode      =   -1  'True
      Begin vkUserContolsXP.vkOptionButton vkOptionButton2 
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   4
         UseUnicode      =   -1  'True
      End
      Begin vkUserContolsXP.vkCheck vkCheck6 
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseUnicode      =   -1  'True
      End
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseUnicode      =   -1  'True
      End
      Begin vkUserContolsXP.vkCommand vkCommand230 
         Height          =   495
         Left            =   240
         TabIndex        =   82
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseUnicode      =   -1  'True
      End
      Begin vkUserContolsXP.vkToggleButton vkToggleButton2 
         Height          =   495
         Left            =   240
         TabIndex        =   81
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
         UseUnicode      =   -1  'True
      End
      Begin vkUserContolsXP.vkListBox vkListBox4 
         Height          =   2055
         Left            =   2280
         TabIndex        =   80
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   0
         UseUnicode      =   -1  'True
      End
      Begin vkUserContolsXP.vkTextBox vkTextBox4 
         Height          =   1215
         Left            =   120
         TabIndex        =   79
         Top             =   2640
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2143
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendText      =   "Example"
         LegendForeColor =   12937777
         LegendType      =   1
         UseUnicode      =   -1  'True
      End
   End
   Begin vkUserContolsXP.vkTimer vkTimer2 
      Left            =   4703
      Top             =   5692
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   20
      Enabled         =   -1  'True
   End
   Begin vkUserContolsXP.vkListBox vkListBox3 
      Height          =   1695
      Left            =   6120
      TabIndex        =   73
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0
   End
   Begin vkUserContolsXP.vkFrame vkFrame8 
      Height          =   1695
      Left            =   7680
      TabIndex        =   69
      Top             =   7680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2990
      Caption         =   "Example 2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkTimer vkTimer1 
         Left            =   600
         Top             =   120
         _ExtentX        =   926
         _ExtentY        =   926
         Interval        =   20
      End
      Begin vkUserContolsXP.vkCommand vkCommand22 
         Height          =   375
         Left            =   120
         TabIndex        =   72
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Stop"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand21 
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Start"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkBar vkBar3 
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Value           =   1
         DisplayLabel    =   2
         BackPicture     =   "Form1.frx":1B02
         FrontPicture    =   "Form1.frx":1B1E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame7 
      Height          =   1695
      Left            =   2760
      TabIndex        =   65
      Top             =   7680
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2990
      Caption         =   "Example"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkListBox vkListBox2 
         Height          =   855
         Left            =   120
         TabIndex        =   68
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1508
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiSelect     =   0   'False
         Sorted          =   0
      End
      Begin vkUserContolsXP.vkCommand vkCommand20 
         Height          =   315
         Left            =   1680
         TabIndex        =   67
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "Add to list"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkTextBox vkTextBox3 
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Text            =   "Here is a line"
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendForeColor =   12937777
      End
   End
   Begin vkUserContolsXP.vkScrollContainer vkScrollContainer1 
      Height          =   3345
      Left            =   9120
      TabIndex        =   45
      Top             =   6000
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   5900
      AreaHeight      =   3000
      AreaWidth       =   10000
      Begin vkUserContolsXP.vkCommand vkCommand19 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   75
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
      End
      Begin vkUserContolsXP.vkCommand vkCommand19 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   74
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   14345190
         BackColorPushed1=   14542053
         BackColorPushed2=   14345442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   7617536
         DisabledBackColor=   15398133
         CustomStyle     =   2
      End
      Begin vkUserContolsXP.vkVScroll vkVScroll4 
         Height          =   2295
         Left            =   8520
         TabIndex        =   64
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   4048
         ArrowColor      =   4210752
         BackColor       =   14737632
         BorderColor     =   12632256
         FrontColor      =   8421504
         LargeChangeColor=   4210752
      End
      Begin vkUserContolsXP.vkVScroll vkVScroll3 
         Height          =   2295
         Left            =   8160
         TabIndex        =   63
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   4048
         ArrowColor      =   128
         BackColor       =   12632319
         BorderColor     =   255
         FrontColor      =   8421631
         MouseHoverColor =   4210816
         DownColor       =   255
         LargeChangeColor=   16384
      End
      Begin vkUserContolsXP.vkFrame vkFrame6 
         Height          =   1095
         Index           =   1
         Left            =   6360
         TabIndex        =   62
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1931
         BackColor1      =   16777215
         BackGradient    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         ShowTitle       =   0   'False
         TitleGradient   =   0
         TitleHeight     =   300
         RoundAngle      =   20
         BorderWidth     =   2
      End
      Begin vkUserContolsXP.vkFrame vkFrame6 
         Height          =   1095
         Index           =   0
         Left            =   6360
         TabIndex        =   61
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1931
         BackColor1      =   14215660
         BackColor2      =   14215660
         BackGradient    =   0
         Caption         =   "   Caption"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         TextPosition    =   0
         TitleColor1     =   13160660
         TitleColor2     =   14215660
         TitleGradient   =   0
         BorderColor     =   10070188
         BreakCorner     =   0   'False
      End
      Begin vkUserContolsXP.vkBar vkBar2 
         Height          =   375
         Index           =   4
         Left            =   3840
         TabIndex        =   60
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Decimals        =   10
         Max             =   117
         Value           =   37
         InteractiveControl=   -1  'True
         BackPicture     =   "Form1.frx":1B3A
         FrontPicture    =   "Form1.frx":1B56
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkBar vkBar2 
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   59
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BorderColor     =   8421504
         LeftColor       =   0
         RightColor      =   16777215
         DisplayLabel    =   3
         GradientMode    =   1
         ForeColor       =   49152
         BackPicture     =   "Form1.frx":1B72
         FrontPicture    =   "Form1.frx":1B8E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkBar vkBar2 
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   58
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         DisplayLabel    =   2
         BackPicture     =   "Form1.frx":1BAA
         FrontPicture    =   "Form1.frx":1BC6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkBar vkBar2 
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   57
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Value           =   75
         InteractiveControl=   -1  'True
         DisplayLabel    =   0
         InteractiveButton=   0
         BackPicture     =   "Form1.frx":1BE2
         FrontPicture    =   "Form1.frx":3371
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkBar vkBar2 
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   56
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BorderColor     =   12632256
         BackColorTop    =   16777215
         BackColorBottom =   16777215
         LeftColor       =   49152
         RightColor      =   12648384
         DisplayLabel    =   0
         BackPicture     =   "Form1.frx":4704
         FrontPicture    =   "Form1.frx":4720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkOptionButton vkOptionButton1 
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   55
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   2
         Group           =   3
      End
      Begin vkUserContolsXP.vkOptionButton vkOptionButton1 
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   54
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Group           =   45
      End
      Begin vkUserContolsXP.vkOptionButton vkOptionButton1 
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   53
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   3
      End
      Begin vkUserContolsXP.vkCheck vkCheck5 
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   52
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   2
      End
      Begin vkUserContolsXP.vkCheck vkCheck5 
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   51
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin vkUserContolsXP.vkCheck vkCheck5 
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   50
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand19 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor1      =   14215660
         BackColorPushed1=   13228765
         BackGradient    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   10070188
         BreakCorner     =   0   'False
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   14215660
         CustomStyle     =   6
      End
      Begin vkUserContolsXP.vkCommand vkCommand19 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor1      =   16761024
         BackColor2      =   13228765
         BackColorPushed1=   16761024
         BackColorPushed2=   16777215
         BackGradient    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   8421504
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         CustomStyle     =   3
      End
      Begin vkUserContolsXP.vkCommand vkCommand19 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor1      =   15597311
         BackColor2      =   13820902
         BackColorPushed1=   12636628
         BackColorPushed2=   13557730
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   8426644
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15332854
         CustomStyle     =   4
      End
      Begin vkUserContolsXP.vkCommand vkCommand19 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame4 
      Height          =   2535
      Left            =   4200
      TabIndex        =   19
      Top             =   5040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4471
      Caption         =   "Some vkListBox methods"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkCommand vkCommand11 
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         Top             =   1920
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Sort"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand10 
         Height          =   375
         Left            =   3120
         TabIndex        =   31
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Invert Checks"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand8 
         Height          =   375
         Left            =   2640
         TabIndex        =   30
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Change Scroll appareance"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand7 
         Height          =   375
         Left            =   2640
         TabIndex        =   28
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Invert selection"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand4 
         Height          =   375
         Left            =   3600
         TabIndex        =   27
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Clear"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand6 
         Height          =   375
         Left            =   2040
         TabIndex        =   26
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Random item"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCheck vkCheck4 
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Scroll"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin vkUserContolsXP.vkCheck vkCheck3 
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Border"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin vkUserContolsXP.vkCheck vkCheck2 
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "FullRowSelect"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin vkUserContolsXP.vkCheck vkCheck1 
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Use vkChecks"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand9 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Add 100 items with icons"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand vkCommand5 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Add 1000 items"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin vkUserContolsXP.vkSysTray vkSysTray1 
      Left            =   12480
      Top             =   4200
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   615
      Left            =   9120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      BorderStyle     =   1
      BackColor       =   16777215
      Caption         =   " So you still will use vbControls ? ;)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkOptionButton vkOption1 
      Height          =   255
      Left            =   9480
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Group           =   1
   End
   Begin vkUserContolsXP.vkCheck vkCheckBox1 
      Height          =   255
      Left            =   10920
      TabIndex        =   2
      Top             =   5640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkBar vkBar1 
      Height          =   375
      Left            =   9360
      TabIndex        =   3
      Top             =   5160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      DisplayLabel    =   3
      BackPicture     =   "Form1.frx":473C
      FrontPicture    =   "Form1.frx":4758
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkCommand vkCommand2 
      Height          =   495
      Left            =   11040
      TabIndex        =   4
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Grayed picture"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Picture         =   "Form1.frx":4774
   End
   Begin vkUserContolsXP.vkVScroll vkVScroll1 
      Height          =   3255
      Left            =   12720
      TabIndex        =   5
      Top             =   840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   5741
   End
   Begin vkUserContolsXP.vkHScroll vkHScroll1 
      Height          =   255
      Left            =   9120
      TabIndex        =   6
      Top             =   4200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   3255
      Left            =   9120
      TabIndex        =   7
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5741
      Caption         =   "Icon and rounded border"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleHeight     =   300
      Picture         =   "Form1.frx":4D0E
      RoundAngle      =   9
      Begin vkUserContolsXP.vkToggleButton vkToggleButton1 
         Height          =   495
         Left            =   120
         TabIndex        =   87
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "ToggleButton"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":52A8
         Value           =   -1  'True
         DisplayMouseHoverIcon=   -1  'True
         MouseHoverPicture=   "Form1.frx":5842
      End
      Begin vkUserContolsXP.vkFrame vkFrame5 
         Height          =   1725
         Left            =   1800
         TabIndex        =   36
         Top             =   1395
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3043
         Caption         =   "TraySystem"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin vkUserContolsXP.vkCommand vkCommand17 
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Reboot explorer"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkCommand vkCommand16 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   880
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Caption         =   "Remove 2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkCommand vkCommand15 
            Height          =   255
            Left            =   840
            TabIndex        =   39
            Top             =   300
            Width           =   680
            _ExtentX        =   1191
            _ExtentY        =   450
            Caption         =   "Add 2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkCommand vkCommand14 
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Caption         =   "Remove 1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkCommand vkCommand13 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   300
            Width           =   680
            _ExtentX        =   1191
            _ExtentY        =   450
            Caption         =   "Add 1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame3 
         Height          =   975
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1720
         Caption         =   "UpDown control"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin vkUserContolsXP.vkUpDown vkUpDown4 
            Height          =   255
            Left            =   960
            TabIndex        =   17
            Top             =   600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Direction       =   0
         End
         Begin vkUserContolsXP.vkUpDown vkUpDown3 
            Height          =   255
            Left            =   960
            TabIndex        =   16
            Top             =   320
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin vkUserContolsXP.vkUpDown vkUpDown2 
            Height          =   375
            Left            =   480
            TabIndex        =   15
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Direction       =   0
         End
         Begin vkUserContolsXP.vkUpDown vkUpDown1 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame2 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3625
         Caption         =   "Enabled=False"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Begin vkUserContolsXP.vkVScroll vkVScroll2 
            Height          =   1575
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   2778
            Enabled         =   0   'False
         End
         Begin vkUserContolsXP.vkCommand vkCommand3 
            Height          =   495
            Left            =   600
            TabIndex        =   11
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin vkUserContolsXP.vkOptionButton vkOption2 
            Height          =   255
            Left            =   600
            TabIndex        =   10
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Group           =   2
         End
         Begin vkUserContolsXP.vkOptionButton vkOption3 
            Height          =   255
            Left            =   600
            TabIndex        =   9
            Top             =   1440
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Value           =   1
            Group           =   2
         End
      End
   End
   Begin vkUserContolsXP.vkCommand vkCommand18 
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   5280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      Caption         =   "Changer vkTextBox's Scroll appareance"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkTextBox vkTextBox2 
      Height          =   2415
      Left            =   120
      TabIndex        =   43
      Top             =   2760
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4260
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   2
      LegendAutosize  =   -1  'True
      LegendForeColor =   12937777
      LegendType      =   2
   End
   Begin vkUserContolsXP.vkTextBox vkTextBox1 
      Height          =   2535
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   2
      LegendText      =   "Sample of vertical legend"
      LegendAlignmentY=   2
      LegendBackColor1=   16761024
      LegendBackColor2=   16711680
      LegendGradient  =   2
      LegendForeColor =   16777215
      LegendType      =   1
   End
   Begin vkUserContolsXP.vkListBox lstFile 
      Height          =   1815
      Left            =   1680
      TabIndex        =   33
      Top             =   5760
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0
      StyleCheckBox   =   -1  'True
      ListType        =   1
      Path            =   "C:\"
   End
   Begin vkUserContolsXP.vkListBox vkListBox1 
      Height          =   1335
      Left            =   120
      TabIndex        =   35
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0
      ListType        =   3
      IconSize        =   32
      ShowReadOnlyFiles=   0   'False
   End
   Begin vkUserContolsXP.vkCommand vkCommand12 
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   5760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Change Path"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkListBox List2 
      Height          =   1695
      Left            =   120
      TabIndex        =   29
      Top             =   7680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   1
      ListType        =   2
      Path            =   "C:\"
      ShowReadOnlyFiles=   0   'False
   End
   Begin vkUserContolsXP.vkListBox List 
      Height          =   4875
      Left            =   4200
      TabIndex        =   18
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8599
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0
      ShowReadOnlyFiles=   0   'False
   End
   Begin vkUserContolsXP.vkCommand vkCommand1 
      Height          =   495
      Left            =   9240
      TabIndex        =   86
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Move mouse hover me !"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":5DDC
      PictureOffsetX  =   390
      DisplayMouseHoverIcon=   -1  'True
      MouseHoverPicture=   "Form1.frx":6376
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   12480
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   31
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6910
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6FB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7306
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7658
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":79AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":804E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":83A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":86F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":90E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":943A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":978C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":9ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":9E30
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":A182
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":A4D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":A826
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":AB78
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":AECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":B21C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":B56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":B8C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":BC12
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":BF64
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":C2B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":C608
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":C95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":CCAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopUp1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "&Menu1"
      End
      Begin VB.Menu tiret 
         Caption         =   "-"
      End
      Begin VB.Menu mnu11 
         Caption         =   "&Menu11"
      End
   End
   Begin VB.Menu mnuPopUp2 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnu2 
         Caption         =   "&Menu2"
      End
      Begin VB.Menu tiret23 
         Caption         =   "-"
      End
      Begin VB.Menu mnu22 
         Caption         =   "&Menu22"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Sub Form_Load()
    
    vkLabel2.Caption = vbNewLine & vbNewLine & "Try to compile vkUserControlsXP to have better performances ;)" & vbNewLine & "Please vote for me if you enjoy my controls ! Thanks ^^    "
        
    Randomize

    List.DisplayVScroll = True
    
    List2.Path = App.Path
    'List2.ListType = FolderListBox

    vkListBox1.Path = App.Path
    lstFile.Path = App.Path
    'vkListBox1.ListType = DriveListBox
    'lstFile.ListType = FileListBox
    
    vkTextBox2.Text = "Here is a sample of a vkTextBox wich draw the line number" & vbNewLine & "It can be helpful to display some source code." & vbNewLine & "Text example : 'This is a text sample wich will be written on several lines ;) Of course you have to use 'LineNumberLegend' style with VerticalScrollbar and Multiline modes.'"
    vkTextBox2.Text = vkTextBox2.Text & vbNewLine & "Usercontrol self-sublcassing to update line numbers when you scroll, enter new texte, add a new line ...etc."
    
    vkTextBox1.Text = "Here is a sample of multiline vkTextBox. Click me to display the CurrentLine on the Form Caption !" & vbNewLine & "Here is an other line." & vbNewLine & "Try to use the new method and property of vkTextBox !" & vbNewLine & "Careful : HScroll of vkTextBox is not finished : there is still some bugs... but sur you will enjoy my controls ;)"
    
    Dim x As Long
    vkTextBox2.UnRefreshControl = True
    For x = 1 To 100
        Call vkTextBox2.AppendLine("New line ! " & CStr(x))
    Next x
    vkTextBox2.UnRefreshControl = False
    Call vkTextBox2.Refreshnum(True)
    
    DoEvents
    Call vkCommand23_Click
    
    vkFrame9.Caption = vkFrame9.Caption & "  " & ChrW$(20013) & ChrW$(25991) & " (" & ChrW$(21488) & ChrW$(28771) & ")"
    vkOptionButton2.Caption = LoadResString(101)
    vkCheck6.Caption = LoadResString(102)
    vkLabel3.Caption = LoadResString(103)
    vkCommand230.Caption = LoadResString(111)
    vkToggleButton2.Caption = LoadResString(105)
    vkTextBox4.LegendText = LoadResString(110)
    With vkListBox4
        .UnRefreshControl = True
        For x = 106 To 117
            Call .AddItem(LoadResString(x))
        Next x
        .UnRefreshControl = False
        Call .Refresh
    End With
End Sub

Private Sub vkFrame9_MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    Call ReleaseCapture
    Call SendMessage(vkFrame9.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub vkFrame9_MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    If Button = vbLeftButton Then _
        vkFrame9.Visible = Not ((x > vkFrame9.Width - 300) And (y < 300))
End Sub

Private Sub List_ItemChek(Item As vkUserContolsXP.vkListItem)
    Me.Caption = "Checked " & Item.Text
End Sub

Private Sub List_ItemClick(Item As vkUserContolsXP.vkListItem)
    'Me.Caption = "Selected " & Item.Text
End Sub

Private Sub List_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    If Button = vbMiddleButton Then Me.Caption = List.ListCount
    If Button = vbRightButton Then
        List.Item(List.ListIndex).BackColor = vbRed
    End If
End Sub

Private Sub List2_ItemDblClick(Item As vkUserContolsXP.vkListItem)
    lstFile.Path = Item.tagString1
End Sub
Private Sub vkCheck1_Change(Value As CheckBoxConstants)
    List.StyleCheckBox = CBool(Value)
End Sub

Private Sub vkCheck2_Change(Value As CheckBoxConstants)
    List.FullRowSelect = CBool(Value)
End Sub

Private Sub vkCheck3_Change(Value As CheckBoxConstants)
    List.DisplayBorder = CBool(Value)
End Sub

Private Sub vkCheck4_Change(Value As CheckBoxConstants)
    List.DisplayVScroll = CBool(Value)
End Sub

Private Sub vkCommand1_Click()
    Me.Caption = "Pushed"
End Sub

Private Sub vkCommand1_MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    Me.Caption = "Up"
End Sub

Private Sub vkCommand10_Click()
    Call List.InvertChecks
End Sub

Private Sub vkCommand11_Click()
Dim l As Long
    Call List.SortItems(Alphabetical)
End Sub

Private Sub vkCommand12_Click()
    List2.Path = BrowseForFolder("Choose a folder", Me.hwnd)
    lstFile.Path = List2.Path
End Sub

Private Sub vkCommand13_Click()
    With vkSysTray1
        .BalloonTipString = "Here's the first Item !"
        Set .Icon = vkCommand1.Picture
        Call .AddToTray(1)
    End With
    'vkFrame5_MouseDown
End Sub

Private Sub vkCommand14_Click()
    Call vkSysTray1.RemoveFromTray(1)
    'vkFrame5_MouseDown
End Sub

Private Sub vkCommand15_Click()
    With vkSysTray1
        .BalloonTipString = "Here is the SECOND Item !"
        Set .Icon = vkFrame1.Picture
        Call .AddToTray(2)
    End With
    'vkFrame5_MouseDown
End Sub

Private Sub vkCommand16_Click()
    Call vkSysTray1.RemoveFromTray(2)
    'vkFrame5_MouseDown
End Sub

Private Sub vkCommand17_Click()
    Call KickProcess("explorer.exe")
    Call Shell("explorer.exe")
End Sub

Private Sub vkCommand18_Click()
Dim tVS As New vkPrivateScroll
    
    With tVS
        .ArrowColor = vbYellow
        .DownColor = vbRed
        .Width = 255
        .Enabled = vkTextBox2.VScroll.Enabled
        .FrontColor = .FrontColor - 100
        .MouseInterval = 1
    End With
    
    vkTextBox2.VScroll = tVS
    Call vkTextBox2.Refreshnum(True)
    Set tVS = Nothing
End Sub

Private Sub vkCommand18_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    If Button = vbMiddleButton Then Me.vkTextBox2.Refreshnum (True)
End Sub

Private Sub vkCommand2_Click()
    Beep
End Sub

Private Sub vkCommand20_Click()
    Me.vkListBox2.AddItem Caption:=Me.vkTextBox3.Text
End Sub

Private Sub vkCommand21_Click()
    vkTimer1.Enabled = True
End Sub

Private Sub vkCommand22_Click()
    vkTimer1.Enabled = False
End Sub

Private Sub vkCommand23_Click()
'liste toutes les polices installées
Dim x As Long
Dim t As StdFont
Dim It As vkListItem

    For x = 0 To Screen.FontCount - 1
        Set t = New StdFont
        t.Name = Screen.Fonts(x)
        Set It = New vkListItem
        With It
            .Font = t
            .Text = t.Name
        End With
        vkListBox3.AddItem Item:=It
    Next x
    
    Call vkListBox3.SortItems(Alphabetical)
    
    Set t = Nothing
    Set It = Nothing
End Sub

Private Sub vkCommand4_Click()
    Call List.Clear
End Sub

Private Sub vkCommand5_Click()
Dim x As Long
Dim l As Long
Dim l1 As Long
Dim b As Boolean
    
    b = List.AcceptAutoSort
    
    l1 = GetTickCount
    List.UnRefreshControl = True
    List.AcceptAutoSort = False
    For x = 1 To 1000
        List.AddItem Rnd
    Next x
    List.UnRefreshControl = False: List.Refresh
    l1 = GetTickCount - l1
    List.AcceptAutoSort = b
    Me.Caption = l & "    " & l1
End Sub

Private Sub vkCommand6_Click()
Dim It As New vkListItem
Dim tFont As New StdFont

    With tFont
        .Bold = Int(Rnd * 2) - 1
        .Italic = Int(Rnd * 2) - 1
        .Name = IIf(Rnd > 0.66, "Courier New", IIf(Rnd < 0.33, "Tahoma", "Times New Roman"))
        .Size = 9 + Int(Rnd * 6)
        .Underline = Int(Rnd * 2) - 1
    End With
    
    With It
        .Alignment = Int(Rnd * 3)
        .Checked = Int(Rnd * 2) - 1
        .Selected = Int(Rnd * 2) - 1
        .Font = tFont
        .ForeColor = IIf(Rnd > 0.4, IIf(Rnd > 0.8, vbRed, vbBlack), vbBlue)
        .Text = "Random Item !"
        .Height = 290 + Int(Rnd * 200)
        .Key = "key1"
        .Icon = IMG.ListImages.Item(Int(Rnd * 30) + 1).Picture
        .pxlIconHeight = 16
        .pxlIconWidth = 16
        '.SelColor = RGB(255, 180, 158)
        .BackColor = IIf(Rnd > 0.7, RGB(221, 240, 255), vbWhite)
        .BorderSelColor = IIf(Rnd > 0.7, vbGreen, vbBlue)
    End With
    
    Call List.AddItem(Item:=It)
End Sub

Private Sub vkCommand7_Click()
    List.InvertSelection
End Sub

Private Sub vkCommand8_Click()
Dim tVS As New vkPrivateScroll
    
    With tVS
        .ArrowColor = vbYellow
        .DownColor = vbRed
        .Width = 400
        .Value = Int(List.VScroll.Max / 2)
        .FrontColor = .FrontColor - 100
        .MouseInterval = 1
    End With
    
    List.VScroll = tVS
    Set tVS = Nothing
End Sub

Private Sub vkCommand9_Click()
Dim x As Long
Dim l As Long
Dim l1 As Long
Dim t As vkListItem
Dim b As Boolean
    
    b = List.AcceptAutoSort
    List.AcceptAutoSort = False
    l1 = GetTickCount
    List.UnRefreshControl = True
    List.Visible = False
    For x = 1 To 100
        Set t = New vkListItem
        With t
            .pxlIconHeight = 16
            .pxlIconWidth = 16
            .Icon = IMG.ListImages.Item(Int(Rnd * 30) + 1).Picture
            .Height = 420
            .Font = Me.Font
            .BackColor = vbWhite
            .Text = CStr(Rnd)
        End With
        List.AddItem , t
    Next x
    List.Visible = True
    l1 = GetTickCount - l1
    Set t = Nothing
    Me.Caption = l & "    " & l1
    List.AcceptAutoSort = b
    List.UnRefreshControl = False: List.Refresh
End Sub

Private Sub vkFrame5_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    Me.Caption = "First one " & CStr(vkSysTray1.IsInTray(1)) & " Second one : " & CStr(vkSysTray1.IsInTray(2))
End Sub

Private Sub vkHScroll1_Change(Value As Currency)
    Me.Caption = Rnd
End Sub

Private Sub vkListBox1_ItemDblClick(Item As vkUserContolsXP.vkListItem)
    lstFile.Path = Item.tagString1
    List2.Path = Item.tagString1
End Sub

Private Sub vkMouseKeyEvents1_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.Caption = Shift
End Sub

Private Sub vkScrollContainer1_HScrollChange(Value As Currency)
    '/!\ THIS LINE IS HERE ONLY BECAUSE VKCONTROLS ARE NOT REFRESH
    'IN THE WM_PAINT MESSAGE
    'YOU CAN REMOVE THIS LINE IF YOU USE OTHER CONTROLS IN VKSCROLLCONTAINER
    '
    '/!\ CETTE LIGNE DE CODE N'EST LA QUE PARCE QUE LES VKCONTROLS NE SONT
    'PAS RAFRAICHIT QUANQ ILS RECOIVENT LE MESSAGE WM_PAINT
    'VOUS POUVEZ (DEVEZ) SUPPRIMER CETTE LIGNE SI VOUS UTILISEZ DES CONTROLES
    'DE BASE DANS LE VKSCROLLCONTAINER
    Me.vkScrollContainer1.Refresh
End Sub

Private Sub vkScrollContainer1_HScrollScroll()
    '/!\ THIS LINE IS HERE ONLY BECAUSE VKCONTROLS ARE NOT REFRESH
    'IN THE WM_PAINT MESSAGE
    'YOU CAN REMOVE THIS LINE IF YOU USE OTHER CONTROLS IN VKSCROLLCONTAINER
    '
    '/!\ CETTE LIGNE DE CODE N'EST LA QUE PARCE QUE LES VKCONTROLS NE SONT
    'PAS RAFRAICHIT QUANQ ILS RECOIVENT LE MESSAGE WM_PAINT
    'VOUS POUVEZ (DEVEZ) SUPPRIMER CETTE LIGNE SI VOUS UTILISEZ DES CONTROLES
    'DE BASE DANS LE VKSCROLLCONTAINER
    Me.vkScrollContainer1.Refresh
End Sub

Private Sub vkScrollContainer1_VScrollChange(Value As Currency)
    '/!\ THIS LINE IS HERE ONLY BECAUSE VKCONTROLS ARE NOT REFRESH
    'IN THE WM_PAINT MESSAGE
    'YOU CAN REMOVE THIS LINE IF YOU USE OTHER CONTROLS IN VKSCROLLCONTAINER
    '
    '/!\ CETTE LIGNE DE CODE N'EST LA QUE PARCE QUE LES VKCONTROLS NE SONT
    'PAS RAFRAICHIT QUANQ ILS RECOIVENT LE MESSAGE WM_PAINT
    'VOUS POUVEZ (DEVEZ) SUPPRIMER CETTE LIGNE SI VOUS UTILISEZ DES CONTROLES
    'DE BASE DANS LE VKSCROLLCONTAINER
    Me.vkScrollContainer1.Refresh
End Sub

Private Sub vkScrollContainer1_VScrollScroll()
    '/!\ THIS LINE IS HERE ONLY BECAUSE VKCONTROLS ARE NOT REFRESH
    'IN THE WM_PAINT MESSAGE
    'YOU CAN REMOVE THIS LINE IF YOU USE OTHER CONTROLS IN VKSCROLLCONTAINER
    '
    '/!\ CETTE LIGNE DE CODE N'EST LA QUE PARCE QUE LES VKCONTROLS NE SONT
    'PAS RAFRAICHIT QUANQ ILS RECOIVENT LE MESSAGE WM_PAINT
    'VOUS POUVEZ (DEVEZ) SUPPRIMER CETTE LIGNE SI VOUS UTILISEZ DES CONTROLES
    'DE BASE DANS LE VKSCROLLCONTAINER
    Me.vkScrollContainer1.Refresh
End Sub

Private Sub vkBar21_ValueIsMax(Value As Double)
    vkTimer2.Enabled = False
    vkLabel2.Visible = False
    Me.vkBar21.Visible = False
End Sub

Private Sub vkSysTray1_MouseUp(Button As MouseButtonConstants, ID As Long)
    If Button = vbRightButton Then
        If ID = 1 Then
            Me.PopupMenu Me.mnuPopUp1
        ElseIf ID = 2 Then
            Me.PopupMenu Me.mnuPopUp2
        End If
    End If
End Sub

Private Sub vkTimer1_Timer()
    With vkBar3
        .Value = .Value + 1
        If .Value = .Max Then .Value = .Min
    End With
End Sub

Private Sub vkTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Caption = vkTextBox1.GetLine(vkTextBox1.LineIndex)
End Sub

Private Sub vkTimer2_Timer()
    vkBar21.Value = vkBar21.Value + 1
End Sub

Private Sub vkVScroll1_Change(Value As Currency)
    Me.Caption = Rnd
End Sub
