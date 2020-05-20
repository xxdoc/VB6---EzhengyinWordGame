VERSION 5.00
Begin VB.Form FormSettings 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   0  'None
   Caption         =   "EzhengyinWordGame"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13560
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormSettings.frx":0000
   LinkTopic       =   "FormSettings"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormSettings.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   9465
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameFonts 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "字体更换  (Beta)"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   6510
      TabIndex        =   57
      Top             =   6720
      Width           =   6840
      Begin VB.CommandButton CmdFontsApply 
         Caption         =   "Apply"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5040
         MouseIcon       =   "FormSettings.frx":015E
         MousePointer    =   99  'Custom
         TabIndex        =   63
         Top             =   1890
         Width           =   1485
      End
      Begin VB.TextBox TextboxFontsEngFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1995
         MousePointer    =   3  'I-Beam
         TabIndex        =   62
         Text            =   "SimHei"
         Top             =   1420
         Width           =   4530
      End
      Begin VB.TextBox TextboxFontsJpnFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1995
         MousePointer    =   3  'I-Beam
         TabIndex        =   60
         Text            =   "SimHei"
         Top             =   980
         Width           =   4530
      End
      Begin VB.CheckBox CheckboxFontsSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "启用字体更换（可能导致程序崩溃，请小心）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":02B0
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   420
         Width           =   4530
      End
      Begin VB.Label LabelFontsEngFont 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "选项字体："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   630
         TabIndex        =   61
         Top             =   1470
         Width           =   1230
      End
      Begin VB.Label LabelFontsJpnFont 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "题面字体："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   630
         TabIndex        =   59
         Top             =   1050
         Width           =   1230
      End
   End
   Begin VB.Frame FrameKanaIncluded 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "包括的内容  (占 250 分)"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   210
      TabIndex        =   18
      Top             =   2625
      Width           =   6000
      Begin VB.CheckBox CheckboxKanaIncluded03 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "其它"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3990
         MouseIcon       =   "FormSettings.frx":0402
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   420
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.CheckBox CheckboxKanaIncluded02 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "辱包"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2100
         MouseIcon       =   "FormSettings.frx":0554
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   420
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.CheckBox CheckboxKanaIncluded01 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "膜蛤"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":06A6
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   420
         Value           =   1  'Checked
         Width           =   960
      End
   End
   Begin VB.Frame FrameGameDifficultyIndexIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "难度指数"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1380
      Left            =   210
      TabIndex        =   2
      Top             =   1050
      Width           =   6000
      Begin VB.CommandButton CmdGameDifficultyIndexIndicatorHelp 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5460
         MouseIcon       =   "FormSettings.frx":07F8
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   315
         Width           =   420
      End
      Begin VB.Timer TimerProgressbarAnimation 
         Interval        =   1
         Left            =   5670
         Top             =   1050
      End
      Begin VB.Shape ShapeGameDifficultyIndexIndicatorProgressbar 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         Height          =   120
         Left            =   315
         Top             =   1050
         Width           =   120
      End
      Begin VB.Shape ShapeGameDifficultyIndexIndicatorBottombar 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         Height          =   120
         Left            =   315
         Top             =   1050
         Width           =   5370
      End
      Begin VB.Label LabelGameDifficultyIndexIndicator1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   690
         Left            =   210
         TabIndex        =   3
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label LabelGameDifficultyIndexIndicator2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "／1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1575
         TabIndex        =   4
         Top             =   735
         Width           =   915
      End
      Begin VB.Label LabelGameDifficultyIndexIndicator3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "当前难度指数的描述..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2625
         TabIndex        =   5
         Top             =   630
         Width           =   2970
      End
   End
   Begin VB.Frame FrameInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "键盘输入"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1380
      Left            =   6510
      TabIndex        =   7
      Top             =   1050
      Width           =   6840
      Begin VB.TextBox TextboxInputOption3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   6090
         MaxLength       =   1
         MousePointer    =   3  'I-Beam
         TabIndex        =   17
         Top             =   735
         Width           =   435
      End
      Begin VB.TextBox TextboxInputOption2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   3885
         MaxLength       =   1
         MousePointer    =   3  'I-Beam
         TabIndex        =   14
         Top             =   735
         Width           =   435
      End
      Begin VB.TextBox TextboxInputOption1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   1680
         MaxLength       =   1
         MousePointer    =   3  'I-Beam
         TabIndex        =   11
         Top             =   735
         Width           =   435
      End
      Begin VB.Label LabelInputOption3Indicator 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   5565
         TabIndex        =   16
         Top             =   735
         Width           =   495
      End
      Begin VB.Label LabelInputOption3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "选项3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4725
         TabIndex        =   15
         Top             =   840
         Width           =   810
      End
      Begin VB.Label LabelInputOption2Indicator 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   3360
         TabIndex        =   13
         Top             =   735
         Width           =   495
      End
      Begin VB.Label LabelInputOption2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "选项2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   810
      End
      Begin VB.Label LabelInputOption1Indicator 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1155
         TabIndex        =   10
         Top             =   735
         Width           =   495
      End
      Begin VB.Label LabelInputOption1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "选项1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   315
         TabIndex        =   9
         Top             =   840
         Width           =   810
      End
      Begin VB.Label LabelInput 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "更改选项对应的按键。提示：您始终可以使用 F6, F7 与 F8."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   450
         Width           =   5700
      End
   End
   Begin VB.Frame FrameDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "显示"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1800
      Left            =   210
      TabIndex        =   48
      Top             =   6300
      Width           =   6000
      Begin VB.CheckBox CheckboxDisplaySpinningSakura 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "旋转的维尼遨游星瀚"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormSettings.frx":094A
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2220
      End
      Begin VB.CheckBox CheckboxDisplaySmoothAnimations 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "平滑动画效果"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":0A9C
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   1260
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.CheckBox CheckboxDisplayHideUnnecessaryInformation 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "隐藏不重要的信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":0BEE
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   840
         Width           =   2220
      End
      Begin VB.CheckBox CheckboxDisplayReduceContrast 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "降低对比度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormSettings.frx":0D40
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   420
         Width           =   1800
      End
      Begin VB.CheckBox CheckboxDisplayBlackOnWhite 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "黑底白字"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":0E92
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   420
         Width           =   1800
      End
   End
   Begin VB.Frame FrameDifficulty 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "难度  (占 500 分)"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3900
      Left            =   6510
      TabIndex        =   31
      Top             =   2625
      Width           =   6840
      Begin VB.CheckBox CheckboxDifficultyIncreaseDifficultyGradually 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "缓慢提升难度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":0FE4
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   840
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.HScrollBar HScrollDifficultyMistakeAllowedAmount 
         Height          =   330
         LargeChange     =   5
         Left            =   3150
         Max             =   10
         MouseIcon       =   "FormSettings.frx":1136
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   3330
         Value           =   3
         Width           =   3375
      End
      Begin VB.HScrollBar HScrollDifficultyInterval 
         Height          =   330
         LargeChange     =   15
         Left            =   3150
         Max             =   30
         Min             =   1
         MouseIcon       =   "FormSettings.frx":1288
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   2800
         Value           =   10
         Width           =   3375
      End
      Begin VB.HScrollBar HScrollDifficultyReachNormalDifficultyAt 
         Height          =   330
         LargeChange     =   50
         Left            =   3150
         Max             =   100
         MouseIcon       =   "FormSettings.frx":13DA
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   2160
         Value           =   20
         Width           =   3375
      End
      Begin VB.HScrollBar HScrollDifficultyInitialDifficulty 
         Height          =   330
         LargeChange     =   20
         Left            =   3150
         Max             =   50
         Min             =   2
         MouseIcon       =   "FormSettings.frx":152C
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   1320
         Value           =   50
         Width           =   3375
      End
      Begin VB.HScrollBar HScrollDifficultyNormalDifficulty 
         Height          =   330
         LargeChange     =   20
         Left            =   3150
         Max             =   50
         Min             =   2
         MouseIcon       =   "FormSettings.frx":167E
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   490
         Value           =   20
         Width           =   3375
      End
      Begin VB.Label LabelDifficultyMistakeAllowedAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "允许失误次数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   210
         TabIndex        =   45
         Top             =   3360
         Width           =   1440
      End
      Begin VB.Label LabelDifficultyInterval 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "间隔时长："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   210
         TabIndex        =   42
         Top             =   2850
         Width           =   1440
      End
      Begin VB.Label LabelDifficultyReachNormalDifficultyAt 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "完成到指定游戏进度时，抵达正常难度："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   525
         TabIndex        =   39
         Top             =   1785
         Width           =   3750
      End
      Begin VB.Label LabelDifficultyInitialDifficulty 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "初始难度："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   525
         TabIndex        =   36
         Top             =   1365
         Width           =   1440
      End
      Begin VB.Label LabelDifficultyNormalDifficulty 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "正常难度："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   210
         TabIndex        =   32
         Top             =   525
         Width           =   1650
      End
      Begin VB.Label LabelDifficultyNormalDifficultyIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1575
         TabIndex        =   33
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label LabelDifficultyInitialDifficultyIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1575
         TabIndex        =   37
         Top             =   1305
         Width           =   1440
      End
      Begin VB.Label LabelDifficultyReachNormalDifficultyAtIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1575
         TabIndex        =   40
         Top             =   2145
         Width           =   1440
      End
      Begin VB.Label LabelDifficultyIntervalIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1575
         TabIndex        =   43
         Top             =   2790
         Width           =   1440
      End
      Begin VB.Label LabelDifficultyMistakeAllowedAmountIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1575
         TabIndex        =   46
         Top             =   3315
         Width           =   1440
      End
   End
   Begin VB.Frame FrameCheating 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "作弊"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   210
      TabIndex        =   54
      Top             =   8295
      Width           =   6000
      Begin VB.CheckBox CheckboxCheatingShowCorrectAnswer 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "指示正确选项"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormSettings.frx":17D0
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   420
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.CheckBox CheckboxCheatingSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "启用作弊功能"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":1922
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   420
         Width           =   1800
      End
   End
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "吼了"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   11865
      MouseIcon       =   "FormSettings.frx":1A74
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
   Begin VB.Frame FrameGameMode 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "游戏模式  (占 250 分)"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2325
      Left            =   210
      TabIndex        =   22
      Top             =   3780
      Width           =   6000
      Begin VB.OptionButton RadioboxGameModeKana 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "遍历所有文字（抽选过且答对过所有文字后胜利）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":1BC6
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   420
         Value           =   -1  'True
         Width           =   5580
      End
      Begin VB.OptionButton RadioboxGameModeTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "固定一段时间（坚持完这段时间即胜利）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":1D18
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1260
         Width           =   5580
      End
      Begin VB.HScrollBar HScrollGameModeSpecifiedTime 
         Height          =   330
         LargeChange     =   15
         Left            =   3570
         Max             =   30
         Min             =   1
         MouseIcon       =   "FormSettings.frx":1E6A
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1750
         Value           =   5
         Width           =   2115
      End
      Begin VB.HScrollBar HScrollGameModeRepeatedTimes 
         Height          =   330
         LargeChange     =   5
         Left            =   3570
         Max             =   10
         Min             =   1
         MouseIcon       =   "FormSettings.frx":1FBC
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   870
         Value           =   1
         Width           =   2115
      End
      Begin VB.Label LabelGameModeSpecifiedTime 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "设定时间："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   525
         TabIndex        =   28
         Top             =   1785
         Width           =   1290
      End
      Begin VB.Label LabelGameModeRepeatedTimes 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "单一文字的出现次数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   525
         TabIndex        =   25
         Top             =   945
         Width           =   2070
      End
      Begin VB.Label LabelGameModeRepeatedTimesIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2730
         TabIndex        =   26
         Top             =   840
         Width           =   705
      End
      Begin VB.Label LabelGameModeSpecifiedTimeIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1995
         TabIndex        =   29
         Top             =   1725
         Width           =   1440
      End
   End
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   13230
      Top             =   9135
   End
   Begin VB.Label LabelSettingsTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "设定"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   315
      TabIndex        =   0
      Top             =   210
      Width           =   15555
   End
   Begin VB.Shape ShapeEdge 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   9465
      Left            =   0
      Top             =   0
      Width           =   13560
   End
End
Attribute VB_Name = "FormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Public windowanimationtargetleft As Integer
Public windowanimationtargettop As Integer
Public windowanimationtargetwidth As Integer
Public windowanimationtargetheight As Integer
Public gamedifficultyindexprogressbaranimationtarget As Integer  'Range: 0~8000

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    'Close button...
    Public Sub CmdClose_Click()
        windowanimationtargetleft = (Screen.Width / 2)
        windowanimationtargettop = (Screen.Height / 2)
        windowanimationtargetwidth = 0
        windowanimationtargetheight = 0
    End Sub

    'Settings...  [!] Other settings are automatically refreshed in FormMainWindow.TimerSettingsRefresher.
    Public Sub CmdGameDifficultyIndexIndicatorHelp_Click()
        FormDifficultyIndexHelp.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormDifficultyIndexHelp.windowanimationtargetleft = (Screen.Width / 2) - (13560 / 2)
        FormDifficultyIndexHelp.windowanimationtargettop = (Screen.Height / 2) - (9465 / 2)
        FormDifficultyIndexHelp.windowanimationtargetwidth = 13560
        FormDifficultyIndexHelp.windowanimationtargetheight = 9465
        FormDifficultyIndexHelp.Show
    End Sub
    Public Sub CmdFontsApply_Click()
        FormMainWindow.LabelKanaDashboard.Font = TextboxFontsJpnFont.Text
        FormMainWindow.CmdOption1.Font = TextboxFontsEngFont.Text
        FormMainWindow.CmdOption2.Font = TextboxFontsEngFont.Text
        FormMainWindow.CmdOption3.Font = TextboxFontsEngFont.Text
        MsgBox "字体已更换！", vbInformation + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerWindowAnimation_Timer()
        If ((Me.Left = windowanimationtargetleft) And (Me.Top = windowanimationtargettop) And (Me.Width = windowanimationtargetwidth) And (Me.Height = windowanimationtargetheight)) Then Exit Sub

        Select Case FormMainWindow.setanimationswitch
            Case True
                If Me.Left > windowanimationtargetleft Then Me.Left = Me.Left - Abs(Me.Left - windowanimationtargetleft) / 4
                If Me.Left < windowanimationtargetleft Then Me.Left = Me.Left + Abs(Me.Left - windowanimationtargetleft) / 4
                If Me.Top > windowanimationtargettop Then Me.Top = Me.Top - Abs(Me.Top - windowanimationtargettop) / 4
                If Me.Top < windowanimationtargettop Then Me.Top = Me.Top + Abs(Me.Top - windowanimationtargettop) / 4
                If Me.Width > windowanimationtargetwidth Then Me.Width = Me.Width - Abs(Me.Width - windowanimationtargetwidth) / 4
                If Me.Width < windowanimationtargetwidth Then Me.Width = Me.Width + Abs(Me.Width - windowanimationtargetwidth) / 4
                If Me.Height > windowanimationtargetheight Then Me.Height = Me.Height - Abs(Me.Height - windowanimationtargetheight) / 4
                If Me.Height < windowanimationtargetheight Then Me.Height = Me.Height + Abs(Me.Height - windowanimationtargetheight) / 4
                If Abs(Me.Left - windowanimationtargetleft) < 10 Then Me.Left = windowanimationtargetleft
                If Abs(Me.Top - windowanimationtargettop) < 10 Then Me.Top = windowanimationtargettop
                If Abs(Me.Width - windowanimationtargetwidth) < 10 Then Me.Width = windowanimationtargetwidth
                If Abs(Me.Height - windowanimationtargetheight) < 10 Then Me.Height = windowanimationtargetheight
            Case False
                Me.Move windowanimationtargetleft, windowanimationtargettop, windowanimationtargetwidth, windowanimationtargetheight
        End Select

        If windowanimationtargetheight = 0 And Me.Height < 10 Then Me.Hide
    End Sub

    Public Sub TimerProgressbarAnimation_Timer()
        If Me.Height < windowanimationtargetheight Then
            ShapeGameDifficultyIndexIndicatorProgressbar.Width = 0
            Exit Sub
        End If

        Select Case FormMainWindow.setanimationswitch
            Case True
                If ShapeGameDifficultyIndexIndicatorProgressbar.Width = gamedifficultyindexprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip1_
                If ShapeGameDifficultyIndexIndicatorProgressbar.Width > gamedifficultyindexprogressbaranimationtarget Then ShapeGameDifficultyIndexIndicatorProgressbar.Width = ShapeGameDifficultyIndexIndicatorProgressbar.Width - Abs(ShapeGameDifficultyIndexIndicatorProgressbar.Width - gamedifficultyindexprogressbaranimationtarget) / 4
                If ShapeGameDifficultyIndexIndicatorProgressbar.Width < gamedifficultyindexprogressbaranimationtarget Then ShapeGameDifficultyIndexIndicatorProgressbar.Width = ShapeGameDifficultyIndexIndicatorProgressbar.Width + Abs(ShapeGameDifficultyIndexIndicatorProgressbar.Width - gamedifficultyindexprogressbaranimationtarget) / 4
                If Abs(ShapeGameDifficultyIndexIndicatorProgressbar.Width - gamedifficultyindexprogressbaranimationtarget) < 10 Then ShapeGameDifficultyIndexIndicatorProgressbar.Width = gamedifficultyindexprogressbaranimationtarget
TimerProgressbarAnimation_Skip1_:

            Case False
                If ShapeGameDifficultyIndexIndicatorProgressbar.Width = gamedifficultyindexprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip2_
                ShapeGameDifficultyIndexIndicatorProgressbar.Width = gamedifficultyindexprogressbaranimationtarget
TimerProgressbarAnimation_Skip2_:

        End Select
    End Sub
