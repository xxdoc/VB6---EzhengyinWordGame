VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "恶政隐文字游戏　v20200520"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   16710
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
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   10095
   ScaleWidth      =   16710
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer TimerSettingsRefresher 
      Interval        =   100
      Left            =   2205
      Top             =   0
   End
   Begin VB.Timer TimerTimer 
      Interval        =   90
      Left            =   14385
      Top             =   1470
   End
   Begin VB.Timer TimerSpinningSakuraAnimation 
      Interval        =   1
      Left            =   16380
      Top             =   9765
   End
   Begin VB.Timer TimerProgressbarAnimation 
      Interval        =   1
      Left            =   15960
      Top             =   9765
   End
   Begin VB.Timer TimerCalculator 
      Interval        =   90
      Left            =   3780
      Top             =   1470
   End
   Begin VB.TextBox TextboxInput 
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
      Left            =   1365
      MaxLength       =   1
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Top             =   735
      Width           =   435
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "停止"
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
      Left            =   210
      MouseIcon       =   "FormMainWindow.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   735
      Width           =   1065
   End
   Begin VB.CommandButton CmdOption3 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   11235
      MouseIcon       =   "FormMainWindow.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   7455
      Width           =   4740
   End
   Begin VB.CommandButton CmdOption1 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   735
      MouseIcon       =   "FormMainWindow.frx":0402
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   7455
      Width           =   4740
   End
   Begin VB.CommandButton CmdOption2 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   5985
      MouseIcon       =   "FormMainWindow.frx":0554
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   7455
      Width           =   4740
   End
   Begin VB.CommandButton CmdEXIT 
      Cancel          =   -1  'True
      Caption         =   "退出"
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
      Left            =   14910
      MouseIcon       =   "FormMainWindow.frx":06A6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   210
      Width           =   1590
   End
   Begin VB.CommandButton CmdStartPauseResume 
      Caption         =   "开始"
      Default         =   -1  'True
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
      Left            =   210
      MouseIcon       =   "FormMainWindow.frx":07F8
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   210
      Width           =   1590
   End
   Begin VB.Timer TimerGameStatusRefresher 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2625
      Top             =   1260
   End
   Begin VB.Timer TimerClock 
      Interval        =   500
      Left            =   16170
      Top             =   840
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   435
      Left            =   1680
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   435
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   767
      _cy             =   767
   End
   Begin VB.Label LabelGameDifficultyIndexIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1260
      TabIndex        =   11
      Top             =   5250
      Width           =   2955
   End
   Begin VB.Label LabelGameDifficultyIndexTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "难度指数"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1260
      MouseIcon       =   "FormMainWindow.frx":094A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "请参考 [设定]→[难度指数]→[?]。"
      Top             =   4620
      Width           =   2955
   End
   Begin VB.Label LabelGameAverageReactionTimeIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   12495
      TabIndex        =   17
      Top             =   5250
      Width           =   2955
   End
   Begin VB.Label LabelGameAverageReactionTimeTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "平均反应速度"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   12495
      MouseIcon       =   "FormMainWindow.frx":0A9C
      MousePointer    =   99  'Custom
      TabIndex        =   16
      ToolTipText     =   "从新的文字与选项显示出来开始，到您作答为止，二者之间的时长（秒）。"
      Top             =   4620
      Width           =   2955
   End
   Begin VB.Label LabelGameTimeElapsedIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   12495
      TabIndex        =   15
      Top             =   3675
      Width           =   2955
   End
   Begin VB.Label LabelGameTimeElapsedTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "本局耗时"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   12495
      MouseIcon       =   "FormMainWindow.frx":0BEE
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   "本局游戏的耗时计时器（分' 秒"" 百毫秒）。此计时器可能很不准确。"
      Top             =   3045
      Width           =   2955
   End
   Begin VB.Label LabelGameProgressIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1260
      TabIndex        =   9
      Top             =   3675
      Width           =   2955
   End
   Begin VB.Label LabelGameProgressTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "游戏进度"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1260
      MouseIcon       =   "FormMainWindow.frx":0D40
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "本局游戏的完成率（％）。"
      Top             =   3045
      Width           =   2955
   End
   Begin VB.Line LineSpinningSakura5 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   5
      X1              =   735
      X2              =   735
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSpinningSakura4 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   5
      X1              =   630
      X2              =   630
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSpinningSakura3 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   5
      X1              =   525
      X2              =   525
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSpinningSakura2 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   5
      X1              =   420
      X2              =   420
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSpinningSakura1 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   5
      X1              =   315
      X2              =   315
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Shape ShapeGameCurrentTimeLeftProgressbar 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   5370
      Left            =   12075
      Top             =   1680
      Width           =   120
   End
   Begin VB.Shape ShapeGameCurrentDifficultyProgressbar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   5370
      Left            =   4515
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label LabelGameCurrentTimeLeftIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   12495
      TabIndex        =   13
      Top             =   2100
      Width           =   2955
   End
   Begin VB.Label LabelGameCurrentTimeLeftTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "剩余时间"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   12495
      MouseIcon       =   "FormMainWindow.frx":0E92
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "当前文字的剩余答题时间（秒）。"
      Top             =   1470
      Width           =   2955
   End
   Begin VB.Label LabelGameCurrentDifficultyIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1260
      TabIndex        =   7
      Top             =   2100
      Width           =   2955
   End
   Begin VB.Label LabelGameCurrentDifficultyTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "当前难度"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1260
      MouseIcon       =   "FormMainWindow.frx":0FE4
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "当前对于单一文字的答题限时（秒）。数值越小难度越高。"
      Top             =   1470
      Width           =   2955
   End
   Begin VB.Label LabelOption3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   11235
      MouseIcon       =   "FormMainWindow.frx":1136
      MousePointer    =   99  'Custom
      TabIndex        =   24
      ToolTipText     =   "键盘操作，可以按这个按键或按 F8 来选择此选项。"
      Top             =   9450
      Width           =   4740
   End
   Begin VB.Label LabelOption2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   5985
      MouseIcon       =   "FormMainWindow.frx":1288
      MousePointer    =   99  'Custom
      TabIndex        =   23
      ToolTipText     =   "键盘操作，可以按这个按键或按 F7 来选择此选项。"
      Top             =   9450
      Width           =   4740
   End
   Begin VB.Shape ShapeGameCurrentTimeLeftBottombar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   5580
      Left            =   12075
      Top             =   1470
      Width           =   120
   End
   Begin VB.Shape ShapeGameCurrentDifficultyBottombar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   5580
      Left            =   4515
      Top             =   1470
      Width           =   120
   End
   Begin VB.Label LabelStatusbar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "载入中..."
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   2205
      MouseIcon       =   "FormMainWindow.frx":13DA
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "状态栏。"
      Top             =   315
      Width           =   12300
   End
   Begin VB.Label LabelOption1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   735
      MouseIcon       =   "FormMainWindow.frx":152C
      MousePointer    =   99  'Custom
      TabIndex        =   22
      ToolTipText     =   "键盘操作，可以按这个按键或按 F6 来选择此选项。"
      Top             =   9450
      Width           =   4740
   End
   Begin VB.Label LabelClock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   14910
      MouseIcon       =   "FormMainWindow.frx":167E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "时钟。"
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label LabelKanaDashboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   270
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5550
      Left            =   5355
      TabIndex        =   18
      Top             =   1470
      Width           =   5970
   End
   Begin VB.Shape ShapeLightIndicatorOption1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   2010
      Left            =   630
      Top             =   7350
      Width           =   4950
   End
   Begin VB.Shape ShapeLightIndicatorOption2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   2010
      Left            =   5880
      Top             =   7350
      Width           =   4950
   End
   Begin VB.Shape ShapeLightIndicatorOption3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   2010
      Left            =   11130
      Top             =   7350
      Width           =   4950
   End
   Begin VB.Shape ShapeGameProgressProgressbar 
      BackColor       =   &H00FF8800&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   2205
      Top             =   945
      Width           =   12090
   End
   Begin VB.Shape ShapeGameProgressBottombar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   2205
      Top             =   945
      Width           =   12300
   End
   Begin VB.Menu MenuGame 
      Caption         =   "游戏 (&G)"
      Begin VB.Menu MenuGameStartPauseResume 
         Caption         =   "开始"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuGameStop 
         Caption         =   "停止"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MenuGame1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuGameChooseOption1 
         Caption         =   "选择 选项1"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuGameChooseOption2 
         Caption         =   "选择 选项2"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenuGameChooseOption3 
         Caption         =   "选择 选项3"
         Enabled         =   0   'False
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu Menu1_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuSoundSwitch 
      Caption         =   "声音 开 (&D)"
   End
   Begin VB.Menu MenuSettings 
      Caption         =   "设定 (&S)..."
   End
   Begin VB.Menu MenuAbout 
      Caption         =   "关于 (&A)"
      Begin VB.Menu MenuAbout1 
         Caption         =   "由于本游戏的敏感性，作者已隐匿身份。"
      End
      Begin VB.Menu MenuAbout1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout2 
         Caption         =   "版本：v20200520"
      End
      Begin VB.Menu MenuAbout3 
         Caption         =   "版权：(C) 2020 Anonym."
      End
      Begin VB.Menu MenuAbout4 
         Caption         =   "协议：GNU GPL v3，CC BY-NC 3.0"
      End
   End
   Begin VB.Menu Menu2_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuLanguage 
      Caption         =   "Ａ字あ (&L)"
      Begin VB.Menu MenuLanguageENG 
         Caption         =   "English (United States)"
         Enabled         =   0   'False
         Shortcut        =   +{F1}
      End
      Begin VB.Menu MenuLanguageCHS 
         Caption         =   "中文（简体）"
         Checked         =   -1  'True
         Shortcut        =   +{F2}
      End
      Begin VB.Menu MenuLanguageCHT 
         Caption         =   "中文（繁體）"
         Enabled         =   0   'False
         Shortcut        =   +{F3}
      End
      Begin VB.Menu MenuLanguageJPN 
         Caption         =   "日本語"
         Enabled         =   0   'False
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu Menu3_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuEXIT 
      Caption         =   "退出 (&X)"
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Option Explicit

'Declare Menu...
Public setlanguage As String
Public soundswitch As Boolean

'Declare Game...
Public gamestatus As Integer  '9-Welcome, 0-Initial, 3-Ready, 1-Ongoing, 2-Interval, 7-Paused, 4-Stopped.
Public gameresult As Integer  '0-None, 1-Winner, 2-Loser.
Public gameprogress As Single  'Range: 0~100.00
Public gametotalkana As Integer
Public gamekanarepeatedtimescount As Integer
Public gametotalcount As Integer
Public gamecombocount As Integer
Public gamecombobest As Integer
Public gamemistakecount As Integer
Public gametimeelapsed As Long  'Unit: 0.1s. eg. 10 -> 1 秒
Public gamedifficultyindex As Integer  'Range: 0~1000
Public gamecurrentdifficulty As Integer  'Unit: 0.1. eg. 50 -> 5 秒
Public gamecurrenttimeleft As Integer  'Unit: 0.1s. eg. 10 -> 1 秒
Public gameaveragereactiontime As Single  'Unit: 0.1s. eg. 10 -> 1 秒

Public lotterytotal As Integer
Public lotterynumber As Integer

Public lotterykana As String
Public lotterykanalocationX As Integer
Public lotterykanalocationY As Integer
Public kanadata As Variant  '(1 To 3, 1 To 74)
Public kanarepeatedtimesdata As Variant  '(1 To 3, 1 To 74)

Public correspondingromaji As String
Public lotteryromajilocationX As Integer
Public lotteryromajilocationY As Integer
Public romajidata As Variant  '(1 To 3, 1 To 74)

Public correctanswer As Integer
Public chosenanswer As Integer

'Declare Display...
Public gameprogressprogressbaranimationtarget As Long  'Range: 0~12300
Public gamecurrentdifficultyprogressbaranimationtarget As Long  'Range: 0~5580
Public gamecurrenttimeleftprogressbaranimationtarget As Long  'Range: 0~5580
Public spinningsakuracurrentangle As Single  'Range: -180.000~180.000. Note: 90.000 means straight up.
Public spinningsakuracurrentangle2 As Single
Public spinningsakuracurrentangle3 As Single
Public spinningsakuracurrentangle4 As Single
Public spinningsakuracurrentangle5 As Single
Public spinningsakuracurrentspeed As Single  'Range: 0.00~10.00
Public spinningsakuratargetspeed As Single  'Range: 0.00~10.00. Note: The maximum spinning speed is based on the current difficulty.

'Declare Settings...
Public gamedifficultyindexindicatordescription As String

Public setinputoption As Variant  '(1 To 3)

Public setkanaswitch As Variant  '(1 To 3)

Public setgamemode As Integer  '1-Kana, 2-Time.
Public setrepeatedtimes As Integer  'Range: 1~10
Public setspecifiedtime As Integer  'Unit: min. Range: 1~30 min.

Public setnormaldifficulty As Integer  'Unit: 0.1. eg. 20 -> 2 秒 Range: 2~50
Public setincreasedifficultygraduallyswitch As Boolean
Public setinitialdifficulty As Integer  'Unit: 0.1. eg. 50 -> 5 秒 Range: 2~50
Public setreachnormaldifficultyat As Integer  'Range: 0~100
Public setinterval As Integer  'Unit: 0.1. eg. 10 -> 1.0 秒 Range: 1~30
Public setmistakeallowedamount As Integer  'Range: 0~10

Public setblackonwhite As Boolean
Public setreducecontrast As Boolean
Public setanimationswitch As Boolean
Public sethideunnecessaryinfo As Boolean
Public setspinningsakuraswitch As Boolean

Public setcheatingswitch As Boolean
Public setcheatingshowcorrectanswer As Boolean

Public setfontswitch As Boolean

'Declare Others...
Public forloop1 As Integer
Public forloop2 As Integer
Public forloop3 As Integer

'Declare Dialog...
Public answer

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Sub Form_Load()
        'Load and Initialization...

        'Initialize Menu...
        setlanguage = "CHS"
        soundswitch = True

        'Initialize Game...
        gamestatus = 9
        gameresult = 0
        gameprogress = 0
        gametotalkana = 0
        gamekanarepeatedtimescount = 0
        gametotalcount = 0
        gamecombocount = 0
        gamecombobest = 0
        gamemistakecount = 0
        gametimeelapsed = 0
        gamedifficultyindex = 0
        gamecurrentdifficulty = 0
        gamecurrenttimeleft = 0
        gameaveragereactiontime = 0

        lotterytotal = 0
        lotterynumber = 0
        lotterykana = "??"
        kanadata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                         Array("!!", "续", "蛤", "青", "改", "夏", "吉", "谦", "另", "高", "苟", "赛", "吼", "基", "钦", "无", "奉", "滋", "削", "图", "身", "西", "华", "谈", "风", "姿", "识", "捉", "跑", "森", "上", "拿", "抱", "长", "经", "碰", "闷", "发", "坠", "负", "特", "连", "要", "表", "民", "新", "批", "安", "不", "得", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--"), _
                         Array("!!", "い", "祈", "小", "维", "猪", "庆", "包", "吸", "禁", "倒", "星", "轻", "易", "通", "宽", "金", "律", "绿", "颐", "气", "冰", "岿", "大", "掀", "池", "风", "雨", "萨", "格", "尔", "吃", "没", "麦", "十", "山", "二", "百", "换", "突", "满", "喷", "梁", "沼", "精", "细", "工", "八", "撸", "不", "自", "困", "艰", "奋", "苦", "逆", "没", "发", "时", "读", "书", "闹", "清", "应", "神", "敬", "坡", "汹", "找", "瞻", "游", "亲", "谭", "麻", "泼"), _
                         Array("!!", "膜", "品", "赵", "共", "称", "言", "粉", "五", "网", "干", "反", "翻", "一", "胡", "法", "坏", "逼", "支", "抖", "辣", "厉", "墙", "六", "坦", "铁", "螳", "当", "腊", "耿", "战", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--") _
                         )
        kanarepeatedtimesdata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                      )

        correspondingromaji = "??"
        romajidata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                           Array("!!", "命", "蟆", "蛙", "变", "威", "他", "虚", "请", "明", "利", "艇", "啊", "本", "点", "可", "告", "磁", "习", "样", "经", "方", "莱", "笑", "生", "势", "得", "急", "快", "破", "台", "衣", "歉", "者", "验", "到", "声", "财", "吼", "泽", "首", "任", "要", "态", "白", "闻", "判", "轨", "行", "罪", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--"), _
                           Array("!!", "よ", "翠", "熊", "尼", "头", "丰", "子", "精", "评", "车", "瀚", "关", "道", "商", "衣", "科", "玉", "玉", "使", "指", "棒", "然", "海", "翻", "塘", "狂", "骤", "格", "尔", "王", "饱", "事", "子", "里", "路", "百", "斤", "肩", "开", "脸", "粪", "家", "气", "甚", "腻", "笔", "千", "袖", "强", "息", "难", "苦", "斗", "吃", "差", "有", "酵", "代", "过", "单", "欢", "单", "验", "明", "畏", "涛", "涌", "准", "养", "泳", "自", "德", "批", "鸡"), _
                           Array("!!", "乎", "韭", "弹", "惨", "帝", "论", "蛆", "毛", "紧", "五", "贼", "车", "派", "言", "轮", "球", "站", "乎", "阴", "椒", "害", "国", "四", "克", "骑", "臂", "车", "肉", "爽", "狼", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--") _
                           )

        correctanswer = 0
        chosenanswer = 0

        'Initialize Display...
        gameprogressprogressbaranimationtarget = 0
        gamecurrentdifficultyprogressbaranimationtarget = 0
        gamecurrenttimeleftprogressbaranimationtarget = 0
        spinningsakuracurrentangle = 90
        spinningsakuracurrentspeed = 0
        spinningsakuratargetspeed = 0

        'Initialize Settings...
        gamedifficultyindexindicatordescription = "当前难度指数的描述..."

        setinputoption = Array("!!", "1", "2", "3")

        setkanaswitch = Array("!!", True, True, True, True, True, True, True, True, True, True, False)

        setgamemode = 1
        setrepeatedtimes = 1
        setspecifiedtime = 5

        setnormaldifficulty = 20
        setincreasedifficultygraduallyswitch = True
        setinitialdifficulty = 50
        setreachnormaldifficultyat = 20
        setinterval = 10
        setmistakeallowedamount = 3

        setblackonwhite = False
        setreducecontrast = False
        setanimationswitch = True
        sethideunnecessaryinfo = False
        setspinningsakuraswitch = True

        setcheatingswitch = False
        setcheatingshowcorrectanswer = True

        setfontswitch = False
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    'CMD Language...
    Public Sub MenuLanguageENG_Click()
        'Call ModuleLoadLanguage.LoadLanguageENG
    End Sub
    Public Sub MenuLanguageCHS_Click()
        'Call ModuleLoadLanguage.LoadLanguageCHS
    End Sub
    Public Sub MenuLanguageCHT_Click()
        'Call ModuleLoadLanguage.LoadLanguageCHT
    End Sub
    Public Sub MenuLanguageJPN_Click()
        'Call ModuleLoadLanguage.LoadLanguageJPN
    End Sub

    'CMD Menu...
    Public Sub MenuEXIT_Click()
        End
    End Sub
    Public Sub CmdEXIT_Click()
        Call MenuEXIT_Click
    End Sub
    Public Sub MenuSettings_Click()
        FormSettings.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormSettings.windowanimationtargetleft = (Screen.Width / 2) - (13560 / 2)
        FormSettings.windowanimationtargettop = (Screen.Height / 2) - (9465 / 2)
        FormSettings.windowanimationtargetwidth = 13560
        FormSettings.windowanimationtargetheight = 9465
        FormSettings.Show
    End Sub
    Public Sub MenuSoundSwitch_Click()
        Select Case soundswitch
            Case True
                soundswitch = False
                MenuSoundSwitch.Caption = "声音 关 (&D)"
            Case False
                soundswitch = True
                MenuSoundSwitch.Caption = "声音 开 (&D)"
        End Select
    End Sub

    'CMD Game...
    Public Sub MenuGameStartPauseResume_Click()
        Select Case gamestatus
            Case 9  'Status: Welcome...
                gamestatus = 9  'Hold it...
            Case 0  'Status: Initial...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Proximity Notification.wav"
                gamestatus = 3  'Into: Ready...
            Case 3  'Status: Ready...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
                gamestatus = 7  'Into: Paused...
            Case 1  'Status: Ongoing...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
                gamestatus = 7  'Into: Paused...
            Case 2  'Status: Interval...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
                gamestatus = 7  'Into: Paused...
            Case 7  'Status: Paused...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
                gamestatus = 2  'Into: Interval...
            Case 4  'Status: Stopped...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Proximity Notification.wav"
                gamestatus = 3  'Into: Ready...
        End Select

        TextboxInput.SetFocus: Call GameStatusRefresher
    End Sub
    Public Sub CmdStartPauseResume_Click()
        Call MenuGameStartPauseResume_Click
    End Sub
    Public Sub MenuGameStop_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Speech Off.wav"
        gamestatus = 4: Call GameStatusRefresher: TextboxInput.SetFocus
    End Sub
    Public Sub CmdStop_Click()
        Call MenuGameStop_Click
    End Sub

    Public Sub MenuGameChooseOption1_Click()
        chosenanswer = 1
        If gamestatus = 1 Then Call GameRespondent
        'TextboxInput.SetFocus
    End Sub
    Public Sub CmdOption1_Click()
        Call MenuGameChooseOption1_Click
    End Sub
    Public Sub MenuGameChooseOption2_Click()
        chosenanswer = 2
        If gamestatus = 1 Then Call GameRespondent
        'TextboxInput.SetFocus
    End Sub
    Public Sub CmdOption2_Click()
        Call MenuGameChooseOption2_Click
    End Sub
    Public Sub MenuGameChooseOption3_Click()
        chosenanswer = 3
        If gamestatus = 1 Then Call GameRespondent
        'TextboxInput.SetFocus
    End Sub
    Public Sub CmdOption3_Click()
        Call MenuGameChooseOption3_Click
    End Sub

    Public Sub TextboxInput_Change()
        Select Case TextboxInput.Text
            Case setinputoption(1)
                Call MenuGameChooseOption1_Click
            Case setinputoption(2)
                Call MenuGameChooseOption2_Click
            Case setinputoption(3)
                Call MenuGameChooseOption3_Click
            Case ""
                Exit Sub
            Case Else
                MsgBox "注意：无效输入。您按下了错误的按键。请确认您的手指是否抵在了正确的三个按键上面。", vbExclamation + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
        End Select

        TextboxInput.Text = ""
    End Sub

'[] TIMERS []

    Public Sub TimerClock_Timer()
        LabelClock.Caption = Format((Hour(Time)), "00") & ":" & Format((Minute(Time)), "00") & ":" & Format((Second(Time)), "00")

        If gamestatus = 9 Then  'Initialize...
            gamestatus = 0: Call GameStatusRefresher
        End If
    End Sub

    Public Sub TimerSettingsRefresher_Timer()
        'Game Difficulty Index Indicator...
            If gamedifficultyindex >= 0 Then gamedifficultyindexindicatordescription = "貌似过于简单了？"
            If gamedifficultyindex >= 200 Then gamedifficultyindexindicatordescription = "初来乍到"
            If gamedifficultyindex >= 400 Then gamedifficultyindexindicatordescription = "普通级别"
            If gamedifficultyindex >= 500 Then gamedifficultyindexindicatordescription = "困难级别"
            If gamedifficultyindex >= 600 Then gamedifficultyindexindicatordescription = "香港记者"
            If gamedifficultyindex >= 700 Then gamedifficultyindexindicatordescription = "究极反贼"

            FormSettings.LabelGameDifficultyIndexIndicator1.Caption = gamedifficultyindex
            FormSettings.gamedifficultyindexprogressbaranimationtarget = gamedifficultyindex / 1000 * 8000
            FormSettings.LabelGameDifficultyIndexIndicator3.Caption = gamedifficultyindexindicatordescription

        'Input...
            If Not (FormSettings.TextboxInputOption1.Text = "") Then setinputoption(1) = FormSettings.TextboxInputOption1.Text: FormSettings.LabelInputOption1Indicator.Caption = setinputoption(1): LabelOption1.Caption = setinputoption(1): FormSettings.TextboxInputOption1.Text = ""
            If Not (FormSettings.TextboxInputOption2.Text = "") Then setinputoption(2) = FormSettings.TextboxInputOption2.Text: FormSettings.LabelInputOption2Indicator.Caption = setinputoption(2): LabelOption2.Caption = setinputoption(2): FormSettings.TextboxInputOption2.Text = ""
            If Not (FormSettings.TextboxInputOption3.Text = "") Then setinputoption(3) = FormSettings.TextboxInputOption3.Text: FormSettings.LabelInputOption3Indicator.Caption = setinputoption(3): LabelOption3.Caption = setinputoption(3): FormSettings.TextboxInputOption3.Text = ""

        'Kana Included...
            If FormSettings.CheckboxKanaIncluded01.Value = 1 Then setkanaswitch(1) = True Else setkanaswitch(1) = False
            If FormSettings.CheckboxKanaIncluded02.Value = 1 Then setkanaswitch(2) = True Else setkanaswitch(2) = False
            If FormSettings.CheckboxKanaIncluded03.Value = 1 Then setkanaswitch(3) = True Else setkanaswitch(3) = False

        'Game Mode...
            If FormSettings.RadioboxGameModeKana.Value = True Then setgamemode = 1
            If FormSettings.RadioboxGameModeTime.Value = True Then setgamemode = 2
            setrepeatedtimes = FormSettings.HScrollGameModeRepeatedTimes.Value
            setspecifiedtime = FormSettings.HScrollGameModeSpecifiedTime.Value
            FormSettings.LabelGameModeRepeatedTimesIndicator.Caption = setrepeatedtimes
            FormSettings.LabelGameModeSpecifiedTimeIndicator.Caption = setspecifiedtime & " 分钟"

        'Difficulty...
            FormSettings.HScrollDifficultyNormalDifficulty.Max = FormSettings.HScrollDifficultyInitialDifficulty.Value

            setnormaldifficulty = FormSettings.HScrollDifficultyNormalDifficulty.Value
            FormSettings.LabelDifficultyNormalDifficultyIndicator.Caption = Format((setnormaldifficulty / 10), "0.0") & " 秒"

            If FormSettings.CheckboxDifficultyIncreaseDifficultyGradually.Value = 1 Then setincreasedifficultygraduallyswitch = True Else setincreasedifficultygraduallyswitch = False
                setinitialdifficulty = FormSettings.HScrollDifficultyInitialDifficulty.Value
                FormSettings.LabelDifficultyInitialDifficultyIndicator.Caption = Format((setinitialdifficulty / 10), "0.0") & " 秒"
                setreachnormaldifficultyat = FormSettings.HScrollDifficultyReachNormalDifficultyAt.Value
                FormSettings.LabelDifficultyReachNormalDifficultyAtIndicator.Caption = setreachnormaldifficultyat & "%"

            setinterval = FormSettings.HScrollDifficultyInterval.Value
            FormSettings.LabelDifficultyIntervalIndicator.Caption = Format((setinterval / 10), "0.0") & " 秒"
            setmistakeallowedamount = FormSettings.HScrollDifficultyMistakeAllowedAmount.Value
            FormSettings.LabelDifficultyMistakeAllowedAmountIndicator.Caption = setmistakeallowedamount

        'Display...
            If FormSettings.CheckboxDisplayBlackOnWhite.Value = 1 Then
                setblackonwhite = True
                LabelKanaDashboard.BackColor = &H0&: LabelKanaDashboard.ForeColor = &HFFFFFF
            Else
                setblackonwhite = False
                LabelKanaDashboard.BackColor = &HFFFFFF: LabelKanaDashboard.ForeColor = &H0&
            End If

            If FormSettings.CheckboxDisplayReduceContrast.Value = 1 Then
                setreducecontrast = True
                If setblackonwhite = True Then LabelKanaDashboard.BackColor = &H404040 Else LabelKanaDashboard.BackColor = &HE0E0E0
            Else
                setreducecontrast = False
                If setblackonwhite = True Then LabelKanaDashboard.BackColor = &H0 Else LabelKanaDashboard.BackColor = &HFFFFFF
            End If

            If FormSettings.CheckboxDisplaySmoothAnimations.Value = 1 Then setanimationswitch = True Else setanimationswitch = False

            If FormSettings.CheckboxDisplayHideUnnecessaryInformation.Value = 1 Then
                sethideunnecessaryinfo = True
                LabelStatusbar.Visible = False
                LabelGameCurrentDifficultyTitle.Visible = False: LabelGameCurrentDifficultyIndicator.Visible = False
                LabelGameProgressTitle.Visible = False: LabelGameProgressIndicator.Visible = False
                LabelGameDifficultyIndexTitle.Visible = False: LabelGameDifficultyIndexIndicator.Visible = False
                LabelGameCurrentTimeLeftTitle.Visible = False: LabelGameCurrentTimeLeftIndicator.Visible = False
                LabelGameTimeElapsedTitle.Visible = False: LabelGameTimeElapsedIndicator.Visible = False
                LabelGameAverageReactionTimeTitle.Visible = False: LabelGameAverageReactionTimeIndicator.Visible = False
                LabelOption1.Visible = False: LabelOption2.Visible = False: LabelOption3.Visible = False
            Else
                sethideunnecessaryinfo = False
                LabelStatusbar.Visible = True
                LabelGameCurrentDifficultyTitle.Visible = True: LabelGameCurrentDifficultyIndicator.Visible = True
                LabelGameProgressTitle.Visible = True: LabelGameProgressIndicator.Visible = True
                LabelGameDifficultyIndexTitle.Visible = True: LabelGameDifficultyIndexIndicator.Visible = True
                LabelGameCurrentTimeLeftTitle.Visible = True: LabelGameCurrentTimeLeftIndicator.Visible = True
                LabelGameTimeElapsedTitle.Visible = True: LabelGameTimeElapsedIndicator.Visible = True
                LabelGameAverageReactionTimeTitle.Visible = True: LabelGameAverageReactionTimeIndicator.Visible = True
                LabelOption1.Visible = True: LabelOption2.Visible = True: LabelOption3.Visible = True
            End If

            If FormSettings.CheckboxDisplaySpinningSakura.Value = 1 Then
                setspinningsakuraswitch = True
                LineSpinningSakura1.Visible = True: LineSpinningSakura2.Visible = True: LineSpinningSakura3.Visible = True: LineSpinningSakura4.Visible = True: LineSpinningSakura5.Visible = True
            Else
                setspinningsakuraswitch = False
                LineSpinningSakura1.Visible = False: LineSpinningSakura2.Visible = False: LineSpinningSakura3.Visible = False: LineSpinningSakura4.Visible = False: LineSpinningSakura5.Visible = False
            End If

        'Cheating...
            If FormSettings.CheckboxCheatingSwitch.Value = 1 Then
                setcheatingswitch = True
                If FormSettings.CheckboxCheatingShowCorrectAnswer.Value = 1 Then setcheatingshowcorrectanswer = True Else setcheatingshowcorrectanswer = False
                FormSettings.CheckboxCheatingShowCorrectAnswer.Enabled = True
            Else
                setcheatingswitch = False: setcheatingshowcorrectanswer = False
                FormSettings.CheckboxCheatingShowCorrectAnswer.Enabled = False
            End If

        'Fonts (Beta)...
            If FormSettings.CheckboxFontsSwitch.Value = 1 Then
                FormSettings.TextboxFontsJpnFont.Enabled = True: FormSettings.TextboxFontsEngFont.Enabled = True: FormSettings.CmdFontsApply.Enabled = True
            Else
                FormSettings.TextboxFontsJpnFont.Enabled = False: FormSettings.TextboxFontsEngFont.Enabled = False: FormSettings.CmdFontsApply.Enabled = False
                FormMainWindow.LabelKanaDashboard.Font = "SimHei": FormMainWindow.CmdOption1.Font = "SimHei": FormMainWindow.CmdOption2.Font = "SimHei": FormMainWindow.CmdOption3.Font = "SimHei"
            End If
    End Sub

    Public Sub TimerCalculator_Timer()
        'Difficulty Index calculator...

            gamedifficultyindex = 0

            'Difficulty index calculation Part 1...
            If setkanaswitch(1) = True Then gamedifficultyindex = gamedifficultyindex + 100
            If setkanaswitch(2) = True Then gamedifficultyindex = gamedifficultyindex + 100
            If setkanaswitch(3) = True Then gamedifficultyindex = gamedifficultyindex + 50

            'Difficulty index calculaton Part 2...
            Select Case setgamemode
                Case 1
                    gamedifficultyindex = gamedifficultyindex + 150 ^ ((setrepeatedtimes - 1) / 9)
                Case 2
                    gamedifficultyindex = gamedifficultyindex + 150 ^ ((setspecifiedtime - 1) / 29)
            End Select

            'Difficulty index calculaton Part 3...
            gamedifficultyindex = gamedifficultyindex + 300 ^ ((50 - setnormaldifficulty) / 48)
            Select Case setincreasedifficultygraduallyswitch
                Case True
                    gamedifficultyindex = gamedifficultyindex + 60 ^ (1 - (setreachnormaldifficultyat / 100) * ((setinitialdifficulty - setnormaldifficulty) / 48))
                    gamedifficultyindex = gamedifficultyindex + 90 ^ (1 - (setreachnormaldifficultyat / 100) * ((setinitialdifficulty - setnormaldifficulty) / 48))
                Case False
                    gamedifficultyindex = gamedifficultyindex + 150
            End Select
            gamedifficultyindex = gamedifficultyindex + 50 ^ ((30 - setinterval) / 29)
            gamedifficultyindex = gamedifficultyindex + 100 ^ (1 - setmistakeallowedamount / 10)

            'Apply calculation result...
            LabelGameDifficultyIndexIndicator.Caption = gamedifficultyindex & " / 1000"

        'Game Progress calculator...

            gametotalkana = 0
            If setkanaswitch(1) = True Then gametotalkana = gametotalkana + 49
            If setkanaswitch(2) = True Then gametotalkana = gametotalkana + 74
            If setkanaswitch(3) = True Then gametotalkana = gametotalkana + 30

            'Prevent disabling all kanaswitch...
            If gametotalkana = 0 Then
                MsgBox "注意：不可以排除所有文字。将恢复 [包括的内容] 到默认设定。", vbExclamation + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
                'setkanaswitch = Array("!!", True, True, True, True, True, True, True, True, True, True, False)
                FormSettings.CheckboxKanaIncluded01.Value = 1: FormSettings.CheckboxKanaIncluded02.Value = 1: FormSettings.CheckboxKanaIncluded03.Value = 1
                Exit Sub
            End If

            Select Case setgamemode
                Case 1
                    gamekanarepeatedtimescount = 0
                    For forloop1 = 1 To 3
                        For forloop2 = 1 To 74
                            If kanarepeatedtimesdata(forloop1)(forloop2) >= setrepeatedtimes Then gamekanarepeatedtimescount = gamekanarepeatedtimescount + 1
                        Next
                    Next
                    gameprogress = (gamekanarepeatedtimescount / gametotalkana) * 100
                Case 2
                    gameprogress = (gametimeelapsed / (setspecifiedtime * 600)) * 100
            End Select

            'Apply calculation result...
            LabelGameProgressIndicator.Caption = Format(gameprogress, "0.00") & "%"
            gameprogressprogressbaranimationtarget = gameprogress / 100 * 12300
            If gameprogressprogressbaranimationtarget < 0 Then gameprogressprogressbaranimationtarget = 0
            If gameprogressprogressbaranimationtarget > 12300 Then gameprogressprogressbaranimationtarget = 12300

        'Current Difficulty calculator...

            Select Case setincreasedifficultygraduallyswitch
                Case True
                    If gameprogress < setreachnormaldifficultyat Then
                        gamecurrentdifficulty = setinitialdifficulty - (setinitialdifficulty - setnormaldifficulty) * (gameprogress / setreachnormaldifficultyat)
                    Else
                        gamecurrentdifficulty = setnormaldifficulty
                    End If
                Case False
                    gamecurrentdifficulty = setnormaldifficulty
            End Select

            LabelGameCurrentDifficultyIndicator.Caption = Format((gamecurrentdifficulty / 10), "0.0")
            If (setinitialdifficulty = setnormaldifficulty) Or (setincreasedifficultygraduallyswitch = False) Then
                gamecurrentdifficultyprogressbaranimationtarget = 0.5 * 5580
            Else
                gamecurrentdifficultyprogressbaranimationtarget = (setinitialdifficulty - gamecurrentdifficulty) / (setinitialdifficulty - setnormaldifficulty) * 5580
            End If
            If gamecurrentdifficultyprogressbaranimationtarget < 0 Then gamecurrentdifficultyprogressbaranimationtarget = 0
            If gamecurrentdifficultyprogressbaranimationtarget > 5580 Then gamecurrentdifficultyprogressbaranimationtarget = 5580

        'Time Left, Time Elapsed, and Average Reaction Time calculator...

            Select Case gamestatus
                Case 3
                    'New Game initialization...
                    gameresult = 0: gameprogress = 0: gamekanarepeatedtimescount = 0: gametotalcount = 0: gamecombocount = 0: gamecombobest = 0: gamemistakecount = 0: gametimeelapsed = 0: gameaveragereactiontime = 0
                    lotterytotal = 0: lotterynumber = 0: lotterykana = "??": correspondingromaji = "??": correctanswer = 0: chosenanswer = 0
                    kanarepeatedtimesdata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                                  )
                    gamecurrenttimeleft = gamecurrenttimeleft + 1
                    LabelStatusbar.Caption = "准备好了么！ --- " & Format(((30 - gamecurrenttimeleft) / 10), "0.0")
                    LabelKanaDashboard.Caption = Format(Int((40 - gamecurrenttimeleft) / 10), "0")

                    LabelGameCurrentTimeLeftIndicator.Caption = Format(((gamecurrenttimeleft / 30) * gamecurrentdifficulty / 10), "0.0")
                    gamecurrenttimeleftprogressbaranimationtarget = gamecurrenttimeleft / 30 * 5580
                    If gamecurrenttimeleftprogressbaranimationtarget < 0 Then gamecurrenttimeleftprogressbaranimationtarget = 0
                    If gamecurrenttimeleftprogressbaranimationtarget > 5580 Then gamecurrenttimeleftprogressbaranimationtarget = 5580

                    If gamecurrenttimeleft >= 30 Then
                        gamecurrenttimeleft = gamecurrentdifficulty
                        Call GameQuestioner: GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                    End If

                    ShapeLightIndicatorOption1.BackColor = &H808080
                    ShapeLightIndicatorOption2.BackColor = &H808080
                    ShapeLightIndicatorOption3.BackColor = &H808080
                Case 1
                    gametimeelapsed = gametimeelapsed + 1
                    gamecurrenttimeleft = gamecurrenttimeleft - 1
                    LabelStatusbar.Caption = "遍历文字" & gamekanarepeatedtimescount & "/" & gametotalkana & " --- 总计数" & gametotalcount & " --- " & gamecombocount & "连击 --- 最高" & gamecombobest & "连击 --- 失误数" & gamemistakecount & "/" & setmistakeallowedamount

                    LabelGameCurrentTimeLeftIndicator.Caption = Format((gamecurrenttimeleft / 10), "0.0")
                    gamecurrenttimeleftprogressbaranimationtarget = gamecurrenttimeleft / gamecurrentdifficulty * 5580
                    If gamecurrenttimeleftprogressbaranimationtarget < 0 Then gamecurrenttimeleftprogressbaranimationtarget = 0
                    If gamecurrenttimeleftprogressbaranimationtarget > 5580 Then gamecurrenttimeleftprogressbaranimationtarget = 5580

                    'Time up judgement...
                    If gamecurrenttimeleft <= 0 Then
                        chosenanswer = 4: Call GameRespondent: GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                    End If
                Case 2
                    gametimeelapsed = gametimeelapsed + 1
                    gamecurrenttimeleft = gamecurrenttimeleft + 1
                    LabelStatusbar.Caption = "遍历文字" & gamekanarepeatedtimescount & "/" & gametotalkana & " --- 总计数" & gametotalcount & " --- " & gamecombocount & "连击 --- 最高" & gamecombobest & "连击 --- 失误数" & gamemistakecount & "/" & setmistakeallowedamount

                    LabelGameCurrentTimeLeftIndicator.Caption = Format(((gamecurrenttimeleft / setinterval) * gamecurrentdifficulty / 10), "0.0")
                    gamecurrenttimeleftprogressbaranimationtarget = gamecurrenttimeleft / setinterval * 5580
                    If gamecurrenttimeleftprogressbaranimationtarget < 0 Then gamecurrenttimeleftprogressbaranimationtarget = 0
                    If gamecurrenttimeleftprogressbaranimationtarget > 5580 Then gamecurrenttimeleftprogressbaranimationtarget = 5580

                    'Time up judgement...
                    If gamecurrenttimeleft >= setinterval Then
                        Call GameQuestioner: GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                    End If
                Case 7
                    GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                Case 9
                    GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                    Case 0
                        GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                        Case 4
                            GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                Case Else
                    MsgBox "错误：Game status is out of range." & vbCrLf & "请向我们提供反馈以帮助解决问题。感谢您的支持！", vbCritical + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
            End Select

TimerCalculator_ForceExitSelectCaseGameStatus_:

            LabelGameTimeElapsedIndicator.Caption = (Format(Int(gametimeelapsed / 600), "00")) & "' " & (Format((Int(gametimeelapsed / 10) Mod 60), "00")) & """ " & (Format((gametimeelapsed Mod 10), "0"))
            LabelGameAverageReactionTimeIndicator.Caption = Format((gameaveragereactiontime / 10), "0.000")
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerProgressbarAnimation_Timer()
        Select Case setanimationswitch
            Case True
                If ShapeGameProgressProgressbar.Width = gameprogressprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip1_
                If ShapeGameProgressProgressbar.Width > gameprogressprogressbaranimationtarget Then ShapeGameProgressProgressbar.Width = ShapeGameProgressProgressbar.Width - Abs(ShapeGameProgressProgressbar.Width - gameprogressprogressbaranimationtarget) / 4
                If ShapeGameProgressProgressbar.Width < gameprogressprogressbaranimationtarget Then ShapeGameProgressProgressbar.Width = ShapeGameProgressProgressbar.Width + Abs(ShapeGameProgressProgressbar.Width - gameprogressprogressbaranimationtarget) / 4
                If Abs(ShapeGameProgressProgressbar.Width - gameprogressprogressbaranimationtarget) < 10 Then ShapeGameProgressProgressbar.Width = gameprogressprogressbaranimationtarget
TimerProgressbarAnimation_Skip1_:

                If ShapeGameCurrentDifficultyProgressbar.Height = gamecurrentdifficultyprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip2_
                If ShapeGameCurrentDifficultyProgressbar.Height > gamecurrentdifficultyprogressbaranimationtarget Then ShapeGameCurrentDifficultyProgressbar.Height = ShapeGameCurrentDifficultyProgressbar.Height - Abs(ShapeGameCurrentDifficultyProgressbar.Height - gamecurrentdifficultyprogressbaranimationtarget) / 4
                If ShapeGameCurrentDifficultyProgressbar.Height < gamecurrentdifficultyprogressbaranimationtarget Then ShapeGameCurrentDifficultyProgressbar.Height = ShapeGameCurrentDifficultyProgressbar.Height + Abs(ShapeGameCurrentDifficultyProgressbar.Height - gamecurrentdifficultyprogressbaranimationtarget) / 4
                If Abs(ShapeGameCurrentDifficultyProgressbar.Height - gamecurrentdifficultyprogressbaranimationtarget) < 10 Then ShapeGameCurrentDifficultyProgressbar.Height = gamecurrentdifficultyprogressbaranimationtarget
                   ShapeGameCurrentDifficultyProgressbar.Top = 7050 - ShapeGameCurrentDifficultyProgressbar.Height
TimerProgressbarAnimation_Skip2_:

                If ShapeGameCurrentTimeLeftProgressbar.Height = gamecurrenttimeleftprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip3_
                If ShapeGameCurrentTimeLeftProgressbar.Height > gamecurrenttimeleftprogressbaranimationtarget Then ShapeGameCurrentTimeLeftProgressbar.Height = ShapeGameCurrentTimeLeftProgressbar.Height - Abs(ShapeGameCurrentTimeLeftProgressbar.Height - gamecurrenttimeleftprogressbaranimationtarget) / 4
                If ShapeGameCurrentTimeLeftProgressbar.Height < gamecurrenttimeleftprogressbaranimationtarget Then ShapeGameCurrentTimeLeftProgressbar.Height = ShapeGameCurrentTimeLeftProgressbar.Height + Abs(ShapeGameCurrentTimeLeftProgressbar.Height - gamecurrenttimeleftprogressbaranimationtarget) / 4
                If Abs(ShapeGameCurrentTimeLeftProgressbar.Height - gamecurrenttimeleftprogressbaranimationtarget) < 10 Then ShapeGameCurrentTimeLeftProgressbar.Height = gamecurrenttimeleftprogressbaranimationtarget
                   ShapeGameCurrentTimeLeftProgressbar.Top = 7050 - ShapeGameCurrentTimeLeftProgressbar.Height
TimerProgressbarAnimation_Skip3_:

            Case False
                If ShapeGameProgressProgressbar.Width = gameprogressprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip4_
                ShapeGameProgressProgressbar.Width = gameprogressprogressbaranimationtarget
TimerProgressbarAnimation_Skip4_:
                If ShapeGameCurrentDifficultyProgressbar.Height = gamecurrentdifficultyprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip5_
                ShapeGameCurrentDifficultyProgressbar.Height = gamecurrentdifficultyprogressbaranimationtarget: ShapeGameCurrentDifficultyProgressbar.Top = 7050 - ShapeGameCurrentDifficultyProgressbar.Height
TimerProgressbarAnimation_Skip5_:
                If ShapeGameCurrentTimeLeftProgressbar.Height = gamecurrenttimeleftprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip6_
                ShapeGameCurrentTimeLeftProgressbar.Height = gamecurrenttimeleftprogressbaranimationtarget: ShapeGameCurrentTimeLeftProgressbar.Top = 7050 - ShapeGameCurrentTimeLeftProgressbar.Height
TimerProgressbarAnimation_Skip6_:

        End Select
    End Sub

    Public Sub TimerSpinningSakuraAnimation_Timer()
        If (gamestatus = 1 Or gamestatus = 2 Or gamestatus = 3) Then
            spinningsakuratargetspeed = (70 - gamecurrentdifficulty) / 68 * 10
        Else
            spinningsakuratargetspeed = 0
        End If

        'Locate (2205+ShapeGameProgressProgressbar.Width, 1000) ...
        LineSpinningSakura1.X1 = 2205 + ShapeGameProgressProgressbar.Width: LineSpinningSakura1.Y1 = 1000
        LineSpinningSakura2.X1 = 2205 + ShapeGameProgressProgressbar.Width: LineSpinningSakura2.Y1 = 1000
        LineSpinningSakura3.X1 = 2205 + ShapeGameProgressProgressbar.Width: LineSpinningSakura3.Y1 = 1000
        LineSpinningSakura4.X1 = 2205 + ShapeGameProgressProgressbar.Width: LineSpinningSakura4.Y1 = 1000
        LineSpinningSakura5.X1 = 2205 + ShapeGameProgressProgressbar.Width: LineSpinningSakura5.Y1 = 1000

        'Make flower (Length set to 250) ...
        spinningsakuracurrentangle2 = spinningsakuracurrentangle - 360 / 5 * 1
        spinningsakuracurrentangle3 = spinningsakuracurrentangle - 360 / 5 * 2
        spinningsakuracurrentangle4 = spinningsakuracurrentangle - 360 / 5 * 3
        spinningsakuracurrentangle5 = spinningsakuracurrentangle - 360 / 5 * 4
        While spinningsakuracurrentangle2 < -180: spinningsakuracurrentangle2 = spinningsakuracurrentangle2 + 360: Wend
        While spinningsakuracurrentangle3 < -180: spinningsakuracurrentangle3 = spinningsakuracurrentangle3 + 360: Wend
        While spinningsakuracurrentangle4 < -180: spinningsakuracurrentangle4 = spinningsakuracurrentangle4 + 360: Wend
        While spinningsakuracurrentangle5 < -180: spinningsakuracurrentangle5 = spinningsakuracurrentangle5 + 360: Wend

        LineSpinningSakura1.X2 = LineSpinningSakura1.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle)
        LineSpinningSakura1.Y2 = LineSpinningSakura1.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle)
        LineSpinningSakura2.X2 = LineSpinningSakura2.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle2)
        LineSpinningSakura2.Y2 = LineSpinningSakura2.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle2)
        LineSpinningSakura3.X2 = LineSpinningSakura3.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle3)
        LineSpinningSakura3.Y2 = LineSpinningSakura3.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle3)
        LineSpinningSakura4.X2 = LineSpinningSakura4.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle4)
        LineSpinningSakura4.Y2 = LineSpinningSakura4.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle4)
        LineSpinningSakura5.X2 = LineSpinningSakura5.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle5)
        LineSpinningSakura5.Y2 = LineSpinningSakura5.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle5)

        'Prevent constant blinking...
        If (spinningsakuratargetspeed = 0 And spinningsakuracurrentspeed = 0) Then Exit Sub

        'Spin...
        Select Case setanimationswitch
            Case True
                spinningsakuracurrentangle = spinningsakuracurrentangle - spinningsakuracurrentspeed
                If spinningsakuracurrentangle <= -180 Then spinningsakuracurrentangle = spinningsakuracurrentangle + 360
            Case False
                spinningsakuracurrentangle = 90
        End Select

        'Adjust spinning speed...
        If spinningsakuracurrentspeed < spinningsakuratargetspeed Then spinningsakuracurrentspeed = spinningsakuracurrentspeed + 0.1
        If spinningsakuracurrentspeed > spinningsakuratargetspeed Then spinningsakuracurrentspeed = spinningsakuracurrentspeed - 0.05
        If spinningsakuracurrentspeed < 0 Then spinningsakuracurrentspeed = 0
        If spinningsakuracurrentspeed > 10 Then spinningsakuracurrentspeed = 10
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] GAME ENGINE []

    Public Sub GameStatusRefresher()
        Select Case gamestatus
            Case 9
                LabelStatusbar.Caption = "载入中..."
            Case 0
                'Reset the time left...
                gamecurrenttimeleft = 0

                LabelStatusbar.Caption = "点击 [开始] 来撸起袖子大干一场！"
                MenuGameStartPauseResume.Caption = "开始": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = False
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = True: MenuAbout.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "开始": CmdStop.Enabled = False
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False: CmdOption1.Caption = "?": CmdOption2.Caption = "?": CmdOption3.Caption = "?": LabelKanaDashboard.Caption = "?"
            Case 3
                LabelStatusbar.Caption = ""
                MenuGameStartPauseResume.Caption = "准备": MenuGameStartPauseResume.Enabled = False: MenuGameStop.Enabled = True
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = False: MenuAbout.Enabled = False
                CmdStartPauseResume.Enabled = False: CmdStartPauseResume.Caption = "准备": CmdStop.Enabled = True
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False: CmdOption1.Caption = "?": CmdOption2.Caption = "?": CmdOption3.Caption = "?"
            Case 1
                LabelStatusbar.Caption = ""
                MenuGameStartPauseResume.Caption = "暂停": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = True
                MenuGameChooseOption1.Enabled = True: MenuGameChooseOption2.Enabled = True: MenuGameChooseOption3.Enabled = True
                MenuSettings.Enabled = False: MenuAbout.Enabled = False
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "暂停": CmdStop.Enabled = True
                CmdOption1.Enabled = True: CmdOption2.Enabled = True: CmdOption3.Enabled = True
            Case 2
                LabelStatusbar.Caption = ""
                MenuGameStartPauseResume.Caption = "暂停": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = True
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = False: MenuAbout.Enabled = False
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "暂停": CmdStop.Enabled = True
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False
            Case 7
                'Reset the time left...
                gamecurrenttimeleft = 0

                LabelStatusbar.Caption = "游戏暂停 --- 点击 [继续] 来一起继续摇摆！"
                MenuGameStartPauseResume.Caption = "继续": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = True
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = False: MenuAbout.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "继续": CmdStop.Enabled = True
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False: CmdOption1.Caption = "?": CmdOption2.Caption = "?": CmdOption3.Caption = "?": LabelKanaDashboard.Caption = "?"
            Case 4
                'Reset the time left...
                gamecurrenttimeleft = 0

                LabelStatusbar.Caption = "游戏已停止 --- 点击 [开始] 重装上阵！"
                MenuGameStartPauseResume.Caption = "开始": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = False
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = True: MenuAbout.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "开始": CmdStop.Enabled = False
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False: CmdOption1.Caption = "?": CmdOption2.Caption = "?": CmdOption3.Caption = "?": LabelKanaDashboard.Caption = "?"
            Case Else
                MsgBox "错误：Game status is out of range." & vbCrLf & "请向我们提供反馈以帮助解决问题。感谢您的支持！", vbCritical + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
        End Select
    End Sub

    Public Sub RandomNumberGenerator()
        Randomize
        lotterynumber = Int((lotterytotal + 1) * Rnd)
        While lotterynumber = 0
            Randomize
            lotterynumber = Int((lotterytotal + 1) * Rnd)
        Wend
    End Sub

    Public Sub GameQuestioner()
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Menu Command.wav"

        gamestatus = 1: Call GameStatusRefresher: gamecurrenttimeleft = gamecurrentdifficulty
        If gameprogress >= 100 Then Exit Sub

        'Clear contents...
        LabelStatusbar.BackColor = &HC0C0C0
        LabelKanaDashboard.Caption = ""
        CmdOption1.Caption = ""
        CmdOption2.Caption = ""
        CmdOption3.Caption = ""
        ShapeLightIndicatorOption1.BackColor = &H808080
        ShapeLightIndicatorOption2.BackColor = &H808080
        ShapeLightIndicatorOption3.BackColor = &H808080

        'Step 1: Kana...
            lotterytotal = 0: lotterynumber = 0: lotterykanalocationX = 0: lotterykanalocationY = 0
            Do Until Not (kanarepeatedtimesdata(lotterykanalocationX)(lotterykanalocationY) >= setrepeatedtimes Or kanadata(lotterykanalocationX)(lotterykanalocationY) = "!!" Or kanadata(lotterykanalocationX)(lotterykanalocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                lotterytotal = 3: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                lotterykanalocationX = lotterynumber
                lotterytotal = 74: lotterynumber = 0: Call RandomNumberGenerator
                lotterykanalocationY = lotterynumber
            Loop
            lotterykana = kanadata(lotterykanalocationX)(lotterykanalocationY)
            correspondingromaji = romajidata(lotterykanalocationX)(lotterykanalocationY)
            LabelKanaDashboard.Caption = lotterykana

        'Step 2: The correct option...
            lotterytotal = 3: lotterynumber = 0: Call RandomNumberGenerator: correctanswer = lotterynumber
            Select Case correctanswer
                Case 1
                    CmdOption1.Caption = correspondingromaji
                Case 2
                    CmdOption2.Caption = correspondingromaji
                Case 3
                    CmdOption3.Caption = correspondingromaji
                Case Else
                    MsgBox "错误：Correct answer is out of range." & vbCrLf & "请向我们提供反馈以帮助解决问题。感谢您的支持！", vbCritical + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
            End Select

        'Step 3: Other option 1...
            Select Case correctanswer
                Case 1
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption1.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 3: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 74: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption2.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
                Case 2
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption2.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 3: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 74: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption1.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
                Case 3
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption3.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 3: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 74: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption1.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
            End Select

        'Step 4: Other option 2...
            Select Case correctanswer
                Case 1
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption1.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption2.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 3: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 74: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption3.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
                Case 2
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption2.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption1.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 3: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 74: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption3.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
                Case 3
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption3.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption1.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 3: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 74: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption2.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
            End Select

        'Cheating...
            If setcheatingswitch = True Then
                LabelStatusbar.BackColor = &HFFFF&

                If setcheatingshowcorrectanswer = True Then
                    Select Case correctanswer
                        Case 1
                            ShapeLightIndicatorOption1.BackColor = &HFFFF&
                        Case 2
                            ShapeLightIndicatorOption2.BackColor = &HFFFF&
                        Case 3
                            ShapeLightIndicatorOption3.BackColor = &HFFFF&
                        Case Else
                            MsgBox "错误：Correct answer is out of range." & vbCrLf & "请向我们提供反馈以帮助解决问题。感谢您的支持！", vbCritical + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
                    End Select
                End If
            End If
    End Sub

    Public Sub GameRespondent()
        'Average reaction time calculation...
        gametotalcount = gametotalcount + 1
        gameaveragereactiontime = (gameaveragereactiontime * (gametotalcount - 1) + (gamecurrentdifficulty - gamecurrenttimeleft)) / gametotalcount

        'Switch game status...
        gamestatus = 2: Call GameStatusRefresher: gamecurrenttimeleft = 0

        'Judgement...
        Select Case correctanswer
            Case 1
                ShapeLightIndicatorOption1.BackColor = &HFF00&
                ShapeLightIndicatorOption2.BackColor = &H808080
                ShapeLightIndicatorOption3.BackColor = &H808080
            Case 2
                ShapeLightIndicatorOption1.BackColor = &H808080
                ShapeLightIndicatorOption2.BackColor = &HFF00&
                ShapeLightIndicatorOption3.BackColor = &H808080
            Case 3
                ShapeLightIndicatorOption1.BackColor = &H808080
                ShapeLightIndicatorOption2.BackColor = &H808080
                ShapeLightIndicatorOption3.BackColor = &HFF00&
            Case Else
                MsgBox "错误：Correct answer is out of range." & vbCrLf & "请向我们提供反馈以帮助解决问题。感谢您的支持！", vbCritical + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
        End Select

        If chosenanswer = correctanswer Then
            'Answer correct sound...
            If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Startup.wav"

            'Combo count...
            gamecombocount = gamecombocount + 1
            If gamecombobest < gamecombocount Then gamecombobest = gamecombocount

            If setgamemode = 1 Then kanarepeatedtimesdata(lotterykanalocationX)(lotterykanalocationY) = kanarepeatedtimesdata(lotterykanalocationX)(lotterykanalocationY) + 1

            'Winner judgement...
            Call TimerCalculator_Timer
            If gameprogress >= 100 Then
                FormGameReport.LabelGameReportWinnerLoser.Caption = "胜  利"
                If setcheatingswitch = True Then
                    FormGameReport.LabelGameReportCheatedIndicator.Visible = True
                Else
                    FormGameReport.LabelGameReportCheatedIndicator.Visible = False
                End If
                FormGameReport.LabelGameDifficultyIndexIndicator1.Caption = gamedifficultyindex
                FormGameReport.LabelGameDifficultyIndexIndicator3.Caption = FormSettings.LabelGameDifficultyIndexIndicator3.Caption
                FormGameReport.LabelGameProgressIndicator.Caption = LabelGameProgressIndicator.Caption
                FormGameReport.LabelGameCurrentDifficultyIndicator.Caption = LabelGameCurrentDifficultyIndicator.Caption & "s"
                FormGameReport.LabelGameAverageReactionTimeIndicator.Caption = LabelGameAverageReactionTimeIndicator.Caption & "s"
                FormGameReport.LabelGameTimeElapsedIndicator.Caption = LabelGameTimeElapsedIndicator.Caption
                FormGameReport.LabelGameTotalCountIndicator.Caption = gametotalcount
                FormGameReport.LabelGameComboBestIndicator.Caption = gamecombobest
                FormGameReport.LabelGameMistakeCountIndicator.Caption = gamemistakecount

                gameresult = 1: gamestatus = 0: Call GameStatusRefresher
                MsgBox "祝贺！！挑战成功！" & vbCrLf & "稍后将显示本局游戏的详情。", vbInformation + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"

                FormGameReport.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
                FormGameReport.windowanimationtargetleft = (Screen.Width / 2) - (13560 / 2)
                FormGameReport.windowanimationtargettop = (Screen.Height / 2) - (9465 / 2)
                FormGameReport.windowanimationtargetwidth = 13560
                FormGameReport.windowanimationtargetheight = 9465
                FormGameReport.Show
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\tada.wav"
            End If
        Else
            'Answer incorrect sound...
            If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\chord.wav"

            Select Case chosenanswer
                Case 1
                    ShapeLightIndicatorOption1.BackColor = &HFF&
                Case 2
                    ShapeLightIndicatorOption2.BackColor = &HFF&
                Case 3
                    ShapeLightIndicatorOption3.BackColor = &HFF&
                Case 4
                    If ShapeLightIndicatorOption1.BackColor = &H808080 Then ShapeLightIndicatorOption1.BackColor = &H80FF&
                    If ShapeLightIndicatorOption2.BackColor = &H808080 Then ShapeLightIndicatorOption2.BackColor = &H80FF&
                    If ShapeLightIndicatorOption3.BackColor = &H808080 Then ShapeLightIndicatorOption3.BackColor = &H80FF&
                Case Else
                    MsgBox "错误：Chosen answer is out of range." & vbCrLf & "请向我们提供反馈以帮助解决问题。感谢您的支持！", vbCritical + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"
            End Select

            'Combo reset... But do not reset best combo (gamecombobest)...
            gamecombocount = 0: gamemistakecount = gamemistakecount + 1

            'Loser judgement...
            Call TimerCalculator_Timer
            If gamemistakecount > setmistakeallowedamount Then
                FormGameReport.LabelGameReportWinnerLoser.Caption = "失  败"
                If setcheatingswitch = True Then
                    FormGameReport.LabelGameReportCheatedIndicator.Visible = True
                Else
                    FormGameReport.LabelGameReportCheatedIndicator.Visible = False
                End If
                FormGameReport.LabelGameDifficultyIndexIndicator1.Caption = gamedifficultyindex
                FormGameReport.LabelGameDifficultyIndexIndicator3.Caption = FormSettings.LabelGameDifficultyIndexIndicator3.Caption
                FormGameReport.LabelGameProgressIndicator.Caption = LabelGameProgressIndicator.Caption
                FormGameReport.LabelGameCurrentDifficultyIndicator.Caption = LabelGameCurrentDifficultyIndicator.Caption & "s"
                FormGameReport.LabelGameAverageReactionTimeIndicator.Caption = LabelGameAverageReactionTimeIndicator.Caption & "s"
                FormGameReport.LabelGameTimeElapsedIndicator.Caption = LabelGameTimeElapsedIndicator.Caption
                FormGameReport.LabelGameTotalCountIndicator.Caption = gametotalcount
                FormGameReport.LabelGameComboBestIndicator.Caption = gamecombobest
                FormGameReport.LabelGameMistakeCountIndicator.Caption = gamemistakecount

                gameresult = 2: gamestatus = 0: Call GameStatusRefresher
                MsgBox "很遗憾... 您未能挑战成功。" & vbCrLf & "您已完成进度 " & Format(gameprogress, "0.00") & "%.", vbInformation + vbOKOnly + vbDefaultButton1, "恶政隐文字游戏"

                FormGameReport.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
                FormGameReport.windowanimationtargetleft = (Screen.Width / 2) - (13560 / 2)
                FormGameReport.windowanimationtargettop = (Screen.Height / 2) - (9465 / 2)
                FormGameReport.windowanimationtargetwidth = 13560
                FormGameReport.windowanimationtargetheight = 9465
                FormGameReport.Show
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Print Complete.wav"
            End If
        End If
    End Sub
