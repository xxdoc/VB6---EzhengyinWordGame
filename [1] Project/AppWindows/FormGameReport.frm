VERSION 5.00
Begin VB.Form FormGameReport 
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
   Icon            =   "FormGameReport.frx":0000
   LinkTopic       =   "FormGameReport"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormGameReport.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   9465
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOK 
      Cancel          =   -1  'True
      Caption         =   "吼啊"
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
      MouseIcon       =   "FormGameReport.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   17640
      Top             =   10395
   End
   Begin VB.Label LabelGameReportCheatedIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "有作弊！"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   9345
      TabIndex        =   3
      Top             =   1785
      Width           =   4005
   End
   Begin VB.Label LabelGameMistakeCountIndicator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000060F0&
      Height          =   2010
      Left            =   9135
      TabIndex        =   21
      Top             =   6405
      Width           =   2640
   End
   Begin VB.Label LabelGameMistakeCountTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "次失误"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000060F0&
      Height          =   645
      Left            =   11025
      TabIndex        =   20
      Top             =   8190
      Width           =   1905
   End
   Begin VB.Label LabelGameComboBestIndicator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   1590
      Left            =   5040
      TabIndex        =   19
      Top             =   7560
      Width           =   4320
   End
   Begin VB.Label LabelGameComboBestTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "最高连击数"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   645
      Left            =   4515
      TabIndex        =   18
      Top             =   7035
      Width           =   3165
   End
   Begin VB.Label LabelGameTotalCountIndicator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   750
      Left            =   1785
      TabIndex        =   17
      Top             =   8085
      Width           =   2325
   End
   Begin VB.Label LabelGameTotalCountTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "总计数"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   645
      Left            =   1470
      TabIndex        =   16
      Top             =   7455
      Width           =   1590
   End
   Begin VB.Label LabelGameTimeElapsedIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C06000&
      Height          =   855
      Left            =   9240
      TabIndex        =   15
      Top             =   5565
      Width           =   3375
   End
   Begin VB.Label LabelGameTimeElapsedTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "本局续命"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C06000&
      Height          =   645
      Left            =   9135
      TabIndex        =   14
      Top             =   4935
      Width           =   3165
   End
   Begin VB.Label LabelGameAverageReactionTimeIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   63.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1380
      Left            =   4515
      TabIndex        =   13
      Top             =   5460
      Width           =   4320
   End
   Begin VB.Label LabelGameAverageReactionTimeTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "平均反应速度"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   645
      Left            =   4200
      TabIndex        =   12
      Top             =   4935
      Width           =   4005
   End
   Begin VB.Label LabelGameCurrentDifficultyIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   63.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1590
      Left            =   945
      TabIndex        =   11
      Top             =   5775
      Width           =   2850
   End
   Begin VB.Label LabelGameCurrentDifficultyTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "难度"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   645
      Left            =   735
      TabIndex        =   10
      Top             =   5250
      Width           =   2535
   End
   Begin VB.Label LabelGameProgressIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   645
      Left            =   10185
      TabIndex        =   9
      Top             =   3990
      Width           =   2535
   End
   Begin VB.Label LabelGameProgressTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "完成进度"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   645
      Left            =   9765
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label LabelGameDifficultyIndexIndicator3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "当前难度指数的描述..."
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   645
      Left            =   3570
      TabIndex        =   7
      Top             =   3675
      Width           =   4215
   End
   Begin VB.Label LabelGameDifficultyIndexIndicator2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "／1000"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   435
      Left            =   3150
      TabIndex        =   6
      Top             =   4410
      Width           =   1590
   End
   Begin VB.Label LabelGameDifficultyIndexIndicator1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1065
      Left            =   840
      TabIndex        =   5
      Top             =   3675
      Width           =   2535
   End
   Begin VB.Label LabelGameDifficultyIndexTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "难度指数"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   645
      Left            =   945
      TabIndex        =   4
      Top             =   3045
      Width           =   3900
   End
   Begin VB.Label LabelGameReportWinnerLoser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2010
      Left            =   735
      TabIndex        =   2
      Top             =   1050
      Width           =   12090
   End
   Begin VB.Label LabelGameReportTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "游戏结束了！"
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
Attribute VB_Name = "FormGameReport"
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

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    Public Sub CmdOK_Click()
        windowanimationtargetleft = (Screen.Width / 2)
        windowanimationtargettop = (Screen.Height / 2)
        windowanimationtargetwidth = 0
        windowanimationtargetheight = 0
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
