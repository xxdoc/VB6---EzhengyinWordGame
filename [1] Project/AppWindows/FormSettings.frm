VERSION 5.00
Begin VB.Form FormSettings 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   0  'None
   Caption         =   "EzhengyinWordGame"
   ClientHeight    =   10725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17970
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
   ScaleHeight     =   10725
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameFonts 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "�������  (Beta)"
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
      Height          =   3585
      Left            =   9135
      TabIndex        =   57
      Top             =   6825
      Width           =   8625
      Begin VB.CommandButton CmdFontsApply 
         Caption         =   "Ӧ��"
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
         Left            =   6825
         MouseIcon       =   "FormSettings.frx":015E
         MousePointer    =   99  'Custom
         TabIndex        =   63
         Top             =   2835
         Width           =   1485
      End
      Begin VB.TextBox TextboxFontsEngFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1680
         MousePointer    =   3  'I-Beam
         TabIndex        =   62
         Text            =   "SimHei"
         Top             =   1620
         Width           =   6630
      End
      Begin VB.TextBox TextboxFontsJpnFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1680
         MousePointer    =   3  'I-Beam
         TabIndex        =   60
         Text            =   "SimHei"
         Top             =   1080
         Width           =   6630
      End
      Begin VB.CheckBox CheckboxFontsSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "����������������ܵ��³����������С�ģ�"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   8100
      End
      Begin VB.Label LabelFontsEngFont 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ�����壺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   315
         TabIndex        =   61
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label LabelFontsJpnFont 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "�������壺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   315
         TabIndex        =   59
         Top             =   1155
         Width           =   1230
      End
   End
   Begin VB.Frame FrameKanaIncluded 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "����������  (ռ 250 ��)"
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
      TabIndex        =   18
      Top             =   2625
      Width           =   8625
      Begin VB.CheckBox CheckboxKanaIncluded03 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   5670
         MouseIcon       =   "FormSettings.frx":0402
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   420
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox CheckboxKanaIncluded02 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2940
         MouseIcon       =   "FormSettings.frx":0554
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   420
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox CheckboxKanaIncluded01 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Ĥ��"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   2535
      End
   End
   Begin VB.Frame FrameGameDifficultyIndexIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "�Ѷ�ָ��"
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
      Height          =   1485
      Left            =   210
      TabIndex        =   2
      Top             =   945
      Width           =   8625
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
         Left            =   8085
         MouseIcon       =   "FormSettings.frx":07F8
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   315
         Width           =   420
      End
      Begin VB.Timer TimerProgressbarAnimation 
         Interval        =   1
         Left            =   8295
         Top             =   1155
      End
      Begin VB.Shape ShapeGameDifficultyIndexIndicatorProgressbar 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         Height          =   120
         Left            =   315
         Top             =   1155
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
         Top             =   1155
         Width           =   8000
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
         Top             =   525
         Width           =   1395
      End
      Begin VB.Label LabelGameDifficultyIndexIndicator2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "��1000"
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
         Top             =   840
         Width           =   915
      End
      Begin VB.Label LabelGameDifficultyIndexIndicator3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ�Ѷ�ָ��������..."
         BeginProperty Font 
            Name            =   "����"
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
         Top             =   735
         Width           =   5640
      End
   End
   Begin VB.Frame FrameInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "��������"
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
      Height          =   1485
      Left            =   9135
      TabIndex        =   7
      Top             =   945
      Width           =   8625
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
         Left            =   7560
         MaxLength       =   1
         MousePointer    =   3  'I-Beam
         TabIndex        =   17
         Top             =   840
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
         Left            =   4830
         MaxLength       =   1
         MousePointer    =   3  'I-Beam
         TabIndex        =   14
         Top             =   840
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
         Left            =   2100
         MaxLength       =   1
         MousePointer    =   3  'I-Beam
         TabIndex        =   11
         Top             =   840
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
         Left            =   6930
         TabIndex        =   16
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LabelInputOption3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ��3"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5775
         TabIndex        =   15
         Top             =   945
         Width           =   1020
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
         Left            =   4200
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LabelInputOption2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ��2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3045
         TabIndex        =   12
         Top             =   945
         Width           =   1020
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
         Left            =   1470
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LabelInputOption1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ��1"
         BeginProperty Font 
            Name            =   "����"
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
         Top             =   945
         Width           =   1020
      End
      Begin VB.Label LabelInput 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ���Ӧ�İ�������ʾ����ʼ�տ���ʹ�� F6, F7 �� F8."
         BeginProperty Font 
            Name            =   "����"
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
         Top             =   525
         Width           =   8115
      End
   End
   Begin VB.Frame FrameDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "��ʾ"
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
      Height          =   1905
      Left            =   210
      TabIndex        =   48
      Top             =   6825
      Width           =   8625
      Begin VB.CheckBox CheckboxDisplaySpinningSakura 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "��ת��ά��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   4410
         MouseIcon       =   "FormSettings.frx":094A
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   1260
         Value           =   1  'Checked
         Width           =   3900
      End
      Begin VB.CheckBox CheckboxDisplaySmoothAnimations 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "ƽ������Ч��"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   3900
      End
      Begin VB.CheckBox CheckboxDisplayHideUnnecessaryInformation 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "���ز���Ҫ����Ϣ"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   8100
      End
      Begin VB.CheckBox CheckboxDisplayReduceContrast 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "���ͶԱȶ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   4410
         MouseIcon       =   "FormSettings.frx":0D40
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   420
         Width           =   3900
      End
      Begin VB.CheckBox CheckboxDisplayBlackOnWhite 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "�ڵװ���"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   3900
      End
   End
   Begin VB.Frame FrameDifficulty 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "�Ѷ�  (ռ 500 ��)"
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
      Height          =   4005
      Left            =   9135
      TabIndex        =   31
      Top             =   2625
      Width           =   8625
      Begin VB.CheckBox CheckboxDifficultyIncreaseDifficultyGradually 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "���������Ѷ�"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   8100
      End
      Begin VB.HScrollBar HScrollDifficultyMistakeAllowedAmount 
         Height          =   330
         Left            =   4935
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
         LargeChange     =   5
         Left            =   4935
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
         LargeChange     =   10
         Left            =   4935
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
         LargeChange     =   5
         Left            =   4935
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
         LargeChange     =   5
         Left            =   4935
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
         Caption         =   "����ʧ�������"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   3015
      End
      Begin VB.Label LabelDifficultyInterval 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ�����ش��ֱ����һ�����ֳ��֣���"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   4065
      End
      Begin VB.Label LabelDifficultyReachNormalDifficultyAt 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "��ɵ�ָ����Ϸ����ʱ���ִ������Ѷȣ�"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   4380
      End
      Begin VB.Label LabelDifficultyInitialDifficulty 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ�Ѷȣ�"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   2700
      End
      Begin VB.Label LabelDifficultyNormalDifficulty 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ѷȣ�"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   3015
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
         Left            =   3360
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
         Left            =   3360
         TabIndex        =   37
         Top             =   1300
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
         Left            =   3360
         TabIndex        =   40
         Top             =   2140
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
         Left            =   3360
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
         Left            =   3360
         TabIndex        =   46
         Top             =   3310
         Width           =   1440
      End
   End
   Begin VB.Frame FrameCheating 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "����"
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
      Height          =   1485
      Left            =   210
      TabIndex        =   54
      Top             =   8925
      Width           =   8625
      Begin VB.CheckBox CheckboxCheatingShowCorrectAnswer 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "ָʾ��ȷѡ��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   315
         MouseIcon       =   "FormSettings.frx":17D0
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   840
         Value           =   1  'Checked
         Width           =   7995
      End
      Begin VB.CheckBox CheckboxCheatingSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "�������׹���"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   8100
      End
   End
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "���"
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
      Left            =   16275
      MouseIcon       =   "FormSettings.frx":1A74
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
   Begin VB.Frame FrameGameMode 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "��Ϸģʽ  (ռ 250 ��)"
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
      Height          =   2430
      Left            =   210
      TabIndex        =   22
      Top             =   4200
      Width           =   8625
      Begin VB.OptionButton RadioboxGameModeKana 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "�����趨��Χ�ڵ��������֣���ѡ���Ҵ�Թ��������ֺ�ʤ����"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   8100
      End
      Begin VB.OptionButton RadioboxGameModeTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "�̶�һ��ʱ�䣨��������ʱ�伴ʤ����"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   8100
      End
      Begin VB.HScrollBar HScrollGameModeSpecifiedTime 
         Height          =   330
         LargeChange     =   5
         Left            =   4935
         Max             =   30
         Min             =   1
         MouseIcon       =   "FormSettings.frx":1E6A
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1750
         Value           =   5
         Width           =   3375
      End
      Begin VB.HScrollBar HScrollGameModeRepeatedTimes 
         Height          =   330
         Left            =   4935
         Max             =   10
         Min             =   1
         MouseIcon       =   "FormSettings.frx":1FBC
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   910
         Value           =   1
         Width           =   3375
      End
      Begin VB.Label LabelGameModeSpecifiedTime 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "�趨ʱ�䣺"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   2700
      End
      Begin VB.Label LabelGameModeRepeatedTimes 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "��һ���ֵĳ��ִ�����"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   3435
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
         Left            =   4095
         TabIndex        =   26
         Top             =   890
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
         Left            =   3360
         TabIndex        =   29
         Top             =   1725
         Width           =   1440
      End
   End
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   17640
      Top             =   10395
   End
   Begin VB.Label LabelSettingsTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "�趨"
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
      Height          =   10725
      Left            =   0
      Top             =   0
      Width           =   17970
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
        FormDifficultyIndexHelp.windowanimationtargetleft = (Screen.Width / 2) - (17970 / 2)
        FormDifficultyIndexHelp.windowanimationtargettop = (Screen.Height / 2) - (10725 / 2)
        FormDifficultyIndexHelp.windowanimationtargetwidth = 17970
        FormDifficultyIndexHelp.windowanimationtargetheight = 10725
        FormDifficultyIndexHelp.Show
    End Sub
    Public Sub CmdFontsApply_Click()
        FormMainWindow.LabelKanaDashboard.Font = TextboxFontsJpnFont.Text
        FormMainWindow.CmdOption1.Font = TextboxFontsEngFont.Text
        FormMainWindow.CmdOption2.Font = TextboxFontsEngFont.Text
        FormMainWindow.CmdOption3.Font = TextboxFontsEngFont.Text
        MsgBox "�����Ѹ�����", vbInformation + vbOKOnly + vbDefaultButton1, "������������Ϸ"
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