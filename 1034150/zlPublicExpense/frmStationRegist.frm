VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmStationRegist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医生站挂号"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   Icon            =   "frmStationRegist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdNewPati 
      Height          =   345
      Left            =   2940
      Picture         =   "frmStationRegist.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "新增病人"
      Top             =   600
      Width           =   350
   End
   Begin VB.PictureBox picPayMoney 
      BackColor       =   &H80000005&
      Height          =   420
      Left            =   6300
      ScaleHeight     =   360
      ScaleWidth      =   1575
      TabIndex        =   36
      Top             =   4942
      Width           =   1635
      Begin VB.Label lblPayMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   780
         TabIndex        =   37
         Top             =   15
         Width           =   720
      End
   End
   Begin VB.PictureBox picInfo 
      Height          =   2925
      Left            =   15
      ScaleHeight     =   2865
      ScaleWidth      =   7845
      TabIndex        =   30
      Top             =   1950
      Width           =   7905
      Begin VB.CheckBox chkBook 
         Caption         =   "购买病历"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6435
         TabIndex        =   8
         Top             =   2543
         Width           =   1485
      End
      Begin VB.ComboBox cboDoctor 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   525
         Width           =   3390
      End
      Begin VB.ComboBox cboAppointStyle 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   525
         Width           =   2355
      End
      Begin VB.ComboBox cboArrangeNo 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   45
         Width           =   3390
      End
      Begin VB.ComboBox cboRemark 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         TabIndex        =   7
         Top             =   2490
         Width           =   5430
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
         Height          =   1440
         Left            =   75
         TabIndex        =   31
         Top             =   975
         Width           =   7770
         _cx             =   13705
         _cy             =   2540
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStationRegist.frx":0B14
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   585
         Width           =   480
      End
      Begin VB.Label lblAppointStyle 
         AutoSize        =   -1  'True
         Caption         =   "预约方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4365
         TabIndex        =   35
         Top             =   585
         Width           =   960
      End
      Begin VB.Label lblArrangeNO 
         AutoSize        =   -1  'True
         Caption         =   "号别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   105
         Width           =   480
      End
      Begin VB.Label lblLimit 
         AutoSize        =   -1  'True
         Caption         =   "限号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4860
         TabIndex        =   33
         Top             =   105
         Width           =   600
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "备注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   32
         Top             =   2550
         Width           =   480
      End
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H80000005&
      Height          =   420
      Left            =   795
      ScaleHeight     =   360
      ScaleWidth      =   1575
      TabIndex        =   28
      Top             =   4950
      Width           =   1635
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   825
         TabIndex        =   29
         Top             =   15
         Width           =   720
      End
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   3
      Left            =   -45
      TabIndex        =   24
      Top             =   5430
      Width           =   11000
   End
   Begin VB.ComboBox cboPayMode 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4335
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4935
      Width           =   1950
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   2
      Left            =   -30
      TabIndex        =   19
      Top             =   1440
      Width           =   11000
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   1
      Left            =   -30
      TabIndex        =   18
      Top             =   480
      Width           =   11000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6180
      TabIndex        =   11
      Top             =   5520
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4800
      TabIndex        =   10
      Top             =   5520
      Width           =   1300
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   705
      TabIndex        =   12
      Top             =   5520
      Width           =   1300
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   0
      Left            =   -60
      TabIndex        =   16
      Top             =   960
      Width           =   11000
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   360
      Left            =   705
      TabIndex        =   15
      Top             =   600
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   635
      Appearance      =   2
      IDKindStr       =   "姓|姓名或就诊卡|0|0|0|0|0|;医|医保号|0|0|0|0|0|;身|身份证号|1|0|0|0|0|;门|门诊号|0|0|0|0|0|"
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "宋体"
      IDKind          =   -1
      DefaultCardType =   "0"
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1290
      TabIndex        =   1
      Top             =   600
      Width           =   1650
   End
   Begin VB.ComboBox cboNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6060
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "↓"
      Height          =   345
      Left            =   7605
      TabIndex        =   27
      Top             =   1568
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   360
      Left            =   5520
      TabIndex        =   3
      Top             =   1560
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92930050
      CurrentDate     =   42121
   End
   Begin VB.PictureBox picRoom 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5505
      ScaleHeight     =   300
      ScaleWidth      =   2325
      TabIndex        =   43
      Top             =   1560
      Width           =   2385
      Begin VB.Label lblRoomName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   45
         TabIndex        =   44
         Top             =   15
         Width           =   120
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   360
      Left            =   675
      TabIndex        =   2
      Top             =   1560
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92930049
      CurrentDate     =   42121
   End
   Begin VB.PictureBox picDept 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   675
      ScaleHeight     =   300
      ScaleWidth      =   3330
      TabIndex        =   41
      Top             =   1560
      Width           =   3390
      Begin VB.Label lblDeptName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   45
         TabIndex        =   42
         Top             =   15
         Width           =   120
      End
   End
   Begin VB.Label lbl急 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "急"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   135
      TabIndex        =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblPayMode 
      AutoSize        =   -1  'True
      Caption         =   "支付方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2970
      TabIndex        =   23
      Top             =   4995
      Width           =   1320
   End
   Begin VB.Label lblSum 
      AutoSize        =   -1  'True
      Caption         =   "合计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   22
      Top             =   4995
      Width           =   660
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "门诊预交余额:0.00     "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3615
      TabIndex        =   17
      Top             =   645
      Width           =   2880
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      Caption         =   "病人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   14
      Top             =   645
      Width           =   480
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "单据号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5310
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "性别:     年龄:       门诊号:              费别: "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   135
      TabIndex        =   38
      Top             =   1110
      Width           =   5880
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   25
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4875
      TabIndex        =   26
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      Caption         =   "诊室"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4875
      TabIndex        =   20
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "科室"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   21
      Top             =   1620
      Width           =   480
   End
End
Attribute VB_Name = "frmStationRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Long
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset, mblnStartFactUseType As Boolean
Private mblnCard As Boolean, mintSysAppLimit As Integer
Private mfrmPatiInfo As frmPatiInfo
Private mstrYBPati As String, mlng挂号ID As Long, mlng领用ID As Long
Private mblnOlnyBJYB As Boolean, mblnSharedInvoice As Boolean
Private mstr险类 As String, mblnAppointment As Boolean, mblnChangeFeeType As Boolean
Private mstrAge As String, mstrFeeType As String, mstrGender As String, mstrClinic As String
Private mstrPassWord As String, mblnUnload As Boolean, mstrInsure As String
Private mlngDept As Long
Private Const SNCOLS = 10
Private Const SnArgCols = 7
Public mlngNewPatiID As Long
Private mrsPlan As ADODB.Recordset, mblnInit As Boolean
Private mrsSNState As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsItems As ADODB.Recordset
Private mrs时间段 As ADODB.Recordset
Private mrsInComes As ADODB.Recordset
Private mcolCardPayMode As Collection
Private mcolArrangeNo As Collection
Private mlng病人ID As Long, mintIDKind As Integer
Private mcur个帐余额 As Currency
Private mblnOK As Boolean, mstrCardPass As String
Private mstrNO As String, mlngSN As Long
Private mintInsure As Integer
Private mdatLast As Date, mblnNewPati As Boolean
Private mblnChangeByCode As Boolean
Private mstrCardNO As String
Private mcur个帐透支 As Currency
Private Enum EM_REGISTFEE_MODE  '挂号费用收取方式
        EM_RG_现收 = 0
        EM_RG_划价 = 1
        EM_RG_记帐 = 2
End Enum
Private Enum EM_PATI_CHARGE_MODE    '病人收费模式
    EM_先结算后诊疗 = 0
    EM_先诊疗后结算 = 1
End Enum
Private mRegistFeeMode As EM_REGISTFEE_MODE '挂号费用收取方式
Private mPatiChargeMode As EM_PATI_CHARGE_MODE    '病人收费模式

Private Type TYPE_MedicarePAR
    医保接口打印票据 As Boolean
    使用个人帐户   As Boolean  'support挂号使用个人帐户
    连续挂号  As Boolean    'support连续挂号
    不收病历费 As Boolean   'support挂号不收取病历费
End Type
Private MCPAR As TYPE_MedicarePAR

Private Enum ViewMode
     V_普通号
     v_专家号
     v_专家号分时段
     V_普通号分时段
End Enum
Private mViewMode As ViewMode

Private Type ty_ModulePara
    bln姓名模糊查找 As Boolean
    lng姓名查找天数 As Long
    bln默认购买病历 As Boolean
    bln默认输入摘要 As Boolean
    byt挂号模式 As Byte
    bln挂号必须刷卡 As Boolean
    bln优先使用预交 As Boolean
    bln住院病人挂号 As Boolean
    bln包含科室安排 As Boolean
    int挂号发票打印 As Integer
    int挂号凭条打印 As Integer
    int预约挂号打印 As Integer
    bln随机序号选择 As Boolean
    lng预约有效时间 As Long
    bln共用收费票据 As Boolean
    bln退号重用 As Boolean
    bln预约时收款 As Boolean
    bln消费验证 As Boolean
    bln输入医生 As Boolean
End Type

Private mty_Para As ty_ModulePara

Public Sub zlShowMe(ByVal frmMain As Object, ByVal lngModul As Long, ByVal strDeptIDs As String, _
                    ByVal blnAppointment As Boolean, ByVal lng病人ID As Long, ByRef strOutNO As String)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    mblnAppointment = blnAppointment
    mlngModul = lngModul
    mlng病人ID = lng病人ID
    If frmMain Is Nothing Then
        Me.Show
    Else
        Me.Show 1, frmMain
    End If
    If mblnOK = True Then
        strOutNO = mstrNO
        Unload Me
    End If
End Sub

Private Sub InitPara()
    Dim strValue As String
    On Error GoTo errH
    With mty_Para
        .bln姓名模糊查找 = Val(gobjDatabase.GetPara("姓名模糊查找", glngSys, 9000, "0")) = 1
        .lng姓名查找天数 = Val(gobjDatabase.GetPara("姓名查找天数", glngSys, 9000, 0))
        .bln默认购买病历 = Val(gobjDatabase.GetPara("默认购买病历", glngSys, 9000, "0")) = 1
        .bln默认输入摘要 = Val(gobjDatabase.GetPara("默认输入摘要", glngSys, 9000, "1")) = 1
        .byt挂号模式 = Val(gobjDatabase.GetPara("挂号模式", glngSys, 9000, "0"))
        .bln优先使用预交 = Val(gobjDatabase.GetPara("优先使用预交", glngSys, 9000, "0")) = 1
        .bln住院病人挂号 = Val(gobjDatabase.GetPara("允许住院病人挂号", glngSys, 9000, "0")) = 1
        .int挂号发票打印 = Val(gobjDatabase.GetPara("挂号发票打印方式", glngSys, 9000, "0"))
        .int挂号凭条打印 = Val(gobjDatabase.GetPara("挂号凭条打印方式", glngSys, 9000, "0"))
        .int预约挂号打印 = Val(gobjDatabase.GetPara("预约挂号单打印方式", glngSys, 9000, "0"))
        .bln随机序号选择 = Val(gobjDatabase.GetPara("随机序号选择", glngSys, 9000, "0")) = 1
        .bln共用收费票据 = Val(gobjDatabase.GetPara("挂号共用收费票据", glngSys, 1121)) = 1
        .bln退号重用 = Val(gobjDatabase.GetPara("已退序号允许挂号", glngSys, 1111)) = 1
        .bln预约时收款 = Val(gobjDatabase.GetPara("预约时收款", glngSys, 9000, "0")) = 1
        .bln包含科室安排 = Val(gobjDatabase.GetPara("包含科室安排", glngSys, 9000, "0")) = 1
        .bln挂号必须刷卡 = Val(gobjDatabase.GetPara("挂号必须刷卡", glngSys, 9000)) = 1
        .bln消费验证 = Val(gobjDatabase.GetPara(28, glngSys)) <> 0
        .bln输入医生 = Val(gobjDatabase.GetPara("输入医生", glngSys, 9000)) = 1
        If .bln默认输入摘要 Then
            cboRemark.TabStop = True
        Else
            cboRemark.TabStop = False
        End If
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_现收
        Else
            If .byt挂号模式 = 0 Then
                mRegistFeeMode = EM_RG_现收
            Else
                mRegistFeeMode = EM_RG_划价
            End If
        End If
    End With
    '刷卡要求输入密码
    mstrCardPass = gobjDatabase.GetPara(46, glngSys, , "0000000000")
    Call gobjControl.PicShowFlat(picInfo)
    '收费和挂号共用票据
    mblnSharedInvoice = gobjDatabase.GetPara("挂号共用收费票据", glngSys, 1121) = "1"
    '本地共用挂号批次ID
    If mblnSharedInvoice Then
        mlng挂号ID = Val(gobjDatabase.GetPara("共用收费票据批次", glngSys, 1121, ""))
    Else
        mlng挂号ID = Val(gobjDatabase.GetPara("共用挂号票据批次", glngSys, mlngModul, ""))
    End If
    mlngDept = Val(gobjDatabase.GetPara("接诊科室", glngSys, 1260, ""))
    If mlng挂号ID > 0 Then
        If Not ExistBill(mlng挂号ID, IIf(mblnSharedInvoice, 1, 4)) Then
            If mblnSharedInvoice Then
                gobjDatabase.SetPara "共用收费票据批次", "0", glngSys, 1121
            Else
                gobjDatabase.SetPara "共用挂号票据批次", "0", glngSys, mlngModul
            End If
            mlng挂号ID = 0
        End If
    End If
    '票号严格控制
    strValue = gobjDatabase.GetPara(24, glngSys, , "00000")
    gblnBill挂号 = (Mid(strValue, IIf(mblnSharedInvoice, 1, 4), 1) = "1")
    mintSysAppLimit = Val(gobjDatabase.GetPara("挂号允许预约天数", glngSys))
    If mblnSharedInvoice Then
        '挂号用门诊票据:42703
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function zlStartFactUseType(ByVal int票种 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了使用类别的
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-10 16:11:47
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select  1 as 存在 From 票据领用记录 where 票种=[1] and nvl(使用类别,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查票据是否启用了使用类别的", int票种)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zl_GetInvoiceUserType(ByVal lng病人ID As Long, ByVal lng主页Id As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的使用类别
    '返回:发票的使用类别
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select  Zl_Billclass([1],[2],[3]) as 使用类别 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取票据使用类别", lng病人ID, lng主页Id, intInsure)
    zl_GetInvoiceUserType = Nvl(rsTemp!使用类别)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function ExistBill(lngID As Long, bytKind As Byte) As Boolean
'功能：判断是否存在指定的票据领用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select ID From 票据领用记录 Where ID=[1] And 票种=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "检查领用ID", lngID, bytKind)
    ExistBill = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function RefreshFact(Optional ByRef strFact As String) As Boolean
'参数：blnNew=是否新单保存时调用,这时对于非严格控制的票据是保存当前号
    Dim strUseType As String
    If mblnStartFactUseType Then
        strUseType = zl_GetInvoiceUserType(Val(mrsInfo!病人ID), 0, mintInsure)
    End If
    If gblnBill挂号 Then
        mlng领用ID = CheckUsedBill(IIf(mblnSharedInvoice, 1, 4), IIf(mlng领用ID > 0, mlng领用ID, mlng挂号ID), , strUseType)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的挂号票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            strFact = "": Exit Function
        Else
            '严格：取下一个号码
            strFact = GetNextBill(mlng领用ID)
        End If
    Else
        If mblnSharedInvoice Then
            strFact = gobjDatabase.GetPara("当前收费票据号", glngSys, 1121)
        Else
            strFact = gobjDatabase.GetPara("当前挂号票据号", glngSys, 1111)
        End If
        strFact = IncStr(strFact)
        If mblnSharedInvoice Then
            gobjDatabase.SetPara "当前收费票据号", strFact, glngSys, 1121
        Else
            gobjDatabase.SetPara "当前挂号票据号", strFact, glngSys, 1111
        End If
    End If
    RefreshFact = True
End Function

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long, strTemp As String
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "姓|姓名|0;医|医保号|0;身|身份证号|0;门|门诊号|0", txtPatient)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If

    Call GetRegInFor(g私有模块, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function

Private Sub cboAppointStyle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboArrangeNo_Click()
    Call ReadLimit
    Call LoadDoctor
    Call LoadFeeItem(Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1)
    Call GetActiveView
    If mblnAppointment Then
        Select Case mViewMode
            Case V_普通号分时段, v_专家号分时段
                cmdTime.Visible = True
            Case Else
                cmdTime.Visible = False
        End Select
        Call InitRegTime
    Else
        cmdTime.Visible = False
    End If
    lblDeptName.Caption = Nvl(mrsPlan!科室)
End Sub

Private Sub InitRegTime()
    Dim dateCur As Date, strNO As String, strDay As String, strSQL As String
    Dim rsTmp As ADODB.Recordset, rsTime As ADODB.Recordset
    On Error GoTo errH
    strDay = zlGet当前星期几(dtpDate.Value)
    strSQL = "Select 时间段,开始时间,缺省时间 From 时间段 Where 号类 Is Null And 站点 Is Null"
    Set rsTime = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If IsNull(mrsPlan.Fields(strDay).Value) Then
        If Format(dtpDate.Value, "yyyy-mm-dd") = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") Then
            '当天不当班,取当前时间
            dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
        Else
            '未来不当班,取默认时间
            rsTime.Filter = "时间段='白天'"
            If rsTime.EOF Then
                dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
            Else
                If IsNull(rsTime!缺省时间) Then
                    dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
                Else
                    dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
                End If
            End If
        End If
    Else
        Select Case mViewMode
            Case V_普通号分时段, v_专家号分时段
            strSQL = "Select Distinct a.序号 As ID, To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
                    "From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
                    "Where a.安排id = b.Id And b.号码 = [1] And" & vbNewLine & _
                    " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
                    "      Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六'," & vbNewLine & _
                    "             Null) = a.星期(+) And Not Exists" & vbNewLine & _
                    " (Select Count(1)" & vbNewLine & _
                    "       From 挂号序号状态" & vbNewLine & _
                    "       Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like a.序号 || '__') Having" & vbNewLine & _
                    "        Count(1) - a.限制数量 >= 0) And Not Exists" & vbNewLine & _
                    " (Select 1" & vbNewLine & _
                    "       From 挂号安排计划 E" & vbNewLine & _
                    "       Where e.安排id = b.Id And e.审核时间 Is Not Null And" & vbNewLine & _
                    "             [2] Between Nvl(e.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                    "             Nvl(e.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')))"
            strSQL = strSQL & " Union " & _
                    "Select Distinct a.序号 As ID, To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
                    "From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C," & vbNewLine & _
                    "     (Select Max(a.生效时间) 生效" & vbNewLine & _
                    "       From 挂号安排计划 A, 挂号安排 B" & vbNewLine & _
                    "       Where a.安排id = b.Id And b.号码 = [1] And a.审核时间 Is Not Null And" & vbNewLine & _
                    "             [2] Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                    "             Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'))) D" & vbNewLine & _
                    "Where a.计划id = b.Id And b.安排id = c.Id And c.号码 = [1] And b.生效时间 = d.生效 And b.审核时间 Is Not Null And" & vbNewLine & _
                    " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
                    "      [2] Between Nvl(b.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                    "      Nvl(b.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And Not Exists" & vbNewLine & _
                    " (Select Count(1)" & vbNewLine & _
                    "       From 挂号序号状态" & vbNewLine & _
                    "       Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like a.序号 || '__') Having" & vbNewLine & _
                    "        Count(1) - a.限制数量 >= 0) And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5'," & vbNewLine & _
                    "                                           '周四', '6', '周五', '7', '周六', Null) = a.星期(+)" & vbNewLine & _
                    "Order By 开始时间"
        
            dateCur = Format(dtpDate, "yyyy-mm-dd")
            strNO = Get号别
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, dateCur)
            If Not rsTmp.EOF Then
                '时段当班有时段,取最小时段
                dtpTime.Value = Format(Nvl(rsTmp!开始时间), "hh:mm:ss")
            Else
                If Format(dtpDate.Value, "yyyy-mm-dd") = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") Then
                    dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
                Else
                    '时段当班无时段,取开始时间
                    rsTime.Filter = "时间段='" & Nvl(mrsPlan.Fields(strDay).Value) & "'"
                    If rsTime.EOF Then
                        dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
                    Else
                        If IsNull(rsTime!缺省时间) Then
                            dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
                        Else
                            dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
                        End If
                    End If
                End If
            End If
            Case Else
                If Format(dtpDate.Value, "yyyy-mm-dd") = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") Then
                    dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
                Else
                    '当班无时段,取开始时间
                    rsTime.Filter = "时间段='" & Nvl(mrsPlan.Fields(strDay).Value) & "'"
                    If rsTime.EOF Then
                        dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
                    Else
                        If IsNull(rsTime!缺省时间) Then
                            dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
                        Else
                            dtpTime.Value = Format(Nvl(rsTime!开始时间), "hh:mm:ss")
                        End If
                    End If
                End If
        End Select
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub GetAll医生()
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.姓名, Upper(a.简码) As 简码,b.部门id,a.编号" & _
            " From 人员表 a, 部门人员 b, 人员性质说明 c" & _
            " Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order By a.简码 Desc"
    Set mrsDoctor = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, "医生")
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub cboArrangeNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub LoadDoctor()
    With cboDoctor
        .Clear
        If Nvl(mrsPlan!医生) = "" Then
            If mty_Para.bln输入医生 Then
                mrsDoctor.Filter = "部门id=" & Val(Nvl(mrsPlan!科室ID))
                
                Do While Not mrsDoctor.EOF
                    .AddItem IIf(IsNull(mrsDoctor!简码), "", mrsDoctor!简码 & "-") & mrsDoctor!姓名
                    .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                    If Nvl(mrsDoctor!姓名) = UserInfo.姓名 Then .ListIndex = .NewIndex
                    mrsDoctor.MoveNext
                Loop
                If .ListIndex < 0 Then
                    .ListIndex = 0
                End If
                .Enabled = True
                lblDoctor.Enabled = True
            Else
                mrsDoctor.Filter = "姓名='" & UserInfo.姓名 & "'"
                If mrsDoctor.RecordCount <> 0 Then
                    .AddItem IIf(IsNull(mrsDoctor!简码), "", mrsDoctor!简码 & "-") & mrsDoctor!姓名
                    .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                    .ListIndex = 0
                End If
                .Enabled = False
                lblDoctor.Enabled = False
            End If
        Else
            mrsDoctor.Filter = "姓名='" & Nvl(mrsPlan!医生) & "'"
            If mrsDoctor.RecordCount <> 0 Then
                .AddItem IIf(IsNull(mrsDoctor!简码), "", mrsDoctor!简码 & "-") & mrsDoctor!姓名
                .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                .ListIndex = 0
            End If
            .Enabled = False
            lblDoctor.Enabled = False
        End If
    End With
End Sub

Private Sub cboPayMode_Click()
    If MCPAR.不收病历费 And cboPayMode.Text = mstrInsure Then
        chkBook.Enabled = False
        chkBook.Value = 0
    Else
        chkBook.Enabled = True
    End If
End Sub

Private Sub cboPayMode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboRemark_Change()
    cboRemark.Tag = ""
End Sub

Private Sub cboRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRemark.Tag <> "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(cboRemark.Text) = "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If SelectMemo(Trim(cboRemark.Text)) = False Then
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
End Sub

Private Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String

    If Val(gobjDatabase.GetPara("输入匹配")) = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Private Function SelectMemo(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择常用摘要
    '入参:strInput-输入串;为空时,表示全部
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If gobjCommFun.IsCharChinese(cboRemark.Text) Then
             strWhere = " And  名称 like [1] "
        ElseIf gobjCommFun.IsNumOrChar(cboRemark.Text) Then
             strWhere = " And (简码 like upper([1]) or 编码 like upper([1]))"
        End If
    End If
    
    strSQL = "" & _
     "   Select RowNum AS ID,编码,名称,简码  " & _
     "   From 常用挂号摘要 " & _
     "   Where 1=1 " & strWhere & _
     "   Order by 缺省标志"
     vRect = GetControlRect(cboRemark.hWnd)
     On Error GoTo Hd
     Set rsInfo = gobjDatabase.ShowSQLSelect(Me, strSQL, 0, "常用挂号摘要", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cboRemark.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "没有设置常用挂号摘要,请在字典管理中设置", vbOKOnly + vbInformation, gstrSysName
        End If
        gobjCommFun.PressKey vbKeyTab: Exit Function
     End If
     gobjControl.CboSetText Me.cboRemark, Nvl(rsInfo!名称)
     cboRemark.Tag = Nvl(rsInfo!名称)
     gobjCommFun.PressKey vbKeyTab
     SelectMemo = True
     Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then Resume
    gobjComlib.SaveErrLog
End Function

Private Sub chkBook_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cmdNewPati_Click()
    If Not mrsInfo Is Nothing Then
        If Val(Nvl(mrsInfo!病人ID)) <> 0 Then
            Call ViewPatiInfo
            Exit Sub
        End If
    End If
    Call CreateNewPati
End Sub

Private Sub ViewPatiInfo()
    '查看病人信息
    Dim bln复诊 As Boolean, lng科室id As Long
    On Error GoTo errH
    If mrsPlan Is Nothing Then
        lng科室id = 0
    Else
        If mrsPlan.RecordCount = 0 Then
            lng科室id = 0
        Else
            lng科室id = Val(Nvl(mrsPlan!科室ID))
        End If
    End If
    
    bln复诊 = Check复诊(Val(Nvl(mrsInfo!病人ID)), lng科室id)
    With mfrmPatiInfo
        Set .mfrmMain = Me
        .mbytFun = 0
        .mlng病人ID = Val(Nvl(mrsInfo!病人ID))
        .mbln复诊 = bln复诊
        .Show 1, Me
    End With
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CreateNewPati()
    '新建病人信息
    On Error GoTo errH

    With mfrmPatiInfo
        Set .mfrmMain = Me
        .mbytFun = 2
        .Show 1, Me
        If mlngNewPatiID <> 0 Then mblnNewPati = True
    End With
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdTime_Click()
    If InitTimePlan = False Then Exit Sub
    If mrs时间段.RecordCount <> 0 Then
        dtpTime.Value = Format(mrs时间段!开始时间, "hh:mm:ss")
    End If
End Sub

Private Sub dtpDate_Change()
    Call LoadRegPlans(False)
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub dtpTime_GotFocus()
    Call cmdTime_Click
End Sub

Private Sub dtpTime_Validate(Cancel As Boolean)
    If Format(dtpDate.Value, "YYYY-MM-DD") = Format(gobjDatabase.CurrentDate, "YYYY-MM-DD") Then
        If Format(dtpTime.Value, "hh:mm:ss") < Format(gobjDatabase.CurrentDate, "hh:mm:ss") Then
            MsgBox "预约时间不能小于当前时间!", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnload Then mblnUnload = False: Unload Me: Exit Sub
    If mblnInit And Not mrsInfo Is Nothing Then
        If cboArrangeNo.Enabled And cboArrangeNo.Visible Then cboArrangeNo.SetFocus
        If cboArrangeNo.ListCount = 1 Then gobjCommFun.PressKeyEx vbKeyTab
    End If
    mblnInit = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsInfo = Nothing
    If Not mobjIDCard Is Nothing Then
         Call mobjIDCard.SetEnabled(False)
         Set mobjIDCard = Nothing
     End If
     If Not mobjICCard Is Nothing Then
         Call mobjICCard.SetEnabled(False)
         Set mobjICCard = Nothing
     End If
     mintIDKind = IDKind.IDKind
     Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strExpand As String
    Dim strOutCardNO As String, strOutPatiInforXML As String
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        '系统IC卡
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(objCard, txtPatient.Text, True)
            End If
        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    
    If txtPatient.Text <> "" Then
        Call GetPatient(objCard, txtPatient.Text, True)
    End If
End Sub

Private Sub txtRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub chkBook_Click()
    If mrsPlan Is Nothing Then Exit Sub
    Call LoadFeeItem(Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1)
End Sub

Private Sub cmdCancel_Click()
    If txtPatient.Text <> "" Then
        If MsgBox("是否清空当前病人信息？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ClearPatient
        End If
        Exit Sub
    End If
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp "zl9RegEvent", Me.hWnd, "frmRegistEdit"
    Exit Sub
End Sub

Private Function CheckBrushCard(ByVal dblMoney As Double, ByVal lng医疗卡类别ID As Long, ByVal bln消费卡 As Boolean, _
                                ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str年龄 As String
    On Error GoTo errHandle
    '68991
    If mRegistFeeMode <> EM_RG_现收 Then CheckBrushCard = True: Exit Function
    If dblMoney = 0 Then
        CheckBrushCard = True: Exit Function
    End If
    If Not (cboPayMode.Visible And cboPayMode.Enabled) Then
        CheckBrushCard = True: Exit Function
    End If
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then
        CheckBrushCard = True: Exit Function
    End If
    If lng医疗卡类别ID = 0 Then
        MsgBox cboPayMode.Text & "异常,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "使用" & cboPayMode.Text & "支付必须先初始化接口部件！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call zlGetClassMoney(rsMoney, rsItems, rsIncomes)
    
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String) As Boolean
    str年龄 = Trim(mstrAge)

   If gobjSquare.objSquareCard.zlBrushCard(Me, glngModul, rsMoney, lng医疗卡类别ID, bln消费卡, _
    txtPatient.Text, NeedName(mstrGender), str年龄, dblMoney, mstrCardNO, mstrPassWord, _
    False, True, False, True, Nothing, False, True) = False Then Exit Function
    
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, glngModul, lng医疗卡类别ID, _
        bln消费卡, mstrCardNO, dblMoney, "", "") = False Then Exit Function

    CheckBrushCard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL  As String
    
    Err = 0: On Error GoTo Errhand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic

        rsItems.Filter = 0
        If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
        Do While Not rsItems.EOF
            rsIncomes.Filter = "项目ID=" & rsItems!项目ID
            rsMoney.Filter = "收费类别='" & Nvl(rsItems!类别, "无") & "'"
            If rsMoney.EOF Then
                .AddNew
            Else
                rsMoney.Filter = 0
            End If
            !收费类别 = Nvl(rsItems!类别, "无")
            !金额 = Val(Nvl(!金额)) + Val(Nvl(rsIncomes!实收))
            .Update
            rsItems.MoveNext
        Loop
    End With
    rsMoney.Filter = 0
    zlGetClassMoney = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim blnSlipPrint As Boolean, blnInvoicePrint As Boolean, int价格父号 As Integer, blnBalance As Boolean
    Dim k As Integer, i As Integer, j As Integer, strNO As String, strFactNO As String
    Dim cllPro As New Collection, strSQL As String, str登记时间 As String, str发生时间 As String
    Dim cur预交 As Currency, cur个帐 As Currency, cur现金 As Currency, str划价NO As String
    Dim lngSN As Long, lng挂号科室ID As Long, lng结帐ID As Long, byt复诊 As Byte
    Dim lng医疗卡类别ID As Long, bln消费卡 As Boolean, blnNoDoc As Boolean, strBalanceStyle As String
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset, cllProAfter As New Collection
    Dim blnTrans As Boolean, blnNotCommit As Boolean, strAdvance As String, lng病人ID As Long
    Dim lng医生ID As Long, blnOneCard As Boolean, rsTmp As ADODB.Recordset
    Dim cllCardPro As Collection, cllTheeSwap As Collection, strNotValiedNos As String
    Dim strDay As String, blnAppointPrint As Boolean, str付款方式 As String
    Dim rs付款方式 As ADODB.Recordset, str医生 As String, blnAdd As Boolean
    
    If CheckValied = False Then Exit Sub
    
    strSQL = "Select 编号,名称,医院编码,结算方式 From 一卡通目录 Where 启用 = 1 And 结算方式 = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, cboPayMode.Text)
    blnOneCard = rsTmp.RecordCount <> 0
    
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        blnSlipPrint = False
    Else
        Select Case Val(mty_Para.int挂号凭条打印)
            Case 0    '不打印
                blnSlipPrint = False
            Case 1    '自动打印
                If InStr(gstrPrivs, ";病人挂号凭条;") > 0 Then
                    blnSlipPrint = True
                Else
                    blnSlipPrint = False
                    MsgBox "你没有挂号凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
                End If
            Case 2    '选择打印
                If MsgBox("要打印挂号凭条吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If InStr(gstrPrivs, ";病人挂号凭条;") > 0 Then
                        blnSlipPrint = True
                    Else
                        blnSlipPrint = False
                        MsgBox "你没有挂号凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
                    End If
                Else
                    blnSlipPrint = False
                End If
        End Select
    End If
    
    If mRegistFeeMode = EM_RG_划价 Or mRegistFeeMode = EM_RG_记帐 Or (mblnAppointment And mty_Para.bln预约时收款 = False) Then
        blnInvoicePrint = False
    Else
        If Not (mintInsure <> 0 And MCPAR.医保接口打印票据) Then
            Select Case Val(mty_Para.int挂号发票打印)
                Case 0    '不打印
                    blnInvoicePrint = False
                Case 1    '自动打印
                    If InStr(gstrPrivs, ";挂号发票打印;") > 0 Then
                        blnInvoicePrint = True
                    Else
                        blnInvoicePrint = False
                        MsgBox "你没有挂号发票打印的权限，请联系管理员！", vbInformation, gstrSysName
                    End If
                Case 2    '选择打印
                    If MsgBox("要打印挂号发票吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        If InStr(gstrPrivs, ";挂号发票打印;") > 0 Then
                            blnInvoicePrint = True
                        Else
                            blnInvoicePrint = False
                            MsgBox "你没有挂号发票打印的权限，请联系管理员！", vbInformation, gstrSysName
                        End If
                    Else
                        blnInvoicePrint = False
                    End If
            End Select
        End If
    End If
    
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        Select Case Val(mty_Para.int预约挂号打印)
            Case 0
                blnAppointPrint = False
            Case 1
                If InStr(gstrPrivs, ";预约挂号单;") > 0 Then
                    blnAppointPrint = True
                Else
                    blnAppointPrint = False
                    MsgBox "你没有预约挂号单打印的权限，请联系管理员！", vbInformation, gstrSysName
                End If
            Case 2
                If InStr(gstrPrivs, ";预约挂号单;") > 0 Then
                    If MsgBox("要打印预约挂号单吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        blnAppointPrint = True
                    Else
                        blnAppointPrint = False
                    End If
                Else
                    MsgBox "你没有预约挂号单打印的权限，请联系管理员！", vbInformation, gstrSysName
                    blnAppointPrint = False
                End If
        End Select
    Else
        blnAppointPrint = False
    End If
    
    If blnInvoicePrint Or (mintInsure <> 0 And MCPAR.医保接口打印票据) Then
        If RefreshFact(strFactNO) = False Then Exit Sub
    End If
    
    If mblnAppointment Then
        If mRegistFeeMode = EM_RG_记帐 And mty_Para.bln预约时收款 Then
            MsgBox "不支持先诊疗后结算病人的预约收款挂号！", vbInformation, gstrSysName
            Exit Sub
        End If
        If mty_Para.bln预约时收款 Then
            If Not mRegistFeeMode = EM_RG_划价 Then
                If cboPayMode.Text = "预交金" Then
                    cur预交 = Val(lblTotal.Caption)
                Else
                    If cboPayMode.Text = mstrInsure Then
                        cur个帐 = Val(lblTotal.Caption)
                    Else
                        blnBalance = True
                        cur现金 = Val(lblTotal.Caption)
                    End If
                End If
            End If
        Else
            blnBalance = False
        End If
    Else
        If Not mRegistFeeMode = EM_RG_划价 Then
            If cboPayMode.Text = "预交金" Then
                cur预交 = Val(lblTotal.Caption)
            Else
                If cboPayMode.Text = "个人帐户" Then
                    cur个帐 = Val(lblTotal.Caption)
                Else
                    blnBalance = True
                    cur现金 = Val(lblTotal.Caption)
                End If
            End If
        End If
    End If
    
    If frmPatiInfo.SaveAfterArrList(mblnNewPati, lng病人ID) = False Then
        MsgBox "保存病人信息失败，请检查！", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnNewPati = False
    mlngNewPatiID = 0
    txtPatient.Text = "-" & lng病人ID
    GetPatient IDKind.GetCurCard, txtPatient.Text, False
    
    If Val(cur预交) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!病人ID), Val(cur预交), mlngModul, 1, , mty_Para.bln消费验证) Then Exit Sub
    End If
    
    ReadRegistPrice Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1, False, mstrFeeType, rsItems, rsIncomes
    
    If mblnAppointment = False Or (mblnAppointment = True And mty_Para.bln预约时收款) Then
        If zlIsAllowPatiChargeFeeMode(ZVal(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!结算模式))) = False Then Exit Sub
    End If

    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lng医疗卡类别ID = mcolCardPayMode.Item(i)(3)
                bln消费卡 = Val(mcolCardPayMode.Item(i)(5)) = 1
                strBalanceStyle = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur现金), lng医疗卡类别ID, bln消费卡, rsItems, rsIncomes) = False Then Exit Sub
    End If
    
    str登记时间 = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
    
    If mblnAppointment Then
        strDay = zlGet当前星期几(dtpDate.Value)
    Else
        strDay = zlGet当前星期几
    End If
    
    '获取发生时间
    blnAdd = False
    If mblnAppointment Then
        mlngSN = 0
        str发生时间 = "To_Date('" & Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
        If mViewMode = v_专家号分时段 Then
            If Val(Nvl(mrsPlan!计划ID)) <> 0 Then
                strSQL = "Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约, 0 As 已约数, Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数," & vbNewLine & _
                        "       Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段" & vbNewLine & _
                        "From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码," & vbNewLine & _
                        "              To_Date(To_Char(" & str发生时间 & ", 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间," & vbNewLine & _
                        "              To_Date(To_Char(" & str发生时间 & ", 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间," & vbNewLine & _
                        "              Sd.限制数量, Sd.是否预约" & vbNewLine & _
                        "       From 挂号安排计划 Jh, 挂号计划时段 Sd" & vbNewLine & _
                        "       Where Jh.Id = Sd.计划id And Jh.Id = [1] And" & vbNewLine & _
                        "             Sd.星期 =" & vbNewLine & _
                        "             Decode(To_Char(" & str发生时间 & ", 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)) Jh," & vbNewLine & _
                        "     挂号序号状态 Zt" & vbNewLine & _
                        "Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Jh.开始时间 = " & str发生时间 & " And Zt.序号(+) = Jh.序号 And Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1" & vbNewLine & _
                        "Order By 序号"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!计划ID)))
                If rsTmp.RecordCount <> 0 Then
                    mlngSN = Val(Nvl(rsTmp!序号))
                Else
                    strSQL = "Select Max(序号) As 序号 From 挂号序号状态 Where 号码 = [1] And Trunc(日期) = Trunc(" & str发生时间 & ")"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(mrsPlan!号别))
                    If rsTmp.RecordCount <> 0 Then mlngSN = Val(Nvl(rsTmp!序号))
                    strSQL = "Select Max(序号) As 序号 From 挂号计划时段 Where 计划ID = [1] And 星期 = Decode(To_Char(" & str发生时间 & ", 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!计划ID)))
                    If mlngSN = 0 Then
                        If rsTmp.RecordCount <> 0 Then mlngSN = Val(Nvl(rsTmp!序号))
                    Else
                        If Val(Nvl(rsTmp!序号)) > mlngSN Then mlngSN = Val(Nvl(rsTmp!序号))
                    End If
                    mlngSN = mlngSN + 1
                    blnAdd = True
                End If
            Else
                strSQL = "Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约, 0 As 已约数, Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数," & vbNewLine & _
                        "       Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段" & vbNewLine & _
                        "From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码," & vbNewLine & _
                        "              To_Date(To_Char(" & str发生时间 & ", 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间," & vbNewLine & _
                        "              To_Date(To_Char(" & str发生时间 & ", 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间," & vbNewLine & _
                        "              Sd.限制数量, Sd.是否预约" & vbNewLine & _
                        "       From 挂号安排 Ap, 挂号安排时段 Sd" & vbNewLine & _
                        "       Where Ap.Id = Sd.安排id And Ap.Id = [1] And" & vbNewLine & _
                        "             Sd.星期 = Decode(To_Char(" & str发生时间 & ", 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7'," & vbNewLine & _
                        "                            '周六', Null)) Ap, 挂号序号状态 Zt" & vbNewLine & _
                        "Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Ap.开始时间 = " & str发生时间 & " And Zt.序号(+) = Ap.序号 And Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1" & vbNewLine & _
                        "Order By 序号"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!ID)))
                If rsTmp.RecordCount <> 0 Then
                    mlngSN = Val(Nvl(rsTmp!序号))
                Else
                    strSQL = "Select Max(序号) As 序号 From 挂号序号状态 Where 号码 = [1] And Trunc(日期) = Trunc(" & str发生时间 & ")"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(mrsPlan!号别))
                    If rsTmp.RecordCount <> 0 Then mlngSN = Val(Nvl(rsTmp!序号))
                    strSQL = "Select Max(序号) As 序号 From 挂号安排时段 Where 安排ID = [1] And 星期 = Decode(To_Char(" & str发生时间 & ", 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!ID)))
                    If mlngSN = 0 Then
                        If rsTmp.RecordCount <> 0 Then mlngSN = Val(Nvl(rsTmp!序号))
                    Else
                        If Val(Nvl(rsTmp!序号)) > mlngSN Then mlngSN = Val(Nvl(rsTmp!序号))
                    End If
                    mlngSN = mlngSN + 1
                    blnAdd = True
                End If
            End If
        End If
        If IsNull(mrsPlan.Fields(strDay).Value) Then blnAdd = True
    Else
        Select Case mViewMode
            Case V_普通号
                str发生时间 = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                If IsNull(mrsPlan.Fields(strDay).Value) Then blnAdd = True
            Case V_普通号分时段
                str发生时间 = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                If IsNull(mrsPlan.Fields(strDay).Value) Then blnAdd = True
            Case v_专家号
                str发生时间 = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                If IsNull(mrsPlan.Fields(strDay).Value) Then blnAdd = True
            Case v_专家号分时段
                If IsNull(mrsPlan.Fields(strDay).Value) Then
                    str发生时间 = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                    blnAdd = True
                Else
                    strSQL = "Select 1" & vbNewLine & _
                            "From 时间段" & vbNewLine & _
                            "Where 号类 Is Null And 站点 Is Null And (('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') Between" & vbNewLine & _
                            "      Decode(Sign(开始时间 - 终止时间), 1, '3000-01-09 ' || To_Char(Nvl(提前时间, 开始时间), 'HH24:MI:SS')," & vbNewLine & _
                            "               '3000-01-10 ' || To_Char(Nvl(提前时间, 开始时间), 'HH24:MI:SS')) And" & vbNewLine & _
                            "      '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')) Or" & vbNewLine & _
                            "      ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') Between '3000-01-10 ' || To_Char(Nvl(提前时间, 开始时间), 'HH24:MI:SS') And" & vbNewLine & _
                            "      Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS')," & vbNewLine & _
                            "               '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')))) And 时间段 = [1]"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mrsPlan.Fields(strDay).Value)
                    '不当班
                    If rsTmp.RecordCount = 0 Then
                        str发生时间 = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                        blnAdd = True
                    Else
                        '取最小可用时间段
                        If Val(Nvl(mrsPlan!计划ID)) <> 0 Then
                            strSQL = "Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约, 0 As 已约数, Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数," & vbNewLine & _
                                    "       Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段" & vbNewLine & _
                                    "From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码," & vbNewLine & _
                                    "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间," & vbNewLine & _
                                    "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间," & vbNewLine & _
                                    "              Sd.限制数量, Sd.是否预约" & vbNewLine & _
                                    "       From 挂号安排计划 Jh, 挂号计划时段 Sd" & vbNewLine & _
                                    "       Where Jh.Id = Sd.计划id And Jh.Id = [1] And" & vbNewLine & _
                                    "             Sd.星期 =" & vbNewLine & _
                                    "             Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)) Jh," & vbNewLine & _
                                    "     挂号序号状态 Zt" & vbNewLine & _
                                    "Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1" & vbNewLine & _
                                    "Order By 序号"
                            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!计划ID)))
                            If rsTmp.RecordCount <> 0 Then
                                mlngSN = Val(Nvl(rsTmp!序号))
                                str发生时间 = "To_Date('" & Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") & " " & Format(Nvl(rsTmp!开始时间), "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                            Else
                                str发生时间 = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                            End If
                        Else
                            strSQL = "Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约, 0 As 已约数, Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数," & vbNewLine & _
                                    "       Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段" & vbNewLine & _
                                    "From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码," & vbNewLine & _
                                    "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间," & vbNewLine & _
                                    "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间," & vbNewLine & _
                                    "              Sd.限制数量, Sd.是否预约" & vbNewLine & _
                                    "       From 挂号安排 Ap, 挂号安排时段 Sd" & vbNewLine & _
                                    "       Where Ap.Id = Sd.安排id And Ap.Id = [1] And" & vbNewLine & _
                                    "             Sd.星期 = Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7'," & vbNewLine & _
                                    "                            '周六', Null)) Ap, 挂号序号状态 Zt" & vbNewLine & _
                                    "Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1" & vbNewLine & _
                                    "Order By 序号"
    
                            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsPlan!ID)))
                            If rsTmp.RecordCount <> 0 Then
                                mlngSN = Val(Nvl(rsTmp!序号))
                                str发生时间 = "To_Date('" & Format(gobjDatabase.CurrentDate, "yyyy-mm-dd") & " " & Format(Nvl(rsTmp!开始时间), "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                            Else
                                str发生时间 = "To_Date('" & gobjDatabase.CurrentDate & "','yyyy-mm-dd hh24:mi:ss')"
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    
    lng挂号科室ID = Val(Nvl(mrsPlan!科室ID))
    lng结帐ID = gobjDatabase.GetNextId("病人结帐记录")
    byt复诊 = IIf(Check复诊(Val(mrsInfo!病人ID), lng挂号科室ID), 1, 0)
    
    '票据处理
    If mRegistFeeMode = EM_RG_划价 Then
        str划价NO = gobjDatabase.GetNextNo(13)
    End If
    lngSN = mlngSN
    strNO = gobjDatabase.GetNextNo(12)
    
    rsItems.Filter = ""
    str医生 = NeedName(cboDoctor.Text)
    If cboDoctor.ListCount = 0 Then
        lng医生ID = 0
    Else
        lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    strSQL = "Select 编码 From 医疗付款方式 Where 名称 = [1]"
    Set rs付款方式 = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, Nvl(mrsInfo!医疗付款方式))
    If rs付款方式.RecordCount <> 0 Then
        str付款方式 = Nvl(rs付款方式!编码)
    Else
        strSQL = "Select 编码 From 医疗付款方式 Where 缺省标志 = 1"
        Set rs付款方式 = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName)
        If rs付款方式.RecordCount <> 0 Then
            str付款方式 = Nvl(rs付款方式!编码)
        End If
    End If
    
    k = 1: rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        int价格父号 = k
        rsIncomes.Filter = "项目ID=" & rsItems!项目ID
        For j = 1 To rsIncomes.RecordCount
            strSQL = _
            "zl_病人挂号记录_INSERT(" & ZVal(Nvl(mrsInfo!病人ID)) & "," & IIf(mstrClinic = "", "NULL", mstrClinic) & ",'" & txtPatient.Text & "','" & mstrGender & "'," & _
                     "'" & mstrAge & "','" & str付款方式 & "','" & mstrFeeType & "','" & strNO & "'," & _
                     "'" & IIf(blnInvoicePrint = False, "", strFactNO) & "'," & k & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & IIf(rsItems!性质 = 2, 1, "NULL") & "," & _
                     "'" & rsItems!类别 & "'," & rsItems!项目ID & "," & rsItems!数次 & "," & rsIncomes!单价 & "," & _
                     rsIncomes!收入项目ID & ",'" & rsIncomes!收据费目 & "','" & IIf(blnBalance, IIf(strBalanceStyle = "", cboPayMode.Text, strBalanceStyle), "") & "'," & _
                     IIf(mRegistFeeMode = EM_RG_划价, 0, rsIncomes!应收) & "," & IIf(mRegistFeeMode = EM_RG_划价, 0, rsIncomes!实收) & "," & _
                     lng挂号科室ID & "," & UserInfo.部门ID & "," & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                     str发生时间 & "," & str登记时间 & "," & _
                     "'" & str医生 & "'," & ZVal(lng医生ID) & "," & IIf(rsItems!性质 = 3, 1, IIf(rsItems!性质 = 4, 2, 0)) & "," & IIf(lbl急.Visible, 1, 0) & "," & _
                     "'" & Get号别 & "','" & IIf(str医生 = UserInfo.姓名, lblRoomName.Caption, "") & "'," & ZVal(lng结帐ID) & "," & IIf(blnInvoicePrint = False, "NULL", ZVal(mlng领用ID)) & "," & _
                     ZVal(IIf(k = 1, cur预交, 0)) & "," & ZVal(IIf(k = 1, cur现金, 0)) & "," & _
                     ZVal(IIf(k = 1, cur个帐, 0)) & "," & ZVal(Nvl(rsItems!保险大类ID, 0)) & "," & _
                     ZVal(Nvl(rsItems!保险项目否, 0)) & "," & ZVal(Nvl(rsIncomes!统筹金额, 0)) & "," & _
                     "'" & IIf(str划价NO <> "", "划价:" & str划价NO, Me.cboRemark.Text) & "'," & IIf(mblnAppointment, IIf(mty_Para.bln预约时收款, 0, 1), 0) & "," & IIf(mty_Para.bln共用收费票据, 1, 0) & ",'" & rsItems!保险编码 & "'," & byt复诊 & "," & ZVal(lngSN) & ",Null," & _
                     IIf(mblnAppointment, 1, 0) & ",'" & IIf(cboAppointStyle.Visible, cboAppointStyle.Text, "") & "'," & _
                     0 & ","
            '卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSQL = strSQL & "" & IIf(lng医疗卡类别ID <> 0 And bln消费卡 = False, lng医疗卡类别ID, "NULL") & ","
            '结算卡序号_In 病人预交记录.结算卡序号%Type := Null,
            strSQL = strSQL & "" & IIf(lng医疗卡类别ID <> 0 And bln消费卡, lng医疗卡类别ID, "NULL") & ","
            '卡号_In       病人预交记录.卡号%Type := Null,
            strSQL = strSQL & "'" & mstrCardNO & "',"
            '交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSQL = strSQL & " NULL,"
            '交易说明_In   病人预交记录.交易说明%Type := Null,
            strSQL = strSQL & " NULL,"
            '合作单位_In   病人预交记录.合作单位%Type := Null
            strSQL = strSQL & " NULL,"
            '  操作类型_In   Number:=0
            strSQL = strSQL & IIf(blnAdd, 1, 0) & ","
            '  险类_IN       病人挂号记录.险类%type:=null,
            strSQL = strSQL & IIf(mintInsure = 0, "NULL", mintInsure) & ","
            '  结算模式_IN   NUMBER :=0,
            strSQL = strSQL & IIf(mPatiChargeMode = EM_先诊疗后结算, 1, 0) & ","
            '  记帐费用_IN Number:=0
            strSQL = strSQL & IIf(mRegistFeeMode = EM_RG_记帐, 1, 0) & ","
            '  退号重用_IN Number:=1
            strSQL = strSQL & IIf(mty_Para.bln退号重用, 1, 0) & ")"
            
            Call zlAddArray(cllPro, strSQL)
            '问题:31187:将挂号汇总单独出来
            If Get号别 <> "" And k = 1 Then
                If Nvl(mrsPlan!医生) = "" Then blnNoDoc = True
                strSQL = "zl_病人挂号汇总_Update("
                '  医生姓名_In   挂号安排.医生姓名%Type,
                strSQL = strSQL & IIf(blnNoDoc, "Null,", "'" & str医生 & "',")
                '  医生id_In     挂号安排.医生id%Type,
                strSQL = strSQL & "" & IIf(blnNoDoc, "0,", ZVal(lng医生ID) & ",")
                '  收费细目id_In 门诊费用记录.收费细目id%Type,
                strSQL = strSQL & "" & Val(Nvl(rsItems!项目ID)) & ","
                '  执行部门id_In 门诊费用记录.执行部门id%Type,
                strSQL = strSQL & "" & IIf(Val(Nvl(rsItems!执行科室ID)) = 0, lng挂号科室ID, Val(Nvl(rsItems!执行科室ID))) & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                strSQL = strSQL & "" & str发生时间 & ","
                '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收,3-收费预约
                strSQL = strSQL & IIf(mblnAppointment, IIf(mty_Para.bln预约时收款, 3, 1), 0) & ","
                '  号码_In       挂号安排.号码%Type := Null
                strSQL = strSQL & "'" & Get号别 & "')"
                Call zlAddArray(cllProAfter, strSQL)
            End If
            
            If mRegistFeeMode = EM_RG_划价 Then
                strSQL = _
                "zl_门诊划价记录_Insert('" & str划价NO & "'," & k & "," & ZVal(Nvl(mrsInfo!病人ID)) & ",NULL," & _
                         IIf(mstrClinic = "", "NULL", mstrClinic) & ",'" & str付款方式 & "'," & _
                         "'" & txtPatient.Text & "','" & mstrGender & "','" & mstrAge & "'," & _
                         "'" & mstrFeeType & "',NULL," & lng挂号科室ID & "," & _
                         IIf(lng挂号科室ID <> 0, lng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & IIf(rsItems!性质 = 2, 1, "NULL") & "," & _
                         rsItems!项目ID & ",'" & rsItems!类别 & "','" & rsItems!计算单位 & "'," & _
                         "NULL,1," & rsItems!数次 & ",NULL," & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & _
                         rsIncomes!收入项目ID & ",'" & rsIncomes!收据费目 & "'," & rsIncomes!单价 & "," & _
                         rsIncomes!应收 & "," & rsIncomes!实收 & "," & str发生时间 & "," & str登记时间 & ",NULL,'" & UserInfo.姓名 & "','挂号:" & strNO & "')"
                Call zlAddArray(cllPro, strSQL)
            End If
            k = k + 1
            rsIncomes.MoveNext
            Next j
        rsItems.MoveNext
    Next i
    
    If Not mblnAppointment Then
        If str医生 = UserInfo.姓名 Then
            strSQL = "ZL_病人挂号记录_更新诊室('" & strNO & "'," & Nvl(mrsInfo!病人ID) & ",'" & lblRoomName.Caption & "','" & UserInfo.姓名 & "','','','" & zl_Get预约方式ByNo(strNO) & "')"    '问题号:48350
            Call zlAddArray(cllPro, strSQL)
            strSQL = "zl_病人接诊(" & Nvl(mrsInfo!病人ID) & ",'" & strNO & "',NULL,'" & UserInfo.姓名 & "','" & lblRoomName.Caption & "')"
            Call zlAddArray(cllPro, strSQL)
        End If
    End If
    
    Err = 0: On Error GoTo ErrFirt:
    
    If cllPro.Count > 0 Then
        Err = 0: On Error GoTo ErrFirt:
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, False

        Err = 0: On Error GoTo errH:
        blnTrans = True
        If blnOneCard And lng医疗卡类别ID <> 0 And mRegistFeeMode = EM_RG_现收 And cur现金 <> 0 Then
            If Not mobjICCard.PaymentSwap(Val(cur现金), Val(cur现金), Val(lng医疗卡类别ID), 0, mstrCardNO, "", lng结帐ID, Nvl(mrsInfo!病人ID)) Then
                gcnOracle.RollbackTrans
                MsgBox "一卡通结算挂号费失败", vbInformation, gstrSysName
                Exit Sub
            Else
                strSQL = "zl_一卡通结算_Update(" & lng结帐ID & ",'" & cboPayMode.Text & "','" & mstrCardNO & "','" & lng医疗卡类别ID & "','" & "" & "'," & cur现金 & ")"
                Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If

        '医保改动
        blnNotCommit = False
        If mintInsure <> 0 And mstrYBPati <> "" And cur个帐 <> 0 Then
            '68991:strAdvance:结算模式(0或1)|挂号费收取方式(0或1) |挂号单号
            strAdvance = ""
            If mRegistFeeMode = EM_RG_记帐 Or mPatiChargeMode = EM_先诊疗后结算 Then
                strAdvance = IIf(mPatiChargeMode = EM_先诊疗后结算, "1", "0")
                strAdvance = strAdvance & "|" & IIf(mRegistFeeMode = EM_RG_记帐, "1", "0")
                strAdvance = strAdvance & "|" & strNO
            End If
            If Not gclsInsure.RegistSwap(lng结帐ID, cur个帐, mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
            End If
            blnNotCommit = True
        End If
        '问题:31187 调用医保成功后,最后作一些数据更新:内部过程中已有提交语句,所以不用再写
        zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
        Set cllCardPro = New Collection: Set cllTheeSwap = New Collection
        If mRegistFeeMode = EM_RG_现收 And Not blnOneCard And Not mPatiChargeMode = EM_先诊疗后结算 And cur现金 <> 0 Then
            If zlInterfacePrayMoney(lng结帐ID, cllCardPro, cllTheeSwap, Val(cur现金), lng医疗卡类别ID, bln消费卡) = False Then
                gcnOracle.RollbackTrans: If cmdOK.Enabled = False Then cmdOK.Enabled = True
                Exit Sub
            End If
            '修正三方交易
            zlExecuteProcedureArrAy cllCardPro, Me.Caption, False, False
        End If
        
        Err = 0: On Error GoTo OthersCommit:
        zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, False, False
OthersCommit:
        gcnOracle.CommitTrans

        If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, True, mintInsure)
        
        blnTrans = False
        On Error GoTo 0
    End If
    '打印单据
    If blnInvoicePrint Then
RePrint:
        If Not (mintInsure <> 0 And MCPAR.医保接口打印票据) And mRegistFeeMode = EM_RG_现收 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me, "NO=" & strNO, 2)
            If gblnBill挂号 Then
                If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
                    If MsgBox("挂号单号为[" & strNotValiedNos & "]票据打印未成功,是否重新进行票据打印!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                End If
            End If
        End If
    End If
    
    If blnSlipPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
    End If
    
    If blnAppointPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    End If
    
    If blnSlipPrint Or blnInvoicePrint Then
        '记录打印的凭条
        gstrSQL = "Zl_凭条打印记录_Update(4,'" & strNO & "',1,'" & UserInfo.姓名 & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    Call ReloadPage
    mstrNO = strNO
    mblnOK = True
    Unload Me
    Exit Sub
ErrFirt:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Sub
errH:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Sub
ErrGo:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ReloadPage()
    On Error GoTo errHandle
    Call LoadRegPlans(False)
    Call ClearPatient
    Call ClearRegInfo
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function Check复诊(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As Boolean
'功能:判断病人是否再次到“相同临床性质的临床科室”挂号
'     包括挂过号的,或住过院的,复诊不好确定时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select a.临床科室id" & vbNewLine & _
    "       From (Select 执行部门id 临床科室id From 病人挂号记录 Where 病人id = [1] and 记录性质=1 and 记录状态=1 " & vbNewLine & _
    "             Union All" & vbNewLine & _
    "             Select 出院科室id 临床科室id From 病案主页 Where 病人id = [1]) a" & vbNewLine & _
    "       Where Exists (Select 1" & vbNewLine & _
    "                    From 临床部门 b" & vbNewLine & _
    "                    Where b.部门id = a.临床科室id And b.工作性质 = (Select 工作性质 From 临床部门 Where 部门id = [2] And Rownum=1))"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng执行部门ID)
    Check复诊 = Not rsTmp.EOF
End Function

Private Function zlIsAllowPatiChargeFeeMode(ByVal lng病人ID As Long, ByVal int原结算模式 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许改变病人收费模式
    '入参:lng病人ID-病人ID
    '       int原结算模式-0表示先结算后诊疗;1表示先诊疗后结算
    '返回:允许调整收费模式,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-12-25 10:06:49
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim dtDate As Date, intDay As Integer
    On Error GoTo errHandle
    
'    If mbytMode = 1 Then zlIsAllowPatiChargeFeeMode = True: Exit Function '预约不处理
    '模式未调整，直接返回true
    If int原结算模式 = mPatiChargeMode Then zlIsAllowPatiChargeFeeMode = True: Exit Function
    
      
    If int原结算模式 = 1 Then
        '原为先诊疗后结算且存在未结费用的,则必须采用记帐模式
        strSQL = "" & _
        "   Select 1 " & _
        "   From 病人未结费用 " & _
        "   Where 病人id = [1] And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
        If rsTemp.EOF = False Then
            MsgBox "注意:" & vbCrLf & "  当前病人的就诊模式为先诊疗后结算且存在未结费用，" & _
                                          vbCrLf & "不允许调整该病人的就诊模式,你可以先对未结费用结帐后" & _
                                          vbCrLf & "再挂号或不调整病人的就诊模式", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        intDay = -1 * Val(Left(gobjDatabase.GetPara(21, glngSys, , "01") & "1", 1))
        dtDate = DateAdd("d", intDay, gobjDatabase.CurrentDate)
        ' 上次为"先诊疗后结算",本次为"先结算后诊疗"的,同时满足未发生医嘱业务数据的 ,
        '   则不允许更改就诊模式
        strSQL = "Select 1 " & _
        " From 病人挂号记录 A, 病人医嘱记录 B " & _
        " Where a.病人id + 0 = b.病人id And a.No || '' = b.挂号单  " & _
        "               And a.记录状态 = 1 And a.记录性质 = 1 And a.登记时间 - 0 >= [2] " & _
        "               And  a.病人id = [1] And rownum<2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, dtDate)
        If rsTemp.EOF Then
            '未发生医嘱数据
            MsgBox "注意:" & vbCrLf & "  当前病人的就诊模式为先诊疗后结算," & vbCrLf & "  不允许调整该病人的就诊模式!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    zlIsAllowPatiChargeFeeMode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub ClearRegInfo()
    If cboArrangeNo.ListCount <> 0 Then cboArrangeNo.ListIndex = 0
    lblDeptName.Caption = ""
    lblRoomName.Caption = ""
    cboRemark.Text = ""
    chkBook.Value = IIf(mty_Para.bln默认购买病历, 1, 0)
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = "0.00"
    lblPayMoney.Caption = "0.00"
    txtPatient.SetFocus
    lbl急.Visible = False
End Sub

Private Function zlIsNotSucceedPrintBill(ByVal bytType As Byte, ByVal strNos As String, ByRef strOutValidNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否已经正常打印
    '入参:bytType-1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '       strNos-本次打印票据的单据,用逗号分离
    '出参:strOutValidNos-打印失败的单据号
    '返回:存在不存功票据的打印,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-01-16 18:06:01
    '问题:44322,44326,44332,44330
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempNos As String, rsTemp As ADODB.Recordset
    Dim strSQL As String, strBillNos As String
    Dim bytBill As Byte
    On Error GoTo errHandle
    strBillNos = Replace(Replace(strNos, "'", ""), " ", "")
    strSQL = "" & _
        "Select  /*+ rule */ distinct  B.NO " & _
        " From 票据使用明细 A,票据打印内容 B,Table( f_Str2list([2])) J" & _
        " Where A.打印ID =b.ID And B.数据性质=[1] And B.No=J.Column_value "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查票据是否打印", bytType, strBillNos)
    
    strTempNos = ""
    With rsTemp
        Do While Not .EOF
            If InStr(1, "," & strBillNos & ",", "," & !NO & ",") = 0 Then
                strTempNos = strTempNos & "," & !NO
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then strTempNos = "," & strBillNos
    End With
    If strTempNos <> "" Then strTempNos = Mid(strTempNos, 2)
    rsTemp.Close: Set rsTemp = Nothing
    strOutValidNos = strTempNos
    zlIsNotSucceedPrintBill = strTempNos <> ""
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckValied() As Boolean
    Dim i As Integer
    '保存前检查
    If mrsInfo Is Nothing Then
        MsgBox "无法确定病人信息,请先选择一个病人！", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "无法确定病人信息,请先选择一个病人！", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan Is Nothing Then
        MsgBox "无法确定号别信息,请先选择一个号别！", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan.State = 0 Then
        MsgBox "无法确定号别信息,请先选择一个号别！", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan.RecordCount = 0 Then
        MsgBox "无法确定号别信息,请先选择一个号别！", vbInformation, gstrSysName
        Exit Function
    End If
    If cboPayMode.Text = "" And cboPayMode.Visible And Val(lblTotal.Caption) <> 0 Then
        MsgBox "没有确定可用的结算方式,不能完成挂号!", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        If IsNull(mrsPlan!排班) Then
            MsgBox "预约不收款模式下,不能挂不当班的号别!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Nvl(mrsInfo!姓名) <> txtPatient.Text Then
        If MsgBox("当前病人姓名已经发生变化,是否重新读取病人信息?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            Call GetPatient(IDKind.GetCurCard, txtPatient.Text, False)
            Exit Function
        Else
            txtPatient.Text = Nvl(mrsInfo!姓名)
        End If
    End If
    
    If InStr(gstrPrivs, ";挂号费别打折;") = 0 Then
        For i = 1 To vsfMoney.Rows - 1
            If Val(vsfMoney.TextMatrix(i, 2)) <> Val(vsfMoney.TextMatrix(i, 1)) Then
                MsgBox "你没有权限给病人使用打折费别,不能完成挂号", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    CheckValied = True
End Function

Private Function zlInterfacePrayMoney(ByVal lng挂号结帐ID As Long, ByRef cllPro As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double, lng医疗卡类别ID As Long, bln消费卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If lng医疗卡类别ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, lng医疗卡类别ID, bln消费卡, mstrCardNO, lng挂号结帐ID, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
     If lng挂号结帐ID <> 0 Then
        '问题:58322
        'mbytMode As Integer '0-挂号,1-预约,2-接收,3-取消预约 ,4-退号 预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
        If Not bln消费卡 Then
            '消费卡已经在插入挂号记录时,已经扣款
            Call zlAddUpdateSwapSQL(False, lng挂号结帐ID, lng医疗卡类别ID, bln消费卡, mstrCardNO, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lng挂号结帐ID, lng医疗卡类别ID, bln消费卡, mstrCardNO, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddThreeSwapSQLToCollection(ByVal bln预交款 As Boolean, _
    ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str卡号 As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方结算数据
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    ' 出参:cllPro-返回SQL集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '先提交,这样避免风险,再更新相关的交易信息
    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
    varData = Split(strExpend, "||")
    Dim str交易信息 As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If gobjCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                    str交易信息 = Mid(str交易信息, 3)
                    'Zl_三方结算交易_Insert
                    strSQL = "Zl_三方结算交易_Insert("
                    '卡类别id_In 病人预交记录.卡类别id%Type,
                    strSQL = strSQL & "" & lng卡类别ID & ","
                    '消费卡_In   Number,
                    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                    '卡号_In     病人预交记录.卡号%Type,
                    strSQL = strSQL & "'" & str卡号 & "',"
                    '结帐ids_In  Varchar2,
                    strSQL = strSQL & "'" & strIDs & "',"
                    '交易信息_In Varchar2:交易项目|交易内容||...
                    strSQL = strSQL & "'" & str交易信息 & "',"
                    '预交款缴款_In Number := 0
                    strSQL = strSQL & IIf(bln预交款, "1", "0") & ")"
                    zlAddArray cllPro, strSQL
                    str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
        End If
    Next
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSQL = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & lng卡类别ID & ","
        '消费卡_In   Number,
        strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSQL = strSQL & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSQL = strSQL & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSQL = strSQL & IIf(bln预交款, "1", "0") & ")"
        zlAddArray cllPro, strSQL
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddUpdateSwapSQL(ByVal bln预交 As Boolean, ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    str卡号 As String, str交易流水号 As String, str交易说明 As String, _
    ByRef cllPro As Collection, Optional int校对标志 As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新三方交易流水号和流水说明
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    '出参:cllPro-返回SQL集
    '编制:刘兴洪
    '日期:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  消费卡_In     Number,
    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSQL = strSQL & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSQL = strSQL & "'" & strIDs & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type
    strSQL = strSQL & "'" & str交易说明 & "',"
    '预交款缴款_In Number := 0
    strSQL = strSQL & "" & IIf(bln预交, 1, 0) & ","
    '退费标志 :1-退费;0-付费
    strSQL = strSQL & "0,"
    '校对标志
    strSQL = strSQL & "" & IIf(int校对标志 = 0, "NULL", int校对标志) & ")"
    zlAddArray cllPro, strSQL
End Function

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
        If cboArrangeNo.ListCount = 1 Then gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub SetControl()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTemp As String, i As Integer
    If mblnAppointment Then
        lblRoom.Visible = False
        picRoom.Visible = False
        lblDept.Visible = False
        picDept.Visible = False
        lblDept.Left = lblLimit.Left
        picDept.Left = lblDept.Left + lblDept.Width + 30
        picDept.Width = Me.Width - 240 - picDept.Left
        chkBook.Value = 0
        chkBook.Visible = False
        cboRemark.Width = 7170
        If mty_Para.bln预约时收款 Then
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
        Else
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
        End If
        cboAppointStyle.Clear
        strSQL = "Select 名称,缺省标志 From 预约方式"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
        Do While Not rsTmp.EOF
            cboAppointStyle.AddItem Nvl(rsTmp!名称)
            If Val(Nvl(rsTmp!缺省标志)) = 1 Then cboAppointStyle.ListIndex = cboAppointStyle.NewIndex
            rsTmp.MoveNext
        Loop
        strTemp = gobjDatabase.GetPara("缺省预约方式", glngSys, 9000, "")
        For i = 0 To cboAppointStyle.ListCount - 1
            If Mid(cboAppointStyle.List(i), InStr(cboAppointStyle.List(i), ".") + 1) = strTemp Then
                cboAppointStyle.ListIndex = i
            End If
        Next i
    Else
        lblDate.Visible = False
        lblTime.Visible = False
        dtpDate.Visible = False
        dtpTime.Visible = False
        cmdTime.Visible = False
        If mty_Para.byt挂号模式 = 0 Then
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
        Else
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Err = 0
    mblnInit = True
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    If mblnAppointment Then
        Me.Caption = "医生站预约"
        lblAppointStyle.Visible = True
        cboAppointStyle.Visible = True
    Else
        Me.Caption = "医生站挂号"
        lblAppointStyle.Visible = False
        cboAppointStyle.Visible = False
    End If
    Call InitPara
    chkBook.Value = IIf(mty_Para.bln默认购买病历, 1, 0)
    Call InitIDKind
    Call InitTime
    Call GetAll医生
    If LoadRegPlans(False) = False Then
        mblnUnload = True
    End If
    Call LoadPayMode
    Call SetControl
    If mblnAppointment And mlng病人ID <> 0 Then
        Call GetPatient(IDKind.GetCurCard, "-" & mlng病人ID, False)
    End If
    cmdNewPati.Enabled = InStr(gstrPrivs, ";挂号病人建档;") > 0
End Sub

Private Sub InitTime()
    dtpDate.Value = Format(gobjDatabase.CurrentDate + mintSysAppLimit, "yyyy-mm-dd")
    dtpDate.minDate = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd")
    dtpTime.Value = Format(gobjDatabase.CurrentDate, "hh:mm:ss")
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.卡号
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call gobjControl.TxtSelAll(txtPatient)
    Call gobjCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：医保身份验卡
    '编制：刘兴洪
    '日期：2010-07-14 11:32:08
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim str病人类型 As String
    Dim rsTmp As ADODB.Recordset
    Dim cur余额 As Currency
    Dim curMoney As Currency
    Dim blnDeposit As Boolean, blnInsure As Boolean
    If mrsInfo Is Nothing Then
        lng病人ID = 0
        str病人类型 = ""
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        str病人类型 = Nvl(mrsInfo!病人类型)
    End If

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False

    Dim strAdvance As String    '结算模式(0-先结算后诊疗或1-先诊疗后结算)|挂号费收取方式(0-现收或1-记帐)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng病人ID, mintInsure, strAdvance)
    mRegistFeeMode = EM_RG_现收: mPatiChargeMode = EM_先结算后诊疗
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng病人ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
    
    txtPatient.Text = "-" & lng病人ID
    Call txtPatient_Validate(False)    '其中的Setfocus调用使本事件(txtPatient_KeyPress)执行完后,不会再次自动执行txtPatient_Validate
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
    Call SetPatiColor(txtPatient, str病人类型, vbRed)
    txtPatient.BackColor = &HE0E0E0
    txtPatient.Locked = True

    If strAdvance <> "" Then
        varData = Split(strAdvance & "|", "|")
        mPatiChargeMode = IIf(Val(varData(0)) = 1, EM_先诊疗后结算, EM_先结算后诊疗)
        mRegistFeeMode = IIf(Val(varData(1)) = 1, EM_RG_记帐, EM_RG_现收)
    End If
    MCPAR.不收病历费 = gclsInsure.GetCapability(support挂号不收取病历费, lng病人ID, mintInsure)
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure)
    mlng领用ID = 0
    curMoney = GetRegistMoney
    Set rsTmp = GetMoneyInfoRegist(lng病人ID, , , 1)

    cur余额 = 0
    Do While Not rsTmp.EOF
        cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
        cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))

        rsTmp.MoveNext
    Loop
    If cur余额 > 0 Then
        lblMoney.Caption = "门诊预交余额:" & Format(cur余额, "0.00")
        If cur余额 >= curMoney Then
            blnDeposit = True
        Else
            blnDeposit = False
        End If
    End If
    
    mcur个帐余额 = gclsInsure.SelfBalance(lng病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur个帐透支, mintInsure)
    lblMoney.Caption = lblMoney.Caption & "/个人帐户余额:" & Format(mcur个帐余额, "0.00")
    If gclsInsure.GetCapability(support挂号使用个人帐户, lng病人ID, mintInsure) = False Then
        blnInsure = False
    Else
        If mcur个帐余额 + mcur个帐透支 >= curMoney Then
            blnInsure = True
        Else
            blnInsure = False
        End If
    End If
    Call LoadPayMode(blnDeposit, blnInsure)
    If mRegistFeeMode = EM_RG_记帐 Then
        lblSum.Caption = "记帐"
        picPayMoney.Visible = False
        cboPayMode.Visible = False
        lblPayMode.Visible = False
    Else
        lblSum.Caption = "合计"
    End If
    If mRegistFeeMode = EM_RG_现收 Then
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_现收
        Else
            If mty_Para.byt挂号模式 = 0 Then
                mRegistFeeMode = EM_RG_现收
                picPayMoney.Visible = True
                cboPayMode.Visible = True
                lblPayMode.Visible = True
            Else
                mRegistFeeMode = EM_RG_划价
                picPayMoney.Visible = False
                cboPayMode.Visible = False
                lblPayMode.Visible = False
            End If
        End If
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '0-门诊号,1-姓名,2-挂号单,3-就诊卡号,4-医保号
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    Static sngBegin As Single
    Dim sngNow As Single
    '医保验证
    If txtPatient.Text = "" And KeyAscii = 13 Then
        KeyAscii = 0
        Call zlInusreIdentify
    End If
    
    If KeyAscii <> 0 And KeyAscii > 32 And mty_Para.bln挂号必须刷卡 Then
        sngNow = Timer
        If txtPatient.Text = "" Then
            sngBegin = sngNow
        ElseIf Format((sngNow - sngBegin) / (Len(txtPatient.Text) + 1), "0.000") >= 0.04 Then    '>0.007>=0.01
            txtPatient.Text = Chr(KeyAscii)
            txtPatient.SelStart = 1
            KeyAscii = 0
            sngBegin = sngNow
        End If
    End If
    
    strKind = IDKind.GetCurCard.名称
    txtPatient.PasswordChar = IIf(IDKind.GetCurCard.卡号密文规则 <> "", "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.bln缺省卡号密文)
        intLen = gobjSquare.int缺省卡号长度
    Case "门诊号"
        If InStr("0123456789-" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "身份证"
    Case "医保号"
    Case Else
            If IDKind.GetCurCard.接口序号 <> 0 Then
                blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.GetCurCard.卡号密文规则 <> "")
                intLen = IDKind.GetCurCard.卡号长度
            End If
    End Select
    
    '刷卡完毕或输入号码后回车
    If (blnCard And Len(txtPatient.Text) = intLen - 1 And KeyAscii <> 8) Or (KeyAscii = 13) Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0: mblnCard = True
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCard)
        mblnCard = False
        gobjControl.TxtSelAll txtPatient
   End If
End Sub

Private Function CheckNoValied(ByVal lngRow As Long) As Boolean
    CheckNoValied = True
End Function

Private Function zl_Get预约方式ByNo(strNO As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据挂号单据号获取病人预约方式
    '入参:strNo-挂号单据号
    '返回:预约方式
    '编制:王吉
    '日期:2012-07-03
    '问题号:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str预约方式 As String
    Dim rsTemp As Recordset
    strSQL = "" & _
        "Select 预约方式 From 病人挂号记录 Where 记录状态=1 And No=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取预约方式", strNO)
    If rsTemp Is Nothing Then zl_Get预约方式ByNo = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_Get预约方式ByNo = "": Exit Function
    While rsTemp.EOF = False
        str预约方式 = Nvl(rsTemp!预约方式)
        rsTemp.MoveNext
    Wend
    zl_Get预约方式ByNo = str预约方式
End Function

Public Function Get失约号(ByVal str号别 As String, ByVal datThis As Date) As Long
   '获取安排在某一天.预约失约数
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strDat  As String
'    If mty_Para.bln失约用于挂号 = False Or mty_Para.lng预约有效时间 <= 0 Then Exit Function
    strSQL = "                " & " SELECT count(1) AS 失约号 "
    strSQL = strSQL & vbNewLine & " FROM 挂号序号状态 "
    strSQL = strSQL & vbNewLine & " WHERE 号码=[1] AND 状态=2 AND 日期-[3]/24/60 <SYSDATE AND To_Char(日期,'yyyy-MM-dd')=[2]"
    strDat = Format(datThis, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, str号别, strDat, mty_Para.lng预约有效时间)
    If rsTmp.EOF Then
        Get失约号 = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    Get失约号 = Val(Nvl(rsTmp!失约号, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub txtPatient_LostFocus()
    Call gobjCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
'        gobjControl.TxtSelAll txtPatient
'    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIndex As Long
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strID:
        If txtPatient.Text = "" Then
            Call mobjIDCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        DoEvents
        If txtPatient.Visible = True And txtPatient.Enabled Then
            Call txtPatient.SetFocus
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdNewPati_Click
    Else
        IDKind.ActiveFastKey
    End If
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIndex As Long
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strNO
        If txtPatient.Text = "" Then
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnInputIDCard As Boolean = False, Optional ByRef Cancel As Boolean)
    '功能：获取病人信息
    '参数：blnCard=是否就诊卡刷卡
    '
    '         blnInputIDCard-是否身份证刷卡
    '出参:Cancel-为true表示返回的放弃读取病人信息
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim strInputInfo As String '保存传入的输入文本 避免在使用身份证号 对病人进行查找后 被替换成"-" 病人ID的情况
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim bln医保号 As Boolean, rsFeeType As ADODB.Recordset
    Dim IntMsg As VbMsgBoxResult
    Dim blnOtherType As Boolean '非法卡类别

    strInputInfo = strInput
    
    On Error GoTo errH
    bln医保号 = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard

    strSQL = "Select  A.病人ID,A.门诊号,A.住院号,A.就诊卡号,A.费别,A.医疗付款方式,A.姓名,A.性别,A.年龄,A.出生日期,A.出生地点,A.身份证号,A.其他证件,A.身份,A.职业,A.民族,A.病人类型, " & _
             "A.国籍,A.籍贯,A.区域,A.学历,A.婚姻状况,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.监护人,A.联系人姓名,A.联系人关系,A.联系人地址,A.联系人电话,A.户口地址, " & _
             "A.户口地址邮编,A.Email,A.QQ,A.合同单位id,A.工作单位,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人,A.担保额,A.担保性质,A.就诊时间,A.就诊状态, " & _
             "A.就诊诊室,A.住院次数,A.当前科室id,A.当前病区id,A.当前床号,A.入院时间,A.出院时间,A.在院,A.IC卡号,A.健康号,A.医保号,A.险类,A.查询密码,A.登记时间,A.停用时间,A.锁定,A.联系人身份证号, " & _
             "B.名称 险类名称,A.查询密码 As 卡验证码,A.结算模式 From 病人信息 A,保险类别 B  Where A.险类 = B.序号(+) And A.停用时间 is NULL"

    If mty_Para.bln住院病人挂号 = False Then
        str非在院 = " And Not Exists(Select 1 From 病案主页 Where 病人ID=A.病人ID   And 主页ID<>0 And 主页ID=A.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    End If
   
    If blnCard And objCard.名称 Like "姓名*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
'        Else
'            lng卡类别ID = gCurSendCard.lng卡类别ID
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0

        If lng病人ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And A.门诊号=[2]" & str非在院
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And A.病人ID=[2]" & _
        IIf(mstrYBPati <> "", "", str非在院)
    ElseIf blnInputIDCard Then  '单独的身份证识别
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        strInput = "-" & lng病人ID
        strSQL = strSQL & " And A.病人ID=[2] " & str非在院
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If Not mty_Para.bln姓名模糊查找 Or mty_Para.bln姓名模糊查找 And Len(txtPatient.Text) < 2 Then
                    Set mrsInfo = Nothing: Exit Sub
                End If
                strPati = _
                    " Select distinct 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                    " From 病人信息 A " & _
                    " Where Rownum <101 And A.停用时间 is NULL And A.姓名 Like [1]" & str非在院 & _
                    IIf(mty_Para.lng姓名查找天数 = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])")
                    
'                strPati = strPati & " Union ALL " & _
'                        "Select 0,0 as ID,-NULL,'[新病人]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                strPati = strPati & " Order by 排序ID,姓名"
                    
                vRect = GetControlRect(txtPatient.hWnd)
                Set rsTmp = gobjDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", mty_Para.lng姓名查找天数)
                If Not rsTmp Is Nothing Then
                    If rsTmp!ID = 0 Then '当作新病人
                        txtPatient.Text = ""
                        MsgBox "没有找到对应的病人信息，请检查输入信息是否正确或者病人是否建档！", vbInformation, gstrSysName
                        Set mrsInfo = Nothing: Exit Sub
                    Else '以病人ID读取
                        strInput = rsTmp!病人ID
                        strSQL = strSQL & " And A.病人ID=[1]"
                    End If
                Else '取消选择
                    txtPatient.Text = ""
                    Set mrsInfo = Nothing: Exit Sub
                End If
            Case "医保号"
                strInput = UCase(strInput)
                bln医保号 = True
                If mblnOlnyBJYB And gobjCommFun.ActualLen(strInput) >= 9 Then
                    strSQL = strSQL & " And A.医保号 like [3] " & str非在院
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And A.医保号=[1]" & str非在院
                End If
                
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSQL = strSQL & " And A.病人ID=[2] " & str非在院
                 
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSQL = strSQL & " And A.病人ID=[2] " & str非在院
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[1]" & str非在院
             Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                strSQL = strSQL & " And A.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
ReadPati:
    If Mid(mstrCardPass, 1, 1) = "1" And strPassWord <> "" Then
        If Not gobjCommFun.VerifyPassWord(Me, "" & strPassWord) Then
            MsgBox "病人身份验证失败！", vbInformation, gstrSysName
            ClearPatient
            Exit Sub
        End If
    End If
    Set mrsInfo = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, Mid(strInput, 2), strTemp)
    strInput = strInputInfo
    If Not mrsInfo.EOF Then
        Set mfrmPatiInfo = New frmPatiInfo
        txtPatient.Text = Nvl(mrsInfo!姓名) '会调用Change事件
        txtPatient.BackColor = &H80000005
        lblSum.Caption = "合计"
        Call SetControl
        '在调用txtPatient_Change事件后在门诊号和病人姓名都为空的情况下 无法识别该病人信息 出现错误
        '对这类数据库数据错误不再进行后续的处理
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(Trim(mstr险类) = "", txtPatient.ForeColor, vbRed))
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        mstrGender = Nvl(mrsInfo!性别)
        txtPatient.PasswordChar = ""
        
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        mstrFeeType = Nvl(mrsInfo!费别)
        If mstrFeeType = "" Then
            strSQL = "Select 名称 From 费别 Where 缺省标志 = 1"
            Set rsFeeType = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If Not rsFeeType.EOF Then
                mstrFeeType = Nvl(rsFeeType!名称)
            End If
        End If
        mstrAge = Nvl(mrsInfo!年龄)
        mstrClinic = Nvl(mrsInfo!门诊号)
        If mstrClinic = "" Then
            mstrClinic = gobjDatabase.GetNextNo(3)
            mblnChangeFeeType = True
        Else
            mblnChangeFeeType = False
        End If
        
        lblInfo.Caption = "性别:" & mstrGender & "   年龄:" & mstrAge & "   门诊号:" & mstrClinic & "   费别:" & mstrFeeType
        
        '病人预交款信息
        Set rsTmp = GetMoneyInfoRegist(mrsInfo!病人ID, , , 1)
        cur余额 = 0
        Do While Not rsTmp.EOF
            cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
            cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))
            rsTmp.MoveNext
        Loop
        If cur余额 > 0 Then
            lblMoney.Caption = "门诊预交余额:" & Format(cur余额, "0.00")
            curMoney = GetRegistMoney
            If cur余额 >= curMoney Then
                Call LoadPayMode(True)
            Else
                Call LoadPayMode
            End If
        Else
            lblMoney.Caption = "门诊预交余额:0.00"
            Call LoadPayMode
        End If
        Call LoadFeeItem(Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1)
        cmdNewPati.ToolTipText = "详细信息"
        cmdNewPati.Enabled = True
        If cboArrangeNo.Enabled And cboArrangeNo.Visible Then cboArrangeNo.SetFocus
        If cboArrangeNo.ListCount = 1 Then gobjCommFun.PressKeyEx vbKeyTab
    Else
NewPati:
        MsgBox "没有找到对应的病人信息，请检查输入信息是否正确或者病人是否建档！", vbInformation, gstrSysName
        ClearPatient
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub ClearPatient()
    Set mfrmPatiInfo = New frmPatiInfo
    txtPatient.Text = ""
    txtPatient.BackColor = &H80000005
    txtPatient.ForeColor = vbBlack
    txtPatient.Locked = False
    mlngNewPatiID = 0
    mstrGender = ""
    mstrAge = ""
    cmdNewPati.ToolTipText = "新增病人"
    cmdNewPati.Enabled = InStr(gstrPrivs, ";挂号病人建档;") > 0
    mstrClinic = ""
    mblnNewPati = False
    mstrFeeType = ""
    lblInfo.Caption = "性别:     年龄:       门诊号:              费别:  "
    lblMoney.Caption = "门诊预交余额:0.00  "
    lblSum.Caption = "合计"
    mintInsure = 0
    mlng领用ID = 0
    chkBook.Enabled = True
    LoadPayMode False, False
    Set mrsInfo = Nothing
    If mblnAppointment Then
        mRegistFeeMode = EM_RG_现收
    Else
        If mty_Para.byt挂号模式 = 0 Then
            mRegistFeeMode = EM_RG_现收
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
        Else
            mRegistFeeMode = EM_RG_划价
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
        End If
    End If
End Sub

Private Function GetRegistMoney(Optional blnOnlyReg As Boolean) As Currency
    '功能：获取当前挂号单的合计金额
    'blnOnlyReg-是否仅仅读取挂号费用
    Dim cur合计 As Currency, i As Integer
    Dim cur应收 As Currency, j As Integer
    Dim k As Integer
    If Not blnOnlyReg Then
        For i = 1 To vsfMoney.Rows - 1
            cur合计 = cur合计 + Val(vsfMoney.TextMatrix(i, 2))
        Next
    Else
        For i = 1 To vsfMoney.Rows - 1
            cur合计 = cur合计 + Val(vsfMoney.TextMatrix(i, 2))
        Next
    End If
    GetRegistMoney = cur合计
End Function

Private Sub LoadPayMode(Optional ByVal blnPrepay As Boolean = False, Optional ByVal blnInsure As Boolean = False)
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, str性质 As String
    
    strSQL = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式 And Instr([2] ,','||B.性质||',')>0" & _
        " Order by B.编码"
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, "挂号", ",3,7,8,")
    
    Set mcolCardPayMode = New Collection
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    
    With cboPayMode
        .Clear: j = 0
'        Do While Not rsTemp.EOF
'            blnFind = False
'            For i = 0 To UBound(varData)
'                varTemp = Split(varData(i) & "|||||", "|")
'                If varTemp(6) = Nvl(rsTemp!名称) Then
'                    blnFind = True
'                    Exit For
'                End If
'            Next
'
'            If Not blnFind Then
'                .AddItem Nvl(rsTemp!名称)
'                mcolCardPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
'                If Val(Nvl(rsTemp!缺省)) = 1 Then
'                    If .ListIndex = -1 Then
'                         .ItemData(.NewIndex) = 1: .ListIndex = .NewIndex
'                    End If
'                End If
'                j = j + 1
'            End If
'            rsTemp.MoveNext
'        Loop
     
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                rsTemp.Filter = "名称='" & varTemp(6) & "'"
                If Not rsTemp.EOF Then
                    mcolCardPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    j = j + 1
                End If
            End If
        Next
    End With
    
    If blnPrepay Then
        cboPayMode.AddItem "预交金"
        If mty_Para.bln优先使用预交 Then
            cboPayMode.ListIndex = cboPayMode.NewIndex
        End If
    End If
    
    If blnInsure Then
        rsTemp.Filter = "性质 = 3"
        If rsTemp.EOF Then
            mstrInsure = ""
            MsgBox "不能加载医保结算方式,请检查!", vbInformation, gstrSysName
        Else
            cboPayMode.AddItem Nvl(rsTemp!名称)
            mstrInsure = Nvl(rsTemp!名称)
            If Not mty_Para.bln优先使用预交 Or blnPrepay = False Then
                cboPayMode.ListIndex = cboPayMode.NewIndex
            End If
            If (mintInsure <> 0 And MCPAR.不收病历费) And cboPayMode.Text = mstrInsure And cboPayMode.Visible Then
                chkBook.Enabled = False
                chkBook.Value = 0
            Else
                chkBook.Enabled = True
            End If
        End If
    End If
    
    If cboPayMode.ListCount > 0 And cboPayMode.ListIndex = -1 Then
        cboPayMode.ListIndex = 0
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function LoadRegPlans(ByVal blnCache As Boolean) As Boolean
    Dim strTime As String, strState As String, strWhere As String
    Dim strSQL As String, strIF As String
    Dim i As Integer, k As Integer
    Dim DateThis As Date, strZero As String
    Dim str挂号安排 As String
    Dim str挂号安排计划 As String
    Dim str排序         As String
    On Error GoTo errH
    
    str排序 = "Decode(医生,Null,3,Decode(科室ID," & mlngDept & ",1,2)),医生,科室,号别,项目,已挂"
    
    If Not blnCache Then
        strSQL = "Zl_挂号安排_Autoupdate"
        gobjDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    If Not blnCache Then
        If gstrDeptIDs <> "" Then strIF = " And Instr(','||[4]||',',','||P.科室ID||',')>0"
        If mty_Para.bln包含科室安排 Then
            strIF = strIF & " And (P.医生姓名 = [1] or P.医生姓名 Is Null)"
        Else
            strIF = strIF & " And (P.医生姓名 = [1])"
        End If
        
        str挂号安排 = "" & _
                "            Select A.ID, A.号码, A.号类, A.科室id, A.项目id, A.医生id, A.医生姓名, A.病案必须, A. 周日, A.周一, A.周二, A.周三, " & _
                "                   A.周四 , A.周五, A.周六, A.分诊方式,a.开始时间,a.终止时间, A.序号控制, B.限号数, B.限约数,a.停用日期 " & vbNewLine & _
                "            From 挂号安排 A, 挂号安排限制 B " & vbNewLine & _
                "            Where a.停用日期 Is Null And " & "[5] Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
                "                 Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                "                  And a.ID = B.安排id(+) And Trunc(Sysdate)+Nvl(A.预约天数," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5] And Decode(To_Char([5], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) = B.限制项目(+)" & vbNewLine
      
        If mblnAppointment Then
            DateThis = Format(dtpDate, "yyyy-mm-dd hh:mm:ss")
        Else
            DateThis = gobjDatabase.CurrentDate
        End If
        '取对应日期安排的时间段
        strSQL = "Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)"
        
        '该部分语句取现在所对应的时间段
        strTime = _
            "Select 时间段 From 时间段 Where 号类 Is Null And 站点 Is Null And " & _
            "    ('3000-01-10 '||To_Char([5],'HH24:MI:SS') Between" & _
            "               Decode(Sign(开始时间-终止时间),1,'3000-01-09 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS'),'3000-01-10 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS'))" & _
            "               And '3000-01-10 '||To_Char(终止时间,'HH24:MI:SS'))" & _
            " Or" & _
            " ('3000-01-10 '||To_Char([5],'HH24:MI:SS')  Between" & _
            "   '3000-01-10 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS') And" & _
            "     Decode(Sign(开始时间-终止时间),1,'3000-01-11 '||To_Char(终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(终止时间,'HH24:MI:SS')))"
            
        '该部分语句当时读取各种安排的挂号情况
        strState = _
        "   Select A.ID as 安排ID,B.已挂数,B.已约数" & _
        "   From (" & str挂号安排 & ") A,病人挂号汇总 B" & _
        "   Where A.科室ID = B.科室ID And A.项目ID = B.项目ID" & _
        "               And Nvl(A.医生ID,0)=Nvl(B.医生ID,0) " & _
        "               And Nvl(A.医生姓名,'医生')=Nvl(B.医生姓名,'医生') " & _
        "               And (A.号码=B.号码 or B.号码 is Null )  And B.日期=[6]"
        
        If mblnAppointment Then
            str挂号安排计划 = " " & _
                "             Select A.ID,A.ID as 计划ID, A.安排id, A.号码, A.项目id, A.安排人, A.安排时间, A. 周日, A.周一, A.周二, A.周三, A.周四, A.周五," & _
                "                    A.周六 , A.分诊方式, A.序号控制, B.限号数, B.限约数, A.生效时间, A.失效时间 ,A.医生姓名,A.医生ID " & _
                "             From 挂号安排计划 A, 挂号计划限制 B," & vbNewLine & _
                "                  (" & vbNewLine & _
                "                      Select Max(生效时间) As 生效时间, 安排id" & _
                "                      From 挂号安排计划 " & vbNewLine & _
                "                      Where 审核时间 Is Not Null And  [5] Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
                "                          Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD'))  " & vbNewLine & _
                "                       Group By 安排id" & vbNewLine & _
                "                   ) C" & _
                "             Where A.审核时间 Is Not Null And ([5] Between  A.生效时间  And A.失效时间)" & _
                "                   And A.ID = B.计划id(+) And " & vbNewLine & _
                "                   Decode(To_Char([5], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6'," & _
                "                  '周五', '7', '周六', Null) = B.限制项目(+) And A.生效时间 = C.生效时间 And A.安排id = C.安排id"

            strSQL = _
            " Select P.ID,0 as 计划ID,P.号码 ,P.号类,P.科室ID,P.项目ID," & _
            "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(P.病案必须,0) as 病案必须," & _
            "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
            "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)  as 排班 " & _
            " From (" & str挂号安排 & ") P" & _
            " Where    Not Exists(Select 1 From 挂号安排计划 where 安排ID=P.id And ([5] BETWEEN 生效时间  and 失效时间)  And 审核时间 is not NULL  ) " & _
            "          And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=P.ID and [5] between 开始停止时间 and 结束停止时间 )" & _
            " Union ALL " & _
            " Select   C.ID,P.计划ID,C.号码,C.号类,C.科室ID,P.项目ID," & _
            "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(C.病案必须,0) as 病案必须," & _
            "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
            "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)  as 排班 " & _
            " From (" & str挂号安排计划 & ") P, 挂号安排 C" & _
            " Where P.安排ID=C.ID  And C.停用日期 Is  NULL  And Trunc(Sysdate)+Nvl(C.预约天数," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]  " & _
            "           And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=C.ID and [5] between 开始停止时间 and 结束停止时间 )"
            strSQL = "(" & strSQL & ") P"
        Else
            strSQL = _
                        " (Select P.ID,0 as 计划ID,P.号码 ,P.号类,P.科室ID,P.项目ID," & _
                        "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(P.病案必须,0) as 病案必须," & _
                        "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
                        "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL) as 排班 " & _
                        " From (" & str挂号安排 & ") P "
            strSQL = strSQL & vbNewLine & "  ) P"
        End If
        
        strSQL = _
                    "Select Distinct " & _
                    "       P.ID,p.计划ID,P.号码 as 号别,P.号类,P.科室ID,B.名称 As 科室,P.项目ID,C.名称 As 项目," & _
                    "       P.医生ID,P.医生姓名 as 医生,Nvl(A.已挂数,0) as 已挂,Nvl(A.已约数,0) as 已约," & _
                    "       P.限号数 as 限号,P.限约数 as 限约,Nvl(P.病案必须,0) as 病案,Nvl(C.项目特性,0) as 急诊," & _
                    "       P.周日 as 日,P.周一 as 一,P.周二 as 二,P.周三 as 三,P.周四 as 四,P.周五 as 五,P.周六 as 六," & _
                    "       Decode(P.分诊方式,1,'指定',2,'动态',3,'平均',NULL) as 分诊,P.序号控制,P.排班" & _
                    " From " & strSQL & "," & vbCrLf & _
                    "           (" & strState & ") A,部门表 B,收费项目目录 C" & _
                    " Where P.ID=A.安排ID(+) And Nvl(B.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And P.科室ID=B.ID And P.项目ID=C.ID" & strIF & strZero & _
                    "           And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & strWhere & _
                    "           And (Nvl(P.医生ID,0)=0 Or Exists(Select 1 From 人员表 Q Where P.医生ID=Q.ID And (Q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.撤档时间 Is Null)))" & _
                    " Order by " & str排序
                    
        Set mrsPlan = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                UserInfo.姓名, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
    Else
        Exit Function
    End If
    If mrsPlan.RecordCount = 0 And mblnAppointment Then
        cboArrangeNo.Clear
        lblDeptName.Caption = ""
        If mblnInit Then MsgBox "当前没有可用的挂号安排，请在挂号安排管理中设置后重试！", vbInformation, gstrSysName
        Exit Function
    End If
    Set mcolArrangeNo = New Collection
    With cboArrangeNo
        .Clear
        Do While Not mrsPlan.EOF
            If Nvl(mrsPlan!医生) = "" Then
                .AddItem "[" & Nvl(mrsPlan!号别) & "]" & Nvl(mrsPlan!项目)
            Else
                .AddItem "[" & Nvl(mrsPlan!号别) & "]" & Nvl(mrsPlan!项目) & "(" & Nvl(mrsPlan!医生) & ")"
            End If
            mcolArrangeNo.Add Nvl(mrsPlan!号别)
            mrsPlan.MoveNext
        Loop
        If .ListCount <> 0 Then
            .ListIndex = 0
        Else
            MsgBox "当前没有可用的挂号安排，请在挂号安排管理中设置后重试！", vbInformation, gstrSysName
            Exit Function
        End If
'        Call GetActiveView
        Call ReadLimit
        Call LoadFeeItem(Val(Nvl(mrsPlan!项目ID)), chkBook.Value = 1)
'        If mblnAppointment Then
'            Select Case mViewMode
'                Case V_普通号分时段, v_专家号分时段
'                    cmdTime.Visible = True
'                Case Else
'                    cmdTime.Visible = False
'            End Select
'            Call InitRegTime
'        Else
'            cmdTime.Visible = False
'        End If

        lblDeptName.Caption = Nvl(mrsPlan!科室)
    End With
    LoadRegPlans = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub ReadLimit()
    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    mrsPlan.Filter = "号别='" & Get号别 & "'"
    If mrsPlan.RecordCount = 0 Then Exit Sub
    If mblnAppointment Then
        If Nvl(mrsPlan!限约) = "" Then
            lblLimit.Caption = "已约:" & Nvl(mrsPlan!已约, 0)
        Else
            lblLimit.Caption = "限约:" & Nvl(mrsPlan!限约) & "  已约:" & Nvl(mrsPlan!已约, 0)
        End If
    Else
        If Nvl(mrsPlan!限号) = "" Then
            lblLimit.Caption = "已挂:" & Nvl(mrsPlan!已挂, 0)
        Else
            lblLimit.Caption = "限号:" & Nvl(mrsPlan!限号) & "  已挂:" & Nvl(mrsPlan!已挂, 0)
        End If
    End If
    If Val(Nvl(mrsPlan!急诊)) = 0 Then
        lbl急.Visible = False
    Else
        lbl急.Visible = True
    End If
End Sub

Private Function Get号别() As String
    If cboArrangeNo.Text = "" Then Exit Function
    Get号别 = Mid(cboArrangeNo.Text, 2, InStr(cboArrangeNo.Text, "]") - 2)
End Function

Private Function GetActiveView()
    '得到当前挂号业务  采取那种类型的流程
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    Dim str号码         As String
    Dim dat            As Date
    
    On Error GoTo errH
    str号码 = Get号别
    If mblnAppointment Then
        dat = dtpDate.Value
    Else
        dat = gobjDatabase.CurrentDate
    End If
    
    strSQL = _
    "       Select   Havedata, 安排id" & vbNewLine & _
    "       From (" & vbNewLine & _
    "               Select 1 As Havedata, b.Id As 安排id " & vbNewLine & _
    "               From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
    "               Where B.号码=[1] And A.安排id = b.ID " & _
    "                And   Decode(To_Char([2], 'D'), '1', '周日', '2'," & _
    "                   '周一', '3', '周二', '4', '周三', '5', '周四', '6','周五', '7', '周六', Null) =a.星期 " & vbNewLine & _
    "                       And Not Exists" & vbNewLine & _
    "                     (Select 1 From 挂号安排计划 C " & vbNewLine & _
    "                         Where c.安排id = b.Id And c.审核时间 Is Not Null And [2] Between " & _
    "                               Nvl(c.生效时间, [2]) And" & _
    "                          Nvl(c.失效时间, To_Date('3000-01-01', 'yyyy-MM-dd')))" & vbNewLine & _
    "               Union All " & vbNewLine & _
    "               Select 1 As Havedata, c.Id As 安排id" & vbNewLine & _
    "               From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C,(" & vbNewLine & _
    "                   SELECT MAX(a.生效时间 ) 生效 FROM 挂号安排计划 a,挂号安排 B  WHERE a.安排Id=b.ID AND b.号码=[1] AND a.审核时间 IS NOT NULL" & vbNewLine & _
    "             And [2] Between nvl(a.生效时间,to_date('1900-01-01','yyyy-mm-dd')) And nvl(a.失效时间,to_date('3000-01-01','yyyy-mm-dd'))" & vbNewLine & _
    "           ) D  " & vbNewLine & _
    "               Where  C.号码=[1] And c.Id = b.安排id And b.Id = a.计划id And b.生效时间=d.生效 And b.审核时间 Is Not Null" & _
    "                    And   Decode(To_Char([2], 'D'), '1', '周日', '2'," & _
    "                   '周一', '3', '周二', '4', '周三', '5', '周四', '6','周五', '7', '周六', Null) =a.星期 " & vbNewLine & _
    "                       And [2] Between Nvl(b.生效时间,[2]) And nvl(b.失效时间,To_Date('3000-01-01', 'yyyy-MM-dd'))) B"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, str号码, dat)
    If rsTmp.RecordCount > 0 And Val(Nvl(mrsPlan!序号控制)) = 1 Then
       '*********************
       '专家号分时段
       '*********************
       mViewMode = v_专家号分时段

    ElseIf rsTmp.RecordCount > 0 And Val(Nvl(mrsPlan!序号控制)) = 0 Then
       '*********************
       '普通号分时段
       '*********************
       mViewMode = V_普通号分时段

    ElseIf Val(Nvl(mrsPlan!序号控制)) = 1 And Nvl(mrsPlan!限号) <> "" Then
       '*********************
       '专家号不分时段
       '*********************
       mViewMode = v_专家号

     Else
       '*********************
       '普通号
       '*********************
       mViewMode = V_普通号

    End If
    Set rsTmp = Nothing
Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
         Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function InitTimePlan() As Boolean
    '**************************************
    '加载时段
    '返回时段是否加载成功或是否有分时段
    '**************************************
     Dim strSQL         As String
     Dim dateCur        As Date
     Dim strNO          As String
     Dim vRect          As RECT
    If Not mblnAppointment Then Exit Function
    strSQL = "Select Distinct a.序号 As ID, To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
            "From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
            "Where a.安排id = b.Id And b.号码 = [1] And" & vbNewLine & _
            " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
            "      Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六'," & vbNewLine & _
            "             Null) = a.星期(+) And Not Exists" & vbNewLine & _
            " (Select Count(1)" & vbNewLine & _
            "       From 挂号序号状态" & vbNewLine & _
            "       Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like a.序号 || '__') Having" & vbNewLine & _
            "        Count(1) - a.限制数量 >= 0) And Not Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From 挂号安排计划 E" & vbNewLine & _
            "       Where e.安排id = b.Id And e.审核时间 Is Not Null And" & vbNewLine & _
            "             [2] Between Nvl(e.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "             Nvl(e.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')))"
    strSQL = strSQL & " Union " & _
            "Select Distinct a.序号 As ID, To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
            "From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C," & vbNewLine & _
            "     (Select Max(a.生效时间) 生效" & vbNewLine & _
            "       From 挂号安排计划 A, 挂号安排 B" & vbNewLine & _
            "       Where a.安排id = b.Id And b.号码 = [1] And a.审核时间 Is Not Null And" & vbNewLine & _
            "             [2] Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "             Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'))) D" & vbNewLine & _
            "Where a.计划id = b.Id And b.安排id = c.Id And c.号码 = [1] And b.生效时间 = d.生效 And b.审核时间 Is Not Null And" & vbNewLine & _
            " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
            "      [2] Between Nvl(b.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
            "      Nvl(b.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And Not Exists" & vbNewLine & _
            " (Select Count(1)" & vbNewLine & _
            "       From 挂号序号状态" & vbNewLine & _
            "       Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like a.序号 || '__') Having" & vbNewLine & _
            "        Count(1) - a.限制数量 >= 0) And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5'," & vbNewLine & _
            "                                           '周四', '6', '周五', '7', '周六', Null) = a.星期(+)" & vbNewLine & _
            "Order By 开始时间"


    dateCur = Format(dtpDate, "yyyy-mm-dd")
    If strSQL = "" Then Exit Function
    strNO = Get号别
    vRect = GetControlRect(dtpTime.hWnd)
    
    On Error GoTo errH
    
    Set mrs时间段 = gobjDatabase.ShowSQLSelect(Me, strSQL, 0, "预约时间选择", False, "", "预约时间选择", _
                                                False, False, True, vRect.Left, vRect.Top - 300, 600, False, True, False, strNO, dateCur)
    If mrs时间段 Is Nothing Then Exit Function
    If mrs时间段.EOF Then Exit Function
    InitTimePlan = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub LoadFeeItem(ByVal lngItemID As Long, ByVal blnBook As Boolean)
    Dim strSQL As String, i As Integer, dblTotal As Double
    Dim rsIncomes As ADODB.Recordset, cur应收 As Currency, cur实收 As Currency
    Dim j As Integer, rsItems As ADODB.Recordset
    If lngItemID = 0 Then Exit Sub
    '性质:1-主挂号费用 2-从项费用 3-病历费
    ReadRegistPrice lngItemID, blnBook, False, mstrFeeType, rsItems, rsIncomes
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = Format(0, "0.00")
    lblPayMoney.Caption = Format(0, "0.00")
    dblTotal = 0
    If rsItems.RecordCount = 0 Then Exit Sub
    rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        With vsfMoney
            .RowData(.Rows - 1) = Nvl(rsItems!项目ID)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = Nvl(rsItems!项目名称)
            rsIncomes.Filter = "项目ID=" & rsItems!项目ID
            cur应收 = 0: cur实收 = 0
            For j = 1 To rsIncomes.RecordCount
                cur应收 = cur应收 + rsIncomes!应收
                cur实收 = cur实收 + rsIncomes!实收
                rsIncomes.MoveNext
            Next j
            .TextMatrix(.Rows - 1, .ColIndex("应收金额")) = Format(cur应收, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("实收金额")) = Format(cur实收, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("性质")) = Nvl(rsItems!性质)
            .Rows = .Rows + 1
        End With
        rsItems.MoveNext
    Next i
    If vsfMoney.Rows > 2 Then vsfMoney.Rows = vsfMoney.Rows - 1
    For i = 1 To vsfMoney.Rows - 1
        dblTotal = dblTotal + Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("实收金额")))
    Next i
    vsfMoney.RowHeightMin = 350
    lblTotal.Caption = Format(dblTotal, "0.00")
    lblPayMoney.Caption = Format(dblTotal, "0.00")
    lblRoomName.Caption = gstrRooms
End Sub


Private Function GetSNState(str号别 As String, datThis As Date, Optional lngSN As Long) As ADODB.Recordset
    Dim strSQL           As String
    Dim datStart         As Date
    Dim datEnd           As Date
    On Error GoTo errH
    datStart = CDate(Format(datThis, "yyyy-MM-dd"))
    datEnd = DateAdd("s", -1, DateAdd("d", 1, datStart))
    strSQL = "    " & vbNewLine & " Select 序号,状态,操作员姓名,Nvl(预约,0) as 预约,TO_Char(日期,'hh24:mi:ss') as 日期  "
    strSQL = strSQL & vbNewLine & " From 挂号序号状态 "
    strSQL = strSQL & vbNewLine & " Where 号码=[1]"
    strSQL = strSQL & vbNewLine & IIf(datThis = CDate(0), " And 日期 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 ", " And 日期 Between  [2] And [3]")
    strSQL = strSQL & vbNewLine & IIf(lngSN > 0, " And 序号=[4]", "")
    Set GetSNState = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, str号别, datStart, datEnd, lngSN)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function zlGet当前星期几(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当日是星期几
    '编制:刘兴洪
    '日期:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bln当前日期 As Boolean, strTemp As String
    If strDate = "" Then
        strSQL = "Select Decode(To_Char(Sysdate,'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六',NULL) as 星期  From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "Select Decode(To_Char([1],'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六','') As 星期 From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strDate))
    End If
    
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!星期)
    zlGet当前星期几 = strTemp
End Function

