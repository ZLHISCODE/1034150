VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#2.1#0"; "ZlPatiAddress.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPageMedRecEdit_YN 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "病案首页"
   ClientHeight    =   60000
   ClientLeft      =   165
   ClientTop       =   1605
   ClientWidth     =   16005
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPageMedRecEdit_YN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   60000
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.VScrollBar vsbMain 
      Height          =   7335
      LargeChange     =   100
      Left            =   0
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   427
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdTop 
      Appearance      =   0  'Flat
      Height          =   500
      Left            =   0
      Picture         =   "frmPageMedRecEdit_YN.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   426
      ToolTipText     =   "回顶部"
      Top             =   1000
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.HScrollBar hsbMain 
      Height          =   255
      LargeChange     =   25
      Left            =   1000
      Max             =   100
      TabIndex        =   451
      Top             =   0
      Width           =   7935
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   58010
      Left            =   600
      ScaleHeight     =   57975
      ScaleWidth      =   12975
      TabIndex        =   428
      Top             =   300
      Width           =   13000
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6855
         Index           =   1
         Left            =   738
         ScaleHeight     =   6855
         ScaleWidth      =   11505
         TabIndex        =   430
         Tag             =   "true"
         Top             =   2535
         Width           =   11500
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   52
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   39
            Top             =   840
            Width           =   210
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   42
            Left            =   7560
            TabIndex        =   452
            TabStop         =   0   'False
            Top             =   4793
            Width           =   285
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   7230
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   4393
            Width           =   285
         End
         Begin VB.CommandButton cmdDateInfo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   2
            Left            =   2995
            Picture         =   "frmPageMedRecEdit_YN.frx":079B
            Style           =   1  'Graphical
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   4380
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   255
            Index           =   2
            Left            =   1215
            TabIndex        =   100
            Tag             =   "####-##-## ##:##"
            Top             =   4380
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   0
            Left            =   6460
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1480
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   5
            Left            =   9225
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   3880
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   2
            Left            =   6460
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   2280
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   3
            Left            =   6460
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   2680
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   1
            Left            =   9225
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1480
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   4
            Left            =   6460
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   3880
            Width           =   270
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   255
            Index           =   4
            Left            =   1215
            TabIndex        =   93
            Top             =   3885
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   255
            Index           =   3
            Left            =   1215
            TabIndex        =   72
            Top             =   2685
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   255
            Index           =   2
            Left            =   1215
            TabIndex        =   64
            Top             =   2280
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   255
            Index           =   1
            Left            =   7485
            TabIndex        =   50
            Top             =   1485
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Items           =   2
            Style           =   1
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   250
            Index           =   0
            Left            =   1215
            TabIndex        =   46
            Top             =   1480
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Items           =   3
            Style           =   1
            MaxLength       =   100
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   255
            Index           =   0
            Left            =   5400
            TabIndex        =   22
            Tag             =   "####-##-## ##:##"
            Top             =   180
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "再入院"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   0
            Left            =   6240
            TabIndex        =   30
            Top             =   565
            Width           =   960
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   6
            Left            =   6460
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   3080
            Width           =   270
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   2
            Left            =   1210
            MaxLength       =   30
            TabIndex        =   101
            Top             =   4380
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   0
            Left            =   5400
            MaxLength       =   30
            TabIndex        =   23
            Top             =   180
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   4
            Left            =   1210
            MaxLength       =   100
            TabIndex        =   94
            Tag             =   "50"
            ToolTipText     =   "按*键显示地区列表"
            Top             =   3880
            Width           =   5520
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   3
            Left            =   1210
            MaxLength       =   100
            TabIndex        =   73
            Tag             =   "50"
            ToolTipText     =   "按*键显示地区列表"
            Top             =   2680
            Width           =   5520
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   2
            Left            =   1210
            MaxLength       =   100
            TabIndex        =   65
            ToolTipText     =   "按*键显示地区列表"
            Top             =   2280
            Width           =   5520
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   1
            Left            =   7485
            MaxLength       =   100
            TabIndex        =   51
            ToolTipText     =   "按*键显示地区列表"
            Top             =   1480
            Width           =   2010
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   61
            ItemData        =   "frmPageMedRecEdit_YN.frx":0891
            Left            =   1210
            List            =   "frmPageMedRecEdit_YN.frx":0893
            TabIndex        =   56
            Top             =   1840
            Width           =   2340
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   5
            Left            =   7485
            MaxLength       =   30
            TabIndex        =   97
            ToolTipText     =   "按*键显示区域列表"
            Top             =   3880
            Width           =   2010
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   0
            Left            =   1210
            MaxLength       =   100
            TabIndex        =   47
            Tag             =   "30"
            ToolTipText     =   "按*键显示地区列表"
            Top             =   1480
            Width           =   5520
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   5
            ItemData        =   "frmPageMedRecEdit_YN.frx":0895
            Left            =   10215
            List            =   "frmPageMedRecEdit_YN.frx":0897
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1440
            Width           =   1150
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   3
            ItemData        =   "frmPageMedRecEdit_YN.frx":0899
            Left            =   7485
            List            =   "frmPageMedRecEdit_YN.frx":089B
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   1845
            Width           =   2010
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   7
            ItemData        =   "frmPageMedRecEdit_YN.frx":089D
            Left            =   1210
            List            =   "frmPageMedRecEdit_YN.frx":08AA
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   4740
            Width           =   2175
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   8
            Left            =   9000
            MaxLength       =   100
            TabIndex        =   107
            Tag             =   "18"
            Top             =   4380
            Width           =   2100
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   3
            Left            =   1210
            MaxLength       =   64
            TabIndex        =   18
            Tag             =   "20"
            Top             =   180
            Width           =   1260
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   1
            Left            =   7485
            MaxLength       =   20
            TabIndex        =   81
            Tag             =   "20"
            Top             =   3080
            Width           =   2010
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   4
            Left            =   10215
            MaxLength       =   6
            TabIndex        =   70
            Top             =   2280
            Width           =   1150
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   2
            Left            =   10215
            MaxLength       =   6
            TabIndex        =   83
            Tag             =   "6"
            Top             =   3080
            Width           =   1150
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   5
            Left            =   10215
            MaxLength       =   6
            TabIndex        =   76
            Tag             =   "6"
            Top             =   2680
            Width           =   1150
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   14
            Left            =   7485
            MaxLength       =   20
            TabIndex        =   91
            Tag             =   "20"
            Top             =   3480
            Width           =   2010
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   7
            Left            =   4935
            TabIndex        =   104
            Tag             =   "18"
            ToolTipText     =   "按*键显示地区列表"
            Top             =   4380
            Width           =   2580
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   6
            Left            =   1210
            MaxLength       =   64
            TabIndex        =   85
            Tag             =   "10"
            Top             =   3480
            Width           =   1260
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   6
            Left            =   1210
            MaxLength       =   100
            TabIndex        =   78
            Tag             =   "50"
            ToolTipText     =   "按*键显示合约单位列表"
            Top             =   3080
            Width           =   5520
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   3
            Left            =   7485
            MaxLength       =   20
            TabIndex        =   68
            Top             =   2280
            Width           =   2010
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   4
            ItemData        =   "frmPageMedRecEdit_YN.frx":08C0
            Left            =   10215
            List            =   "frmPageMedRecEdit_YN.frx":08C2
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   135
            Width           =   1150
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "入院前是否经外院治疗"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   1
            Left            =   8535
            TabIndex        =   112
            Top             =   4765
            Width           =   2520
         End
         Begin VB.ComboBox cboSpecificInfo 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   16
            ItemData        =   "frmPageMedRecEdit_YN.frx":08C4
            Left            =   4200
            List            =   "frmPageMedRecEdit_YN.frx":08C6
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   930
            Width           =   800
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   16
            Left            =   3000
            MaxLength       =   20
            TabIndex        =   38
            Tag             =   "年龄"
            Top             =   960
            Width           =   360
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   17
            Left            =   6960
            MaxLength       =   25
            TabIndex        =   42
            Tag             =   "新生儿出生体重    "
            Top             =   960
            Width           =   1050
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   18
            Left            =   10080
            MaxLength       =   25
            TabIndex        =   44
            Tag             =   "新生儿入院体重"
            Top             =   960
            Width           =   1050
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   10215
            MaxLength       =   5
            TabIndex        =   35
            Top             =   580
            Width           =   795
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   8220
            MaxLength       =   5
            TabIndex        =   32
            Top             =   580
            Width           =   690
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   4
            Left            =   4545
            MaxLength       =   20
            TabIndex        =   58
            Tag             =   "18"
            Top             =   1885
            Width           =   2175
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   6
            ItemData        =   "frmPageMedRecEdit_YN.frx":08C8
            Left            =   3240
            List            =   "frmPageMedRecEdit_YN.frx":08CA
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   3440
            Width           =   2030
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   2
            ItemData        =   "frmPageMedRecEdit_YN.frx":08CC
            Left            =   10215
            List            =   "frmPageMedRecEdit_YN.frx":08CE
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   1840
            Width           =   1150
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   1
            ItemData        =   "frmPageMedRecEdit_YN.frx":08D0
            Left            =   3240
            List            =   "frmPageMedRecEdit_YN.frx":08D2
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   140
            Width           =   1005
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   15
            Left            =   8220
            MaxLength       =   20
            TabIndex        =   26
            Tag             =   "5"
            Top             =   180
            Width           =   450
         End
         Begin VB.ComboBox cboSpecificInfo 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   15
            ItemData        =   "frmPageMedRecEdit_YN.frx":08D4
            Left            =   8715
            List            =   "frmPageMedRecEdit_YN.frx":08D6
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   160
            Width           =   800
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   42
            Left            =   4920
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   4780
            Width           =   2625
         End
         Begin VB.CommandButton cmdDateInfo 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   0
            Left            =   7215
            Picture         =   "frmPageMedRecEdit_YN.frx":08D8
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   180
            Width           =   285
         End
         Begin VB.PictureBox PicOut 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   0
            ScaleHeight     =   855
            ScaleWidth      =   11505
            TabIndex        =   115
            Top             =   6000
            Width           =   11500
            Begin VB.CommandButton cmdInfo 
               Appearance      =   0  'Flat
               Caption         =   "…"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   9
               Left            =   7245
               TabIndex        =   122
               TabStop         =   0   'False
               Top             =   0
               Width           =   270
            End
            Begin VB.CommandButton cmdDateInfo 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   3
               Left            =   2995
               Picture         =   "frmPageMedRecEdit_YN.frx":09CE
               Style           =   1  'Graphical
               TabIndex        =   119
               TabStop         =   0   'False
               Top             =   0
               Width           =   270
            End
            Begin MSMask.MaskEdBox mskDateInfo 
               Height          =   255
               Index           =   3
               Left            =   1215
               TabIndex        =   117
               Tag             =   "####-##-## ##:##"
               Top             =   0
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   450
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   16
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "####-##-## ##:##"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   10
               Left            =   9030
               MaxLength       =   100
               TabIndex        =   124
               Tag             =   "18"
               Top             =   0
               Width           =   2100
            End
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00ECFFCC&
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   19
               Left            =   1210
               Locked          =   -1  'True
               TabIndex        =   126
               TabStop         =   0   'False
               Tag             =   "4"
               Top             =   400
               Width           =   2055
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   11
               Left            =   4935
               MaxLength       =   20
               TabIndex        =   128
               Top             =   400
               Width           =   2580
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   9
               Left            =   4935
               TabIndex        =   121
               Tag             =   "18"
               ToolTipText     =   "按*键显示地区列表"
               Top             =   0
               Width           =   2580
            End
            Begin VB.TextBox txtDateInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   3
               Left            =   1210
               MaxLength       =   30
               TabIndex        =   118
               Top             =   0
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "病房"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   10
               Left            =   8550
               TabIndex        =   123
               Top             =   20
               Width           =   420
            End
            Begin VB.Label lblSpecificInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院天数"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   19
               Left            =   310
               TabIndex        =   125
               Top             =   420
               Width           =   840
            End
            Begin VB.Label lblDateInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "出院时间"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   310
               TabIndex        =   116
               Top             =   20
               Width           =   840
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "出院科室"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   9
               Left            =   4035
               TabIndex        =   120
               Top             =   20
               Width           =   840
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "医保号"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   11
               Left            =   4245
               TabIndex        =   127
               Top             =   420
               Width           =   630
            End
         End
         Begin VB.PictureBox picRelation 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   5280
            ScaleHeight     =   240
            ScaleMode       =   0  'User
            ScaleWidth      =   1430.103
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   3480
            Visible         =   0   'False
            Width           =   1435
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   41
               Left            =   0
               MaxLength       =   100
               TabIndex        =   89
               Top             =   0
               Width           =   1445
            End
            Begin VB.Line lineRelation 
               X1              =   0
               X2              =   1440.034
               Y1              =   225
               Y2              =   225
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTransfer 
            Height          =   710
            Left            =   1215
            TabIndex        =   114
            Top             =   5205
            Width           =   9975
            _cx             =   17595
            _cy             =   1252
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   1000
            ColWidthMax     =   1500
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmPageMedRecEdit_YN.frx":0AC4
            ScrollTrack     =   0   'False
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            Editable        =   2
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
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            Caption         =   "30"
            Height          =   180
            Index           =   52
            Left            =   3720
            TabIndex        =   453
            Top             =   1125
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Line lineH 
            Index           =   0
            X1              =   0
            X2              =   14200
            Y1              =   1350
            Y2              =   1350
         End
         Begin VB.Line lineH 
            Index           =   1
            X1              =   0
            X2              =   14200
            Y1              =   4250
            Y2              =   4250
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院途径"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   7
            Left            =   310
            TabIndex        =   108
            Top             =   4800
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "邮编"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   9735
            TabIndex        =   69
            Top             =   2295
            Width           =   420
         End
         Begin VB.Label lblAdressInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "现住址"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   520
            TabIndex        =   63
            Top             =   2300
            Width           =   630
         End
         Begin VB.Label lblAdressInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "籍贯"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   7005
            TabIndex        =   49
            Top             =   1500
            Width           =   420
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(年龄不足一周岁的)年龄"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   16
            Left            =   600
            TabIndex        =   37
            Top             =   980
            Width           =   2310
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "新生儿出生体重             克"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   17
            Left            =   5280
            TabIndex        =   41
            Top             =   980
            Width           =   3045
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "新生儿入院体重            克"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   18
            Left            =   8520
            TabIndex        =   43
            Top             =   980
            Width           =   2940
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "kg"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   9
            Left            =   11130
            TabIndex        =   36
            Top             =   600
            Width           =   210
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "体重"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   8
            Left            =   9735
            TabIndex        =   34
            Top             =   600
            Width           =   420
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "cm"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   7
            Left            =   8970
            TabIndex        =   33
            Top             =   600
            Width           =   210
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "身高"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   7740
            TabIndex        =   31
            Top             =   600
            Width           =   420
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "其它证件"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   3645
            TabIndex        =   57
            Top             =   1905
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   7005
            TabIndex        =   67
            Top             =   2295
            Width           =   420
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院科室"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   7
            Left            =   4035
            TabIndex        =   103
            Top             =   4400
            Width           =   840
         End
         Begin VB.Label lblDateInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院时间"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   310
            TabIndex        =   99
            Top             =   4400
            Width           =   840
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病房"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   8
            Left            =   8550
            TabIndex        =   106
            Top             =   4400
            Width           =   420
         End
         Begin VB.Label lblAdressInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "出生地"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   520
            TabIndex        =   45
            Top             =   1500
            Width           =   630
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "身份证"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   520
            TabIndex        =   55
            Top             =   1900
            Width           =   630
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "职业"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   7005
            TabIndex        =   59
            Top             =   1905
            Width           =   420
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "年龄"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   15
            Left            =   7740
            TabIndex        =   25
            Top             =   195
            Width           =   420
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "联系人姓名"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   100
            TabIndex        =   84
            Top             =   3500
            Width           =   1050
         End
         Begin VB.Label lblAdressInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "联系人地址"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   100
            TabIndex        =   92
            Top             =   3900
            Width           =   1050
         End
         Begin VB.Label lblAdressInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "户口地址"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   310
            TabIndex        =   71
            Top             =   2700
            Width           =   840
         End
         Begin VB.Label lblAdressInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "工作单位"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   310
            TabIndex        =   77
            Top             =   3100
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "国籍"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   9735
            TabIndex        =   28
            Top             =   195
            Width           =   420
         End
         Begin VB.Label lblAdressInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "区域"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   7005
            TabIndex        =   96
            Top             =   3900
            Width           =   420
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "邮编"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   9735
            TabIndex        =   82
            Top             =   3105
            Width           =   420
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "邮编"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   9735
            TabIndex        =   75
            Top             =   2700
            Width           =   420
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   7005
            TabIndex        =   80
            Top             =   3105
            Width           =   420
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   14
            Left            =   7005
            TabIndex        =   90
            Top             =   3495
            Width           =   420
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "关系"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   2760
            TabIndex        =   86
            Top             =   3495
            Width           =   420
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "姓名"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   730
            TabIndex        =   17
            Top             =   200
            Width           =   420
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "民族"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   9735
            TabIndex        =   53
            Top             =   1500
            Width           =   420
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "婚姻"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   9735
            TabIndex        =   61
            Top             =   1905
            Width           =   420
         End
         Begin VB.Label lblDateInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "出生日期"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   4500
            TabIndex        =   21
            Top             =   195
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2760
            TabIndex        =   19
            Top             =   195
            Width           =   420
         End
         Begin VB.Label lblTansfer 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "转科情况"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   310
            TabIndex        =   113
            Top             =   5200
            Width           =   840
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "转入"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   42
            Left            =   4455
            TabIndex        =   110
            Top             =   4800
            Width           =   420
         End
      End
      Begin VB.PictureBox picInfectInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   840
         ScaleHeight     =   3465
         ScaleWidth      =   4065
         TabIndex        =   131
         Top             =   3000
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ListBox lstInfectParts 
            Appearance      =   0  'Flat
            Height          =   2190
            ItemData        =   "frmPageMedRecEdit_YN.frx":0B7A
            Left            =   240
            List            =   "frmPageMedRecEdit_YN.frx":0B81
            Style           =   1  'Checkbox
            TabIndex        =   135
            Top             =   840
            Width           =   3615
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   10
            ItemData        =   "frmPageMedRecEdit_YN.frx":0B95
            Left            =   1920
            List            =   "frmPageMedRecEdit_YN.frx":0B97
            Style           =   2  'Dropdown List
            TabIndex        =   133
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "感染与死亡的关系"
            Height          =   210
            Index           =   8
            Left            =   120
            TabIndex        =   132
            Top             =   165
            Width           =   1680
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "感染部位"
            Height          =   210
            Index           =   26
            Left            =   120
            TabIndex        =   134
            Top             =   480
            Width           =   840
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   8355
         Index           =   18
         Left            =   738
         ScaleHeight     =   8355
         ScaleWidth      =   11505
         TabIndex        =   447
         Tag             =   "true"
         Top             =   49155
         Width           =   11500
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   250
            Index           =   32
            Left            =   6975
            TabIndex        =   375
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   480
            Width           =   270
         End
         Begin VB.CommandButton cmdLastDiag 
            Appearance      =   0  'Flat
            Caption         =   "获取上次住院主要诊断"
            Height          =   330
            Left            =   705
            TabIndex        =   424
            TabStop         =   0   'False
            Top             =   7515
            Width           =   3300
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "上一次住本院与本次住院是因同一疾病(主要诊断)"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   23
            Left            =   4920
            TabIndex        =   423
            Tag             =   "0"
            Top             =   7005
            Width           =   5505
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   50
            ItemData        =   "frmPageMedRecEdit_YN.frx":0B99
            Left            =   2520
            List            =   "frmPageMedRecEdit_YN.frx":0B9B
            Style           =   2  'Dropdown List
            TabIndex        =   422
            Top             =   6960
            Width           =   1485
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   51
            ItemData        =   "frmPageMedRecEdit_YN.frx":0B9D
            Left            =   8340
            List            =   "frmPageMedRecEdit_YN.frx":0B9F
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   379
            TabStop         =   0   'False
            Top             =   840
            Width           =   2445
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "非预期的重返重症医学科"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   25
            Left            =   2880
            TabIndex        =   377
            TabStop         =   0   'False
            Top             =   865
            Width           =   2850
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "发生人工气道脱出 "
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   24
            Left            =   8340
            TabIndex        =   376
            TabStop         =   0   'False
            Top             =   465
            Width           =   2130
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   32
            Left            =   4800
            MaxLength       =   100
            TabIndex        =   374
            Top             =   480
            Width           =   2445
         End
         Begin VB.PictureBox PicAdvEvent 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3255
            Left            =   5760
            ScaleHeight     =   3225
            ScaleWidth      =   5670
            TabIndex        =   411
            TabStop         =   0   'False
            Top             =   3480
            Width           =   5700
            Begin VB.ListBox lstAdvEvent 
               Appearance      =   0  'Flat
               Height          =   1710
               ItemData        =   "frmPageMedRecEdit_YN.frx":0BA1
               Left            =   -15
               List            =   "frmPageMedRecEdit_YN.frx":0BA8
               Style           =   1  'Checkbox
               TabIndex        =   412
               Top             =   -15
               Width           =   5700
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   46
               ItemData        =   "frmPageMedRecEdit_YN.frx":0BB9
               Left            =   3795
               List            =   "frmPageMedRecEdit_YN.frx":0BBB
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   416
               TabStop         =   0   'False
               Top             =   1875
               Width           =   1620
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   48
               ItemData        =   "frmPageMedRecEdit_YN.frx":0BBD
               Left            =   1575
               List            =   "frmPageMedRecEdit_YN.frx":0BBF
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   420
               TabStop         =   0   'False
               Top             =   2775
               Width           =   3840
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   47
               ItemData        =   "frmPageMedRecEdit_YN.frx":0BC1
               Left            =   1575
               List            =   "frmPageMedRecEdit_YN.frx":0BC3
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   418
               TabStop         =   0   'False
               Top             =   2325
               Width           =   3840
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   45
               ItemData        =   "frmPageMedRecEdit_YN.frx":0BC5
               Left            =   1575
               List            =   "frmPageMedRecEdit_YN.frx":0BC7
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   414
               TabStop         =   0   'False
               Top             =   1875
               Width           =   1560
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "压疮发生期间"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   45
               Left            =   420
               TabIndex        =   413
               Top             =   1935
               Width           =   1080
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "跌倒或坠床原因"
               Height          =   180
               Index           =   48
               Left            =   240
               TabIndex        =   419
               Top             =   2835
               Width           =   1260
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "跌倒或坠床伤害"
               Height          =   180
               Index           =   47
               Left            =   240
               TabIndex        =   417
               Top             =   2385
               Width           =   1260
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分期"
               Height          =   180
               Index           =   46
               Left            =   3360
               TabIndex        =   415
               Top             =   1935
               Width           =   360
            End
         End
         Begin VB.ListBox lstInfection 
            Appearance      =   0  'Flat
            Height          =   1470
            ItemData        =   "frmPageMedRecEdit_YN.frx":0BC9
            Left            =   0
            List            =   "frmPageMedRecEdit_YN.frx":0BD0
            Style           =   1  'Checkbox
            TabIndex        =   404
            Top             =   3480
            Width           =   5500
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "3.彩色多普勒"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   15
            Left            =   2910
            TabIndex        =   408
            Top             =   5295
            Width           =   1875
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "2.MRI"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   14
            Left            =   1620
            TabIndex        =   407
            Top             =   5295
            Width           =   1035
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "1.CT"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   13
            Left            =   420
            TabIndex        =   406
            Top             =   5295
            Width           =   795
         End
         Begin VB.PictureBox PicRestrain 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   0
            ScaleHeight     =   1545
            ScaleWidth      =   5475
            TabIndex        =   381
            Top             =   1560
            Width           =   5500
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "住院期间使用物理约束"
               ForeColor       =   &H80000008&
               Height          =   280
               Index           =   26
               Left            =   120
               TabIndex        =   382
               Top             =   300
               Width           =   2535
            End
            Begin VB.ComboBox cboBaseInfo 
               BackColor       =   &H8000000F&
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   54
               ItemData        =   "frmPageMedRecEdit_YN.frx":0BE2
               Left            =   3720
               List            =   "frmPageMedRecEdit_YN.frx":0BE4
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   390
               TabStop         =   0   'False
               Top             =   1140
               Width           =   1605
            End
            Begin VB.ComboBox cboBaseInfo 
               BackColor       =   &H8000000F&
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   53
               ItemData        =   "frmPageMedRecEdit_YN.frx":0BE6
               Left            =   3720
               List            =   "frmPageMedRecEdit_YN.frx":0BE8
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   388
               TabStop         =   0   'False
               Top             =   765
               Width           =   1605
            End
            Begin VB.ComboBox cboBaseInfo 
               BackColor       =   &H8000000F&
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   52
               ItemData        =   "frmPageMedRecEdit_YN.frx":0BEA
               Left            =   3720
               List            =   "frmPageMedRecEdit_YN.frx":0BEC
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   386
               TabStop         =   0   'False
               Top             =   375
               Width           =   1605
            End
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   43
               Left            =   3960
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   384
               TabStop         =   0   'False
               Top             =   40
               Width           =   615
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "约束原因"
               Height          =   210
               Index           =   54
               Left            =   2760
               TabIndex        =   389
               Top             =   1200
               Width           =   840
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "约束工具"
               Height          =   210
               Index           =   53
               Left            =   2760
               TabIndex        =   387
               Top             =   825
               Width           =   840
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "约束方式"
               Height          =   210
               Index           =   52
               Left            =   2760
               TabIndex        =   385
               Top             =   420
               Width           =   840
            End
            Begin VB.Label lblSpecificInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "约束总时间        小时"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   43
               Left            =   2760
               TabIndex        =   383
               Top             =   60
               Width           =   2310
            End
         End
         Begin VB.PictureBox PicPath 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   5760
            ScaleHeight     =   1545
            ScaleWidth      =   5670
            TabIndex        =   392
            Top             =   1560
            Width           =   5700
            Begin VB.Frame fraCbo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   250
               Index           =   0
               Left            =   3360
               TabIndex        =   401
               Top             =   975
               Visible         =   0   'False
               Width           =   1965
               Begin VB.ComboBox cboBaseInfo 
                  Height          =   330
                  IMEMode         =   3  'DISABLE
                  Index           =   62
                  ItemData        =   "frmPageMedRecEdit_YN.frx":0BEE
                  Left            =   -30
                  List            =   "frmPageMedRecEdit_YN.frx":0BF0
                  Style           =   2  'Dropdown List
                  TabIndex        =   402
                  Top             =   -25
                  Width           =   2035
               End
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "进入路径"
               ForeColor       =   &H80000008&
               Height          =   280
               Index           =   19
               Left            =   120
               TabIndex        =   393
               Top             =   140
               Width           =   1260
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "完成路径"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   280
               Index           =   20
               Left            =   960
               TabIndex        =   395
               TabStop         =   0   'False
               Top             =   550
               Width           =   1260
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "变异"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   280
               Index           =   21
               Left            =   960
               TabIndex        =   398
               TabStop         =   0   'False
               Top             =   955
               Width           =   900
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   30
               Left            =   3360
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   397
               TabStop         =   0   'False
               Top             =   565
               Width           =   1965
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   31
               Left            =   3360
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   400
               TabStop         =   0   'False
               Top             =   970
               Width           =   1965
            End
            Begin VB.CommandButton cmdAutoLoad 
               Caption         =   "自动提取"
               Height          =   330
               Index           =   3
               Left            =   1920
               TabIndex        =   394
               TabStop         =   0   'False
               Top             =   105
               Width           =   1200
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "退出原因"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   30
               Left            =   2460
               TabIndex        =   396
               Top             =   585
               Width           =   840
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "变异原因"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   31
               Left            =   2460
               TabIndex        =   399
               Top             =   990
               Width           =   840
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTSJC 
            Height          =   1095
            Left            =   60
            TabIndex        =   409
            Top             =   5640
            Width           =   5500
            _cx             =   9701
            _cy             =   1931
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmPageMedRecEdit_YN.frx":0BF2
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            Editable        =   2
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
         Begin VB.Line lineH 
            Index           =   18
            X1              =   0
            X2              =   13080
            Y1              =   6840
            Y2              =   6840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "距上一次住本院的时间"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   50
            Left            =   360
            TabIndex        =   421
            Top             =   7020
            Width           =   2100
         End
         Begin VB.Label lblDiagInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "上次诊断"
            ForeColor       =   &H80000008&
            Height          =   690
            Left            =   4920
            TabIndex        =   425
            Top             =   7515
            Width           =   6195
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "重返间隔时间"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   51
            Left            =   7020
            TabIndex        =   378
            Top             =   900
            Width           =   1260
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "重症监护室名称"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   32
            Left            =   3240
            TabIndex        =   373
            Top             =   500
            Width           =   1470
         End
         Begin VB.Line lineH 
            Index           =   14
            X1              =   0
            X2              =   13080
            Y1              =   350
            Y2              =   350
         End
         Begin VB.Label lblVsTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入住重症监护室（ICU）情况："
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   240
            TabIndex        =   372
            Top             =   500
            Width           =   2835
         End
         Begin VB.Line lineCheck 
            X1              =   1320
            X2              =   5520
            Y1              =   5160
            Y2              =   5160
         End
         Begin VB.Label lblAdvEvent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "不良事件"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5880
            TabIndex        =   410
            Top             =   3240
            Width           =   840
         End
         Begin VB.Line lineAdvEvent 
            X1              =   6720
            X2              =   11985
            Y1              =   3360
            Y2              =   3330
         End
         Begin VB.Line lineInfection 
            X1              =   840
            X2              =   5520
            Y1              =   3330
            Y2              =   3330
         End
         Begin VB.Label lblInfection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "感染因素"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   403
            Top             =   3240
            Width           =   840
         End
         Begin VB.Label lblVsTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "特殊检查情况"
            Height          =   210
            Index           =   5
            Left            =   0
            TabIndex        =   405
            Top             =   5040
            Width           =   1260
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "附页"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   18
            Left            =   0
            TabIndex        =   371
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblRestrain 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院期间物理约束"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   380
            Top             =   1320
            Width           =   1680
         End
         Begin VB.Label lblPath 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "临床路径"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5760
            TabIndex        =   391
            Top             =   1320
            Width           =   840
         End
         Begin VB.Line lineRestrain 
            X1              =   1680
            X2              =   5520
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line linePath 
            X1              =   6600
            X2              =   11760
            Y1              =   1410
            Y2              =   1410
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4040
         Index           =   3
         Left            =   738
         ScaleHeight     =   4035
         ScaleWidth      =   11505
         TabIndex        =   432
         Tag             =   "true"
         Top             =   13050
         Width           =   11500
         Begin VB.CommandButton cmdDateInfo 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   4
            Left            =   10890
            Picture         =   "frmPageMedRecEdit_YN.frx":0C60
            Style           =   1  'Graphical
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   320
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   255
            Index           =   4
            Left            =   9045
            TabIndex        =   144
            Tag             =   "####-##-## ##:##"
            Top             =   315
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   250
            Index           =   5
            Left            =   1305
            TabIndex        =   175
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##:##"
            Top             =   3220
            Visible         =   0   'False
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   20
            Left            =   11055
            TabIndex        =   181
            TabStop         =   0   'False
            Top             =   3220
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   250
            Index           =   22
            Left            =   11055
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   2720
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   21
            Left            =   11055
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   1740
            Width           =   270
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   330
            Index           =   20
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D56
            Left            =   4455
            List            =   "frmPageMedRecEdit_YN.frx":0D58
            Style           =   2  'Dropdown List
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   3180
            Width           =   1515
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   22
            Left            =   3375
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   170
            TabStop         =   0   'False
            Top             =   2720
            Width           =   615
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   21
            Left            =   1620
            MaxLength       =   5
            TabIndex        =   168
            Top             =   2720
            Width           =   555
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   15
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D5A
            Left            =   9825
            List            =   "frmPageMedRecEdit_YN.frx":0D5C
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   1180
            Width           =   1515
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   14
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D5E
            Left            =   6105
            List            =   "frmPageMedRecEdit_YN.frx":0D60
            Style           =   2  'Dropdown List
            TabIndex        =   154
            Top             =   1180
            Width           =   1515
         End
         Begin VB.CommandButton cmdDeliceryInfo 
            Caption         =   "分娩信息(&Z)"
            Height          =   330
            Left            =   10005
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1320
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   12
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D62
            Left            =   2385
            List            =   "frmPageMedRecEdit_YN.frx":0D64
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   780
            Width           =   1515
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   13
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D66
            Left            =   6105
            List            =   "frmPageMedRecEdit_YN.frx":0D68
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   780
            Width           =   1515
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   330
            Index           =   21
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D6A
            Left            =   4455
            List            =   "frmPageMedRecEdit_YN.frx":0D6C
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   185
            TabStop         =   0   'False
            Top             =   3580
            Width           =   1515
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   19
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D6E
            Left            =   6105
            List            =   "frmPageMedRecEdit_YN.frx":0D70
            Style           =   2  'Dropdown List
            TabIndex        =   164
            Top             =   2180
            Width           =   1515
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   18
            IntegralHeight  =   0   'False
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D72
            Left            =   9825
            List            =   "frmPageMedRecEdit_YN.frx":0D74
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   2180
            Width           =   1515
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   16
            Left            =   2385
            Style           =   2  'Dropdown List
            TabIndex        =   152
            Top             =   1180
            Width           =   1515
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "新发肿瘤"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   5
            Left            =   4440
            TabIndex        =   141
            Top             =   305
            Width           =   1215
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   11
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D76
            Left            =   2385
            List            =   "frmPageMedRecEdit_YN.frx":0D78
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   280
            Width           =   1515
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "医院感染作病原学检查"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   3
            Left            =   1275
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   1725
            Width           =   2505
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "确诊"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   2
            Left            =   6240
            TabIndex        =   142
            Top             =   305
            Value           =   1  'Checked
            Width           =   795
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   19
            Left            =   2385
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   2220
            Width           =   1515
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   20
            Left            =   7500
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   3220
            Width           =   3825
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   21
            Left            =   6105
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   1740
            Width           =   5220
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   22
            Left            =   6105
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   172
            TabStop         =   0   'False
            Top             =   2720
            Width           =   5220
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   4
            Left            =   9045
            MaxLength       =   30
            TabIndex        =   145
            Top             =   320
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   5
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   176
            TabStop         =   0   'False
            Top             =   3225
            Width           =   1845
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   330
            Index           =   60
            ItemData        =   "frmPageMedRecEdit_YN.frx":0D7A
            Left            =   1560
            List            =   "frmPageMedRecEdit_YN.frx":0D7C
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   3580
            Width           =   915
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "死亡患者尸检"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   60
            Left            =   240
            TabIndex        =   182
            Top             =   3640
            Width           =   1260
         End
         Begin VB.Label lblDateInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "死亡时间"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   360
            TabIndex        =   174
            Top             =   3240
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "成功次数"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   22
            Left            =   2475
            TabIndex        =   169
            Top             =   2740
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "抢救次数"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   21
            Left            =   720
            TabIndex        =   167
            Top             =   2740
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与出院"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   14
            Left            =   4995
            TabIndex        =   153
            Top             =   1240
            Width           =   1050
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院与出院"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   15
            Left            =   8715
            TabIndex        =   155
            Top             =   1240
            Width           =   1050
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病理号"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   19
            Left            =   1695
            TabIndex        =   161
            Top             =   2240
            Width           =   630
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "分化程度"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   12
            Left            =   1485
            TabIndex        =   147
            Top             =   840
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "最高诊断依据"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   13
            Left            =   4785
            TabIndex        =   149
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "抢救病因"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   22
            Left            =   5205
            TabIndex        =   171
            Top             =   2740
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "临床与尸检"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   21
            Left            =   3345
            TabIndex        =   184
            Top             =   3640
            Width           =   1050
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "临床与病理"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   19
            Left            =   4995
            TabIndex        =   163
            Top             =   2240
            Width           =   1050
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "放射与病理"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   18
            Left            =   8715
            TabIndex        =   165
            Top             =   2240
            Width           =   1050
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与入院"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   16
            Left            =   1275
            TabIndex        =   151
            Top             =   1240
            Width           =   1050
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "死亡期间"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   20
            Left            =   3555
            TabIndex        =   177
            Top             =   3240
            Width           =   840
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "死亡原因"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   20
            Left            =   6600
            TabIndex        =   179
            Top             =   3240
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院情况"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   11
            Left            =   1485
            TabIndex        =   139
            Top             =   340
            Width           =   840
         End
         Begin VB.Label lblDateInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "主要诊断确认日期"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   7275
            TabIndex        =   143
            Top             =   345
            Width           =   1680
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "医院感染病原学诊断"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   21
            Left            =   4155
            TabIndex        =   158
            Top             =   1740
            Width           =   1890
         End
         Begin VB.Line lineH 
            Index           =   6
            X1              =   0
            X2              =   14400
            Y1              =   3090
            Y2              =   3090
         End
         Begin VB.Line lineH 
            Index           =   5
            X1              =   0
            X2              =   14400
            Y1              =   2590
            Y2              =   2590
         End
         Begin VB.Line lineH 
            Index           =   4
            X1              =   0
            X2              =   14400
            Y1              =   2090
            Y2              =   2090
         End
         Begin VB.Line lineH 
            Index           =   3
            X1              =   0
            X2              =   14400
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Line lineH 
            Index           =   2
            X1              =   0
            X2              =   14200
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "西医诊断符合情况"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   0
            TabIndex        =   138
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2500
         Index           =   8
         Left            =   738
         ScaleHeight     =   2505
         ScaleWidth      =   11505
         TabIndex        =   437
         Tag             =   "true"
         Top             =   26590
         Width           =   11500
         Begin VB.CommandButton cmdDateInfo 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   8
            Left            =   11145
            Picture         =   "frmPageMedRecEdit_YN.frx":0D7E
            Style           =   1  'Graphical
            TabIndex        =   280
            TabStop         =   0   'False
            Top             =   1680
            Width           =   270
         End
         Begin VB.CommandButton cmdDateInfo 
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
            Index           =   9
            Left            =   5325
            Picture         =   "frmPageMedRecEdit_YN.frx":0E74
            Style           =   1  'Graphical
            TabIndex        =   286
            TabStop         =   0   'False
            Top             =   2085
            Width           =   270
         End
         Begin VB.CommandButton cmdDateInfo 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   10
            Left            =   8205
            Picture         =   "frmPageMedRecEdit_YN.frx":0F6A
            Style           =   1  'Graphical
            TabIndex        =   290
            TabStop         =   0   'False
            Top             =   2080
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   255
            Index           =   9
            Left            =   3990
            TabIndex        =   284
            Tag             =   "####-##-##"
            Top             =   2085
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   255
            Index           =   10
            Left            =   6870
            TabIndex        =   288
            Tag             =   "####-##-##"
            Top             =   2085
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   255
            Index           =   8
            Left            =   9810
            TabIndex        =   278
            Tag             =   "####-##-##"
            Top             =   1680
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   3
            Left            =   3990
            Sorted          =   -1  'True
            TabIndex        =   264
            Top             =   740
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   10
            Left            =   1020
            Sorted          =   -1  'True
            TabIndex        =   262
            Top             =   740
            Width           =   1600
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   33
            ItemData        =   "frmPageMedRecEdit_YN.frx":1060
            Left            =   1020
            List            =   "frmPageMedRecEdit_YN.frx":106D
            Style           =   2  'Dropdown List
            TabIndex        =   272
            Top             =   1640
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   1
            Left            =   1020
            TabIndex        =   254
            Top             =   340
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   8
            Left            =   3990
            Sorted          =   -1  'True
            TabIndex        =   274
            Top             =   1640
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   9
            Left            =   6870
            Sorted          =   -1  'True
            TabIndex        =   276
            Top             =   1640
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   2
            Left            =   3990
            Sorted          =   -1  'True
            TabIndex        =   256
            Top             =   340
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   7
            Left            =   6870
            Sorted          =   -1  'True
            TabIndex        =   266
            Top             =   740
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   6
            Left            =   9810
            Sorted          =   -1  'True
            TabIndex        =   268
            Top             =   735
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   0
            ItemData        =   "frmPageMedRecEdit_YN.frx":1083
            Left            =   1020
            List            =   "frmPageMedRecEdit_YN.frx":1085
            TabIndex        =   270
            Top             =   1140
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   4
            Left            =   6870
            TabIndex        =   258
            Top             =   340
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   5
            Left            =   9810
            TabIndex        =   260
            Top             =   340
            Width           =   1600
         End
         Begin VB.ComboBox cboManInfo 
            Height          =   330
            Index           =   11
            Left            =   1020
            Sorted          =   -1  'True
            TabIndex        =   282
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1600
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   8
            Left            =   9810
            MaxLength       =   30
            TabIndex        =   279
            Top             =   1680
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   10
            Left            =   6870
            MaxLength       =   30
            TabIndex        =   289
            Top             =   2080
            Visible         =   0   'False
            Width           =   1600
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   9
            Left            =   3990
            MaxLength       =   30
            TabIndex        =   285
            Top             =   2080
            Visible         =   0   'False
            Width           =   1600
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "责任护士"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   261
            Top             =   795
            Width           =   840
         End
         Begin VB.Label lblDateInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "质控日期"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   8
            Left            =   8910
            TabIndex        =   277
            Top             =   1695
            Width           =   840
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "质控医师"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   8
            Left            =   3090
            TabIndex        =   273
            Top             =   1695
            Width           =   840
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "质控护士"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   9
            Left            =   5970
            TabIndex        =   275
            Top             =   1695
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病案质量"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   33
            Left            =   120
            TabIndex        =   271
            Top             =   1695
            Width           =   840
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "主任(副主任)医师"
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   2
            Left            =   3000
            TabIndex        =   255
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "实习医师"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   7
            Left            =   5970
            TabIndex        =   265
            Top             =   795
            Width           =   840
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "研究生医师"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   8700
            TabIndex        =   267
            Top             =   795
            Width           =   1050
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "门诊医师"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   269
            Top             =   1200
            Width           =   840
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "进修医师"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   3090
            TabIndex        =   263
            Top             =   795
            Width           =   840
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "科主任"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   330
            TabIndex        =   253
            Top             =   405
            Width           =   630
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院医师"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   8910
            TabIndex        =   259
            Top             =   405
            Width           =   840
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "主治医师"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   5970
            TabIndex        =   257
            Top             =   405
            Width           =   840
         End
         Begin VB.Line lineH 
            Index           =   10
            X1              =   0
            X2              =   14400
            Y1              =   1550
            Y2              =   1550
         End
         Begin VB.Label lblManInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "编目员"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   11
            Left            =   330
            TabIndex        =   281
            Top             =   2100
            Width           =   630
         End
         Begin VB.Label lblDateInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "编目日期"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   9
            Left            =   3090
            TabIndex        =   283
            Top             =   2100
            Width           =   840
         End
         Begin VB.Label lblDateInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "收回日期"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   10
            Left            =   5970
            TabIndex        =   287
            Top             =   2115
            Width           =   840
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "签名信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   8
            Left            =   0
            TabIndex        =   252
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3660
         Index           =   2
         Left            =   738
         ScaleHeight     =   3660
         ScaleWidth      =   11505
         TabIndex        =   431
         Tag             =   "true"
         Top             =   9390
         Width           =   11500
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   11120
            Picture         =   "frmPageMedRecEdit_YN.frx":1087
            Style           =   1  'Graphical
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   1680
            Width           =   375
         End
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   11120
            Picture         =   "frmPageMedRecEdit_YN.frx":37B2
            Style           =   1  'Graphical
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   1200
            Width           =   375
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
            Height          =   3100
            Left            =   0
            TabIndex        =   130
            Top             =   360
            Width           =   11055
            _cx             =   19500
            _cy             =   5468
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   9
            Cols            =   26
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPageMedRecEdit_YN.frx":5D5A
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "西医诊断"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   0
            TabIndex        =   129
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2000
         Index           =   6
         Left            =   738
         ScaleHeight     =   1995
         ScaleWidth      =   11505
         TabIndex        =   435
         Tag             =   "true"
         Top             =   22345
         Width           =   11500
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "无过敏记录"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   30
            Left            =   1560
            TabIndex        =   221
            Top             =   120
            Width           =   1455
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAller 
            Height          =   1400
            Left            =   0
            TabIndex        =   222
            Top             =   400
            Width           =   11490
            _cx             =   20267
            _cy             =   2469
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmPageMedRecEdit_YN.frx":6104
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "药物过敏"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   0
            TabIndex        =   220
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Index           =   5
         Left            =   738
         ScaleHeight     =   2655
         ScaleWidth      =   11505
         TabIndex        =   434
         Tag             =   "true"
         Top             =   19690
         Width           =   11500
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   23
            ItemData        =   "frmPageMedRecEdit_YN.frx":61BD
            Left            =   7020
            List            =   "frmPageMedRecEdit_YN.frx":61BF
            Style           =   2  'Dropdown List
            TabIndex        =   195
            Top             =   285
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   22
            ItemData        =   "frmPageMedRecEdit_YN.frx":61C1
            Left            =   3270
            List            =   "frmPageMedRecEdit_YN.frx":61C3
            Style           =   2  'Dropdown List
            TabIndex        =   193
            Top             =   285
            Width           =   1560
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "疑难"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   8
            Left            =   6600
            TabIndex        =   199
            Top             =   810
            Width           =   810
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "急症"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   7
            Left            =   4560
            TabIndex        =   198
            Top             =   810
            Width           =   810
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "危重"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   6
            Left            =   2400
            TabIndex        =   197
            Top             =   810
            Width           =   810
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   26
            ItemData        =   "frmPageMedRecEdit_YN.frx":61C5
            Left            =   9795
            List            =   "frmPageMedRecEdit_YN.frx":61C7
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   1285
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   25
            ItemData        =   "frmPageMedRecEdit_YN.frx":61C9
            Left            =   6420
            List            =   "frmPageMedRecEdit_YN.frx":61CB
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   1285
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   24
            ItemData        =   "frmPageMedRecEdit_YN.frx":61CD
            Left            =   2790
            List            =   "frmPageMedRecEdit_YN.frx":61CF
            Style           =   2  'Dropdown List
            TabIndex        =   202
            Top             =   1285
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   27
            ItemData        =   "frmPageMedRecEdit_YN.frx":61D1
            Left            =   2790
            List            =   "frmPageMedRecEdit_YN.frx":61D3
            Style           =   2  'Dropdown List
            TabIndex        =   209
            Top             =   1785
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   29
            ItemData        =   "frmPageMedRecEdit_YN.frx":61D5
            Left            =   6420
            List            =   "frmPageMedRecEdit_YN.frx":61D7
            Style           =   2  'Dropdown List
            TabIndex        =   211
            Top             =   1785
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   31
            ItemData        =   "frmPageMedRecEdit_YN.frx":61D9
            Left            =   9795
            List            =   "frmPageMedRecEdit_YN.frx":61DB
            Style           =   2  'Dropdown List
            TabIndex        =   213
            Top             =   1785
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   28
            ItemData        =   "frmPageMedRecEdit_YN.frx":61DD
            Left            =   2790
            List            =   "frmPageMedRecEdit_YN.frx":61DF
            Style           =   2  'Dropdown List
            TabIndex        =   215
            Top             =   2185
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   30
            ItemData        =   "frmPageMedRecEdit_YN.frx":61E1
            Left            =   6420
            List            =   "frmPageMedRecEdit_YN.frx":61E3
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   2185
            Width           =   1560
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   32
            ItemData        =   "frmPageMedRecEdit_YN.frx":61E5
            Left            =   9795
            List            =   "frmPageMedRecEdit_YN.frx":61E7
            Style           =   2  'Dropdown List
            TabIndex        =   219
            Top             =   2185
            Width           =   1560
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与出院"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   22
            Left            =   2160
            TabIndex        =   192
            Top             =   345
            Width           =   1050
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院与出院"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   23
            Left            =   5910
            TabIndex        =   194
            Top             =   345
            Width           =   1050
         End
         Begin VB.Line lineH 
            Index           =   9
            X1              =   0
            X2              =   14760
            Y1              =   1695
            Y2              =   1695
         End
         Begin VB.Line lineH 
            Index           =   7
            X1              =   0
            X2              =   14400
            Y1              =   695
            Y2              =   695
         End
         Begin VB.Line lineH 
            Index           =   8
            X1              =   0
            X2              =   14400
            Y1              =   1125
            Y2              =   1125
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "方药"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   26
            Left            =   9315
            TabIndex        =   205
            Top             =   1350
            Width           =   420
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "治法"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   25
            Left            =   5940
            TabIndex        =   203
            Top             =   1350
            Width           =   420
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "辨证"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   24
            Left            =   2310
            TabIndex        =   201
            Top             =   1350
            Width           =   420
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "治疗类别"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   27
            Left            =   1890
            TabIndex        =   208
            Top             =   1845
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "抢救方法"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   29
            Left            =   5520
            TabIndex        =   210
            Top             =   1845
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "自制中药制剂"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   31
            Left            =   8475
            TabIndex        =   212
            Top             =   1845
            Width           =   1260
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "使用中医诊疗设备"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   28
            Left            =   1050
            TabIndex        =   214
            Top             =   2250
            Width           =   1680
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "使用中医诊疗技术"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   30
            Left            =   4680
            TabIndex        =   216
            Top             =   2250
            Width           =   1680
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "辨证施护"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   32
            Left            =   8895
            TabIndex        =   218
            Top             =   2250
            Width           =   840
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "中医诊断符合情况"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   0
            TabIndex        =   191
            Top             =   0
            Width           =   1800
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院期间病情:"
            Height          =   210
            Index           =   600
            Left            =   75
            TabIndex        =   196
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "准确度:"
            Height          =   210
            Index           =   601
            Left            =   705
            TabIndex        =   200
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "治疗方法:"
            Height          =   210
            Index           =   602
            Left            =   495
            TabIndex        =   207
            Top             =   1845
            Width           =   945
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2600
         Index           =   4
         Left            =   738
         ScaleHeight     =   2595
         ScaleWidth      =   11505
         TabIndex        =   433
         Tag             =   "true"
         Top             =   17090
         Width           =   11500
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   2
            Left            =   11120
            Picture         =   "frmPageMedRecEdit_YN.frx":61E9
            Style           =   1  'Graphical
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   3
            Left            =   11120
            Picture         =   "frmPageMedRecEdit_YN.frx":8791
            Style           =   1  'Graphical
            TabIndex        =   190
            TabStop         =   0   'False
            Top             =   1320
            Width           =   375
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
            Height          =   1995
            Left            =   0
            TabIndex        =   188
            Top             =   405
            Width           =   11055
            _cx             =   19500
            _cy             =   3519
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   26
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPageMedRecEdit_YN.frx":AEBC
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "中医诊断"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   0
            TabIndex        =   187
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2830
         Index           =   9
         Left            =   738
         ScaleHeight     =   2835
         ScaleWidth      =   11505
         TabIndex        =   438
         Tag             =   "true"
         Top             =   29090
         Width           =   11500
         Begin VB.CommandButton cmdOPSMove 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   11120
            Picture         =   "frmPageMedRecEdit_YN.frx":B20E
            Style           =   1  'Graphical
            TabIndex        =   293
            TabStop         =   0   'False
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdOPSMove 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   11120
            Picture         =   "frmPageMedRecEdit_YN.frx":D7B6
            Style           =   1  'Graphical
            TabIndex        =   294
            TabStop         =   0   'False
            Top             =   1440
            Width           =   375
         End
         Begin VB.PictureBox PicOPS 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   810
            Left            =   0
            ScaleHeight     =   810
            ScaleWidth      =   11505
            TabIndex        =   295
            Top             =   1920
            Width           =   11500
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "发生术后猝死"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   5880
               TabIndex        =   297
               Top             =   0
               Width           =   1935
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "发生围术期死亡"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   3000
               TabIndex        =   296
               Top             =   0
               Width           =   2055
            End
            Begin VB.ComboBox cboBaseInfo 
               Height          =   330
               Index           =   17
               ItemData        =   "frmPageMedRecEdit_YN.frx":FEE1
               Left            =   1065
               List            =   "frmPageMedRecEdit_YN.frx":FEE3
               Style           =   2  'Dropdown List
               TabIndex        =   299
               Top             =   0
               Width           =   1515
            End
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   23
               Left            =   555
               MaxLength       =   3
               TabIndex        =   301
               Top             =   405
               Width           =   405
            End
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   25
               Left            =   4350
               MaxLength       =   3
               TabIndex        =   305
               Top             =   405
               Width           =   405
            End
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   26
               Left            =   6285
               MaxLength       =   3
               TabIndex        =   307
               Top             =   405
               Width           =   405
            End
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   27
               Left            =   7605
               MaxLength       =   3
               TabIndex        =   309
               Top             =   405
               Width           =   405
            End
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   28
               Left            =   9000
               MaxLength       =   3
               TabIndex        =   311
               Top             =   405
               Width           =   405
            End
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   250
               Index           =   24
               Left            =   2460
               MaxLength       =   3
               TabIndex        =   303
               Top             =   405
               Width           =   405
            End
            Begin VB.Label lblBaseInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "术前与术后"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   17
               Left            =   0
               TabIndex        =   298
               Top             =   60
               Width           =   1050
            End
            Begin VB.Label lblSpecificInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "特护    天"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   23
               Left            =   120
               TabIndex        =   300
               Top             =   420
               Width           =   1050
            End
            Begin VB.Label lblSpecificInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "一级护理    天"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   24
               Left            =   1605
               TabIndex        =   302
               Top             =   420
               Width           =   1470
            End
            Begin VB.Label lblSpecificInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "二级护理    天"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   25
               Left            =   3495
               TabIndex        =   304
               Top             =   420
               Width           =   1470
            End
            Begin VB.Label lblSpecificInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "三级护理    天"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   26
               Left            =   5445
               TabIndex        =   306
               Top             =   420
               Width           =   1470
            End
            Begin VB.Label lblSpecificInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ICU    天"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   27
               Left            =   7260
               TabIndex        =   308
               Top             =   420
               Width           =   945
            End
            Begin VB.Label lblSpecificInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "CCU     天"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   28
               Left            =   8625
               TabIndex        =   310
               Top             =   420
               Width           =   1050
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsOPS 
            Height          =   1515
            Left            =   -60
            TabIndex        =   292
            Top             =   360
            Width           =   11055
            _cx             =   19500
            _cy             =   2672
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   43
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPageMedRecEdit_YN.frx":FEE5
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            Editable        =   2
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
            Begin VB.PictureBox picCopy 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               Picture         =   "frmPageMedRecEdit_YN.frx":10599
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   448
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "手术记录"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   291
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2245
         Index           =   7
         Left            =   738
         ScaleHeight     =   2250
         ScaleWidth      =   11505
         TabIndex        =   436
         Tag             =   "true"
         Top             =   24345
         Width           =   11500
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   36
            ItemData        =   "frmPageMedRecEdit_YN.frx":1128B
            Left            =   1965
            List            =   "frmPageMedRecEdit_YN.frx":1128D
            Style           =   2  'Dropdown List
            TabIndex        =   225
            Top             =   225
            Width           =   1605
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   38
            ItemData        =   "frmPageMedRecEdit_YN.frx":1128F
            Left            =   5655
            List            =   "frmPageMedRecEdit_YN.frx":11291
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   225
            Width           =   1605
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   40
            ItemData        =   "frmPageMedRecEdit_YN.frx":11293
            Left            =   1965
            List            =   "frmPageMedRecEdit_YN.frx":112A0
            Style           =   2  'Dropdown List
            TabIndex        =   229
            Top             =   1025
            Width           =   1605
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   33
            Left            =   5655
            MaxLength       =   5
            TabIndex        =   243
            Tag             =   "11"
            Top             =   1825
            Width           =   1605
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   29
            Left            =   1965
            MaxLength       =   5
            TabIndex        =   227
            Tag             =   "11"
            Top             =   670
            Width           =   1365
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   30
            Left            =   5655
            MaxLength       =   5
            TabIndex        =   237
            Tag             =   "11"
            Top             =   670
            Width           =   1605
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   31
            Left            =   5655
            MaxLength       =   5
            TabIndex        =   239
            Tag             =   "11"
            Top             =   1065
            Width           =   1605
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   32
            Left            =   5655
            MaxLength       =   5
            TabIndex        =   241
            Tag             =   "11"
            Top             =   1465
            Width           =   1605
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   23
            Left            =   9300
            MaxLength       =   30
            TabIndex        =   251
            Top             =   1825
            Width           =   1980
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   41
            ItemData        =   "frmPageMedRecEdit_YN.frx":112B2
            Left            =   1965
            List            =   "frmPageMedRecEdit_YN.frx":112BF
            Style           =   2  'Dropdown List
            TabIndex        =   231
            Top             =   1425
            Width           =   1605
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   37
            ItemData        =   "frmPageMedRecEdit_YN.frx":112D1
            Left            =   9300
            List            =   "frmPageMedRecEdit_YN.frx":112D3
            Style           =   2  'Dropdown List
            TabIndex        =   247
            ToolTipText     =   "丙型肝炎病毒抗体"
            Top             =   1025
            Width           =   1980
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   42
            ItemData        =   "frmPageMedRecEdit_YN.frx":112D5
            Left            =   1965
            List            =   "frmPageMedRecEdit_YN.frx":112D7
            Style           =   2  'Dropdown List
            TabIndex        =   233
            Top             =   1785
            Width           =   1605
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   35
            ItemData        =   "frmPageMedRecEdit_YN.frx":112D9
            Left            =   9300
            List            =   "frmPageMedRecEdit_YN.frx":112DB
            Style           =   2  'Dropdown List
            TabIndex        =   245
            ToolTipText     =   "乙型肝炎表面抗原"
            Top             =   625
            Width           =   1980
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   39
            ItemData        =   "frmPageMedRecEdit_YN.frx":112DD
            Left            =   9300
            List            =   "frmPageMedRecEdit_YN.frx":112DF
            Style           =   2  'Dropdown List
            TabIndex        =   249
            ToolTipText     =   "获得性人类免疫缺陷病毒抗体"
            Top             =   1425
            Width           =   1980
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "血型"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   36
            Left            =   1485
            TabIndex        =   224
            Top             =   285
            Width           =   420
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rh"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   38
            Left            =   5385
            TabIndex        =   234
            Top             =   285
            Width           =   210
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输液反应"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   40
            Left            =   1065
            TabIndex        =   228
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "自体回收                ml"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   33
            Left            =   4770
            TabIndex        =   242
            Top             =   1845
            Width           =   2730
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输血反应"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   41
            Left            =   1065
            TabIndex        =   230
            Top             =   1485
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输红细胞               单位"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   29
            Left            =   1065
            TabIndex        =   226
            Top             =   690
            Width           =   2835
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输血小板                单位"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   30
            Left            =   4770
            TabIndex        =   236
            Top             =   690
            Width           =   2940
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输血浆                ml"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   31
            Left            =   4965
            TabIndex        =   238
            Top             =   1080
            Width           =   2520
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输全血                ml"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   32
            Left            =   4965
            TabIndex        =   240
            Top             =   1485
            Width           =   2520
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输其他"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   23
            Left            =   8610
            TabIndex        =   250
            Top             =   1845
            Width           =   630
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输血前的9项检查"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   42
            Left            =   330
            TabIndex        =   232
            Top             =   1845
            Width           =   1575
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "HIV-Ab"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   39
            Left            =   8610
            TabIndex        =   248
            ToolTipText     =   "获得性人类免疫缺陷病毒抗体"
            Top             =   1485
            Width           =   630
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "HCV-Ab"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   37
            Left            =   8610
            TabIndex        =   246
            ToolTipText     =   "丙型肝炎病毒抗体"
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "HBsAg"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   35
            Left            =   8715
            TabIndex        =   244
            ToolTipText     =   "乙型肝炎表面抗原"
            Top             =   690
            Width           =   525
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输血信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   7
            Left            =   0
            TabIndex        =   223
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3505
         Index           =   11
         Left            =   738
         ScaleHeight     =   3510
         ScaleWidth      =   11505
         TabIndex        =   440
         Tag             =   "true"
         Top             =   34520
         Width           =   11500
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   255
            Index           =   6
            Left            =   9450
            TabIndex        =   325
            Tag             =   "####-##-##"
            Top             =   390
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   255
            Index           =   7
            Left            =   10725
            TabIndex        =   326
            Tag             =   "##:##"
            Top             =   390
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   250
            Index           =   24
            Left            =   11115
            TabIndex        =   335
            TabStop         =   0   'False
            Top             =   1320
            Width           =   270
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   55
            ItemData        =   "frmPageMedRecEdit_YN.frx":112E1
            Left            =   5940
            List            =   "frmPageMedRecEdit_YN.frx":112E3
            Style           =   2  'Dropdown List
            TabIndex        =   357
            Top             =   3040
            Width           =   1605
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "住院期间出现危重"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   22
            Left            =   240
            TabIndex        =   355
            Top             =   3070
            Width           =   2265
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   40
            Left            =   9765
            MaxLength       =   6
            TabIndex        =   354
            Tag             =   "入院后(分钟)"
            Top             =   2680
            Width           =   510
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   43
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   330
            Top             =   735
            Width           =   1980
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   24
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   334
            TabStop         =   0   'False
            Top             =   1320
            Width           =   5625
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   39
            Left            =   8775
            MaxLength       =   6
            TabIndex        =   353
            Tag             =   "入院后(小时)"
            Top             =   2680
            Width           =   510
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   37
            Left            =   5565
            MaxLength       =   6
            TabIndex        =   350
            Tag             =   "入院前(分钟)"
            Top             =   2680
            Width           =   510
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   36
            Left            =   4455
            MaxLength       =   6
            TabIndex        =   349
            Tag             =   "入院前(小时)"
            Top             =   2680
            Width           =   510
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   35
            Left            =   3525
            MaxLength       =   6
            TabIndex        =   348
            Top             =   2680
            Width           =   510
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   38
            Left            =   7845
            MaxLength       =   6
            TabIndex        =   352
            Top             =   2680
            Width           =   510
         End
         Begin VB.OptionButton optInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "无"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   0
            Left            =   3780
            TabIndex        =   338
            Tag             =   "1"
            Top             =   1765
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "有，目的"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   1
            Left            =   4440
            TabIndex        =   339
            Top             =   1765
            Width           =   1215
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   25
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   340
            TabStop         =   0   'False
            Top             =   1780
            Width           =   5625
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   49
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   337
            Top             =   1740
            Width           =   2295
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   41
            Left            =   6810
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   345
            TabStop         =   0   'False
            Tag             =   "5"
            Top             =   2280
            Width           =   800
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "随诊"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   12
            Left            =   4365
            TabIndex        =   343
            Top             =   2285
            Width           =   795
         End
         Begin VB.ComboBox cboSpecificInfo 
            BackColor       =   &H80000004&
            Height          =   330
            Index           =   41
            ItemData        =   "frmPageMedRecEdit_YN.frx":112E5
            Left            =   7660
            List            =   "frmPageMedRecEdit_YN.frx":112E7
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   346
            TabStop         =   0   'False
            Top             =   2260
            Width           =   735
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "科研病案"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   10
            Left            =   5040
            TabIndex        =   322
            Top             =   370
            Width           =   1215
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "疑难病例"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   11
            Left            =   6480
            TabIndex        =   323
            Top             =   370
            Width           =   1275
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   34
            ItemData        =   "frmPageMedRecEdit_YN.frx":112E9
            Left            =   1200
            List            =   "frmPageMedRecEdit_YN.frx":112EB
            Style           =   2  'Dropdown List
            TabIndex        =   320
            Top             =   360
            Width           =   1980
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   34
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   342
            Tag             =   "呼吸机使用时间"
            Top             =   2280
            Width           =   735
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "示教病案"
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   9
            Left            =   3600
            TabIndex        =   321
            Top             =   370
            WhatsThisHelpID =   9
            Width           =   1215
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   44
            ItemData        =   "frmPageMedRecEdit_YN.frx":112ED
            Left            =   1200
            List            =   "frmPageMedRecEdit_YN.frx":112FA
            Style           =   2  'Dropdown List
            TabIndex        =   332
            Top             =   1245
            Width           =   1980
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   7
            Left            =   10725
            MaxLength       =   30
            TabIndex        =   328
            Top             =   390
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   6
            Left            =   9450
            MaxLength       =   30
            TabIndex        =   327
            Top             =   385
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "产科新生儿情况(离院方式)"
            Height          =   210
            Index           =   55
            Left            =   3360
            TabIndex        =   356
            Top             =   3105
            Width           =   2520
         End
         Begin VB.Label lblDateInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "发病时间"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   8505
            TabIndex        =   324
            Top             =   405
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "生育状况"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   43
            Left            =   300
            TabIndex        =   329
            Top             =   795
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "昏迷时间(颅脑损伤患者)   入院前       天      小时      分钟"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   35
            Left            =   240
            TabIndex        =   347
            Top             =   2700
            Width           =   6300
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "离院方式"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   44
            Left            =   300
            TabIndex        =   331
            Top             =   1305
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "随诊期限"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   41
            Left            =   5925
            TabIndex        =   344
            Top             =   2300
            Width           =   840
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病例分型"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   34
            Left            =   300
            TabIndex        =   319
            Top             =   405
            Width           =   840
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "呼吸机使用时间        小时"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   34
            Left            =   285
            TabIndex        =   341
            Top             =   2300
            Width           =   2730
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院后       天      小时      分钟"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   38
            Left            =   7125
            TabIndex        =   351
            Top             =   2700
            Width           =   3675
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院情况"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   11
            Left            =   0
            TabIndex        =   318
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "转入"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   24
            Left            =   5235
            TabIndex        =   333
            Top             =   1305
            Width           =   420
         End
         Begin VB.Line lineH 
            Index           =   11
            X1              =   0
            X2              =   14000
            Y1              =   1150
            Y2              =   1150
         End
         Begin VB.Line lineH 
            Index           =   12
            X1              =   0
            X2              =   14120
            Y1              =   1650
            Y2              =   1650
         End
         Begin VB.Line lineH 
            BorderStyle     =   2  'Dash
            DrawMode        =   1  'Blackness
            Index           =   13
            X1              =   0
            X2              =   14120
            Y1              =   2150
            Y2              =   2150
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "是否有"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   10
            Left            =   510
            TabIndex        =   336
            Top             =   1800
            Width           =   630
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3375
         Index           =   12
         Left            =   738
         ScaleHeight     =   3375
         ScaleWidth      =   11505
         TabIndex        =   441
         Tag             =   "true"
         Top             =   38025
         Width           =   11500
         Begin VSFlex8Ctl.VSFlexGrid vsChemoth 
            Height          =   2835
            Left            =   0
            TabIndex        =   360
            Top             =   330
            Width           =   11490
            _cx             =   20267
            _cy             =   5001
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmPageMedRecEdit_YN.frx":11310
            ScrollTrack     =   0   'False
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
         Begin VB.Label lblEdit 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "提示信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   0
            Left            =   1995
            TabIndex        =   359
            Top             =   0
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "化疗记录信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   12
            Left            =   0
            TabIndex        =   358
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2175
         Index           =   17
         Left            =   738
         ScaleHeight     =   2175
         ScaleWidth      =   11505
         TabIndex        =   446
         Tag             =   "true"
         Top             =   46980
         Width           =   11500
         Begin VSFlex8Ctl.VSFlexGrid vsfMain 
            Height          =   1605
            Left            =   0
            TabIndex        =   370
            Top             =   360
            Width           =   11490
            _cx             =   20267
            _cy             =   2831
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483642
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   100
            ColWidthMax     =   2400
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmPageMedRecEdit_YN.frx":1143D
            ScrollTrack     =   0   'False
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病案附加项目"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   17
            Left            =   0
            TabIndex        =   369
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Index           =   14
         Left            =   738
         ScaleHeight     =   2535
         ScaleWidth      =   11505
         TabIndex        =   443
         Tag             =   "true"
         Top             =   43935
         Width           =   11500
         Begin VB.CommandButton cmdAutoLoad 
            Caption         =   "自动提取"
            Height          =   330
            Index           =   0
            Left            =   9720
            TabIndex        =   365
            TabStop         =   0   'False
            Top             =   120
            Width           =   1200
         End
         Begin VSFlex8Ctl.VSFlexGrid vsKSS 
            Height          =   1845
            Left            =   0
            TabIndex        =   366
            Top             =   480
            Width           =   11490
            _cx             =   20267
            _cy             =   3254
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmPageMedRecEdit_YN.frx":1154D
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "抗菌药物使用情况（按DDD数降序排列）"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   14
            Left            =   0
            TabIndex        =   364
            Top             =   0
            Width           =   3960
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Index           =   13
         Left            =   738
         ScaleHeight     =   2535
         ScaleWidth      =   11505
         TabIndex        =   442
         Tag             =   "true"
         Top             =   41400
         Width           =   11500
         Begin VSFlex8Ctl.VSFlexGrid vsRadioth 
            Height          =   1965
            Left            =   0
            TabIndex        =   363
            Top             =   375
            Width           =   11490
            _cx             =   20267
            _cy             =   3466
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmPageMedRecEdit_YN.frx":1166C
            ScrollTrack     =   0   'False
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
         Begin VB.Label lblEdit 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "提示信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   1
            Left            =   2115
            TabIndex        =   362
            Top             =   0
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "放疗记录信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   13
            Left            =   0
            TabIndex        =   361
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   738
         ScaleHeight     =   255
         ScaleWidth      =   11505
         TabIndex        =   444
         Tag             =   "true"
         Top             =   46470
         Width           =   11500
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "抗精神病治疗情况"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   15
            Left            =   0
            TabIndex        =   367
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   738
         ScaleHeight     =   255
         ScaleWidth      =   11505
         TabIndex        =   445
         Tag             =   "true"
         Top             =   46725
         Width           =   11500
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重症监护情况"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   0
            TabIndex        =   368
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Index           =   0
         Left            =   720
         ScaleHeight     =   2535
         ScaleWidth      =   11505
         TabIndex        =   429
         Tag             =   "true"
         Top             =   0
         Width           =   11500
         Begin VB.CommandButton cmdSpecificInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   20
            Left            =   2070
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   2130
            Width           =   270
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   250
            Index           =   20
            Left            =   960
            MaxLength       =   18
            TabIndex        =   11
            Tag             =   "18"
            ToolTipText     =   "按F4弹出病人选择器"
            Top             =   2130
            Width           =   1380
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   1
            Left            =   5205
            MaxLength       =   20
            TabIndex        =   14
            Tag             =   "20"
            Top             =   2130
            Width           =   1380
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   2
            Left            =   10005
            MaxLength       =   20
            TabIndex        =   16
            Top             =   2130
            Width           =   1380
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H000000FF&
            Height          =   250
            Index           =   0
            Left            =   10005
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "18"
            Top             =   1730
            Width           =   1380
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H000000FF&
            Height          =   250
            Index           =   11
            Left            =   5205
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Tag             =   "5"
            Text            =   "1"
            Top             =   1730
            Width           =   465
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   330
            Index           =   0
            ItemData        =   "frmPageMedRecEdit_YN.frx":11793
            Left            =   960
            List            =   "frmPageMedRecEdit_YN.frx":11795
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1290
            Width           =   2600
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   63
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1730
            Width           =   1380
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院号"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   20
            Left            =   240
            TabIndex        =   10
            Top             =   2150
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "档案号"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   4560
            TabIndex        =   13
            Top             =   2145
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "X线号"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   9435
            TabIndex        =   15
            Top             =   2145
            Width           =   525
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病案号"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   9330
            TabIndex        =   8
            Top             =   1755
            Width           =   630
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "第       次住院"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   11
            Left            =   4800
            TabIndex        =   6
            Top             =   1755
            Width           =   1575
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "付费方式"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   1350
            Width           =   840
         End
         Begin VB.Label lblHead 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病 案 首 页"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   4500
            TabIndex        =   0
            Tag             =   "241,88"
            Top             =   360
            Width           =   2085
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "健康卡号"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   60
            TabIndex        =   4
            Top             =   1750
            Width           =   840
         End
         Begin VB.Label lblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "提示：只能对已经接收的病案进行编目"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   600
            TabIndex        =   1
            Top             =   720
            Width           =   3570
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2600
         Index           =   10
         Left            =   738
         ScaleHeight     =   2595
         ScaleWidth      =   11505
         TabIndex        =   439
         Tag             =   "true"
         Top             =   31920
         Width           =   11500
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECFFCC&
            BorderStyle     =   0  'None
            Height          =   250
            Index           =   42
            Left            =   5955
            Locked          =   -1  'True
            TabIndex        =   315
            TabStop         =   0   'False
            Tag             =   "11"
            Top             =   242
            Width           =   1635
         End
         Begin VB.CommandButton cmdFeeEdit 
            Caption         =   "住院费用"
            Height          =   330
            Left            =   8760
            TabIndex        =   316
            TabStop         =   0   'False
            Top             =   202
            Width           =   1200
         End
         Begin VB.CheckBox chkFeeEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "允许编辑下表中的费用信息"
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   1200
            TabIndex        =   313
            Top             =   227
            Width           =   3015
         End
         Begin VSFlex8Ctl.VSFlexGrid vsFees 
            Height          =   1800
            Left            =   0
            TabIndex        =   317
            Top             =   600
            Width           =   11490
            _cx             =   20267
            _cy             =   3175
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   325
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmPageMedRecEdit_YN.frx":11797
            ScrollTrack     =   0   'False
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
         Begin MSComctlLib.ListView lvwFee 
            Height          =   3015
            Left            =   1920
            TabIndex        =   450
            Top             =   2280
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   5318
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "费用名称"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "金额"
               Object.Width           =   1270
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "住院次数"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院总费用"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   42
            Left            =   4845
            TabIndex        =   314
            Top             =   262
            Width           =   1050
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院费用"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   10
            Left            =   0
            TabIndex        =   312
            Top             =   0
            Width           =   900
         End
      End
      Begin MSComCtl2.MonthView monInfo 
         Height          =   2460
         Left            =   840
         TabIndex        =   449
         Top             =   0
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   4339
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollRate      =   1
         StartOfWeek     =   59637761
         TitleBackColor  =   8421504
         TitleForeColor  =   16777215
         CurrentDate     =   38003
         MaxDate         =   73415
         MinDate         =   -18260
      End
   End
   Begin zlSubclass.Subclass subcMain 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Image imgButtonNew 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   500
      Picture         =   "frmPageMedRecEdit_YN.frx":11880
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonDel 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   0
      Picture         =   "frmPageMedRecEdit_YN.frx":11E0A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmPageMedRecEdit_YN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'gclsPros.MainInfoRec.RecordCount =132，gclsPros.SecdInfoRec.RecordCount =29，不包括加载扩展的控件的次级信息

Private Sub cboManInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call ManInfoKeyDown(Index, KeyCode)
End Sub

Private Sub cmdTop_Click()
    Call cmdTopClick
End Sub

Private Sub cmdTop_GotFocus()
    Call cmdTopGotFocus
End Sub

Private Sub Form_Load()
    Call FormLoad
End Sub

Private Sub Form_Resize()
    Call FormResize
End Sub

Private Sub hsbMain_Change()
    Call hsbMainChange
End Sub

Private Sub padrInfo_SetInput(Index As Integer, ByVal intLevel As Integer, rsReturn As ADODB.Recordset)
    Call SetYoubian(Index, intLevel, rsReturn)
End Sub


Private Sub PicPage_Resize(Index As Integer)
    Call PicPageResize(Index)
End Sub

Private Sub padrInfo_Change(Index As Integer)
    Call CheckValueChange
End Sub

Private Sub txtAdressInfo_Change(Index As Integer)
    Call txtAdressInfoChange(Index)
End Sub

Private Sub txtAdressInfo_GotFocus(Index As Integer)
    Call txtAdressInfoGotFocus(Index)
End Sub

Private Sub txtDateInfo_Change(Index As Integer)
    Call CheckValueChange(txtDateInfo(Index))
End Sub

Private Sub txtDateInfo_GotFocus(Index As Integer)
    Call txtDateInfoGotFocus(Index)
End Sub

Private Sub vsAller_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsAller)
End Sub

Private Sub vsbMain_Change()
   Call vsbMainChange
End Sub

Private Sub subcMain_WndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
    Call SubCMainWndProc(msg, wParam, lParam, Result)
End Sub

Private Sub cboBaseInfo_Click(Index As Integer)
    Call CboBaseInfoClick(Index)
End Sub

Private Sub cboBaseInfo_GotFocus(Index As Integer)
    Call CboBaseInfoGotFocus(Index)
End Sub

Private Sub cboBaseInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboBaseInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub cboBaseInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CboBaseInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cboBaseInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboBaseInfoKeyUp(Index, KeyCode, Shift)
End Sub

Private Sub cboBaseInfo_Validate(Index As Integer, Cancel As Boolean)
    Call cboBaseInfoValidate(Index, Cancel)
End Sub

Private Sub cboManInfo_Click(Index As Integer)
    Call ManInfoClick(Index)
End Sub

Private Sub cboManInfo_DropDown(Index As Integer)
    Call ManInfoDropDown(Index)
End Sub

Private Sub cboManInfo_GotFocus(Index As Integer)
    Call ManInfoGotFocus(Index)
End Sub

Private Sub cboManInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call ManInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cboManInfo_LostFocus(Index As Integer)
    Call ManInfoLostFocus(Index)
End Sub

Private Sub cboManInfo_Validate(Index As Integer, Cancel As Boolean)
    Call ManInfoValidate(Index, Cancel)
End Sub

Private Sub cboBaseInfo_Change(Index As Integer)
    Call CboBaseInfoChange(Index)
End Sub

Private Sub CboSpecificInfo_Click(Index As Integer)
    Call CboSpecificInfoClick(Index)
End Sub

Private Sub cboSpecificInfo_GotFocus(Index As Integer)
    Call CboSpecificInfoGotFocus(Index)
End Sub

Private Sub cboSpecificInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboSpecificInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub chkFeeEdit_Click()
    Call ChkFeeEditClick
End Sub

Private Sub chkFeeEdit_KeyPress(KeyAscii As Integer)
    Call ChkFeeEditKeyPress(KeyAscii)
End Sub

Private Sub chkInfo_Click(Index As Integer)
    Call chkInfoClick(Index)
End Sub

Private Sub chkInfo_GotFocus(Index As Integer)
    Call ChkInfoGotFocus(Index)
End Sub

Private Sub chkInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call ChkInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cmdAdressInfo_Click(Index As Integer)
    Call CmdAdressInfoClick(Index)
End Sub

Private Sub cmdAutoLoad_Click(Index As Integer)
    Call CmdAutoLoadClick(Index)
End Sub

Private Sub cmdDateInfo_Click(Index As Integer)
    Call DateInfoClick(Index)
End Sub

Private Sub cmdDeliceryInfo_Click()
    Call CmdDeliceryInfoClick
End Sub

Private Sub cmdDeliceryInfo_GotFocus()
    Call CmdDeliceryInfoGotFocus
End Sub

Private Sub cmdDiagMove_Click(Index As Integer)
    Call CmdDiagMoveClick(Index)
End Sub

Private Sub cmdDiagMove_GotFocus(Index As Integer)
    Call CmdDiagMoveGotFocus(Index)
End Sub

Private Sub cmdDown_Click()
    Call CmdDownClick
End Sub

Private Sub cmdDown_GotFocus()
    Call CmdDownGotFocus
End Sub

Private Sub cmdFeeEdit_Click()
    Call cmdFeeEditClick
End Sub

Private Sub cmdHelp_Click()
    Call CmdHelpClick
End Sub

Private Sub cmdHelp_GotFocus()
    Call CmdHelpGotFocus
End Sub

Private Sub cmdInfo_Click(Index As Integer)
    Call CmdInfoClick(Index)
End Sub

Private Sub cmdLastDiag_Click()
    Call cmdLastDiagClick
End Sub

Private Sub cmdOPSMove_Click(Index As Integer)
    Call cmdOPSMoveClick(Index)
End Sub

Private Sub cmdSpecificInfo_Click(Index As Integer)
    Call SpecificInfoClick(Index, True)
End Sub

Private Sub cmdUp_Click()
    Call CmdUPClick
End Sub

Private Sub cmdUp_GotFocus()
    Call CmdUPGotFocus
End Sub

Private Sub Form_Activate()
    Call FormActivate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FormKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call FormKeyPress(KeyAscii)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FormUnLoad(Cancel)
End Sub

Private Sub lstAdvEvent_GotFocus()
    Call LstGotFocus(lstAdvEvent)
End Sub

Private Sub lstAdvEvent_ItemCheck(Item As Integer)
    Call LstItemCheck(lstAdvEvent, Item)
End Sub

Private Sub lstAdvEvent_KeyDown(KeyCode As Integer, Shift As Integer)
    Call LstKeyDown(lstAdvEvent, KeyCode, Shift)
End Sub

Private Sub lstAdvEvent_KeyPress(KeyAscii As Integer)
    Call LstKeyPress(lstAdvEvent, KeyAscii)
End Sub

Private Sub lstInfection_GotFocus()
    Call LstGotFocus(lstInfection)
End Sub

Private Sub lstInfection_ItemCheck(Item As Integer)
    Call LstItemCheck(lstInfection, Item)
End Sub

Private Sub lstInfection_KeyDown(KeyCode As Integer, Shift As Integer)
    Call LstKeyDown(lstInfection, KeyCode, Shift)
End Sub

Private Sub lstInfection_KeyPress(KeyAscii As Integer)
    Call LstKeyPress(lstInfection, KeyAscii)
End Sub

Private Sub lstInfectParts_GotFocus()
    Call LstGotFocus(lstInfectParts)
End Sub

Private Sub lstInfectParts_ItemCheck(Item As Integer)
    Call LstItemCheck(lstInfectParts, Item)
End Sub

Private Sub lstInfectParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Call LstKeyDown(lstInfectParts, KeyCode, Shift)
End Sub

Private Sub lstInfectParts_KeyPress(KeyAscii As Integer)
    Call LstKeyPress(lstInfectParts, KeyAscii)
End Sub

Private Sub monInfo_DateClick(ByVal DateClicked As Date)
    Call monInfoDateClick(DateClicked)
End Sub

Private Sub monInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call monInfoKeyDown(KeyCode, Shift)
End Sub

Private Sub monInfo_KeyPress(KeyAscii As Integer)
    Call monInfoKeyPress(KeyAscii)
End Sub

Private Sub monInfo_Validate(Cancel As Boolean)
    Call monInfoValidate(Cancel)
End Sub

Private Sub lvwFee_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call lvwFeeItemCheck(Item)
End Sub

Private Sub mskDateInfo_Change(Index As Integer)
    Call DateInfoChange(Index)
End Sub

Private Sub mskDateInfo_GotFocus(Index As Integer)
    Call DateInfoGotFocus(Index)
End Sub

Private Sub mskDateInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call DateInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub mskDateInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call DateInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub mskDateInfo_Validate(Index As Integer, Cancel As Boolean)
    Call DateInfoValidate(Index, Cancel)
End Sub

Private Sub optInput_Click(Index As Integer)
    Call OptInputClick(Index)
End Sub

Private Sub optInput_KeyPress(Index As Integer, KeyAscii As Integer)
    Call OptInputKeyPress(Index, KeyAscii)
End Sub

Private Sub picCopy_Click()
    Call picCopyClick
End Sub

Private Sub txtAdressInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call txtAdressInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtAdressInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call txtAdressInfoMouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub txtAdressInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call txtAdressInfoMouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub txtInfo_Change(Index As Integer)
    Call TxtInfoChange(Index)
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call TxtInfoGotFocus(Index)
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call TxtInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call TxtInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TxtInfoMouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub txtInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TxtInfoMouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Call TxtInfoValidate(Index, Cancel)
End Sub

Private Sub txtSpecificInfo_Change(Index As Integer)
    Call SpecificInfoChange(Index)
End Sub

Private Sub txtSpecificInfo_GotFocus(Index As Integer)
    Call SpecificInfoGotFocus(Index)
End Sub

Private Sub txtSpecificInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call SpecificInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub txtSpecificInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call SpecificInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtSpecificInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SpecificInfoMouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub txtSpecificInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SpecificInfoMouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub txtSpecificInfo_Validate(Index As Integer, Cancel As Boolean)
    Call SpecificInfoValidate(Index, Cancel)
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call AllerAfterEdit(vsAller, Row, Col)
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call AllerAfterRowColChange(vsAller, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call AllerCellButtonClick(vsAller, Row, Col)
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Call AllerKeyDown(vsAller, KeyCode, Shift)
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    Call AllerKeyPress(vsAller, KeyAscii)
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call AllerKeyPressEdit(vsAller, Row, Col, KeyAscii)
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call AllerSetupEditWindow(vsAller, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call AllerStartEdit(vsAller, Row, Col, Cancel)
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call AllerValidateEdit(vsAller, Row, Col, Cancel)
End Sub

Private Sub vsChemoth_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call ChemothAfterEdit(vsChemoth, Row, Col)
End Sub

Private Sub vsChemoth_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call ChemothAfterRowColChange(vsChemoth, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsChemoth_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsChemoth)
End Sub

Private Sub vsChemoth_GotFocus()
    Call VSFlxGotFocus(vsChemoth)
End Sub

Private Sub vsChemoth_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ChemothKeyDown(vsChemoth, KeyCode, Shift)
End Sub

Private Sub vsChemoth_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Call ChemothKeyDownEdit(vsChemoth, Row, Col, KeyCode, Shift)
End Sub

Private Sub vsChemoth_KeyPress(KeyAscii As Integer)
    Call ChemothKeyPress(vsChemoth, KeyAscii)
End Sub

Private Sub vsChemoth_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call ChemothKeyPressEdit(vsChemoth, Row, Col, KeyAscii)
End Sub

Private Sub vsChemoth_LostFocus()
    Call ChemothLostFocus(vsChemoth)
End Sub

Private Sub vsChemoth_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call ChemothValidateEdit(vsChemoth, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagXY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsDiagXY)
End Sub

Private Sub vsDiagXY_Click()
    Call DiagClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_DblClick()
    Call DiagDblClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_GotFocus()
    Call DiagGotFocus(vsDiagXY)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagXY, KeyCode, Shift)
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagXY, KeyAscii)
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagXY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagXY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagZY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsDiagZY)
End Sub

Private Sub vsDiagZY_Click()
    Call DiagClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_DblClick()
    Call DiagDblClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_GotFocus()
    Call DiagGotFocus(vsDiagZY)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagZY, KeyCode, Shift)
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagZY, KeyAscii)
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagZY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagZY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsFees_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange
End Sub

Private Sub vsFees_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call vsFeesComboDropDown(Row, Col)
End Sub

Private Sub vsFees_KeyDown(KeyCode As Integer, Shift As Integer)
    Call vsFeesKeyDown(KeyCode, Shift)
End Sub

Private Sub vsFees_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call vsFeesKeyPressEdit(Row, Col, KeyAscii)
End Sub

Private Sub vsFees_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsFeesStartEdit(Row, Col, Cancel)
End Sub

Private Sub vsFees_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsFeesValidateEdit(Row, Col, Cancel)
End Sub

Private Sub vsfMain_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsfMain)
End Sub

Private Sub vsfMain_EnterCell()
    Call vsfMainEnterCell(vsfMain)
End Sub

Private Sub vsfMain_GotFocus()
    Call VSFlxGotFocus(vsfMain)
End Sub

Private Sub vsfMain_KeyPress(KeyAscii As Integer)
    Call vsfMainKeyPress(vsfMain, KeyAscii)
End Sub

Private Sub vsfMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsfMainStartEdit(vsfMain, Row, Col, Cancel)
End Sub

Private Sub vsfMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsfMainValidateEdit(vsfMain, Row, Col, Cancel)
End Sub

Private Sub vsKSS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call KSSAfterEdit(vsKSS, Row, Col)
End Sub

Private Sub vsKSS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call KSSAfterRowColChange(vsKSS, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsKSS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call KSSCellButtonClick(vsKSS, Row, Col)
End Sub

Private Sub vsKSS_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsKSS)
End Sub

Private Sub vsKSS_GotFocus()
    Call VSFlxGotFocus(vsKSS)
End Sub

Private Sub vsKSS_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KSSKeyDown(vsKSS, KeyCode, Shift)
End Sub

Private Sub vsKSS_KeyPress(KeyAscii As Integer)
    Call KSSKeyPress(vsKSS, KeyAscii)
End Sub

Private Sub vsKSS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call KSSKeyPressEdit(vsKSS, Row, Col, KeyAscii)
End Sub

Private Sub vsKSS_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call KSSSetupEditWindow(vsKSS, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsKSS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call KSSValidateEdit(vsKSS, Row, Col, Cancel)
End Sub

Private Sub vsOPS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call OPSAfterEdit(vsOPS, Row, Col)
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call OPSAfterRowColChange(vsOPS, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsOPS_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call OPSBeforeUserResize(vsOPS, Row, Col, Cancel)
End Sub

Private Sub vsOPS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call OPSCellButtonClick(vsOPS, Row, Col)
End Sub

Private Sub vsOPS_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsOPS)
End Sub

Private Sub vsOPS_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call OPSComboDropDown(vsOPS, Row, Col)
End Sub

Private Sub vsOPS_DblClick()
    Call OPSDblClick(vsOPS)
End Sub

Private Sub vsOPS_KeyDown(KeyCode As Integer, Shift As Integer)
    Call OPSKeyDown(vsOPS, KeyCode, Shift)
End Sub

Private Sub vsOPS_KeyPress(KeyAscii As Integer)
    Call OPSKeyPress(vsOPS, KeyAscii)
End Sub

Private Sub vsOPS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call OPSKeyPressEdit(vsOPS, Row, Col, KeyAscii)
End Sub

Private Sub vsOPS_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call OPSSetupEditWindow(vsOPS, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsOPS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call OPSStartEdit(vsOPS, Row, Col, Cancel)
End Sub

Private Sub vsOPS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call OPSValidateEdit(vsOPS, Row, Col, Cancel)
End Sub

Private Sub vsRadioth_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call RadiothAfterEdit(vsRadioth, Row, Col)
End Sub

Private Sub vsRadioth_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call RadiothAfterRowColChange(vsRadioth, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsRadioth_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsRadioth)
End Sub

Private Sub vsRadioth_GotFocus()
    Call VSFlxGotFocus(vsRadioth)
End Sub

Private Sub vsRadioth_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RadiothKeyDown(vsRadioth, KeyCode, Shift)
End Sub

Private Sub vsRadioth_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Call RadiothKeyDownEdit(vsRadioth, Row, Col, KeyCode, Shift)
End Sub

Private Sub vsRadioth_KeyPress(KeyAscii As Integer)
    Call RadiothKeyPress(vsRadioth, KeyAscii)
End Sub

Private Sub vsRadioth_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call RadiothKeyPressEdit(vsRadioth, Row, Col, KeyAscii)
End Sub

Private Sub vsRadioth_LostFocus()
    Call RadiothLostFocus(vsRadioth)
End Sub

Private Sub vsRadioth_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call RadiothValidateEdit(vsRadioth, Row, Col, Cancel)
End Sub

Private Sub vsTransfer_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsTransferAfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsTransfer_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call vsTransferCellButtonClick(Row, Col)
End Sub

Private Sub vsTransfer_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange
End Sub

Private Sub vsTransfer_KeyDown(KeyCode As Integer, Shift As Integer)
    Call vsTransferKeyDown(KeyCode, Shift)
End Sub

Private Sub vsTransfer_KeyPress(KeyAscii As Integer)
    Call vsTransferKeyPress(KeyAscii)
End Sub

Private Sub vsTransfer_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsTransferStartEdit(Row, Col, Cancel)
End Sub

Private Sub vsTransfer_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsTransferValidateEdit(Row, Col, Cancel)
End Sub

Private Sub vsTSJC_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call TSJCAfterEdit(vsTSJC, Row, Col)
End Sub

Private Sub vsTSJC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call TSJCAfterRowColChange(vsTSJC, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsTSJC_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call TSJCCellButtonClick(vsTSJC, Row, Col)
End Sub

Private Sub vsTSJC_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsTSJC)
End Sub

Private Sub vsTSJC_GotFocus()
    Call VSFlxGotFocus(vsTSJC)
End Sub

Private Sub vsTSJC_KeyDown(KeyCode As Integer, Shift As Integer)
    Call TSJCKeyDown(vsTSJC, KeyCode, Shift)
End Sub

Private Sub vsTSJC_KeyPress(KeyAscii As Integer)
    Call TSJCKeyPress(vsTSJC, KeyAscii)
End Sub

Private Sub vsTSJC_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call TSJCKeyPressEdit(vsTSJC, Row, Col, KeyAscii)
End Sub

Private Sub vsTSJC_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call TSJCSetupEditWindow(vsTSJC, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsTSJC_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call TSJCValidateEdit(vsTSJC, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseDown(vsDiagXY, Button, Shift, X, Y)
End Sub

Private Sub vsDiagXY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseUp(vsDiagXY, Button, Shift, X, Y)
End Sub

Private Sub vsDiagZY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseDown(vsDiagZY, Button, Shift, X, Y)
End Sub

Private Sub vsDiagZY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseUp(vsDiagZY, Button, Shift, X, Y)
End Sub
