VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLR_S 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   4140
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5010
      ScaleWidth      =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1140
      Width           =   45
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6150
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReport.frx":014A
            Text            =   "????????"
            TextSave        =   "????????"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11298
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "????"
            TextSave        =   "????"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "????"
            TextSave        =   "????"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmReport.frx":09DE
      Height          =   5175
      LargeChange     =   20
      Left            =   9225
      Max             =   100
      SmallChange     =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   250
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmReport.frx":0CE8
      Height          =   250
      LargeChange     =   20
      Left            =   4185
      Max             =   100
      SmallChange     =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5895
      Width           =   4995
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   5145
      Left            =   4230
      ScaleHeight     =   5085
      ScaleWidth      =   4950
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   735
      Width           =   5010
      Begin VSFlex8Ctl.VSFlexGrid msh 
         Height          =   1575
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   3135
         _cx             =   1989547418
         _cy             =   1989544666
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "frmReport.frx":0FF2
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16777215
         ForeColorFixed  =   0
         BackColorSel    =   10251637
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   4005
         ScaleHeight     =   765
         ScaleWidth      =   330
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1815
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Index           =   0
         Left            =   255
         ScaleHeight     =   3390
         ScaleWidth      =   3315
         TabIndex        =   6
         Top             =   165
         Width           =   3315
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   -8888
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   30
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   330
         ScaleHeight     =   3390
         ScaleWidth      =   3315
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   3315
      End
   End
   Begin VB.PictureBox picGroup 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   0
      ScaleHeight     =   5010
      ScaleWidth      =   4140
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "Save"
      Top             =   1140
      Width           =   4140
      Begin VB.PictureBox picPar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F4F4F4&
         Height          =   3090
         Left            =   45
         ScaleHeight     =   3030
         ScaleWidth      =   4050
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "Save"
         Top             =   2325
         Width           =   4110
         Begin VB.CommandButton cmdSelAll 
            Caption         =   "????"
            Height          =   350
            Left            =   120
            TabIndex        =   34
            Top             =   930
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmdSelNone 
            Cancel          =   -1  'True
            Caption         =   "????"
            Height          =   350
            Left            =   765
            TabIndex        =   33
            Top             =   930
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmdLoad 
            BackColor       =   &H00F4F4F4&
            Caption         =   "????(&O)"
            Height          =   350
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   930
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdDefault 
            BackColor       =   &H00F4F4F4&
            Caption         =   "????(&D)"
            Height          =   350
            Left            =   2850
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   930
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.Frame fraGroup 
            BackColor       =   &H00F4F4F4&
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   16
            Top             =   -60
            Visible         =   0   'False
            Width           =   3825
         End
         Begin VB.Frame fra 
            BackColor       =   &H00F4F4F4&
            ForeColor       =   &H00800000&
            Height          =   645
            Index           =   0
            Left            =   210
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   3825
            Begin VB.OptionButton opt 
               BackColor       =   &H00F4F4F4&
               Caption         =   "#"
               Height          =   180
               Index           =   0
               Left            =   105
               MaskColor       =   &H8000000F&
               TabIndex        =   18
               Top             =   270
               Visible         =   0   'False
               Width           =   1150
            End
         End
         Begin VB.CheckBox chk 
            BackColor       =   &H00F4F4F4&
            Caption         =   "#"
            Height          =   195
            Index           =   0
            Left            =   1455
            TabIndex        =   23
            Top             =   255
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cbo 
            BackColor       =   &H00F4F4F4&
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   21
            Top             =   195
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00F4F4F4&
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   20
            Top             =   195
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00F4F4F4&
            Caption         =   "??"
            Height          =   240
            Index           =   0
            Left            =   4425
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "?? F2 ??????????"
            Top             =   225
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   22
            Top             =   195
            Visible         =   0   'False
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16053492
            CalendarTitleBackColor=   12946264
            CalendarTitleForeColor=   16053492
            CustomFormat    =   "yyyy??MM??dd?? HH:mm:ss"
            Format          =   380502019
            CurrentDate     =   36731
         End
         Begin VB.Frame fraSplit 
            BackColor       =   &H00F4F4F4&
            Height          =   75
            Left            =   -180
            TabIndex        =   25
            Top             =   750
            Visible         =   0   'False
            Width           =   10000
         End
         Begin VB.Label lblName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "????????"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   675
            TabIndex        =   24
            Top             =   255
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1845
         Left            =   45
         TabIndex        =   11
         Tag             =   "Save"
         Top             =   225
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3254
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "????"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "????"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "????"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblPar_S 
         BackColor       =   &H009B6737&
         Caption         =   " ????????"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         MousePointer    =   7  'Size N S
         TabIndex        =   13
         Top             =   2100
         Width           =   4080
      End
      Begin VB.Label lblGroup_S 
         BackColor       =   &H009B6737&
         Caption         =   " ??????"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   12
         Top             =   15
         Width           =   4095
      End
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   2011
      _CBWidth        =   9480
      _CBHeight       =   1140
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   4500
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Caption2        =   "????"
      Child2          =   "cboFormat"
      MinWidth2       =   2505
      MinHeight2      =   315
      Width2          =   4005
      NewRow2         =   0   'False
      Caption3        =   "????"
      Child3          =   "txtFind"
      MinWidth3       =   1005
      MinHeight3      =   330
      Width3          =   1935
      NewRow3         =   0   'False
      Begin VB.TextBox txtFind 
         Height          =   330
         Left            =   585
         TabIndex        =   32
         Top             =   780
         Width           =   8805
      End
      Begin MSComctlLib.ImageCombo cboFormat 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   225
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   16053492
         Locked          =   -1  'True
         ImageList       =   "img16"
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Preview"
               Description     =   "????"
               Object.ToolTipText     =   "????????"
               Object.Tag             =   "????"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Print"
               Description     =   "????"
               Object.ToolTipText     =   "????"
               Object.Tag             =   "????"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Graph"
               Description     =   "????"
               Object.ToolTipText     =   "??????????????????????"
               Object.Tag             =   "????"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Par"
               Description     =   "????"
               Object.ToolTipText     =   "????????"
               Object.Tag             =   "????"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Par_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "ColWidth"
               Description     =   "????"
               Object.ToolTipText     =   "????"
               Object.Tag             =   "????"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Object.Tag             =   "????????"
                     Text            =   "????????"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Fill"
                     Object.Tag             =   "????????"
                     Text            =   "????????"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Def"
                     Object.Tag             =   "????????"
                     Text            =   "????????"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "SelMode"
               Description     =   "????"
               Object.ToolTipText     =   "????????????????"
               Object.Tag             =   "????"
               ImageKey        =   "SelMode"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RowMode"
                     Object.Tag             =   "????????"
                     Text            =   "????????"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ColMode"
                     Object.Tag             =   "????????"
                     Text            =   "????????"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Style"
               Object.ToolTipText     =   "??????????????????"
               Object.Tag             =   "????"
               ImageKey        =   "Style"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Large"
                     Object.Tag             =   "??????"
                     Text            =   "??????"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "??????"
                     Text            =   "??????"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "????"
                     Text            =   "????"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "????????"
                     Text            =   "????????"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Style_"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Pre"
               Description     =   "????"
               Object.ToolTipText     =   "????????????????(Page Up)"
               Object.Tag             =   "????"
               ImageKey        =   "Pre"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Next"
               Description     =   "????"
               Object.ToolTipText     =   "????????????????(Page Down)"
               Object.Tag             =   "????"
               ImageKey        =   "Next"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Page_"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Help"
               Description     =   "????"
               Object.ToolTipText     =   "????????????"
               Object.Tag             =   "????"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "????"
               Key             =   "Quit"
               Description     =   "????"
               Object.ToolTipText     =   "????"
               Object.Tag             =   "????"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   705
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":18CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2134
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":234E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2568
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2782
            Key             =   "Style"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":299C
            Key             =   "Pre"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2BB6
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2DD0
            Key             =   "SelMode"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   75
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3204
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":341E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3638
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3852
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3A6C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3EA0
            Key             =   "Style"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":40BA
            Key             =   "Pre"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":42D4
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":44EE
            Key             =   "SelMode"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timHead 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSScriptControlCtl.ScriptControl Srt 
      Left            =   6855
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   2745
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":4708
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2100
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":4A22
            Key             =   "Format"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":4B7C
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin C1Chart2D8.Chart2D Chart 
      Height          =   1230
      Index           =   0
      Left            =   4275
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   1650
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   2910
      _ExtentY        =   2170
      _StockProps     =   0
      ControlProperties=   "frmReport.frx":4CD6
   End
   Begin VB.Image imgCode 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   4230
      Stretch         =   -1  'True
      Top             =   2415
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   4230
      Stretch         =   -1  'True
      Top             =   2415
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Line lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   4380
      X2              =   5655
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Shape Shp 
      FillColor       =   &H80000005&
      Height          =   315
      Index           =   0
      Left            =   4365
      Top             =   1995
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4380
      MouseIcon       =   "frmReport.frx":5335
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "????(&F)"
      Begin VB.Menu mnuFile_Setup 
         Caption         =   "????????(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_Preview 
         Caption         =   "????????(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "????????(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "Excel????????(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile_Graph 
         Caption         =   "Excel????????(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "????(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "????(&E)"
      Begin VB.Menu mnuEdit_Par 
         Caption         =   "????????(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuEdit_Par_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SetCol 
         Caption         =   "????????(&C)"
         Begin VB.Menu mnuEdit_SetCol_Auto 
            Caption         =   "????????(&A)"
         End
         Begin VB.Menu mnuEdit_SetCol_Fill 
            Caption         =   "????????(&I)"
         End
         Begin VB.Menu mnuEdit_SetCol_Def 
            Caption         =   "????????(&D)"
         End
      End
      Begin VB.Menu mnuEdit_SelMode 
         Caption         =   "????????(&S)"
         Begin VB.Menu mnuEdit_SelMode_Row 
            Caption         =   "????????(&R)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuEdit_SelMode_Col 
            Caption         =   "????????(&C)"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "????(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "??????(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "????????(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolFormat 
            Caption         =   "????????(&F)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolGroup 
            Caption         =   "??????(&G)"
            Checked         =   -1  'True
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "????????(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "??????(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "??????(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "??????(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "????(&L)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "????????(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuViewStyle_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_Pre 
         Caption         =   "??????(&P)"
      End
      Begin VB.Menu mnuView_Next 
         Caption         =   "??????(&N)"
      End
      Begin VB.Menu mnuView_Page_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_reFlash 
         Caption         =   "????(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "????(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "????????(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "WEB????????"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "????????(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "????????(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "????????(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "????(&A)..."
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "????????"
      Visible         =   0   'False
      Begin VB.Menu mnuPop_Cond 
         Caption         =   "????1"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop_Split1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop_Save 
         Caption         =   "????(&S)"
      End
      Begin VB.Menu mnuPop_SaveAs 
         Caption         =   "??????(&A)"
      End
      Begin VB.Menu mnuPop_Del 
         Caption         =   "????(&C)"
      End
      Begin VB.Menu mnuPop_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop_Default 
         Caption         =   "????(&D)"
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Option Compare Text
'??????????----------------------------------------------------------
Public mobjCurDLL As clsReport '??????????????????????????????????????(DLL:clsReport,????????????????)
Public mbytStyle As Byte    '??????????????????????

Public mblnDisabledPrint As Boolean     '????????????????
Public mblnPrintEmpty As Boolean '????????????????????????
Public bytFormat As Byte '????????????????????????
Public marrPars As Variant '????????????,????????????????

Public frmParent As Object '????????
Public mobjReport As Report '????????????(??????????????????)

Private arrReport() As Report '????????,??????????????????????
Private arrLibDatas() As LibDatas '????????????????????????????(????????????,????????????????????????,??????????????)
Private arrDefPars() As RPTPars '??????????????????????????????

Public intReport As Integer '>=0,????????????????????,????picPaper??????

'??????????(????????????????Excel??????????)-------------------------
Public mLibDatas As LibDatas '??????????????????????????,??????????????
Public marrPage As Variant   '????PageCells??????????,????????????????,??????????????
Public marrPageCard As Variant   '??;??????????
Public mcolRowIDs As New Collection '????????????????????????ID(ID????????????????,??????????????????)

'????????------------------------------------------------------------
Private mstrExcelFile As String
Private mblnAllFormat As Boolean
Private lngPreX As Long, lngPreY As Long
Private intGridCount As Integer '??????????????????????????(????????????????)
Private intGridID As Integer '????????????????????,??????????ID
Private objCurGrid As Object
Private mobjPars As RPTPars '??????????????????????????
Private mobjDefPars As RPTPars '??????????????????????????,??????????????
Private objScript As clsScript
Private blnMatch As Boolean, blnExcel As Boolean
Private blnRefresh As Boolean
Private lngCurInx As Long
Private lngTmpColor As Long
Private mstrPDFFile As String
Private mobjfrmShow As frmPreview
Private mlngRPTID As Long               '??????????ID??????????ID
Private mintCurMenuIndex As Integer
Private mintCurCondID As Integer

Private Const CON_SETFOCES As Long = &H9C6D75

Public Sub ShowMe(objParent As Object, objCurDLL As clsReport, arrPars As Variant, ByVal bytStyle As Byte)
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    
    On Error Resume Next
    
    If mbytStyle <> 0 Then
        Load Me
        If Err.Number = 0 Then
            If mbytStyle = 1 Then '????????
                mnuFile_Preview_Click
            ElseIf mbytStyle = 2 Then '????????
                mnuFile_Print_Click
            ElseIf mbytStyle = 3 Then '??????Excel
                mnuFile_Excel_Click
            ElseIf mbytStyle = 4 Then '??????????PDF
                mnuFile_Print_Click
            End If
        ElseIf Err.Number <> 0 Then
            '364:??????????(??Form_Load????Unload,??????????????)
            Err.Clear
        End If
        Unload Me
    Else
        '??????????????????????
        If frmParent Is Nothing Then
            Me.Show
        ElseIf frmParent.name = "frmDesign" Then
            Me.Show 1, frmParent
        Else
            Me.Show , frmParent
        End If
        
        '????????????????????????
        If Err.Number = 373 Or Err.Number = 401 Then
            '373:??????????????????????????????????(??????????zlReport.dll,??????????????)
            '401:????????????????????????????????????
            '??????Load????????????????????Form_Load????
            Err.Clear: Me.Show 1
        ElseIf Err.Number = 364 Then
            '364:??????????(??Form_Load????Unload,??????????????)
            Err.Clear
        ElseIf Err.Number <> 0 Then
            Err.Clear: Unload Me '??????Load????????????????????
        End If
    End If
End Sub

Private Sub CopyLibDatas(objS As LibDatas, objO As LibDatas)
'??????????????????????????????????
    Dim tmpData As LibData
    
    Set objO = New LibDatas
    
    For Each tmpData In objS
        objO.Add tmpData.DataName, tmpData.DataSet.Clone, "_" & tmpData.DataName
    Next
End Sub

Private Sub CboFormat_Click()
    Dim strErr As String
    
    If CByte(Mid(cboFormat.SelectedItem.Key, 2)) = bytFormat Then Exit Sub
    bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
    mobjReport.bytFormat = bytFormat
    
    mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).???? <> 0)
    tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).???? <> 0)
    
    If mobjReport.blnLoad Then
        '??????????????????????(??????????????????????????????????)
        strErr = OpenReportData(False)
        If strErr <> "" Then
            MsgBox "??????????????""" & strErr & """??????????????,??????????????", vbInformation, App.Title
            Exit Sub
        End If

        '????????????????
        If Not mobjCurDLL Is Nothing Then
            mobjCurDLL.Act_CommitCondition mobjReport.????, GetParsStr(MakeNamePars(mobjReport, True)), Me
        End If
        
        timHead.Enabled = False
        Call ShowItems
        timHead.Enabled = True
    End If
End Sub

Private Sub CboFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set cboFormat.SelectedItem = cboFormat.ComboItems("_" & bytFormat)
        KeyAscii = 0
    End If
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Chart_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngX As Single, sngY As Single
    Dim lngSeries As Long, lngPoint As Long, lngDS As Long
    Dim strSeries As String, vArea As RegionConstants
    Dim dblX As Double, strX As String, strY As String
    Dim strLabelX As String, strLabelY As String
    
    With Chart(Index).ChartGroups(1)
        sngX = X / Screen.TwipsPerPixelX
        sngY = Y / Screen.TwipsPerPixelY
        vArea = .CoordToDataIndex(sngX, sngY, oc2dFocusXY, lngSeries, lngPoint, lngDS)
        If vArea = oc2dRegionInChartArea Then
            If lngDS <= 3 Then
                strSeries = ""
                If lngSeries <= .SeriesLabels.count Then
                    strSeries = .SeriesLabels(lngSeries).Text & ":"
                End If
                                
                If .Data.Layout = oc2dDataGeneral Then
                    dblX = .Data.X(lngSeries, lngPoint)
                Else
                    dblX = .Data.X(1, lngPoint)
                End If
                
                If Chart(Index).ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels Then '??1970-01-01 08:00:00??????????
                    strX = Format(DateAdd("s", dblX, CDate("1970-01-01 08:00:00")), "yyyy-MM-dd HH:mm:ss")
                    strX = Replace(strX, " 00:00:00", "")
                    strX = Replace(strX, ":00:00", "")
                    strX = Replace(strX, ":00", "")
                Else
                    strX = dblX
                End If
                strY = .Data.Y(lngSeries, lngPoint)
                
                If Chart(Index).ChartArea.Axes("X").Title.Text <> "" Then
                    strLabelX = Chart(Index).ChartArea.Axes("X").Title.Text & "="
                End If
                If Chart(Index).ChartArea.Axes("Y").Title.Text <> "" Then
                    strLabelY = Chart(Index).ChartArea.Axes("Y").Title.Text & "="
                End If
                
                sta.Panels(3).Text = strSeries & strLabelX & strX & "," & strLabelY & strY
            Else
                sta.Panels(3).Text = ""
            End If
        Else
            sta.Panels(3).Text = ""
        End If
    End With
End Sub

Private Sub cmdDefault_Click()
    Dim sngTop As Single
    
    sngTop = cmdDefault.Top + cmdDefault.Height + picPar.Top + IIF(cbr.Visible, cbr.Height, 0) + 15
    Call Me.PopupMenu(mnuPop, , cmdDefault.Left + 30, sngTop)
End Sub

Private Sub cmdLoad_Click()
    mnuView_reFlash_Click
End Sub

Private Sub cmdSelAll_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        chkTmp.Value = 1
    Next
End Sub

Private Sub cmdSelNone_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        chkTmp.Value = 0
    Next
End Sub

Private Sub Form_Activate()
    Dim tmpMsh As Object
    Static blnAct As Boolean
    
    If blnExcel Then blnExcel = False: Exit Sub
    
    cbr.Bands(2).Width = cbr.Bands(2).Width + 15
    cbr.Bands(2).Width = cbr.Bands(2).Width - 15
    
    '????????
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_ReportActive(mobjReport.????, Me)
    End If
    
    If cbr.Bands(1).Visible Then cbr.Bands(1).MinHeight = tbr.ButtonHeight

    '??????????????????
    If Not blnAct Then
        blnAct = True
        For Each tmpMsh In msh
            If tmpMsh.Index <> 0 And tmpMsh.Container Is picPaper(intReport) And Not tmpMsh.Tag Like "H_*" Then
                Call msh_EnterCell(tmpMsh.Index)
                On Error Resume Next
                tmpMsh.SetFocus: Exit For
            End If
        Next
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (scrVsc.Visible And scrHsc.Visible) And KeyCode <> vbKeyF3 Then Exit Sub
    Select Case KeyCode
        Case vbKeyUp
            If scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then
                If Shift = 2 Then
                    scrVsc.Value = IIF(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIF(scrVsc.Value - scrVsc.SmallChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.SmallChange)
                End If
            End If
        Case vbKeyDown
            If scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then
                If Shift = 2 Then
                    scrVsc.Value = IIF(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIF(scrVsc.Value + scrVsc.SmallChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.SmallChange)
                End If
            End If
        Case vbKeyLeft
            If scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then
                If Shift = 2 Then
                    scrHsc.Value = IIF(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIF(scrHsc.Value - scrHsc.SmallChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.SmallChange)
                End If
            End If
        Case vbKeyRight
            If scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then
                If Shift = 2 Then
                    scrHsc.Value = IIF(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIF(scrHsc.Value + scrHsc.SmallChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.SmallChange)
                End If
            End If
        Case vbKeyF3
            Call FindItem(txtFind.Text, True)
    End Select
End Sub

Private Sub Form_Load()
    Dim strErr As String, i As Integer, j As Integer
    Dim objItem As Object, rsTmp As ADODB.Recordset
    Dim strPrivs As String, lng????ID As Long, lng????ID As Long
    Dim blnPriv As Boolean, bytMode As Byte
    Dim strSQL As String, lngReport As Long
    Dim rsReport As New ADODB.Recordset
    Dim frmNewParInput As New frmParInput
    Dim strBasePrivs As String
    Dim strTmp As String
    
    Set objScript = New clsScript
    Srt.AddObject "clsScript", objScript, True
    
    garrBill = Empty
    
    mblnPrintEmpty = False
    bytFormat = 0
    
    blnExcel = False
    timHead.Enabled = False
    
    '????????????
    If gobjReport Is Nothing Then
        '??????????
        '??????????????
        Set rsTmp = GetGroupInfo(glngGroup)
        If rsTmp Is Nothing Then Unload Me: Exit Sub '????????
        Caption = rsTmp!????
        lblGroup_S.Caption = lblGroup_S.Caption & ":" & rsTmp!????
        Me.Tag = rsTmp!???? '??????????????
        
        lng????ID = IIF(IsNull(rsTmp!????), 0, rsTmp!????)
        lng????ID = IIF(IsNull(rsTmp!????ID), 0, rsTmp!????ID)
        
        '????????????
        bytMode = GetSetting("ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & Me.Tag, "????????", 0)
        mnuEdit_SelMode_Row.Checked = (bytMode = 0)
        mnuEdit_SelMode_Col.Checked = (bytMode = 1)
        
        '??????????????
        Set rsTmp = GetSubReport(glngGroup)
        If rsTmp Is Nothing Then
            MsgBox "????????????????????????????????", vbInformation, App.Title
            Unload Me: Exit Sub '????????
        End If
        Screen.MousePointer = 11
        strPrivs = GetPrivFunc(lng????ID, lng????ID)
        i = 0
        Do While Not rsTmp.EOF
            '????????????????????????????????????????????????????
            If InStr(";" & strPrivs & ";", ";" & Nvl(rsTmp!????, "NONE") & ";") <= 0 _
                And Not mobjCurDLL Is Nothing Then
                GoTo makContinue
            End If
            
            blnPriv = True
            '??????????
            blnPriv = CheckPass(rsTmp!????ID)
            '????????
            If lng????ID > 0 And Not IsNull(rsTmp!????) And blnPriv Then
                blnPriv = (InStr(";" & strPrivs & ";", ";" & rsTmp!???? & ";") > 0)
            End If
            If blnPriv Then
                If i = 0 Then
                    ReDim arrReport(0)
                    ReDim arrLibDatas(0) '??????????????
                    ReDim arrDefPars(0)
                Else
                    Load picPaper(i): picPaper(i).Visible = False
                    ReDim Preserve arrReport(i)
                    ReDim Preserve arrLibDatas(i) '??????????????
                    ReDim Preserve arrDefPars(i)
                End If
                
                '????????
                Set arrReport(i) = New Report
                Set arrReport(i) = ReadReport(rsTmp!????ID)
                Call ReplaceSysNo(arrReport(i)) '????????????????????????
                Call GetUserName(arrReport(i).????, gstrUserName, gstrUserNO)
                Call SetReportIndex(i, arrReport(i))
                
                '????????????
                Set arrDefPars(i) = New RPTPars
                Set arrDefPars(i) = MakeNamePars(arrReport(i))
                
                Set objItem = lvw.ListItems.Add(, "_" & rsTmp!????ID, arrReport(i).????, "Report", "Report")
                objItem.SubItems(1) = arrReport(i).????
                objItem.SubItems(2) = arrReport(i).????
                
                 '????????????????????????????????
                If arrReport(i).Datas.count > 0 Then Call DelUnUseData(arrReport(i))
                '??????????????????????????
                If ParCount(arrReport(i)) > 0 Then Call ReplaceUserPars(arrReport(i))
                
                '??????????????????,????1
                arrReport(i).bytFormat = CByte(GetSetting("ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & arrReport(i).????, "????", 1))
                
                i = i + 1
            End If
            
makContinue:
            rsTmp.MoveNext
        Loop
        
        Screen.MousePointer = 0
        If rsTmp.RecordCount > 0 And lvw.ListItems.count = 0 Then
            MsgBox "????????????????????????????????????????????????????????????????????????", vbInformation, App.Title
            Unload Me: Exit Sub '????????
        ElseIf lvw.ListItems.count = 0 Then
            Unload Me: Exit Sub '????????
        End If
        
        mnuEdit_Par.Visible = False
        mnuEdit_Par_.Visible = False
        tbr.Buttons("Par").Caption = "????"
        tbr.Buttons("Par").Tag = "????"
        
        lvw.ColumnHeaders(2).Position = 1
        RestoreWinState Me, App.ProductName, Me.Tag

        SetView lvw.View
        
        '????????????????????????????????????????
        lvw.Height = lvw.ListItems.count * 350
        If lvw.Height < 1000 Then lvw.Height = 1000
        If lvw.Height > picGroup.Height / 2 Then
            lvw.Height = picGroup.Height / 2
        End If
        
        picLR_S.Visible = mnuViewToolGroup.Checked
        picGroup.Visible = mnuViewToolGroup.Checked
        
        If Not lvw.SelectedItem Is Nothing Then Call lvw_ItemClick(lvw.SelectedItem)
    Else
        '????????????
        picBack.BorderStyle = 0
        picLR_S.Visible = False
        picGroup.Visible = False
        For i = 0 To mnuViewStyle.UBound
            mnuViewStyle(i).Visible = False
        Next
        mnuViewStyle_.Visible = False
        mnuView_Pre.Visible = False
        mnuView_Next.Visible = False
        mnuView_Page_.Visible = False
        mnuViewToolGroup.Visible = False

        tbr.Buttons("Style").Visible = False
        tbr.Buttons("Style_").Visible = False
        tbr.Buttons("Pre").Visible = False
        tbr.Buttons("Next").Visible = False
        tbr.Buttons("Page_").Visible = False
        
        intReport = 0
        Call CopyReport(gobjReport, mobjReport)
        Call ReplaceSysNo(mobjReport) '????????????????????????
        Call GetUserName(mobjReport.????, gstrUserName, gstrUserNO)
        Call SetReportIndex(intReport, mobjReport)
        Caption = mobjReport.????
        
        If mbytStyle = 0 Then '??????????????????????????????
            RestoreWinState Me, App.ProductName, mobjReport.????
        End If
        
        '????????????????????????????????
        If Format(mobjReport.????????????, "HH:mm:ss") <> "00:00:00" Or Format(mobjReport.????????????, "HH:mm:ss") <> "00:00:00" Then
            If CDate(Format(mobjReport.????????????, "HH:mm:ss")) > CDate(Format(mobjReport.????????????, "HH:mm:ss")) Then
                If Between(CDate(Format(Currentdate, "HH:mm:ss")), CDate(Format(mobjReport.????????????, "HH:mm:ss")), CDate(Format(mobjReport.????????????, "HH:mm:ss"))) Then
                    MsgBox "??????????" & CDate(Format(mobjReport.????????????, "HH:mm:ss")) & "-" & CDate(Format(mobjReport.????????????, "HH:mm:ss")) & "????????????????????????????????", vbInformation, App.Title
                    Unload Me: Exit Sub
                End If
            Else
                If CDate(Format(Currentdate, "HH:mm:ss")) < CDate(Format(mobjReport.????????????, "HH:mm:ss")) Or CDate(Format(Currentdate, "HH:mm:ss")) > CDate(Format(mobjReport.????????????, "HH:mm:ss")) Then
                    MsgBox "??????????" & CDate(Format(mobjReport.????????????, "HH:mm:ss")) & "-??????" & CDate(Format(mobjReport.????????????, "HH:mm:ss")) & "????????????????????????????????", vbInformation, App.Title
                    Unload Me: Exit Sub
                End If
            End If
        End If
        
        '????????????
        bytMode = GetSetting("ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.????, "????????", 0)
        mnuEdit_SelMode_Row.Checked = (bytMode = 0)
        mnuEdit_SelMode_Col.Checked = (bytMode = 1)
    
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_BeforeReportLoad(mobjReport.????, Me)
        End If
    
         '????????????????????????????????
        If mobjReport.Datas.count > 0 Then Call DelUnUseData(mobjReport)
    
        '??????????????????
        bytFormat = 1
        
        '????????????????????????????????
        '??????????????????????????????????????????????????????????????
        strTmp = GetSetting("ZLSOFT", "????????\" & App.ProductName & "\LocalSet\" & mobjReport.????, "AllFormat", "")
        If strTmp = "" Then strTmp = GetSetting("ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & mobjReport.????, "AllFormat", 0)
        mblnAllFormat = Val(strTmp) = 1
        If Not (mbytStyle = 2 And mblnAllFormat) Then
            '????????????????????
            bytFormat = CByte(GetSetting("ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.????, "????", 1))
            '?????????????????????? 'If mobjReport.???? Then
            strTmp = GetSetting("ZLSOFT", "????????\" & App.ProductName & "\LocalSet\" & mobjReport.????, "Format", "")
            If strTmp = "" Then
                i = Val(GetSetting("ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & mobjReport.????, "Format", -1))
            Else
                i = Val(strTmp)
            End If
            If i <> -1 Then bytFormat = i
        End If
        
        '??????????ID
        lngReport = 0
        If ReportReaded(, mobjReport.????, mobjReport.????) Then
            lngReport = grsReport!ID '????????
        Else
            strSQL = "Select ID,????,????,????,????,??????,????,????,????????,????,????ID,????,????????,????????,????????????,???????????? From zlReports Where ????=[1] And Nvl(????,0)=[2]"
            Set rsReport = OpenSQLRecord(strSQL, Me.Caption, mobjReport.????, mobjReport.????)
            If Not rsReport.EOF Then '????????
                Set grsReport = New ADODB.Recordset
                Set grsReport = rsReport
                gdatModiTime = grsReport!????????
                
                lngReport = rsReport!ID
            End If
        End If
        mlngRPTID = lngReport
        
        '??????????????????????????????
        If IsArray(marrPars) Then
            If UBound(marrPars) <> -1 Then
                For i = 0 To UBound(marrPars)
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        'ReportFormat
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("ReportFormat") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                bytFormat = CByte(Trim(Mid(CStr(marrPars(i)), j + 1)))
                                mblnAllFormat = False '????????????????????????????
                            End If
                        'DisabledPrint
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("DisabledPrint") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                mblnDisabledPrint = CByte(Trim(Mid(CStr(marrPars(i)), j + 1))) = 1
                            End If
                        'PrintEmpty
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("PrintEmpty") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                mblnPrintEmpty = CByte(Trim(Mid(CStr(marrPars(i)), j + 1))) = 1
                            End If
                        'ExcelFile
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("ExcelFile") Then
                            mstrExcelFile = Trim(Mid(CStr(marrPars(i)), j + 1))
                        'PDF
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("PDF") Then
                            mstrPDFFile = Trim(Mid(CStr(marrPars(i)), j + 1))
                        End If
                    End If
                Next
            End If
        End If
        
        '????????????
        For i = 1 To mobjReport.Fmts.count
            Set objItem = cboFormat.ComboItems.Add(, "_" & mobjReport.Fmts(i).????, mobjReport.Fmts(i).????, "Format")
            If mobjReport.Fmts(i).???? = bytFormat Then objItem.Selected = True
        Next
        If cboFormat.SelectedItem Is Nothing And cboFormat.ComboItems.count > 0 Then
            cboFormat.ComboItems(1).Selected = True
            bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
        End If
        mobjReport.bytFormat = bytFormat
        mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).???? <> 0)
        tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).???? <> 0)
                
'        If cboFormat.ComboItems.count = 1 Then
'            mnuViewToolFormat.Checked = False
'            cbr.Bands(2).Visible = False
'        End If
        cboFormat.Locked = cboFormat.ComboItems.count > 1
                
        '????????
        If ParCount(mobjReport) > 0 Then
            If Not ReplaceUserPars(mobjReport) Then
                '????????????????????,??????????????
                
                Set mobjPars = MakeNamePars(mobjReport)
                Call CopyPars(mobjPars, mobjDefPars)
                frmNewParInput.mlngReport = lngReport
                Set frmNewParInput.mobjPars = mobjPars
                Set frmNewParInput.mobjDefPars = mobjDefPars
                Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
                
                frmNewParInput.mstrTitle = mobjReport.????
                frmNewParInput.mblnReset = False
                frmNewParInput.Show 1, Me
                If frmNewParInput.mblnOK Then
                    '????????????????
                    If Not mobjCurDLL Is Nothing Then
                        mobjCurDLL.Act_CommitCondition mobjReport.????, GetParsStr(frmNewParInput.mobjPars), Me
                    End If
                    
                    ReplaceInputPars frmNewParInput.mobjPars
                    Unload frmNewParInput
                Else
                    Unload Me: Exit Sub '????????????????
                End If
            Else
                '????????????????,????????????????????
                tbr.Buttons("Par").Visible = False
                mnuEdit_Par.Visible = False
                tbr.Buttons("Par_").Visible = False
                mnuEdit_Par_.Visible = False
                
                Set mobjDefPars = MakeNamePars(mobjReport)
                
                '????????????????
                If Not mobjCurDLL Is Nothing Then
                    mobjCurDLL.Act_CommitCondition mobjReport.????, GetParsStr(MakeNamePars(mobjReport, True)), Me
                End If
            End If
        Else
            '????????????????
            If Not mobjCurDLL Is Nothing Then
                mobjCurDLL.Act_CommitCondition mobjReport.????, "ReportFormat=" & bytFormat, Me
            End If
            
            '????????????????,????????????????????
            tbr.Buttons("Par").Visible = False
            mnuEdit_Par.Visible = False
            tbr.Buttons("Par_").Visible = False
            mnuEdit_Par_.Visible = False
        End If
        
        '??????????????????????????????????????????????????(mobjReport)
        '????????
        If Not frmParent Is Nothing Then frmParent.Refresh
        Me.Refresh
        strErr = OpenReportData(False)
        If strErr <> "" Then
            MsgBox "??????????????""" & strErr & """??????????????,??????????????", vbInformation, App.Title
            Unload Me: Exit Sub
        End If
        
        '????????
        Call ShowItems
    
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_AfterReportLoad(mobjReport.????, Me)
        End If
    End If
    
    'Excel??????????????????
    strBasePrivs = GetPrivFunc(0, 16)
    If InStr(";" & strBasePrivs & ";", ";Excel????;") = 0 Then
        mnuFile_Excel.Visible = False
        mnuFile_Graph.Visible = False
        mnuFile_1.Visible = False
        tbr.Buttons("Graph").Visible = False
        tbr.Buttons(5).Visible = False
    End If
    If InStr(";" & strBasePrivs & ";", ";????;") = 0 Or mblnDisabledPrint Then
        mnuFile_Print.Visible = False
        tbr.Buttons("Print").Visible = False
    End If
    
    timHead.Enabled = True
End Sub

Private Sub lbl_Click(Index As Integer)
    Dim objRelations As RPTRelations
    Dim i As Long
    Dim lngRec As Long, strDataName As String
    Dim strLisName As String
    Dim tmpData As RPTData, tmpPar As RPTPar
    
    If lbl(Index).Tag <> "" Then
        Set objRelations = mobjReport.Items("_" & Index).Relations
        lngRec = Val(lbl(Index).Tag)
        
        If Not CheckReportPriv(objRelations.Item(1).????????ID) Then
            MsgBox "????????????????????????????????????????", vbInformation, App.Title: Exit Sub
        End If
        '????????
        If CheckPass(objRelations.Item(1).????????ID) = False Then
            MsgBox "??????????????????????????????", vbInformation, App.Title: Exit Sub
        End If
        
        Set gobjReport = ReadReport(objRelations.Item(1).????????ID)
        '??????????
        garrPars = Array()
        '??????????
        On Error Resume Next
        For i = 1 To objRelations.count
            If InStr(objRelations.Item(i).??????????, ".") > 0 Then
                strDataName = Mid(objRelations.Item(i).??????????, 1, InStr(objRelations.Item(i).??????????, ".") - 1)
            End If
            If strDataName <> "" Then Exit For
        Next

        '????????????????
        If strDataName <> "" Then mLibDatas("_" & strDataName).DataSet.AbsolutePosition = lngRec
        
        For i = 1 To objRelations.count
            With objRelations.Item(i)
                strLisName = ""
                If InStr(.??????????, ".") > 0 Then
                    If mLibDatas("_" & strDataName).DataSet.RecordCount > 0 Then
                        strLisName = mLibDatas("_" & strDataName).DataSet.Fields(Mid(.??????????, InStr(.??????????, ".") + 1)).Value
                    End If
                ElseIf InStr(.??????????, "=") = 1 Then
                    For Each tmpData In mobjReport.Datas
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.???? = Mid(.??????????, 2) Then
                                strLisName = tmpPar.??????
                                Exit For
                            End If
                        Next
                        If strLisName <> "" Then Exit For
                    Next
                End If
                ReDim Preserve garrPars(UBound(garrPars) + 1)
                garrPars(UBound(garrPars)) = .?????? & "=" & strLisName
            End With
        Next
        
        
        If Not ShowReport(Me) Then MsgBox "??????????????", vbInformation, App.Title
    End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl(Index).Tag <> "" Then lbl(Index).MousePointer = 99
End Sub

Private Sub lblPar_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lngPreY = Y
End Sub

Private Sub lblPar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If lblPar_S.Top + Y - lngPreY < 1000 Or picPar.Height - (Y - lngPreY) < 1000 Then Exit Sub
        lblPar_S.Top = lblPar_S.Top + Y - lngPreY
        lvw.Height = lvw.Height + Y - lngPreY
        picPar.Top = picPar.Top + Y - lngPreY
        picPar.Height = picPar.Height - (Y - lngPreY)
        Me.Refresh
    End If
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, objItem As Object
    
    Set objCurGrid = Nothing
    
    LockWindowUpdate Me.hwnd
    
    '????????????????????????
    If Not mobjReport Is Nothing Then
        Call CopyReport(mobjReport, arrReport(intReport)) '????
        Call CopyPars(mobjDefPars, arrDefPars(intReport)) '????????????
        If mLibDatas Is Nothing Then '??????????
            Set arrLibDatas(intReport) = Nothing
        Else
            Call CopyLibDatas(mLibDatas, arrLibDatas(intReport))
        End If
    End If
    
    '????????????????????
    intReport = Item.Index - 1
    Call CopyReport(arrReport(intReport), mobjReport) '????
    Call CopyPars(arrDefPars(intReport), mobjDefPars) '????????????
    If arrLibDatas(intReport) Is Nothing Then '??????????
        Set mLibDatas = Nothing
    Else
        Call CopyLibDatas(arrLibDatas(intReport), mLibDatas)
    End If
    
    bytFormat = mobjReport.bytFormat
    intGridCount = mobjReport.intGridCount
    intGridID = mobjReport.intGridID
        
    '????????
    cboFormat.ComboItems.Clear
    For i = 1 To mobjReport.Fmts.count
        Set objItem = cboFormat.ComboItems.Add(, "_" & mobjReport.Fmts(i).????, mobjReport.Fmts(i).????, "Format")
        If mobjReport.Fmts(i).???? = bytFormat Then objItem.Selected = True
    Next
    If cboFormat.SelectedItem Is Nothing And cboFormat.ComboItems.count > 0 Then
        cboFormat.ComboItems(1).Selected = True
        bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
        mobjReport.bytFormat = bytFormat
    End If
    cboFormat.Refresh
    cboFormat.Locked = cboFormat.ComboItems.count > 1
    
    mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).???? <> 0)
    tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).???? <> 0)
    
    '????????????
    picBack.Visible = False
    For i = 0 To picPaper.UBound
        picPaper(i).Visible = (i = intReport)
    Next
    picPaper(intReport).ZOrder
    
    scrVsc.Visible = Not (intGridCount = 1 And Not mobjReport.????)
    scrHsc.Visible = Not (intGridCount = 1 And Not mobjReport.????)
    picShadow.Visible = Not (intGridCount = 1 And Not mobjReport.????)
    If Not (intGridCount = 1 And Not mobjReport.????) Then
        scrVsc.Value = scrVsc.Min
        scrHsc.Value = scrHsc.Min
        Call scrhsc_Change
        Call scrVsc_Change
    End If
    
    '????????
    Call Form_Resize
    picBack.Visible = True

    '????????
    picPar.Visible = False
    
    Call CopyPars(mobjDefPars, mobjPars)
    mlngRPTID = Val(Mid(lvw.SelectedItem.Key, 2))
    Call InitReportPars
    picPar.Visible = True
    
    LockWindowUpdate 0

    '??????????????????
    For Each objItem In msh
        If objItem.Index <> 0 And objItem.Container Is picPaper(intReport) And Not objItem.Tag Like "H_*" Then
            Call msh_EnterCell(objItem.Index)
            Exit For
        End If
    Next
End Sub

Private Sub mnuEdit_SelMode_Col_Click()
    Dim tmpMsh As Object
    
    If mnuEdit_SelMode_Col.Checked Then Exit Sub
    '(????????)????????
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_LeaveCell(tmpMsh.Index)
        End If
    Next
    
    mnuEdit_SelMode_Col.Checked = True
    mnuEdit_SelMode_Row.Checked = False
    
    '(????????)????????
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
        End If
    Next
End Sub

Private Sub mnuEdit_SelMode_Row_Click()
    Dim tmpMsh As Object
    If mnuEdit_SelMode_Row.Checked Then Exit Sub
    
    '(????????)????????
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_LeaveCell(tmpMsh.Index)
        End If
    Next
    
    mnuEdit_SelMode_Row.Checked = True
    mnuEdit_SelMode_Col.Checked = False
    
    '(????????)????????
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
        End If
    Next
End Sub

Private Sub mnuEdit_SetCol_Auto_Click()
'????????????????????????????????????,????????????????????(??????)
    Dim tmpMsh As Object, i As Integer

    If Not mobjReport.blnLoad Then Exit Sub
    
    On Error Resume Next
    
    timHead.Enabled = False
    Screen.MousePointer = 11
    
    LockWindowUpdate Me.hwnd
    
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 Then
            If tmpMsh.Container Is picPaper(intReport) Then
                Call SetColWidth(tmpMsh)
            ElseIf UCase(tmpMsh.Container.name) = "PIC" Then
                If tmpMsh.Container.Container Is picPaper(intReport) Then
                    Call SetColWidth(tmpMsh)
                End If
            End If
        End If
    Next
    '??????????????????????????????
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") And tmpMsh.FixedRows = 0 Then
            For i = 0 To tmpMsh.Cols - 1
                If tmpMsh.ColWidth(i) > msh(tmpMsh.Tag).ColWidth(i) Then
                    msh(tmpMsh.Tag).ColWidth(i) = tmpMsh.ColWidth(i)
                Else
                    tmpMsh.ColWidth(i) = msh(tmpMsh.Tag).ColWidth(i)
                End If
            Next
            tmpMsh.LeftCol = 0: msh(tmpMsh.Tag).LeftCol = 0
        End If
    Next
    Screen.MousePointer = 0
    timHead.Enabled = True
    
    LockWindowUpdate 0
End Sub

Private Sub mnuEdit_SetCol_Def_Click()
'????????????????????????????????
    Dim objItem As RPTItem, objCurItem As RPTItem
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim i As Integer, j As Integer, strWidth As String
    Dim lngColB As Long, lngColE As Long
    
    If Not mobjReport.blnLoad Then Exit Sub
        
    On Error Resume Next
    
    LockWindowUpdate Me.hwnd
    
    For Each objItem In mobjReport.Items
        If objItem.?????? = bytFormat Then
            If objItem.???? = 4 Then
                With objItem
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                        msh(.ID).ColWidth(tmpItem.????) = tmpItem.W
                        msh(.SubIDs(1).ID).ColWidth(tmpItem.????) = tmpItem.W
                        msh(.ID).LeftCol = 0: msh(.SubIDs(1).ID).LeftCol = 0
                    Next
                End With
            ElseIf objItem.???? = 5 And objItem.???? = 0 Then
                For i = 0 To UBound(Split(objItem.????, "|"))
                    Set objCurItem = mobjReport.Items("_" & Split(Split(objItem.????, "|")(i), ",")(0))
                    With objCurItem
                        strWidth = ""
                        For Each tmpID In .SubIDs
                            Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                            Select Case tmpItem.????
                                Case 7
                                    If i = 0 Then msh(objItem.ID).ColWidth(tmpItem.????) = tmpItem.W
                                Case 9
                                    strWidth = strWidth & "," & tmpItem.W
                            End Select
                        Next
                        strWidth = Mid(strWidth, 2)
                        
                        If i = 0 Then
                            lngColB = msh(objItem.ID).FixedCols
                        Else
                            lngColB = lngColE + 1
                        End If
                        lngColE = CLng(Split(Split(objItem.????, "|")(i), ",")(1)) - 1
                        
                        For j = lngColB To lngColE
                            msh(objItem.ID).ColWidth(j) = _
                                CLng(Split(strWidth, ",")((j - lngColB) Mod (UBound(Split(strWidth, ",")) + 1)))
                        Next
                    End With
                Next
            End If
        End If
    Next
    
    '????????????????????
    Call SetGridAlign
    
    LockWindowUpdate 0
End Sub

Private Sub mnuEdit_SetCol_Fill_Click()
'??????????????????(??????????????????????????,????????????????????????)
    Dim tmpMsh As VSFlexGrid, i As Integer
    Dim lngCurW As Long, sngScale As Single
    
    If Not mobjReport.blnLoad Then Exit Sub
    
    On Error Resume Next
    
    LockWindowUpdate Me.hwnd
    
    timHead.Enabled = False
    '????
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") Then
            tmpMsh.Redraw = False
            
            lngCurW = GetGridColWidth(tmpMsh)
            If lngCurW < tmpMsh.Width - 300 Then
                sngScale = (tmpMsh.Width - 300) / lngCurW
                For i = 0 To tmpMsh.Cols - 1
                    tmpMsh.ColWidth(i) = tmpMsh.ColWidth(i) * sngScale
                Next
                tmpMsh.ColWidth(tmpMsh.Cols - 1) = _
                    tmpMsh.ColWidth(tmpMsh.Cols - 1) + tmpMsh.Width - 300 - GetGridColWidth(tmpMsh)
            End If
            
            tmpMsh.Redraw = True
        End If
    Next
    timHead.Enabled = True
    
    LockWindowUpdate 0
End Sub

Private Sub mnuFile_Graph_Click()
    Dim objHead As Object
    Dim objItem As RPTItem
    Dim bytKind As Byte
    Dim tmpMsh As Object
    
    If Not mobjReport.blnLoad Then Exit Sub
    
    If zlRegInfo("????????") <> "1" Then
        MsgBox "??????????????????????????????", vbInformation, App.Title
        Exit Sub
    End If
    
    If intGridCount = 0 Then
        MsgBox "??????????????????????????????????", vbInformation, App.Title
        Exit Sub
    End If
    If objCurGrid Is Nothing Then
        If msh.count > 1 Then
            For Each tmpMsh In msh
                If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") And Not tmpMsh.Tag Like "H_*" Then
                    Set objCurGrid = tmpMsh
                    Exit For
                End If
            Next
        End If
        If objCurGrid Is Nothing Then
            MsgBox "????????????????????????????????????", vbInformation, App.Title
            Exit Sub
        End If
    End If
    If objCurGrid.Tag Like "H_*" Then
        MsgBox "??????????????????????????", vbInformation, App.Title
        Exit Sub
    End If
    
    Set objItem = mobjReport.Items("_" & objCurGrid.Index)
    If objItem.???? = 4 Then
        bytKind = GetGridStyle(mobjReport, objItem.ID)
        If bytKind = 0 Then Set objHead = msh(CInt(objCurGrid.Tag))
    End If
    blnExcel = True
    Call ExcelChart(Me, objCurGrid, objHead, IIF(mobjReport.Items("_" & objCurGrid.Index).???? = 5, 1, 2), mobjReport.????, mobjReport.Fmts("_" & bytFormat).????)
End Sub

Private Sub mnuFile_Preview_Click()
    Dim frmShow As New frmPreview

    If Not mobjReport.blnLoad Then Exit Sub
    
    If mobjReport.Items.count = 0 Then Exit Sub
    
    If Not InitPrinter(Me) Then
        gblnError = True
        MsgBox "??????????????.????????????????????????????????????????????", vbInformation, App.Title: Exit Sub
    End If
    
    If Not CalcCellPage Then
        gblnError = True
        MsgBox "??????????????????,??????????????", vbInformation, App.Title: Exit Sub
    End If
    If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
        lbl(lngCurInx).BackColor = lngTmpColor
        lngCurInx = 0: lngTmpColor = 0
    End If
    
    timHead.Enabled = False
    SetRedraw False
    
    Set frmShow.frmParent = Me
    
    If mbytStyle = 1 Then
        If Not frmParent Is Nothing Then
            On Error Resume Next
            frmShow.Show 1, frmParent
            If Err.Number <> 0 Then
                Err.Clear
                frmShow.Show 1
            End If
        Else
            frmShow.Show 1
        End If
    Else
        frmShow.Show 1, Me
    End If
    
    timHead.Enabled = True
    SetRedraw True
End Sub

Private Sub mnuFile_Print_Click()
    Dim objItem As RPTItem, strSource As String
    Dim lngPrintH As Long, blnReset As Boolean
    Dim blnExit As Boolean, intCopy As Integer
    Dim blnDo As Boolean, blnCancel As Boolean
    Dim k As Integer, i As Integer, j As Integer
    Dim arrBill As Variant, strItem As String
    Dim objFmt As RPTFmt, blnGoOn As Boolean
    Dim blnPrint As Boolean, blnALLEmpty As Boolean
    Dim strTmp As String
    Dim strDefault As String
    Dim lngEndPage As Long
    
    If Not mobjReport.blnLoad Then Exit Sub
    If mobjReport.Items.count = 0 Then Exit Sub
    
    If Not mobjCurDLL Is Nothing Then
        mobjCurDLL.DataIsEmpty = False
    End If
    blnALLEmpty = True
    
    strDefault = mobjReport.Fmts(mobjReport.bytFormat).????
    strTmp = GetRegPrinterInfo("PaperCopy", mobjReport.????, strDefault)
    intCopy = Val(strTmp)
    If intCopy < 1 Then intCopy = 1
    If gblnSingleTask Then intCopy = 1 '????????????????????????????????
    If mobjReport.???? Then intCopy = 1 '??????????????????????1??
    
    blnGoOn = True
    Do While blnGoOn
        Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)

        '??????????,??????????????????????,????????
        blnExit = False
        'If (Not mblnPrintEmpty Or mobjReport.???????? = 1) And mbytStyle = 2 Then
        If mblnPrintEmpty = False And mobjReport.???????? = Val("1-??????????") And InStr(";0;2;4;", ";" & mbytStyle & ";") > 0 Then
            strSource = ""
            For Each objItem In mobjReport.Items
                If objItem.?????? = bytFormat Then
                    If objItem.???? = 4 Then '????????
                        strItem = GetGridSource(objItem, True) '"????????,????????,..."
                        If strItem <> "" Then strSource = strSource & "," & strItem
                    ElseIf objItem.???? = 5 Then '????????
                        strSource = strSource & "," & objItem.????
                    End If
                End If
            Next
            '????????????(??????)??????,????????????????????????????
            If strSource <> "" Then
                blnExit = True
                strSource = Mid(strSource, 2)
                For i = 0 To UBound(Split(strSource, ","))
                    On Error Resume Next
                    blnExit = blnExit And mLibDatas("_" & Split(strSource, ",")(i)).DataSet.RecordCount = 0
                    Err.Clear: On Error GoTo 0
                Next
            End If
            If blnExit Then GoTo NextFormat
        End If
        blnALLEmpty = False
        
        On Error GoTo errH
        
        '????????????????
        If Not InitPrinter(Me, intCopy) Then
            MsgBox "??????????????.????????????????????????????????????????????", vbInformation, App.Title
            gblnError = True: GoTo ExitHandle
        End If
        '??????PDF
        If mbytStyle = 4 Then
            If PDFInitialize() Then
                Call PDFFile(mstrPDFFile)
            Else
                Exit Sub
            End If
        End If
        
        k = intCopy '??????????????????k??
        If Printer.Copies = intCopy Then k = 1 '????????????????????
        
        '????????????
        If Not CalcCellPage Then
            gblnError = True
            MsgBox "??????????????????,??????????????", vbInformation, App.Title: GoTo ExitHandle
        End If
        If mbytStyle <> 2 And mbytStyle <> 4 Then
            If MsgBox("????????????????,????????????????????", vbQuestion + vbYesNo, App.Title) = vbNo Then GoTo ExitHandle
        End If
        
        If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
            lbl(lngCurInx).BackColor = lngTmpColor
            lngCurInx = 0: lngTmpColor = 0
        End If
        
        '????????????????????????????
        If Not mobjCurDLL Is Nothing Then
            arrBill = Empty: blnCancel = False: i = 1
            If IsArray(marrPage) Then i = UBound(marrPage) + 1
            Call mobjCurDLL.Act_BeforePrint(mobjReport.????, i * intCopy, blnCancel, arrBill)
            If blnCancel Then GoTo ExitHandle
            
            '????????????????????
            If IsArray(arrBill) Then garrBill = arrBill
        End If
        
        timHead.Enabled = False
        SetRedraw False
        
        '????????????
        If mbytStyle <> 2 Then Screen.MousePointer = 11
        
        j = 0
        blnReset = False
        Do
            k = k - 1
            j = j + 1
            If Not IsArray(marrPage) Then
                If IsArray(marrPageCard) Then
                    '????
                    GoTo makPage
                End If
                
                If mbytStyle <> 2 Then
                    If Printer.Copies <> intCopy And intCopy <> 1 Then
                        ShowFlash "????" & mobjReport.???? & ",?? 1 ?? " & intCopy & " ??,?????? " & j & " ??", j / intCopy, Me
                    Else
                        ShowFlash "????" & mobjReport.???? & "??", 1, Me
                    End If
                End If
                
                '??????????????????????
                If objFmt.???????? And objFmt.???? = 1 Then
                    Call PrintPage(0, Me, Me, 1, False, True, lngPrintH)
                    blnDo = lngPrintH > 0 And lngPrintH < objFmt.H
                    If blnDo Then '????????????30mm??????????????1/8
                        blnDo = objFmt.H - lngPrintH > 30 * Twip_mm And objFmt.H - lngPrintH > objFmt.H / 8
                    End If
                    If blnDo Then
                        lngPrintH = lngPrintH + 567 '??????????????10mm????
                        If Not SetPrinterPaper(Me.hwnd, mobjReport, lngPrintH, intCopy) Then
                            '????????????????????????
                            Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                        End If
                    End If
                End If
                blnPrint = True
                Call PrintPage(0, Printer, Me)
                
            Else
makPage:
                If IsArray(marrPage) Then
                    lngEndPage = UBound(marrPage)
                ElseIf IsArray(marrPageCard) Then
                    lngEndPage = UBound(marrPageCard)
                Else
                    lngEndPage = -1
                End If
                
                For i = 0 To lngEndPage
                    If mbytStyle <> 2 Then
                        If Printer.Copies <> intCopy And intCopy <> 1 Then
                            ShowFlash "????" & mobjReport.???? & ",?? " & lngEndPage + 1 & " ?? " & intCopy & " ??,?????? " & j & " ??", ((i + 1) + ((j - 1) * (lngEndPage + 1))) / ((lngEndPage + 1) * intCopy), Me
                        Else
                            ShowFlash "????" & mobjReport.???? & ",?? " & lngEndPage + 1 & " ??,?????? " & i + 1 & " ????", (i + 1) / (lngEndPage + 1), Me
                        End If
                    End If
                    
                    '??????????????????????
                    If objFmt.???????? And objFmt.???? = 1 Then
                        Call PrintPage(i, Me, Me, 1, False, True, lngPrintH)
                        blnDo = lngPrintH > 0 And lngPrintH < objFmt.H
                        If blnDo Then '????????????30mm??????????????1/8
                            blnDo = objFmt.H - lngPrintH > 30 * Twip_mm And objFmt.H - lngPrintH > objFmt.H / 8
                        End If
                        If blnDo Then
                            lngPrintH = lngPrintH + 567 '??????????????10mm????
                            If Not SetPrinterPaper(Me.hwnd, mobjReport, lngPrintH, intCopy) Then
                                '????????????????????????
                                Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                                blnReset = False
                            Else
                                blnReset = True '????????????????????,????????????????????????????????
                            End If
                        ElseIf blnReset Then
                            Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                            blnReset = False
                        End If
                    End If
                    
                    blnPrint = True
                    If Not PrintPage(i, Printer, Me) Then Exit For
                    If i <> lngEndPage Then Printer.NewPage: blnPrint = True '????
                Next
            End If
            If k > 0 Then Printer.NewPage: blnPrint = True '????
        Loop Until k = 0
        
NextFormat:
        '????????????????????????
        blnGoOn = False
        'If mbytStyle = 2 And mblnAllFormat And cboFormat.ComboItems.count > 1 And cboFormat.SelectedItem.Index < cboFormat.ComboItems.count Then
        If InStr(";0;2;4;", ";" & mbytStyle & ";") > 0 And mblnAllFormat And cboFormat.ComboItems.count > 1 And cboFormat.SelectedItem.Index < cboFormat.ComboItems.count Then
            cboFormat.ComboItems(cboFormat.SelectedItem.Index + 1).Selected = True
            Call CboFormat_Click: blnGoOn = True
            If Not (mblnPrintEmpty = False And mobjReport.???????? = 1) Or blnExit = False Then
                Printer.NewPage
            End If
            blnPrint = True '??????????????????????????,??????????????????
        End If
    Loop

    If Not mobjCurDLL Is Nothing Then
        mobjCurDLL.DataIsEmpty = blnALLEmpty
    End If

ExitHandle:
    If blnPrint Then
        If gblnSingleTask Then
            Printer.NewPage '??????????????????,??????????????????
        Else
            Printer.EndDoc
        End If
        
        '????????????????????????????
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_AfterPrint(mobjReport.????)
        End If
    End If

    If mbytStyle <> 2 Then ShowFlash
    timHead.Enabled = True
    SetRedraw True
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If mbytStyle <> 2 Then Call ShowFlash
    Printer.KillDoc
    timHead.Enabled = True
    SetRedraw True
    MsgBox Err.Number & ":" & Err.Description & vbCrLf & "????????????????????", vbExclamation, App.Title
    Err.Clear
    gblnError = True
End Sub

Private Sub mnuHelpTitle_Click()
    If Me.Tag = "" Then
        Call ShowHelpRpt(Me.hwnd, mobjReport.????, Int((mobjReport.????) / 100))
    Else
        Call ShowHelpRpt(Me.hwnd, Me.Tag, Int((mobjReport.????) / 100))
    End If
End Sub

Private Sub mnuPop_Cond_Click(Index As Integer)
    Set mobjPars = mdlPublic.RPTParsCondExec(mlngRPTID, Val(mnuPop_Cond(Index).Tag), mobjDefPars)
    If Not mobjPars Is Nothing Then
        mintCurMenuIndex = Index
        mintCurCondID = Val(mnuPop_Cond(Index).Tag)
        Call InitReportPars
        If cmdLoad.Enabled And cmdLoad.Visible Then cmdLoad.SetFocus
    End If
End Sub

Private Sub mnuPop_Default_Click()
    '??????????????????????????????????
    Call CopyPars(mobjDefPars, mobjPars)
    If Not mobjPars Is Nothing Then
        mintCurMenuIndex = 0
        mintCurCondID = 0
        Call InitReportPars
        If cmdLoad.Enabled And cmdLoad.Visible Then cmdLoad.SetFocus
    End If
End Sub

Private Sub mnuPop_Del_Click()
    If mdlPublic.RPTParsCondDel(mlngRPTID, mintCurCondID) Then
        Call mnuPop_Default_Click
    End If
End Sub

Private Sub mnuPop_Save_Click()
    '????????
    If mdlPublic.RPTParsCondSave(mlngRPTID, mintCurCondID, mobjPars, mobjDefPars, Me) Then
        '????????????
        If mintCurCondID = 0 Then
            '??????????????????????????????????
            Call mnuPop_Cond_Click(mnuPop_Cond.count - 1)
        Else
            '????????????????
            Call mnuPop_Cond_Click(mintCurCondID)
        End If
    End If
End Sub

Private Sub mnuPop_SaveAs_Click()
    If mdlPublic.RPTParsCondSave(mlngRPTID, mintCurCondID, mobjPars, mobjDefPars, Me, True) Then
        '????????????
        If mintCurCondID = 0 Then
            '??????????????????????????????????
            Call mnuPop_Cond_Click(mnuPop_Cond.count - 1)
        Else
            '????????????????
            Call mnuPop_Cond_Click(mintCurCondID)
        End If
    End If
End Sub

Private Sub mnuView_Next_Click()
    Dim intIdx As Integer
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    intIdx = lvw.SelectedItem.Index
    If intIdx + 1 <= lvw.ListItems.count Then
        lvw.ListItems(intIdx + 1).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuView_Pre_Click()
    Dim intIdx As Integer
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    intIdx = lvw.SelectedItem.Index
    If intIdx - 1 >= 1 Then
        lvw.ListItems(intIdx - 1).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuView_reFlash_Click()
    Dim strErr As String, strCond As String
    Dim tmpMsh As Object
    
    timHead.Enabled = False

    '??????????
    If lblName.UBound > 0 Then
        If Not ReSetReportPars Then Exit Sub
    End If
    
    '????????????????
    If Not mobjCurDLL Is Nothing Then
        strCond = GetParsStr(MakeNamePars(mobjReport, True))
        mobjCurDLL.Act_CommitCondition mobjReport.????, strCond, Me
    End If
    
     '????????????????????????????????
    If Format(mobjReport.????????????, "HH:mm:ss") <> "00:00:00" Or Format(mobjReport.????????????, "HH:mm:ss") <> "00:00:00" Then
        If CDate(Format(mobjReport.????????????, "HH:mm:ss")) > CDate(Format(mobjReport.????????????, "HH:mm:ss")) Then
            If Between(CDate(Format(Currentdate, "HH:mm:ss")), CDate(Format(mobjReport.????????????, "HH:mm:ss")), CDate(Format(mobjReport.????????????, "HH:mm:ss"))) Then
                MsgBox "??????????" & CDate(Format(mobjReport.????????????, "HH:mm:ss")) & "-" & CDate(Format(mobjReport.????????????, "HH:mm:ss")) & "????????????????????????????????", vbInformation, App.Title
                Exit Sub
            End If
        Else
            If CDate(Format(Currentdate, "HH:mm:ss")) < CDate(Format(mobjReport.????????????, "HH:mm:ss")) Or CDate(Format(Currentdate, "HH:mm:ss")) > CDate(Format(mobjReport.????????????, "HH:mm:ss")) Then
                MsgBox "??????????" & CDate(Format(mobjReport.????????????, "HH:mm:ss")) & "-??????" & CDate(Format(mobjReport.????????????, "HH:mm:ss")) & "????????????????????????????????", vbInformation, App.Title
                Exit Sub
            End If
        End If
    End If
    
    '????????
    strErr = OpenReportData(True)
    If strErr <> "" Then
        MsgBox "??????????????""" & strErr & """??????????????,??????????????", vbInformation, App.Title
        Exit Sub
    End If
    '????????
    Call ShowItems
    
    If lblName.UBound > 0 Then
        '????????????(??????????????????)
        picPar.Visible = False
        Set mobjPars = New RPTPars
        Set mobjPars = MakeNamePars(mobjReport)
        Call InitReportPars
        picPar.Visible = True
        
        '????????????,??????????????????????????
        Call KeepParsSame
    End If
    
    '??????????????????
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And tmpMsh.Container Is picPaper(intReport) And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
            On Error Resume Next
            tmpMsh.SetFocus: Exit For
        End If
    Next
    
    timHead.Enabled = True
End Sub

Private Sub KeepParsSame()
'??????????????????????????????????????????,????????????????????????????
'??????mobjPars=??????????????????????
'??????1.??????????????????,????????????????????????????
    Dim objPar As RPTPar, tmpPar As RPTPar
    Dim objData As RPTData, i As Integer
    For i = 0 To UBound(arrReport)
        If i <> intReport Then
            For Each objData In arrReport(i).Datas
                For Each objPar In objData.Pars
                    For Each tmpPar In mobjPars
                        If tmpPar.???? = objPar.???? _
                            And tmpPar.???? = objPar.???? _
                            And objPar.???? <> 3 Then
                            objPar.?????? = tmpPar.??????
                            objPar.Reserve = tmpPar.Reserve
                        End If
                    Next
                Next
            Next
        End If
    Next
End Sub

Private Sub mnuViewToolFormat_Click()
    mnuViewToolFormat.Checked = Not mnuViewToolFormat.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolGroup_Click()
    mnuViewToolGroup.Checked = Not mnuViewToolGroup.Checked
    picLR_S.Visible = Not picLR_S.Visible
    picGroup.Visible = Not picGroup.Visible
    Call Form_Resize
End Sub

Private Sub msh_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    If mobjReport.Items("_" & Index).???? = 6 Then
        msh(Mid(msh(Index).Tag, 3)).Col = Col
        msh(Mid(msh(Index).Tag, 3)).Sort = Order
    End If
End Sub

Private Sub msh_Click(Index As Integer)
    Dim objRelations As RPTRelations
    Dim i As Long
    Dim lngRec As Long, strDataName As String
    Dim strLisName As String
    Dim tmpData As RPTData, tmpPar As RPTPar
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim strFilter As String, lngItemID As Long
    Dim lngMouseRow As Long, lngMouseCol As Long
    
    lngMouseRow = msh(Index).MouseRow: lngMouseCol = msh(Index).MouseCol
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Sub
    If grsObject.State = adStateClosed Then
        Set grsObject = Nothing
        Set grsObject = UserObject
        If grsObject Is Nothing Then Exit Sub
    End If
    
    If lngMouseRow > -1 And lngMouseCol > -1 Then
        If msh(Index).Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
            If mobjReport.Items("_" & Index).???? = 4 Then
                Set objRelations = msh(Index).Cell(flexcpData, lngMouseRow, lngMouseCol)(2)
                '????????????????????????
                lngRec = msh(Index).Cell(flexcpData, lngMouseRow, lngMouseCol)(1)
            Else
                '??????????????????????????????
                lngRec = msh(Index).FixedRows
                If TypeName(msh(Index).Cell(flexcpData, lngRec, lngMouseCol)) = "Empty" Then Exit Sub
                
                Set objRelations = msh(Index).Cell(flexcpData, lngRec, lngMouseCol).Relations
                lngItemID = msh(Index).Cell(flexcpData, lngRec, lngMouseCol).ID
            End If
            If Not CheckReportPriv(objRelations.Item(1).????????ID) Then
                MsgBox "????????????????????????????????????????", vbInformation, App.Title: Exit Sub
            End If
            '????????
            If CheckPass(objRelations.Item(1).????????ID) = False Then
                MsgBox "??????????????????????????????", vbInformation, App.Title: Exit Sub
            End If
            
            Set gobjReport = ReadReport(objRelations.Item(1).????????ID)
            '??????????
            garrPars = Array()
            '??????????
            On Error Resume Next
            For i = 1 To objRelations.count
                If InStr(objRelations.Item(i).??????????, ".") > 0 Then
                    strDataName = Mid(objRelations.Item(i).??????????, 1, InStr(objRelations.Item(i).??????????, ".") - 1)
                End If
                If strDataName <> "" Then Exit For
            Next
            If mobjReport.Items("_" & Index).???? = 4 Then
                '??????????????????????????????
                If strDataName <> "" Then mLibDatas("_" & strDataName).DataSet.AbsolutePosition = lngRec
            Else
                '??????????????????????????????????????
                If strDataName <> "" Then
                    For Each tmpID In mobjReport.Items("_" & Index).SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                        Select Case mobjReport.Items("_" & lngItemID).????
                            Case 7 '????????
                                If tmpItem.???? = 7 Then
                                    If Decode(Trim(msh(Index).TextMatrix(lngMouseRow, tmpItem.????)), "????", 1, "??????", 2, "??????", 3, "??????", 4, "??????", 5, 0) > 0 Then
                                        '????????????????????????
                                        lngMouseRow = lngMouseRow - 1
                                    End If
                                    strFilter = strFilter & " And " & tmpItem.???? & "='" & msh(Index).TextMatrix(lngMouseRow, tmpItem.????) & "'"
                                End If
                            Case 8 '????????
                                If tmpItem.???? = 8 Then
                                    strFilter = strFilter & " And " & tmpItem.???? & "='" & msh(Index).TextMatrix(tmpItem.????, lngMouseCol) & "'"
                                End If
                            Case 9 '??????
                                '????????????????????????????????
                                If tmpItem.???? = 7 Then
                                    If Decode(Trim(msh(Index).TextMatrix(lngMouseRow, tmpItem.????)), "????", 1, "??????", 2, "??????", 3, "??????", 4, "??????", 5, 0) > 0 Then
                                        '????????????????????????
                                        lngMouseRow = lngMouseRow - 1
                                    End If
                                    strFilter = strFilter & " And " & tmpItem.???? & "='" & msh(Index).TextMatrix(lngMouseRow, tmpItem.????) & "'"
                                ElseIf tmpItem.???? = 8 Then
                                    strFilter = strFilter & " And " & tmpItem.???? & "='" & msh(Index).TextMatrix(tmpItem.????, lngMouseCol) & "'"
                                End If
                        End Select
                    Next
                    mLibDatas("_" & strDataName).DataSet.Filter = Mid(strFilter, 6)
                End If
            End If
            For i = 1 To objRelations.count
                With objRelations.Item(i)
                    strLisName = ""
                    If InStr(.??????????, ".") > 0 Then
                        If mLibDatas("_" & strDataName).DataSet.RecordCount > 0 Then
                            strLisName = mLibDatas("_" & strDataName).DataSet.Fields(Mid(.??????????, InStr(.??????????, ".") + 1)).Value
                        End If
                    ElseIf InStr(.??????????, "=") = 1 Then
                        For Each tmpData In mobjReport.Datas
                            For Each tmpPar In tmpData.Pars
                                If tmpPar.???? = Mid(.??????????, 2) Then
                                    strLisName = tmpPar.??????
                                    Exit For
                                End If
                            Next
                            If strLisName <> "" Then Exit For
                        Next
                    End If
                    ReDim Preserve garrPars(UBound(garrPars) + 1)
                    garrPars(UBound(garrPars)) = .?????? & "=" & strLisName
                End With
            Next
            
            
            If Not ShowReport(Me) Then MsgBox "??????????????", vbInformation, App.Title
        End If
    End If
End Sub

Private Sub msh_DblClick(Index As Integer)
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetDblClick(mobjReport.????, msh(Index), Me)
        msh(Index).SetFocus
    End If
End Sub

Private Sub msh_EnterCell(Index As Integer)
    Dim i As Integer, j As Integer, strRowText As String
    Dim intRow As Integer, intCol As Integer, strText As String
    Static strRow As String
    Static strCol As String
    
    If blnRefresh = False Then Exit Sub
    Set objCurGrid = msh(Index)
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        intRow = msh(Index).Row
        intCol = msh(Index).Col
        strText = msh(Index).Text
        Call mobjCurDLL.Act_EnterCell(mobjReport.????, intRow, intCol, strText)
        '??????????
        If intRow >= 0 And intRow <= msh(Index).Rows - 1 And intCol >= 0 And intCol <= msh(Index).Cols - 1 Then
            msh(Index).Row = intRow
            msh(Index).Col = intCol
            msh(Index).Text = strText
        End If
        
        If strRow <> Index & "," & msh(Index).Row Then
            For i = 0 To msh(Index).Cols - 1
                strRowText = strRowText & "|" & msh(Index).TextMatrix(msh(Index).Row, i)
            Next
            Call mobjCurDLL.Act_EnterRow(mobjReport.????, msh(Index).Row, Mid(strRowText, 2), msh(Index))
            strRow = Index & "," & msh(Index).Row
        End If
        
        If strCol <> Index & "," & msh(Index).Col Then
            Call mobjCurDLL.Act_EnterCol(mobjReport.????, msh(Index).Col, msh(Index))
            strCol = Index & "," & msh(Index).Col
        End If
    End If
    
'    '????????????
'    If mnuEdit_SelMode_Row.Checked Then
'        If msh(Index).Row >= msh(Index).FixedRows And msh(Index).Col >= msh(Index).FixedCols Then
'            msh(Index).Redraw = False
'            For i = msh(Index).FixedCols To msh(Index).Cols - 1
'                If msh(Index).ColData(i) <> 0 Or msh(Index).RowData(msh(Index).Row) <> 0 Then
'                    msh(Index).Cell(flexcpBackColor, msh(Index).Row, i) = &H808080 '????????????????
'                    msh(Index).Cell(flexcpForeColor, msh(Index).Row, i) = msh(Index).ForeColorSel
'                Else
'                    msh(Index).Cell(flexcpBackColor, msh(Index).Row, i) = msh(Index).BackColorSel
'                    msh(Index).Cell(flexcpForeColor, msh(Index).Row, i) = msh(Index).ForeColorSel
'                End If
'                If msh(Index).Cell(flexcpFontUnderline, msh(Index).Row, i) = True Then
'                    msh(Index).Cell(flexcpForeColor, msh(Index).Row, i) = &H00FF0001&
'                End If
'            Next
'            msh(Index).Redraw = True
'        End If
'    ElseIf mnuEdit_SelMode_Col.Checked Then
'        If msh(Index).Row >= msh(Index).FixedRows And msh(Index).Col >= msh(Index).FixedCols Then
'            msh(Index).Redraw = False
'            For i = msh(Index).FixedRows To msh(Index).Rows - 1
'                If msh(Index).ColData(msh(Index).Col) <> 0 Or msh(Index).RowData(i) <> 0 Then
'                    msh(Index).Cell(flexcpBackColor, i, msh(Index).Col) = &H808080 '????????????????
'                    msh(Index).Cell(flexcpForeColor, i, msh(Index).Col) = msh(Index).ForeColorSel
'                Else
'                    msh(Index).Cell(flexcpBackColor, i, msh(Index).Col) = msh(Index).BackColorSel
'                    msh(Index).Cell(flexcpForeColor, i, msh(Index).Col) = msh(Index).ForeColorSel
'                End If
'                If msh(Index).Cell(flexcpFontUnderline, i, msh(Index).Col) = True Then
'                    msh(Index).Cell(flexcpForeColor, i, msh(Index).Col) = &H00FF0001&
'                End If
'            Next
'            msh(Index).Redraw = True
'        End If
'    End If
End Sub

Private Sub msh_GotFocus(Index As Integer)
    On Error Resume Next
    If msh(Index).Tag Like "H_*" Then
        msh(CInt(Mid(msh(Index).Tag, 3))).SetFocus
    Else
        Call msh_EnterCell(Index)
    End If
End Sub

Private Sub msh_LeaveCell(Index As Integer)
    Dim i As Integer, j As Integer
    Dim intPre As Integer
    
    If blnRefresh = False Then Exit Sub
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_LevelCell(mobjReport.????, msh(Index).Row, msh(Index).Col, msh(Index).Text)
    End If
    
'    If mnuEdit_SelMode_Row.Checked Then
'        If msh(Index).Row >= msh(Index).FixedRows And msh(Index).Col >= msh(Index).FixedCols Then
'            msh(Index).Redraw = False
'            On Error Resume Next
'            msh(Index).Cell(flexcpBackColor, msh(Index).Row, msh(Index).FixedCols, msh(Index).Row, msh(Index).Cols - 1) = msh(Index).BackColor
'            msh(Index).Cell(flexcpForeColor, msh(Index).Row, msh(Index).FixedCols, msh(Index).Row, msh(Index).Cols - 1) = msh(Index).ForeColor
'            For i = msh(Index).FixedCols To msh(Index).Cols - 1
'                If msh(Index).Cell(flexcpFontUnderline, msh(Index).Row, i) = True Then
'                    msh(Index).Cell(flexcpForeColor, msh(Index).Row, i) = &H00FF0001&
'                End If
'            Next
'            msh(Index).Redraw = True
'        End If
'    ElseIf mnuEdit_SelMode_Col.Checked Then
'        If msh(Index).Row >= msh(Index).FixedRows And msh(Index).Col >= msh(Index).FixedCols Then
'            msh(Index).Redraw = False
'            On Error Resume Next
'            msh(Index).Cell(flexcpBackColor, msh(Index).FixedRows, msh(Index).Col, msh(Index).Rows - 1, msh(Index).Col) = msh(Index).BackColor
'            msh(Index).Cell(flexcpForeColor, msh(Index).FixedRows, msh(Index).Col, msh(Index).Rows - 1, msh(Index).Col) = msh(Index).ForeColor
'            For i = msh(Index).FixedRows To msh(Index).Rows - 1
'                If msh(Index).Cell(flexcpFontUnderline, i, msh(Index).Col) = True Then
'                    msh(Index).Cell(flexcpForeColor, i, msh(Index).Col) = &H00FF0001&
'                End If
'            Next
'            msh(Index).Redraw = True
'        End If
'    End If
End Sub

Private Sub msh_LostFocus(Index As Integer)
    If Not msh(Index).Tag Like "H_*" Then Call msh_LeaveCell(Index)
End Sub

Private Sub msh_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseDown(mobjReport.????, Button, Shift, X, Y, msh(Index), Me)
    End If
End Sub

Private Sub msh_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    msh(Index).ToolTipText = msh(Index).TextMatrix(msh(Index).MouseRow, msh(Index).MouseCol)
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseMove(mobjReport.????, Button, Shift, X, Y, msh(Index), Me)
    End If
    If msh(Index).MouseRow > -1 And msh(Index).MouseCol > -1 Then
        If msh(Index).Cell(flexcpFontUnderline, msh(Index).MouseRow, msh(Index).MouseCol) = True Then
            msh(Index).MousePointer = 99
        Else
            msh(Index).MousePointer = 0
        End If
    Else
        msh(Index).MousePointer = 0
    End If
End Sub

Private Sub msh_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseUp(mobjReport.????, Button, Shift, X, Y, msh(Index), Me)
    End If
End Sub

Private Sub msh_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim intPre As Integer
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetScroll(mobjReport.????, msh(Index))
    End If
    
    If IsNumeric(msh(Index).Tag) Then
        intPre = msh(msh(Index).Tag).LeftCol
        msh(msh(Index).Tag).LeftCol = msh(Index).LeftCol
        If msh(msh(Index).Tag).LeftCol = intPre Then msh(Index).LeftCol = intPre
    ElseIf Left(msh(Index).Tag, 2) = "H_" Then
        intPre = msh(Mid(msh(Index).Tag, 3)).LeftCol
        msh(Mid(msh(Index).Tag, 3)).LeftCol = msh(Index).LeftCol
        If msh(Mid(msh(Index).Tag, 3)).LeftCol = intPre Then msh(Index).LeftCol = intPre
    End If
End Sub

Private Sub opt_GotFocus(Index As Integer)
    If opt(Index).Value Then
        '????????????????????TAB????????????????????????
        opt(Index).Value = False
        opt(Index).Value = True
    End If
End Sub

Private Sub picLR_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objTmp As Object
    
    On Error Resume Next
    
    If Button = 1 Then
        If picGroup.Width + X < 1000 Or picBack.Width - X < 3000 Then Exit Sub
        picLR_S.Left = picLR_S.Left + X

        picGroup.Width = picGroup.Width + X
        picBack.Left = picBack.Left + X
        picBack.Width = picBack.Width - X
        scrHsc.Left = scrHsc.Left + X
        scrHsc.Width = scrHsc.Width - X
        
        lblGroup_S.Width = lblGroup_S.Width + X
        lvw.Width = lvw.Width + X
        lblPar_S.Width = lblPar_S.Width + X
        picPar.Width = picPar.Width + X
        
        lvw.ColumnHeaders(1).Width = lvw.Width - 500    '????????
        
        For Each objTmp In fraGroup
            objTmp.Width = picGroup.ScaleWidth - objTmp.Left * 2
        Next
        For Each objTmp In fra
            objTmp.Width = picGroup.ScaleWidth - objTmp.Left * 2
        Next
        
        picPaper(intReport).Cls
        Call SetPaper
        Call SetPlace
        Me.Refresh
    End If
End Sub

Private Sub picPaper_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnPop As Boolean
    
    lngPreX = X: lngPreY = Y
    
    If Not mobjCurDLL Is Nothing Then
        blnPop = True
        Call mobjCurDLL.Act_PaperMouseDown(mobjReport.????, Button, Shift, X, Y, blnPop)
        If blnPop Then
            If Button = 2 Then PopupMenu mnuEdit, 2
        End If
    Else
        If Button = 2 Then PopupMenu mnuEdit, 2
    End If
End Sub

Private Sub picPaper_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_PaperMouseMove(mobjReport.????, Button, Shift, X, Y)
    End If
    If Button = 1 Then
        If scrVsc.Enabled And scrVsc.Visible Then
            If (Y - lngPreY) / 15 > 0 Then
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - lngPreY) / 15)
            Else
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - lngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled And scrHsc.Visible Then
            If (X - lngPreX) / 15 > 0 Then
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - lngPreX) / 15)
            Else
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - lngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub mnuEdit_Par_Click()
'??????????????????
    Dim strErr As String, objPars As RPTPars
    Dim strCond As String, blnInhere As Boolean
    Dim lngReport As Long, strSQL As String
    Dim rsReport As New ADODB.Recordset
    Dim frmNewParInput As New frmParInput
    
    '??????????ID
    lngReport = 0
    strSQL = "Select ID from zlReports Where ????=[1]"
    Set rsReport = OpenSQLRecord(strSQL, Me.Caption, mobjReport.????)
    If Not rsReport.EOF Then lngReport = rsReport!ID
    
    If Not mobjCurDLL Is Nothing Then
        blnInhere = True
        Set objPars = MakeNamePars(mobjReport, True)
        strCond = GetParsStr(objPars)
        
        '????????????????
        mobjCurDLL.Act_ResetCondition mobjReport.????, strCond, blnInhere, Me
        
        If Not blnInhere Then
             '????????????????????,??????????????????
            If strCond = "" Or Not strCond Like "*=*" Then Exit Sub
            
            Set objPars = SetStrPars(strCond, objPars)
            timHead.Enabled = False
            
            '????????????????
            strCond = GetParsStr(objPars)
            mobjCurDLL.Act_CommitCondition mobjReport.????, strCond, Me
            
            ReplaceInputPars objPars
            
            Me.Refresh
            strErr = OpenReportData(True)
            If strErr <> "" Then MsgBox "??????????????""" & strErr & """??????????????,??????????????", vbInformation, App.Title: Exit Sub
            Call ShowItems
            timHead.Enabled = True
        Else
            timHead.Enabled = False
            
            Set objPars = MakeNamePars(mobjReport) '????????????????????
            
            frmNewParInput.mlngReport = lngReport
            Set frmNewParInput.mobjPars = objPars
            Set frmNewParInput.mobjDefPars = mobjDefPars
            Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
            
            frmNewParInput.mstrTitle = mobjReport.????
            frmNewParInput.mblnReset = True
            frmNewParInput.Show 1, Me
            If frmNewParInput.mblnOK Then
                '????????????????
                strCond = GetParsStr(frmNewParInput.mobjPars)
                mobjCurDLL.Act_CommitCondition mobjReport.????, strCond, Me
                
                ReplaceInputPars frmNewParInput.mobjPars
                Unload frmNewParInput
                
                '????????
                Me.Refresh
                strErr = OpenReportData(True)
                If strErr <> "" Then MsgBox "??????????????""" & strErr & """??????????????,??????????????", vbInformation, App.Title: Exit Sub
                Call ShowItems
            End If
            timHead.Enabled = True
        End If
    Else
        timHead.Enabled = False
        
        frmNewParInput.mlngReport = lngReport
        Set frmNewParInput.mobjPars = MakeNamePars(mobjReport)
        Set frmNewParInput.mobjDefPars = mobjDefPars
        Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
        
        frmNewParInput.mstrTitle = mobjReport.????
        frmNewParInput.mblnReset = True
        frmNewParInput.Show 1, Me
        If frmNewParInput.mblnOK Then
            ReplaceInputPars frmNewParInput.mobjPars
            Unload frmNewParInput
           
            '????????
            Me.Refresh
            strErr = OpenReportData(True)
            If strErr <> "" Then MsgBox "??????????????""" & strErr & """??????????????,??????????????", vbInformation, App.Title: Exit Sub
            Call ShowItems
        End If
        timHead.Enabled = True
    End If
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '??????????????
    Dim staH As Long '??????????????
    Dim lngTmp As Long
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    '??????????????????
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(sta.Visible, sta.Height, 0)
    
    lblGroup_S.Width = picGroup.ScaleWidth - lblGroup_S.Left * 2
    lblPar_S.Width = lblGroup_S.Width
    
    lvw.Top = lblGroup_S.Top + lblGroup_S.Height + 15
    lvw.Width = picGroup.ScaleWidth
    lvw.Height = lblPar_S.Top - lblGroup_S.Top - lblGroup_S.Height - 15 * 2
    
    picPar.Top = lblPar_S.Top + lblPar_S.Height + 15
    picPar.Left = 0
    picPar.Width = lvw.Width
    picPar.Height = ScaleHeight - staH - cbrH - (lblGroup_S.Height + 30) - (lblPar_S.Height + 30) - lvw.Height
    
    picBack.Top = ScaleTop + cbrH
    picBack.Left = ScaleLeft + IIF(picGroup.Visible, picGroup.Width + picLR_S.Width, 0)
    picBack.Width = ScaleWidth - IIF(scrVsc.Visible, scrVsc.Width, 0) - IIF(picGroup.Visible, picGroup.Width + picLR_S.Width, 0)
    picBack.Height = ScaleHeight - staH - cbrH - IIF(scrHsc.Visible, scrHsc.Height, 0)
    
    If scrVsc.Visible Then
        scrVsc.Top = picBack.Top
        scrVsc.Left = ScaleWidth - scrVsc.Width
        scrVsc.Height = picBack.Height
        
        scrHsc.Left = picBack.Left
        scrHsc.Top = picBack.Top + picBack.Height
        scrHsc.Width = picBack.Width
    End If
    
    On Error GoTo 0
    
    If Not mobjReport Is Nothing And Visible Then
        picPaper(intReport).Cls
        Call SetPaper
        Call SetPlace
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, bytMode As Byte
    
    timHead.Enabled = False
    bytMode = IIF(mnuEdit_SelMode_Row.Checked, 0, 1)
    
    If lvw.ListItems.count > 0 Then
        SaveWinState Me, App.ProductName, Me.Tag
        SaveSetting "ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & Me.Tag, "????????", bytMode
        For i = 0 To UBound(arrReport)
            SaveSetting "ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & arrReport(i).????, "????", arrReport(i).bytFormat
        Next
    ElseIf Not mobjReport Is Nothing Then
        If mbytStyle = 0 Then '??????????????????????????????
            SaveWinState Me, App.ProductName, mobjReport.????
        End If
        SaveSetting "ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.????, "????????", bytMode
    End If
    
    If Not mobjReport Is Nothing Then
        SaveSetting "ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.????, "????", bytFormat
    End If
    
    '????????????????
    If Not mobjCurDLL Is Nothing And Not mobjReport Is Nothing Then
        Call mobjCurDLL.Act_ReportUnload(mobjReport.????, Me)
    End If
    
    '????????????
    '---------------------------------------------------
    mbytStyle = 0
    mstrExcelFile = ""
    mstrPDFFile = ""
    
    Unload frmFlash
    
    Set frmParent = Nothing
    Set mobjCurDLL = Nothing
    Set mobjReport = Nothing
    Set mLibDatas = Nothing
    Set objCurGrid = Nothing
    Set mobjPars = Nothing
    Set mobjDefPars = Nothing
    Set objScript = Nothing
    
    Erase arrReport, arrLibDatas, arrDefPars

    If IsArray(marrPars) Then Erase marrPars
    If IsArray(marrPage) Then Erase marrPage
    marrPars = Empty
    marrPage = Empty

    Err.Clear
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuFile_Setup_Click()
    Dim objFmt As RPTFmt
    Dim strTmp As String
    Dim strDefault As String
    
    Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)
    strDefault = mobjReport.Fmts(mobjReport.bytFormat).????
    strTmp = GetRegPrinterInfo("Printer", mobjReport.????, objFmt.????, mobjReport)
    If Not ReportLocalSet(mobjReport.????, mobjReport.????, False, mobjReport.bytFormat, Me) Then Exit Sub
    sta.Panels(2) = "??????:" & strTmp & _
        "   ????:" & GetPaperName(objFmt.????, objFmt.W, objFmt.H) & " " & _
        IIF(objFmt.???? = 256, CInt(objFmt.W / Twip_mm) & "mm ?? " & CInt(objFmt.H / Twip_mm) & "mm", "") & _
        IIF(objFmt.???? = 1, "   ????", "   ????")
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta.Visible = Not sta.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.count
        tbr.Buttons(i).Caption = IIF(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub picPaper_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_PaperMouseUp(mobjReport.????, Button, Shift, X, Y)
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Par"
            If lvw.ListItems.count = 0 Then
                mnuEdit_Par_Click '??????????????????
            Else
                mnuView_reFlash_Click
            End If
        Case "Preview"
            mnuFile_Preview_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Graph"
            mnuFile_Graph_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Style"
            Call SetView((lvw.View + 1) Mod 4)
        Case "Pre"
            mnuView_Pre_Click
        Case "Next"
            mnuView_Next_Click
        Case "ColWidth"
            mnuEdit_SetCol_Auto_Click
        Case "SelMode"
            If mnuEdit_SelMode_Row.Checked Then
                mnuEdit_SelMode_Col_Click
            Else
                mnuEdit_SelMode_Row_Click
            End If
    End Select
End Sub

Private Sub SetPaper()
'??????????????????????,????,????
'??????????????????????
    Dim strPrinter As String
    Dim strDefault As String
    
    strDefault = mobjReport.Fmts(mobjReport.bytFormat).????
    strPrinter = GetRegPrinterInfo("Printer", mobjReport.????, strDefault, mobjReport)
    With mobjReport.Fmts("_" & mobjReport.bytFormat)
        sta.Panels(2).Text = "??????:" & strPrinter & "   ????:" & GetPaperName(.????, .W, .H) & " " & _
            IIF(.???? = 256, CInt(.W / Twip_mm) & "mm ?? " & CInt(.H / Twip_mm) & "mm", "") & _
            IIF(.???? = 1, "   ????", "   ????")
    End With
    On Error GoTo errH
    
    If intGridCount = 1 And Not mobjReport.???? Then
        picPaper(intReport).Top = 45
        picPaper(intReport).Left = 45
        picPaper(intReport).Width = picBack.ScaleWidth - picPaper(intReport).Left * 2
        picPaper(intReport).Height = picBack.ScaleHeight - picPaper(intReport).Top * 2
    Else
        With mobjReport.Fmts("_" & mobjReport.bytFormat)
            If .???? = 1 Then
                picPaper(intReport).Width = .W
                picPaper(intReport).Height = .H
            Else
                picPaper(intReport).Width = .H
                picPaper(intReport).Height = .W
            End If
        End With
        picShadow.Width = picPaper(intReport).Width
        picShadow.Height = picPaper(intReport).Height
        
        If picBack.ScaleWidth >= picPaper(intReport).Width + 180 Then
            picPaper(intReport).Left = (picBack.ScaleWidth - (picPaper(intReport).Width + 180)) / 2 + 60
            scrHsc.Enabled = False
        Else
            picPaper(intReport).Left = 60
            scrHsc.Max = (picPaper(intReport).Width + 180 - picBack.ScaleWidth) / 15
            If scrHsc.Max / 3 < scrHsc.SmallChange Then
                scrHsc.LargeChange = scrHsc.SmallChange
            Else
                scrHsc.LargeChange = scrHsc.Max / 3
            End If
            scrHsc.Enabled = True
        End If
        
        If picBack.ScaleHeight >= picPaper(intReport).Height + 180 Then
            picPaper(intReport).Top = (picBack.ScaleHeight - (picPaper(intReport).Height + 180)) / 2 + 60
            scrVsc.Enabled = False
        Else
            picPaper(intReport).Top = 60
            scrVsc.Max = (picPaper(intReport).Height + 180 - picBack.ScaleHeight) / 15
            If scrVsc.Max / 3 < scrVsc.SmallChange Then
                scrVsc.LargeChange = scrVsc.SmallChange
            Else
                scrVsc.LargeChange = scrVsc.Max / 3
            End If
            scrVsc.Enabled = True
        End If
        
        picShadow.Top = picPaper(intReport).Top + 60
        picShadow.Left = picPaper(intReport).Left + 60
    End If
    Exit Sub
errH:
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub scrhsc_Change()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrHsc.Value / (scrHsc.Max - scrHsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.????, 0, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrHsc.Value = (scrHsc.Max - scrHsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Left = -scrHsc.Value * 15# + 60
    picShadow.Left = picPaper(intReport).Left + 60
    Me.Refresh
End Sub

Private Sub scrhsc_Scroll()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrHsc.Value / (scrHsc.Max - scrHsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.????, 0, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrHsc.Value = (scrHsc.Max - scrHsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Left = -scrHsc.Value * 15# + 60
    picShadow.Left = picPaper(intReport).Left + 60
    Me.Refresh
End Sub

Private Sub scrVsc_Change()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrVsc.Value / (scrVsc.Max - scrVsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.????, 1, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrVsc.Value = (scrVsc.Max - scrVsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Top = -scrVsc.Value * 15# + 60
    picShadow.Top = picPaper(intReport).Top + 60
    Me.Refresh
End Sub

Private Sub scrVsc_Scroll()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrVsc.Value / (scrVsc.Max - scrVsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.????, 1, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrVsc.Value = (scrVsc.Max - scrVsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Top = -scrVsc.Value * 15# + 60
    picShadow.Top = picPaper(intReport).Top + 60
    Me.Refresh
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Auto"
            mnuEdit_SetCol_Auto_Click
        Case "Def"
            mnuEdit_SetCol_Def_Click
        Case "Fill"
            mnuEdit_SetCol_Fill_Click
        Case "Large"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
        Case "RowMode"
            mnuEdit_SelMode_Row_Click
        Case "ColMode"
            mnuEdit_SelMode_Col_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub timHead_Timer()
'??????????????????????????????????
    Dim tmpMsh As Object, sngWidth As Single
    Dim lngRow As Long, lngCol As Long
    Dim lngTopRow As Long, lngLeftCol As Long
    Dim i As Integer, j As Integer
    
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") _
            And tmpMsh.FixedRows = 0 And IsNumeric(tmpMsh.Tag) Then
            
'            tmpMsh.Redraw = False
'            lngRow = tmpMsh.Row
'            lngCol = tmpMsh.Col
'            lngTopRow = tmpMsh.TopRow
'            lngLeftCol = tmpMsh.LeftCol
            For i = 0 To tmpMsh.Cols - 1
                If tmpMsh.ColWidth(i) <> msh(tmpMsh.Tag).ColWidth(i) Then
                    sngWidth = msh(tmpMsh.Tag).ColWidth(i)
                    If Not mobjCurDLL Is Nothing Then
                        Call mobjCurDLL.Act_ColResize(mobjReport.????, i, sngWidth, tmpMsh.ColWidth(i))
                    End If

                    tmpMsh.ColWidth(i) = sngWidth
                    msh(tmpMsh.Tag).ColWidth(i) = sngWidth
                    
'                    tmpMsh.Col = i
'                    For j = 0 To tmpMsh.Rows - 1
'                        tmpMsh.Row = j
'                        If Not tmpMsh.CellPicture Is Nothing Then
'                            Me.picTemp.Cls '??????????
'                            Me.picTemp.Width = tmpMsh.CellWidth
'                            Me.picTemp.Height = tmpMsh.CellHeight
'                            Me.picTemp.PaintPicture tmpMsh.CellPicture, 0, 0, tmpMsh.CellWidth, tmpMsh.CellHeight
'
'                            Set tmpMsh.CellPicture = Me.picTemp.Image
'                            tmpMsh.CellPictureAlignment = 4
'                        End If
'                    Next
                End If
            Next
'            tmpMsh.Row = lngRow
'            tmpMsh.Col = lngCol
'            tmpMsh.TopRow = lngTopRow
'            tmpMsh.LeftCol = lngLeftCol
'            tmpMsh.Redraw = True
        End If
    Next
End Sub

Private Function GetGridSource(objItem As RPTItem, Optional ByVal blnHead As Boolean) As String
'????????????????????????????????????????
'??????objItem=????????????
'      blnHead=????????????????????
'??????"????????,????????,...",""
    Dim tmpID As RelatID
    Dim strSource As String, strFormula As String
    
    For Each tmpID In objItem.SubIDs
        strFormula = mobjReport.Items("_" & tmpID.ID).????
        Do While InStr(strFormula, "[") > 0
            strSource = Trim(Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1))
            strFormula = Mid(strFormula, InStr(strFormula, "]") + 1)
            If InStr(strSource, ".") > 0 Then
                If InStr(GetGridSource & ",", "," & Left(strSource, InStr(strSource, ".") - 1) & ",") = 0 Then
                    GetGridSource = GetGridSource & "," & Left(strSource, InStr(strSource, ".") - 1)
                End If
            End If
        Loop
        
        If blnHead Then
            strFormula = mobjReport.Items("_" & tmpID.ID).????
            Do While InStr(strFormula, "[") > 0
                strSource = Trim(Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1))
                strFormula = Mid(strFormula, InStr(strFormula, "]") + 1)
                If InStr(strSource, ".") > 0 Then
                    If InStr(GetGridSource & ",", "," & Left(strSource, InStr(strSource, ".") - 1) & ",") = 0 Then
                        GetGridSource = GetGridSource & "," & Left(strSource, InStr(strSource, ".") - 1)
                    End If
                End If
            Loop
        End If
    Next
    If GetGridSource <> "" Then GetGridSource = Mid(GetGridSource, 2)
End Function

Private Function ReplaceUserPars(objReport As Report) As Boolean
'??????????????????????????????????????
'????????????????????(????)(????)????
'??????????????"="??????????????????,????????????????????????Split????,????Instr????
    Dim tmpData As RPTData, tmpPar As RPTPar
    Dim i As Integer, j As Integer, k As Integer
    Dim blnCur As Boolean, blnALL As Boolean
    Dim strTmp As String
    
    If Not IsArray(marrPars) Then Exit Function
    If UBound(marrPars) <> -1 Then
        '??????????????????
        blnALL = True
        For Each tmpData In objReport.Datas
            For Each tmpPar In tmpData.Pars
                blnCur = False: k = k + 1
                For i = 0 To UBound(marrPars)
                    '????????????????????????????
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase(tmpPar.????) Then
                            strTmp = Trim(Mid(CStr(marrPars(i)), j + 1))
                            If strTmp <> "" Then
                                If InStr(strTmp, "|") > 0 And (tmpPar.?????? = "????????????" Or tmpPar.?????? = "????????????") Then
                                    blnCur = True: Exit For
                                Else
                                    Select Case tmpPar.????
                                        Case 0, 3
                                            blnCur = True: Exit For
                                        Case 1
                                            If IsNumeric(strTmp) Then blnCur = True: Exit For
                                        Case 2
                                            If IsDate(strTmp) Then blnCur = True: Exit For
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next
                blnALL = blnALL And blnCur
            Next
        Next
        
        '??????
        For Each tmpData In objReport.Datas
            For Each tmpPar In tmpData.Pars
                k = k + 1
                For i = 0 To UBound(marrPars)
                    '????????????????????????????
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase(tmpPar.????) Then
                            strTmp = Trim(Mid(CStr(marrPars(i)), j + 1))
                            If strTmp <> "" Then
                                '????????????????????????,??????????????????
                                If InStr(strTmp, "|") = 0 And (tmpPar.?????? = "????????????" Or tmpPar.?????? = "????????????") Then
                                    If tmpPar.?????? = "????????????" Then
                                        For j = 0 To UBound(Split(tmpPar.??????, "|"))
                                            If Split(Split(tmpPar.??????, "|")(j), ",")(1) = strTmp Then
                                                strTmp = Split(Split(tmpPar.??????, "|")(j), ",")(0) & "|" & strTmp
                                                If Left(strTmp, 1) = "??" Then strTmp = Mid(strTmp, 2)
                                                Exit For
                                            End If
                                        Next
                                    Else
                                        '????????????????????????,????????????????????????
                                        strTmp = "????????|" & strTmp
                                    End If
                                End If
                                If InStr(strTmp, "|") > 0 And (tmpPar.?????? = "????????????" Or tmpPar.?????? = "????????????") Then
                                    '????????,??????????????????
                                    If Not blnALL Then
                                        '????????????????????,????????????????????????????????
                                        tmpPar.Reserve = strTmp
                                    Else
                                        '??????????????,????????????????????????????????????
                                        tmpPar.Reserve = tmpPar.?????? & "|" & Split(strTmp, "|")(0)
                                        tmpPar.?????? = Split(strTmp, "|")(1)
                                    End If
                                    Exit For '??????????????,????????
                                Else
                                    '????????????,??????????,????????????????
                                    '????????????,????????????????
                                    If tmpPar.Reserve = "" And Left(tmpPar.??????, 1) = "&" Then
                                        tmpPar.Reserve = tmpPar.??????
                                    End If
                                    Select Case tmpPar.????
                                        Case 0, 3
                                            tmpPar.?????? = strTmp: Exit For
                                        Case 1
                                            If IsNumeric(strTmp) Then tmpPar.?????? = strTmp: Exit For
                                        Case 2
                                            If IsDate(strTmp) Then tmpPar.?????? = strTmp: Exit For
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        Next
    End If
    ReplaceUserPars = blnALL
End Function

Private Function ParCount(objReport As Report) As Integer
'????????????????????????????????????????
    Dim tmpPar As RPTPar, tmpData As RPTData, StrPar As String
    
    If objReport.Datas.count = 0 Then ParCount = 0: Exit Function
    For Each tmpData In objReport.Datas
        For Each tmpPar In tmpData.Pars
            If InStr(StrPar & ",", "," & tmpPar.???? & ",") = 0 Then
                StrPar = StrPar & "," & tmpPar.????
                ParCount = ParCount + 1
            End If
        Next
    Next
End Function

Private Sub ReplaceInputPars(objPars As RPTPars)
'??????????????????????????????????(????????)??????????????????????
    Dim tmpData As RPTData, tmpPar As RPTPar, objPar As RPTPar
    
    For Each tmpData In mobjReport.Datas
        For Each tmpPar In tmpData.Pars
            '??????????????????
            For Each objPar In objPars
                If objPar.???? = tmpPar.???? Then
                    tmpPar.?????? = objPar.??????
                    tmpPar.Reserve = objPar.Reserve
                    Exit For '??????????????
                End If
            Next
        Next
    Next
End Sub

Private Function OpenReportData(Optional ByVal blnAllReLoad As Boolean = True) As String
'??????????????????(mobjReport)????????????????????,????????????????????
'??????blnAllReLoad=????????????????????????(????????,??????????,????????????????????????)
'??????????="",????="????????"
    Dim tmpData As RPTData, strName As String
    Dim rsTmp As ADODB.Recordset
    Dim blnDo As Boolean, i As Integer
    
    '??????????????
    mobjReport.blnLoad = True '????????????????????????
    If mobjReport.Datas.count = 0 Then Exit Function
    
    If blnAllReLoad Then
        Set mLibDatas = Nothing
        Set mLibDatas = New LibDatas
    ElseIf mLibDatas Is Nothing Then
        Set mLibDatas = New LibDatas
    End If
            
    For Each tmpData In mobjReport.Datas
        '??????????????????????
        blnDo = True
        For i = 1 To mLibDatas.count
            If mLibDatas(i).Key = tmpData.???? Then
                blnDo = False: Exit For
            End If
        Next
        '????????????????????????
        If blnDo And DataUsed(mobjReport, tmpData.????, True) Then
            strName = tmpData.????
            Set rsTmp = Nothing
            Set rsTmp = OpenReportSQL(tmpData)
            If rsTmp Is Nothing Then
                OpenReportData = tmpData.????
                mobjReport.blnLoad = False
                Call ShowFlash: Exit Function
            End If
            mLibDatas.Add strName, rsTmp, "_" & strName
        End If
    Next
    
    Call ShowFlash
End Function

Private Function OpenReportSQL(objData As RPTData) As ADODB.Recordset
'??????????????????????????????????
'????????????ADO.Command??????????????????Clone????????????Clone????????????????????????Command??????????????????????
'1.????????????,????????????,Clone??????????????????,????????????????????
'2.??????????????????????,????Static????.????????????????????????,??????????????Static????????????
'  ????????????????????????????????,????????????????????????????
'3.??????????????????????????select '[0]' ???? from ...

    Dim rsTmp As New ADODB.Recordset
    Dim cmdData As New ADODB.Command
    Dim strLeft As String, strRight As String
    Dim StrPar As String, strParOld As String, bytPar As Byte
    Dim strSQL As String, strLog As String
    Dim intMax As Integer
    Dim strSQLtmp As String, i As Long, arrStr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim intDateType  As Integer  '0=????????????1=??????????2=????????????????,3=??????????????????????????????????????????
    Dim j As Long, k As Long, datValue As Date
    Dim L As Long

    If mbytStyle = 0 Or mbytStyle = 1 Then
        ShowFlash "????????????""" & objData.???? & """??????????????", , Me
    End If

    On Error GoTo errHandle

    '????????SQL
    'strSql = SQLOwner(TrimChar(objData.SQL), objData.????)
    strSQL = SQLOwner(RemoveNote(objData.SQL), objData.????)
    
    '??????????????????????????????????????/*+ XXX*/??????????????????
    strSQLtmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLtmp, 7)), 1, 2) <> "/*" And Mid(strSQLtmp, 1, 6) = "SELECT" Then
        arrStr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrStr)
            strSQLtmp1 = strSQLtmp
            Do While InStr(strSQLtmp1, arrStr(i)) > 0
                '????????????????IN ??????????Rule
                '??????????????SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrStr(i)) - 1)
                strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)               '??????3??????
                
                If strTmp = "IN(" Then '????in(select??????????????????????????????????????????????????????????????????
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrStr(i)) + Len(arrStr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrStr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
    
    strLog = strSQL
        
    i = 1
    Do While i <= Len(strLog)
        If InStr(i, strLog, "[") <= 0 Then
            i = i + 1
            GoTo makContinue1
        End If
        strLeft = Left(strLog, InStr(i, strLog, "[") - 1)
        strTmp = Mid(strLog, InStr(i, strLog, "["))
        If mdlPublic.AtString(strLeft) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '????????????????????????????[0-99]
            i = i + 1
            GoTo makContinue1
        End If
        
        If InStr(i, strLog, "]") <= 0 Then
            i = i + 1
            GoTo makContinue1
        End If
        strRight = Mid(strLog, InStr(i, strLog, "]") + 1)
        If strRight <> "" And mdlPublic.AtString(strRight) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '????????????????????????????[0-99]
            i = i + 1
            GoTo makContinue1
        End If
        
        '????????????????
        i = InStr(i, strLog, "[")
        strRight = Mid(strLog, InStr(i, strLog, "]") + 1)
        
        StrPar = Mid(strLog, InStr(i, strLog, "[") + 1, InStr(i, strLog, "]") - InStr(i, strLog, "[") - 1)
        strParOld = StrPar
        bytPar = Val(StrPar)
        Select Case objData.Pars("_" & CInt(bytPar)).????
            Case 0 '????
                StrPar = "'" & Replace(objData.Pars("_" & CInt(bytPar)).??????, "'", "''") & "'"
            Case 1 '????
                StrPar = objData.Pars("_" & CInt(bytPar)).??????
            Case 2 '????
                If Left(objData.Pars("_" & CInt(bytPar)).??????, 1) = "&" Then
                    StrPar = GetParSQLMacro(objData.Pars("_" & CInt(bytPar)).??????)
                Else
                    If Format(objData.Pars("_" & CInt(bytPar)).??????, "HH:mm:ss") = "00:00:00" Then
                        '??????????
                        StrPar = "To_Date('" & Format(objData.Pars("_" & CInt(bytPar)).??????, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                    Else
                        '??????????
                        StrPar = "To_Date('" & Format(objData.Pars("_" & CInt(bytPar)).??????, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                End If
            Case 3 '??????
                StrPar = objData.Pars("_" & CInt(bytPar)).??????
        End Select
        strLog = strLeft & StrPar & strRight
        
        i = Len(strLeft & strParOld)
        
makContinue1:
    Loop
        
    If InStr(UCase(objData.SQL), "--UNBOUND") > 0 Then GoTo LineOld

    '????????????SQL
    '????????????:????????????????
    cmdData.CommandText = "" '??????????????????????
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    L = 1
    Do While L <= Len(strSQL)
        If InStr(L, strSQL, "[") <= 0 Then
            L = L + 1
            GoTo makContinue2
        End If
        strLeft = Left(strSQL, InStr(L, strSQL, "[") - 1)
        strTmp = Mid(strSQL, InStr(L, strSQL, "["))
        If mdlPublic.AtString(strLeft) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '????????????????????????????[0-99]
            L = L + 1
            GoTo makContinue2
        End If
        
        If InStr(L, strSQL, "]") <= 0 Then
            L = L + 1
            GoTo makContinue2
        End If
        strRight = Mid(strSQL, InStr(L, strSQL, "]") + 1)
        If strRight <> "" And mdlPublic.AtString(strRight) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '????????????????????????????[0-99]
            L = L + 1
            GoTo makContinue2
        End If
        
        '????????????????
        L = InStr(L, strSQL, "[")
        strRight = Mid(strSQL, InStr(L, strSQL, "]") + 1)
        
        StrPar = Mid(strSQL, InStr(L, strSQL, "[") + 1, InStr(L, strSQL, "]") - InStr(L, strSQL, "[") - 1)
        strParOld = StrPar
        bytPar = Val(StrPar)
        intDateType = 0
        datValue = CDate(0)
        strTmp = ""
        
        Select Case objData.Pars("_" & CInt(bytPar)).????
            Case 0 '????
                StrPar = objData.Pars("_" & CInt(bytPar)).??????
                intMax = LenB(StrConv(StrPar, vbFromUnicode))
                
                If intMax <= 2000 Then
                    intMax = IIF(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarChar, adParamInput, intMax, StrPar)
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adLongVarChar, adParamInput, intMax, StrPar)
                End If

                strSQL = strLeft & "?" & strRight
            Case 1 '????
                StrPar = objData.Pars("_" & CInt(bytPar)).??????
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarNumeric, adParamInput, 30, Val(StrPar))

                strSQL = strLeft & "?" & strRight
            Case 2 '????
                If Left(objData.Pars("_" & CInt(bytPar)).??????, 1) = "&" Then
                    StrPar = GetParVBMacro(objData.Pars("_" & CInt(bytPar)).??????)
                Else
                    If Format(objData.Pars("_" & CInt(bytPar)).??????, "HH:mm:ss") = "00:00:00" Then
                        '??????????
                        StrPar = Format(objData.Pars("_" & CInt(bytPar)).??????, "yyyy-MM-dd")
                    Else
                        '??????????
                        StrPar = Format(objData.Pars("_" & CInt(bytPar)).??????, "yyyy-MM-dd HH:mm:ss")
                    End If
                End If
'                1????????????????????????????????????????????????????????
'                2????????????????????????????????????????????????????????SQL????????????????????????????????????????XX+1/24??????????????????????????????????????
'                      ??????????????????????????????????????sysdate??????????????????????????????????????????(sql????)??
'                      ??????????????????????????????????????????????????????????????????????????????????????????
 '                     ??????????????????????1??+1-1/24/60/60  2?? -1/24/60/60+1 ??3??-1/24/60/60  4??????????????????????????????
 '                     ????????????????????????????+1/24 ????????????????????????????????????
 '????SQL??
'                select * from ?????? where ???????? >1+ [0]- 1  and  ID>0
'                Union All
'                select * from ?????? where  [0]- 1=????????  and  ID>0
'                Union All
'                select * from ?????? where  ????????>[0]+1 - 1/24 /60 /60  and  ID>0
'                Union All
'                select * from ?????? where  ????????>[0] - 1/24 /60 /60+1  and  ID>0
'                Union All
'                select * from ?????? where  ????????>[0] - 1 /24 /60 /60  and  ID>0
'                Union All
'                select * from ?????? where  ????????>[0] - 1 /24 /60   and  ID>0
'                Union All
'                select * from ?????? where  ????????>1+[0] - 1 /24 /60/60   and  ID>0

                '????????????????????????
                datValue = CDate(StrPar)
                
                For i = 1 To Len(strRight)
                    If Mid(strRight, i, 1) <> " " Then
                        If InStr("+-", Mid(strRight, i, 1)) > 0 Then
                            For j = i + 1 To Len(strRight)
                                If Mid(strRight, j, 1) <> " " Then
                                    '????????????
                                    For k = j + 1 To Len(strRight)
                                        If Mid(strRight, k, 1) = " " Or (IsNumeric(Mid(strRight, j, 1)) And Not IsNumeric(Mid(strRight, j, k - j + 1))) Then
                                            If Not Mid(strRight, k, 1) = " " And Not IsNumeric(Mid(strRight, k - 1, 1)) Then
                                                k = k - 1
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    If IsNumeric(Mid(strRight, j, k - j)) Then
                                        intDateType = 1
                                        
                                        '??????????
                                        '??????????????????????????
                                        If InStr(Replace(strRight, " ", ""), "+1-1/24/60/60") = 1 Then
                                            datValue = datValue + 1 - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "60") + 2), "60") + 2 + InStr(strRight, "60") + 1)
                                        ElseIf InStr(Replace(strRight, " ", ""), "-1/24/60/60+1") = 1 Then
                                            datValue = datValue + 1 - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "+") + 1), "1") + InStr(strRight, "+") + 1)
                                        ElseIf InStr(Replace(strRight, " ", ""), "-1/24/60/60") = 1 Then
                                            datValue = datValue - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "60") + 2), "60") + 2 + InStr(strRight, "60") + 1)
                                        Else
                                            If Mid(strRight, i, 1) = "+" Then
                                                datValue = datValue + Val(Mid(strRight, j, k - j))
                                            Else
                                                datValue = datValue - Val(Mid(strRight, j, k - j))
                                            End If
                                            strTmp = Mid(strRight, k)
                                        End If
                                        If InStr("+-*/", Mid(Replace(strTmp, " ", ""), 1, 1)) > 0 And Replace(strRight, " ", "") <> "" Then
                                            '????????????+-*/????????????????????,??????????????????
                                            intDateType = 3
                                        End If
                                    Else
                                        intDateType = 2
                                    End If
                                    Exit For
                                End If
                            Next
                        Else
                            Exit For
                        End If
                        Exit For
                    End If
                Next
                '????????????????????????????????????????????
                If intDateType <> 2 Then
                    For i = Len(strLeft) To 1 Step -1
                        If Mid(strLeft, i, 1) <> " " Then
                            If InStr("+-", Mid(strLeft, i, 1)) > 0 Then
                               intDateType = 3
                            End If
                            Exit For
                        End If
                    Next
                End If
                If intDateType = 2 Then
                    '??????????????
                    strSQL = strLeft & "To_Date('" & Format(datValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & strRight
                ElseIf intDateType = 3 Then
                    '??????????????????????SQL??????,??????????????????????????????????????
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarChar, adParamInput, Len(StrPar), StrPar)
                    If StrPar Like "*:*:*" Then
                        strSQL = strLeft & "To_Date(?,'YYYY-MM-DD HH24:MI:SS')" & strRight
                    Else
                        strSQL = strLeft & "To_Date(?,'YYYY-MM-DD')" & strRight
                    End If
                Else
                    '???????????????????? ????????????????????????????????
                    If intDateType = 1 Then strRight = strTmp
                    '????????????????????????????????????????????????
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adDBTimeStamp, adParamInput, , datValue)
                    strSQL = strLeft & "?" & strRight
                End If
            Case 3 '??????
                StrPar = objData.Pars("_" & CInt(bytPar)).??????

                strSQL = strLeft & StrPar & strRight
        End Select
        
        L = Len(strLeft & strParOld)
        
makContinue2:
    Loop

    '??????????????
'    If cmdData.ActiveConnection Is Nothing Then
'        Set cmdData.ActiveConnection = gcnOracle '??????????
'    End If
    Set cmdData.ActiveConnection = mdlPublic.GetDBConnection(objData.????????????)
    cmdData.CommandText = strSQL

LineBand:
    Call SQLTest(App.ProductName, "OpenReportSQL", strLog)
    Set rsTmp = cmdData.Execute
    Call SQLTest
    Set OpenReportSQL = rsTmp
    Exit Function
    
LineOld:
    Call OpenRecord(rsTmp, strLog, "OpenReportSQL", objData.????????????)
    Set OpenReportSQL = rsTmp
    Exit Function
    
LineBlob:
    Dim cn As New ADODB.Connection
    Dim strServerName As String, strUserName As String, strUserPwd As String
    Dim arrTmp As Variant
    
    If objData.???????????? <= 0 Then
        arrTmp = Split(gcnOracle.ConnectionString, ";")
        strServerName = Replace(Split(arrTmp(2), "Server=")(1), """", "")
        strUserName = Split(arrTmp(4), "User ID=")(1)
        strUserPwd = Split(arrTmp(5), "Password=")(1)
        cn.CursorLocation = adUseClient
        cn.Provider = "OraOLEDB.Oracle"
        cn.Open "PLSQLRSet=1;DistribTx=0;Persist Security Info=True;Data Source=" & strServerName, strUserName, strUserPwd
        Set cmdData.ActiveConnection = cn
    Else
        Set cmdData.ActiveConnection = mdlPublic.GetDBConnectionEx(Val("1-OracleOLEDB"), objData.????????????)
    End If
    
    If Not cmdData.ActiveConnection Is Nothing Then
        Set rsTmp = cmdData.Execute
        Set OpenReportSQL = rsTmp
    End If
    Exit Function

errHandle:
    'ORA-00979:???? GROUP BY ??????
    'SQL????"?"????????Oracle????ADO????????":P1,:P2"????,Group by??????????????????????,??????????????
    '????????????,ADO??SQL??????????":P"????????(????????????????)
    '????????????,ADO??SQL??????????":P"????????,????Parameters????????????????????????,??????????.
    '    ????????????????Group????????????????????????,??????????????????????????????????
    If Err.Description Like "*ORA-00979*" Then Err.Clear: GoTo LineOld
    
    'ORA-00932: ??????????????: ???? NUMBER, ???????? -
    '??Group By Rollup??Decode????????????????????????????????????????,??????????????
    '??????????????????????????SQL??????????????????????????????
    If Err.Description Like "*ORA-00932*" Then Err.Clear: GoTo LineOld
    'MS??ODBC??????????BLOB????????????(????????????????????????????????????????)??????????????OraOLEDB??????????????
    If Err.Number = -2147467259 Then Err.Clear: GoTo LineBlob

    Call ShowFlash
    If Err.Description Like "*ORA-00920*" Then
        MsgBox "??????????????????????????????????""" & objData.???? & """??", vbExclamation, App.Title
    ElseIf ErrCenter() = 1 Then
        If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "????????????""" & objData.???? & """??????????????", , Me
        Resume
    End If
    Call SaveErrLog
End Function

Private Function EvalFormula(ByVal strFormula As String, idx As Integer, Row As Long) As String
'????????????????????
'??????strFormula=????????,idx:??????????????????????????,Row=??????
'????????????????,??????????????
'??????mLibDatas
    Dim strLeft As String, strRight As String, strVar As String
    
    On Error Resume Next
    
    strFormula = Trim(strFormula)
    
    If strFormula = "" Then '????
        Exit Function
    ElseIf InStr(strFormula, "[") = 0 Then '????????
        EvalFormula = Srt.Eval(strFormula)
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         '????????????????
         EvalFormula = GetFieldValue(Me, Mid(strFormula, 2, Len(strFormula) - 2))
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         '??????????????
         EvalFormula = msh(idx).TextMatrix(Row, CInt(Mid(strFormula, 2, Len(strFormula) - 2)))
    Else '????????
        Do While InStr(strFormula, "[") > 0
            strLeft = Left(strFormula, InStr(strFormula, "[") - 1)
            strRight = Mid(strFormula, InStr(strFormula, "]") + 1)
            strVar = Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1)
            
            If IsNumeric(Mid(strVar, 2)) And Left(strVar, 1) = "@" Then
                If Row = msh(idx).FixedRows Then
                    strVar = "" '????????????????
                Else
                    If InStr(strFormula, """[" & strVar & "]""") > 0 And InStr(msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))), """") > 0 Then
                        '??????????????????????????????
                        strVar = Replace(msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))), """", """""")
                    Else
                        strVar = msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))) '??????????????
                    End If
                End If
                If strVar = "" Then strVar = 0
            ElseIf IsNumeric(strVar) Then
                If InStr(strFormula, """[" & strVar & "]""") > 0 And InStr(msh(idx).TextMatrix(Row, CInt(strVar)), """") > 0 Then
                    '??????????????????????????????
                    strVar = Replace(msh(idx).TextMatrix(Row, CInt(strVar)), """", """""")
                Else
                    strVar = msh(idx).TextMatrix(Row, CInt(strVar)) '??????????????
                End If
                If strVar = "" Then strVar = 0
            ElseIf InStr(strVar, ".") > 0 Then
                '????????,????"Null",??????????????????
                strVar = GetFieldValue(Me, strVar, True) '??????????????????????????,????????????
            End If
            
            '??????????????
            If InStr(strVar, "[") > 0 Or InStr(strVar, "]") > 0 Then
                strVar = Replace(strVar, "[", Chr(1) & "SKIPCYCLEFT" & Chr(1))
                strVar = Replace(strVar, "]", Chr(1) & "SKIPCYCRIGHT" & Chr(1))
            End If
            strFormula = strLeft & strVar & strRight
        Loop
        strFormula = Replace(strFormula, Chr(1) & "SKIPCYCLEFT" & Chr(1), "[")
        strFormula = Replace(strFormula, Chr(1) & "SKIPCYCRIGHT" & Chr(1), "]")
        EvalFormula = Srt.Eval(strFormula)
    End If
End Function

Private Function SortFormula(objItem As RPTItem) As Variant
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrFormula() As String, strTmp As String
    Dim strReferCols As String, intReferCols As Integer
    Dim intCol As Integer, intCur As Integer
    Dim i As Integer, j As Integer
    Dim strDie As String, strOrder As String
    
    
    ReDim arrFormula(objItem.SubIDs.count - 1) As String
    
    '????????????????
    For Each tmpID In objItem.SubIDs
        Set tmpItem = mobjReport.Items("_" & tmpID.ID)
        arrFormula(tmpItem.????) = tmpItem.???? & "|" & tmpItem.???? & "|" & tmpItem.???? & "|" & tmpItem.????
    Next
    
    '????"????"??????????????
    i = 0
    strOrder = GetOrder(arrFormula)
    Do While i <= UBound(arrFormula)
        '????????????
        strReferCols = GetReferCols(CStr(Split(arrFormula(i), "|")(0)))
        intReferCols = UBound(Split(strReferCols, ","))
        
        intCur = i '????????????
        For j = 0 To intReferCols
            '??????????????
            intCol = GetReferLoc(arrFormula, CInt(Split(strReferCols, ",")(j)))
            If intCol > intCur Then
                strTmp = arrFormula(intCur)
                arrFormula(intCur) = arrFormula(intCol)
                arrFormula(intCol) = strTmp
                intCur = intCol
            End If
        Next
        '????????????????,????????????,????????????????
        '????????????????????????????
        strDie = GetOrder(arrFormula)
        If intCur = i Or (intCur <> i And strOrder = strDie) Then
            i = i + 1
            strOrder = strDie
        End If
    Loop
    
    SortFormula = arrFormula
End Function

Private Function GetOrder(arrFormula() As String) As String
'????????????????????????????????????????,??????????????
    Dim i As Integer
    For i = 0 To UBound(arrFormula)
        GetOrder = GetOrder & "," & CInt(Split(arrFormula(i), "|")(2))
    Next
    GetOrder = Mid(GetOrder, 2)
End Function

Private Function GetReferLoc(arrFormula() As String, intCol As Integer) As Integer
'????????????????intCol????????????????????
    Dim i As Integer
    For i = 0 To UBound(arrFormula)
        If CInt(Split(arrFormula(i), "|")(2)) = intCol Then
            GetReferLoc = i: Exit Function
        End If
    Next
End Function

Private Function GetReferCols(ByVal strFormula As String) As String
'??????????????????????????,??"3,5,6"
    Dim strRight As String, strCol As String, strCols As String
    
    strFormula = Trim(strFormula)
    
    Do While InStr(strFormula, "[") > 0
        strRight = Mid(strFormula, InStr(strFormula, "]") + 1)
        strCol = Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1)
        If IsNumeric(strCol) Then strCols = strCols & "," & strCol
        strFormula = strRight
    Loop
    GetReferCols = Mid(strCols, 2)
End Function

Private Sub SetRedraw(blnDraw As Boolean)
    Dim obj As Object
    For Each obj In msh
        If obj.Index <> 0 And (obj.Container Is picPaper(intReport) Or UCase(obj.Container.name) = "PIC") Then obj.Redraw = blnDraw
    Next
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Function GetLR(msh As Object, Col As Integer) As Byte
    Select Case msh.ColAlignment(Col)
        Case 0, 1, 2 '??????
            GetLR = 2 '????????
        Case 3, 4, 5 '??????
            GetLR = 1 '????????
        Case 6, 7, 8 '??????
            GetLR = 0 '????????
    End Select
End Function

Private Function GetRowText(msh As Object, Row As Long, Col As Long) As String
    Dim i As Integer
    Dim strTmp As String
    
    For i = 0 To Col
        strTmp = strTmp & Trim(msh.TextMatrix(Row, i))
    Next
    GetRowText = strTmp
End Function

Private Function GetColText(msh As Object, Row As Long, Col As Long) As String
    Dim i As Integer
    Dim strTmp As String
    
    For i = 0 To Row
        strTmp = strTmp & Trim(msh.TextMatrix(i, Col))
    Next
    GetColText = strTmp
End Function

Private Function GetColType(ByVal strFormula As String) As Byte
'??????????????????????????????????
'??????strFormula=??????????
'??????0=??????,1-????(????),2=????,3=????
'??????mLibDatas
    Dim varR As Variant, strData As String, strField As String
    
    On Error Resume Next
    
    strFormula = Trim(strFormula)
    
    If strFormula = "" Then '????
        GetColType = 1
    ElseIf InStr(strFormula, "[") = 0 Then '????????
        varR = Srt.Eval(strFormula)
        If IsNumeric(varR) Then
            GetColType = 2
        ElseIf IsDate(varR) Then
            GetColType = 3
        Else
            GetColType = 1
        End If
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         '????????????????
        strFormula = Mid(strFormula, 2, Len(strFormula) - 2)
        strData = Left(strFormula, InStr(strFormula, ".") - 1)
        strField = Mid(strFormula, InStr(strFormula, ".") + 1)
        
        Select Case mLibDatas("_" & strData).DataSet.Fields(strField).type
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                GetColType = 1
            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                GetColType = 2
            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                GetColType = 3
        End Select
    End If
End Function

Private Function GetParsStr(ByVal objPars As RPTPars) As String
'??????????????????????????????????????????????
'??????"??????=??????|??????=??????..."
'??????????????????????????,????????"ReportFormat=x"
    Dim tmpPar As RPTPar
    Dim strPars As String
    
    If mobjReport.Fmts.count > 1 Then
        strPars = strPars & "|ReportFormat=" & bytFormat
    End If
    
    For Each tmpPar In objPars
        If tmpPar.?????? Like "&*" And tmpPar.???? = 2 Then
            strPars = strPars & "|" & tmpPar.???? & "=" & GetParVBMacro(tmpPar.??????)
        Else
            strPars = strPars & "|" & tmpPar.???? & "=" & tmpPar.??????
        End If
    Next
    GetParsStr = Mid(strPars, 2)
End Function

Private Function SetStrPars(ByVal strPars As String, ByVal objPars As RPTPars) As RPTPars
'??????????????????????????????????????????????
'??????strPars="??????=??????|??????=??????..."
'????????????????????????
'??????????????????????????????????????????????????,??????
    Dim tmpPar As RPTPar, tmpPars As RPTPars
    Dim i As Integer, j As Integer
    Dim bytTmp As Byte, strTmp As String
    
    If strPars = "" Or Not strPars Like "*=*" Then Set SetStrPars = objPars: Exit Function
    
    Set tmpPars = objPars
    
    For i = 0 To UBound(Split(strPars, "|"))
        For Each tmpPar In tmpPars
            strTmp = Split(strPars, "|")(i)
            If UCase(Split(strTmp, "=")(0)) = UCase("ReportFormat") And mobjReport.Fmts.count > 1 Then
                If IsNumeric(Split(strTmp, "=")(1)) Then
                    bytTmp = CByte(Split(strTmp, "=")(1))
                    For j = 1 To cboFormat.ComboItems.count
                        If CByte(Mid(cboFormat.ComboItems(j).Key, 2)) = bytTmp Then
                            cboFormat.ComboItems(j).Selected = True
                            bytFormat = bytTmp: mobjReport.bytFormat = bytFormat: Exit For
                        End If
                    Next
                End If
            ElseIf UCase(tmpPar.????) = UCase(Split(strTmp, "=")(0)) Then
                Select Case tmpPar.????
                    Case 1 '??????
                        If IsNumeric(Split(strTmp, "=")(1)) Then tmpPar.?????? = Split(strTmp, "=")(1)
                    Case 2 '??????
                        If IsDate(Split(strTmp, "=")(1)) Then tmpPar.?????? = Split(strTmp, "=")(1)
                    Case Else
                        tmpPar.?????? = Split(strTmp, "=")(1)
                End Select
            End If
        Next
    Next
    Set SetStrPars = tmpPars
End Function

Private Sub mnuFile_Excel_Click()
    Dim intRow As Integer, intCol As Integer
    Dim bytKind As Byte, tmpMsh As Object
    Dim i As Long, j As Long
    
    '????????????
    If Not mobjReport.blnLoad Then Exit Sub
    
    If zlRegInfo("????????") <> "1" Then
        MsgBox "??????????????????????????????", vbInformation, App.Title
        Exit Sub
    End If
    
    If isExporting Then
        gblnError = True
        MsgBox "?????????????????????? Excel,????????????????????", vbInformation, App.Title
        Exit Sub
    End If
    
    If intGridCount = 0 Then
        MsgBox "?????????????????????????????? Excel??", vbInformation, App.Title
        Exit Sub
    End If
    If objCurGrid Is Nothing Then
        If msh.count > 1 Then
            For Each tmpMsh In msh
                If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") And Not tmpMsh.Tag Like "H_*" Then
                    Set objCurGrid = tmpMsh
                    Exit For
                End If
            Next
        End If
        If objCurGrid Is Nothing Then
            MsgBox "???????????????????? Excel??????????", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    If Not HaveExcel Then
        gblnError = True
        MsgBox "?????????????????????? Microsoft Excel ????,??????????????", vbInformation, App.Title
        Exit Sub
    End If
    
    '????[????]??????
    Set gobjHead = Nothing
    Set gobjBody = Nothing
    
    Set gobjBody = objCurGrid
    If Val(objCurGrid.Tag) > 0 Then
        bytKind = GetGridStyle(mobjReport, objCurGrid.Index)
        If bytKind <> 2 Then Set gobjHead = msh(CInt(objCurGrid.Tag))
    End If
    
    timHead.Enabled = False
    
    '????????????????
    Call MakeAppend(Me, picPaper(intReport))
    
    '??????Excel
    intRow = gobjBody.Row
    intCol = gobjBody.Col
    If Not gobjHead Is Nothing Then gobjHead.Redraw = False
    gobjBody.Redraw = False
    
    '??????????????
    If Not gobjHead Is Nothing Then
        For i = 0 To gobjHead.Rows - 1
            For j = 0 To gobjHead.Cols - 1
                gobjHead.TextMatrix(i, j) = Replace(Replace(Replace(gobjHead.TextMatrix(i, j), vbCrLf, "<??????????>"), vbLf, "<??????????>"), vbCr, "<??????????>")
            Next
        Next
    End If
    For i = 0 To gobjBody.Rows - 1
        For j = 0 To gobjBody.Cols - 1
            gobjBody.TextMatrix(i, j) = Replace(Replace(Replace(gobjBody.TextMatrix(i, j), vbCrLf, "<??????????>"), vbLf, "<??????????>"), vbCr, "<??????????>")
        Next
    Next
    
    blnExcel = True
    Call ExportExcel(Me, IIF(mbytStyle = 3, mstrExcelFile, ""))
    
    gobjBody.Row = intRow
    gobjBody.Col = intCol
    Call msh_EnterCell(gobjBody.Index)
    If Not gobjHead Is Nothing Then
        gobjHead.Redraw = True
    
        '??????????????
        For i = 0 To gobjHead.Rows - 1
            For j = 0 To gobjHead.Cols - 1
                gobjHead.TextMatrix(i, j) = Replace(gobjHead.TextMatrix(i, j), "<??????????>", vbCrLf)
            Next
        Next
    End If
    For i = 0 To gobjBody.Rows - 1
        For j = 0 To gobjBody.Cols - 1
            gobjBody.TextMatrix(i, j) = Replace(gobjBody.TextMatrix(i, j), "<??????????>", vbCrLf)
        Next
    Next
    gobjBody.Redraw = True
    
    timHead.Enabled = True
End Sub

Public Function DelUnUseData(objReport As Report) As Boolean
'????????????mobjReport????????????????????????
'????????????????????????????
'??????1.????????????????????????????
'      2.??????????????????????????????
    Dim tmpData As RPTData
    
    If objReport Is Nothing Then Exit Function
    
    For Each tmpData In objReport.Datas
        If Not DataUsed(objReport, tmpData.????) Then objReport.Datas.Remove "_" & tmpData.Key
    Next
End Function

Private Function GetStatText(strStat As String) As String
    Select Case strStat
        Case "SUM"
            GetStatText = "????"
        Case "AVG"
            GetStatText = "??????"
        Case "MAX"
            GetStatText = "??????"
        Case "MIN"
            GetStatText = "??????"
        Case "COUNT"
            GetStatText = "??????"
    End Select
End Function

Public Sub AddCol(msh As Object, Optional ByVal intCol As Integer = -1, Optional ByVal intCols As Integer = 1)
'????????????????msh??????intCols????,????????????????????intCol,????????intCol????,????????
'??????????????,????????????,??????????????(????????)
    Dim i As Integer, j As Integer, k As Integer
    
    If intCol >= msh.Cols Then intCol = -1
    msh.Cols = msh.Cols + intCols
    If intCol = -1 Then Exit Sub
    '????????
    For j = msh.Cols - 1 To intCol + intCols Step -intCols
        For i = 0 To msh.FixedRows - 1
            For k = 0 To intCols - 1
                msh.TextMatrix(i, j - k) = msh.TextMatrix(i, j - k - intCols)
                msh.ColData(j - k) = msh.ColData(j - k - intCols)
            Next
        Next
    Next
    '????????????
    For j = intCol To intCol + intCols - 1
        For i = 0 To msh.FixedRows - 1
            msh.TextMatrix(i, j) = ""
        Next
    Next
End Sub

Private Sub ShowFreeGrid(objItem As RPTItem, lngW As Long, lngH As Long)
'??????????????????????????????????????
    Dim strData As String, strTmp As String, bytKind As Byte
    Dim lngCol As Long, strState As String, arrState() As Variant
    Dim mshBody As Object, mshHead As Object
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim strValue As String, lngHead As Long, arrHead() As String
    Dim strSource As String, arrSource() As String, arrFormula() As String
    Dim strFormula As String, strFormat As String, arrType() As Long
    Dim i As Long, j As Long, k As Long, L As Long, blnDo As Boolean
    Dim arrRowIDs() As Variant, strIDSource As String, strFirstSource As String
    Dim objPic As StdPicture
    Dim colRelation As Collection
    Dim objColProtertys As RPTColProtertys
    Dim varIFValue As Variant
    Dim blnTmp As Boolean, blnRPTLink As Boolean
    
    arrRowIDs = Array()
    
    With objItem
        bytKind = GetGridStyle(mobjReport, .ID)
        Load msh(.ID) '????????
        Load msh(.SubIDs(1).ID) '????????
        Set msh(.ID).Container = picPaper(intReport)
        Set msh(.SubIDs(1).ID).Container = picPaper(intReport)
        If .??ID <> 0 Then
            Set msh(.ID).Container = pic(.??ID)
            Set msh(.SubIDs(1).ID).Container = pic(.??ID)
        End If
        Set mshBody = msh(.ID)
        Set mshHead = msh(.SubIDs(1).ID) '????????????ID????????????
        
        mshBody.Redraw = False
        mshHead.Redraw = False
                            
        mshHead.Tag = "H_" & mshBody.Index '????????????????????
        mshBody.Tag = mshHead.Index
        
        '????????
        '????
        mshHead.Left = .X: mshHead.Top = .Y
        mshHead.Width = .W: mshHead.Height = .H '????????????????
        
        mshHead.Cols = .SubIDs.count
        mshHead.FixedCols = 0
        mshHead.Rows = UBound(Split(mobjReport.Items("_" & .SubIDs(1).ID).????, "|")) + 2
        mshHead.RowHeight(mshHead.Rows - 1) = 0
        mshHead.FixedRows = mshHead.Rows - 1
        
        mshHead.ForeColor = .????
        mshHead.ForeColorFixed = .????
        mshHead.BackColor = .????
        mshHead.BackColorFixed = .????
        mshHead.GridColor = .????
        mshHead.GridColorFixed = IIF(.???? = "", .????, Val(.????))
        mshHead.Font.name = .????
        mshHead.Font.Size = .????
        mshHead.Font.Bold = .????
        mshHead.Font.Italic = .????
        mshHead.Font.Underline = .????
        mshHead.GridLineWidth = IIF(.??????????, 2, 1)
        'Set mshHead.FontFixed = mshHead.Font
        '????????
        mshHead.ExplorerBar = flexExSortShow
        
        '????????????(????????????????????,????)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.ID)
            If tmpItem.Relations.count > 0 Then
                If blnRPTLink = False Then blnRPTLink = True
            End If
            arrHead = Split(tmpItem.????, "|")
            lngHead = 0 '??????????????
            For i = 0 To UBound(arrHead) '????^????^????
                mshHead.Col = tmpItem.????: mshHead.Row = i
                mshHead.CellAlignment = CInt(Split(arrHead(i), "^")(0))
                
                mshHead.RowHeight(i) = CLng(Split(arrHead(i), "^")(1))
                lngHead = lngHead + mshHead.RowHeight(i)
                
                If CStr(Split(arrHead(i), "^")(2)) = "#" Then '????
                    mshHead.TextMatrix(i, tmpItem.????) = ""
                ElseIf CStr(Split(arrHead(i), "^")(2)) = "??" Then '????????????????
                    mshHead.TextMatrix(i, tmpItem.????) = mshHead.TextMatrix(i, tmpItem.???? - 1)
                ElseIf CStr(Split(arrHead(i), "^")(2)) = "??" Then '????????????????
                    mshHead.TextMatrix(i, tmpItem.????) = mshHead.TextMatrix(i - 1, tmpItem.????)
                Else
                    strValue = CStr(Split(arrHead(i), "^")(2))
                    
                    '????????????(????????????????????????????)
                    '??????????????(??????????????????)
                    strData = GetLabelDataName(strValue)
                    If strData <> "" Then
                        For j = 0 To UBound(Split(strData, "|"))
                            strTmp = Split(Split(strData, "|")(j), ".")(0)
                            If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                mLibDatas("_" & strTmp).DataSet.MoveFirst
                            End If
                            strTmp = GetFieldValue(Me, CStr(Split(strData, "|")(j)))
                            strValue = Replace(strValue, "[" & Split(strData, "|")(j) & "]", strTmp)
                        Next
                    End If
                    
                    '??????????????:[=??????]??[n>=0]??[??????????][????????]
                    strValue = GetLabelMacro(Me, strValue)
                    
                    mshHead.TextMatrix(i, tmpItem.????) = strValue
                    
                End If
                If UBound(Split(arrHead(i), "^")) > 3 Then
                    '??????????????????
                    If Split(arrHead(i), "^")(3) = 1 Then
                        mshHead.Cell(flexcpFontBold, i, tmpItem.????) = True
                    End If
                    mshHead.Cell(flexcpForeColor, i, tmpItem.????) = Val(Split(arrHead(i), "^")(4))
                End If
            Next
            mshHead.ColWidth(tmpItem.????) = tmpItem.W
        Next
        
        '????????????
        For i = 0 To mshHead.FixedRows - 1
            mshHead.MergeRow(i) = True
        Next
        For i = 0 To mshHead.Cols - 1
            mshHead.MergeCol(i) = True
        Next
        
        '????????
        If bytKind = 2 Then '????????
            mshBody.Top = .Y: mshBody.Left = .X
            mshBody.Height = .H: mshBody.Width = .W
        Else
            mshBody.Top = .Y + lngHead: mshBody.Left = .X
            If .H - lngHead + 15 < 0 Then
                mshBody.Height = 0
            Else
                mshBody.Height = .H - lngHead + 15
            End If
            mshBody.Width = .W
        End If
        mshBody.Cols = .SubIDs.count: mshBody.FixedCols = 0
        mshBody.Rows = 1: mshBody.FixedRows = 0 '??????????????????
        mshBody.RowHeight(0) = .????
        
        mshBody.ForeColor = .????
        mshBody.ForeColorFixed = .????
        mshBody.BackColor = .????
        mshBody.BackColorFixed = .????
        mshBody.GridColor = .????
        mshBody.GridColorFixed = .????
        mshBody.Font.name = .????
        mshBody.Font.Size = .????
        mshBody.Font.Bold = .????
        mshBody.Font.Italic = .????
        mshBody.Font.Underline = .????
        mshBody.GridLineWidth = IIF(.??????????, 2, 1)
        'Set mshBody.FontFixed = mshBody.Font
        
        '????????(????????????)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.ID)
            mshBody.ColData(tmpItem.????) = tmpItem
            mshBody.ColWidth(tmpItem.????) = tmpItem.W
            mshBody.ColAlignment(tmpItem.????) = Switch(tmpItem.???? = 0, 1, tmpItem.???? = 1, 4, tmpItem.???? = 2, 7)
            If mshBody.FixedRows - 1 >= 0 And mshBody.Rows - 1 >= 0 Then mshBody.Cell(flexcpAlignment, mshBody.FixedRows - 1, tmpItem.????, mshBody.Rows - 1, tmpItem.????) = mshBody.ColAlignment(tmpItem.????)
            mshBody.MergeCol(tmpItem.????) = tmpItem.????
        Next
        
        '--------------------------------------------------------------------------------------
        '????????????
        '--------------------------------------------------------------------------------------
        '1.??????????????????????
        strSource = GetGridSource(objItem) '"????????,????????,..."
        
        '2.??????????????????????
        arrFormula = SortFormula(objItem) '(????????="????|????|????|????")
        
        '3.????????????
        ReDim arrState(.SubIDs.count - 1)
        ReDim arrType(.SubIDs.count - 1) '????????????(0=??????,1-????(????),2-????,3-????)
        If strSource <> "" Then
            arrSource = Split(strSource, ",")
            strFirstSource = arrSource(0)
            ''??????????????????????:??????????????,????????
            blnDo = False
            For i = 0 To UBound(arrSource)
                If mLibDatas("_" & arrSource(i)).DataSet.RecordCount > 0 Then
                    mLibDatas("_" & arrSource(i)).DataSet.MoveFirst '????????
                End If
                blnDo = blnDo Or Not mLibDatas("_" & arrSource(i)).DataSet.EOF
                
                '??????ID????????????:????????????
                If strIDSource = "" Then
                    For j = 0 To mLibDatas("_" & arrSource(i)).DataSet.Fields.count - 1
                        If UCase(mLibDatas("_" & arrSource(i)).DataSet.Fields(j).name) = "ID" Then
                            If IsType(mLibDatas("_" & arrSource(i)).DataSet.Fields(j).type, adNumeric) Then
                                strIDSource = arrSource(i): Exit For
                            End If
                        End If
                    Next
                End If
            Next
        Else
            blnDo = True
        End If
        
        '4.????????
        j = 0
        Do While blnDo
            If j > 0 Then
                mshBody.Rows = mshBody.Rows + 1 '??????????
                mshBody.RowHeight(mshBody.Rows - 1) = .????
            End If
            
            '????????????ID??????????????
            ReDim Preserve arrRowIDs(UBound(arrRowIDs) + 1)
            arrRowIDs(UBound(arrRowIDs)) = 0
            If strIDSource <> "" Then
                If Not mLibDatas("_" & strIDSource).DataSet.EOF Then
                    arrRowIDs(UBound(arrRowIDs)) = Val(Nvl(mLibDatas("_" & strIDSource).DataSet.Fields("ID").Value, 0))
                End If
            End If
            
            For i = 0 To UBound(arrFormula)
                strFormula = Split(arrFormula(i), "|")(0) '????
                strFormat = Split(arrFormula(i), "|")(1) '????
                lngCol = CInt(Split(arrFormula(i), "|")(2)) '????
                strState = Split(arrFormula(i), "|")(3) '????
                
                
                '??????????
                strValue = EvalFormula(strFormula, mshBody.Index, j)
                If gobjFile.FileExists(strValue) Then
                    Set objPic = Nothing
                    On Error Resume Next
                    Set objPic = LoadPicture(strValue)
                    gobjFile.DeleteFile strValue, True
                    On Error GoTo 0
                    
                    If Not objPic Is Nothing Then
                        mshBody.Row = j: mshBody.Col = lngCol
                        
                        Me.picTemp.Cls '??????????
                        If objPic.Height / objPic.Width < mshBody.CellHeight / mshBody.CellWidth Then
                            Me.picTemp.Width = mshBody.CellWidth
                            Me.picTemp.Height = (objPic.Height / objPic.Width) * mshBody.CellWidth
                        Else
                            Me.picTemp.Height = mshBody.CellHeight
                            Me.picTemp.Width = (objPic.Width / objPic.Height) * mshBody.CellHeight
                        End If
                        Me.picTemp.PaintPicture objPic, 0, 0, Me.picTemp.Width, Me.picTemp.Height
                                            
                        Set mshBody.CellPicture = Me.picTemp.Image
                        mshBody.CellPictureAlignment = 4 '??????????
                    End If
                Else
                    mshBody.TextMatrix(j, lngCol) = strValue
                    If (strIDSource <> "" Or strFirstSource <> "") And blnRPTLink = True Then
                        Set colRelation = New Collection
                        colRelation.Add mLibDatas("_" & IIF(strIDSource = "", strFirstSource, strIDSource)).DataSet.AbsolutePosition
                        mshBody.Cell(flexcpData, j, lngCol) = colRelation
                    End If
                    '??????????
                    Set objColProtertys = mshBody.ColData(lngCol).ColProtertys
                    If objColProtertys.count > 0 Then
                        For L = 1 To objColProtertys.count
                            If InStr(objColProtertys.Item(L).??????, strIDSource & ".") > 0 Then
                                varIFValue = EvalFormula("[" & objColProtertys.Item(L).?????? & "]", mshBody.Index, j)
                            Else
                                varIFValue = objColProtertys.Item(L).??????
                            End If
                            If CheckColProtertys(EvalFormula("[" & objColProtertys.Item(L).???????? & "]", mshBody.Index, j), objColProtertys.Item(L).????????, varIFValue) Then
                                If objColProtertys.Item(L).???????????? Then
                                    mshBody.Cell(flexcpBackColor, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(L).????????
                                    mshBody.Cell(flexcpForeColor, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(L).????????
                                    mshBody.Cell(flexcpFontBold, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(L).????????
                                Else
                                    mshBody.Cell(flexcpBackColor, j, lngCol) = objColProtertys.Item(L).????????
                                    mshBody.Cell(flexcpForeColor, j, lngCol) = objColProtertys.Item(L).????????
                                    mshBody.Cell(flexcpFontBold, j, lngCol) = objColProtertys.Item(L).????????
                                End If
                            End If
                        Next
                    End If
                End If
                
                '????????????
                If j = 0 And (strState = "MAX" Or strState = "MIN") Then
                    arrType(lngCol) = GetColType(strFormula)
                    arrState(lngCol) = "??????"
                End If
                If strState = "MAX" Or strState = "MIN" Then
                    If arrType(lngCol) = 0 Then
                        If IsNumeric(mshBody.TextMatrix(j, lngCol)) Then
                            arrType(lngCol) = 2
                        ElseIf IsDate(mshBody.TextMatrix(j, lngCol)) Then
                            arrType(lngCol) = 3
                        Else
                            arrType(lngCol) = 1
                        End If
                    End If
                End If
                
                '????????
                On Error Resume Next
                If mshBody.TextMatrix(j, lngCol) <> "" Then
                    Select Case strState
                        Case "SUM", "AVG" '??????????(????)
                            If IsNumeric(mshBody.TextMatrix(j, lngCol)) Then
                                arrState(lngCol) = arrState(lngCol) + CDbl(mshBody.TextMatrix(j, lngCol))
                            ElseIf IsDate(mshBody.TextMatrix(j, lngCol)) Then
                                arrState(lngCol) = arrState(lngCol) + CDate(mshBody.TextMatrix(j, lngCol))
                            Else
                                arrState(lngCol) = arrState(lngCol) + mshBody.TextMatrix(j, lngCol)
                            End If
                        Case "MAX"
                            If arrState(lngCol) = "??????" Then
                                If arrType(lngCol) = 2 Then
                                    arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                ElseIf arrType(lngCol) = 3 Then
                                    arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                Else
                                    arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                End If
                            Else
                                If arrType(lngCol) = 2 Then
                                    If CDbl(mshBody.TextMatrix(j, lngCol)) > arrState(lngCol) Then
                                        arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                    End If
                                ElseIf arrType(lngCol) = 3 Then
                                    If CDate(mshBody.TextMatrix(j, lngCol)) > arrState(lngCol) Then
                                        arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                    End If
                                Else
                                    If mshBody.TextMatrix(j, lngCol) > arrState(lngCol) Then
                                        arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                    End If
                                End If
                            End If
                        Case "MIN"
                            If arrState(lngCol) = "??????" Then
                                If arrType(lngCol) = 2 Then
                                    arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                ElseIf arrType(lngCol) = 3 Then
                                    arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                Else
                                    arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                End If
                            Else
                                If arrType(lngCol) = 2 Then
                                    If CDbl(mshBody.TextMatrix(j, lngCol)) < arrState(lngCol) Then
                                        arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                    End If
                                ElseIf arrType(lngCol) = 3 Then
                                    If CDate(mshBody.TextMatrix(j, lngCol)) < arrState(lngCol) Then
                                        arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                    End If
                                Else
                                    If mshBody.TextMatrix(j, lngCol) < arrState(lngCol) Then
                                        arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                    End If
                                End If
                            End If
                        Case "COUNT"
                            arrState(lngCol) = arrState(lngCol) + 1
                    End Select
                End If
                
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                
                '??????????????,????????
                If strFormat <> "" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = Format(mshBody.TextMatrix(j, lngCol), strFormat)
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
            Next
            
            If strSource <> "" Then
                '??????????????,????????
                blnDo = False
                For i = 0 To UBound(arrSource)
                    If Not mLibDatas("_" & arrSource(i)).DataSet.EOF Then
                        mLibDatas("_" & arrSource(i)).DataSet.MoveNext
                    End If
                    blnDo = blnDo Or Not mLibDatas("_" & arrSource(i)).DataSet.EOF
                Next
            Else
                '??????????????????,??????????????
                blnDo = False
            End If
            
            j = j + 1
        Loop
        
        '5.??????????,????????????????,??????
        blnDo = False
        For i = 0 To UBound(arrFormula)
            blnDo = blnDo Or (Split(arrFormula(i), "|")(3) <> "")
        Next
        If blnDo Then
            mshBody.Rows = mshBody.Rows + 1
            mshBody.RowHeight(mshBody.Rows - 1) = .????
            For i = 0 To UBound(arrFormula)
                strState = Split(arrFormula(i), "|")(3) '????
                lngCol = Split(arrFormula(i), "|")(2) '????
                strFormat = Split(arrFormula(i), "|")(1) '????
                strFormula = Split(arrFormula(i), "|")(0) '????
                If strState = "AVG" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = arrState(lngCol) / j
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                ElseIf strState <> "" Then
                    If TypeName(arrState(lngCol)) = "String" Then
                        If arrState(lngCol) = "??????" Then arrState(lngCol) = ""
                    End If
                    mshBody.TextMatrix(j, lngCol) = arrState(lngCol)
                ElseIf strFormula <> "" Then
                    '????????????????????????????????????????????
                    strValue = EvalFormula(strFormula, mshBody.Index, j)
                    If gobjFile.FileExists(strValue) Then
                        '????????
                        On Error Resume Next
                        gobjFile.DeleteFile strValue, True
                        On Error GoTo 0
                    Else
                        mshBody.TextMatrix(j, lngCol) = strValue
                    End If
                End If
                '????????????
                If strFormat <> "" And mshBody.TextMatrix(j, lngCol) <> "" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = Format(mshBody.TextMatrix(j, lngCol), strFormat)
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
            Next
            '????????????
            For k = 0 To mshBody.Cols - 1
                If mshBody.ColWidth(k) > 0 Then Exit For
            Next
            If mshBody.TextMatrix(j, k) = "" Then
                blnDo = True: L = 0
                For i = 0 To UBound(arrFormula)
                    If Split(arrFormula(i), "|")(3) <> "" Then
                        If L = 0 Then
                            strState = Split(arrFormula(i), "|")(3)
                        Else
                            blnDo = blnDo And (Split(arrFormula(i), "|")(3) = strState)
                        End If
                        L = L + 1
                    End If
                Next
                If blnDo Then '????????????
                    mshBody.TextMatrix(j, k) = Switch(strState = "SUM", "????", strState = "AVG", "??????", strState = "MAX", "??????", strState = "MIN", "??????", strState = "COUNT", "??????")
                Else '????????????
                    mshBody.TextMatrix(j, k) = "????"
                End If
                mshBody.Row = j: mshBody.Col = k: mshBody.CellAlignment = 4
            End If
        End If
        
        For i = 0 To mshBody.Rows - 1
            mshBody.RowHeight(i) = .????
        Next
        '--------------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------

        '????????
        mshHead.WordWrap = True
        mshBody.WordWrap = True
        
        mshHead.ScrollBars = flexScrollBarHorizontal
        mshBody.MergeCells = flexMergeRestrictRows
        mshBody.ScrollBars = flexScrollBarBoth

        mshHead.Row = mshHead.FixedRows
        mshBody.Row = 0: mshBody.Col = 0
        mshBody.Redraw = True
        mshHead.Redraw = True
         '????????(????????????)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.ID)
            '??????????????????????
            If tmpItem.Relations.count > 0 Then
                For i = 0 To mshBody.Rows - 1
                    '????????????
                    If TypeName(mshBody.Cell(flexcpData, i, tmpItem.????)) <> "Empty" Then
                        If mshBody.Cell(flexcpForeColor, i, tmpItem.????) = 0 Then
                            mshBody.Cell(flexcpForeColor, i, tmpItem.????) = &HFF0001
                        End If
                        mshBody.Cell(flexcpFontUnderline, i, tmpItem.????) = True
                        On Error Resume Next
                        mshBody.Cell(flexcpData, i, tmpItem.????).Remove 2
                        On Error GoTo 0
                        mshBody.Cell(flexcpData, i, tmpItem.????).Add tmpItem.Relations
                    End If
                Next
                
            End If
            '??????????????????????
            blnTmp = False
            If tmpItem.???? = 1 Then
                For i = mshBody.FixedRows To mshBody.Rows - 1
                    If mshBody.TextMatrix(i, tmpItem.????) <> "" Then
                        blnTmp = True: Exit For
                    End If
                Next
                If blnTmp = False Then
                    mshBody.ColHidden(tmpItem.????) = True
                    mshHead.ColHidden(tmpItem.????) = True
                    mshBody.ColWidth(tmpItem.????) = 0
                    mshHead.ColWidth(tmpItem.????) = 0
                End If
            End If
        Next
        
        If bytKind <> 2 Then '????????
            mshHead.ZOrder
            mshHead.Visible = True
        End If
        If bytKind <> 1 Then '????????
            mshBody.ZOrder
            mshBody.Visible = True
        End If
    End With
    
    '????????????????
    If UBound(arrRowIDs) + 1 < mshBody.Rows Then
        ReDim Preserve arrRowIDs(UBound(arrRowIDs) + (mshBody.Rows - (UBound(arrRowIDs) + 1)))
    End If
    mcolRowIDs.Add arrRowIDs, "_" & mshBody.Index
End Sub

Private Function CheckColProtertys(ByVal var???????? As Variant, ByVal str???????? As String, ByVal var?????? As Variant) As Boolean
'??????????????????????????????????????
    
    Select Case str????????
        Case ""
            CheckColProtertys = True
        Case "????"
            If IsNumeric(var????????) Then var?????? = ValEx(var??????): var???????? = ValEx(var????????)
            CheckColProtertys = (var???????? = var??????)
        Case "????"
            var?????? = ValEx(var??????)
            var???????? = ValEx(var????????)
            CheckColProtertys = (var???????? > var??????)
        Case "????"
            var?????? = ValEx(var??????)
            var???????? = ValEx(var????????)
            CheckColProtertys = (var???????? < var??????)
        Case "??????"
            var?????? = Val(var??????)
            var???????? = Val(var????????)
            CheckColProtertys = (var???????? <> var??????)
        Case "????????"
            var?????? = ValEx(var??????)
            var???????? = ValEx(var????????)
            CheckColProtertys = (var???????? >= var??????)
        Case "????????"
            var?????? = ValEx(var??????)
            var???????? = ValEx(var????????)
            CheckColProtertys = (var???????? <= var??????)
        Case "??????"
            If var???????? <> "" And var?????? <> "" Then
                CheckColProtertys = (var???????? Like var?????? & "*")
            End If
        Case "????????"
            If var???????? <> "" And var?????? <> "" Then
                CheckColProtertys = (var???????? Like "*" & var?????? & "*")
            End If
    End Select
End Function

Private Sub ShowItems()
    Dim i As Integer, lngW As Long, lngH As Long
    Dim objItem As RPTItem, objLoad As Object
    Dim strData As String, strTmp As String
    Dim strValue As String, objPic As StdPicture
    Dim objFmt As RPTFmt, objFont As StdFont
    Dim lngSize As Long, sngWidth As Single
    Dim lngRec As Long
    
    On Error GoTo errH
    blnRefresh = False
    If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "????????????????,????????????", , Me

    LockWindowUpdate Me.hwnd
    
    Set mcolRowIDs = New Collection
    For Each objLoad In msh
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In lbl
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In img
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In imgCode
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In lin
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In Shp
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In Chart
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In pic
        If objLoad.Index <> 0 And objLoad.Container Is picPaper(intReport) Then Unload objLoad
    Next
    
    picPaper(intReport).Cls
    intGridCount = 0
    intGridID = 0
    Set objCurGrid = Nothing
    
    Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)
    If objFmt.???? = 1 Then
        lngW = objFmt.W
        lngH = objFmt.H
    Else
        lngW = objFmt.H
        lngH = objFmt.W
    End If
    '????????
    For Each objItem In mobjReport.Items
        '??????2??????????????,??????????????????????????
        If objItem.?????? = bytFormat Then
            With objItem
                If .???? = 14 Then
                    Load pic(.ID)
                    Set pic(.ID).Container = picPaper(intReport)
                    Set objLoad = pic(.ID)
                    .???? = "111"
                    objLoad.Left = .X
                    objLoad.Top = .Y
                    
                    objLoad.Height = IIF(.H > lngH, lngH, .H)
                    objLoad.Width = IIF(.W > lngW, lngW, .W)
                    objLoad.BorderStyle = IIF(.????, 1, 0)
                    
                    objLoad.ZOrder
                    objLoad.Visible = True
                End If
            End With
        End If
    Next
    
    For Each objItem In mobjReport.Items
        '??????2??????????????,??????????????????????????
        If objItem.?????? = bytFormat Then
            With objItem
                Select Case .????
                    Case 1 '????
                        Load lin(.ID)
                        Set lin(.ID).Container = picPaper(intReport)
                        If .??ID <> 0 Then
                            Set lin(.ID).Container = pic(.??ID)
                        End If
                        Set objLoad = lin(.ID)
                        objLoad.X1 = .X
                        objLoad.X2 = IIF(.X + .W - IIF(.W > 0, Screen.TwipsPerPixelX, 0) > lngW, lngW, .X + .W - IIF(.W > 0, Screen.TwipsPerPixelX, 0))
                        objLoad.Y1 = .Y
                        objLoad.Y2 = IIF(.Y + .H - IIF(.H > 0, Screen.TwipsPerPixelY, 0) > lngH, lngH, .Y + .H - IIF(.H > 0, Screen.TwipsPerPixelY, 0))
                        objLoad.BorderColor = .????
                        If .???? Then objLoad.BorderWidth = 2
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 10 '????
                        Load Shp(.ID)
                        Set Shp(.ID).Container = picPaper(intReport)
                        If .??ID <> 0 Then
                            Set Shp(.ID).Container = pic(.??ID)
                        End If
                        Set objLoad = Shp(.ID)
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderColor = 0
                        If .???? Then objLoad.BorderWidth = 2
                        objLoad.Shape = IIF(.????, ShapeConstants.vbShapeOval, ShapeConstants.vbShapeRectangle)
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 11 '????
                        Load img(.ID)
                        Set img(.ID).Container = picPaper(intReport)
                        If .??ID <> 0 Then
                            Set img(.ID).Container = pic(.??ID)
                        End If
                        Set objLoad = img(.ID)
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        
                        Set objPic = LoadPictureFromPar(Me, .????)
                        If objPic Is Nothing Then Set objPic = .????
                        If .???? And Not objPic Is Nothing Then
                            .W = objPic.Width * (15 / 26.46)
                            .H = objPic.Height * (15 / 26.46)
                        End If
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderStyle = IIF(.????, 1, 0)
                        
                        '????????
                        If Not objPic Is Nothing Then
                            If .???? Then
                                Set objLoad.Picture = ScalePicture(picTemp, objPic, objLoad.Width, objLoad.Height)
                            Else
                                Set objLoad.Picture = objPic
                            End If
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 2, 3 '????,????????????
                        strValue = .????
                        
                        '????????????(????????????????????????????)
                        strData = GetLabelDataName(strValue)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = Split(Split(strData, "|")(i), ".")(0)
                                If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                    If Val(.?????? & "") <> 0 Then
                                        If mLibDatas("_" & strTmp).DataSet.RecordCount >= Val(.?????? & "") Then
                                           mLibDatas("_" & strTmp).DataSet.AbsolutePosition = Val(.?????? & "")
                                        Else
                                            mLibDatas("_" & strTmp).DataSet.MoveFirst
                                        End If
                                    Else
                                        mLibDatas("_" & strTmp).DataSet.MoveFirst
                                    End If
                                End If
                            Next
                            lngRec = mLibDatas("_" & strTmp).DataSet.AbsolutePosition
                        End If
                        
                        '??????????????(??????????????????)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = GetFieldValue(Me, CStr(Split(strData, "|")(i)))
                                If .???? <> "" Then
                                    On Error Resume Next
                                    strTmp = Format(strTmp, .????)
                                    If Err.Number <> 0 Then Err.Clear
                                    On Error GoTo errH
                                End If
                                strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                            Next
                        End If
                        
                        
                        '??????????????:[=??????]??[n>=0]??[??????????]??[????????]
                        strValue = GetLabelMacro(Me, strValue)
                        
                        If gobjFile.FileExists(strValue) Then
                            '??????????????????
                            On Error Resume Next
                            Set .???? = LoadPicture(strValue)
                            If .???? Is Nothing Then Set .???? = New StdPicture '??????????????????????
                            Kill strValue
                            Err.Clear
                            On Error GoTo errH
                            
                            If .???? Then
                                .W = .????.Width * (15 / 26.46)
                                .H = .????.Height * (15 / 26.46)
                            End If
                            
                            Load img(.ID)
                            Set img(.ID).Container = picPaper(intReport)
                            Set objLoad = img(.ID)
                            objLoad.BorderStyle = IIF(.????, 1, 0)
                            
                            '????????
                            If .???? Then
                                Set objLoad.Picture = ScalePicture(picTemp, .????, objLoad.Width, objLoad.Height)
                            Else
                                Set objLoad.Picture = .????
                            End If
                        Else
                            Set .???? = Nothing '??????????????????????
                            
                            If .???? Then Call ItemAutoSize(objItem, strValue, picBack)
                            
                            Load lbl(.ID)
                            Set lbl(.ID).Container = picPaper(intReport)
                            If .??ID <> 0 Then
                                Set lbl(.ID).Container = pic(.??ID)
                            End If
                            Set objLoad = lbl(.ID)
                            
                            objLoad.FontName = .????
                            objLoad.FontSize = .????
                            objLoad.FontBold = .????
                            objLoad.FontItalic = .????
                            objLoad.FontUnderline = .????
                            
                            objLoad.Alignment = IIF(.???? = 2, 1, IIF(.???? = 1, 2, 0))
                            objLoad.BorderStyle = IIF(.????, 1, 0)
                            objLoad.ForeColor = .????
                            objLoad.BackColor = .????
                            objLoad.Caption = strValue
                            '??????????????
                            If objItem.Relations.count > 0 Then
                                objLoad.ForeColor = &HFF0001
                                objLoad.FontUnderline = True
                                objLoad.Tag = lngRec
                            End If
                        
                            '????????????????
'                            If .???? = 1 Then
'                                Set objFont = GetAutoFont(objLoad.Caption, IIF(.W > lngW, lngW, .W), IIF(.H > lngH, lngH, .H), objLoad.Font, picTemp, True)
'                                objLoad.Font.Size = objFont.Size
'                            End If
                        End If
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 4 '????????(????????6??????)
                        If objItem.???? = 0 Then
                            If .??ID = 0 Then
                                '??????????????????????
                                intGridCount = intGridCount + 1
                                intGridID = objItem.ID
                            End If
                        End If
                        Call ShowFreeGrid(objItem, lngW, lngH)
                    Case 5 '????????(????????7,8,9??????)
                        If objItem.???? = 0 Then
                            intGridCount = intGridCount + 1
                            intGridID = objItem.ID
                            Call ShowStatGrid(objItem, lngW, lngH)
                        End If
                    Case 12 '????@@@
                        Load Chart(.ID)
                        Set Chart(.ID).Container = picPaper(intReport)
                        If .??ID <> 0 Then
                            Set Chart(.ID).Container = pic(.??ID)
                        End If
                        Set objLoad = Chart(.ID)
                        
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                                                                
                        strTmp = GetChartFileFromPar(Me, .????)
                        If strTmp <> "" Then
                            Call objLoad.Load(strTmp)
                            objLoad.Height = IIF(.H > lngH, lngH, .H)
                            objLoad.Width = IIF(.W > lngW, lngW, .W)
                        Else
                            If objItem.???? <> "" Then
                                Call GetChartDataName(objItem.????, , , , strTmp)
                            End If
                            If strTmp <> "" Then
                                Call SetChartStyleAndData(objLoad, objItem, mLibDatas("_" & strTmp).DataSet)
                            Else
                                Call SetChartStyleAndData(objLoad, objItem, , , , True)
                            End If
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 13 '????
                        Load imgCode(.ID)
                        Set imgCode(.ID).Container = picPaper(intReport)
                        If .??ID <> 0 Then
                            Set imgCode(.ID).Container = pic(.??ID)
                        End If
                        Set objLoad = imgCode(.ID)
                        
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderStyle = 0
                        
                        '????????????
                        strValue = .????
                        
                        '????????????(????????????????????????????)
                        strData = GetLabelDataName(strValue)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = Split(Split(strData, "|")(i), ".")(0)
                                If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                    mLibDatas("_" & strTmp).DataSet.MoveFirst
                                End If
                            Next
                        End If
                        
                        '??????????????(??????????????????)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = GetFieldValue(Me, CStr(Split(strData, "|")(i)))
                                If .???? <> "" Then
                                    On Error Resume Next
                                    strTmp = Format(strTmp, .????)
                                    If Err.Number <> 0 Then Err.Clear
                                    On Error GoTo errH
                                End If
                                strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                            Next
                        End If
                        
                        '??????????????:[=??????]??[n>=0]??[??????????]??[????????]
                        strValue = GetLabelMacro(Me, strValue)
                        '[????]??[????]????????????
                        strValue = Replace(strValue, "[????]", "")
                        strValue = Replace(strValue, "[????]", "")
                        
                        Set objPic = Nothing
                        If strValue <> "" Then
                            Unload frmFlash '????????Picture????????????????????
                            If .???? = 1 Then
                                Set objPic = DrawBarCode128(frmFlash.picTemp, 3, strValue, Mid(.????, 1, 1) = "1")
                            ElseIf .???? = 2 Then
                                Set objPic = DrawBarCode39(frmFlash.picTemp, 3, strValue, Mid(.????, 2, 1) = "1", Mid(.????, 1, 1) = "1")
                            ElseIf .???? = 3 Then
                                Set objPic = DrawBarCode128Auto(frmFlash.picTemp, strValue, sngWidth, .????, Mid(.????, 1, 1) = "1")
                            ElseIf .???? = 10 Then
                                Set objPic = DrawBarCode2D(strValue, frmFlash.picTemp, lngSize)
                            End If
                            If Val(Mid(.????, 3, 1)) <> 0 Then
                                Set objPic = PictureSpin(objPic, Val(Mid(.????, 3, 1)), frmFlash.picTemp)
                            End If
                        End If
                        Set objLoad.Picture = objPic
                        
                        If .???? = 3 Then
                            '128??????????????
                            If Val(Mid(.????, 3, 1)) = 0 Then
                                .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                                objLoad.Width = .W
                            Else
                                .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                                objLoad.Height = .H
                            End If
                        ElseIf .???? = 10 And .???? Then
                            '????????????????????????
                            objLoad.Width = lngSize
                            objLoad.Height = lngSize
                            .W = lngSize: .H = lngSize
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                End Select
            End With
        End If
    Next
    
    '????????????????????????????????????
    For Each objItem In mobjReport.Items
        If objItem.???? = 4 Or objItem.???? = 5 Then
            Call mdlPublic.SetCellValue(Val("0-??????"), Me, objItem)
        End If
    Next
    
    scrVsc.Visible = Not (intGridCount = 1 And Not mobjReport.????)
    scrHsc.Visible = Not (intGridCount = 1 And Not mobjReport.????)
    picShadow.Visible = Not (intGridCount = 1 And Not mobjReport.????)
    
    '??????????????????
    Call SetGridAlign
        
    mobjReport.intGridCount = intGridCount
    mobjReport.intGridID = intGridID
        
    Call Form_Resize
    
    ShowFlash
    blnRefresh = True
    LockWindowUpdate 0
    Exit Sub
errH:
    ShowFlash
    LockWindowUpdate 0
    If ErrCenter() = 1 Then
        If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "????????????????,????????????", , Me
        LockWindowUpdate Me.hwnd
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetFixHeight(objGrid As Object) As Long
'????????????????????????????????
    Dim i As Integer, lngH As Long
    
    For i = 0 To objGrid.FixedRows - 1
        lngH = lngH + objGrid.RowHeight(i)
    Next
    GetFixHeight = lngH
End Function

Private Sub GetGridCurSize(ByVal intID As Integer, ByRef X As Long, ByRef Y As Long, _
    ByRef W As Long, ByRef H As Long, Optional ByRef Bottom As Long)
'????????????????????????????????????(????????????????????????????)
'??????X,Y,W,H??Bottom(??????)
    Dim objItem As RPTItem, tmpItem As RPTItem, lngCurH As Long, lngBottom As Long
    
    X = msh(intID).Left
    W = msh(intID).Width
    
    If Val(msh(intID).Tag) = 0 Then
        Y = msh(intID).Top
        lngCurH = msh(intID).Height
    Else
        Y = msh(CInt(msh(intID).Tag)).Top
        lngCurH = msh(CInt(msh(intID).Tag)).Height
    End If
        
    lngBottom = mobjReport.Items("_" & intID).Y + mobjReport.Items("_" & intID).H
    
    Set objItem = mobjReport.Items("_" & intID)
    W = W * objItem.????
    
    '????????????????????
    For Each tmpItem In mobjReport.Items '??????????????
        If tmpItem.?????? = bytFormat And tmpItem.???? = 4 _
            And tmpItem.???? = 1 And tmpItem.???? = objItem.???? Then
            lngCurH = lngCurH + msh(CInt(msh(tmpItem.ID).Tag)).Height
            lngBottom = lngBottom + tmpItem.H
        End If
    Next
    H = lngCurH
    Bottom = lngBottom
End Sub

Private Function GetDependID(strName As String) As Integer
'??????????????????,??????????.
    Dim objItem As RPTItem
    
    For Each objItem In mobjReport.Items
        If objItem.?????? = bytFormat And objItem.???? = strName _
            And (objItem.???? = 4 Or objItem.???? = 5) And objItem.???? = 0 Then
            GetDependID = objItem.ID: Exit Function
        End If
    Next
End Function

Private Sub SetPlace()
'????????????????????????????????????????????
    Dim objItem As RPTItem, tmpItem As RPTItem
    Dim lngDesignH As Long, lngShowH As Long, lngAppH As Long
    Dim lngCurH As Long, lngCurTop As Long, bytKind As Byte
    Dim strAppGrid As String, lngFixH As Long
    Dim strGridScale As String, i As Integer
    Dim intCurID As Integer, sngScale As Single
    Dim lngTX As Long, lngTY As Long, lngTW As Long, lngTH As Long
    Dim lngBottom As Long
    
    On Error GoTo errH
    
    If mobjReport Is Nothing Then Exit Sub
    If Not mobjReport.blnLoad Then Exit Sub
    
    '????????????????????
    If intGridCount = 1 And Not mobjReport.???? Then
        '??????????????????(????????????????????)??
        '1:??Top??Left????????????
        '2.??Width??????????Left??????????
        '3.??????????????????????????????????,????????????????????????????????
        Set objItem = mobjReport.Items("_" & intGridID)
        
        '??????????????????????(????????????????????????????)
        lngDesignH = objItem.H '????????
        If Val(msh(intGridID).Tag) > 0 Then
            '????????????????????????
            lngShowH = msh(CInt(msh(intGridID).Tag)).Height
        Else
            lngShowH = msh(intGridID).Height
        End If
        
        For Each tmpItem In mobjReport.Items '??????????????
            If tmpItem.?????? = bytFormat And tmpItem.???? = 4 _
                And tmpItem.???? = 1 And tmpItem.???? = objItem.???? Then
                lngDesignH = lngDesignH + tmpItem.H
                lngShowH = lngShowH + msh(CInt(msh(tmpItem.ID).Tag)).Height
                strAppGrid = strAppGrid & "," & tmpItem.ID '??????????
            End If
        Next
        strAppGrid = Mid(strAppGrid, 2)
        
        '??????????????????????
        For Each tmpItem In mobjReport.Items
            If tmpItem.?????? = bytFormat And tmpItem.???? = 2 _
                And tmpItem.???? Is Nothing And tmpItem.Y >= objItem.Y + lngDesignH Then
                If tmpItem.Y + lbl(tmpItem.ID).Height > lngAppH Then
                    lngAppH = tmpItem.Y + lbl(tmpItem.ID).Height '??????????????
                End If
            End If
        Next
        
        If lngAppH > 0 Then lngAppH = lngAppH - (objItem.Y + lngDesignH)
        
        '??????????????????????????
        If lngAppH > lngShowH / 2 Then lngAppH = lngShowH / 2
        
        msh(intGridID).Width = (picPaper(intReport).ScaleWidth - msh(intGridID).Left * 2) / objItem.????

        '????????????????????
        If Val(msh(intGridID).Tag) = 0 Then
            lngCurH = picPaper(intReport).ScaleHeight - msh(intGridID).Top - (lngAppH + 200)
        Else
            msh(CInt(msh(intGridID).Tag)).Width = msh(intGridID).Width
            lngCurH = picPaper(intReport).ScaleHeight - msh(CInt(msh(intGridID).Tag)).Top - (lngAppH + 200)
        End If
        
        If strAppGrid = "" Then '??????????????
            If objItem.???? = 5 Then
                msh(intGridID).Height = lngCurH
            Else
                bytKind = GetGridStyle(mobjReport, intGridID)
                If bytKind = 2 Then
                    msh(intGridID).Height = lngCurH
                Else
                    lngFixH = GetFixHeight(msh(CInt(msh(intGridID).Tag))) '??????????
                    If lngCurH < lngFixH + 300 Then lngCurH = lngFixH + 300 '??????????????????????(????????)
                    msh(CInt(msh(intGridID).Tag)).Height = lngCurH
                    msh(intGridID).Height = lngCurH - lngFixH
                End If
            End If
        Else
            '????????????
            '??????????????????????????????
            strGridScale = "|" & objItem.ID & "," & objItem.H / lngDesignH
            For i = 0 To UBound(Split(strAppGrid, ","))
                Set tmpItem = mobjReport.Items("_" & Split(strAppGrid, ",")(i))
                strGridScale = strGridScale & "|" & tmpItem.ID & "," & tmpItem.H / lngDesignH
            Next
            strGridScale = Mid(strGridScale, 2) '"??ID,????|??ID,????..."
            lngCurTop = objItem.Y
            For i = 0 To UBound(Split(strGridScale, "|"))
                intCurID = CInt(Split(Split(strGridScale, "|")(i), ",")(0))
                sngScale = CSng(Split(Split(strGridScale, "|")(i), ",")(1))
                
                If i > 0 Then
                    msh(intCurID).Width = msh(intGridID).Width
                    msh(CInt(msh(intCurID).Tag)).Width = msh(intCurID).Width
                End If
                
                bytKind = GetGridStyle(mobjReport, intCurID)
                
                If Val(msh(intCurID).Tag) = 0 Then '????????,????????????,??????????????
                    msh(intCurID).Height = lngCurH * sngScale
                    lngCurTop = lngCurTop + msh(intCurID).Height
                Else
                    lngFixH = GetFixHeight(msh(CInt(msh(intCurID).Tag)))
                    If lngCurH * sngScale < lngFixH + 300 Then '??????????????????????(????????)
                        msh(CInt(msh(intCurID).Tag)).Height = lngFixH + 300
                    Else
                        msh(CInt(msh(intCurID).Tag)).Height = lngCurH * sngScale
                    End If
                    msh(CInt(msh(intCurID).Tag)).Top = lngCurTop
                    lngCurTop = lngCurTop + msh(CInt(msh(intCurID).Tag)).Height
                    
                    bytKind = GetGridStyle(mobjReport, intCurID)
                    If bytKind = 2 Then
                        msh(intCurID).Top = msh(CInt(msh(intCurID).Tag)).Top
                        msh(intCurID).Height = msh(CInt(msh(intCurID).Tag)).Height
                    Else
                        msh(intCurID).Top = msh(CInt(msh(intCurID).Tag)).Top + lngFixH
                        msh(intCurID).Height = msh(CInt(msh(intCurID).Tag)).Height - lngFixH
                    End If
                End If
            Next
        End If
    End If
    
    For Each tmpItem In mobjReport.Items
        If tmpItem.?????? = bytFormat And tmpItem.???? = 4 And tmpItem.???? > 1 _
            And tmpItem.???? = 0 And tmpItem.???? = "" Then
            '????????
            For i = 2 To tmpItem.????
                With msh(tmpItem.ID)
                    DrawCell picPaper(intReport), "????????????", tmpItem.X + ((i - 1) * .Width), tmpItem.Y, .Width, _
                        msh(CInt(.Tag)).Height - 15, , , .GridColor, .ForeColor, .BackColor, .Font, , 1, 1
                End With
            Next
        End If
    Next
    
    '??????????????????????????(????????)
    For Each tmpItem In mobjReport.Items
        If tmpItem.?????? = bytFormat And tmpItem.???? = 2 And tmpItem.???? Is Nothing Then
            '??????????????????????????????????
            If tmpItem.???? <> 0 And tmpItem.???? <> "" Then
                GetGridCurSize GetDependID(tmpItem.????), lngTX, lngTY, lngTW, lngTH
                Select Case tmpItem.????
                    Case 11, 21 '????
                        lbl(tmpItem.ID).Left = lngTX
                    Case 12, 22 '????
                        lbl(tmpItem.ID).Left = lngTX + (lngTW - lbl(tmpItem.ID).Width) / 2
                    Case 13, 23 '????
                        lbl(tmpItem.ID).Left = lngTX + lngTW - lbl(tmpItem.ID).Width
                End Select
            End If
            '??????????????????????????????????????(????????,??????????)
            If intGridCount = 1 Then
                GetGridCurSize intGridID, lngTX, lngTY, lngTW, lngTH, lngBottom
                If tmpItem.Y >= lngBottom Then
                    lbl(tmpItem.ID).Top = lngTY + lngTH + (tmpItem.Y - lngBottom)
                End If
            End If
        End If
    Next
    Exit Sub
errH:
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GridHaveApp(intID As Integer) As Boolean
'??????????????????????????????????
    Dim tmpItem As RPTItem, strName As String
    
    strName = mobjReport.Items("_" & intID).????
    For Each tmpItem In mobjReport.Items
        If tmpItem.?????? = bytFormat And tmpItem.???? = 4 And tmpItem.???? = 1 And tmpItem.???? = strName Then
            GridHaveApp = True: Exit Function
        End If
    Next
End Function

Private Function GetGridDesignWidth(objItem As RPTItem) As Long
'??????????????????????????????????????????????
    Dim lngW As Long, tmpItem As RPTItem
    
    lngW = objItem.W
    For Each tmpItem In mobjReport.Items
        If tmpItem.?????? = bytFormat And tmpItem.???? = 5 _
            And tmpItem.???? = 2 And tmpItem.???? = objItem.???? Then
            lngW = lngW + tmpItem.W
        End If
    Next
    GetGridDesignWidth = lngW
End Function

Private Function GetPreAppGrid(intID As Integer, arrGrids As Variant) As Long
'??????????????????????????????????(????????????????????????)
'??????arrGrids=??XY??????????????????????????
'??????1.????????????????????????????????Y????????????
'      2.??????????????????????????,??????????????????????
    Dim objItem As RPTItem, tmpItem As RPTItem, i As Integer
    
    Set objItem = mobjReport.Items("_" & intID)
    For i = 0 To UBound(arrGrids)
        Set tmpItem = mobjReport.Items("_" & arrGrids(i))
        If tmpItem.?????? = bytFormat And _
            ((tmpItem.???? = 4 And tmpItem.???? = 1 And tmpItem.???? = objItem.????) Or _
            (tmpItem.???? = 0 And objItem.???? = tmpItem.????)) Then
            If tmpItem.ID <> intID Then
                GetPreAppGrid = tmpItem.ID
            Else '??????????????????????,????????????????
                Exit Function
            End If
        End If
    Next
End Function

Private Function GetGridDesignHeight(intID As Integer) As Long
'??????????????????????????(????????????????)
'????????????????????????????????????
    Dim objItem As RPTItem, tmpItem As RPTItem
    Dim lngH As Long
    
    Set objItem = mobjReport.Items("_" & intID)
    If objItem.???? = 1 And objItem.???? <> "" Then
        Set objItem = mobjReport.Items("_" & GetDependID(objItem.????))
    End If
    
    lngH = objItem.H
    For Each tmpItem In mobjReport.Items
        If tmpItem.?????? = bytFormat And tmpItem.???? = 4 _
            And tmpItem.???? = 1 And tmpItem.???? = objItem.???? Then
            lngH = lngH + tmpItem.H
        End If
    Next
    GetGridDesignHeight = lngH
End Function

Private Function GetGridPageCol(objItem As RPTItem) As Integer
'??????????????????????????????????????
'??????objItem=????????
'??????-1=????
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    GetGridPageCol = -1
    If objItem.???? <> 4 Then Exit Function
    
    For Each tmpID In objItem.SubIDs
        Set tmpItem = mobjReport.Items("_" & tmpID.ID)
        If tmpItem.???? Then
            GetGridPageCol = tmpItem.????
            Exit For
        End If
    Next
End Function

Private Function CalcCellPage() As Boolean
'??????????????????????????????
'??????mobjreport=????????
'      marrPage=????????
'????????????????????????????(??????????????????????????????)
'????????????????????????isArray(marrPage)=False,??????????????????
    Dim objBody As Control, objPageCell As PageCell, arrPage As Variant '????????????????
    Dim lngFixW As Long, lngFixH As Long '????????????????????
    Dim lngRowB As Long, lngRowE As Long
    Dim lngColB As Long, lngColE As Long '????????
    Dim lngBodyW As Long, lngBodyH As Long '??????????????????????
    Dim lngCurW As Long, lngCurH As Long '????????????????????????????
    Dim lngOutX As Long, lngOutY As Long '??????????????????????????(????????????????)
    Dim bytKind As Byte, intPage As Integer  '??????????????(0-N)
    Dim i As Long, j As Long, k As Long, strTmp As String
    Dim objItem As RPTItem, blnHaveApp As Boolean, blnHorPage As Boolean
    Dim blnApp As Boolean, lngMinH As Long, arrGrids As Variant
    Dim lngPreID As Long, intDepend As Integer, lngDesignH As Long
    Dim tmpPageCell As PageCell
    Dim lngL As Long, lngW As Long, lngC As Long, lngZ As Long
    Dim lngTop As Long, lngLeft As Long
    Dim lngCount As Long
    Dim blnData As Boolean, tmpSubID As RelatID
    Dim Y As Long, X As Long, Z As Long
    
    '??????????????????????????
    Dim strCurText As String, blnNewPage As Boolean, lngPageCol As Long
    Dim lngBaseRows As Long, lngVRowE As Long
    Dim colCardRow As New Collection  '??????????????????????????
    Dim lngLastID As Long, lngRow As Long
    Dim lngRowCount As Long, colCard As New Collection
    Dim blnRePage As Boolean, blnPage As Boolean
    Dim arrPageTmp As Variant, arrTmp As Variant
    
    '????????X,Y????????????
    arrGrids = Array()
    For Each objBody In msh
        If objBody.Index <> 0 And (objBody.Container Is picPaper(intReport) Or objBody.Container.name = "pic") And Left(objBody.Tag, 2) <> "H_" Then
            ReDim Preserve arrGrids(UBound(arrGrids) + 1)
            arrGrids(UBound(arrGrids)) = objBody.Left & "," & objBody.Top & "," & objBody.Index
        End If
    Next
    For i = 0 To UBound(arrGrids) - 1
        For j = i To UBound(arrGrids)
            If CLng(Split(arrGrids(j), ",")(0)) < CLng(Split(arrGrids(i), ",")(0)) Then
                strTmp = arrGrids(i): arrGrids(i) = arrGrids(j): arrGrids(j) = strTmp
            End If
        Next
    Next
    For i = 0 To UBound(arrGrids) - 1
        For j = i To UBound(arrGrids)
            If CLng(Split(arrGrids(j), ",")(1)) < CLng(Split(arrGrids(i), ",")(1)) Then
                strTmp = arrGrids(i): arrGrids(i) = arrGrids(j): arrGrids(j) = strTmp
            End If
        Next
    Next
    For i = 0 To UBound(arrGrids)
        arrGrids(i) = CInt(Split(arrGrids(i), ",")(2))
    Next
    
    arrPage = Empty
    marrPage = Empty
    marrPageCard = Empty
    
    For k = 0 To UBound(arrGrids)
        '????????????
        Set objBody = msh(arrGrids(k))
        strTmp = ""
        lngLeft = 0: lngTop = 0
        If objBody.Container.name = "pic" Then
            If objBody.Container.Container Is picPaper(intReport) Then
                lngLeft = mobjReport.Items("_" & objBody.Container.Index).X
                lngTop = mobjReport.Items("_" & objBody.Container.Index).Y
            End If
        End If
        Set objItem = mobjReport.Items("_" & objBody.Index)
        blnApp = (objItem.???? = 4 And objItem.???? = 1 And objItem.???? <> "") '????????????
        
        '??????????????????????????
        lngFixW = 0: lngFixH = 0
        lngDesignH = GetGridDesignHeight(objItem.ID) '??????????????????
        
        If objItem.???? = 5 Then
            For i = 0 To objBody.FixedCols - 1
                lngFixW = lngFixW + objBody.ColWidth(i)
            Next
            For i = 0 To objBody.FixedRows - 1
                lngFixH = lngFixH + objBody.RowHeight(i)
            Next
            '????????????????????????????????????(????????)
            lngBodyW = GetGridDesignWidth(objItem) - lngFixW
            lngBodyH = lngDesignH - lngFixH
        Else
            bytKind = GetGridStyle(mobjReport, objBody.Index)
            For i = 0 To msh(CInt(objBody.Tag)).FixedRows - 1
                lngFixH = lngFixH + msh(objBody.Tag).RowHeight(i)
            Next
            Select Case bytKind
                Case 0
                    lngBodyH = lngDesignH - lngFixH
                Case 1
                    lngBodyH = 0
                Case 2
                    lngBodyH = lngDesignH
                    lngFixH = 0
            End Select
            lngBodyW = objItem.W
        End If
        
        If objItem.???? = 4 Then blnHaveApp = GridHaveApp(objItem.ID)
        
        lngPageCol = GetGridPageCol(objItem) '??????????,??????-1
        lngRowB = objBody.FixedRows
        lngColB = objBody.FixedCols
        lngRowE = lngRowB - 1
        lngColE = lngColB - 1
        
        '????????????????????
        If blnApp Then
            '??????????ID
            intDepend = GetDependID(objItem.????)
            '????????????????????ID
            '????????????????????,????????????????????????
            lngPreID = GetPreAppGrid(objItem.ID, arrGrids)
            intPage = -1
            If objItem.??ID <> 0 Then
                arrTmp = arrPageTmp
            Else
                arrTmp = arrPage
            End If
            For i = 0 To UBound(arrTmp)
                For Each objPageCell In arrTmp(i)
                    If objPageCell.ID = lngPreID Then
                        '????????(????????)??????????????????????????
                        '(??????????????????????,????????????????????????)
                        If objPageCell.RowE >= msh(objPageCell.ID).Rows - 1 _
                            And objPageCell.ColB = msh(objPageCell.ID).FixedCols Then
                            '??????????????????????(????????????????????)
                            Select Case bytKind
                                Case 0 '????????
                                    lngMinH = lngFixH + objItem.????
                                Case 1 '??????????
                                    lngMinH = lngFixH
                                Case 2 '??????????
                                    lngMinH = objItem.????
                            End Select
                            If lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y) >= lngMinH Then
                                lngOutX = objPageCell.X + lngLeft
                                lngOutY = objPageCell.Y + objPageCell.H + lngTop
                                Select Case bytKind
                                    Case 0 '????????
                                        lngBodyH = lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y) - lngFixH
                                    Case 1 '??????????
                                        lngBodyH = 0
                                    Case 2 '??????????
                                        lngBodyH = lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y)
                                End Select
                                intPage = i
                            Else
                                '????????????????,????????????????
                                lngOutX = mobjReport.Items("_" & intDepend).X + lngLeft
                                lngOutY = mobjReport.Items("_" & intDepend).Y + lngTop
                                Select Case bytKind
                                    Case 0 '????????
                                        lngBodyH = lngDesignH - lngFixH
                                    Case 1 '??????????
                                        lngBodyH = 0
                                    Case 2 '??????????
                                        lngBodyH = lngDesignH
                                End Select
                                intPage = i + 1
                                
                                '????????????????????????????????
                                For j = intPage To UBound(arrTmp)
                                    For Each tmpPageCell In arrTmp(j)
                                        If tmpPageCell.ID = intDepend Then
                                            If tmpPageCell.ColB <> msh(intDepend).FixedCols Then
                                                intPage = intPage + 1
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                            Exit For
                        End If
                    End If
                Next
                If intPage <> -1 Then Exit For
            Next
            If intPage = -1 Then intPage = 0
        Else
            lngOutX = objItem.X + lngLeft
            lngOutY = objItem.Y + lngTop
            intPage = 0
        End If
        
        '????????(????????????????????)
        lngBaseRows = Int(lngBodyH / objItem.????) '??????????????????????????????
        Do
            '????????(????DO)
            
            '????????????????
            lngCurH = 0
            blnNewPage = False
            Do
                If lngPageCol <> -1 Then
                    If lngRowE + 1 = lngRowB Then
                        '????????????lngRowE=lngRowB-1,????????,??????????????????????
                        strCurText = objBody.TextMatrix(lngRowE + 1, lngPageCol)
                    ElseIf lngRowE + 1 > lngRowB Then
                        If strCurText <> objBody.TextMatrix(lngRowE + 1, lngPageCol) Then
                            blnNewPage = True
                        End If
                    End If
                End If
                If Not blnNewPage Then
                    lngCurH = lngCurH + objBody.RowHeight(lngRowE + 1)
                    If lngCurH <= lngBodyH Then
                        lngRowE = lngRowE + 1
                        If lngPageCol <> -1 Then
                            strCurText = objBody.TextMatrix(lngRowE, lngPageCol)
                        End If
                    End If
                End If
            Loop Until (lngCurH > lngBodyH) Or (lngRowE = objBody.Rows - 1) Or blnNewPage
            
            '??????????
            If lngCurH > lngBodyH Then lngCurH = lngCurH - objBody.RowHeight(lngRowE + 1)
            
            '????????????,??????????????
            If lngRowE < lngRowB Then lngRowE = lngRowB
            
            '????????????????????????????????????????????
            lngVRowE = 0 '??????????????????????????????
            If objItem.???? > 1 Then
                '????????????????(??????????????????????)
                If lngPageCol <> -1 Then
                    strCurText = objBody.TextMatrix(lngRowE, lngPageCol)
                    For i = lngRowE + 1 To objBody.Rows - 1
                        If i - lngRowB + 1 > lngBaseRows * objItem.???? Then
                            lngRowE = i - 1: Exit For
                        ElseIf strCurText <> objBody.TextMatrix(i, lngPageCol) Then
                            lngRowE = i - 1: Exit For
                        Else
                            lngRowE = i
                        End If
                        strCurText = objBody.TextMatrix(i, lngPageCol)
                    Next
                Else
                    For i = lngRowE + 1 To objBody.Rows - 1
                        If i - lngRowB + 1 > lngBaseRows * objItem.???? Then
                            lngRowE = i - 1: Exit For
                        Else
                            lngRowE = i
                        End If
                    Next
                End If
                '????????????????(??????????????????)
                If mobjReport.???? Then
                    '????????????????????????????
                    lngVRowE = lngRowE + (lngBaseRows * objItem.???? - (lngRowE - lngRowB + 1))
                Else
                    '????????????????????????????????????????????
                    If lngRowE - lngRowB + 1 <= lngBaseRows Then
                        lngVRowE = lngRowE + (lngRowE - lngRowB + 1) * (objItem.???? - 1)
                    Else
                        lngVRowE = lngRowE + (lngBaseRows * objItem.???? - (lngRowE - lngRowB + 1))
                    End If
                End If
            Else
                '??????????????????????????
                '??????????????????????????
                If mobjReport.???? Then
                    lngVRowE = lngRowE + (lngBaseRows - (lngRowE - lngRowB + 1))
                End If
            End If
            If lngVRowE = lngRowE Then lngVRowE = 0
            
            '??????????(??????????????)
            Do
                '????????????????
                lngCurW = 0
                Do
                    lngCurW = lngCurW + objBody.ColWidth(lngColE + 1)
                    If lngCurW <= lngBodyW Then lngColE = lngColE + 1
                Loop Until lngCurW > lngBodyW Or lngColE = objBody.Cols - 1
                
                '??????????
                If lngCurW > lngBodyW Then lngCurW = lngCurW - objBody.ColWidth(lngColE + 1)
                
                '????????????,??????????????
                If lngColE < lngColB Then lngColE = lngColB
                
                If objItem.??ID = 0 Then
                    '??????????????????
                    blnPage = True
                Else
                    '??????????????????????????????????
                    If mobjReport.Items("_" & objItem.??ID).?????? = "" Then
                        blnPage = True
                    Else
                        blnPage = False
                    End If
                End If
                
                '????????????
                If blnPage Then
                    If Not IsArray(arrPage) Then
                        ReDim arrPage(intPage) As PageCells  '????????????
                        Set arrPage(intPage) = New PageCells
                    ElseIf intPage > UBound(arrPage) Then
                        '????????????????????????,????????????
                        ReDim Preserve arrPage(intPage) As PageCells
                        Set arrPage(intPage) = New PageCells
                    End If
                Else
                    If intPage = 0 Then
                        If Not IsArray(arrPageTmp) Then
                            ReDim arrPageTmp(intPage) As PageCells  '????????????
                            Set arrPageTmp(intPage) = New PageCells
                        ElseIf intPage > UBound(arrPageTmp) Then
                            '????????????????????????,????????????
                            ReDim Preserve arrPageTmp(intPage) As PageCells
                            Set arrPageTmp(intPage) = New PageCells
                        End If
                    End If
                End If
                blnData = False
                If objBody.Container.name = "pic" Then
                    If objBody.Container.Container Is picPaper(intReport) Then
                        If mobjReport.Items("_" & objBody.Index).SubIDs.count <> 0 And mobjReport.Items("_" & objBody.Container.Index).?????? <> "" And lngLastID <> objBody.Index Then
                            For Each tmpSubID In mobjReport.Items("_" & objBody.Index).SubIDs
                                If mobjReport.Items("_" & tmpSubID.ID).???? <> "" Then
                                    With mobjReport.Items("_" & tmpSubID.ID)
                                        X = InStr(1, .????, "]")
                                        Y = InStr(1, .????, ".")
                                        Z = InStr(1, .????, "[")
                                        If X > Z And X > Y And X <> 0 And Z <> 0 Then
                                            If Mid(.????, Z + 1, Y - Z - 1) = mobjReport.Items("_" & objBody.Container.Index).?????? Then
                                                blnData = True
                                                Exit For
                                            End If
                                        End If
                                    End With
                                End If
                            Next
                            If blnData Then
                                On Error Resume Next
                                If lngCurH \ mobjReport.Items("_" & objBody.Index).???? < colCardRow("_" & objBody.Container.Index) Then
                                    If Err.Number = 0 Then colCardRow.Remove "_" & objBody.Container.Index
                                    colCardRow.Add lngCurH \ mobjReport.Items("_" & objBody.Index).????, "_" & objBody.Container.Index
                                End If
                                Err.Clear: On Error GoTo 0
                            End If
                        End If
                    End If
                End If
                lngLastID = objBody.Index
                
                '??????????????????
                '??????????????????????
                If blnPage Then
                    arrPage(intPage).Add objBody.Index, lngOutX, lngOutY, lngCurW + lngFixW, lngCurH + lngFixH, _
                        lngDesignH, lngRowB, lngRowE, lngVRowE, lngColB, lngColE, _
                        lngFixW, lngFixH, objItem.????, "_" & objBody.Index
                Else
                    If intPage = 0 Then
                        arrPageTmp(intPage).Add objBody.Index, lngOutX, lngOutY, lngCurW + lngFixW, lngCurH + lngFixH, _
                            lngDesignH, lngRowB, lngRowE, lngVRowE, lngColB, lngColE, _
                            lngFixW, lngFixH, objItem.????, "_" & objBody.Index
                    End If
                End If
                lngColB = lngColE + 1
                lngColE = lngColB - 1
                
                intPage = intPage + 1
            
                If blnApp Then
                    '????????????????????????????????
                    If objItem.??ID <> 0 Then
                        arrTmp = arrPageTmp
                    Else
                        arrTmp = arrPage
                    End If
                    For i = intPage To UBound(arrTmp)
                        For Each objPageCell In arrTmp(i)
                            If objPageCell.ID = intDepend Then
                                If objPageCell.ColB <> msh(intDepend).FixedCols Then
                                    intPage = intPage + 1
                                End If
                            End If
                        Next
                    Next
                    '??????????????????????
                    '????????????????,????????????????
                    lngOutX = mobjReport.Items("_" & intDepend).X
                    lngOutY = mobjReport.Items("_" & intDepend).Y
                    Select Case bytKind
                        Case 0 '????????
                            lngBodyH = lngDesignH - lngFixH
                        Case 1 '??????????
                            lngBodyH = 0
                        Case 2 '??????????
                            lngBodyH = lngDesignH
                    End Select
                End If
            
            '????????????????????????????????????????
            Loop Until lngColB > objBody.Cols - 1 Or _
                (objItem.???? = 4 And (objItem.???? > 1 Or objItem.???? = 1 Or blnHaveApp))
            
            lngColB = objBody.FixedCols
            lngColE = lngColB - 1
                            
            lngRowB = lngRowE + 1
            lngRowE = lngRowB - 1
        '??????????????,????????????????????????????????????
        Loop Until lngRowB > objBody.Rows - 1 Or (objItem.???? = 4 And bytKind = 1)
    Next
    
    '????????????
    For Each objBody In pic
        '??????????????
        If objBody.Index <> 0 And Not mobjReport.Items("_" & objBody.Index) Is Nothing Then
            If mobjReport.Items("_" & objBody.Index).?????? <> "" Then
                lngRowB = 0
                lngRowE = 0
                intPage = 0
                lngCount = 0
                If mobjReport.Items("_" & objBody.Index).???????? = 0 Then
                    If mobjReport.Fmts.Item("_" & bytFormat).???? = 1 Then
                        lngL = (mobjReport.Fmts.Item("_" & bytFormat).W - objBody.Left + mobjReport.Items("_" & objBody.Index).????????) \ (objBody.Width + mobjReport.Items("_" & objBody.Index).????????)
                    Else
                        lngL = (mobjReport.Fmts.Item("_" & bytFormat).H - objBody.Left + mobjReport.Items("_" & objBody.Index).????????) \ (objBody.Width + mobjReport.Items("_" & objBody.Index).????????)
                    End If
                Else
                    lngL = mobjReport.Items("_" & objBody.Index).????????
                End If
                If mobjReport.Items("_" & objBody.Index).???????? = 0 Then
                    If mobjReport.Fmts.Item("_" & bytFormat).???? = 1 Then
                        lngW = (mobjReport.Fmts.Item("_" & bytFormat).H - objBody.Top + mobjReport.Items("_" & objBody.Index).????????) \ (objBody.Height + mobjReport.Items("_" & objBody.Index).????????)
                    Else
                        lngW = (mobjReport.Fmts.Item("_" & bytFormat).W - objBody.Top + mobjReport.Items("_" & objBody.Index).????????) \ (objBody.Height + mobjReport.Items("_" & objBody.Index).????????)
                    End If
                Else
                    lngW = mobjReport.Items("_" & objBody.Index).????????
                End If
                '????????????????????
                lngC = lngW * lngL
                With mLibDatas("_" & mobjReport.Items("_" & objBody.Index).??????).DataSet
                    If .RecordCount > 0 Then .MoveFirst
                    On Error Resume Next
                    If !???????? & "" <> "" Or !???????? & "" = "" Then
                        If Err.Number = 0 Then
                            '????????????????????????????????????????????????????????????????
                            If .RecordCount > 0 Then
                                strTmp = "????????????"
                                lngRow = 0
                                lngRowCount = 0
                                On Error Resume Next
                                '??????????????????????
                                lngRowCount = colCardRow("_" & objBody.Index)
                                If lngRowCount = 0 Then lngRowCount = .RecordCount
                                Err.Clear: On Error GoTo 0
                                For i = 1 To .RecordCount
                                    If strTmp <> !???????? & "" Then
                                        If strTmp <> "????????????" Then
                                            colCard.Add lngRow & "-" & .AbsolutePosition - 1
                                        End If
                                        lngRow = .AbsolutePosition
                                    Else
                                        '??????????????
                                        If .AbsolutePosition + 1 - lngRow > lngRowCount Then
                                            colCard.Add lngRow & "-" & .AbsolutePosition - 1
                                            lngRow = .AbsolutePosition
                                        End If
                                    End If
rePage:
                                    If colCard.count >= lngC Or blnRePage Then
                                        If Not IsArray(marrPageCard) Then
                                            ReDim marrPageCard(intPage) As PageCards  '????????????
                                            Set marrPageCard(intPage) = New PageCards
                                        ElseIf intPage > UBound(marrPageCard) Then
                                            '????????????????????????,????????????
                                            ReDim Preserve marrPageCard(intPage) As PageCards
                                            Set marrPageCard(intPage) = New PageCards
                                        End If
                                        For j = 1 To colCard.count
                                            If j = 1 Then lngRowB = Val(Mid(colCard(j), 1, InStr(colCard(j), "-") - 1))
                                            If j = colCard.count Then lngRowE = Val(Mid(colCard(j), InStr(colCard(j), "-") + 1, Len(colCard(j))))
                                        Next
                                        marrPageCard(intPage).Add objBody.Index, objBody.Left, objBody.Top, objBody.Width, objBody.Height, lngRowB, lngRowE, lngW, lngL, colCard, _
                                                            "_" & objBody.Index
                                        Set colCard = New Collection
                                        intPage = intPage + 1
                                        If blnRePage = True Then blnRePage = False: Exit For
                                    End If
                                    If i = .RecordCount Then
                                        colCard.Add lngRow & "-" & .AbsolutePosition
                                        blnRePage = True
                                        GoTo rePage
                                    End If
                                    
                                    strTmp = !???????? & ""
                                    .MoveNext
                                Next
                                .MoveFirst
                            End If
                        Else
                            '????????????????????
                            lngZ = .RecordCount
                            Err.Clear: On Error GoTo 0
                            Set colCard = New Collection
                            If lngZ > 0 Then
                                Do While lngZ > 0
                                    '????????????
                                    If Not IsArray(marrPageCard) Then
                                        ReDim marrPageCard(intPage) As PageCards  '????????????
                                        Set marrPageCard(intPage) = New PageCards
                                    ElseIf intPage > UBound(marrPageCard) Then
                                        '????????????????????????,????????????
                                        ReDim Preserve marrPageCard(intPage) As PageCards
                                        Set marrPageCard(intPage) = New PageCards
                                    End If
                                    If lngZ <= lngC Then
                                        If lngRowE <> 0 Then
                                            lngRowB = lngRowE + 1
                                        Else
                                            lngRowB = 0
                                        End If
                                        If lngRowE <> 0 Then
                                            lngRowE = lngRowE + lngZ
                                        Else
                                            lngRowE = lngZ - 1
                                        End If
                                        For i = lngRowB To lngRowE
                                            colCard.Add i + 1 & "-" & i + 1
                                        Next
                                        '??????????????????
                                        marrPageCard(intPage).Add objBody.Index, objBody.Left, objBody.Top, objBody.Width, objBody.Height, lngRowB, lngRowE, lngW, lngL, colCard, _
                                             "_" & objBody.Index

                                        lngZ = 0
                                    Else
                                        If lngRowE <> 0 Then
                                            lngRowB = lngRowE + 1
                                        Else
                                            '??????????????????0??????????????????
                                            If lngC = 1 And UBound(marrPageCard) = 1 Then
                                                lngRowB = 1
                                            Else
                                                lngRowB = 0
                                            End If
                                        End If
                                        If lngRowE <> 0 Then
                                            lngRowE = lngRowE + lngC
                                        Else
                                            If lngC = 1 And UBound(marrPageCard) = 1 Then
                                                lngRowE = 1
                                            Else
                                                lngRowE = lngC - 1
                                            End If
                                        End If
                                        For i = lngRowB To lngRowE
                                            colCard.Add i + 1 & "-" & i + 1
                                        Next
                                        
                                       '??????????????????
                                        
                                        marrPageCard(intPage).Add objBody.Index, objBody.Left, objBody.Top, objBody.Width, objBody.Height, lngRowB, lngRowE, lngW, lngL, colCard, _
                                             "_" & objBody.Index
                                             
                                        lngZ = lngZ - lngC
                                    End If
                                    Set colCard = New Collection
                                    intPage = intPage + 1
                                Loop
                            End If
                        End If
                    End If
                End With
            End If
'makNothing:
        End If
    Next
    
    '??????????????????????
    If IsArray(arrPageTmp) Then
        If arrPageTmp(0).count > 0 Then
            For Each objPageCell In arrPageTmp(0)
                If IsArray(marrPageCard) Then
                    For i = 0 To UBound(marrPageCard)
                        On Error Resume Next
                        If marrPageCard(i).Item("_" & mobjReport.Items("_" & objPageCell.ID).??ID).ID <> 0 Then
                            If Err.Number = 0 Then
                                With marrPageCard(i).Item("_" & mobjReport.Items("_" & objPageCell.ID).??ID)
                                        For j = 1 To .Item.count
                                            If Not IsArray(arrPage) Then
                                                ReDim arrPage(i) As PageCells  '????????????
                                                Set arrPage(i) = New PageCells
                                            ElseIf i > UBound(arrPage) Then
                                                '????????????????????????,????????????
                                                ReDim Preserve arrPage(i) As PageCells
                                                Set arrPage(i) = New PageCells
                                            End If
                                            If mobjReport.???? Then
                                                lngVRowE = (mobjReport.Items("_" & objPageCell.ID).H - objPageCell.FixH) \ mobjReport.Items("_" & objPageCell.ID).???? _
                                                        - (Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) + 1) _
                                                        + Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - 1
                                            Else
                                                lngVRowE = objPageCell.VRowE
                                            End If
                                            arrPage(i).Add objPageCell.ID, objPageCell.X + ((j - 1) Mod .Col) * (mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.ID).??ID).W + mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.ID).??ID).????????), _
                                                objPageCell.Y + ((j - 1) \ .Col) * (mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.ID).??ID).H + mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.ID).??ID).????????), _
                                                objPageCell.W, objPageCell.FixH + (Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) + 1) * mobjReport.Items("_" & objPageCell.ID).????, _
                                                objPageCell.MaxH, Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) - 1, Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - 1, lngVRowE, objPageCell.ColB, objPageCell.ColE, _
                                                objPageCell.FixW, objPageCell.FixH, objPageCell.Copys, "_" & objPageCell.ID + (j - 1)
                                        Next
                                    
                                End With
                            Else
                                Err.Clear: On Error GoTo 0
                            End If
                        End If
                    Next
                End If
            Next
        End If
    End If
    
    marrPage = arrPage
    CalcCellPage = True
End Function

Private Sub SetReportIndex(intIndex As Integer, objReport As Report)
'??????????????????????????,??????????????????????
'??????????????????????????????????????????????
'????????????????????????????????????
    Dim tmpItem As RPTItem, objItems As RPTItems
    Dim tmpSubID As RelatID, objSubIDs As RelatIDs
    Dim tmpCopyID As RelatID, objCopyIDs As RelatIDs
    
    Set objItems = New RPTItems
    For Each tmpItem In objReport.Items
        With tmpItem
            Set objSubIDs = New RelatIDs
            For Each tmpSubID In .SubIDs
                objSubIDs.Add tmpSubID.ID & intIndex, "_" & tmpSubID.ID & intIndex
            Next
            Set objCopyIDs = New RelatIDs
            For Each tmpCopyID In .CopyIDs
                objCopyIDs.Add tmpCopyID.ID & intIndex, "_" & tmpCopyID.ID & intIndex
            Next
            objItems.Add .ID & intIndex, .??????, .????, .????ID, .????, .????, .????, .????, .????, .????, .X, .Y, .W, .H, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .??????????, .????, .????, IIF(.??ID = 0, 0, .??ID & intIndex), objSubIDs, objCopyIDs, "_" & .ID & intIndex, .??????, .????????, .????????, .??????, .????????, .????????, .Relations, .ColProtertys
        End With
    Next
    
    Set objReport.Items = New RPTItems
    For Each tmpItem In objItems
        With tmpItem
            objReport.Items.Add .ID, .??????, .????, .????ID, .????, .????, .????, .????, .????, .????, .X, .Y, .W, .H, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .????, .??????????, .????, .????, .??ID, .SubIDs, .CopyIDs, "_" & .ID, .??????, .????????, .????????, .??????, .????????, .????????, .Relations, .ColProtertys
        End With
    Next
End Sub

Private Sub mnuViewStyle_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub SetView(bytStyle As Byte)
'??????????????????????????
'??????bytstyle=0-??????,1-??????,2-????,3-????????
    mnuViewStyle(0).Checked = False
    mnuViewStyle(1).Checked = False
    mnuViewStyle(2).Checked = False
    mnuViewStyle(3).Checked = False
    mnuViewStyle(bytStyle).Checked = True
    lvw.View = bytStyle
End Sub

Private Function GetSubReport(lngGroup As Long) As ADODB.Recordset
'????????????????ID,??????????????????
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH

    strSQL = "Select ??ID,????ID,????,???? From zlRPTSubs Where ??ID=[1] Order by ????"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngGroup)
    If Not rsTmp.EOF Then Set GetSubReport = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetGroupInfo(lngGroup As Long) As ADODB.Recordset
'????????????????ID??????????
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ID,????,????,????,????,????ID,???????? From zlRPTGroups Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngGroup)
    If Not rsTmp.EOF Then Set GetGroupInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitReportPars()
'????????????????,????????????????????????????????????
    Dim i As Integer, j As Integer
    Dim tmpPar As RPTPar, strTmp As String
    Dim lngCurH As Long, objTmp As Object
    Dim intCurTab As Integer, objLoad As Object
    Dim strGroup As String, objGroup As Object
    Dim blnCmd As Boolean, blnExist As Boolean
    Dim strPre As String, strCur As String
    Dim blnTmp As Boolean, lngTmp As Long
    
    For Each objLoad In lblName
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In txt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cmd
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cbo
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In dtp
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In opt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In chk
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fra
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fraGroup
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    
    blnMatch = False
    
    '????????????????
    i = 0: lngCurH = lblName(0).Top
    For Each tmpPar In mobjPars
        i = i + 1
        
        Load lblName(i)
        lblName(i).Caption = tmpPar.???? & "(&" & i & ")"
        lblName(i).ToolTipText = tmpPar.????
        lblName(i).Left = txt(0).Left - lblName(i).Width - 30
        lblName(i).Top = lngCurH
        lblName(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
        lblName(i).Visible = True
        
        If tmpPar.?????? = "????????????" Then
            If tmpPar.???? = 0 Then '??????
                Load cbo(i): Set objTmp = cbo(i)
                If tmpPar.???????? Then objTmp.Enabled = False
                cbo(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                cbo(i).Left = cbo(0).Left: cbo(i).Top = lblName(i).Top - (cbo(i).Height - lblName(i).Height) / 2
                '????????????
                For j = 0 To UBound(Split(tmpPar.??????, "|"))
                    strTmp = Split(Split(tmpPar.??????, "|")(j), ",")(0)
                    
                    If Left(strTmp, 1) = "??" Then
                        cbo(i).AddItem Mid(strTmp, 2)
                        If cbo(i).ListIndex = -1 Then cbo(i).ListIndex = cbo(i).NewIndex
                    Else
                        cbo(i).AddItem strTmp
                    End If
                    '??????????Reserve??????"??????|??????"
                    '??????????????????????????
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "??" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then cbo(i).ListIndex = cbo(i).NewIndex
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then cbo(i).ListIndex = cbo(i).NewIndex
                        End If
                        
                        '????????????????????????????????,??????
                        '??????????????????????????????,??????????????
                        If Split(tmpPar.Reserve, "|")(0) = Split(Split(tmpPar.??????, "|")(j), ",")(1) Then
                            cbo(i).ListIndex = cbo(i).NewIndex
                        End If
                    End If
                Next
                cbo(i).Visible = True
            ElseIf tmpPar.???? = 1 Then '??????
                Load fra(i): Set objTmp = fra(i)
                fra(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                fra(i).Left = fra(0).Left: fra(i).Top = lblName(i).Top - 50
                
                lblName(i).Visible = False
                fra(i).Caption = lblName(i).Caption
                                
                j = UBound(Split(tmpPar.??????, "|")) + 1 '??????
                j = CInt((j / 3) + 0.4) '????
                
                fra(i).Height = fra(0).Height + (j - 1) * (opt(0).Height * 1.6) - opt(0).Height * 0.3
                
                blnExist = False '????????????????????????????????
                '????????????
                For j = 0 To UBound(Split(tmpPar.??????, "|"))
                    strTmp = Split(Split(tmpPar.??????, "|")(j), ",")(0)
                    
                    Load opt(opt.UBound + 1)
                    If tmpPar.???????? Then opt(opt.UBound).Enabled = False
                    Set opt(opt.UBound).Container = fra(i)
                    opt(opt.UBound).TabIndex = intCurTab: intCurTab = intCurTab + 1
                    opt(opt.UBound).Tag = Split(Split(tmpPar.??????, "|")(j), ",")(1) '??????????
                    
                    If InStr(",0,1,3,", "," & UBound(Split(tmpPar.??????, "|")) & ",") > 0 Then
                        '????1,2,4????????????????
                        If j = 0 Or j = 1 Then 'Top
                            opt(opt.UBound).Top = opt(0).Top
                        Else
                            opt(opt.UBound).Top = opt(0).Top + opt(0).Height * 1.6
                        End If
                        If j = 0 Or j = 2 Then 'Left
                            opt(opt.UBound).Left = opt(0).Left + 150
                        Else
                            opt(opt.UBound).Left = opt(0).Left + (opt(0).Width * 1.4 + 60) + 150
                        End If
                        
                        If Left(strTmp, 1) = "??" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    Else
                        opt(opt.UBound).Top = opt(0).Top + (CInt(((j + 1) / 3) + 0.4) - 1) * (opt(0).Height * 1.6)
                        opt(opt.UBound).Left = opt(0).Left + (IIF(((j + 1) Mod 3) = 0, 3, ((j + 1) Mod 3)) - 1) * (opt(0).Width + 60)
                        
                        If Left(strTmp, 1) = "??" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    End If

                    opt(opt.UBound).Width = TextWidth(opt(opt.UBound).Caption) + 300
                    
                    '??????????Reserve??????"??????|??????"
                    '??????????????????????????
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "??" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                opt(opt.UBound).Value = True
                                blnExist = True
                            End If
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                opt(opt.UBound).Value = True
                                blnExist = True
                            End If
                        End If
                    End If
                    
                    opt(opt.UBound).Visible = True
                Next
                
                fra(i).ZOrder 1 '??????????
                fra(i).Visible = True
            ElseIf tmpPar.???? = 2 Then '??????????
                lblName(i).Visible = False
                
                blnTmp = True
                Load chk(i): Set objTmp = chk(i)
                If tmpPar.???????? Then objTmp.Enabled = False
                chk(i).Caption = lblName(i).Caption
                chk(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                chk(i).Left = chk(0).Left: chk(i).Top = lblName(i).Top - (chk(i).Height - lblName(i).Height) / 2
                chk(i).Width = TextWidth(chk(i).Caption) + 230
                
                '????????????
                If Left(Split(Split(tmpPar.??????, "|")(0), ",")(0), 1) = "??" Then chk(i).Value = 1
                For j = 0 To 1
                    strTmp = Split(Split(tmpPar.??????, "|")(j), ",")(0)
                    '??????????Reserve??????????"??????|??????"
                    '??????????????????????????????
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "??" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                If Left(strTmp, 1) = "??" Then
                                    chk(i).Value = 1
                                Else
                                    chk(i).Value = 0
                                End If
                            End If
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                If Left(strTmp, 1) = "??" Then
                                    chk(i).Value = 1
                                Else
                                    chk(i).Value = 0
                                End If
                            End If
                        End If
                    End If
                Next
                chk(i).Visible = True
            End If
        ElseIf tmpPar.?????? = "????????????" Then
            Load txt(i): Set objTmp = txt(i)
            If tmpPar.???????? Then objTmp.Enabled = False
            txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
            txt(i).Left = txt(0).Left: txt(i).Top = lblName(i).Top - (txt(i).Height - lblName(i).Height) / 2
            txt(i).ToolTipText = "?? F2 ??????????"
            txt(i).Locked = True
                                                
            blnCmd = True
            If tmpPar.Reserve Like "*|*" Then
                If Split(tmpPar.Reserve, "|")(0) <> "" Then
                    '??????????Reserve??????"??????|??????"
                    txt(i).Text = Split(tmpPar.Reserve, "|")(0)
                    txt(i).Tag = Split(tmpPar.Reserve, "|")(1)
                    
                    '??????????,??????????????????????????
                    strTmp = ""
                    If InStr(tmpPar.????, "|") > 0 Then strTmp = Split(tmpPar.????, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.????SQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.????, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.????????, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                Else
                    '????????????????????
                    If tmpPar.?????? Like "*|*" Then
                        txt(i).Text = Split(tmpPar.??????, "|")(0)
                        txt(i).Tag = Split(tmpPar.??????, "|")(1)
                    ElseIf tmpPar.????SQL <> "" Then
                        '??????SQL??????????????,????????????,????????
                        strTmp = ""
                        If InStr(tmpPar.????, "|") > 0 Then strTmp = Split(tmpPar.????, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.????SQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.????, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.????????, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                        If strTmp <> "" Then
                            txt(i).Text = Split(strTmp, "|")(0)
                            txt(i).Tag = Split(strTmp, "|")(1)
                            If tmpPar.???? = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                            blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                        Else
                            blnCmd = False
                        End If
                    End If
                End If
            Else
                If tmpPar.?????? Like "*|*" Then
                    '????????????????????
                    txt(i).Text = Split(tmpPar.??????, "|")(0)
                    txt(i).Tag = Split(tmpPar.??????, "|")(1)
                    
                    '??????????,??????????????????????????
                    strTmp = ""
                    If InStr(tmpPar.????, "|") > 0 Then strTmp = Split(tmpPar.????, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.????SQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.????, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.????????, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                ElseIf tmpPar.????SQL <> "" Then
                    '??????SQL??????????????,????????????,????????
                    strTmp = ""
                    If InStr(tmpPar.????, "|") > 0 Then strTmp = Split(tmpPar.????, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.????SQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.????, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.????????, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        txt(i).Text = Split(strTmp, "|")(0)
                        txt(i).Tag = Split(strTmp, "|")(1)
                        If tmpPar.???? = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                    Else
                        blnCmd = False
                    End If
                End If
            End If
                        
            Load cmd(i)
            If tmpPar.???????? Then cmd(i).Enabled = False
            cmd(i).Top = txt(i).Top + 30
            cmd(i).Left = txt(i).Left + txt(i).Width - cmd(i).Width - 30
            cmd(i).Height = txt(i).Height - 45
            cmd(i).TabStop = False
            cmd(i).ZOrder
            
            txt(i).Visible = True
            cmd(i).Visible = blnCmd
            
            '????????????
            txt(i).Locked = Not ((InStr(tmpPar.????SQL, "[*]") > 0 Or InStr(tmpPar.????SQL, "[*]") > 0) And blnCmd)
        Else
            If tmpPar.???? = 2 Then
                Load dtp(i): Set objTmp = dtp(i)
                If tmpPar.???????? Then objTmp.Enabled = False
                dtp(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                dtp(i).Left = dtp(0).Left: dtp(i).Top = lblName(i).Top - (dtp(i).Height - lblName(i).Height) / 2
                If InStr(tmpPar.??????, ":") > 0 Or InStr(tmpPar.??????, "????") > 0 Then
                    dtp(i).CustomFormat = "yyyy??MM??dd?? HH:mm:ss"
                    dtp(i).Width = 2460
                Else
                    dtp(i).CustomFormat = "yyyy??MM??dd??"
                    dtp(i).Width = 1635
                End If
                If tmpPar.?????? <> "" Then
                    If Left(tmpPar.??????, 1) = "&" Then
                        dtp(i).Value = GetParVBMacro(tmpPar.??????)
                    Else
                        dtp(i).Value = Format(tmpPar.??????, dtp(i).CustomFormat)
                    End If
                Else
                    dtp(i).Value = Currentdate
                End If
                
'                '????????????
'                If dtp(i).CustomFormat Like "*HH:mm:ss" And Left(tmpPar.??????, 1) <> "&" Then
'                    strTmp = GetSetting("ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.????, lblName(i).ToolTipText & "????", Format(dtp(i).Value, "HH:mm:ss"))
'                    dtp(i).Value = CDate(Format(dtp(i).Value, Left(dtp(i).CustomFormat, InStr(dtp(i).CustomFormat, "HH:mm:ss") - 1)) & strTmp)
'                End If
                
                dtp(i).Visible = True
            Else
                Load txt(i): Set objTmp = txt(i)
                If tmpPar.???????? Then objTmp.Enabled = False
                txt(i).Left = txt(0).Left: txt(i).Top = lblName(i).Top - (txt(i).Height - lblName(i).Height) / 2
                txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                txt(i).Text = tmpPar.??????
                txt(i).Visible = True
            End If
        End If
        If objTmp.name = "fra" Then
            lngCurH = lngCurH + objTmp.Height + 180
        Else
            lngCurH = lngCurH + txt(0).Height + 150
        End If
        
        lblName(i).Tag = tmpPar.???? & "," & objTmp.name
        If tmpPar.?????? = "????????????" Then lblName(i).Tag = lblName(i).Tag & ",cmd"
    Next
    
    fraSplit.Top = lngCurH
    
    '??????????
    For i = 1 To lblName.UBound
        strCur = ""
        If strGroup <> CStr(Split(lblName(i).Tag, ",")(0)) And CStr(Split(lblName(i).Tag, ",")(0)) <> "" Then
            Load fraGroup(fraGroup.UBound + 1)
            Set objGroup = fraGroup(fraGroup.UBound)
            objGroup.Caption = CStr(Split(lblName(i).Tag, ",")(0))
            objGroup.Top = lblName(i).Top - 150
            objGroup.ZOrder 1
            objGroup.Visible = True
            
            Select Case CStr(Split(lblName(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            lngCurH = 195 '????Top????
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            objTmp.Left = 1250
            
            Set lblName(i).Container = objGroup
            lblName(i).Top = objTmp.Top + (objTmp.Height - lblName(i).Height) / 2
            lblName(i).Left = objTmp.Left - lblName(i).Width - 30
            lblName(i).Caption = GetLenStr(lblName(i).ToolTipText, 900, Me) & Mid(lblName(i).Caption, InStr(lblName(i).Caption, "("))
            
            If UBound(Split(lblName(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If

            lngCurH = lngCurH + txt(0).Height + 50 '????Top????
        ElseIf strGroup = CStr(Split(lblName(i).Tag, ",")(0)) And CStr(Split(lblName(i).Tag, ",")(0)) <> "" Then
            strCur = "Add"
            Select Case CStr(Split(lblName(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            objTmp.Left = 1250
            
            Set lblName(i).Container = objGroup
            lblName(i).Top = objTmp.Top + (objTmp.Height - lblName(i).Height) / 2
            lblName(i).Left = objTmp.Left - lblName(i).Width - 30
            lblName(i).Caption = GetLenStr(lblName(i).ToolTipText, 900, Me) & Mid(lblName(i).Caption, InStr(lblName(i).Caption, "("))
            
            If UBound(Split(lblName(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If
                        
            lngCurH = lngCurH + txt(0).Height + 50 '????Top????
            
            objGroup.Height = objTmp.Top + objTmp.Height + 90  '??????
            
            '??????????????????????????
            For j = i + 1 To lblName.UBound
                If Split(lblName(j).Tag, ",")(0) <> "fra" Then
                    lblName(j).Top = lblName(j).Top + 60
                    Select Case CStr(Split(lblName(j).Tag, ",")(1))
                        Case "txt"
                            txt(j).Top = txt(j).Top + 60
                        Case "cbo"
                            cbo(j).Top = cbo(j).Top + 60
                        Case "dtp"
                            dtp(j).Top = dtp(j).Top + 60
                        Case "chk"
                            chk(j).Top = chk(j).Top + 60
                    End Select
                    If UBound(Split(lblName(j).Tag, ",")) = 2 Then
                        cmd(j).Top = cmd(j).Top + 60
                    End If
                End If
            Next
        End If
        If strPre = "Add" And strCur = "" Then
            fraSplit.Top = fraSplit.Top + 60
        End If
        strPre = strCur
        strGroup = CStr(Split(lblName(i).Tag, ",")(0))
    Next
    
    '??????????????????????????,??????????
    If fraGroup.UBound = 0 And fra.UBound > 0 Then
        For Each objTmp In fra
            objTmp.Left = txt(0).Left - 1000
        Next
    End If
    
    cmdLoad.Top = fraSplit.Top + 180
    cmdDefault.Top = fraSplit.Top + 180
    
    fraSplit.Visible = (lblName.UBound > 0)
    cmdLoad.Visible = (lblName.UBound > 0)
    cmdDefault.Visible = (lblName.UBound > 0)
    
    cmdLoad.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdDefault.TabIndex = intCurTab
    
    cmdSelAll.Top = cmdLoad.Top: cmdSelNone.Top = cmdSelAll.Top
    cmdSelAll.Visible = blnTmp
    cmdSelNone.Visible = blnTmp
    If Me.Visible Then
        On Error Resume Next
        If picPar.Height < cmdLoad.Top + cmdLoad.Height + 100 Then
            lngTmp = cmdLoad.Top + cmdLoad.Height + 100 - picPar.Height
            picPar.Height = picPar.Height + lngTmp
            picPar.Top = picPar.Top - lngTmp: lblPar_S.Top = lblPar_S.Top - lngTmp
            lvw.Height = lvw.Height - lngTmp
        End If
    End If
    
    '????????????
    Call LoadCondsMenu
End Sub

Private Function ReSetReportPars() As Boolean
'??????????????????????????????????????
    Dim i As Integer, j As Integer
    Dim strTmp As String, strDisp As String
    Dim strParName As String, curDate As Date
    
    '????????????
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).?????? = "????????????" Then
            Select Case mobjPars("_" & strParName).????
                Case 0
                    If Trim(cbo(i).Text) = "" Then
                        MsgBox "??????""" & strParName & """??????????", vbInformation, App.Title
                        If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                        Exit Function
                    End If
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '????????????
                        '????????
                        Select Case mobjPars("_" & strParName).????
                            Case 1
                                If Not IsNumeric(cbo(i).Text) Then
                                    MsgBox "????????""" & strParName & """??????????????????????????", vbInformation, App.Title
                                    If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                                    Exit Function
                                End If
                            Case 2
                                If Not IsDate(cbo(i).Text) Then
                                    MsgBox "????????""" & strParName & """??????????????????????????", vbInformation, App.Title
                                    If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                                    Exit Function
                                End If
                        End Select
                    End If
            End Select
        ElseIf mobjPars("_" & strParName).?????? = "????????????" Then
            If Trim(txt(i).Text) = "" Then
                MsgBox "??????""" & strParName & """??????????", vbInformation, App.Title
                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                Exit Function
            End If
            If txt(i).Tag = "" Then '????????????
                If mobjPars("_" & strParName).?????? Like "*|*" Then
                    If Split(mobjPars("_" & strParName).??????, "|")(0) <> txt(i).Text Then
                        '????????
                        Select Case mobjPars("_" & strParName).????
                            Case 1
                                If Not IsNumeric(txt(i).Text) Then
                                    MsgBox "????????""" & strParName & """??????????????????????????", vbInformation, App.Title
                                    If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                    Exit Function
                                End If
                            Case 2
                                If Not IsDate(txt(i).Text) Then
                                    MsgBox "????????""" & strParName & """??????????????????????????", vbInformation, App.Title
                                    If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                    Exit Function
                                End If
                        End Select
                    Else
                        '????????????????????????,??????????????
                        txt(i).Tag = Split(mobjPars("_" & strParName).??????, "|")(1)
                    End If
                Else
                    '????????
                    Select Case mobjPars("_" & strParName).????
                        Case 1
                            If Not IsNumeric(txt(i).Text) Then
                                MsgBox "????????""" & strParName & """??????????????????????????", vbInformation, App.Title
                                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                Exit Function
                            End If
                        Case 2
                            If Not IsDate(txt(i).Text) Then
                                MsgBox "????????""" & strParName & """??????????????????????????", vbInformation, App.Title
                                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                Exit Function
                            End If
                    End Select
                End If
            End If
        Else
            Select Case mobjPars("_" & strParName).????
                Case 0, 3
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "??????""" & strParName & """??????????", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If TLen(txt(i).Text) > 255 Then
                        MsgBox """" & strParName & """????????????????????255????????", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                Case 1
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "??????""" & strParName & """??????????", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If TLen(txt(i).Text) > 255 Then
                        MsgBox """" & strParName & """????????????????????255????????", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If Not IsNumeric(txt(i).Text) Then
                        MsgBox """" & strParName & """??????????????????????????", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                Case 2 '??????????????????
                    curDate = Currentdate
                    If Not (mobjPars("_" & strParName).?????? Like "&????*" Or mobjPars("_" & strParName).Reserve Like "&????*" Or _
                        mobjPars("_" & strParName).?????? Like "&????*" Or mobjPars("_" & strParName).Reserve Like "&????*" Or _
                        mobjPars("_" & strParName).?????? Like "&*????*" Or mobjPars("_" & strParName).Reserve Like "&*????*" Or _
                        mobjPars("_" & strParName).?????? Like "&*????*" Or mobjPars("_" & strParName).?????? Like "&*????*" Or _
                        mobjPars("_" & strParName).Reserve Like "&*????*" Or mobjPars("_" & strParName).Reserve Like "&*????*") Then
                        
                        If mobjPars("_" & strParName).?????? Like "*????*" Or mobjPars("_" & strParName).Reserve Like "*????*" Then
                            If Format(dtp(i).Value, "yyyy-MM-dd HH:mm:ss") > Format(curDate, "yyyy-MM-dd HH:mm:ss") Then
                                MsgBox """" & strParName & """ ??????????????????????????", vbInformation, App.Title
                                If dtp(i).Enabled And dtp(i).Visible Then dtp(i).SetFocus
                                Exit Function
                            End If
                        Else
                            If Format(dtp(i).Value, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
                                MsgBox """" & strParName & """ ??????????????????????????", vbInformation, App.Title
                                If dtp(i).Enabled And dtp(i).Visible Then dtp(i).SetFocus
                                Exit Function
                            End If
                        End If
                    End If
            End Select
        End If
    Next
        
    '??????
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).?????? = "????????????" Then '????????????
            Select Case mobjPars("_" & strParName).????
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '????????????
                        'Reserve??????????????????"????????|??????"
                        mobjPars("_" & strParName).Reserve = "????????????|" & cbo(i).Text
                        mobjPars("_" & strParName).?????? = cbo(i).Text
                    Else
                        '????????
                        'Reserve??????????????????"????????|??????"
                        mobjPars("_" & strParName).Reserve = "????????????|" & cbo(i).Text
                        strTmp = mobjPars("_" & strParName).??????
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "??" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                mobjPars("_" & strParName).?????? = Split(Split(strTmp, "|")(j), ",")(1)
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                'Reserve??????????????????"????????|??????"
                                mobjPars("_" & strParName).Reserve = "????????????|" & opt(j).ToolTipText
                                mobjPars("_" & strParName).?????? = opt(j).Tag
                            End If
                        End If
                    Next
                Case 2
                    'Reserve??????????????????"????????|??????"
                    strTmp = mobjPars("_" & strParName).??????
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "??" Then
                                mobjPars("_" & strParName).Reserve = "????????????|" & strDisp
                                mobjPars("_" & strParName).?????? = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        Else
                            If Left(strDisp, 1) = "??" Then
                                mobjPars("_" & strParName).Reserve = "????????????|" & Mid(strDisp, 2)
                                mobjPars("_" & strParName).?????? = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).?????? = "????????????" Then
            If txt(i).Tag = "" Then '????????????
                'Reserve??????????????????"????????|??????"
                mobjPars("_" & strParName).Reserve = "????????????|"
                mobjPars("_" & strParName).?????? = txt(i).Text
            Else
                '????????
                'Reserve??????????????????"????????|??????"
                mobjPars("_" & strParName).Reserve = "????????????|" & txt(i).Text
                mobjPars("_" & strParName).?????? = txt(i).Tag
            End If
        Else
            Select Case mobjPars("_" & strParName).????
                Case 0, 1, 3
                    mobjPars("_" & strParName).?????? = txt(i).Text
                Case 2
                    If mobjPars("_" & strParName).?????? Like "&*" Then
                        mobjPars("_" & strParName).Reserve = mobjPars("_" & strParName).??????
                    End If
                    mobjPars("_" & strParName).?????? = Format(dtp(i).Value, dtp(i).CustomFormat)
                    '????????????
                    If dtp(i).CustomFormat Like "*HH:mm:ss" Then
                        SaveSetting "ZLSOFT", "????????\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.????, lblName(i).ToolTipText & "????", Format(dtp(i).Value, "HH:mm:ss")
                    End If
            End Select
        End If
    Next
    
    Call ReplaceInputPars(mobjPars)
    
    ReSetReportPars = True
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim LngIdx As Long
    
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
    If InStr("~`!@#$^&"";|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If mobjPars("_" & lblName(Index).ToolTipText).???? = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii <> 8 Then
        If SendMessage(cbo(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
        LngIdx = MatchIndex(cbo(Index), KeyAscii)
        If LngIdx <> -2 Then cbo(Index).ListIndex = LngIdx
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Function GetValues() As Collection
'??????????????????????????????
    Dim i As Integer, j As Integer
    Dim strParName As String, strTmp As String
    Dim strDisp As String, colValue As New Collection
     
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).?????? = "????????????" Then
            Select Case mobjPars("_" & strParName).????
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '????????????
                        'Reserve??????????????????"????????|??????"
                        colValue.Add cbo(i).Text, "_" & strParName
                    Else
                        '????????
                        'Reserve??????????????????"????????|??????"
                        '????????????
                        strTmp = mobjPars("_" & strParName).??????
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "??" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                colValue.Add opt(j).Tag, "_" & strParName
                            End If
                        End If
                    Next
                Case 2
                    'Reserve??????????????????"????????|??????"
                    '????????????
                    strTmp = mobjPars("_" & strParName).??????
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "??" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        Else
                            If Left(strDisp, 1) = "??" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).?????? = "????????????" Then
            If txt(i).Tag = "" Then '????????????
                'Reserve??????????????????"????????|??????"
                colValue.Add txt(i).Text, "_" & strParName
            Else
                '????????
                'Reserve??????????????????"????????|??????"
                colValue.Add txt(i).Tag, "_" & strParName
            End If
        Else
            Select Case mobjPars("_" & strParName).????
                Case 0, 1, 3
                    colValue.Add txt(i).Text, "_" & strParName
                Case 2
                    colValue.Add Format(dtp(i).Value, dtp(i).CustomFormat), "_" & strParName
            End Select
        End If
    Next
    Set GetValues = colValue
End Function

Private Sub cmd_Click(Index As Integer)
    Dim tmpPar As RPTPar, str???????? As String, str???????? As String
    Dim frmNewSelect As New frmSelect
    Dim strSQL???? As String, strSQL???? As String
    Dim colValue As New Collection    '????????????
    
    For Each tmpPar In mobjPars
        If tmpPar.???? = lblName(Index).ToolTipText Then
            If blnMatch And txt(Index).Tag = "" Then frmNewSelect.strMatch = txt(Index).Text
            
            If InStr(tmpPar.????, "|") > 0 Then
                str???????? = Split(tmpPar.????, "|")(0)
                str???????? = Split(tmpPar.????, "|")(1)
            End If
            strSQL???? = tmpPar.????SQL
            strSQL???? = tmpPar.????SQL
            Set colValue = GetValues
            Call CheckParsRela(strSQL????, Nothing, tmpPar.????, True, colValue, mobjPars)
            Call CheckParsRela(strSQL????, Nothing, tmpPar.????, True, colValue, mobjPars)
            frmNewSelect.strSQLList = SQLOwner(RemoveNote(strSQL????), str????????)
            frmNewSelect.strSQLTree = SQLOwner(RemoveNote(strSQL????), str????????)
            frmNewSelect.strFLDList = tmpPar.????????
            frmNewSelect.strFLDTree = tmpPar.????????
            frmNewSelect.strParName = tmpPar.????
            frmNewSelect.bytType = tmpPar.????
            frmNewSelect.mblnMulti = tmpPar.???? = 1
            frmNewSelect.mintConnect = GetDBConnectNo(tmpPar, mobjReport.Datas)
            frmNewSelect.lngSeekHwnd = cmd(Index).hwnd
            
            On Error Resume Next
            Err.Clear
            
            frmNewSelect.Show 1, Me
            If frmNewSelect.mblnOK Then
                txt(Index).Text = frmNewSelect.strOutDisp
                txt(Index).Tag = frmNewSelect.strOutBand
                Unload frmNewSelect
                SendKeys "{Tab}"
            ElseIf blnMatch Then
                txt(Index).Text = ""
                txt(Index).Tag = ""
            End If
            
            blnMatch = False
            Exit For
        End If
    Next
    txt(Index).SetFocus
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub dtp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SelAll txt(Index)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And txt(Index).ToolTipText <> "" Then
        If cmd(Index).Enabled And cmd(Index).Visible Then Call cmd_Click(Index)
    End If
    If txt(Index).Locked Then Exit Sub
    
    '??????????(??????)??????????????????????????????
    '144=Num;112-123=F1-F12;229=????????????
    If KeyCode >= 48 And KeyCode <> 144 And KeyCode <> 229 _
        And Not (KeyCode >= 112 And KeyCode <= 123) Then
        txt(Index).Tag = ""
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
            '??????????
            KeyAscii = 0
            If txt(Index).Text <> "" Then
                If cmd(Index).Enabled And cmd(Index).Visible Then
                    blnMatch = True
                    Call cmd_Click(Index)
                End If
            End If
            Exit Sub
        Else
            '??????????
            KeyAscii = 0: SendKeys "{Tab}": Exit Sub
        End If
    End If
    
    If txt(Index).Locked Then Exit Sub
    
    If InStr("~`!@#$^&"";|'" & Chr(3) & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If txt(Index).ToolTipText = "" And mobjPars("_" & lblName(Index).ToolTipText).???? = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '??????????(??????)??????????????????????????????
    '??????????????,??????KeyDown??????
    If KeyAscii < 0 Then txt(Index).Tag = ""
End Sub

Private Sub ShowStatGrid(objItem As RPTItem, lngW As Long, lngH As Long)
'??????????????????????????????????????(????????????????)
    Dim mshBody As Object, tmpItem As RPTItem, tmpID As RelatID
    Dim rsGroup As ADODB.Recordset, rsVsc As ADODB.Recordset, rsHsc As ADODB.Recordset
    Dim arrStat() As Variant, strVscStat As String, strHscStat As String, strStat As String
    Dim strVsc As String, strHsc As String, strVscOrder As String, strHscOrder As String
    Dim strFilter As String, strAlign As String, strTmp As String
    Dim i As Long, j As Long, k As Long, L As Long, M As Long
    Dim X As Long, Y As Long, Z As Long '??????????
    Dim strFormat As String, strSort As String, blnHide As Boolean, blnDo As Boolean
    Dim arrLevel() As String, arrMerge() As String, arrCount() As Long
        
    '??????????????????????
    Dim lngCurCols As Long, objCurItem As RPTItem
    Dim strLink As String, lngMaxY As Long
    Dim lngGrid As Long, strTopRow As String
    Dim lngColB As Long, lngColE As Long, lngDiff As Long
    Dim lngStatistics As Long
    
    '????????????????????
    Dim colVsc As Collection, colHsc As Collection
    Dim strKey As String, lngRow As Long, lngCol As Long, StrFmt As String
    
    Dim varIFValue As Variant
    Dim objColProp As RPTColProterty
    Dim objStatusGridItem As RPTItem
    
    With objItem
        Load msh(.ID)
        Set msh(.ID).Container = picPaper(intReport)
        Set mshBody = msh(.ID)
        
        mshBody.Redraw = False
        
        mshBody.ForeColor = .????
        mshBody.ForeColorFixed = .????
        mshBody.BackColor = .????
        mshBody.BackColorFixed = .????
        mshBody.GridColor = .????
        mshBody.GridColorFixed = .????
        mshBody.Font.name = .????
        mshBody.Font.Size = .????
        mshBody.Font.Bold = .????
        mshBody.Font.Italic = .????
        mshBody.Font.Underline = .????
        mshBody.GridLineWidth = IIF(.??????????, 2, 1)
        'Set mshBody.FontFixed = mshBody.Font
        
        mshBody.Left = .X: mshBody.Top = .Y
        mshBody.Height = .H: mshBody.Width = 0
        mshBody.FixedRows = 0
    
        '??????????????????????
        strLink = strLink & "|" & .ID
        For Each tmpItem In mobjReport.Items
            If tmpItem.?????? = bytFormat And tmpItem.???? = 5 And tmpItem.???? = 2 And tmpItem.???? = .???? Then
                strLink = strLink & "|" & tmpItem.ID
            End If
        Next
        strLink = Mid(strLink, 2)
    End With
        
    objItem.???? = ""
    strTopRow = ""
    
    blnHide = True
    lngCurCols = 0
    lngMaxY = 0
    For lngGrid = 0 To UBound(Split(strLink, "|"))
        Set objCurItem = mobjReport.Items("_" & Split(strLink, "|")(lngGrid))
        With objCurItem
            mshBody.Width = mshBody.Width + .W
            
            '??????????
            '????????????????????????????
            If lngGrid = 0 Then
                strVsc = "": strVscOrder = "": X = 0
            End If
            strHsc = "": strHscOrder = "" '????????????????????????????????
            Y = 0: Z = 0
            For Each tmpID In .SubIDs
                Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                Select Case tmpItem.????
                    Case 7 '????????
                        If lngGrid = 0 Then
                            X = X + 1
                            If tmpItem.???? <> "" Then strVscOrder = strVscOrder & "|" & tmpItem.????
                            strVsc = strVsc & "|" & tmpItem.????
                        End If
                    Case 8 '????????
                        Y = Y + 1
                        If tmpItem.???? <> "" Then strHscOrder = strHscOrder & "|" & tmpItem.????
                        strHsc = strHsc & "|" & tmpItem.????
                    Case 9 '??????
                        Z = Z + 1
                End Select
            Next
            If Y > lngMaxY Then lngMaxY = Y
            If lngGrid = 0 Then
                strVsc = Mid(strVsc, 2)
                strVscOrder = Mid(strVscOrder, 2)
            End If
            strHsc = Mid(strHsc, 2)
            strHscOrder = Mid(strHscOrder, 2)
            
            '????????????
            If lngGrid = 0 Then
                mshBody.FixedRows = Y + 1
            Else
                If Y + 1 > mshBody.FixedRows Then
                    lngDiff = Y + 1 - mshBody.FixedRows '????????????(????????????????????????)
                    For i = 1 To Y + 1 - mshBody.FixedRows
                        mshBody.AddItem "", mshBody.FixedRows
                        mshBody.FixedRows = mshBody.FixedRows + 1
                        For j = 0 To mshBody.Cols - 1 '????????????????????????????????????????
                            mshBody.TextMatrix(mshBody.FixedRows - 1, j) = mshBody.TextMatrix(mshBody.FixedRows - 2, j)
                        Next
                    Next
                End If
            End If
            mshBody.Cols = lngCurCols + IIF(lngGrid = 0, X, 0) + Z
            If lngGrid = 0 Then
                mshBody.Rows = Y + 2
                mshBody.FixedCols = X
            End If
            lngStatistics = 0
            For Each tmpID In .SubIDs
                Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                Select Case tmpItem.????
                    Case 7 '????????
                        If lngGrid = 0 Then
                            For i = 0 To Y
                                mshBody.TextMatrix(i, tmpItem.????) = tmpItem.????
                            Next
                        End If
                    Case 8 '????????
                    Case 9 '??????
                        lngStatistics = lngStatistics + 1
                        For i = mshBody.FixedRows - 1 To Y
                            mshBody.TextMatrix(i, lngCurCols + IIF(lngGrid = 0, X, 0) + tmpItem.????) = tmpItem.????
                        Next
                End Select
            Next
            
            '-------------------------------------------------------------------------------------
            '????????????
            '-------------------------------------------------------------------------------------
            If mLibDatas("_" & .????).DataSet.RecordCount > 0 Then
                Set rsGroup = Nothing
                Set rsGroup = mLibDatas("_" & .????).DataSet.Clone
                
                '1.????????????
                
                '1.1:????????????????????(????????????)
                If lngGrid = 0 Then
                    Set rsVsc = Nothing
                    Set rsVsc = New ADODB.Recordset
                    '1.1.1:??????????????????
                    For i = 0 To UBound(Split(strVscOrder, "|"))
                        If Left(Split(strVscOrder, "|")(i), 1) = "," Then
                            With rsGroup.Fields(Mid(Split(strVscOrder, "|")(i), 2))
                                '??????adNumeric????????????,????????adBigInt??adSingle/adDouble
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        Else
                            With rsGroup.Fields(Split(strVscOrder, "|")(i))
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    '1.1.2:??????????????????(????????????????)
                    For i = 0 To UBound(Split(strVsc, "|"))
                        If InStr("|" & Replace(strVscOrder, ",", "") & "|", "|" & Split(strVsc, "|")(i) & "|") = 0 Then
                            With rsGroup.Fields(Split(strVsc, "|")(i))
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    rsVsc.CursorLocation = adUseClient
                    rsVsc.LockType = adLockBatchOptimistic
                    rsVsc.CursorType = adOpenStatic
                    rsVsc.Open
                End If
                
                '1.2:????????????????????(????????????)
                Set rsHsc = Nothing
                If strHsc <> "" Then
                    Set rsHsc = New ADODB.Recordset
                    '1.2.1:??????????????????
                    For i = 0 To UBound(Split(strHscOrder, "|"))
                        If Left(Split(strHscOrder, "|")(i), 1) = "," Then
                            With rsGroup.Fields(Mid(Split(strHscOrder, "|")(i), 2))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        Else
                            With rsGroup.Fields(Split(strHscOrder, "|")(i))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    '1.2.2:??????????????????(????????????????)
                    For i = 0 To UBound(Split(strHsc, "|"))
                        If InStr("|" & Replace(strHscOrder, ",", "") & "|", "|" & Split(strHsc, "|")(i) & "|") = 0 Then
                            With rsGroup.Fields(Split(strHsc, "|")(i))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    rsHsc.CursorLocation = adUseClient
                    rsHsc.LockType = adLockBatchOptimistic
                    rsHsc.CursorType = adOpenStatic
                    rsHsc.Open
                End If
                
                '1.3:??????????????
                rsGroup.MoveFirst
                For i = 1 To rsGroup.RecordCount
                    '????????
                    If Not rsVsc Is Nothing And lngGrid = 0 Then
                        strFilter = "" '??????????????????????
                        For j = 0 To UBound(Split(strVsc, "|")) '????????????????
                            strFilter = strFilter & " And " & Split(strVsc, "|")(j) & "="
                            Select Case rsGroup.Fields(Split(strVsc, "|")(j)).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "'" & Replace(rsGroup.Fields(Split(strVsc, "|")(j)).Value, " ", "????") & "'"
                                    Else
                                        strFilter = strFilter & "'#'"
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        strFilter = strFilter & rsGroup.Fields(Split(strVsc, "|")(j)).Value
                                    Else
                                        strFilter = strFilter & "123456707654321"
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        '????????????????????????,??#02-4-9#??????"2009-02-04"
                                        strFilter = strFilter & "#" & Format(rsGroup.Fields(Split(strVsc, "|")(j)).Value, "yyyy-MM-dd HH:mm:ss") & "#"
                                    Else
                                        strFilter = strFilter & "#3000-05-05#"
                                    End If
                            End Select
                        Next
                        rsVsc.Filter = Replace(Mid(strFilter, 6), "????", " ")
                        If rsVsc.EOF Then
                            rsVsc.AddNew
                            For j = 0 To rsVsc.Fields.count - 1 '????????????????
                                If Not IsNull(rsGroup.Fields(rsVsc.Fields(j).name).Value) Then
                                    rsVsc.Fields(j).Value = rsGroup.Fields(rsVsc.Fields(j).name).Value
                                Else
                                    Select Case rsGroup.Fields(rsVsc.Fields(j).name).type
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            rsVsc.Fields(j).Value = "#" '??????
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            rsVsc.Fields(j).Value = 123456707654321# '??????
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            rsVsc.Fields(j).Value = #5/5/3000#   '??????
                                    End Select
                                End If
                            Next
                        End If
                    End If
                    '????????
                    If Not rsHsc Is Nothing Then
                        strFilter = "" '??????????????????????
                        For j = 0 To UBound(Split(strHsc, "|")) '????????????????
                            strFilter = strFilter & " And " & Split(strHsc, "|")(j) & "="
                            Select Case rsGroup.Fields(Split(strHsc, "|")(j)).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "'" & rsGroup.Fields(Split(strHsc, "|")(j)).Value & "'"
                                    Else
                                        strFilter = strFilter & "'#'"
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & rsGroup.Fields(Split(strHsc, "|")(j)).Value
                                    Else
                                        strFilter = strFilter & "123456707654321"
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "#" & Format(rsGroup.Fields(Split(strHsc, "|")(j)).Value, "yyyy-MM-dd HH:mm:ss") & "#"
                                    Else
                                        strFilter = strFilter & "#3000-05-05#"
                                    End If
                            End Select
                        Next
                        rsHsc.Filter = Mid(strFilter, 6)
                        If rsHsc.EOF Then
                            rsHsc.AddNew
                            For j = 0 To rsHsc.Fields.count - 1 '????????????????
                                If Not IsNull(rsGroup.Fields(rsHsc.Fields(j).name).Value) Then
                                    rsHsc.Fields(j).Value = rsGroup.Fields(rsHsc.Fields(j).name).Value
                                Else
                                    Select Case rsGroup.Fields(rsHsc.Fields(j).name).type
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            rsHsc.Fields(j).Value = "#" '??????
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            rsHsc.Fields(j).Value = 123456707654321# '??????
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            rsHsc.Fields(j).Value = #5/5/3000#   '??????
                                    End Select
                                End If
                            Next
                        End If
                    End If
                    rsGroup.MoveNext
                Next
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    rsVsc.UpdateBatch adAffectAllChapters
                    rsVsc.Filter = 0
                End If
                If Not rsHsc Is Nothing Then
                    rsHsc.UpdateBatch adAffectAllChapters
                    rsHsc.Filter = 0
                End If
                
                '1.4:??????????????????
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    strSort = ""
                    For i = 0 To UBound(Split(strVscOrder, "|"))
                        If Left(Split(strVscOrder, "|")(i), 1) = "," Then
                            strSort = strSort & "," & Mid(Split(strVscOrder, "|")(i), 2) & " Desc"
                        Else
                            strSort = strSort & "," & Split(strVscOrder, "|")(i)
                        End If
                    Next
                    If strSort <> "" Then rsVsc.Sort = Mid(strSort, 2)
                    rsVsc.MoveFirst
                End If
                If Not rsHsc Is Nothing Then
                    strSort = ""
                    For i = 0 To UBound(Split(strHscOrder, "|"))
                        If Left(Split(strHscOrder, "|")(i), 1) = "," Then
                            strSort = strSort & "," & Mid(Split(strHscOrder, "|")(i), 2) & " Desc"
                        Else
                            strSort = strSort & "," & Split(strHscOrder, "|")(i)
                        End If
                    Next
                    If strSort <> "" Then rsHsc.Sort = Mid(strSort, 2)
                    rsHsc.MoveFirst
                End If
                
                '1.5:??????????????????????
                '????????
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    Set colVsc = New Collection
                    '????????
                    strVscStat = ""
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                        If tmpItem.???? = 7 Then strVscStat = strVscStat & "," & tmpItem.????
                    Next
                    strVscStat = Mid(strVscStat, 2)
                    
                    '????????
                    k = Y  '????????????????
                    ReDim arrLevel(X - 1) '??????????????????????????????????
                    ReDim arrMerge(X - 1) '????????????????????????????????
                    For i = 1 To X - 1
                        arrMerge(i) = Space(i Mod 2)
                    Next
                    For i = 1 To rsVsc.RecordCount
                        k = k + 1
                        If mshBody.Rows - 1 < k Then mshBody.Rows = mshBody.Rows + 1
                        strKey = ""
                        For j = 0 To X - 1
                            strTmp = Trim(mshBody.TextMatrix(Y, j))
                            Select Case rsVsc.Fields(strTmp).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If rsVsc.Fields(strTmp).Value = "#" Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " " '????????????????????
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "??")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If rsVsc.Fields(strTmp).Value = 123456707654321# Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " "
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "??")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If rsVsc.Fields(strTmp).Value = #5/5/3000# Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " "
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "??")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                            End Select
                        Next
                        
                        '??????????(????????????)
                        For j = X - 1 To 1 Step -1 '??????????
                            strTmp = GetRowText(mshBody, k, j - 1)
                            If strTmp <> arrLevel(j) And k > Y + 1 Then
                                If strVscStat <> "" Then
                                    If Split(strVscStat, ",")(j) <> "" Then
                                        mshBody.AddItem "", k
                                        mshBody.Row = k
                                        For L = 0 To j - 1
                                            mshBody.TextMatrix(k, L) = mshBody.TextMatrix(k - 1, L)
                                        Next
                                        For L = j To X - 1
                                            mshBody.Col = L
                                            mshBody.CellAlignment = 4
                                            'mshBody.TextMatrix(k, L) = Space(j Mod 2) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j Mod 2)
                                            mshBody.TextMatrix(k, L) = Space(j) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j)
                                        Next
                                        mshBody.RowData(k) = j + 1
                                        mshBody.MergeRow(k) = True
                                        
                                        k = k + 1
                                    End If
                                End If
                                arrMerge(j) = IIF(arrMerge(j) = "", " ", "")
                            End If
                        Next
                        
                        '??????k??????????????????????(??????????????????????????)
                        colVsc.Add k, "_" & Mid(strKey, 2) '??????????????????
                        
                        '????K??????????(????????)
                        For j = 1 To X - 1
                            mshBody.TextMatrix(k, j) = mshBody.TextMatrix(k, j) & arrMerge(j)
                            arrLevel(j) = GetRowText(mshBody, k, j - 1)
                        Next
                        
                        rsVsc.MoveNext
                    Next
                    
                    '????????????????
                    k = mshBody.Rows
                    If strVscStat <> "" And k > Y + 1 Then
                        For j = X - 1 To 0 Step -1
                            If Split(strVscStat, ",")(j) <> "" Then
                                mshBody.AddItem "", k
                                mshBody.Row = k
                                For L = 0 To j - 1
                                    mshBody.TextMatrix(k, L) = mshBody.TextMatrix(k - 1, L)
                                Next
                                For L = j To X - 1
                                    mshBody.Col = L
                                    mshBody.CellAlignment = 4
                                    '????????????0,2??????,1????????,????0,2????????????????,????????????????????
                                    '????????????????????????????????????????????????
                                    'mshBody.TextMatrix(k, L) = Space(j Mod 2) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j Mod 2)
                                    mshBody.TextMatrix(k, L) = Space(j) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j)
                                Next
                                mshBody.RowData(k) = j + 1
                                mshBody.MergeRow(k) = True
                                
                                k = k + 1
                            End If
                        Next
                    End If
                End If
                
                '????????
                If Y > 0 And Not rsHsc Is Nothing Then
                    Set colHsc = New Collection
                    '????????
                    strHscStat = ""
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                        If tmpItem.???? = 8 Then strHscStat = strHscStat & "," & tmpItem.????
                    Next
                    strHscStat = Mid(strHscStat, 2)
                    
                    '????????
                    ReDim arrLevel(Y - 1) '??????????????????????????????????
                    ReDim arrMerge(Y - 1) '????????????????????????????????
                    For i = 1 To Y - 1
                        arrMerge(i) = Space(i Mod 2)
                    Next
                    L = lngCurCols + IIF(lngGrid = 0, X, 0) - Z    '????????????????
                    For i = 1 To rsHsc.RecordCount
                        L = L + Z
                        If mshBody.Cols - 1 < L Then mshBody.Cols = mshBody.Cols + Z
                        strKey = "" '????????????????(??????????????),??????????????????
                        For j = 0 To Y - 1
                            For k = 0 To Z - 1
                                Select Case rsHsc.Fields(CStr(Split(strHsc, "|")(j))).type
                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = "#" Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, L + k) = " " '????????????????????
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "??")
                                            mshBody.TextMatrix(j, L + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = 123456707654321# Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, L + k) = " "
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "??")
                                            mshBody.TextMatrix(j, L + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = #5/5/3000# Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, L + k) = " "
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "??")
                                            mshBody.TextMatrix(j, L + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                End Select
                            Next
                        Next
                        
                        '??????????
                        For j = Y - 1 To 1 Step -1 '??????????
                            strTmp = GetColText(mshBody, j - 1, L)
                            If strTmp <> arrLevel(j) And L > lngCurCols + IIF(lngGrid = 0, X, 0) Then
                                If strHscStat <> "" Then
                                    If Split(strHscStat, ",")(j) <> "" Then
                                        AddCol mshBody, L, Z
                                        For k = 0 To Z - 1
                                            For M = 0 To j - 1
                                                mshBody.TextMatrix(M, L + k) = mshBody.TextMatrix(M, L + k - Z)
                                            Next
                                            mshBody.Col = L + k
                                            mshBody.Row = j
                                            mshBody.CellAlignment = 4
                                            mshBody.TextMatrix(j, L + k) = Space((j + 1) Mod 2) & GetStatText(CStr(Split(strHscStat, ",")(j))) & Space((j + 1) Mod 2)
                                            mshBody.ColData(L + k) = j + 1
                                            mshBody.MergeCol(L + k) = True
                                        Next
                                        L = L + Z
                                    End If
                                End If
                                arrMerge(j) = IIF(arrMerge(j) = "", " ", "")
                            End If
                        Next
                        
                        '??????L??????????????????????(??????????????????????????)
                        colHsc.Add L, "_" & Mid(strKey, 2) '??????????????????
                        
                        '????L??????????(??????????)
                        For j = 1 To Y - 1
                            For k = 0 To Z - 1
                                mshBody.TextMatrix(j, L + k) = mshBody.TextMatrix(j, L + k) & arrMerge(j)
                            Next
                            arrLevel(j) = GetColText(mshBody, j - 1, L)
                        Next
                        rsHsc.MoveNext
                    Next
                    '????????????????
                    L = mshBody.Cols
                    If strHscStat <> "" And L > lngCurCols + IIF(lngGrid = 0, X, 0) Then
                        For j = Y - 1 To 0 Step -1
                            If Split(strHscStat, ",")(j) <> "" Then
                                AddCol mshBody, L, Z
                                For k = 0 To Z - 1
                                    For M = 0 To j - 1
                                        mshBody.TextMatrix(M, L + k) = mshBody.TextMatrix(M, L + k - Z)
                                    Next
                                    mshBody.Col = L + k
                                    mshBody.Row = j
                                    mshBody.CellAlignment = 4
                                    mshBody.TextMatrix(j, L + k) = Space((j + 1) Mod 2) & GetStatText(CStr(Split(strHscStat, ",")(j))) & Space((j + 1) Mod 2)
                                    mshBody.ColData(L + k) = j + 1
                                    mshBody.MergeCol(L + k) = True
                                Next
                                L = L + Z
                            End If
                        Next
                    End If
                End If
                
                '??????????????
                strFormat = ""
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                    If tmpItem.???? = 9 Then
                        strFormat = strFormat & "|~" & tmpItem.???? & "~"
                        For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                            '??????????????????????????????????,??????????????????????????
                            For j = Y To mshBody.FixedRows - 1
                                mshBody.TextMatrix(j, tmpItem.???? + i) = Space((tmpItem.???? + i) Mod 2) & tmpItem.???? & Space((tmpItem.???? + i) Mod 2)
                            Next
                            '??????????
                            If mshBody.ColData(i) > 0 Then
                                For M = Y - 1 To mshBody.ColData(i) Step -1
                                    For k = 0 To Z - 1
                                        mshBody.TextMatrix(M, i + k) = mshBody.TextMatrix(M + 1, i + k)
                                    Next
                                Next
                            End If
                        Next
                    End If
                Next
                strFormat = Mid(strFormat, 2)
                
                '??????????????
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                    '??????????????????????
                    If tmpItem.???? = 9 Then strFormat = Replace(strFormat, "~" & tmpItem.???? & "~", tmpItem.????)
                Next
                
                '????(????????)??????
                strAlign = ""
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                    Select Case tmpItem.????
                        Case 7 '????????
                           If lngGrid = 0 Then mshBody.ColWidth(tmpItem.????) = tmpItem.W
                        Case 9 '??????
                            strAlign = strAlign & "," & tmpItem.????
                            For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                                mshBody.ColAlignment(i + tmpItem.????) = Switch(tmpItem.???? = 0, 1, tmpItem.???? = 1, 4, tmpItem.???? = 2, 7)
                                If mshBody.FixedRows - 1 >= 0 And mshBody.Rows - 1 >= 0 Then mshBody.Cell(flexcpAlignment, mshBody.FixedRows - 1, i + tmpItem.????, mshBody.Rows - 1, i + tmpItem.????) = mshBody.ColAlignment(i + tmpItem.????)
                                mshBody.ColWidth(i + tmpItem.????) = tmpItem.W
                            Next
                            
                            '????????????????
                            Set objStatusGridItem = tmpItem
                    End Select
                Next
                strAlign = Mid(strAlign, 2)
                
                '????????????
                rsGroup.MoveFirst
                For i = 1 To rsGroup.RecordCount
                    '????
                    strKey = ""
                    For j = 0 To UBound(Split(strVsc, "|"))
                        strKey = strKey & "^" & IIF(IsNull(rsGroup.Fields(CStr(Split(strVsc, "|")(j))).Value), "", Replace(Nvl(rsGroup.Fields(CStr(Split(strVsc, "|")(j))).Value, ""), " ", "??"))
                    Next
                    
                    '??????????????,????????????????????????,????????????,??????????????,??????????????
                    lngRow = 0
                    If lngGrid > 0 Then On Local Error Resume Next
                    lngRow = CLng(colVsc("_" & Mid(strKey, 2))) + lngDiff
                    On Error GoTo 0
                    If lngRow > 0 Then
                        '????
                        lngCol = lngCurCols + IIF(lngGrid = 0, X, 0)
                        If strHsc <> "" Then
                            strKey = ""
                            For j = 0 To UBound(Split(strHsc, "|"))
                                strKey = strKey & "^" & IIF(IsNull(rsGroup.Fields(CStr(Split(strHsc, "|")(j))).Value), "", Replace(rsGroup.Fields(CStr(Split(strHsc, "|")(j))).Value & "", " ", "??"))
                            Next
                            lngCol = CLng(colHsc("_" & Mid(strKey, 2)))
                        End If
                        
                        '????(????????????????)
                        For j = 0 To Z - 1
                            strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + j))
                            If Not IsNull(rsGroup.Fields(strTmp).Value) Then
                                '????????????
                                StrFmt = ""
                                If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(j))
                                If StrFmt <> "" Then
                                    On Local Error Resume Next
                                    Select Case rsGroup.Fields(strTmp).type
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Format(Val(Replace(mshBody.TextMatrix(lngRow, lngCol + j), ",", "")) + rsGroup.Fields(strTmp).Value, StrFmt)
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Format(Val(mshBody.TextMatrix(lngRow, lngCol + j)) + Val(rsGroup.Fields(strTmp).Value), StrFmt)
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            If mshBody.TextMatrix(lngRow, lngCol + j) = "" Then
                                                mshBody.TextMatrix(lngRow, lngCol + j) = Format(CDate(rsGroup.Fields(strTmp).Value), StrFmt)
                                            Else
                                                mshBody.TextMatrix(lngRow, lngCol + j) = Format(CDate(mshBody.TextMatrix(lngRow, lngCol + j)) + rsGroup.Fields(strTmp).Value, StrFmt)
                                            End If
                                    End Select
                                    On Local Error GoTo 0
                                Else
                                    Select Case rsGroup.Fields(strTmp).type
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Val(Replace(mshBody.TextMatrix(lngRow, lngCol + j), ",", "")) + rsGroup.Fields(strTmp).Value
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Val(mshBody.TextMatrix(lngRow, lngCol + j)) + Val(rsGroup.Fields(strTmp).Value)
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            If mshBody.TextMatrix(lngRow, lngCol + j) = "" Then
                                                mshBody.TextMatrix(lngRow, lngCol + j) = CDate(rsGroup.Fields(strTmp).Value)
                                            Else
                                                mshBody.TextMatrix(lngRow, lngCol + j) = CDate(mshBody.TextMatrix(lngRow, lngCol + j)) + rsGroup.Fields(strTmp).Value
                                            End If
                                    End Select
                                End If
                                '????????????????????
                                Select Case CByte(Split(strAlign, ",")(j))
                                    Case 0 '??????
                                        mshBody.TextMatrix(lngRow, lngCol + j) = mshBody.TextMatrix(lngRow, lngCol + j) & Space((lngRow + lngCol + j) Mod 2)
                                    Case 1 '??????
                                        mshBody.TextMatrix(lngRow, lngCol + j) = Space((lngRow + lngCol + j) Mod 2) & mshBody.TextMatrix(lngRow, lngCol + j) & Space((lngRow + lngCol + j) Mod 2)
                                    Case 2 '??????
                                        mshBody.TextMatrix(lngRow, lngCol + j) = Space((lngRow + lngCol + j) Mod 2) & mshBody.TextMatrix(lngRow, lngCol + j)
                                End Select
                                
                                '??????????????????????
                                If Not objStatusGridItem Is Nothing Then
                                    For k = 1 To objStatusGridItem.ColProtertys.count
                                        Set objColProp = objStatusGridItem.ColProtertys.Item(k)
                                        If InStr(objColProp.??????, objCurItem.???? & ".") > 0 Then
                                            varIFValue = GetStatGridData(mshBody.Index, objColProp.??????, lngRow, lngCol + j)
                                        Else
                                            varIFValue = objColProp.??????
                                        End If
                                        If lngCol + j = mshBody.FixedCols And objColProp.???????????? Then
                                            If CheckColProtertys(mshBody.TextMatrix(lngRow, lngCol + j), objColProp.????????, varIFValue) Then
                                                If objColProp.???????? <> vbWhite Then
                                                    mshBody.Cell(flexcpBackColor, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.????????
                                                End If
                                                If objColProp.???????? <> vbBlack Then
                                                    mshBody.Cell(flexcpForeColor, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.????????
                                                End If
                                                If objColProp.???????? Then
                                                    mshBody.Cell(flexcpFontBold, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.????????
                                                End If
                                            End If
                                        Else
                                            If CheckColProtertys(mshBody.TextMatrix(lngRow, lngCol + j), objColProp.????????, varIFValue) Then
                                                If objColProp.???????? <> vbWhite Then
                                                    mshBody.Cell(flexcpBackColor, lngRow, lngCol + j) = objColProp.????????
                                                End If
                                                If objColProp.???????? <> vbBlack Then
                                                    mshBody.Cell(flexcpForeColor, lngRow, lngCol + j) = objColProp.????????
                                                End If
                                                If objColProp.???????? Then
                                                    mshBody.Cell(flexcpFontBold, lngRow, lngCol + j) = objColProp.????????
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                                
                            End If
                        Next
                    End If
                    rsGroup.MoveNext
                Next
                
                '????????????????(????????)
                '??????????
                If strHsc <> "" And strHscStat <> "" Then
                    For L = UBound(Split(strHsc, "|")) To 0 Step -1
                        strStat = CStr(Split(strHscStat, ",")(L))
                        If strStat <> "" Then
                            ReDim arrStat(mshBody.FixedRows To mshBody.Rows - 1, Z - 1)  '????????????
                            ReDim arrCount(mshBody.FixedRows To mshBody.Rows - 1, Z - 1) '????????????????
                            blnDo = False
                            For j = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                                For i = mshBody.FixedRows To mshBody.Rows - 1 '??????????????????????,Y????,??FixedRows
                                    '??????????????
                                    If mshBody.ColData(j) = L + 1 Then
                                        For k = 0 To Z - 1
                                            If strStat = "AVG" Then
                                                strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + k))
                                                Select Case rsGroup.Fields(strTmp).type
                                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                        arrStat(i, k) = Val(arrStat(i, k) / arrCount(i, k))
                                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                        arrStat(i, k) = Val(arrStat(i, k) / arrCount(i, k))
                                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                        arrStat(i, k) = CDate(arrStat(i, k) / arrCount(i, k))
                                                End Select
                                            End If
                                            StrFmt = ""
                                            If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(k))
                                            If StrFmt <> "" Then
                                                On Local Error Resume Next
                                                Select Case Split(strAlign, ",")(k)
                                                    Case 0 '??
                                                        mshBody.TextMatrix(i, j + k) = Format(arrStat(i, k), StrFmt) & Space((i + j + k) Mod 2)
                                                    Case 1 '??
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & Format(arrStat(i, k), StrFmt) & Space((i + j + k) Mod 2)
                                                    Case 2 '??
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & Format(arrStat(i, k), StrFmt)
                                                End Select
                                                On Local Error GoTo 0
                                            Else
                                                Select Case Split(strAlign, ",")(k)
                                                    Case 0 '??
                                                        mshBody.TextMatrix(i, j + k) = arrStat(i, k) & Space((i + j + k) Mod 2)
                                                    Case 1 '??
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & arrStat(i, k) & Space((i + j + k) Mod 2)
                                                    Case 2 '??
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & arrStat(i, k)
                                                End Select
                                            End If
                                        Next
                                    '????????????
                                    ElseIf mshBody.ColData(j) = 0 Then
                                        For k = 0 To Z - 1
                                            If Trim(mshBody.TextMatrix(i, j + k)) <> "" Then
                                                strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + k))
                                                arrCount(i, k) = arrCount(i, k) + 1
                                                Select Case strStat
                                                    Case "SUM", "AVG"
                                                        Select Case rsGroup.Fields(strTmp).type
                                                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                                arrStat(i, k) = arrStat(i, k) + Val(Replace(Trim(mshBody.TextMatrix(i, j + k)), ",", ""))
                                                            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                                arrStat(i, k) = arrStat(i, k) + Val(Trim(mshBody.TextMatrix(i, j + k)))
                                                            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                                arrStat(i, k) = arrStat(i, k) + CDate(Trim(mshBody.TextMatrix(i, j + k)))
                                                        End Select
                                                    Case "MIN"
                                                        If Not blnDo Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k)): blnDo = True
                                                        If Trim(mshBody.TextMatrix(i, j + k)) < arrStat(i, k) Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k))
                                                    Case "MAX"
                                                        If Not blnDo Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k)): blnDo = True
                                                        If Trim(mshBody.TextMatrix(i, j + k)) > arrStat(i, k) Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k))
                                                    Case "COUNT"
                                                        arrStat(i, k) = arrStat(i, k) + 1
                                                End Select
                                            End If
                                        Next
                                    End If
                                Next
                                If mshBody.ColData(j) = L + 1 Then
                                    ReDim arrStat(mshBody.FixedRows To mshBody.Rows - 1, Z - 1)  '????????????
                                    ReDim arrCount(mshBody.FixedRows To mshBody.Rows - 1, Z - 1) '????????????????
                                    blnDo = False
                                End If
                            Next
                        End If
                    Next
                End If

                '??????????
                If strVscStat <> "" Then
                    For L = UBound(Split(strVsc, "|")) To 0 Step -1
                        strStat = CStr(Split(strVscStat, ",")(L))
                        If strStat <> "" Then
                            ReDim arrStat(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1) '????????????
                            ReDim arrCount(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1) '????????????????
                            blnDo = False
                            For i = mshBody.FixedRows To mshBody.Rows - 1 '??????????????????????,Y????,??FixedRows
                                For j = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1
                                    '??????????????
                                    If mshBody.RowData(i) = L + 1 Then
                                        If strStat = "AVG" Then
                                            strTmp = Trim(mshBody.TextMatrix(Y, j))
                                            Select Case rsGroup.Fields(strTmp).type
                                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                    arrStat(j) = Val(arrStat(j) / arrCount(j))
                                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                    arrStat(j) = Val(arrStat(j) / arrCount(j))
                                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                    arrStat(j) = CDate(arrStat(j) / arrCount(j))
                                            End Select
                                        End If
                                        k = 0
                                        If Z > 1 Then k = ((j - (lngCurCols + IIF(lngGrid = 0, X, 0)) + 1) Mod Z) - 1
                                        If k = -1 Then k = Z - 1
                                        StrFmt = ""
                                        If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(k))
                                        If StrFmt <> "" Then
                                            On Local Error Resume Next
                                            Select Case Split(strAlign, ",")(k)
                                                Case 0 '??
                                                    mshBody.TextMatrix(i, j) = Format(arrStat(j), StrFmt) & Space((i + j) Mod 2)
                                                Case 1 '??
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & Format(arrStat(j), StrFmt) & Space((i + j) Mod 2)
                                                Case 2 '??
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & Format(arrStat(j), StrFmt)
                                            End Select
                                            On Local Error GoTo 0
                                        Else
                                            Select Case Split(strAlign, ",")(k)
                                                Case 0 '??
                                                    mshBody.TextMatrix(i, j) = arrStat(j) & Space((i + j) Mod 2)
                                                Case 1 '??
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & arrStat(j) & Space((i + j) Mod 2)
                                                Case 2 '??
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & arrStat(j)
                                            End Select
                                        End If
                                    '????????????
                                    ElseIf mshBody.RowData(i) = 0 And Trim(mshBody.TextMatrix(i, j)) <> "" Then
                                        strTmp = Trim(mshBody.TextMatrix(Y, j))
                                        arrCount(j) = arrCount(j) + 1
                                        Select Case strStat
                                            Case "SUM", "AVG"
                                                Select Case rsGroup.Fields(strTmp).type
                                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                        arrStat(j) = arrStat(j) + Val(Replace(Trim(mshBody.TextMatrix(i, j)), ",", ""))
                                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                        arrStat(j) = arrStat(j) + Val(Trim(mshBody.TextMatrix(i, j)))
                                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                        arrStat(j) = arrStat(j) + CDate(Trim(mshBody.TextMatrix(i, j)))
                                                End Select
                                            Case "MIN"
                                                If Not blnDo Then arrStat(j) = Trim(mshBody.TextMatrix(i, j)): blnDo = True
                                                If Trim(mshBody.TextMatrix(i, j)) < arrStat(j) Then arrStat(j) = Trim(mshBody.TextMatrix(i, j))
                                            Case "MAX"
                                                If Not blnDo Then arrStat(j) = Trim(mshBody.TextMatrix(i, j)): blnDo = True
                                                If Trim(mshBody.TextMatrix(i, j)) > arrStat(j) Then arrStat(j) = Trim(mshBody.TextMatrix(i, j))
                                            Case "COUNT"
                                                arrStat(j) = arrStat(j) + 1
                                        End Select
                                    End If
                                Next
                                If mshBody.RowData(i) = L + 1 Then
                                    ReDim arrStat(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1)
                                    ReDim arrCount(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1)
                                    blnDo = False
                                End If
                            Next
                        End If
                    Next
                End If
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                    Select Case tmpItem.????
                        Case 7 '????????
                            '??????????????????????
                            If tmpItem.Relations.count > 0 Then
                                mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.????, mshBody.Rows - 1, tmpItem.????) = &HFF0001
                                mshBody.Cell(flexcpFontUnderline, mshBody.FixedRows, tmpItem.????, mshBody.Rows - 1, tmpItem.????) = True
                                mshBody.Cell(flexcpData, mshBody.FixedRows, tmpItem.????, mshBody.Rows - 1, tmpItem.????) = tmpItem
                            End If
                            '??????????????????????????????????
                            mshBody.Cell(flexcpFontBold, mshBody.FixedRows, tmpItem.????, mshBody.Rows - 1, tmpItem.????) = tmpItem.????
                            If tmpItem.???? <> 0 Then mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.????, mshBody.Rows - 1, tmpItem.????) = tmpItem.????
                        Case 8 '????????
                            If tmpItem.Relations.count > 0 Then
                                mshBody.Cell(flexcpForeColor, tmpItem.????, mshBody.FixedCols, tmpItem.????, mshBody.Cols - 1) = &HFF0001
                                mshBody.Cell(flexcpFontUnderline, tmpItem.????, mshBody.FixedCols, tmpItem.????, mshBody.Cols - 1) = True
                                mshBody.Cell(flexcpData, tmpItem.????, mshBody.FixedCols, tmpItem.????, mshBody.Cols - 1) = tmpItem
                            End If
                            '??????????????????????????????????
                            mshBody.Cell(flexcpFontBold, tmpItem.????, mshBody.FixedCols, tmpItem.????, mshBody.Cols - 1) = tmpItem.????
                            If tmpItem.???? <> 0 Then mshBody.Cell(flexcpForeColor, tmpItem.????, mshBody.FixedCols, tmpItem.????, mshBody.Cols - 1) = tmpItem.????
                        Case 9 '??????
                            For j = mshBody.FixedCols To mshBody.Cols - 1 Step lngStatistics
                                On Error Resume Next
                                If tmpItem.Relations.count > 0 Then
                                    mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.???? + j, mshBody.Rows - 1, tmpItem.???? + j) = &HFF0001
                                    mshBody.Cell(flexcpFontUnderline, mshBody.FixedRows, tmpItem.???? + j, mshBody.Rows - 1, tmpItem.???? + j) = True
                                End If
                                '??????????????????????????????
                                mshBody.Cell(flexcpData, mshBody.FixedRows, tmpItem.???? + j, mshBody.FixedRows, tmpItem.???? + j) = tmpItem
                            
'                                 '??????????????????????????????????
'                                mshBody.Cell(flexcpFontBold, mshBody.FixedRows, tmpItem.???? + j, mshBody.Rows - 1, tmpItem.???? + j) = tmpItem.????
'                                If tmpItem.???? <> 0 Then mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.???? + j, mshBody.Rows - 1, tmpItem.???? + j) = tmpItem.????
                                On Error GoTo 0
                            Next
                    End Select
                Next
                '??????????????????
                For i = 0 To mshBody.FixedCols - 1
                    For j = 0 To mshBody.Rows - 1
                        If Decode(Trim(mshBody.TextMatrix(j, i)), "????", 1, "??????", 2, "??????", 3, "??????", 4, "??????", 5, 0) > 0 Then
                            mshBody.Cell(flexcpForeColor, j, i, j, mshBody.Cols - 1) = mshBody.ForeColor
                            mshBody.Cell(flexcpFontUnderline, j, i, j, mshBody.Cols - 1) = False
                            mshBody.Cell(flexcpData, j, i, j, mshBody.Cols - 1) = Empty
                            mshBody.Cell(flexcpFontBold, j, i, j, mshBody.Cols - 1) = False
                        End If
                    Next
                Next
                For j = 0 To mshBody.FixedRows - 1
                    For i = 0 To mshBody.Cols - 1
                        If Decode(Trim(mshBody.TextMatrix(j, i)), "????", 1, "??????", 2, "??????", 3, "??????", 4, "??????", 5, 0) > 0 Then
                            mshBody.Cell(flexcpForeColor, j, i, mshBody.Rows - 1, i) = mshBody.ForeColor
                            mshBody.Cell(flexcpFontUnderline, j, i, mshBody.Rows - 1, i) = False
                            mshBody.Cell(flexcpData, j, i, mshBody.Rows - 1, i) = Empty
                            mshBody.Cell(flexcpFontBold, j, i, mshBody.Rows - 1, i) = False
                        End If
                    Next
                Next
                
                '??????????????????
                If Z = 1 And Y > 0 Then
                    For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1
                        For j = mshBody.FixedRows - 1 To Y Step -1
                            mshBody.TextMatrix(j, i) = mshBody.TextMatrix(Y - 1, i)
                        Next
                    Next
                Else
                    blnHide = False
                End If
                
                '??????????"????"????????????????????????,????"????????"????
                objItem.???? = objItem.???? & "|" & .ID & "," & mshBody.Cols
                strTopRow = strTopRow & "|" & objCurItem.???? & "," & mshBody.Cols
                                
                '??????????????????????
                lngCurCols = mshBody.Cols
            Else
                blnHide = False
                Call SetHeadCenter(mshBody)
                Exit For '??????????????????,??????????????????????????
            End If
        End With
    Next
    
    objItem.???? = Mid(objItem.????, 2)
    
    '????????????,????????(??????????,??SQL????)
'    strTopRow = Mid(strTopRow, 2)
'    strTmp = ""
'    For i = 0 To UBound(Split(strTopRow, "|"))
'        If InStr(strTmp & "|", "|" & Split(Split(strTopRow, "|")(i), ",")(0) & "|") = 0 Then
'            strTmp = strTmp & "|" & Split(Split(strTopRow, "|")(i), ",")(0)
'        End If
'    Next
'    If UBound(Split(Mid(strTmp, 2), "|")) > 0 Then
'        '??????
'        mshBody.AddItem "", mshBody.FixedRows
'        mshBody.FixedRows = mshBody.FixedRows + 1
'        For i = mshBody.FixedRows - 1 To 1 Step -1
'            For j = 0 To mshBody.Cols - 1
'                mshBody.TextMatrix(i, j) = mshBody.TextMatrix(i - 1, j)
'                mshBody.RowHeight(i) = mshBody.RowHeight(i - 1)
'                mshBody.RowData(i) = mshBody.RowData(i - 1)
'            Next
'        Next
'        mshBody.RowData(0) = 0
'        mshBody.RowHeight(0) = objItem.????
'        mshBody.MergeRow(0) = True
'        For j = mshBody.FixedCols To mshBody.Cols - 1
'            mshBody.TextMatrix(0, j) = ""
'        Next
'
'        '????????
'        For i = 0 To UBound(Split(strTopRow, "|"))
'            If i = 0 Then
'                lngColB = mshBody.FixedCols
'            Else
'                lngColB = lngColE + 1
'            End If
'            lngColE = CLng(Split(Split(strTopRow, "|")(i), ",")(1)) - 1
'            For j = lngColB To lngColE
'                mshBody.TextMatrix(0, j) = CStr(Split(Split(strTopRow, "|")(i), ",")(0))
'            Next
'        Next
'    End If
    
    '????????????
    For j = 0 To mshBody.Cols - 1
        mshBody.MergeCol(j) = True
    Next
    For i = 0 To mshBody.FixedRows - 2
        mshBody.MergeRow(i) = True
    Next
    
    '????(????????)
    For i = 0 To mshBody.Rows - 1
        mshBody.RowHeight(i) = objItem.????
    Next
    
    '????????????????????
    '------??????????????-----------------
    blnHide = True
    For i = mshBody.FixedRows - 1 To 1 Step -1
        For j = 0 To mshBody.Cols - 1
            If mshBody.TextMatrix(i, j) <> mshBody.TextMatrix(i - 1, j) Then
                blnHide = False: Exit For
            End If
        Next
        If blnHide Then
            mshBody.RowHeight(i) = 0
        Else
            Exit For
        End If
    Next
    '------??????????????-----------------
    'If blnHide Then mshBody.RowHeight(mshBody.FixedRows - 1) = 0
    
    '????????????
    mshBody.Cell(flexcpAlignment, 0, 0, mshBody.FixedRows - 1, mshBody.Cols - 1) = flexAlignCenterCenter
    mshBody.Cell(flexcpAlignment, 0, 0, mshBody.Rows - 1, mshBody.FixedCols - 1) = flexAlignCenterCenter
    
    '????????????(????????)
    For i = mshBody.FixedRows To mshBody.Rows - 1
        If mshBody.RowData(i) = 0 Then
            mshBody.Row = i
            For j = 0 To mshBody.FixedCols - 1
                mshBody.Col = j
                mshBody.CellAlignment = 1
            Next
        End If
    Next
    
    mshBody.WordWrap = True
    
    mshBody.MergeCells = flexMergeFree
    mshBody.ScrollBars = flexScrollBarBoth
    mshBody.Row = mshBody.FixedRows
    mshBody.Col = mshBody.FixedCols
    mshBody.Redraw = flexRDBuffered
    mshBody.ZOrder
    mshBody.Visible = True
End Sub

Private Function GetGridColWidth(objGrid As Object) As Long
'??????????????????????????????
    Dim i As Integer, lngW As Long
    For i = 0 To objGrid.Cols - 1
        lngW = lngW + objGrid.ColWidth(i)
    Next
    GetGridColWidth = lngW
End Function

Private Sub SetGridAlign()
'????????????????????????????(??????????????????)
'????????????????????,????????????????,??????????????????????
    Dim tmpMsh As VSFlexGrid, tmpItem As RPTItem
    Dim lngMaxW As Long, lngCurW As Long, sngScale As Single
    Dim strIDs As String, intCurID As Integer
    Dim j As Integer, i As Integer
    
    If Not mobjReport.blnLoad Then Exit Sub

    timHead.Enabled = False
    
    For Each tmpItem In mobjReport.Items
        If tmpItem.?????? = bytFormat And (tmpItem.???? = 4 Or tmpItem.???? = 5) _
            And tmpItem.???? = "" And tmpItem.???? = 0 Then
            
            If GridHaveApp(tmpItem.ID) Then
                strIDs = GetGridAppIDs(tmpItem.????)
                strIDs = tmpItem.ID & "," & strIDs
                '????????????(??????????????)
                lngMaxW = -1
                For i = 0 To UBound(Split(strIDs, ","))
                    intCurID = CInt(Split(strIDs, ",")(i))
                    Set tmpMsh = msh(intCurID)
                    lngCurW = GetGridColWidth(tmpMsh)
                    If lngCurW <= mobjReport.Items("_" & intCurID).W Then
                        If lngMaxW = -1 Then
                            lngMaxW = lngCurW
                        ElseIf lngCurW > lngMaxW Then
                            lngMaxW = lngCurW
                        End If
                    End If
                Next
                '??????????????
                If lngMaxW <> -1 Then
                    For i = 0 To UBound(Split(strIDs, ","))
                        intCurID = CInt(Split(strIDs, ",")(i))
                        If mobjReport.Items("_" & intCurID).???? = 4 Then
                             '??????????????????
                            Set tmpMsh = msh(CInt(msh(intCurID).Tag))
                        ElseIf mobjReport.Items("_" & intCurID).???? = 5 Then
                            Set tmpMsh = msh(intCurID)
                        End If
                        tmpMsh.Redraw = False
                        lngCurW = GetGridColWidth(tmpMsh)
                        If lngCurW <= mobjReport.Items("_" & intCurID).W And lngCurW <> 0 Then
                            sngScale = lngMaxW / lngCurW
                            For j = 0 To tmpMsh.Cols - 1
                                tmpMsh.ColWidth(j) = tmpMsh.ColWidth(j) * sngScale
                            Next
                            tmpMsh.ColWidth(tmpMsh.Cols - 1) = tmpMsh.ColWidth(tmpMsh.Cols - 1) + lngMaxW - GetGridColWidth(tmpMsh)
                        End If
                        tmpMsh.Redraw = True
                    Next
                End If
            End If
        End If
    Next
    Call timHead_Timer
    timHead.Enabled = True
End Sub

Private Function GetGridAppIDs(strName As String) As String
'????????????????????????????????????
'??????strName=??????
    Dim tmpItem As RPTItem
    Dim strIDs As String
    
    For Each tmpItem In mobjReport.Items
        If tmpItem.?????? = bytFormat And tmpItem.???? = 4 _
            And tmpItem.???? = 1 And tmpItem.???? = strName Then
            strIDs = strIDs & "," & tmpItem.ID
        End If
    Next
    GetGridAppIDs = Mid(strIDs, 2)
End Function

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txt(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
        '????????????
        If txt(Index).Text <> "" Then
            If cmd(Index).Enabled And cmd(Index).Visible Then
                blnMatch = True
                Call cmd_Click(Index)
            End If
            Cancel = True
        End If
    End If
End Sub

Private Sub ReplaceSysNo(objReport As Report)
    Dim i As Integer, j As Integer
    For i = 1 To objReport.Datas.count
        objReport.Datas(i).SQL = Replace(objReport.Datas(i).SQL, "[????]", IIF(glngSys <> 0, glngSys, objReport.????))
        For j = 1 To objReport.Datas(i).Pars.count
            objReport.Datas(i).Pars(j).????SQL = Replace(objReport.Datas(i).Pars(j).????SQL, "[????]", IIF(glngSys <> 0, glngSys, objReport.????))
            objReport.Datas(i).Pars(j).????SQL = Replace(objReport.Datas(i).Pars(j).????SQL, "[????]", IIF(glngSys <> 0, glngSys, objReport.????))
        Next
    Next
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '????:??????????????
    '??????:??????
    '????????:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Function GetStatGridData(ByVal intIndex As Integer, ByVal strFiled As String, ByVal lngRow As Long, ByVal lngCol As Long)
'????????????????????????????????????????????????
    Dim str???????? As String
    Dim i As Long, strFiledTmp As String
    
    With msh(intIndex)
        strFiledTmp = Mid(strFiled, InStr(strFiled, ".") + 1)
        str???????? = .TextMatrix(0, lngCol)
        
        For i = lngCol To .Cols - 1
            If str???????? = .TextMatrix(0, i) And strFiledTmp = Trim(.TextMatrix(.FixedRows - 1, i)) Then
                GetStatGridData = .TextMatrix(lngRow, i)
                Exit Function
            End If
        Next
        
        For i = lngCol To .FixedCols Step -1
            If str???????? = .TextMatrix(0, i) And strFiledTmp = Trim(.TextMatrix(.FixedRows - 1, i)) Then
                GetStatGridData = .TextMatrix(lngRow, i)
                Exit Function
            End If
        Next
        GetStatGridData = strFiled
    End With
End Function

Private Sub FindItem(ByVal strFind As String, Optional ByVal blnNext As Boolean)
'????????????????????????
'??????blnNext=??????????
    Static lngindex As Long
    Static lngMshRow As Long
    Static lngMshcol As Long
    Static strFindLast As String
    
    Dim objControl As Object
    Dim blnTmp As Boolean
    Dim i As Long, j As Long, k As Long
    
    If Trim(strFind) = "" Then Exit Sub
    If strFindLast <> strFind Then lngindex = 0
    strFindLast = strFind
    If lngCurInx <> 0 And lbl(lngCurInx).BackColor = CON_SETFOCES Then lbl(lngCurInx).BackColor = lngTmpColor
    For Each objControl In Me.Controls
        i = i + 1
        '????????????????
        If i >= lngindex Then
            If objControl.name = "lbl" Then
                If i > lngindex Then
                    If objControl.Caption Like "*" & strFind & "*" Then
                        lngCurInx = objControl.Index
                        lngTmpColor = objControl.BackColor
                        objControl.BackColor = CON_SETFOCES
                        lngindex = i
                        blnTmp = True
                        Exit Sub
                    End If
                End If
            ElseIf objControl.name = "msh" Then
                If lngindex <> i Then lngMshRow = 0: lngMshcol = 0
                If lngMshRow < objControl.Rows - 1 Or lngMshcol < objControl.Cols - 1 Then
                    For j = objControl.FixedRows To objControl.Rows - 1
                        For k = objControl.FixedCols To objControl.Cols - 1
                            If j = lngMshRow And k > lngMshcol Or j > lngMshRow Then
                                If objControl.TextMatrix(j, k) Like "*" & strFind & "*" Then
                                    objControl.Row = j: objControl.Col = k
                                    objControl.ShowCell j, k
                                    lngindex = i
                                    blnTmp = True
                                    lngMshRow = j: lngMshcol = k
                                    objControl.SetFocus
                                    Exit Sub
                                End If
                            End If
                            
                        Next
                    Next
                End If
            End If
        End If
    Next
    If blnTmp = False Then
        If lngindex <> 0 Then
            MsgBox "??????????????????", vbInformation, App.Title
        Else
            MsgBox "??????????????????????", vbInformation, App.Title
        End If
        lngindex = 0
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call FindItem(txtFind.Text)
    End If
End Sub

Public Function GetReportForm(objParent As Object, objCurDLL As clsReport, LibDatas As Object, arrPars As Variant, ByVal bytStyle As Byte) As Object
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    
    On Error Resume Next
        Load Me
    If Err.Number = 0 Then
        Set mobjfrmShow = New frmPreview
        If Not mobjReport.blnLoad Then Exit Function
    
        If mobjReport.Items.count = 0 Then Exit Function
        
        If Not InitPrinter(Me) Then
            gblnError = True
            MsgBox "??????????????.????????????????????????????????????????????", vbInformation, App.Title: Exit Function
        End If
        
        If Not CalcCellPage Then
            gblnError = True
            MsgBox "??????????????????,??????????????", vbInformation, App.Title: Exit Function
        End If
        If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
            lbl(lngCurInx).BackColor = lngTmpColor
            lngCurInx = 0: lngTmpColor = 0
        End If
        mobjfrmShow.BorderStyle = FormBorderStyleConstants.vbBSNone '????????????
        mobjfrmShow.Caption = mobjfrmShow.Caption       '????????????
        Set mobjfrmShow.frmParent = Me
        Load mobjfrmShow
        mobjfrmShow.LoadForm 1
        Set LibDatas = mLibDatas
        Set GetReportForm = mobjfrmShow
    ElseIf Err.Number <> 0 Then
        '364:??????????(??Form_Load????Unload,??????????????)
        Err.Clear
    End If
End Function

Public Sub PrintReportForRec(objParent As Object, objCurDLL As clsReport, LibDatas As Object, arrPars As Variant, ByVal bytStyle As Byte)
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    
    On Error Resume Next
    
    If mbytStyle <> 0 Then
        Set mLibDatas = LibDatas
        Load Me
        If Err.Number = 0 Then
            If mbytStyle = 1 Then       '????????
                mnuFile_Preview_Click
            ElseIf mbytStyle = 2 Then   '????????
                mnuFile_Print_Click
            ElseIf mbytStyle = 3 Then   '??????Excel
                mnuFile_Excel_Click
            ElseIf mbytStyle = 4 Then   '??????????PDF
                mnuFile_Print_Click
            End If
        ElseIf Err.Number <> 0 Then
            '364:??????????(??Form_Load????Unload,??????????????)
            Err.Clear
        End If
        Unload Me
    Else
        '??????????????????????
        If frmParent Is Nothing Then
            Me.Show
        ElseIf frmParent.name = "frmDesign" Then
            Me.Show 1, frmParent
        Else
            Me.Show , frmParent
        End If
        
        '????????????????????????
        If Err.Number = 373 Or Err.Number = 401 Then
            '373:??????????????????????????????????(??????????zlReport.dll,??????????????)
            '401:????????????????????????????????????
            '??????Load????????????????????Form_Load????
            Err.Clear: Me.Show 1
        ElseIf Err.Number = 364 Then
            '364:??????????(??Form_Load????Unload,??????????????)
            Err.Clear
        ElseIf Err.Number <> 0 Then
            Err.Clear: Unload Me '??????Load????????????????????
        End If
    End If
End Sub

Private Sub LoadCondsMenu()
    Dim strSQL As String
    Dim i As Integer
    Dim rsPara As ADODB.Recordset
    Dim blnRetry As Boolean
    
    If mlngRPTID = 0 Then Exit Sub
    
    On Error GoTo hErr
    
    '????????????
    For i = mnuPop_Cond.count - 1 To 1 Step -1
        Unload mnuPop_Cond(i)
    Next
    
    blnRetry = True
    strSQL = "Select Distinct ??????, ???????? From zlRptConds Where ????ID=[1] Order by ??????"
    Set rsPara = OpenSQLRecord(strSQL, "??????????????????????", mlngRPTID)
    blnRetry = False
    
    With rsPara
        If .RecordCount = 0 Then
            mnuPop_Split1.Visible = False
            mnuPop_Del.Enabled = False
            mintCurCondID = 0
            mintCurMenuIndex = 0
        Else
            mnuPop_Split1.Visible = True
            mnuPop_Del.Enabled = mintCurCondID > 0
            Do While .EOF = False
                i = .AbsolutePosition
                Load mnuPop_Cond(i)
                mnuPop_Cond(i).Caption = Nvl(!????????) & "(&" & i & ")"
                mnuPop_Cond(i).Visible = True
                mnuPop_Cond(i).Tag = Nvl(!??????, 0)
                
                If mintCurCondID = Nvl(!??????, 0) Then
                    mnuPop_Cond(i).Checked = True
                Else
                    mnuPop_Cond(i).Checked = False
                End If
                
                .MoveNext
            Loop
        End If
        .Close
    End With
            
    mnuPop_Default.Checked = mintCurCondID = 0
    
    Exit Sub
    
hErr:
    If blnRetry Then
        If ErrCenter = 1 Then Resume
    Else
        Call ErrCenter
    End If
End Sub
