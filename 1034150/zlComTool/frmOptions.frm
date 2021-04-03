VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "系统选项"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5775
      Index           =   1
      Left            =   180
      TabIndex        =   27
      Top             =   540
      Width           =   5565
      Begin VB.PictureBox picPreview 
         Height          =   3375
         Left            =   630
         ScaleHeight     =   221
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   269
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1710
         Width           =   4095
      End
      Begin VB.PictureBox picBorder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3765
         Left            =   420
         ScaleHeight     =   251
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   301
         TabIndex        =   33
         Top             =   1500
         Width           =   4515
      End
      Begin VB.PictureBox pic风格 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   450
         ScaleHeight     =   480
         ScaleWidth      =   4485
         TabIndex        =   34
         Top             =   900
         Width           =   4485
         Begin VB.OptionButton optStyle 
            Caption         =   "Windows风格"
            Height          =   195
            Index           =   1
            Left            =   1515
            TabIndex        =   2
            Tag             =   "zlwin"
            Top             =   60
            Width           =   1305
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "传统风格"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   1
            Tag             =   "zlBrw"
            Top             =   60
            Width           =   1275
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "MDI风格"
            Height          =   195
            Index           =   2
            Left            =   3330
            TabIndex        =   3
            Tag             =   "zlmdi"
            Top             =   60
            Width           =   945
         End
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "    按照个人的兴趣爱好选择自己喜欢的导航风格，使你心情更愉快、工作更轻松。"
         Height          =   480
         Left            =   870
         TabIndex        =   26
         Top             =   120
         Width           =   4140
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5775
      Index           =   3
      Left            =   180
      TabIndex        =   35
      Top             =   540
      Width           =   5565
      Begin VB.PictureBox pic简码 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   1485
         ScaleHeight     =   630
         ScaleWidth      =   3345
         TabIndex        =   50
         Top             =   2460
         Width           =   3345
         Begin VB.OptionButton opt简码 
            Caption         =   "拼音，取每字的首字母构成简码(&P)"
            Height          =   210
            Index           =   0
            Left            =   15
            TabIndex        =   17
            Top             =   90
            Value           =   -1  'True
            Width           =   3150
         End
         Begin VB.OptionButton opt简码 
            Caption         =   "五笔，取每字的首字母构成简码(&W)"
            Height          =   210
            Index           =   1
            Left            =   15
            TabIndex        =   18
            Top             =   330
            Width           =   3150
         End
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   0
         Left            =   1950
         TabIndex        =   49
         Top             =   645
         Width           =   2835
      End
      Begin VB.ComboBox cmbIME 
         Height          =   300
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1770
         Width           =   3015
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   1
         Left            =   1365
         TabIndex        =   48
         Top             =   1455
         Width           =   3420
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   2
         Left            =   1560
         TabIndex        =   47
         Top             =   2265
         Width           =   3225
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   3
         Left            =   1800
         TabIndex        =   46
         Top             =   3900
         Width           =   3015
      End
      Begin VB.CheckBox chkAutoStart 
         Caption         =   "在 Windows 启动时自动运行(&A)"
         Height          =   210
         Left            =   1530
         TabIndex        =   20
         Top             =   4185
         Width           =   2865
      End
      Begin VB.CheckBox chkShutDown 
         Caption         =   "退出程序时自动关闭 Windows (&S)"
         Height          =   210
         Left            =   1530
         TabIndex        =   21
         Top             =   4440
         Width           =   3045
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   4
         Left            =   1950
         TabIndex        =   45
         Top             =   4800
         Width           =   2880
      End
      Begin VB.TextBox txtTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1770
         MaxLength       =   4
         TabIndex        =   22
         Text            =   "30"
         ToolTipText     =   "检查周期最少10秒，最多3600秒"
         Top             =   5250
         Width           =   540
      End
      Begin VB.PictureBox pic2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Left            =   255
         Picture         =   "frmOptions.frx":044E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   0
         Left            =   645
         Picture         =   "frmOptions.frx":1118
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   915
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   1
         Left            =   645
         Picture         =   "frmOptions.frx":2A9A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1650
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   2
         Left            =   615
         Picture         =   "frmOptions.frx":441C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2475
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   3
         Left            =   630
         Picture         =   "frmOptions.frx":5D9E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4125
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   4
         Left            =   630
         Picture         =   "frmOptions.frx":7720
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   5100
         Width           =   480
      End
      Begin VB.PictureBox pic匹配方式 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1455
         ScaleHeight     =   315
         ScaleWidth      =   3495
         TabIndex        =   38
         Top             =   960
         Width           =   3495
         Begin VB.OptionButton opt 
            Caption         =   "从左匹配(&L)"
            Height          =   210
            Index           =   1
            Left            =   1815
            TabIndex        =   15
            Top             =   60
            Width           =   1320
         End
         Begin VB.OptionButton opt 
            Caption         =   "双向匹配(&D)"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   14
            Top             =   60
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   5
         Left            =   600
         Picture         =   "frmOptions.frx":A112
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3330
         Width           =   480
      End
      Begin VB.CheckBox chkAutoHide 
         Caption         =   "允许界面区域提供自动隐藏功能(&I)"
         Height          =   195
         Left            =   1515
         TabIndex        =   19
         Top             =   3405
         Width           =   3090
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   5
         Left            =   1560
         TabIndex        =   36
         Top             =   3105
         Width           =   3225
      End
      Begin VB.Label lblTime 
         Caption         =   "每       秒检查消息通知"
         Height          =   255
         Left            =   1560
         TabIndex        =   58
         Top             =   5250
         Width           =   2145
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":A9DC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "    用户可以根据自身的习惯来选择输入的匹配方式、简码类型、输入法等，以提高工作效率"
         Height          =   480
         Left            =   870
         TabIndex        =   57
         Top             =   120
         Width           =   4335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目输入匹配方式"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   56
         Top             =   645
         Width           =   1440
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "汉字输入法"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   55
         Top             =   1470
         Width           =   900
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "简码处理方式"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   54
         Top             =   2265
         Width           =   1080
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "与Windows的结合"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   53
         Top             =   3900
         Width           =   1350
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "消息通知检查周期"
         Height          =   180
         Index           =   4
         Left            =   300
         TabIndex        =   52
         Top             =   4800
         Width           =   1440
      End
      Begin VB.Line lineTime 
         X1              =   1740
         X2              =   2355
         Y1              =   5445
         Y2              =   5445
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "界面区域隐藏"
         Height          =   180
         Index           =   5
         Left            =   300
         TabIndex        =   51
         Top             =   3105
         Width           =   1080
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5775
      Index           =   2
      Left            =   180
      TabIndex        =   28
      Top             =   540
      Width           =   5565
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   3
         Left            =   4890
         Picture         =   "frmOptions.frx":AFF1
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "向后移动"
         Top             =   5160
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   2
         Left            =   4890
         Picture         =   "frmOptions.frx":B13E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "向前移动"
         Top             =   4680
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   1
         Left            =   4890
         Picture         =   "frmOptions.frx":B28B
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "删除程序"
         Top             =   3210
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   0
         Left            =   4890
         Picture         =   "frmOptions.frx":B331
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "新增程序"
         Top             =   2730
         Width           =   345
      End
      Begin MSComctlLib.ListView lvwFavorite 
         Height          =   2805
         Left            =   3000
         TabIndex        =   9
         Top             =   2700
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   4948
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "程序"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "生成快捷方式"
         Height          =   1335
         Left            =   3030
         TabIndex        =   29
         Top             =   750
         Width           =   2265
         Begin VB.CommandButton cmdStartup 
            Caption         =   "到启动菜单(&S)"
            Height          =   350
            Left            =   270
            TabIndex        =   8
            Top             =   840
            Width           =   1725
         End
         Begin VB.CommandButton cmdDesktop 
            Caption         =   "到桌面(&D)"
            Height          =   350
            Left            =   270
            TabIndex        =   7
            Top             =   360
            Width           =   1725
         End
      End
      Begin VB.ComboBox cboGroup 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   750
         Width           =   1620
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   2160
         Top             =   2910
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView tvwMain 
         Height          =   4410
         Left            =   300
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1140
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   7779
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
      Begin VB.Label lblFavorite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "常用程序列表(&R)"
         Height          =   180
         Left            =   3030
         TabIndex        =   32
         Top             =   2430
         Width           =   1350
      End
      Begin VB.Label lblNote 
         Caption         =   $"frmOptions.frx":B3DE
         Height          =   570
         Left            =   900
         TabIndex        =   31
         Top             =   30
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "菜单体系"
         Height          =   180
         Left            =   360
         TabIndex        =   30
         Top             =   810
         Width           =   720
      End
      Begin VB.Image imgNote 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":B466
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   390
      TabIndex        =   25
      Top             =   6510
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   6255
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   11033
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "导航风格"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "菜单选择"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "使用习惯"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4620
      TabIndex        =   24
      Top             =   6510
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3300
      TabIndex        =   23
      Top             =   6510
      Width           =   1100
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum constOpt
    opt_双向匹配 = 0
    opt_从左匹配 = 1
End Enum

Private Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Private Declare Function OSfRemoveShellLink Lib "vb6stkit.dll" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long
Dim mintIndex As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDesktop_Click()
    Dim strPath As String
    Dim StrName As String
    
    '创建快捷方式
    strPath = GetSetting("ZLSOFT", "公共全局", "程序路径", "C:\AppSoft\ZLHIS90.exe")
    StrName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")
    
    '到桌面的
    If Not OSfCreateShellLink("..\DeskTop", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)") Then
        Call OSfCreateShellLink("..\桌面", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)")
    End If
End Sub

Private Sub cmdStartup_Click()
    Dim strPath As String
    Dim StrName As String
    
    '创建快捷方式
    strPath = GetSetting("ZLSOFT", "公共全局", "程序路径", "C:\AppSoft\ZLHIS90.exe")
    StrName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")

    '到启动菜单
    If Not OSfCreateShellLink("\Startup", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Programs)") Then
        Call OSfCreateShellLink("\启动", StrName & "导航台(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Programs)")
    End If
End Sub

Private Sub cmdFavorite_Click(Index As Integer)
    Dim lst As ListItem, lngIndex As Long
    
    If Index <> 0 And lvwFavorite.SelectedItem Is Nothing Then Exit Sub
    
    Select Case Index
        Case 0 '新增
            If tvwMain.SelectedItem Is Nothing Then Exit Sub
            
            With tvwMain.SelectedItem
                For Each lst In lvwFavorite.ListItems
                    If lst.Tag = .Tag Then
                        MsgBox "“" & .Text & "”已经存在。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Next
                If Right(.Tag, 1) = "_" Then Exit Sub
                
                Set lst = lvwFavorite.ListItems.Add(, , .Text, .Image, .Image)
                lst.Tag = .Tag
                lst.Selected = True
                lst.EnsureVisible
            End With
        Case 1 '删除
            lngIndex = lvwFavorite.SelectedItem.Index
            lvwFavorite.ListItems.Remove lngIndex
            
            If lvwFavorite.ListItems.Count = 0 Then Exit Sub
            If lngIndex > lvwFavorite.ListItems.Count Then
                lvwFavorite.ListItems.item(lngIndex - 1).Selected = True
            Else
                lvwFavorite.ListItems.item(lngIndex).Selected = True
            End If
        Case 2 '前移
            With lvwFavorite.SelectedItem
                If .Index = 1 Then Exit Sub
                Set lst = lvwFavorite.ListItems.Add(.Index - 1, , .Text, .Icon, .SmallIcon)
                lst.Tag = .Tag
                
                lngIndex = .Index
                lst.Selected = True
                lvwFavorite.ListItems.Remove lngIndex
                lst.EnsureVisible
            End With
        Case 3 '后移
            With lvwFavorite.SelectedItem
                If .Index = lvwFavorite.ListItems.Count Then Exit Sub
                Set lst = lvwFavorite.ListItems.Add(.Index + 2, , .Text, .Icon, .SmallIcon)
                lst.Tag = .Tag
                
                lngIndex = .Index
                lst.Selected = True
                lvwFavorite.ListItems.Remove lngIndex
                lst.EnsureVisible
            End With
    End Select
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub cmdOK_Click()
    Dim lst As ListItem, i As Integer
    Dim str系统 As String, str序号 As String, str图标 As String, str标题 As String
    
    '保存导航风格
    For i = optStyle.LBound To optStyle.UBound
        If optStyle(i).Value = True Then
            Call zlDatabase.SetPara("导航台", optStyle(i).Tag)
            Exit For
        End If
    Next
    
    '保存常用模块
    For Each lst In lvwFavorite.ListItems
        str系统 = str系统 & "," & Val(Mid(lst.Tag, 1, InStr(lst.Tag, "_") - 1))
        str序号 = str序号 & "," & Mid(lst.Tag, InStr(lst.Tag, "_") + 1)
        str图标 = str图标 & "," & Mid(lst.Icon, 2)
        str标题 = str标题 & "," & lst.Text
    Next
    If str系统 <> "" Then
        str系统 = Mid(str系统, 2)
        str序号 = Mid(str序号, 2)
        str图标 = Mid(str图标, 2)
        str标题 = Mid(str标题, 2)
    End If
    Call zlDatabase.SetPara("常用功能模块", str系统 & "|" & str序号 & "|" & str图标 & "|" & str标题)
    
    Call SaveRegister
    Unload Me
End Sub

Private Sub Form_Activate()
    Call tabMain_Click
End Sub

Private Sub Form_Load()
    Dim intIndex As Integer
    Dim strStyle As String
    Dim rsTemp As New ADODB.Recordset
    
    Call InitIcon
    Call FillCommon
    
    gstrSQL = "select 组别 from ZLMENUS group by  组别 "
    
    rsTemp.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cboGroup.AddItem rsTemp("组别")
        If rsTemp("组别") = gstrMenuSys Then intIndex = cboGroup.NewIndex
        rsTemp.MoveNext
    Loop
    mintIndex = -3
    cboGroup.ListIndex = intIndex
    
    strStyle = UCase(zlDatabase.GetPara("导航台"))
    
    For intIndex = optStyle.LBound To optStyle.UBound
        If UCase(optStyle(intIndex).Tag) = strStyle Then
            optStyle(intIndex).Value = True
            Exit For
        End If
    Next
    Call optStyle_Click(0)

    Call LoadCustom
End Sub

Private Sub InitIcon()
'功能：把图标装入到控件中
    Dim i As Long
    ils16.ListImages.Clear
    ils16.ImageWidth = 16
    ils16.ImageHeight = 16
    
    With ils16.ListImages
        For i = glngLBound To glngUBound
            .Add , "K" & i, LoadResPicture(i, vbResIcon)
        Next
    End With
End Sub

Private Sub FillTree(ByVal str组别 As String)
    Dim strSQL As String
    Dim strTemp As String
    Dim rsMenus As New ADODB.Recordset
    
    gstrSQL = "select * " & _
            " from zlMenus" & _
            " start with 上级ID is null and 组别=[1]" & _
            " connect by prior ID =上级ID  and 组别=[1] order by level,ID"
    rsMenus.CursorLocation = adUseClient
    Set rsMenus = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str组别)
    
    With rsMenus
        tvwMain.Nodes.Clear
        Do While Not .EOF
            If IsNull(rsMenus("图标")) Or rsMenus("图标") = 0 Then
                strTemp = IIf(IsNull(.Fields("模块").Value), "K99", "K100")
            Else
                strTemp = "K" & rsMenus("图标")
            End If
            If IsNull(.Fields("上级ID")) Then
                tvwMain.Nodes.Add , tvwChild, "C" & .Fields("ID").Value, .Fields("标题").Value, strTemp, strTemp
            Else
               tvwMain.Nodes.Add "C" & .Fields("上级ID").Value, tvwChild, "C" & .Fields("ID").Value, .Fields("标题").Value, strTemp, strTemp
            End If
            tvwMain.Nodes("C" & .Fields("ID").Value).Tag = IIf(IsNull(.Fields("系统")), "", .Fields("系统")) & "_" & IIf(IsNull(.Fields("模块")), "", .Fields("模块"))
            .MoveNext
        Loop
    End With
End Sub

Private Sub cboGroup_Click()
    If mintIndex <> cboGroup.ListIndex Then FillTree cboGroup.Text
    mintIndex = cboGroup.ListIndex
End Sub

Private Sub optStyle_Click(Index As Integer)
    Dim i As Integer
    
    For i = optStyle.LBound To optStyle.UBound
        If optStyle(i).Value = True Then
            picPreview.Picture = LoadResPicture(optStyle(i).Tag, vbResBitmap)
            Exit For
        End If
    Next
End Sub

Private Sub picBorder_Paint()
    Dim rc As RECT
    
    With picBorder
        rc.Left = .ScaleLeft + 1
        rc.Right = .ScaleWidth - 2
        rc.Top = .ScaleTop + 1
        rc.Bottom = .ScaleHeight - 2
    End With
    DrawEdge picBorder.hDC, rc, EDGE_RAISED, BF_RECT
End Sub

Private Sub tabMain_Click()
    fra(1).Visible = False
    fra(2).Visible = False
    fra(3).Visible = False
    fra(tabMain.SelectedItem.Index).Visible = True
End Sub

Private Sub FillCommon()
'功能：装入常用的程序
    Dim var系统 As Variant, var序号 As Variant, var图标 As Variant, var标题 As Variant
    Dim lngMax As Long, lngCount As Long, lst As ListItem, strValue As String
    
    strValue = zlDatabase.GetPara("常用功能模块")
    If UBound(Split(strValue, "|")) < 3 Then Exit Sub
    var系统 = Split(Split(strValue, "|")(0), ",")
    var序号 = Split(Split(strValue, "|")(1), ",")
    var图标 = Split(Split(strValue, "|")(2), ",")
    var标题 = Split(Split(strValue, "|")(3), ",")
    
    lngMax = IIf(UBound(var系统) > UBound(var序号), UBound(var系统), UBound(var序号))
    lngMax = IIf(lngMax > UBound(var图标), lngMax, UBound(var图标))
    lngMax = IIf(lngMax > UBound(var标题), lngMax, UBound(var标题))
    If lngMax = -1 Then Exit Sub
    
    For lngCount = 0 To lngMax
        Set lst = lvwFavorite.ListItems.Add(, , var标题(lngCount), "K" & var图标(lngCount), "K" & var图标(lngCount))
        lst.Tag = var系统(lngCount) & "_" & var序号(lngCount)
    Next
End Sub

Private Sub tvwMain_DblClick()
    Call cmdFavorite_Click(0)
End Sub

Private Sub LoadCustom()
'完成用户习惯的初始化工作
    Dim lng简码 As Long
    Dim strPath As String
    
    '输入匹配
    If Val(zlDatabase.GetPara("输入匹配")) = 0 Then
        opt(opt_双向匹配).Value = True
        opt(opt_从左匹配).Value = False
    Else
        opt(opt_双向匹配).Value = False
        opt(opt_从左匹配).Value = True
    End If
    
    '汉字输入法
    Call ChooseIME(cmbIME)
    
    '简码生成方式
    lng简码 = Val(zlDatabase.GetPara("简码生成"))
    If lng简码 = 0 Then
        opt简码(0).Value = True
    Else
        opt简码(1).Value = True
    End If
    
    '自动运行本程序
    Call zlCommFun.GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer", strPath)
    chkAutoStart.Value = IIf(Trim(strPath) <> "", 1, 0)
    
    '自动关闭Windows
    chkShutDown.Value = IIf(Val(zlDatabase.GetPara("关闭Windows")) = 1, 1, 0)
    
    '界面区域隐藏
    chkAutoHide.Value = IIf(Val(zlDatabase.GetPara("界面区域隐藏")) = 1, 1, 0)
    
    '通知检查周期
    txtTime.Text = zlDatabase.GetPara("邮件消息检查周期")
    If Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 60 Then txtTime.Text = 30
    
End Sub

Private Sub SaveRegister()
'保存到注册表中的信息
    Dim lng简码 As Long
    Dim strExeName As String

    Call zlDatabase.SetPara("输入匹配", IIf(opt(opt_双向匹配).Value = True, "0", "1"))
    Call zlDatabase.SetPara("输入法", IIf(cmbIME.Text = "不自动开启", "", cmbIME.Text))

    For lng简码 = opt简码.LBound To opt简码.UBound
        If opt简码(lng简码).Value = True Then
            Call zlDatabase.SetPara("简码方式", lng简码)
            Exit For
        End If
    Next

    '自动运行本程序
    If chkAutoStart.Value = 1 Then
        strExeName = GetSetting("ZLSOFT", "公共全局", "程序路径", "C:\AppSoft\ZLHIS90.exe")
        strExeName = Replace(strExeName, ":\\", ":\")
        Call zlCommFun.SetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer", strExeName)
    Else
        Call zlCommFun.DeleteRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer")
    End If
    
    '界面区域隐藏
    Call zlDatabase.SetPara("界面区域隐藏", chkAutoHide.Value)
    
    '自动关闭Windows
    Call zlDatabase.SetPara("关闭Windows", chkShutDown.Value)
    
    '消息通知检查周期
    If Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 60 Then txtTime.Text = 30
    Call zlDatabase.SetPara("通知检查周期", Val(txtTime.Text))
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTime_Validate(Cancel As Boolean)
    If Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 60 Then txtTime.Text = 30
End Sub

Private Sub txtTime_GotFocus()
    Call zlControl.TxtSelAll(txtTime)
End Sub
