VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ϵͳѡ��"
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
   StartUpPosition =   1  '����������
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
      Begin VB.PictureBox pic��� 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   450
         ScaleHeight     =   480
         ScaleWidth      =   4485
         TabIndex        =   34
         Top             =   900
         Width           =   4485
         Begin VB.OptionButton optStyle 
            Caption         =   "Windows���"
            Height          =   195
            Index           =   1
            Left            =   1515
            TabIndex        =   2
            Tag             =   "zlwin"
            Top             =   60
            Width           =   1305
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "��ͳ���"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   1
            Tag             =   "zlBrw"
            Top             =   60
            Width           =   1275
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "MDI���"
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
         Caption         =   "    ���ո��˵���Ȥ����ѡ���Լ�ϲ���ĵ������ʹ���������졢���������ɡ�"
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
      Begin VB.PictureBox pic���� 
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
         Begin VB.OptionButton opt���� 
            Caption         =   "ƴ����ȡÿ�ֵ�����ĸ���ɼ���(&P)"
            Height          =   210
            Index           =   0
            Left            =   15
            TabIndex        =   17
            Top             =   90
            Value           =   -1  'True
            Width           =   3150
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "��ʣ�ȡÿ�ֵ�����ĸ���ɼ���(&W)"
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
         Caption         =   "�� Windows ����ʱ�Զ�����(&A)"
         Height          =   210
         Left            =   1530
         TabIndex        =   20
         Top             =   4185
         Width           =   2865
      End
      Begin VB.CheckBox chkShutDown 
         Caption         =   "�˳�����ʱ�Զ��ر� Windows (&S)"
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
         ToolTipText     =   "�����������10�룬���3600��"
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
      Begin VB.PictureBox picƥ�䷽ʽ 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1455
         ScaleHeight     =   315
         ScaleWidth      =   3495
         TabIndex        =   38
         Top             =   960
         Width           =   3495
         Begin VB.OptionButton opt 
            Caption         =   "����ƥ��(&L)"
            Height          =   210
            Index           =   1
            Left            =   1815
            TabIndex        =   15
            Top             =   60
            Width           =   1320
         End
         Begin VB.OptionButton opt 
            Caption         =   "˫��ƥ��(&D)"
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
         Caption         =   "������������ṩ�Զ����ع���(&I)"
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
         Caption         =   "ÿ       ������Ϣ֪ͨ"
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
         Caption         =   "    �û����Ը��������ϰ����ѡ�������ƥ�䷽ʽ���������͡����뷨�ȣ�����߹���Ч��"
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
         Caption         =   "��Ŀ����ƥ�䷽ʽ"
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
         Caption         =   "�������뷨"
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
         Caption         =   "���봦��ʽ"
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
         Caption         =   "��Windows�Ľ��"
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
         Caption         =   "��Ϣ֪ͨ�������"
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
         Caption         =   "������������"
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
         ToolTipText     =   "����ƶ�"
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
         ToolTipText     =   "��ǰ�ƶ�"
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
         ToolTipText     =   "ɾ������"
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
         ToolTipText     =   "��������"
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
            Text            =   "����"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "���ɿ�ݷ�ʽ"
         Height          =   1335
         Left            =   3030
         TabIndex        =   29
         Top             =   750
         Width           =   2265
         Begin VB.CommandButton cmdStartup 
            Caption         =   "�������˵�(&S)"
            Height          =   350
            Left            =   270
            TabIndex        =   8
            Top             =   840
            Width           =   1725
         End
         Begin VB.CommandButton cmdDesktop 
            Caption         =   "������(&D)"
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
         Caption         =   "���ó����б�(&R)"
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
         Caption         =   "�˵���ϵ"
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
      Caption         =   "����(&H)"
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
            Caption         =   "�������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�˵�ѡ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ʹ��ϰ��"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4620
      TabIndex        =   24
      Top             =   6510
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
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
    opt_˫��ƥ�� = 0
    opt_����ƥ�� = 1
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
    
    '������ݷ�ʽ
    strPath = GetSetting("ZLSOFT", "����ȫ��", "����·��", "C:\AppSoft\ZLHIS90.exe")
    StrName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    
    '�������
    If Not OSfCreateShellLink("..\DeskTop", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)") Then
        Call OSfCreateShellLink("..\����", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)")
    End If
End Sub

Private Sub cmdStartup_Click()
    Dim strPath As String
    Dim StrName As String
    
    '������ݷ�ʽ
    strPath = GetSetting("ZLSOFT", "����ȫ��", "����·��", "C:\AppSoft\ZLHIS90.exe")
    StrName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")

    '�������˵�
    If Not OSfCreateShellLink("\Startup", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Programs)") Then
        Call OSfCreateShellLink("\����", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Programs)")
    End If
End Sub

Private Sub cmdFavorite_Click(Index As Integer)
    Dim lst As ListItem, lngIndex As Long
    
    If Index <> 0 And lvwFavorite.SelectedItem Is Nothing Then Exit Sub
    
    Select Case Index
        Case 0 '����
            If tvwMain.SelectedItem Is Nothing Then Exit Sub
            
            With tvwMain.SelectedItem
                For Each lst In lvwFavorite.ListItems
                    If lst.Tag = .Tag Then
                        MsgBox "��" & .Text & "���Ѿ����ڡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Next
                If Right(.Tag, 1) = "_" Then Exit Sub
                
                Set lst = lvwFavorite.ListItems.Add(, , .Text, .Image, .Image)
                lst.Tag = .Tag
                lst.Selected = True
                lst.EnsureVisible
            End With
        Case 1 'ɾ��
            lngIndex = lvwFavorite.SelectedItem.Index
            lvwFavorite.ListItems.Remove lngIndex
            
            If lvwFavorite.ListItems.Count = 0 Then Exit Sub
            If lngIndex > lvwFavorite.ListItems.Count Then
                lvwFavorite.ListItems.item(lngIndex - 1).Selected = True
            Else
                lvwFavorite.ListItems.item(lngIndex).Selected = True
            End If
        Case 2 'ǰ��
            With lvwFavorite.SelectedItem
                If .Index = 1 Then Exit Sub
                Set lst = lvwFavorite.ListItems.Add(.Index - 1, , .Text, .Icon, .SmallIcon)
                lst.Tag = .Tag
                
                lngIndex = .Index
                lst.Selected = True
                lvwFavorite.ListItems.Remove lngIndex
                lst.EnsureVisible
            End With
        Case 3 '����
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
    Dim strϵͳ As String, str��� As String, strͼ�� As String, str���� As String
    
    '���浼�����
    For i = optStyle.LBound To optStyle.UBound
        If optStyle(i).Value = True Then
            Call zlDatabase.SetPara("����̨", optStyle(i).Tag)
            Exit For
        End If
    Next
    
    '���泣��ģ��
    For Each lst In lvwFavorite.ListItems
        strϵͳ = strϵͳ & "," & Val(Mid(lst.Tag, 1, InStr(lst.Tag, "_") - 1))
        str��� = str��� & "," & Mid(lst.Tag, InStr(lst.Tag, "_") + 1)
        strͼ�� = strͼ�� & "," & Mid(lst.Icon, 2)
        str���� = str���� & "," & lst.Text
    Next
    If strϵͳ <> "" Then
        strϵͳ = Mid(strϵͳ, 2)
        str��� = Mid(str���, 2)
        strͼ�� = Mid(strͼ��, 2)
        str���� = Mid(str����, 2)
    End If
    Call zlDatabase.SetPara("���ù���ģ��", strϵͳ & "|" & str��� & "|" & strͼ�� & "|" & str����)
    
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
    
    gstrSQL = "select ��� from ZLMENUS group by  ��� "
    
    rsTemp.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cboGroup.AddItem rsTemp("���")
        If rsTemp("���") = gstrMenuSys Then intIndex = cboGroup.NewIndex
        rsTemp.MoveNext
    Loop
    mintIndex = -3
    cboGroup.ListIndex = intIndex
    
    strStyle = UCase(zlDatabase.GetPara("����̨"))
    
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
'���ܣ���ͼ��װ�뵽�ؼ���
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

Private Sub FillTree(ByVal str��� As String)
    Dim strSQL As String
    Dim strTemp As String
    Dim rsMenus As New ADODB.Recordset
    
    gstrSQL = "select * " & _
            " from zlMenus" & _
            " start with �ϼ�ID is null and ���=[1]" & _
            " connect by prior ID =�ϼ�ID  and ���=[1] order by level,ID"
    rsMenus.CursorLocation = adUseClient
    Set rsMenus = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str���)
    
    With rsMenus
        tvwMain.Nodes.Clear
        Do While Not .EOF
            If IsNull(rsMenus("ͼ��")) Or rsMenus("ͼ��") = 0 Then
                strTemp = IIf(IsNull(.Fields("ģ��").Value), "K99", "K100")
            Else
                strTemp = "K" & rsMenus("ͼ��")
            End If
            If IsNull(.Fields("�ϼ�ID")) Then
                tvwMain.Nodes.Add , tvwChild, "C" & .Fields("ID").Value, .Fields("����").Value, strTemp, strTemp
            Else
               tvwMain.Nodes.Add "C" & .Fields("�ϼ�ID").Value, tvwChild, "C" & .Fields("ID").Value, .Fields("����").Value, strTemp, strTemp
            End If
            tvwMain.Nodes("C" & .Fields("ID").Value).Tag = IIf(IsNull(.Fields("ϵͳ")), "", .Fields("ϵͳ")) & "_" & IIf(IsNull(.Fields("ģ��")), "", .Fields("ģ��"))
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
'���ܣ�װ�볣�õĳ���
    Dim varϵͳ As Variant, var��� As Variant, varͼ�� As Variant, var���� As Variant
    Dim lngMax As Long, lngCount As Long, lst As ListItem, strValue As String
    
    strValue = zlDatabase.GetPara("���ù���ģ��")
    If UBound(Split(strValue, "|")) < 3 Then Exit Sub
    varϵͳ = Split(Split(strValue, "|")(0), ",")
    var��� = Split(Split(strValue, "|")(1), ",")
    varͼ�� = Split(Split(strValue, "|")(2), ",")
    var���� = Split(Split(strValue, "|")(3), ",")
    
    lngMax = IIf(UBound(varϵͳ) > UBound(var���), UBound(varϵͳ), UBound(var���))
    lngMax = IIf(lngMax > UBound(varͼ��), lngMax, UBound(varͼ��))
    lngMax = IIf(lngMax > UBound(var����), lngMax, UBound(var����))
    If lngMax = -1 Then Exit Sub
    
    For lngCount = 0 To lngMax
        Set lst = lvwFavorite.ListItems.Add(, , var����(lngCount), "K" & varͼ��(lngCount), "K" & varͼ��(lngCount))
        lst.Tag = varϵͳ(lngCount) & "_" & var���(lngCount)
    Next
End Sub

Private Sub tvwMain_DblClick()
    Call cmdFavorite_Click(0)
End Sub

Private Sub LoadCustom()
'����û�ϰ�ߵĳ�ʼ������
    Dim lng���� As Long
    Dim strPath As String
    
    '����ƥ��
    If Val(zlDatabase.GetPara("����ƥ��")) = 0 Then
        opt(opt_˫��ƥ��).Value = True
        opt(opt_����ƥ��).Value = False
    Else
        opt(opt_˫��ƥ��).Value = False
        opt(opt_����ƥ��).Value = True
    End If
    
    '�������뷨
    Call ChooseIME(cmbIME)
    
    '�������ɷ�ʽ
    lng���� = Val(zlDatabase.GetPara("��������"))
    If lng���� = 0 Then
        opt����(0).Value = True
    Else
        opt����(1).Value = True
    End If
    
    '�Զ����б�����
    Call zlCommFun.GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer", strPath)
    chkAutoStart.Value = IIf(Trim(strPath) <> "", 1, 0)
    
    '�Զ��ر�Windows
    chkShutDown.Value = IIf(Val(zlDatabase.GetPara("�ر�Windows")) = 1, 1, 0)
    
    '������������
    chkAutoHide.Value = IIf(Val(zlDatabase.GetPara("������������")) = 1, 1, 0)
    
    '֪ͨ�������
    txtTime.Text = zlDatabase.GetPara("�ʼ���Ϣ�������")
    If Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 60 Then txtTime.Text = 30
    
End Sub

Private Sub SaveRegister()
'���浽ע����е���Ϣ
    Dim lng���� As Long
    Dim strExeName As String

    Call zlDatabase.SetPara("����ƥ��", IIf(opt(opt_˫��ƥ��).Value = True, "0", "1"))
    Call zlDatabase.SetPara("���뷨", IIf(cmbIME.Text = "���Զ�����", "", cmbIME.Text))

    For lng���� = opt����.LBound To opt����.UBound
        If opt����(lng����).Value = True Then
            Call zlDatabase.SetPara("���뷽ʽ", lng����)
            Exit For
        End If
    Next

    '�Զ����б�����
    If chkAutoStart.Value = 1 Then
        strExeName = GetSetting("ZLSOFT", "����ȫ��", "����·��", "C:\AppSoft\ZLHIS90.exe")
        strExeName = Replace(strExeName, ":\\", ":\")
        Call zlCommFun.SetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer", strExeName)
    Else
        Call zlCommFun.DeleteRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer")
    End If
    
    '������������
    Call zlDatabase.SetPara("������������", chkAutoHide.Value)
    
    '�Զ��ر�Windows
    Call zlDatabase.SetPara("�ر�Windows", chkShutDown.Value)
    
    '��Ϣ֪ͨ�������
    If Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 60 Then txtTime.Text = 30
    Call zlDatabase.SetPara("֪ͨ�������", Val(txtTime.Text))
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
