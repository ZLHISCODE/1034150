VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmChartSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ͼ������"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "frmChartSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtWidth 
      Height          =   275
      Left            =   3180
      TabIndex        =   60
      Text            =   "1"
      Top             =   4870
      Width           =   330
   End
   Begin VB.ComboBox cboFS2 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   58
      Top             =   1560
      Width           =   2400
   End
   Begin VB.ComboBox cboFY2 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtColor 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   54
      ToolTipText     =   "���е���ɫ˳������������ԴSQL��ѯ����������˳����û�����ã���������ʾĬ����ɫ"
      Top             =   2799
      Width           =   2150
   End
   Begin VB.CommandButton cmdChoose 
      Appearance      =   0  'Flat
      Caption         =   "��"
      Height          =   285
      Left            =   3225
      TabIndex        =   53
      Top             =   2799
      Width           =   255
   End
   Begin VB.PictureBox picAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   120
      ScaleHeight     =   1725
      ScaleWidth      =   2370
      TabIndex        =   42
      Top             =   5400
      Visible         =   0   'False
      Width           =   2400
      Begin VB.CommandButton cmdDelte 
         Caption         =   "ɾ��"
         Height          =   300
         Left            =   600
         TabIndex        =   52
         ToolTipText     =   "ɾ��һ����ɫ"
         Top             =   1320
         Width           =   550
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���"
         Height          =   300
         Left            =   0
         TabIndex        =   51
         ToolTipText     =   "���һ����ɫ"
         Top             =   1320
         Width           =   550
      End
      Begin VB.PictureBox picUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2145
         Picture         =   "frmChartSetup.frx":058A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   50
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picDown 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2145
         Picture         =   "frmChartSetup.frx":0F8C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   49
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   0
         ScaleHeight     =   1140
         ScaleWidth      =   2130
         TabIndex        =   45
         Top             =   0
         Width           =   2135
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   705
            Left            =   0
            ScaleHeight     =   705
            ScaleWidth      =   1065
            TabIndex        =   46
            Top             =   0
            Width           =   1065
            Begin VB.CommandButton cmdColor 
               Height          =   300
               Index           =   0
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   80
               Width           =   330
            End
            Begin VB.Label lblColor 
               AutoSize        =   -1  'True
               BackColor       =   &H80000014&
               Caption         =   "1"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   48
               Top             =   390
               Width           =   90
            End
         End
      End
      Begin VB.CommandButton cmdS 
         Caption         =   "ȷ��"
         Height          =   300
         Left            =   1200
         TabIndex        =   44
         Top             =   1320
         Width           =   550
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "ȡ��"
         Height          =   300
         Left            =   1800
         TabIndex        =   43
         Top             =   1320
         Width           =   550
      End
      Begin VB.Line Line2 
         X1              =   2135
         X2              =   2135
         Y1              =   0
         Y2              =   1215
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   2400
         Y1              =   1215
         Y2              =   1215
      End
   End
   Begin VB.CheckBox chkLabel 
      Caption         =   "��ʾ��ǩ"
      Height          =   180
      Left            =   3630
      TabIndex        =   39
      Top             =   4900
      Width           =   1095
   End
   Begin VB.CheckBox chklaLine 
      Caption         =   "��ʾ��ǩ����"
      Enabled         =   0   'False
      Height          =   180
      Left            =   7215
      TabIndex        =   38
      Top             =   4900
      Width           =   1380
   End
   Begin VB.TextBox txtLen 
      Enabled         =   0   'False
      Height          =   275
      Left            =   8610
      TabIndex        =   36
      Text            =   "1"
      Top             =   4870
      Width           =   400
   End
   Begin VB.ComboBox cboLabel 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   4830
      Width           =   1335
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "XY�ύ��"
      Height          =   195
      Index           =   1
      Left            =   4830
      TabIndex        =   25
      Top             =   4485
      Width           =   1020
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "��άЧ��"
      Height          =   195
      Index           =   0
      Left            =   3630
      TabIndex        =   24
      Top             =   4485
      Width           =   1020
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1900
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4830
      Width           =   330
   End
   Begin VB.CommandButton cmdFore 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4830
      Width           =   330
   End
   Begin VB.CommandButton cmdFont 
      Height          =   315
      Left            =   3150
      Picture         =   "frmChartSetup.frx":198E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "��������"
      Top             =   3996
      Width           =   330
   End
   Begin VB.TextBox txtFont 
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3996
      Width           =   2085
   End
   Begin VB.TextBox txtFontTitle 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3597
      Width           =   2085
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   9360
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   15
      TabIndex        =   34
      Top             =   5280
      Width           =   9765
   End
   Begin VB.ComboBox cboLocate 
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4395
      Width           =   2070
   End
   Begin C1Chart2D8.Chart2D Chart 
      Height          =   3960
      Left            =   3645
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   360
      Width           =   5880
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   10372
      _ExtentY        =   6985
      _StockProps     =   0
      ControlProperties=   "frmChartSetup.frx":1F18
   End
   Begin VB.CheckBox chkNode 
      Caption         =   "��ʾ���"
      Height          =   195
      Left            =   8400
      TabIndex        =   28
      Top             =   4485
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.CheckBox chkLine 
      Caption         =   "��ʾ����"
      Height          =   195
      Left            =   7215
      TabIndex        =   27
      Top             =   4485
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.CheckBox chkSample 
      Alignment       =   1  'Right Justify
      Caption         =   "��ʾͼ��"
      Height          =   195
      Left            =   255
      TabIndex        =   22
      Top             =   4425
      Width           =   1050
   End
   Begin VB.CommandButton cmdFontTitle 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3150
      Picture         =   "frmChartSetup.frx":2577
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "��������"
      Top             =   3597
      Width           =   330
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3198
      Width           =   2400
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ӧ��(&A)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8310
      TabIndex        =   33
      Top             =   5490
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7020
      TabIndex        =   32
      Top             =   5490
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5895
      TabIndex        =   31
      Top             =   5490
      Width           =   1100
   End
   Begin VB.ComboBox cboStyle 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2400
      Width           =   2400
   End
   Begin VB.ComboBox cboFY 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin VB.ComboBox cboFS 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2400
   End
   Begin VB.ComboBox cboFX 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   465
      Width           =   2400
   End
   Begin VB.ComboBox cboData 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   105
      Width           =   2400
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "��ʾ����"
      Height          =   195
      Left            =   6015
      TabIndex        =   26
      Top             =   4485
      Width           =   1020
   End
   Begin MSComCtl2.UpDown UpDown 
      Height          =   300
      Left            =   9030
      TabIndex        =   37
      Top             =   4830
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtLen"
      BuddyDispid     =   196629
      OrigLeft        =   5280
      OrigTop         =   4440
      OrigRight       =   5535
      OrigBottom      =   4815
      Max             =   50
      Enabled         =   0   'False
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "������ϸ"
      Height          =   180
      Left            =   2400
      TabIndex        =   61
      Top             =   4905
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��2�����ֶ�"
      Height          =   180
      Left            =   15
      TabIndex        =   59
      Top             =   1620
      Width           =   990
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��2ֵ�ֶ�"
      Height          =   180
      Left            =   195
      TabIndex        =   57
      Top             =   1980
      Width           =   810
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "������ɫ"
      Height          =   180
      Left            =   285
      TabIndex        =   55
      ToolTipText     =   "���е���ɫ˳������������ԴSQL��ѯ����������˳����û�����ã���������ʾĬ����ɫ"
      Top             =   2845
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "mm"
      Height          =   180
      Left            =   9330
      TabIndex        =   41
      Top             =   4920
      Width           =   180
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "��ǩ��ʽ"
      Height          =   180
      Left            =   4830
      TabIndex        =   40
      Top             =   4900
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ɫ"
      Height          =   180
      Left            =   1320
      TabIndex        =   20
      Top             =   4900
      Width           =   540
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǰ��ɫ"
      Height          =   180
      Left            =   255
      TabIndex        =   18
      Top             =   4900
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͼ������"
      Height          =   180
      Left            =   285
      TabIndex        =   15
      Top             =   4030
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   285
      TabIndex        =   12
      Top             =   3635
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   3600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   3420
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͼ��ʾ����"
      Height          =   180
      Left            =   3645
      TabIndex        =   29
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ı�"
      Height          =   180
      Left            =   285
      TabIndex        =   10
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͼ����ʽ"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   285
      TabIndex        =   8
      Top             =   2450
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��1ֵ�ֶ�"
      Height          =   180
      Left            =   195
      TabIndex        =   6
      Top             =   1260
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��1�����ֶ�"
      Height          =   180
      Left            =   15
      TabIndex        =   4
      Top             =   900
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ֵ�ֶ�"
      Height          =   180
      Left            =   285
      TabIndex        =   2
      Top             =   525
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Դ"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   285
      TabIndex        =   0
      Top             =   165
      Width           =   720
   End
End
Attribute VB_Name = "frmChartSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mobjChart As Object 'byRef:In/Out
Private mobjDatas As RPTDatas 'In
Private mobjItem As RPTItem 'byRef:In/Out
Private mtmpItem As RPTItem
Private mblnAdd As Boolean '���ڼ��ؿؼ�
Private mstr��ͷ As String
Private mstrColor As String '������ɫ������ ��ͷ �ֶ�

Private Property Let ItemChange(ByVal vData As Boolean)
    cmdApply.Enabled = vData
    If vData Then
        Call SetChartStyleAndData(Chart, mtmpItem)
    End If
End Property

Private Property Get ItemChange() As Boolean
    ItemChange = cmdApply.Enabled
End Property

Public Function ShowMe(frmParent As Object, ByVal objDatas As RPTDatas, objChart As Object, objItem As RPTItem) As Boolean
    Set mobjDatas = objDatas
    Set mobjChart = objChart
    Set mobjItem = objItem
    
    Me.Show 1, frmParent
    If mblnOK Then '����ж�ػ������ʱ,�����ж�����������ù�ϵ���ж�
        Call CopyItem(objItem, mobjItem)
    End If
    ShowMe = mblnOK
End Function

Private Sub cboFS2_Click()
    If cboFS2.Text <> "" Then
        chkFormat(0).Enabled = False
        chkFormat(0).Value = 0
    End If
    Call SetChartData
End Sub

Private Sub cboFY2_Click()
    If cboFY2.Text <> "" Then
        chkFormat(0).Enabled = False
        chkFormat(0).Value = 0
    End If
    SetChartData
End Sub

Private Sub cboLabel_Click()
    If Visible Then
        mtmpItem.���� = chkLabel.Value & "|" & cboLabel.ListIndex & "|" & chklaLine.Value & "|" & txtLen.Text
        ItemChange = True
    End If
End Sub

Private Sub chkLabel_Click()
    If chkLabel.Value = 1 Then
        cboLabel.Enabled = True
        chklaLine.Enabled = True
        txtLen.Enabled = True
        UpDown.Enabled = True
    Else
        cboLabel.ListIndex = -1
        chklaLine.Value = 0
        cboLabel.Enabled = False
        chklaLine.Enabled = False
        txtLen.Enabled = False
        UpDown.Enabled = False
    End If
    
    If Visible Then
        mtmpItem.���� = chkLabel.Value & "|" & cboLabel.ListIndex & "|" & chklaLine.Value & "|" & txtLen.Text
        ItemChange = True
    End If
End Sub

Private Sub chklaLine_Click()
    If Visible Then
        mtmpItem.���� = chkLabel.Value & "|" & cboLabel.ListIndex & "|" & chklaLine.Value & "|" & txtLen.Text
        ItemChange = True
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim i As Long, lngMax As Long
    
    lngMax = cmdColor.UBound
    mblnAdd = True
    For i = lngMax + 1 To lngMax + 5
        Load cmdColor(i)
        Load lblColor(i)
        cmdColor(i).Visible = True
        cmdColor(i).BackColor = &H8000000F
        lblColor(i).Visible = True
        lblColor(i).Caption = i + 1
        If i = lngMax + 1 Then
            cmdColor(i).Move cmdColor(0).Left, cmdColor(0).Top + (cmdColor(0).Height + lblColor(0).Height + 90) * Fix((lngMax + 1) / 5) + 1, cmdColor(0).Width, cmdColor(0).Height
        Else
            cmdColor(i).Move cmdColor(i - 1).Left + cmdColor(i - 1).Width + 60, cmdColor(i - 1).Top, cmdColor(0).Width, cmdColor(0).Height
        End If
        lblColor(i).Move cmdColor(i).Left + IIF(i < 9, 120, 80), cmdColor(i).Top + cmdColor(i).Height + 20
    Next
    picColor.Height = picColor.Height + (cmdColor(0).Height + lblColor(0).Height + 90)
    picColor.Top = picBack.Height - picColor.Height
    If cmdColor.UBound > 10 Then
        cmdDelte.Enabled = True
        picUp.Visible = True
    Else
        picUp.Visible = False
    End If
    picDown.Visible = False
    mblnAdd = False
End Sub

Private Sub cmdApply_Click()
    If Not CheckInput Then Exit Sub
    Call CopyItem(mobjItem, mtmpItem)
    Call SetChartStyleAndData(mobjChart, mobjItem, , , True)
    mblnOK = True
    ItemChange = False
End Sub

Private Sub cmdC_Click()
    picAll.Visible = False
End Sub

Private Sub cmdChoose_Click()

    picColor.Top = 0
    SetColor
    picAll.Visible = True
    picAll.Top = txtColor.Top + txtColor.Height
    picAll.Left = txtColor.Left
    picAll.SetFocus
    picAll.ZOrder
End Sub

Private Sub cmdColor_Click(Index As Integer)
    Dim i As Long
    
    On Error Resume Next
    If mblnAdd Then Exit Sub
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = cmdColor(Index).BackColor
    cdg.ShowColor
    If Err.Number = 0 Then
        cmdColor(Index).BackColor = cdg.Color
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdDelte_Click()
    Dim i As Long
    
    For i = cmdColor.UBound - 4 To cmdColor.UBound
        Unload cmdColor(i)
        Unload lblColor(i)
    Next
    picColor.Height = picColor.Height - (cmdColor(0).Height + lblColor(0).Height + 90)
    picColor.Top = picBack.Height - picColor.Height
    If cmdColor.UBound < 10 Then
        cmdDelte.Enabled = False
        picUp.Visible = False
    Else
        picUp.Visible = True
    End If
    picDown.Visible = False
End Sub

Private Sub cmdOK_Click()
    If Not CheckInput Then Exit Sub
    Call CopyItem(mobjItem, mtmpItem)
    Call SetChartStyleAndData(mobjChart, mobjItem, , , True)
    mblnOK = True
    Unload Me
End Sub

Private Sub SetOptionEnabled()
    '0-Plot(ɢ��ͼ),1-Plot(����ͼ),2-Bar(����ͼ),3-Pie(��ͼ),4-StackingBar(���ͼ),5-Area(���ͼ)
    '6-HiLo(�ɼ�ͼ-�̸�,�̵�),7-HiLoOpenClose(�ɼ�ͼ-�̸�,�̵�,����,����),8-Candle(�ɼ�ͼ-������ͼ:�̸�,�̵�,����,����)
    '9-Polar(����ͼ),10-Radar(�״�ͼ),11-FilledRadar(����״�ͼ),12-Bubble(����ͼ)
    
    '������ͼ������ά��ʽ
    chkFormat(0).Enabled = InStr(",1,2,3,4,5,", "," & cboStyle.ListIndex & ",") > 0
    If Not chkFormat(0).Enabled Then
        chkFormat(0).Value = 0
    End If
    
    '������ͼ��XY�ύ����Ч
    chkFormat(1).Enabled = InStr(",3,9,10,11,", "," & cboStyle.ListIndex & ",") = 0
    If Not chkFormat(1).Enabled Then
        chkFormat(1).Value = 0
    End If
    
    '��ͼ������
    chkGrid.Enabled = cboStyle.ListIndex <> 3
    If Not chkGrid.Enabled Then chkGrid.Value = 0
    
    '������ͼ��������
    chkLine.Enabled = InStr(",2,3,4,5,", "," & cboStyle.ListIndex & ",") = 0
    If Not chkLine.Enabled Then chkLine.Value = 0
    
    '������ͼ���н��
    chkNode.Enabled = InStr(",2,3,4,5,6,7,8,11,", "," & cboStyle.ListIndex & ",") = 0
    If Not chkNode.Enabled Then chkNode.Value = 0
    
    '������ͼ����ʽ֧��˫Y��
    If InStr(",1,2,4,5,6,7,8,12,", "," & cboStyle.ListIndex & ",") = 0 Then
        cboFS2.Enabled = False
        cboFY2.Enabled = False
        cboFS2.ListIndex = -1
        cboFY2.ListIndex = -1
        Label16.ToolTipText = "��ǰͼ����ʽ��֧��˫Y��"
        Label15.ToolTipText = "��ǰͼ����ʽ��֧��˫Y��"
    Else
        Label16.ToolTipText = ""
        Label15.ToolTipText = ""
    End If
    If cboFS2.Text <> "" Or cboFY2.Text <> "" Then
        chkFormat(0).Enabled = False
        chkFormat(0).Value = 0
    End If
End Sub

Private Sub chkFormat_Click(Index As Integer)
    Dim i As Integer
    If Visible Then
        mtmpItem.��ʽ = ""
        For i = 0 To chkFormat.UBound
            mtmpItem.��ʽ = mtmpItem.��ʽ & CStr(chkFormat(i).Value)
        Next
        ItemChange = True
    End If
End Sub

Private Sub cboData_Click()
    Dim arrField As Variant, strField As String
    Dim strFX As String, strFY As String, strFS As String, strFY2 As String, strFS2 As String
    Dim i As Long
    
    If cboData.ListIndex = -1 Then
        Call CboSetIndex(cboFX.hwnd, -1)
        Call CboSetIndex(cboFS.hwnd, -1)
        Call CboSetIndex(cboFY.hwnd, -1)
        Call CboSetIndex(cboFS2.hwnd, -1)
        Call CboSetIndex(cboFY2.hwnd, -1)
        mtmpItem.���� = ""
        Call SetChartStyleAndData(Chart, mtmpItem)
        Exit Sub
    End If
    
    '������ʾ�����ֶ�
    cboFX.Clear: cboFY.Clear: cboFS.Clear: cboFY2.Clear: cboFS2.Clear '������ἤ��Click
    cboFY2.AddItem ""
    cboFS2.AddItem ""
    strField = mobjDatas("_" & cboData.Text).�ֶ�
    If strField <> "" Then
        arrField = Split(strField, "|")
        For i = 0 To UBound(arrField)
            strField = Split(arrField(i), ",")(0)
            Select Case Val(Split(arrField(i), ",")(1))
                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                    cboFX.AddItem strField
                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    cboFX.AddItem strField
                    cboFY.AddItem strField
                    cboFY2.AddItem strField
                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                    cboFS.AddItem strField
                    cboFS2.AddItem strField
            End Select
        Next
    End If
            
    '���ݶ���ֵ��λ�ֶ�
    Call GetChartDataName(mtmpItem.����, strFX, strFS, strFY, , strFY2, strFS2)
    If strFX <> "" Then
        i = GetCboIndex(cboFX, strFX)
        Call CboSetIndex(cboFX.hwnd, i)
    End If
    If strFS <> "" Then
        i = GetCboIndex(cboFS, strFS)
        Call CboSetIndex(cboFS.hwnd, i)
    End If
    If strFY <> "" Then
        i = GetCboIndex(cboFY, strFY)
        Call CboSetIndex(cboFY.hwnd, i)
    End If
    If strFS2 <> "" Then
        i = GetCboIndex(cboFS2, strFS2)
        Call CboSetIndex(cboFS2.hwnd, i)
    End If
    If strFY2 <> "" Then
        i = GetCboIndex(cboFY2, strFY2)
        Call CboSetIndex(cboFY2.hwnd, i)
    End If
    '��������ֵ�������
    Call SetChartData
End Sub

Private Sub cboFX_Click()
    Call SetChartData
End Sub

Private Sub cboFS_Click()
    Call SetChartData
End Sub

Private Sub cboFY_Click()
    Call SetChartData
End Sub

Private Sub SetChartData()
'���ܣ����ݵ�ǰ�������������,����Chartʾ����ʾ
    Dim strFX As String, strFY As String, strFS As String, strFY2 As String, strFS2 As String
    Dim str���� As String

    strFX = cboFX.Text
    strFS = cboFS.Text
    strFY = cboFY.Text
    strFY2 = cboFY2.Text
    strFS2 = cboFS2.Text
    If strFX <> "" Then
        str���� = str���� & "|" & cboData.Text & "." & strFX
    Else
        str���� = str���� & "|"
    End If
    If strFS <> "" Then
        str���� = str���� & "|" & cboData.Text & "." & strFS
    Else
        str���� = str���� & "|"
    End If
    If strFY <> "" Then
        str���� = str���� & "|" & cboData.Text & "." & strFY
    Else
        str���� = str���� & "|"
    End If
    If strFS2 <> "" Then
        str���� = str���� & "|" & cboData.Text & "." & strFS2
    Else
        str���� = str���� & "|"
    End If
    If strFY2 <> "" Then
        str���� = str���� & "|" & cboData.Text & "." & strFY2
    Else
        str���� = str���� & "|"
    End If
    str���� = Mid(str����, 2)
    
    '����б仯(�����Ŀ������Դ),������ͼ��
    If str���� <> mtmpItem.���� Then
        mtmpItem.���� = str����
        ItemChange = True
    End If
End Sub

Private Sub cboLocate_Click()
    mtmpItem.���� = cboLocate.ListIndex
    ItemChange = True
End Sub

Private Sub cboStyle_Click()
    mtmpItem.��� = cboStyle.ListIndex
        
    Call SetOptionEnabled
    If Visible Then '����ȱʡֵ
        If chkLine.Enabled And chkLine.Value = 0 Then chkLine.Value = 1
        If chkNode.Enabled And chkNode.Value = 0 Then chkNode.Value = 1
    End If
    
    ItemChange = True
End Sub

Private Sub chkGrid_Click()
    If Visible Then
        mtmpItem.���� = IIF(chkGrid.Value = 1, 1, 0)
        ItemChange = True
    End If
End Sub

Private Sub chkLine_Click()
    If Visible Then
        mtmpItem.���� = chkLine.Value = 1
        ItemChange = True
    End If
End Sub

Private Sub chkNode_Click()
    If Visible Then
        mtmpItem.�Ե� = chkNode.Value = 1
        ItemChange = True
    End If
End Sub

Private Sub chkSample_Click()
    If Visible Then
        cboLocate.Enabled = chkSample.Value = 1
        mtmpItem.���� = IIF(chkSample.Value = 1, 2, 1)
        ItemChange = True
    End If
End Sub

Private Sub cmdBack_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = mtmpItem.����
    cdg.ShowColor
    If Err.Number = 0 Then
        mtmpItem.���� = cdg.Color
        cmdBack.BackColor = cdg.Color
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H3 Or &H400 Or &H200 Or &H10000
    
    cdg.FontName = mtmpItem.����
    cdg.FontSize = mtmpItem.�ֺ�
    cdg.FontBold = mtmpItem.����
    cdg.FontItalic = mtmpItem.б��

    cdg.ShowFont
    If Err.Number = 0 Then
        On Error GoTo 0
        mtmpItem.���� = cdg.FontName
        mtmpItem.�ֺ� = cdg.FontSize
        mtmpItem.���� = cdg.FontBold
        mtmpItem.б�� = cdg.FontItalic
        txtFont.Text = cdg.FontName & "," & cdg.FontSize & IIF(cdg.FontBold, ",����", "") & IIF(cdg.FontItalic, ",б��", "")
        Call SelAll(txtFont)
        txtFont.SetFocus
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdFore_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = mtmpItem.ǰ��
    cdg.ShowColor
    If Err.Number = 0 Then
        mtmpItem.ǰ�� = cdg.Color
        cmdFore.BackColor = cdg.Color
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdFontTitle_Click()
    Dim arrFont As Variant
    
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H3 Or &H400 Or &H200 Or &H10000
    
    arrFont = Split(Split(mstr��ͷ, "|")(1), ",")
    cdg.FontName = arrFont(0)
    cdg.FontSize = Val(arrFont(1))
    cdg.FontBold = Val(arrFont(2)) <> 0
    cdg.FontItalic = Val(arrFont(3)) <> 0

    cdg.ShowFont
    If Err.Number = 0 Then
        On Error GoTo 0
        mstr��ͷ = Split(mstr��ͷ, "|")(0) & "|" & cdg.FontName & "," & cdg.FontSize & "," & IIF(cdg.FontBold, 1, 0) & "," & IIF(cdg.FontItalic, 1, 0)
        mtmpItem.��ͷ = mstr��ͷ & ";������ɫ��" & mstrColor
        txtFontTitle.Text = cdg.FontName & "," & cdg.FontSize & IIF(cdg.FontBold, ",����", "") & IIF(cdg.FontItalic, ",б��", "")
        Call SelAll(txtFontTitle)
        txtFontTitle.SetFocus
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdS_Click()
    Dim i As Long
    Dim strColor As String
    
    For i = 0 To cmdColor.UBound
        If cmdColor(i).BackColor = &H8000000F Then
            Exit For
        End If
        strColor = IIF(strColor = "", "", strColor & "|") & cmdColor(i).BackColor
    Next
    mstrColor = strColor
    txtColor.Text = mstrColor
    txtColor.ToolTipText = mstrColor
    mtmpItem.��ͷ = mstr��ͷ & ";������ɫ��" & mstrColor
    ItemChange = True
    Call cmdC_Click
End Sub

Private Sub picDown_Click()
    picColor.Top = picColor.Top - (cmdColor(0).Height + lblColor(0).Height + 90) * 2
    If picColor.Height + picColor.Top > picBack.Height Then
        picDown.Visible = True
    Else
        picDown.Visible = False
    End If
    If picColor.Top < 0 Then
        picUp.Visible = True
    Else
        picUp.Visible = False
    End If
End Sub

Private Sub picUp_Click()
    picColor.Top = picColor.Top + (cmdColor(0).Height + lblColor(0).Height + 90) * 2
    If picColor.Top > 0 Then picColor.Top = 0
    If picColor.Height + picColor.Top > picBack.Height Then
        picDown.Visible = True
    Else
        picDown.Visible = False
    End If
    If picColor.Top < 0 Then
        picUp.Visible = True
    Else
        picUp.Visible = False
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub txtFont_GotFocus()
    SelAll txtFont
End Sub

Private Sub txtFont_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 And cmdFont.Enabled Then
        Call cmdFont_Click
    End If
End Sub

Private Sub txtFontTitle_GotFocus()
    SelAll txtFontTitle
End Sub

Private Sub txtFontTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 And cmdFontTitle.Enabled Then
        Call cmdFontTitle_Click
    End If
End Sub

Private Sub txtLen_Change()
        
    txtLen.Text = Val(txtLen.Text)
    If txtLen.Text < UpDown.Min Then txtLen.Text = UpDown.Min
    If txtLen.Text > UpDown.Max Then txtLen.Text = UpDown.Max
    UpDown.Value = txtLen.Text
    If Visible Then
        mtmpItem.���� = chkLabel.Value & "|" & cboLabel.ListIndex & "|" & chklaLine.Value & "|" & txtLen.Text
        ItemChange = True
    End If
End Sub

Private Sub txtLen_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub txtTitle_Change()
    Dim arrFont As Variant
    
    cmdFontTitle.Enabled = txtTitle.Text <> ""
    txtFontTitle.Enabled = txtTitle.Text <> ""
    
    If Visible Then
        If txtTitle.Text <> "" Then
            If mstr��ͷ = "" Then
                mstr��ͷ = txtTitle.Text & "|����,9,0,0"
            Else
                mstr��ͷ = txtTitle.Text & "|" & Split(mstr��ͷ, "|")(1)
            End If
        Else
            mtmpItem.��ͷ = ""
        End If
        mtmpItem.��ͷ = mstr��ͷ & ";������ɫ��" & mstrColor
        If mstr��ͷ <> "" Then
            arrFont = Split(Split(mstr��ͷ, "|")(1), ",")
            txtFontTitle.Text = arrFont(0) & "," & Val(arrFont(1)) & IIF(Val(arrFont(2)) <> 0, ",����", "") & IIF(Val(arrFont(3)) <> 0, ",б��", "")
        Else
            txtFontTitle.Text = ""
        End If
        ItemChange = True
    End If
End Sub

Private Sub txtTitle_GotFocus()
    Call SelAll(txtTitle)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call PressKey(vbKeyTab)
    Else
        If InStr("'|,;", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strData As String, i As Long
    Dim arrFont As Variant, varTemp As Variant
    
    mblnOK = False
    Call CboSetWidth(cboStyle.hwnd, 3400)
    Call CboSetHeight(cboStyle, Screen.Height)
    Call CopyItem(mtmpItem, mobjItem)
    
    '������Դ
    For i = 1 To mobjDatas.count
        cboData.AddItem mobjDatas(i).����
    Next
    If mtmpItem.���� <> "" Then
        Call GetChartDataName(mtmpItem.����, , , , strData)
        cboData.ListIndex = GetCboIndex(cboData, strData)
    End If
        
    'ͼ����ʽ
    cboStyle.AddItem "ɢ��ͼ(��һX,Y��������)"
    cboStyle.AddItem "����ͼ"
    cboStyle.AddItem "����ͼ"
    cboStyle.AddItem "��ͼ"
    cboStyle.AddItem "���ͼ"
    cboStyle.AddItem "���ͼ"
    cboStyle.AddItem "�ɼ�ͼ(�̸�,�̵�)"
    cboStyle.AddItem "�ɼ�ͼ(�̸�,�̵�,����,����)"
    cboStyle.AddItem "�ɼ�ͼ(������ͼ:�̸�,�̵�,����,����)"
    cboStyle.AddItem "����ͼ"
    cboStyle.AddItem "�״�ͼ"
    cboStyle.AddItem "����״�ͼ"
    cboStyle.AddItem "����ͼ"
    Call CboSetIndex(cboStyle.hwnd, mtmpItem.���)
    
    cboLabel.AddItem "None", 0
    cboLabel.AddItem "3dOut", 1
    cboLabel.AddItem "3dIn", 2
    cboLabel.AddItem "Shadow", 3
    cboLabel.AddItem "Plain", 4
    cboLabel.AddItem "EtchedIn", 5
    cboLabel.AddItem "EtchedOut", 6
    cboLabel.AddItem "FrameIn", 7
    cboLabel.AddItem "FrameOut", 8
    cboLabel.AddItem "Bevel", 9
    cboLabel.ListIndex = 0
    
    If mtmpItem.���� <> "" Then
        varTemp = Split(mtmpItem.����, "|")
        chkLabel.Value = varTemp(0)
        cboLabel.ListIndex = varTemp(1)
        chklaLine.Value = varTemp(2)
        txtLen.Text = Val(varTemp(3))
        UpDown.Value = Val(varTemp(3))
    End If
    
    txtWidth.Text = Val(mtmpItem.����)
    '����
    If mtmpItem.��ͷ <> "" Then
        If InStr(mtmpItem.��ͷ, ";������ɫ��") > 0 Then
            mstr��ͷ = Mid(mtmpItem.��ͷ, 1, InStr(mtmpItem.��ͷ, ";������ɫ��") - 1)
            mstrColor = Mid(Replace(mtmpItem.��ͷ, mstr��ͷ, ""), 7)
            txtColor.Text = mstrColor
            txtColor.ToolTipText = mstrColor
        Else
            mstr��ͷ = mtmpItem.��ͷ
        End If
    End If
    If mstr��ͷ <> "" Then
        txtTitle.Text = Split(mstr��ͷ, "|")(0)
        arrFont = Split(Split(mstr��ͷ, "|")(1), ",")
        txtFontTitle.Text = arrFont(0) & "," & Val(arrFont(1)) & IIF(Val(arrFont(2)) <> 0, ",����", "") & IIF(Val(arrFont(3)) <> 0, ",б��", "")
    End If
            
    'ͼ������
    txtFont.Text = mtmpItem.���� & "," & mtmpItem.�ֺ� & IIF(mtmpItem.����, ",����", "") & IIF(mtmpItem.б��, ",б��", "")
            
    'ͼ����ɫ
    cmdFore.BackColor = mtmpItem.ǰ��
    cmdBack.BackColor = mtmpItem.����
    
    'ͼ��
    chkSample.Value = IIF(mtmpItem.���� <= 1, 0, 1)
    cboLocate.Enabled = chkSample.Value = 1
    cboLocate.AddItem "1-����"
    cboLocate.AddItem "2-����"
    cboLocate.AddItem "3-����"
    cboLocate.AddItem "4-����"
    cboLocate.AddItem "5-���½�"
    cboLocate.AddItem "6-���½�"
    'cboLocate.AddItem "7-���Ͻ�"
    'cboLocate.AddItem "8-���Ͻ�"
    Call CboSetIndex(cboLocate.hwnd, mtmpItem.����)
        
    '������ʽ������λ��,��άЧ��|XY�ụ��
    '��άЧ��
    chkFormat(0).Value = IIF(Val(Mid(Format(mtmpItem.��ʽ, "00"), 1, 1)) = 0, 0, 1)
    'XY�ụ��
    chkFormat(1).Value = IIF(Val(Mid(Format(mtmpItem.��ʽ, "00"), 2, 1)) = 0, 0, 1)
        
    '������
    chkGrid.Value = IIF(mtmpItem.���� <> 0, 1, 0)
    chkLine.Value = IIF(mtmpItem.����, 1, 0)
    chkNode.Value = IIF(mtmpItem.�Ե�, 1, 0)
                
    '���ÿ�ѡ��
    Call SetOptionEnabled
    
    ItemChange = False
    Call SetChartStyleAndData(Chart, mtmpItem)
    Call SetColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    Set mtmpItem = Nothing
    For i = 1 To cmdColor.UBound
        Unload cmdColor(i)
        Unload lblColor(i)
    Next
    mstr��ͷ = ""
    mstrColor = ""
End Sub

Private Function CheckInput() As Boolean
    If cboFX.Text = "" Then
        MsgBox "��ָ����ֵ�ֶ���Դ��", vbInformation, App.Title
        cboFX.SetFocus: Exit Function
    End If
    If cboFS.Text = "" Then
        MsgBox "��ָ����1�����ֶ���Դ��", vbInformation, App.Title
        cboFS.SetFocus: Exit Function
    End If
    If cboFY.Text = "" Then
        MsgBox "��ָ����1ֵ�ֶ���Դ��", vbInformation, App.Title
        cboFY.SetFocus: Exit Function
    End If
    If cboFX.Text = cboFY.Text Then
        MsgBox "��1ֵ�ֶ����ֵ�ֶβ�����ͬ��", vbInformation, App.Title
        cboFY.SetFocus: Exit Function
    End If
    If cboFX.Text = cboFY2.Text Then
        MsgBox "��2ֵ�ֶ����ֵ�ֶβ�����ͬ��", vbInformation, App.Title
        cboFY2.SetFocus: Exit Function
    End If
    If cboFS2.Text = cboFS.Text Then
        MsgBox "��ָ����2�����ֶ���Դ��", vbInformation, App.Title
        cboFS2.SetFocus: Exit Function
    End If
    If cboFY.Text = cboFY2.Text Then
        MsgBox "��2ֵ�ֶ����1ֵ�ֶβ�����ͬ��", vbInformation, App.Title
        cboFY2.SetFocus: Exit Function
    End If
    If cboFY2.Text <> "" Or cboFS2.Text <> "" Then
        If cboFY2.Text = "" Then
            MsgBox "��2�����ֶβ�Ϊ�գ���2ֵ����Ϊ��", vbInformation, App.Title
            cboFY2.SetFocus: Exit Function
        End If
        If cboFS2.Text = "" Then
            MsgBox "��2ֵ��Ϊ�գ���2�����ֶβ���Ϊ��", vbInformation, App.Title
            cboFS2.SetFocus: Exit Function
        End If
    End If
    CheckInput = True
End Function

Private Sub txtWidth_Change()

    If Visible Then
        If Val(txtWidth.Text) > 50 Then
            MsgBox "������ϸ���ܳ���50��", vbInformation, Me.Caption
            txtWidth.Text = 50
        End If
        mtmpItem.���� = Val(txtWidth.Text)
        ItemChange = True
    End If
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub UpDown_Change()
    txtLen.Text = UpDown.Value
End Sub

Private Sub SetColor()
'����������ɫ
'�������е���ɫ
    Dim varColor As Variant
    Dim lngNum As Long, i As Long, lngMax As Long
    
    For i = 1 To cmdColor.UBound
        Unload cmdColor(i)
        Unload lblColor(i)
    Next
        
    varColor = Split(mstrColor, "|")
    lngMax = UBound(varColor) / 5
    If Fix(UBound(varColor) / 5) <> UBound(varColor) / 5 Then lngMax = Fix(lngMax) + 1
    If lngMax < 2 Then lngMax = 2
    '�ȼ��ؿؼ�
    For i = 0 To lngMax * 5 - 1
        If i = 0 Then
            cmdColor(0).Visible = True
            cmdColor(0).BackColor = &H8000000F
            cmdColor(0).ToolTipText = "���е���ɫ˳������������ԴSQL��ѯ����������˳����û�����ã���������ʾĬ����ɫ"
            lblColor(0).Visible = True
            lblColor(0).Move cmdColor(0).Left + 120, cmdColor(0).Top + cmdColor(0).Height + 20
        Else
            Load cmdColor(i)
            Load lblColor(i)
            cmdColor(i).Visible = True
            cmdColor(i).BackColor = &H8000000F
            cmdColor(i).ToolTipText = cmdColor(0).ToolTipText
            lblColor(i).Visible = True
            lblColor(i).Caption = i + 1
            lngNum = Fix(i / 5)
            If lngNum = i / 5 Then
                cmdColor(i).Move cmdColor(0).Left, cmdColor(0).Top + (cmdColor(0).Height + lblColor(0).Height + 90) * lngNum, cmdColor(0).Width, cmdColor(0).Height
            Else
                cmdColor(i).Move cmdColor(i - 1).Left + cmdColor(i - 1).Width + 60, cmdColor(i - 1).Top, cmdColor(0).Width, cmdColor(0).Height
            End If
            lblColor(i).Move cmdColor(i).Left + IIF(i < 9, 120, 80), cmdColor(i).Top + cmdColor(i).Height + 20
        End If
    Next
    picUp.Visible = False
    If lngMax > 2 Then
        picDown.Visible = True
        picColor.Height = (cmdColor(0).Height + lblColor(0).Height + 90) * lngMax
    Else
        picDown.Visible = False
        cmdDelte.Enabled = False
        picColor.Height = (cmdColor(0).Height + lblColor(0).Height + 90) * 2
    End If
    '�ؼ�������������ð�ť��ɫ
    For i = LBound(varColor) To UBound(varColor)
        cmdColor(i).BackColor = varColor(i)
    Next
    picColor.Width = picBack.Width
End Sub
