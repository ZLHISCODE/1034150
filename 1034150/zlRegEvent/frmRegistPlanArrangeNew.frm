VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlanArrangeNew 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20100
   Icon            =   "frmRegistPlanArrangeNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11340
   ScaleWidth      =   20100
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chk������� 
      Caption         =   "������������"
      Height          =   285
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   1650
   End
   Begin VB.PictureBox picBaseBack 
      BorderStyle     =   0  'None
      Height          =   9180
      Left            =   240
      ScaleHeight     =   9180
      ScaleWidth      =   8610
      TabIndex        =   5
      Top             =   480
      Width           =   8610
      Begin VB.PictureBox picBase 
         BorderStyle     =   0  'None
         Height          =   7725
         Left            =   -120
         ScaleHeight     =   7725
         ScaleWidth      =   8130
         TabIndex        =   6
         Top             =   0
         Width           =   8130
         Begin VB.OptionButton opt��Чʱ�� 
            Caption         =   "ָ��ʱ��"
            Height          =   180
            Index           =   1
            Left            =   2040
            TabIndex        =   40
            Top             =   6600
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.Frame Frame1 
            Caption         =   "������Ϣ"
            Height          =   1455
            Left            =   60
            TabIndex        =   27
            Top             =   105
            Width           =   7890
            Begin VB.TextBox txt�ű� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   660
               MaxLength       =   5
               TabIndex        =   34
               Top             =   270
               Width           =   960
            End
            Begin VB.ComboBox cboItem 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   3900
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   675
               Width           =   2580
            End
            Begin VB.ComboBox cboDoctor 
               Height          =   300
               Left            =   660
               TabIndex        =   32
               Top             =   1035
               Width           =   2400
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   660
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   660
               Width           =   2400
            End
            Begin VB.CheckBox chk���� 
               Caption         =   "�Һ�ʱ���뽨����"
               Height          =   195
               Left            =   3870
               TabIndex        =   30
               Top             =   1080
               Width           =   1845
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   3900
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   270
               Width           =   2595
            End
            Begin VB.CheckBox chk��ſ��� 
               Caption         =   "��ſ���"
               Height          =   255
               Left            =   1750
               TabIndex        =   28
               Top             =   293
               Width           =   1095
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "�ű�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   210
               TabIndex        =   39
               Top             =   330
               Width           =   390
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "����"
               Height          =   180
               Left            =   240
               TabIndex        =   38
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "��Ŀ"
               Height          =   180
               Left            =   3480
               TabIndex        =   37
               Top             =   750
               Width           =   360
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "ҽ��"
               Height          =   180
               Left            =   240
               TabIndex        =   36
               Top             =   1080
               Width           =   360
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "����"
               Height          =   180
               Left            =   3465
               TabIndex        =   35
               Top             =   330
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Ӧ������:"
            Height          =   2610
            Left            =   60
            TabIndex        =   21
            Top             =   3840
            Width           =   7875
            Begin VB.OptionButton opt���� 
               Caption         =   "������"
               Height          =   180
               Index           =   0
               Left            =   1020
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "ָ������"
               Height          =   180
               Index           =   1
               Left            =   2010
               TabIndex        =   24
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "��̬����"
               Height          =   180
               Index           =   2
               Left            =   3180
               TabIndex        =   23
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "ƽ������"
               Height          =   180
               Index           =   3
               Left            =   4335
               TabIndex        =   22
               Top             =   0
               Width           =   1020
            End
            Begin MSComctlLib.ListView lvwDept 
               Height          =   2190
               Left            =   150
               TabIndex        =   26
               Top             =   300
               Width           =   7605
               _ExtentX        =   13414
               _ExtentY        =   3863
               View            =   2
               Arrange         =   2
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ӧ��ʱ��"
            Height          =   2070
            Left            =   60
            TabIndex        =   12
            Top             =   1635
            Width           =   7875
            Begin VB.OptionButton opt�� 
               Caption         =   "ÿ��(&D)"
               Height          =   315
               Left            =   225
               TabIndex        =   17
               Top             =   285
               Width           =   960
            End
            Begin VB.OptionButton opt�� 
               Caption         =   "ÿ��(&W)"
               Height          =   315
               Left            =   225
               TabIndex        =   16
               Top             =   630
               Width           =   930
            End
            Begin VB.ComboBox cbo�� 
               Height          =   300
               Left            =   1170
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   270
               Width           =   1110
            End
            Begin VB.TextBox txt�޺� 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   3030
               MaxLength       =   5
               TabIndex        =   14
               Top             =   270
               Width           =   1215
            End
            Begin VB.TextBox txt��Լ 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   5145
               MaxLength       =   5
               TabIndex        =   13
               Top             =   270
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid vsPlan 
               Height          =   1275
               Left            =   1170
               TabIndex        =   18
               Top             =   660
               Width           =   6600
               _cx             =   11642
               _cy             =   2249
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   4
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   300
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmRegistPlanArrangeNew.frx":06EA
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
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "�޺�"
               Height          =   180
               Left            =   2595
               TabIndex        =   20
               Top             =   330
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "��Լ"
               Height          =   180
               Left            =   4710
               TabIndex        =   19
               Top             =   330
               Width           =   360
            End
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   990
            TabIndex        =   11
            Top             =   7020
            Width           =   2370
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   990
            TabIndex        =   10
            Top             =   7410
            Width           =   2370
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   5565
            TabIndex        =   9
            Top             =   7020
            Width           =   2370
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   5580
            TabIndex        =   8
            Top             =   7410
            Width           =   2370
         End
         Begin VB.OptionButton opt��Чʱ�� 
            Caption         =   "����ִ��"
            Height          =   360
            Index           =   0
            Left            =   990
            TabIndex        =   7
            Top             =   6525
            Width           =   1530
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   300
            Left            =   5565
            TabIndex        =   41
            Top             =   6555
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   117112835
            CurrentDate     =   401769
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   3120
            TabIndex        =   42
            Top             =   6555
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   117112835
            CurrentDate     =   38091
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   0
            Left            =   300
            TabIndex        =   48
            Top             =   7080
            Width           =   540
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����ʱ��"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   7470
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Index           =   2
            Left            =   4950
            TabIndex        =   46
            Top             =   7080
            Width           =   540
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "���ʱ��"
            Height          =   180
            Index           =   3
            Left            =   4785
            TabIndex        =   45
            Top             =   7470
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�ƻ�ʱ��"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   44
            Top             =   6600
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   5
            Left            =   5265
            TabIndex        =   43
            Top             =   6615
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   8775
      TabIndex        =   2
      Top             =   1530
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8775
      TabIndex        =   1
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8775
      TabIndex        =   0
      Top             =   1005
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   8700
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8655
      _Version        =   589884
      _ExtentX        =   15266
      _ExtentY        =   15346
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmRegistPlanArrangeNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit 'Ҫ���������
'Private mstr�ƻ�ID As String, mlng����ID As Long, mblnSucces As Boolean, mblnFirst As Boolean
'Private mlngModule As Long, mstrPrivs As String
'Private mblnActive As Boolean
'Private Enum mPageIndex
'    EM_�ƻ� = 0
'    EM_ʱ�� = 1
'End Enum
'Private mfrmTime As frmResistPlanTimeSet    '�ƻ�ʱ����
'Private mrsRegOldData As ADODB.Recordset '�������ݼ�����,ԭʼ�ҺŰ���
'Private mrsRegNewData As ADODB.Recordset '�������ݼ����� �������ú�İ���
'Private mrsRegHistory As ADODB.Recordset '���ιҺŵ����ݼ�
'Private mblnChangeByCode As Boolean
'Public Enum mRegEditType
'    ed_�ƻ����� = 0
'    Ed_�����޸� = 1
'    Ed_����ɾ�� = 2
'    Ed_������� = 3
'    Ed_����ȡ�� = 4
'    ed_���Ų��� = 5
'End Enum
'Private Enum midxTxt
'    idx_������ = 0
'    idx_����ʱ�� = 1
'    idx_����� = 2
'    idx_���ʱ�� = 3
'End Enum
'Private mEditType As mRegEditType
'Private mstr����ID As String
'Private mblnCboClick As Boolean     '�����cbo��keypress�¼������˵����б���API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
''                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
'Private mrsDoctor As ADODB.Recordset
'
'
'Private Type PlanInfo               '���Ÿı���Ҫ�Աȵ���Ϣ
'    str�Ű�         As String       '�Ű���Ϣ
'    str�޺�         As String       '�޺���Ϣ
'    bln���         As Boolean      '�Ƿ���ſ���
'    blnʱ���       As Boolean      '�Ƿ�������ʱ���
'End Type
'Private mPlanInfo     As PlanInfo '����ʱ���ڱ���ԭʼ������Ϣ  �޸�ʱ ����ԭʼ�ļƻ���Ϣ �ڱ���ʱ �Ƚ���Ӧ��Ϣ
'Private Enum mPgIndex
'    Pg_�ƻ����� = 1
'    Pg_�ƻ�ʱ�� = 2
'End Enum
'Private Sub InitPage()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ʼ��ҳ��ؼ�
'    '����:���˺�
'    '����:2009-09-09 11:01:36
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim i As Long, objItem As TabControlItem, objForm As Object
'    Err = 0: On Error GoTo ErrHand:
'
'    Set objItem = tbPage.InsertItem(mPgIndex.Pg_�ƻ�����, "�ƻ�����", picBaseBack.hWnd, 0)
'    objItem.Tag = mPgIndex.Pg_�ƻ�����
'
'    Set mfrmTime = New frmResistPlanTimeSet
'    Set objItem = tbPage.InsertItem(mPgIndex.Pg_�ƻ�ʱ��, "ʱ������", mfrmTime.hWnd, 0)
'    objItem.Tag = mPgIndex.Pg_�ƻ�ʱ��
'     With tbPage
'        tbPage.Item(0).Selected = True
'        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
'        .PaintManager.BoldSelected = True
'        .PaintManager.Layout = xtpTabLayoutAutoSize
'        .PaintManager.StaticFrame = False
'        .PaintManager.ClientFrame = xtpTabFrameBorder
'    End With
'    Exit Sub
'ErrHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Sub
'
'
'Public Function ShowCard(ByVal mfrmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
'    ByVal EditType As mRegEditType, Optional lng����ID As Long, Optional ByVal str�ƻ�Id As String = "") As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ʾ��Ҫ�޸ĵļƻ�����
'    '���:mfrmMain-���õ�������
'    '     lngModule-ģ���
'    '     strPrivs-Ȩ�޴�
'    '     EditType-�༭������
'    '     lng����ID-�ҺŰ���ID.
'    '     str�ƻ�Id-����ʱΪ��,����,����Ϊָ���ļƻ�ID
'    '����:
'    '����:�ɹ�,����true,���򷵻�False
'    '����:���˺�
'    '����:2009-09-14 14:31:59
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    mEditType = EditType: mlngModule = lngModule: mstrPrivs = strPrivs: mstr�ƻ�ID = str�ƻ�Id: mblnSucces = False: mlng����ID = lng����ID
'    Me.Show 1, mfrmMain
'    ShowCard = mblnSucces
'End Function
'
'Private Function LoadData() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:���ؼƻ�����������Ϣ
'    '����:���˺�
'    '����:2009-09-14 14:40:46
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim rsTemp          As New ADODB.Recordset
'    Dim strSQL          As String
'    Dim i               As Long
'    Dim rs�޺�          As ADODB.Recordset
'    Dim strTemp         As String
'    Dim blnÿ��         As Boolean
'    Dim bln�޺�         As Boolean
'    Dim str�޺�         As String
'    Dim bln��Լ         As Boolean
'    Dim str��Լ         As String
'    Err = 0: On Error GoTo ErrHand:
'
'    '���ذ���
'    If mEditType = ed_�ƻ����� Then
'       '��������
'        strSQL = " " & _
'        "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,A.��ĿID as �ƻ���ĿID,   A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id ,   " & _
'        "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,A.Ĭ��ʱ�μ��, " & _
'        "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ����,NULL��as ��Чʱ��,'3000-01-01 00:00:00' as ʧЧʱ�� ," & _
'        "           NULL as ������,NULL as ����ʱ��,NULL �����,NULL ���ʱ��" & _
'        "   From �ҺŰ��� A,�շ���ĿĿ¼ B,�ҺŰ��żƻ� C,���ű� D " & _
'        "   Where A.Id=C.����ID(+) And A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
'        "         And A.Id=[1]"
'    Else
'         '������
'        strSQL = " " & _
'        "Select a.����id, a.Id As �ƻ�id, a.����, �ƻ���Ŀid, a.����, a.����id, a.��Ŀid, a.ҽ������, a.ҽ��id,   a.����, a.��һ, a.�ܶ�, a.����," & _
'        "  a.����, a.����, a.����, a.��������, a.���﷽ʽ, a.��ſ���, a.��ʼʱ��, a.��ֹʱ��, b.���� As ��Ŀ, d.���� As ����, ��Чʱ��, a.ʧЧʱ��, a.������, a.����ʱ��," & _
'        " a.����� , ���ʱ��,A.Ĭ��ʱ�μ��" & _
'        " From (Select c.����id, c.Id, a.����, Nvl(c.��Ŀid, a.��Ŀid) As �ƻ���Ŀid, c.����, a.����id, Nvl(c.��Ŀid, a.��Ŀid) As ��Ŀid, C.ҽ������, C.ҽ��id," & _
'        "       c.����, c.��һ, c.�ܶ�, c.����, c.����, c.����, c.����, a.��������, c.���﷽ʽ, c.��ſ���, a.��ʼʱ��, a.��ֹʱ��, Nvl(C.Ĭ��ʱ�μ��,5) as Ĭ��ʱ�μ��," & _
'        "      To_Char(c.��Чʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Чʱ��, To_Char(c.ʧЧʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʧЧʱ��, c.������," & _
'        "      To_Char(c.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, c.�����, To_Char(c.���ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ���ʱ��" & _
'        " From �ҺŰ��� A, �ҺŰ��żƻ� C " & _
'        " Where a.Id = c.����id) A, �շ���ĿĿ¼ B, ���ű� D " & _
'        " Where a.��Ŀid = b.Id(+) And a.����id = d.Id(+) " & _
'        "  and a.id=[2]"
'    End If
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mstr�ƻ�ID))
'    If rsTemp.EOF Then
'        If mEditType = ed_�ƻ����� Then
'            MsgBox "ע��:" & vbCrLf & _
'                   "    �ҺŰ��ſ����Ѿ�������ɾ��,�����ٽ��мƻ�����", vbInformation + vbOKOnly, gstrSysName
'        Else
'            MsgBox "ע��:" & vbCrLf & _
'                   "    �Һżƻ����ſ����Ѿ�������ɾ��,����!", vbInformation + vbOKOnly, gstrSysName
'        End If
'        Exit Function
'    End If
'    If mEditType = ed_�ƻ����� Then
'        strSQL = "Select ������Ŀ,�޺���,  ��Լ�� From  �ҺŰ������� where ����ID=[1]       "
'    Else
'        strSQL = "Select ������Ŀ,�޺���,  ��Լ�� From  �Һżƻ����� where �ƻ�ID=[2]       "
'    End If
'    Set rs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mstr�ƻ�ID))
'
'    '�������һЩ����
'    If mEditType = Ed_�����޸� And Nvl(rsTemp!���ʱ��) <> "" Then
'            MsgBox "ע��:" & vbCrLf & _
'                   "    �Һżƻ������Ѿ����������,�����ٽ��мƻ��޸ģ�", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
'    End If
'    If mEditType = Ed_����ɾ�� And Nvl(rsTemp!���ʱ��) <> "" Then
'            MsgBox "ע��:" & vbCrLf & _
'                   "    �Һżƻ������Ѿ����������,�����ٽ��мƻ�ɾ����", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
'    End If
'
'    If mEditType = Ed_������� And Nvl(rsTemp!���ʱ��) <> "" Then
'            MsgBox "ע��:" & vbCrLf & _
'                   "    �Һżƻ������Ѿ����������,�����ٽ��мƻ���ˣ�", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
'    End If
'
'    If mEditType = Ed_����ȡ�� And Nvl(rsTemp!���ʱ��) = "" Then
'            MsgBox "ע��:" & vbCrLf & _
'                   "    �Һżƻ������Ѿ�������ȡ�����,�����ٽ��мƻ����ȡ����", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
'    End If
'
'    '�������ݵ��ؼ���
'    txt�ű�.Text = Nvl(rsTemp!����)
'    cbo����.AddItem Nvl(rsTemp!����): cbo����.ListIndex = cbo����.NewIndex
'    chk��ſ���.Value = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, 1, 0)
'    '��ȡ�İ��Ż��߼ƻ��Ƿ���ſ���
'    mPlanInfo.bln��� = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, True, False)
'
'    chk����.Value = IIf(Val(Nvl(rsTemp!��������)) = 1, 1, 0)
'
'
'    txtEdit(midxTxt.idx_������).Text = Nvl(rsTemp!������)
'    txtEdit(midxTxt.idx_����ʱ��).Text = Nvl(rsTemp!����ʱ��)
'    If mEditType = ed_�ƻ����� Then
'        txtEdit(midxTxt.idx_������) = UserInfo.����
'        txtEdit(midxTxt.idx_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
'    End If
'    txtEdit(midxTxt.idx_�����) = Nvl(rsTemp!�����)
'    txtEdit(midxTxt.idx_���ʱ��) = Nvl(rsTemp!���ʱ��)
'    If mEditType = Ed_������� Then
'        txtEdit(midxTxt.idx_�����) = UserInfo.����
'        txtEdit(midxTxt.idx_���ʱ��) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
'    End If
'
'    With cbo����
'        .AddItem Nvl(rsTemp!����): .ItemData(.NewIndex) = Val(Nvl(rsTemp!����ID)): .ListIndex = .NewIndex
'    End With
'    With cboItem
'         If mEditType = Ed_�����޸� Or mEditType = ed_�ƻ����� Then
'            zlControl.CboSetText cboItem, rsTemp!��Ŀ
'        Else
'            .AddItem Nvl(rsTemp!��Ŀ): .ItemData(.NewIndex) = Val(Nvl(rsTemp!��ĿID)): .ListIndex = .NewIndex
'        End If
'
'    End With
'    With cboDoctor
'       If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then
'          LoadDoctor
'          zlControl.CboSetText cboDoctor, Nvl(rsTemp!ҽ������)
'        Else
'            .AddItem Nvl(rsTemp!ҽ������): .ItemData(.NewIndex) = Val(Nvl(rsTemp!ҽ��ID)): .ListIndex = .NewIndex
'        End If
'    End With
'
'    '����ԭʼ���ݵ����ݼ�
'     With mrsRegOldData
'        Set mrsRegOldData = New ADODB.Recordset
'        mrsRegOldData.Fields.Append "ID", adBigInt, 18
'        mrsRegOldData.Fields.Append "������Ŀ", adVarChar, 20
'        mrsRegOldData.Fields.Append "�޺���", adBigInt, 10
'        mrsRegOldData.Fields.Append "��Լ��", adBigInt, 18
'        mrsRegOldData.Fields.Append "��ſ���", adBigInt, 18
'        mrsRegOldData.CursorLocation = adUseClient
'        mrsRegOldData.LockType = adLockOptimistic
'        mrsRegOldData.CursorType = adOpenStatic
'        mrsRegOldData.Open
'
'
'        rs�޺�.Filter = 0
'        If rs�޺�.RecordCount > 0 Then rs�޺�.MoveFirst
'        Do While Not rs�޺�.EOF
'            With mrsRegOldData
'                .AddNew
'                !ID = Val(mstr�ƻ�ID)
'                !������Ŀ = Nvl(rs�޺�!������Ŀ)
'                !�޺��� = Val(Nvl(rs�޺�!�޺���))
'                !��Լ�� = Val(Nvl(rs�޺�!��Լ��))
'                !��ſ��� = Val(Nvl(rsTemp!��ſ���))
'                .Update
'            End With
'            rs�޺�.MoveNext
'        Loop
'    End With
'
'    Call LoadRegHistory
'    '---------------------------------------------------
'    '�ж� ÿ�հ��� �޺��� ��Լ�� ���Ƿ�һ��
'    '---------------------------------------------------
'    rs�޺�.Filter = 0
'    If rs�޺�.RecordCount > 0 Then rs�޺�.MoveFirst
'
'    blnÿ�� = Nvl(rsTemp!����) <> Nvl(rsTemp!��һ) Or Nvl(rsTemp!����) <> Nvl(rsTemp!�ܶ�) _
'        Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) _
'        Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Or Nvl(rsTemp!����) <> Nvl(rsTemp!����)
'
'    If blnÿ�� = False Then
'             rs�޺�.Filter = "������Ŀ='����'"
'             If Not rs�޺�.EOF Then
'                str�޺� = Nvl(rs�޺�!�޺���)
'                str��Լ = Nvl(rs�޺�!��Լ��)
'             End If
'            For i = 1 To 6
'                strTemp = Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
'                rs�޺�.Filter = "������Ŀ='" & "��" & strTemp & "'"
'                If Not rs�޺�.EOF Then
'                    bln�޺� = Nvl(rs�޺�!�޺���) = str�޺�
'                    bln��Լ = Nvl(rs�޺�!��Լ��) = str��Լ
'                    If bln��Լ = False Or bln�޺� = False Then Exit For
'                End If
'            Next
'          blnÿ�� = True
'         If bln�޺� And bln��Լ Then blnÿ�� = False
'
'    End If
'
'    If blnÿ�� Or mrsRegHistory.RecordCount > 0 Then
'        'ÿ��
'        opt��.Value = True:
'        txt�޺�.Enabled = False: txt��Լ.Enabled = False
'        With vsPlan
'            For i = 1 To 7
'                '��֪ʲôԭ��,��.colkey(i)����,Ҫ���ĳ�������.
'                strTemp = "��" & Replace(.ColKey(i), "����", "��")
'                .TextMatrix(1, i) = Nvl(rsTemp.Fields(strTemp))
'                rs�޺�.Filter = "������Ŀ='" & strTemp & "'"
'                If Not rs�޺�.EOF Then
'
'                    .TextMatrix(2, i) = Nvl(rs�޺�!�޺���)
'                    .TextMatrix(3, i) = Nvl(rs�޺�!��Լ��)
'                End If
'            Next
'        End With
'    Else
'        'ÿ��
'        opt��.Value = True:  cbo��.ListIndex = GetCboIndex(cbo��, Nvl(rsTemp!����)): cbo��.Enabled = True
'        If rs�޺�.RecordCount <> 0 Then rs�޺�.MoveFirst
'        If rs�޺�.EOF = False Then
'            txt�޺�.Text = Nvl(rs�޺�!�޺���)
'            txt��Լ.Text = Nvl(rs�޺�!��Լ��)
'        End If
'    End If
'
'     '------------------------------
'    '��ȡ�޸Ļ�������ǰ�� ʱ��κ� �޺���
'    '�����ڱ���ʱ �Ա��޺���Լ����ſ����Լ�ʱ����Ƿ����˱仯
'    '��������˱仯����Ҫ��ʾ  ����Ա��������ʱ����Ϣ
'    '------------------------------
'   mPlanInfo.str�Ű� = ""
'   mPlanInfo.str�޺� = ""
'
'    If blnÿ�� = False Or mrsRegHistory.RecordCount > 0 Then
'        For i = 1 To 7
'             mPlanInfo.str�Ű� = mPlanInfo.str�Ű� & ",'" & Trim(cbo��.Text) & "'"
'             mPlanInfo.str�޺� = mPlanInfo.str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
'             mPlanInfo.str�޺� = mPlanInfo.str�޺� & "," & Val(txt�޺�.Text) & "," & Val(txt��Լ.Text)
'        Next
'    Else
'        For i = 1 To vsPlan.Cols - 1
'            mPlanInfo.str�Ű� = mPlanInfo.str�Ű� & ",'" & Trim(vsPlan.TextMatrix(1, i)) & "'"
'            If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
'                mPlanInfo.str�޺� = mPlanInfo.str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
'                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
'                     mPlanInfo.str�޺� = mPlanInfo.str�޺� & ",0,0"
'                Else
'                     mPlanInfo.str�޺� = mPlanInfo.str�޺� & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Val(Trim(vsPlan.TextMatrix(3, i)))
'                End If
'            End If
'        Next
'    End If
'    If mPlanInfo.str�޺� <> "" Then mPlanInfo.str�޺� = Mid(mPlanInfo.str�޺�, 2)
'    '-------------------------------
'
'    If IsNull(rsTemp!��Чʱ��) Then
'        dtpBegin.Value = Format(zlGetNextWeekDate, "yyyy-mm-dd HH:MM:SS")
'    Else
'        dtpBegin.Value = CDate(Nvl(rsTemp!��Чʱ��))
'    End If
'    dtpEndDate.Value = CDate(Nvl(rsTemp!ʧЧʱ��, "3000-01-01"))
'
'    Select Case Val(Nvl(rsTemp!���﷽ʽ))     '0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
'        Case 0  '"������"
'            opt����(0).Value = True
'        Case 1  ' "ָ������"
'            opt����(1).Value = True
'        Case 2 '"��̬����"
'            opt����(2).Value = True
'        Case 3 ' "ƽ������"
'            opt����(3).Value = True
'    End Select
'
'    If mEditType = ed_�ƻ����� Then
'        strSQL = "Select nvl(��Чʱ��,Sysdate) as ��Чʱ�� ,nvl(ʧЧʱ��,to_date('3000-01-01','yyyy-mm-dd')) as ʧЧʱ�� From �ҺŰ��żƻ� where ID=(Select Max(ID) From �ҺŰ��żƻ� where ����ID=[1]) "
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
'        If Not rsTemp.EOF Then
'            If Format(rsTemp!ʧЧʱ��, "yyyy-mm-dd") < "3000-01-01" Then
'                '��һ���ƻ�����ֹ����,���Ǳ�������Чʱ��
'                dtpBegin.Value = Format(rsTemp!ʧЧʱ��, "yyyy-mm-dd HH:MM:SS")
'            Else '����һ������Чʱ�����һ��Ϊ׼
'                dtpBegin.Value = zlGetNextWeekDate(Format(rsTemp!��Чʱ��, "yyyy-mm-dd HH:MM:SS"))
'            End If
'        End If
'
'        strSQL = "Select �ű�ID as ID,�������ҡ�From �ҺŰ������� Where �ű�ID=[1]"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
'    Else
'        strSQL = "Select �ƻ�ID as ID,�������ҡ�From �Һżƻ����� Where �ƻ�ID=[2]"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mstr�ƻ�ID))
'    End If
'
'    Do While Not rsTemp.EOF
'        For i = 1 To lvwDept.ListItems.Count
'            If Nvl(rsTemp!��������) = lvwDept.ListItems(i).Text Then
'                lvwDept.ListItems(i).Checked = True
'            End If
'        Next
'        rsTemp.MoveNext
'    Loop
'    If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then mPlanInfo.blnʱ��� = Checkʱ��()
'    If mrsRegHistory.RecordCount > 0 Then opt��.Enabled = False
'    LoadData = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Function
'Private Function InitData() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:���س�ʼ������
'    '����:���˺�
'    '����:2009-09-14 15:50:31
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String, rsTemp As New ADODB.Recordset, i As Long
'
'    Err = 0: On Error GoTo ErrHand:
'
'    strSQL = "Select '    ' ʱ��� From dual Union All  " & _
'             " Select ʱ��� From ʱ���"
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
'    Do While Not rsTemp.EOF
'        cbo��.AddItem rsTemp!ʱ���
'        rsTemp.MoveNext
'    Loop
'
'    With vsPlan
'        .ColComboList(1) = .BuildComboList(rsTemp, "ʱ���")
'        For i = 2 To .Cols - 1
'            .ColComboList(i) = .ColComboList(1)
'        Next
'        .Tag = .ColComboList(1)
'    End With
'
'
'    '��������
'    strSQL = "Select ����,���ơ�From �������� Where (վ��='" & gstrNodeNo & "' Or վ�� is Null) Order by ����"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'    lvwDept.ListItems.Clear
'    For i = 1 To rsTemp.RecordCount
'        lvwDept.ListItems.Add , "D" & Nvl(rsTemp!����), Nvl(rsTemp!����)
'        rsTemp.MoveNext
'    Next
'
'
'    '�Һ���Ŀ
'    If mEditType = Ed_�����޸� Or mEditType = ed_�ƻ����� Then
'        strSQL = "Select ID as ���,���� From �շ���ĿĿ¼ " & _
'            " Where ���='1' And (Sysdate Between ����ʱ�� And ����ʱ�� Or ����ʱ��<Sysdate And ����ʱ�� Is Null)" & _
'            " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
'            " Order by ����"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'
'        If rsTemp.EOF Then
'            MsgBox "û�п��õĹҺ���Ŀ��Ϣ,���ȵ��Һ���Ŀ�����г�ʼ��", vbInformation, gstrSysName
'            Exit Function
'        End If
'
'        cboItem.Clear
'        For i = 1 To rsTemp.RecordCount
'            cboItem.AddItem rsTemp!����
'            cboItem.ItemData(cboItem.NewIndex) = rsTemp!���
'            rsTemp.MoveNext
'        Next
'    End If
'
'    'cmdCancel.Caption = "�˳�(&X)"
'    If mEditType = Ed_������� Then
'        Me.Caption = Me.Caption & "�������"
'    ElseIf mEditType = Ed_����ɾ�� Then
'        Me.Caption = Me.Caption & "����ɾ��"
'        'cmdOK.Caption = "ɾ��(&D)"
'    ElseIf mEditType = Ed_����ȡ�� Then
'        Me.Caption = Me.Caption & "����ȡ�����"
'    ElseIf mEditType = ed_���Ų��� Then
'        cmdOK.Visible = False
'        cmdCancel.Top = cmdOK.Top
'    End If
'
'    InitData = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    SaveErrLog
'End Function
'
'
'
'Private Sub Form_Load()
'    Call InitPage
'    opt��Чʱ��(0).Enabled = True: opt��Чʱ��(1).Enabled = True
'    mblnFirst = True
'End Sub
'Private Sub Form_Activate()
'    If mblnFirst = False Then Exit Sub
'    mblnFirst = False
'    If InitData = False Then Unload Me: Exit Sub
'    If LoadData = False Then Unload Me: Exit Sub
'    Call SetCtrlEnabled
'
'    If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then
'        zlCtlSetFocus chk��ſ���
'    Else
'        zlCtlSetFocus cmdOK
'    End If
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
'End Sub
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = Asc("'") Then KeyAscii = 0
'End Sub
'Private Sub SetCtrlEnabled()
'    '���ÿؼ���Enabled����
'    Dim ctl As Control
'
'    For Each ctl In Me.Controls
'        Select Case UCase(TypeName(ctl))
'        Case "TEXTBOX"
'            ctl.Enabled = False
'            '�޸Ļ��������ƻ�ʱ �����޺š���Լ�ı��� ���޸�
'            If ctl Is Me.txt�޺� Or ctl Is txt��Լ Then
'               ctl.Enabled = mEditType = Ed_�����޸� Or mEditType = ed_�ƻ�����
'            End If
'        Case UCase("ComboBox")
'            If ctl Is cbo�� And mEditType = ed_�ƻ����� Then
'                   ctl.Enabled = opt��.Value = 1
'              ElseIf ctl Is cboItem Or ctl Is cboDoctor Then
'                 '-----------------------------------------------------
'                 'Ϊ�޸Ļ��� ����ģʽʱ ���Ŷ� ��Ŀ��ҽ���ĸ���
'
'                 '------------------------------------------------------
'                   If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then
'                       ctl.Enabled = True
'                   Else
'                       ctl.Enabled = False
'                   End If
'               Else:
'                   ctl.Enabled = False
'               End If
'        Case UCase("ListView")
'            ctl.Enabled = False
'        Case UCase("DTPicker")
'            ctl.Enabled = False
'        Case UCase("optionbutton"), UCase("CheckBox")
'            ctl.Enabled = False
'
'        Case Else
'        End Select
'    Next
'
'    Select Case mEditType
'    Case ed_�ƻ�����, Ed_�����޸�
'        chk��ſ���.Enabled = True
'        txt�޺�.Enabled = IIf(opt��.Value = True, True, False): txt��Լ.Enabled = IIf(opt��.Value = True, True, False)
'        cbo��.Enabled = IIf(opt��.Value = True, True, False)
'        dtpBegin.Enabled = IIf(opt��Чʱ��(0).Value = 1, True, False)
'        dtpEndDate.Enabled = True
'        lvwDept.Enabled = True
'        opt����(0).Enabled = True: opt����(1).Enabled = True: opt����(2).Enabled = True: opt����(3).Enabled = True
'        opt��.Enabled = True: opt��.Enabled = True
'        dtpBegin.Enabled = True: opt��Чʱ��(0).Enabled = True
'
'        '�Է����������:
'        '   ָ��ҽ��ʱ���������ó�,��̬�����ƽ������
'        If Trim(cboDoctor.Text) <> "" Then
'            opt����(2).Enabled = False: opt����(3).Enabled = False
'            If opt����(2).Value Or opt����(3).Value Then opt����(0).Value = True
'        Else
'            opt����(2).Enabled = True: opt����(3).Enabled = True
'        End If
'        If opt��.Value = True Then cbo��.Enabled = True
'    Case Else
'    End Select
'
'    '���ñ༭����ɫ
'    For Each ctl In Me.Controls
'        Select Case UCase(TypeName(ctl))
'        Case "TEXTBOX", UCase("ComboBox")
'            Call zlSetCtrolBackColor(ctl)
'        Case UCase("ListView")
'        Case UCase("DTPicker")
'        Case Else
'        End Select
'    Next
'End Sub
'
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdHelp_Click()
'    ShowHelp App.ProductName, Me.hWnd, Me.Name
'End Sub
'Private Function CheckPlanValied() As Boolean
'    '------------------------------------------------------------------------------------------------------------------------
'    '���ܣ����ƻ��ĺϷ���
'    '���أ��ƻ����źϷ�,����True,���򷵻�False
'    '���ƣ����˺�
'    '���ڣ�2010-07-21 17:49:30
'    '˵����
'    '------------------------------------------------------------------------------------------------------------------------
'    Dim rsTemp As New ADODB.Recordset
'    Dim strSQL As String
'    If mEditType <> Ed_�����޸� And mEditType <> ed_�ƻ����� Then
'        CheckPlanValied = True: Exit Function
'    End If
'
'    If dtpBegin.Value > dtpEndDate.Value Then
'        ShowMsgbox "ע��:" & vbCrLf & "    ��Чʱ��С����ʧЧʱ��,����!"
'        If dtpEndDate.Enabled And dtpEndDate.Visible Then dtpEndDate.SetFocus
'        Exit Function
'    End If
'    If zlDatabase.Currentdate > dtpBegin.Value Then
'        ShowMsgbox "ע��:" & vbCrLf & "    ��Чʱ��С���˵�ǰϵͳʱ��,����!"
'        If dtpBegin.Enabled And dtpBegin.Visible Then dtpBegin.SetFocus
'        Exit Function
'    End If
'    Set rsTemp = Nothing
'     CheckPlanValied = True: Exit Function
'End Function
'
'Private Function isValied() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:�����������ݵĺϷ���
'    '����:���ݺϷ�,����true,���򷵻�False
'    '����:���˺�
'    '����:2009-09-14 16:31:50
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long, intCount As Integer
'    Dim strTmp As String
'
'    Err = 0: On Error GoTo ErrHand:
'    If Trim(txt�ű�) = "" Then
'        MsgBox "�ű���Ϊ�գ�", vbInformation, gstrSysName
'        txt�ű�.SetFocus: Exit Function
'    End If
'    If cbo����.ListIndex = -1 Then
'        MsgBox "δ���úű�����Ӧ�Ŀ��ң�", vbInformation, gstrSysName
'        cbo����.SetFocus: Exit Function
'    End If
'    If cboItem.ListIndex = -1 Then
'        MsgBox "δ���úű�����Ӧ�ĹҺ���Ŀ��", vbInformation, gstrSysName
'        cboItem.SetFocus: Exit Function
'    End If
'
'    If opt��.Value Then
'        If cbo��.ListIndex = -1 Then
'            MsgBox "�úű�ÿ���Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
'            If txt�޺�.Enabled Then txt�޺�.SetFocus
'            Exit Function
'        End If
'        If chk��ſ���.Value = 1 Then
'            If Val(txt�޺�.Text) = 0 And Val(txt��Լ.Text) = 0 Then
'                MsgBox "ʹ����ſ���ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
'                If txt�޺�.Enabled Then txt�޺�.SetFocus
'                Exit Function
'            End If
'        End If
'        '�޺���Լ����
'        If Trim(txt�޺�.Text) <> "" Then
'            If Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
'                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
'               If txt��Լ.Enabled Then txt��Լ.SetFocus
'                Exit Function
'            End If
'        ElseIf Trim(txt��Լ.Text) <> "" Then
'            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
'            If txt�޺�.Enabled Then txt�޺�.SetFocus
'            Exit Function
'        End If
'    Else
'     With vsPlan
'            strTmp = ""
'            For i = 1 To .Cols - 1
'                If Trim(.TextMatrix(1, i)) <> "" Then
'                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))
'                    If chk��ſ���.Value = 1 Then
'                          If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
'                              MsgBox "ʹ����ſ���ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
'                              .Row = 2: .Col = i
'                              .SetFocus: Exit Function
'                          End If
'                      End If
'                        '�޺���Լ����
'                        If Val(.TextMatrix(2, i)) <> 0 Then
'                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
'                                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
'                                .Row = 2: .Col = i
'                                .SetFocus: Exit Function
'                            End If
'                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
'                            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
'                            .Row = 2: .Col = i
'                            .SetFocus: Exit Function
'                        End If
'                End If
'            Next
'            If strTmp = "" Then
'                MsgBox "�úű�ÿ�ܵ�Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
'                vsPlan.SetFocus: Exit Function
'            End If
'        End With
'    End If
'    '�����ж�
'    If opt����(1).Value Or opt����(2).Value Or opt����(3).Value Then
'        intCount = 0
'        For i = 1 To lvwDept.ListItems.Count
'            If lvwDept.ListItems(i).Checked Then intCount = intCount + 1
'        Next
'        If opt����(1).Value Then
'            If intCount = 0 Then
'                MsgBox "ָ������ʱ����ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
'                lvwDept.SetFocus: Exit Function
'            ElseIf intCount > 1 Then
'                MsgBox "ָ������ʱֻ��ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
'                lvwDept.SetFocus: Exit Function
'            End If
'        ElseIf opt����(2).Value Or opt����(3).Value Then
'            If intCount < 2 Then
'                MsgBox "��̬�����ƽ������ʱ����Ҫѡ��������Ӧ���������ң�", vbInformation, gstrSysName
'                lvwDept.SetFocus: Exit Function
'            End If
'        End If
'    End If
'
'    '��Ŀ�۸��ж�
'    If ReadRegistPrice(cboItem.ItemData(cboItem.ListIndex), False, False) = 0 Then
'        MsgBox "��Ŀ""" & cboItem.Text & """δ������Ч�۸�,���ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
'        cboItem.SetFocus: Exit Function
'    End If
'    If opt��Чʱ��(0).Value = 0 Then
'        If Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
'            ShowMsgbox "��Чʱ�䲻��С�ڵ�ǰϵͳʱ��,����!"
'            Exit Function
'        End If
'    End If
'    '�����صļƻ�
'    If CheckPlanValied = False Then Exit Function
'    Dim blnMulitPlan As Boolean
'    isValied = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    If 1 = 2 Then
'        Resume
'    End If
'End Function
'
'Private Function SavePlan() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:����ƻ�����
'    '����:����ɹ�������true,���򷵻�False
'    '����:���˺�
'    '����:2009-09-14 16:41:22
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String, strʱ��� As String, str���� As String, i As Long, int���� As Integer
'    Dim lng�ƻ�ID As Long, str�޺� As String
'    Dim strҽ������         As String
'    Dim strҽ��ID           As String
'    Dim blnChange           As Boolean
'    Dim BytType             As Byte
'    Dim vMsgResult          As VbMsgBoxResult
'    Dim strMsg              As String
'    Dim colPro              As Collection
'    'bytType 0-����ʱ ��ʱ�β����д��� �޸�ʱ ��ʱ��ֻɾ���Ѿ�ȥ�����Ű���Ϣ
'    '        1-����ʱ ��ȡԭ���ŵ�ʱ����Ϣ  �޸�ʱ �Լƻ���ʱ�ν���ɾ��
'
'    Err = 0: On Error GoTo ErrHand:
'
'    strʱ��� = "": str�޺� = ""
'    If opt��.Value Then
'        For i = 1 To 7
'            strʱ��� = strʱ��� & ",'" & Trim(cbo��.Text) & "'"
'            str�޺� = str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
'            str�޺� = str�޺� & "," & Val(txt�޺�.Text) & "," & Val(txt��Լ.Text)
'        Next
'    Else
'        With vsPlan
'            For i = 1 To .Cols - 1
'                strʱ��� = strʱ��� & ",'" & Trim(.TextMatrix(1, i)) & "'"
'                If Trim(.TextMatrix(1, i)) <> "" Then
'                    str�޺� = str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
'                    str�޺� = str�޺� & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Val(Trim(vsPlan.TextMatrix(3, i)))
'                End If
'            Next
'        End With
'    End If
'    If str�޺� <> "" Then str�޺� = Mid(str�޺�, 2)
'
'    If mPlanInfo.blnʱ��� Then
'        '�ж����Ѿ��ı� �ƻ���Ϣ
'      blnChange = (mPlanInfo.str�Ű� <> strʱ���) Or (mPlanInfo.str�޺� <> str�޺�) Or (IIf(mPlanInfo.bln���, 1, 0) <> chk��ſ���.Value)
'    End If
'    With lvwDept
'        'ȡ�Һ�����
'        For i = 1 To .ListItems.Count
'            If .ListItems(i).Checked Then
'                str���� = str���� & ";" & .ListItems(i).Text
'            End If
'        Next
'        str���� = Mid(str����, 2)
'    End With
'
'    'ȡ���﷽ʽ
'    int���� = 0
'    For i = 0 To opt����.UBound
'        If opt����(i).Value Then int���� = i: Exit For
'    Next
'
'    '�ڼƻ����߰���������ʱ��ʱ ��ʱ�δ����Ĵ�������
''    If mPlanInfo.blnʱ��� And mEditType = ed_�ƻ����� And blnChange = False Then
''        '���ԭ�ƻ����߰���ʱ ������ʱ�� ��ʾ����ԭ���д���
''        strMsg = "������������ʱ��,�Ƿ���ȡ���ŵ�ʱ����Ϊ�ƻ���ʱ����Ϣ? " & vbCrLf
''        strMsg = strMsg & "[��(Y)]��ȡ���ŵ�ʱ����Ϣ��Ϊ�ƻ���ʱ��" & vbCrLf
''        strMsg = strMsg & "[��(N)]����ȡ���ŵ�ʱ��,��������ʱ��" & vbCrLf
''        vMsgResult = MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
''        BytType = IIf(vMsgResult = vbYes, 1, 0)
''    End If
'    If mEditType = Ed_�����޸� Then
'      BytType = IIf(IIf(mPlanInfo.bln���, 1, 0) <> chk��ſ���.Value, 1, 0)
'    End If
'    'ȡʱ�䷶Χ
'    If mEditType = ed_�ƻ����� Then
'        lng�ƻ�ID = zlDatabase.GetNextId("�ҺŰ��żƻ�")
'    Else
'        lng�ƻ�ID = Val(mstr�ƻ�ID)
'    End If
'     If cboDoctor.ListIndex = -1 Then
'        strҽ������ = ""
'        strҽ��ID = "0"
'     Else
'        strҽ������ = cboDoctor.Text
'        strҽ��ID = Val(cboDoctor.ItemData(cboDoctor.ListIndex))
'     End If
'    'Zl_�ҺŰ��żƻ�_Insert
'    strSQL = "Zl_�ҺŰ��żƻ�_Insert("
'    '  Id_In       In �ҺŰ��żƻ�.ID%Type,
'    strSQL = strSQL & "" & lng�ƻ�ID & ","
'    '  ����id_In   In �ҺŰ��żƻ�.����id%Type,
'    strSQL = strSQL & "" & mlng����ID & ","
'    '  ����_In     In �ҺŰ��żƻ�.����%Type,
'    strSQL = strSQL & "'" & txt�ű�.Text & "',"
'    '  ��Чʱ��_In In �ҺŰ��żƻ�.��Чʱ��%Type,
'    If opt��Чʱ��(0).Value = 1 Then
'        strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
'    Else
'        strSQL = strSQL & "to_date('" & dtpBegin.Value & "','yyyy-mm-dd hh24:mi:ss'),"
'    End If
'    '  ʧЧʱ��_In In �ҺŰ��żƻ�.ʧЧʱ��%Type
'    strSQL = strSQL & "to_date('" & dtpEndDate.Value & "','yyyy-mm-dd hh24:mi:ss') "
'    '  ����_In     In �ҺŰ��żƻ�.����%Type,
'    '  ��һ_In     In �ҺŰ��żƻ�.��һ%Type,
'    '  �ܶ�_In     In �ҺŰ��żƻ�.�ܶ�%Type,
'    '  ����_In     In �ҺŰ��żƻ�.����%Type,
'    '  ����_In     In �ҺŰ��żƻ�.����%Type,
'    '  ����_In     In �ҺŰ��żƻ�.����%Type,
'    '  ����_In     In �ҺŰ��żƻ�.����%Type,
'    strSQL = strSQL & strʱ��� & ","
'    '   �޺ſ���_In In Varchar2,
'    strSQL = strSQL & "'" & str�޺� & "',"
'    '  ���﷽ʽ_In In �ҺŰ��żƻ�.���﷽ʽ%Type,
'    strSQL = strSQL & "" & int���� & ","
'    '  ��ſ���_In In �ҺŰ��żƻ�.��ſ���%Type,
'    strSQL = strSQL & "" & IIf(chk��ſ���.Value = 1, 1, 0) & ","
'    '  ��ĿID_In   In �ҺŰ��żƻ�.��ĿID%Type,
'    strSQL = strSQL & Me.cboItem.ItemData(cboItem.ListIndex) & ","
'    'ҽ������_In In �ҺŰ��żƻ�.ҽ������%Type,
'    strSQL = strSQL & IIf(strҽ������ = "", "NULL,", "'" & strҽ������ & "',")
'    'ҽ��id_In   In �ҺŰ��żƻ�.ҽ��id%Type,
'    strSQL = strSQL & strҽ��ID & ","
'    '  ����_In     Varchar2,
'    strSQL = strSQL & "'" & str���� & "',"
'    '  ����_In Number:=1,��������
'    strSQL = strSQL & "" & IIf(mEditType = ed_�ƻ�����, 1, 0) & "," & BytType & ","
'    '��������_In Number:=0,
'    strSQL = strSQL & "" & IIf(opt��Чʱ��(0).Value = True, 1, 0) & ")"
'    '�������_In Number:=0
'
'    Set colPro = New Collection
'    zlAddArray colPro, strSQL
'    If Not mfrmTime.IsInit Then
'         Call LoadTimePlan
'    End If
'    If mfrmTime.zlSaveData(lng�ƻ�ID, colPro) = False Then Exit Function
'    SavePlan = True
'    zlExecuteProcedureArrAy colPro, Me.Caption
'
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    SaveErrLog
'End Function
'Private Function SaveVerify() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��˹ҺŰ��żƻ�
'    '����:��˳ɹ�,����true, ���򷵻�False
'    '����:���˺�
'    '����:2009-09-14 17:11:24
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String
'    Err = 0: On Error GoTo ErrHand:
'    'Zl_�ҺŰ��żƻ�_Verify(Id_In In �ҺŰ��żƻ�.ID%Type)
'    strSQL = "Zl_�ҺŰ��żƻ�_Verify(" & Val(mstr�ƻ�ID) & ")"
'    zlDatabase.ExecuteProcedure strSQL, Me.Caption
'    SaveVerify = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    SaveErrLog
'End Function
'Private Function SaveCancel() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:ȡ����˹ҺŰ��żƻ�
'    '����:ȡ����˳ɹ�,����true, ���򷵻�False
'    '����:���˺�
'    '����:2009-09-14 17:11:24
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String
'    Err = 0: On Error GoTo ErrHand:
'    'Zl_�ҺŰ��żƻ�_Cancel(Id_In In �ҺŰ��żƻ�.ID%Type) Is
'    strSQL = "Zl_�ҺŰ��żƻ�_Cancel(" & Val(mstr�ƻ�ID) & ")"
'    zlDatabase.ExecuteProcedure strSQL, Me.Caption
'    SaveCancel = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'     SaveErrLog
'End Function
'Private Function SaveDelete() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:ȡ����˹ҺŰ��żƻ�
'    '����:ȡ����˳ɹ�,����true, ���򷵻�False
'    '����:���˺�
'    '����:2009-09-14 17:11:24
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String
'    Err = 0: On Error GoTo ErrHand:
'    'Zl_�ҺŰ��żƻ�_Delete(Id_In In �ҺŰ��żƻ�.ID%Type) Is
'    strSQL = "Zl_�ҺŰ��żƻ�_Delete(" & Val(mstr�ƻ�ID) & ")"
'    zlDatabase.ExecuteProcedure strSQL, Me.Caption
'    SaveDelete = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'     SaveErrLog
'End Function
'
'Private Sub cmdOK_Click()
'
'
'    If mEditType = ed_���Ų��� Then Unload Me: Exit Sub
'    If mEditType = Ed_����ɾ�� Then
'        If SaveDelete = False Then Exit Sub
'        mblnSucces = True
'        Unload Me: Exit Sub
'    End If
'
'    If mEditType = Ed_������� Then
'        If SaveVerify = False Then Exit Sub
'        mblnSucces = True
'        Unload Me: Exit Sub
'    End If
'
'    If mEditType = Ed_����ȡ�� Then
'        If SaveCancel = False Then Exit Sub
'        mblnSucces = True
'        Unload Me: Exit Sub
'    End If
'    If isValied = False Then Exit Sub
'
'    If SavePlan = False Then Exit Sub
'    mblnSucces = True
'    Unload Me
'
'End Sub
'
'Private Sub Form_Resize()
'    Err = 0: On Error Resume Next
'    With cmdOK
'        .Left = ScaleWidth - .Width - 100
'        cmdCancel.Left = .Left
'        cmdHelp.Left = .Left
'    End With
'
'    With tbPage
'        .Top = 50
'        .Height = ScaleHeight - 100
'        .Left = 50
'        .Width = cmdOK.Left - .Left - 100
'    End With
'
'End Sub
'
'Private Sub lvwDept_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    Dim i As Integer
'    If opt����(1).Value Then
'        For i = 1 To lvwDept.ListItems.Count
'            If lvwDept.ListItems(i).Key <> Item.Key Then
'                lvwDept.ListItems(i).Checked = False
'            End If
'        Next
'    End If
'    Set lvwDept.SelectedItem = Item
'End Sub
'Private Sub opt����_Click(Index As Integer)
'    Dim i As Integer, strKey As String
'    If opt����(1).Value Then
'        For i = 1 To lvwDept.ListItems.Count
'            If lvwDept.ListItems(i).Checked Then
'                If strKey = "" Then
'                    strKey = lvwDept.ListItems(i).Key
'                Else
'                    lvwDept.ListItems(i).Checked = False
'                End If
'            End If
'        Next
'        If strKey <> "" Then
'            Set lvwDept.SelectedItem = lvwDept.ListItems(strKey)
'            lvwDept.SelectedItem.EnsureVisible
'        End If
'    End If
'End Sub
'
'Private Sub opt��Чʱ��_Click(Index As Integer)
'     dtpBegin.Enabled = opt��Чʱ��(0).Value = 0
'End Sub
'
'Private Sub opt��_Click()
'    Dim i As Integer
'    Dim strPlan As String
'    Dim ctl As Control
'
'    With vsPlan
'        For i = 1 To .Cols - 1
'            If Trim(.TextMatrix(1, i)) <> "" Then
'                If strPlan = "" Then
'                    strPlan = .TextMatrix(1, i)
'                Else
'                    If .TextMatrix(1, i) <> strPlan Then
'                        strPlan = "": Exit For
'                    End If
'                End If
'            End If
'        Next
'        For i = 1 To .Cols - 1
'            .TextMatrix(1, i) = ""
'            .TextMatrix(2, i) = ""
'            .TextMatrix(3, i) = ""
'        Next
'        .Enabled = False: .TabStop = False
'    End With
'    opt��.Value = -True: txt�޺�.Enabled = True: txt��Լ.Enabled = True
'    cbo��.Enabled = True
'    opt��.Value = False
'    cbo��.ListIndex = GetCboIndex(cbo��, strPlan)
'    cbo��.SetFocus
'
'    '���ñ༭����ɫ
'    For Each ctl In Me.Controls
'        Select Case UCase(TypeName(ctl))
'        Case "TEXTBOX", UCase("ComboBox")
'            Call zlSetCtrolBackColor(ctl)
'        Case UCase("ListView")
'        Case UCase("DTPicker")
'        Case Else
'        End Select
'    Next
'End Sub
'
'Private Sub opt��_Click()
'    Dim i As Integer
'    Dim ctl As Control
'
'    If Trim(cbo��.Text) <> "" Then
'        With vsPlan
'            For i = 0 To .Cols - 1
'                .TextMatrix(1, i) = cbo��.Text
'                .TextMatrix(2, i) = txt�޺�.Text
'                .TextMatrix(3, i) = txt��Լ.Text
'            Next
'            .Enabled = True: .TabStop = True
'            .Col = 1: .SetFocus
'        End With
'    End If
'    opt��.Value = False: txt�޺�.Enabled = False: txt��Լ.Enabled = False
'    cbo��.Enabled = False: cbo��.ListIndex = -1
'    opt��.Value = True: vsPlan.Enabled = True
'
'    '���ñ༭����ɫ
'    For Each ctl In Me.Controls
'        Select Case UCase(TypeName(ctl))
'        Case "TEXTBOX", UCase("ComboBox")
'            Call zlSetCtrolBackColor(ctl)
'        Case UCase("ListView")
'        Case UCase("DTPicker")
'        Case Else
'        End Select
'    Next
'End Sub
'
'
'
'Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'
'
'      If mblnChangeByCode Then Exit Sub
'    PageChange Item
'End Sub
'
'Private Sub PageChange(ByVal Item As XtremeSuiteControls.ITabControlItem)
'
'    If mblnChangeByCode Then Exit Sub
'
'    If Item.Index = mPageIndex.EM_ʱ�� Then
'       mblnChangeByCode = True
'       tbPage.Item(mPageIndex.EM_�ƻ�).Selected = True
'        If isValied() = False Then
'            mblnChangeByCode = False
'            Exit Sub
'        End If
'        tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
'        mblnChangeByCode = False
'        Call LoadTimePlan
'    Else
'        If mfrmTime.mblnChange = False Then Exit Sub
'        If mfrmTime.zlPageSelectedChanged() = False Then
'             mblnChangeByCode = True
'            tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
'             mblnChangeByCode = False
'        End If
'    End If
'End Sub
'
'
'
'Private Sub LoadTimePlan()
'    Dim i As Long
'    Dim lng�޺��� As Long
'    Dim lng��Լ�� As Long
'    Dim strTemp As String
'    Dim str���� As String
'    Dim str�Ű� As String
'
'    If Not mrsRegNewData Is Nothing Then Set mrsRegNewData = Nothing
'
'    If mrsRegNewData Is Nothing Then
'        Set mrsRegNewData = New ADODB.Recordset
'        mrsRegNewData.Fields.Append "ID", adBigInt, 18
'        mrsRegNewData.Fields.Append "������Ŀ", adVarChar, 20
'        mrsRegNewData.Fields.Append "�Ű�", adVarChar, 20
'        mrsRegNewData.Fields.Append "�޺���", adBigInt, 10
'        mrsRegNewData.Fields.Append "��Լ��", adBigInt, 18
'        mrsRegNewData.Fields.Append "��ſ���", adBigInt, 18
'        mrsRegNewData.CursorLocation = adUseClient
'        mrsRegNewData.LockType = adLockOptimistic
'        mrsRegNewData.CursorType = adOpenStatic
'        mrsRegNewData.Open
'     End If
'
'     If opt��.Value = True Then
'          lng�޺��� = Val(txt�޺�.Text)
'          lng��Լ�� = Val(txt��Լ.Text)
'          str�Ű� = Me.cbo��.Text
'          For i = 0 To 6
'            strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
'            '��һ,�޺���,��Լ��|�ܶ�,�޺���,��Լ��|....
'            str���� = str���� & "|" & strTemp & "," & lng�޺��� & "," & lng��Լ��
'             With mrsRegNewData
'                .AddNew
'                !ID = Val(mstr�ƻ�ID)
'                !������Ŀ = strTemp
'                !�Ű� = str�Ű�
'                !�޺��� = lng�޺���
'                !��Լ�� = lng��Լ��
'                !��ſ��� = Me.chk��ſ���.Value
'                .Update
'            End With
'          Next
'
'        Else
'
'           With vsPlan
'            For i = 1 To .Cols - 1
'                If Trim(.TextMatrix(1, i)) <> "" Then
'                    strTemp = Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
'                    lng�޺��� = Val(Trim(vsPlan.TextMatrix(2, i)))
'                    lng��Լ�� = Val(Trim(vsPlan.TextMatrix(3, i)))
'                    str�Ű� = Trim(vsPlan.TextMatrix(1, i))
'                    str���� = str���� & "|" & strTemp & "," & lng�޺��� & "," & lng��Լ��
'                    With mrsRegNewData
'                        .AddNew
'                        !ID = Val(mstr�ƻ�ID)
'                        !������Ŀ = strTemp
'                        !�Ű� = str�Ű�
'                        !�޺��� = lng�޺���
'                        !��Լ�� = lng��Լ��
'                        !��ſ��� = Me.chk��ſ���.Value
'                        .Update
'                    End With
'                End If
'            Next
'        End With
'     End If
'     If str���� <> "" Then str���� = Mid(str����, 2)
''Public Enum mRegEditType
''Ed_�ƻ����� = 0
''Ed_�����޸� = 1
''Ed_����ɾ�� = 2
''Ed_������� = 3
''Ed_����ȡ�� = 4
''Ed_���Ų��� = 5
''End Enum
'
'     mfrmTime.zlShowPagePlan str����, mrsRegNewData, mrsRegHistory, chk��ſ���.Value = 1, Switch(mEditType = ed_�ƻ�����, EM_�ƻ�_����, mEditType = Ed_�����޸�, EM_�ƻ�_�޸�, True, EM_�ƻ�_����), mlng����ID, Val(mstr�ƻ�ID)
'End Sub
'
''Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
''    Dim i As Long
''    Dim lng�޺��� As Long
''    Dim lng��Լ�� As Long
''    Dim strTemp As String
''    Dim str���� As String
''    If Item.Index <> mPageIndex.EM_ʱ�� Then Exit Sub
''    If Not mrsRegNewData Is Nothing Then Set mrsRegNewData = Nothing
''    If mrsRegNewData Is Nothing Then
''        With mrsRegNewData
''        Set mrsRegNewData = New ADODB.Recordset
''        mrsRegNewData.Fields.Append "ID", adBigInt, 18
''        mrsRegNewData.Fields.Append "������Ŀ", adVarChar, 20
''        mrsRegNewData.Fields.Append "�޺���", adBigInt, 10
''        mrsRegNewData.Fields.Append "��Լ��", adBigInt, 18
''        mrsRegNewData.Fields.Append "��ſ���", adBigInt, 18
''        mrsRegNewData.CursorLocation = adUseClient
''        mrsRegNewData.LockType = adLockOptimistic
''        mrsRegNewData.CursorType = adOpenStatic
''        mrsRegNewData.Open
''        If opt��.Value = True Then
''          lng�޺��� = Val(txt�޺�.Text)
''          lng��Լ�� = Val(txt��Լ.Text)
''          For i = 0 To 6
''            strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
''            .AddNew
''            !ID = Val(mstr�ƻ�ID)
''            !������Ŀ = strTemp
''            !�޺��� = lng�޺���
''            !��Լ�� = lng��Լ��
''            !��ſ��� = Me.chk��ſ���.Value
''            .Update
''          Next
''
''        Else
''
''           With vsPlan
''            For i = 1 To .Cols - 1
''                If Trim(.TextMatrix(1, i)) <> "" Then
''                    strTemp = Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
''                    lng�޺��� = Val(Trim(vsPlan.TextMatrix(2, i)))
''                    lng��Լ�� = Val(Trim(vsPlan.TextMatrix(3, i)))
''                    With mrsRegNewData
''                        .AddNew
''                        !ID = Val(mstr�ƻ�ID)
''                        !������Ŀ = strTemp
''                        !�޺��� = lng�޺���
''                        !��Լ�� = lng��Լ��
''                        !��ſ��� = Me.chk��ſ���.Value
''                        .Update
''                    End With
''                End If
''            Next
''        End With
''
''        End If
''    End With
''
''    End If
''    If mfrmTime Is Nothing Then
''        Set mfrmTime = New frmResistPlanTimeSet
''    End If
''End Sub
'
'Private Sub txt�ű�_GotFocus()
'    zlControl.TxtSelAll txt�ű�
'End Sub
'Private Sub txt�ű�_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'End Sub
'
'Private Sub txt�޺�_GotFocus()
'    zlControl.TxtSelAll txt�޺�
'End Sub
'Private Sub txt�޺�_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'End Sub
'
'Private Sub txt�޺�_Validate(Cancel As Boolean)
'    If Trim(txt�޺�.Text) = "" And Trim(txt��Լ.Text) <> "" Then
'        MsgBox "��Լ�����޺�!", vbInformation, gstrSysName
'        Cancel = True: Exit Sub
'    End If
'    If Trim(txt�޺�.Text) <> "" And Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
'        MsgBox "��Լ������С���޺���!", vbInformation, gstrSysName
'        Cancel = True: Exit Sub
'    End If
'End Sub
'Private Sub txt��Լ_GotFocus()
'    zlControl.TxtSelAll txt��Լ
'End Sub
'
'Private Sub txt��Լ_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'    If Val(txt�޺�.Text) = 0 Then KeyAscii = 0
'End Sub
'
'Private Sub txt��Լ_Validate(Cancel As Boolean)
'    If Val(txt�޺�.Text) < Val(txt��Լ.Text) And _
'        Trim(txt�޺�.Text) <> "" And Trim(txt��Լ.Text) <> "" Then
'        MsgBox "��Լ������С���޺���!", vbInformation, gstrSysName
'        Cancel = True: Exit Sub
'    End If
'End Sub
'
'Private Function zlCheckRegistPlanIsValied(ByRef blnMulitNumPlan As Boolean) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��鵱ǰ������ĺ����Ƿ�Ϸ�
'    '����:blnMulitNumPlan-�����Ƿ��ж����ͬ(ͬһ��Ŀ,ͬһ����,ͬһ��,��ͬ��)�İ���
'    '����:�Ϸ�����,�򷵻�true,���򷵻�False
'    '����:���˺�
'    '����:2010-12-29 10:26:45
'    '������ͬһ��Ŀ,ͬһ����,ͬһ��,��ͬ�ţ�:
'    '     1.ͬ���ڲ����н���İ���
'    '����Ŀ:35057
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String, rsTemp As ADODB.Recordset, strҽ�� As String
'    Dim lng��Ŀid As Long, lng����ID As Long, lngҽ��ID As Long
'    Dim str�ű� As String, strTemp As String, strTemp1 As String
'    Dim i As Long, bytCheckType As Byte '0-���ƻ��Ƿ�Ϸ�;1-��鰲��������ִ����Ŀ�Ƿ�Ϸ�.
'    Dim strTittle As String
'
'    On Error GoTo errHandle
'    lng����ID = cbo����.ItemData(cbo����.ListIndex)
'    lng��Ŀid = cboItem.ItemData(cboItem.ListIndex)
'    lngҽ��ID = 0: strҽ�� = Trim(cboDoctor.Text)
'    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
'
'    '���ƻ����Ƿ�����ظ�
'    bytCheckType = 0
'goReCheck:
'    If bytCheckType <> 0 Then
'
'        strSQL = "" & _
'        "   Select Distinct A.����, A.���� D0, A.��һ D1, A.�ܶ� D2, A.���� D3, A.���� D4, A.���� D5, A.���� D6, " & _
'        "                 Nvl(To_Char(a.��ʼʱ��, 'YYYY-MM-DD HH24:MI:SS'), '1901-01-01') ��Чʱ��, " & _
'        "                 Nvl(To_Char(a.��ֹʱ��, 'YYYY-MM-DD HH24:MI:SS'), '3000-01-01 00:00:00') ʧЧʱ�� " & _
'        "   From �ҺŰ��� A,�ҺŰ��� B " & _
'        "   Where A.����id = b.����id And A.ҽ������ = b.ҽ������ And Nvl(A.ҽ��id, 0) = nvl(b.ҽ��id,0) " & _
'        "               And a.ID + 0 <> [1]   And B.ID = [1]  " & _
'        "   Order By ����"
'            strTittle = "����"
'    Else
'        strSQL = "" & _
'            "   Select  distinct A.����,A.���� D0,A.��һ D1,A.�ܶ� D2,A.���� D3,A.���� D4,A.���� D5,A.���� D6," & _
'            "           To_Char(A.��Чʱ��,'YYYY-MM-DD HH24:MI:SS') ��Чʱ��,To_Char(A.ʧЧʱ��,'YYYY-MM-DD HH24:MI:SS') ʧЧʱ��" & _
'            "   From �ҺŰ��żƻ� A, �ҺŰ��� B,�ҺŰ��� C " & _
'            "   Where A.����ID=B.ID and B.����ID=C.����ID and B.ҽ������=C.ҽ������ and nvl(B.ҽ��ID,0)=nvl(C.ҽ��ID,0) " & _
'            "           And B.ID+0<>[1] and C.ID=[1]  " & _
'            "   Order by ����"
'            strTittle = "�ƻ�����"
'    End If
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
'    blnMulitNumPlan = Not rsTemp.EOF
'    If blnMulitNumPlan = False And bytCheckType = 0 Then
'        bytCheckType = bytCheckType + 1
'        GoTo goReCheck:
'    End If
'    If blnMulitNumPlan = False Then zlCheckRegistPlanIsValied = True: Exit Function
'    str�ű� = ""
'    Do While Not rsTemp.EOF
'        str�ű� = str�ű� & "," & Nvl(rsTemp!����)
'        If (Nvl(rsTemp!��Чʱ��) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!��Чʱ��) < Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS")) Or _
'           (Nvl(rsTemp!ʧЧʱ��) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!ʧЧʱ��) < Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS")) Or _
'           (Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!��Чʱ��) And Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!ʧЧʱ��)) Or _
'           (Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!��Чʱ��) And Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!ʧЧʱ��)) Then
'           'ʱ���ڲ��ܽ���
'            If opt��.Value Then
'                If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D0)
'                If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & "  ��һ:" & Nvl(rsTemp!D1)
'                If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & "  �ܶ�:" & Nvl(rsTemp!D2)
'                If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D3)
'                If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D4)
'                If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D5)
'                If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D6)
'                If strTemp <> "" Then
'                    strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ����������" & strTittle & ":" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & vbCrLf & "  ��Чʱ��:" & IIf(Nvl(rsTemp!��Чʱ��) = "1901-01-01", "����", Nvl(rsTemp!��Чʱ��) & "-" & Nvl(rsTemp!ʧЧʱ��)) & vbCrLf
'                    Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺżƻ����� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˼ƻ�����.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
'                    zlCheckRegistPlanIsValied = False: Exit Function
'                End If
'            Else
'                With vsPlan
'                    For i = 0 To 6
'                        strTemp1 = "  ��" & Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
'                        If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
'                            '����,�϶��ظ���
'                            strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
'                        End If
'                    Next
'                End With
'                If strTemp <> "" Then
'                    strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ����������" & strTittle & ":" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & "  ��Чʱ��:" & IIf(Nvl(rsTemp!��Чʱ��) = "1901-01-01", "����", Nvl(rsTemp!��Чʱ��) & "-" & Nvl(rsTemp!ʧЧʱ��)) & vbCrLf
'                    Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˼ƻ�����.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
'                    zlCheckRegistPlanIsValied = False: Exit Function
'                End If
'            End If
'        End If
'        rsTemp.MoveNext
'    Loop
'    If bytCheckType = 0 Then
'        bytCheckType = bytCheckType + 1
'        GoTo goReCheck:
'    End If
'    zlCheckRegistPlanIsValied = True
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'     SaveErrLog
'End Function
'
'Private Sub vsPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsPlan
'        If mEditType <> ed_�ƻ����� And mEditType <> Ed_�����޸� Then Cancel = True: Exit Sub
'        If Not opt��.Value = True Then Cancel = True: Exit Sub
'    End With
'End Sub
'Private Sub vsPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'      '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:������صĸ�ʽ
'    '����:���˺�
'    '����:2011-11-11 11:33:11
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    With vsPlan
'        If Row = 1 Then
'              If Trim(.EditText) = "" Then
'               .TextMatrix(2, Col) = ""
'               .TextMatrix(3, Col) = ""
'            End If
'            Exit Sub
'        End If
'        .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), "###;;;")
'    End With
'    Exit Sub
'End Sub
'Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'   Call zl_VsGridRowChange(vsPlan, OldRow, NewRow, OldCol, NewCol)
'    vsPlan.ColComboList(NewCol) = ""
'    If OldRow = 1 And Trim(vsPlan.TextMatrix(1, OldCol)) = "" Then
'        vsPlan.TextMatrix(2, OldCol) = ""
'        vsPlan.TextMatrix(3, OldCol) = ""
'    End If
'    If OldRow = 2 And Trim(vsPlan.TextMatrix(3, OldCol)) = "" Then
'        vsPlan.TextMatrix(3, OldCol) = vsPlan.TextMatrix(2, OldCol)
'    End If
'    If NewRow <> 1 Then Exit Sub
'    vsPlan.ColComboList(NewCol) = vsPlan.Tag
'End Sub
'Private Sub vsPlan_GotFocus()
'    Call zl_VsGridGotFocus(vsPlan)
'End Sub
'Private Sub vsPlan_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
'    With vsPlan
'        If KeyCode = vbKeyDelete Then
'            .TextMatrix(.Row, .Col) = ""
'        End If
'    End With
'    If KeyCode <> vbKeyReturn Then Exit Sub
'
'    With vsPlan
'        If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
'        If .Row < 3 Then
'            .Row = .Row + 1
'        Else
'            .Row = 1
'            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
'         End If
'    End With
'End Sub
'
'Private Sub vsPlan_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'    '�༭����
'    Dim intCol As Integer, strKey As String, lngRow As Long
'
'    If KeyCode <> vbKeyReturn Then Exit Sub
'    With vsPlan
'            If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
'        If .Row < 3 Then
'            .Row = .Row + 1
'        Else
'            .Row = 1
'            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
'         End If
'    End With
'End Sub
'Private Sub vsPlan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then KeyAscii = 0
'End Sub
'Private Sub vsPlan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    With vsPlan
'        If Row <= 1 Then Exit Sub
'        VsFlxGridCheckKeyPress vsPlan, Row, Col, KeyAscii, m����ʽ
'    End With
'End Sub
'Private Sub vsPlan_LostFocus()
'    zlCommFun.OpenIme False
'    Call zl_VsGridLOSTFOCUS(vsPlan)
'End Sub
'
'Private Sub vsPlan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    Dim strKey As String, intCol As Integer, strTemp As String
'    Dim str������Ŀ As String
'    Dim lng��Լ��  As Long
'    '������֤
'    With vsPlan
'        str������Ŀ = Switch(Col = 1, "����", Col = 2, "��һ", Col = 3, "�ܶ�", Col = 4, "����", Col = 5, "����", Col = 6, "����", True, "����")
'        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
'        If .Row <= 1 Then Exit Sub
'        If zlDblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
'            Cancel = True: Exit Sub
'        End If
'        strKey = Format(Abs(Val(strKey)), "####;;;")
'        If Row = 2 Then
'            If mrsRegHistory.RecordCount <> 0 Then
'                mrsRegHistory.Filter = "������Ŀ='" & str������Ŀ & "'"
'                If mrsRegHistory.RecordCount <> 0 Then
'                     lng��Լ�� = Val(Nvl(mrsRegHistory!ͳ��))
'                     If lng��Լ�� > Val(strKey) Then
'                        Call MsgBox("�޺���С�����Ѿ�ԤԼ��ȥ������[" & lng��Լ�� & "],���ܼ���!", vbOKOnly, gstrSysName)
'                        mrsRegHistory.Filter = 0: Cancel = True: Exit Sub
'                      End If
'                End If
'                mrsRegHistory.Filter = 0
'            End If
'            If Val(strKey) < Val(.TextMatrix(3, Col)) Then
'                If MsgBox("�޺���С������Լ��,�Ƿ������Լ��?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True: Exit Sub
'                .TextMatrix(3, Col) = ""
'            End If
'        ElseIf Row = 3 Then
'            If mrsRegHistory.RecordCount <> 0 Then
'                mrsRegHistory.Filter = "������Ŀ='" & str������Ŀ & "'"
'                If mrsRegHistory.RecordCount <> 0 Then
'                     lng��Լ�� = Val(Nvl(mrsRegHistory!ͳ��))
'                     If lng��Լ�� > Val(strKey) Then
'                        Call MsgBox("��Լ��С�����Ѿ�ԤԼ��ȥ������[" & lng��Լ�� & "],���ܼ���!", vbOKOnly, gstrSysName)
'                        mrsRegHistory.Filter = 0: Cancel = True: Exit Sub
'                      End If
'                End If
'                mrsRegHistory.Filter = 0
'            End If
'
'            If Val(strKey) > Val(.TextMatrix(2, Col)) Then
'                Call MsgBox("�޺���С������Լ��,���ܼ���", vbOKOnly, gstrSysName)
'                Cancel = True: Exit Sub
'            End If
'        End If
'        .EditText = strKey
'    End With
'End Sub
'
'
'
'
'Private Sub cboDoctor_Validate(Cancel As Boolean)
'
'    'ָ��ҽ��ʱ����ָ���������
'    If Trim(cboDoctor.Text) <> "" Then
'        opt����(2).Enabled = False
'        opt����(3).Enabled = False
'        If opt����(2).Value Or opt����(3).Value Then opt����(0).Value = True
'    Else
'        opt����(2).Enabled = True
'        opt����(3).Enabled = True
'    End If
'End Sub
'
'Private Sub LoadDoctor()
'    Set mrsDoctor = GetDoctor(Val(cbo����.ItemData(cbo����.ListIndex)), "")
'    cboDoctor.Clear
'    Do While Not mrsDoctor.EOF
'        cboDoctor.AddItem mrsDoctor!����
'        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
'        mrsDoctor.MoveNext
'    Loop
'End Sub
'
'Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
'    Dim lngIdx As Long, lngҽ��ID As Long
'    If KeyAscii <> 13 Then Exit Sub
'    If cboDoctor.ListIndex <> -1 Then
'        zlCommFun.PressKey vbKeyTab: Exit Sub
'    End If
'    If mrsDoctor Is Nothing Then Exit Sub
'    If Trim(cboDoctor.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
'
'    If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, cboDoctor.Text, True, "") = False Then
'        KeyAscii = 0: Exit Sub
'    End If
'    Exit Sub
'End Sub
'
'Private Function Checkʱ��() As Boolean
'    '�����Ӽƻ�ʱ ��ȡԭ�еİ����Ƿ����ʱ��
'    '�޸ļƻ�ʱ ��ȡԭ�ƻ��Ƿ����ʱ��
'   Dim strSQL           As String
'   Dim rsTmp            As ADODB.Recordset
'   If mEditType <> Ed_�����޸� And mEditType <> ed_�ƻ����� Then Exit Function
'    On Error GoTo Hd
'    If mEditType = ed_�ƻ����� Then
'        strSQL = " Select 1 As Hdata From �ҺŰ���ʱ�� Where ����id =[1] And Rownum=1"
'    Else
'        strSQL = "Select 1  as haveData From �Һżƻ�ʱ�� Where �ƻ�ID=[2] and Rownum=1"
'    End If
'     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mstr�ƻ�ID))
'     Checkʱ�� = Not rsTmp.EOF
'    Set rsTmp = Nothing
'
'   Exit Function
'Hd:
'   If ErrCenter() = 1 Then
'        Resume
'   End If
'   SaveErrLog
'End Function
'
'
'
'
'
'Private Function LoadRegHistory() As Boolean
'    Dim strSQL As String
'    strSQL = " Select Decode(To_Char(a.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',"
'    strSQL = strSQL & vbCrLf & "                       '7', '����') As ������Ŀ, Max(Nvl(a.����, 0)) As ������, Count(1) As ͳ��,to_char(Max(����ʱ��),'hh24:mi:ss') as ����ʱ��"
'    strSQL = strSQL & vbCrLf & " From ���˹Һż�¼ a, �ҺŰ��� b"
'    strSQL = strSQL & vbCrLf & " Where a.��¼״̬ = 1 And a.����ʱ�� Between Sysdate And Sysdate + " & IIf(gintԤԼ���� = 0, 15, gintԤԼ����) & " And a.�ű� = b.���� And b.Id=[1]"
'    strSQL = strSQL & vbCrLf & " Group By Decode(To_Char(a.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',"
'    strSQL = strSQL & vbCrLf & "                             '7', '����')"
'
'    On Error GoTo Hd:
'    Set mrsRegHistory = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
'    LoadRegHistory = True
'Exit Function
'Hd:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Function
'