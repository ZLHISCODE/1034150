VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffPlanCard 
   Caption         =   "���Ĳɹ��ƻ�"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmStuffPlanCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh������ 
      Height          =   2325
      Left            =   240
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   4101
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf��Ӧ��ѡ�� 
      Height          =   2565
      Left            =   6360
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   8
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   4
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   5
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9930
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   1
         Top             =   950
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   3
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lbl���Ʒ��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ʒ���:"
         Height          =   180
         Left            =   8070
         TabIndex        =   24
         Top             =   660
         Width           =   810
      End
      Begin VB.Label txt���Ʒ��� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ٽ��ڼ�ƽ�����շ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9000
         TabIndex        =   23
         Top             =   660
         Width           =   2355
      End
      Begin VB.Label txt�ƻ����� 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1080
         TabIndex        =   22
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "���ϼƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   19
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   18
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   17
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   16
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   15
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "���Ĳɹ��ƻ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label Lbl�ƻ����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ƻ�����:"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   13
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   12
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   7365
         TabIndex        =   11
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9240
         TabIndex        =   10
         Top             =   4500
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1000
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffPlanCard.frx":22EA
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffPlanCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffPlanCard.frx":3080
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCode 
      Caption         =   "����"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmStuffPlanCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelectStock As String           '�Ƿ��ѡ�ⷿ
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean                '��һ����ʾ
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ����� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintUnit As Integer            '0-ɢװ��λ,1-��װ��λ
Private mbln���� As Boolean                 '����ȡ���ڴ������޵�ҩƷ
Private mint���� As Integer
Private mint���� As Integer

Private mlng�ƻ�ID As Long
Private mlng�ⷿid As Long
Private mint�ƻ����� As Integer
Private mint���Ʒ��� As Integer
Private mstr������ID As String      '��id�ָ�
Private mbln�б굥λ As Boolean '�����б깩����,Ҫ��mstr������λһ��������.
Private mstr�ڼ�  As String                  '������λ��ʾ,������λ��ʾ,������λ��ʾ
Dim mstrPrivs As String                     'Ȩ��
Private Const mlngModule = 1724
Private mintУ�鷽ʽ As Integer     '0-����飻1�����ѣ�2����ֹ
Private mblnCheck As Boolean
Private mblnFirstCheck As Boolean
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-�����鿴 false-�������鿴
Private mbln�ƻ����� As Boolean         'true-�����ƻ����� false-�������ƻ�����
Private mstrNow As String               '��¼��ǰ����
Private Const mstrCaption As String = "���Ĳɹ��ƻ�"

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

'=========================================================================================
Private Enum mHeadCol
    ��� = 1
    У�� = 2
    ���� = 3
    ��� = 4
    ���� = 5
    ��λ = 6
    ����ϵ�� = 7
    �б���� = 8
    �洢���� = 9
    �洢���� = 10
    ǰ������ = 11
    �������� = 12
    ������� = 13
    �������� = 14
    �������� = 15
    �ƻ����� = 16
    ���� = 17
    ��� = 18
    �ϴι�Ӧ�� = 19
End Enum

Private Const mconIntColS  As Integer = 20     '������

Private Function CheckQualifications() As Boolean
    '���ݲ���У�����ģ������̣���Ӧ����Ϣ������Ч��
    Dim dateCurrent As Date
    Dim strCheck As String
    Dim strCheck_���� As String
    Dim strCheck_������ As String
    Dim strCheck_��Ӧ�� As String
    Dim intCheckType As Integer
    Dim arrColumn
    Dim rsTmp As ADODB.Recordset
    Dim intRow As Integer
    Dim strTmp_���� As String
    Dim strTmp_������ As String
    Dim strTmp_��Ӧ�� As String
    Dim strMsg_���� As String
    Dim strMsg_������ As String
    Dim strMsg_��Ӧ�� As String
    Dim intCount As Integer
    Dim blnFlag As Boolean
    Dim n As Integer
    Dim strMsgInfo As String
    Dim str�������б� As String
    Dim str��Ӧ���б� As String
    Dim intCount_���� As Integer
    Dim intCount_������ As Integer
    Dim intCount_��Ӧ�� As Integer

'    On Error Resume Next
    On Error GoTo errHandle
    '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    strCheck = zlDatabase.GetPara("����У��", glngSys, mlngModule, "")

    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�鷽ʽ��0-����飻1�����ѣ�2����ֹ
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))

    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�����ݣ�
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '�ֱ�ȡ���ģ������̣���Ӧ����ҪУ�������
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
            If Split(arrColumn(n), ",")(0) = "����" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_���� = IIf(strCheck_���� = "", "", strCheck_���� & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "����������" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_������ = IIf(strCheck_������ = "", "", strCheck_������ & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "���Ĺ�Ӧ��" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_��Ӧ�� = IIf(strCheck_��Ӧ�� = "", "", strCheck_��Ӧ�� & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next

    dateCurrent = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))

    '�ֱ�У�����ģ������̣���Ӧ��
    With mshBill
        .Redraw = False
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                If strCheck_���� <> "" Then
                    gstrSQL = "Select A.����֤��, A.����֤��Ч�� " & _
                              "From �������� A " & _
                              "Where A.����ID = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "У����������", Val(.TextMatrix(intRow, 0)))
                    
                    strTmp_���� = ""
'                    strMsg_���� = ""
                    blnFlag = False
                    
                    If Not rsTmp.EOF Then
                        If NVL(rsTmp!����֤��) = "" And InStr(strCheck_����, "����֤��") > 0 Then
                            strTmp_���� = .TextMatrix(intRow, mHeadCol.����) & "��" & "������֤��"
                            blnFlag = True
                        End If
                        
                        If NVL(rsTmp!����֤��Ч��) <> "" Then
                            If DateDiff("d", rsTmp!����֤��Ч��, dateCurrent) > 0 And InStr(strCheck_����, "����֤��") > 0 Then
                                strTmp_���� = IIf(strTmp_���� = "", .TextMatrix(intRow, mHeadCol.����) & "��", strTmp_���� & ",") & "����֤�ѹ���"
                            blnFlag = True
                            End If
                        End If
                    End If
    
                    If strTmp_���� <> "" Then
                        If intCount_���� <= 5 Then
                            strMsg_���� = strMsg_���� & strTmp_���� & vbCrLf
                        End If
                        intCount_���� = intCount_���� + 1
                    End If
                    If blnFlag = True Then SetBilCheckFlag intRow, mHeadCol.����, False
                End If
                
                If strCheck_������ <> "" And .TextMatrix(intRow, mHeadCol.����) <> "" Then
                    gstrSQL = "Select A.������ҵ����֤, A.������ҵ����֤Ч��,a.��Ӫ����֤, a.��Ӫ����֤Ч��, a.��ҵ����ִ��, a.��ҵ����ִ��Ч�� " & _
                              "From ���������� A " & _
                              "Where A.���� = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "У������������", .TextMatrix(intRow, mHeadCol.����))
                    
                    strTmp_������ = ""
'                    strMsg_������ = ""
                    blnFlag = False
                    
                    If Not rsTmp.EOF Then
                        If NVL(rsTmp!������ҵ����֤) = "" And InStr(strCheck_������ & ";", "������ҵ����֤" & ";") > 0 Then
                            strTmp_������ = .TextMatrix(intRow, mHeadCol.����) & "��" & "��������ҵ����֤"
                            blnFlag = True
                        End If
                        If NVL(rsTmp!������ҵ����֤Ч��) <> "" Then
                            If DateDiff("d", rsTmp!������ҵ����֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "������ҵ����֤Ч��" & ";") > 0 Then
                                strTmp_������ = IIf(strMsg_������ = "", .TextMatrix(intRow, mHeadCol.����) & "��", strTmp_������ & ",") & "������ҵ����֤�ѹ���"
                                blnFlag = True
                            End If
                        End If
                        
                        If NVL(rsTmp!��Ӫ����֤) = "" And InStr(strCheck_������ & ";", "��Ӫ����֤" & ";") > 0 Then
                            strTmp_������ = .TextMatrix(intRow, mHeadCol.����) & "��" & "�޾�Ӫ����֤"
                            blnFlag = True
                        End If
                        If NVL(rsTmp!��Ӫ����֤Ч��) <> "" Then
                            If DateDiff("d", rsTmp!��Ӫ����֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "��Ӫ����֤Ч��" & ";") > 0 Then
                                strTmp_������ = IIf(strMsg_������ = "", .TextMatrix(intRow, mHeadCol.����) & "��", strTmp_������ & ",") & "��Ӫ����֤�ѹ���"
                                blnFlag = True
                            End If
                        End If
                        
                        If NVL(rsTmp!��ҵ����ִ��) = "" And InStr(strCheck_������ & ";", "��ҵ����ִ��" & ";") > 0 Then
                            strTmp_������ = .TextMatrix(intRow, mHeadCol.����) & "��" & "����ҵ����ִ��"
                            blnFlag = True
                        End If
                        If NVL(rsTmp!��ҵ����ִ��Ч��) <> "" Then
                            If DateDiff("d", rsTmp!��ҵ����ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "��ҵ����ִ��Ч��" & ";") > 0 Then
                                strTmp_������ = IIf(strMsg_������ = "", .TextMatrix(intRow, mHeadCol.����) & "��", strTmp_������ & ",") & "��ҵ����ִ���ѹ���"
                                blnFlag = True
                            End If
                        End If
                    End If
                    
                    If strTmp_������ <> "" Then
                        If InStr(1, str�������б�, .TextMatrix(intRow, mHeadCol.����)) = 0 Then
                            str�������б� = IIf(str�������б� = "", "", str�������б� & ",") & .TextMatrix(intRow, mHeadCol.����)
                            
                            If intCount_������ <= 5 Then
                                strMsg_������ = strMsg_������ & strTmp_������ & vbCrLf
                            End If
                            intCount_������ = intCount_������ + 1
                        End If
                    End If
                    If blnFlag = True Then SetBilCheckFlag intRow, mHeadCol.����, False
                    
                End If
                
                If strCheck_��Ӧ�� <> "" And .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) <> "" Then
                    gstrSQL = "Select ˰��ǼǺ�, ����֤��, ִ�պ�, ��Ȩ��, ������֤��, ҩ��ֱ�����, ����֤Ч��, ִ��Ч��, ��Ȩ�� " & _
                              "From ��Ӧ�� " & _
                              "Where (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And ���� = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ӧ����Ϣ", .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��))
                    
                    strTmp_��Ӧ�� = ""
'                    strMsg_��Ӧ�� = ""
                    blnFlag = False
                    
                    If Not rsTmp.EOF Then
                        If NVL(rsTmp!˰��ǼǺ�) = "" And InStr(strCheck_��Ӧ��, "˰��ǼǺ�") > 0 Then
                            strTmp_��Ӧ�� = .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) & "��" & "��˰��ǼǺ�"
                            blnFlag = True
                        End If
                        
                        If NVL(rsTmp!����֤��) = "" And InStr(strCheck_��Ӧ��, "����֤��") > 0 Then
                            strTmp_��Ӧ�� = IIf(strTmp_��Ӧ�� = "", .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) & "��", strTmp_��Ӧ�� & ",") & "������֤��"
                            blnFlag = True
                        End If
                        
                        If NVL(rsTmp!ִ�պ�) = "" And InStr(strCheck_��Ӧ��, "ִ�պ�") > 0 Then
                            strTmp_��Ӧ�� = IIf(strTmp_��Ӧ�� = "", .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) & "��", strTmp_��Ӧ�� & ",") & "��ִ�պ�"
                            blnFlag = True
                        End If
                        
                        If NVL(rsTmp!��Ȩ��) = "" And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                            strTmp_��Ӧ�� = IIf(strTmp_��Ӧ�� = "", .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) & "��", strTmp_��Ӧ�� & ",") & "����Ȩ��"
                            blnFlag = True
                        End If
                        
                        If NVL(rsTmp!ҩ��ֱ�����) = "" And InStr(strCheck_��Ӧ��, "ҩ��ֱ�����") > 0 Then
                            strTmp_��Ӧ�� = IIf(strTmp_��Ӧ�� = "", .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) & "��", strTmp_��Ӧ�� & ",") & "��ҩ��ֱ�����"
                            blnFlag = True
                        End If
                        
                        If NVL(rsTmp!����֤Ч��) <> "" Then
                            If DateDiff("d", rsTmp!����֤Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "����֤Ч��") > 0 Then
                                strTmp_��Ӧ�� = IIf(strTmp_��Ӧ�� = "", .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) & "��", strTmp_��Ӧ�� & ",") & "����֤�ѹ���"
                                blnFlag = True
                            End If
                        End If
                        
                        If NVL(rsTmp!ִ��Ч��) <> "" Then
                            If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ִ��Ч��") > 0 Then
                                strTmp_��Ӧ�� = IIf(strTmp_��Ӧ�� = "", .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) & "��", strTmp_��Ӧ�� & ",") & "ִ���ѹ���"
                                blnFlag = True
                            End If
                        End If
                        
                        If NVL(rsTmp!��Ȩ��) <> "" Then
                            If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                                strTmp_��Ӧ�� = IIf(strTmp_��Ӧ�� = "", .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) & "��", strTmp_��Ӧ�� & ",") & "��Ȩ�ѹ���"
                                blnFlag = True
                            End If
                        End If
                    End If
                    
                    If strTmp_��Ӧ�� <> "" Then
                        If InStr(1, str��Ӧ���б�, .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��)) = 0 Then
                            str��Ӧ���б� = IIf(str��Ӧ���б� = "", "", str��Ӧ���б� & ",") & .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��)
                            
                            If intCount_��Ӧ�� <= 5 Then
                                strMsg_��Ӧ�� = strMsg_��Ӧ�� & strTmp_��Ӧ�� & vbCrLf
                            End If
                            intCount_��Ӧ�� = intCount_��Ӧ�� + 1
                        End If
                    End If
                    If blnFlag = True Then SetBilCheckFlag intRow, mHeadCol.�ϴι�Ӧ��, False
                End If
                 
                If strTmp_���� = "" And strTmp_������ = "" And strTmp_��Ӧ�� = "" Then
                    SetBilCheckFlag intRow, mHeadCol.У��, True
                End If
            End If
        Next
        .Redraw = True
    End With
    
    If strMsg_���� <> "" Then
        strMsg_���� = "���ģ�" & vbCrLf & strMsg_����
        If intCount_���� > 5 Then strMsg_���� = strMsg_���� & vbCrLf & "....."
        strMsgInfo = IIf(strMsgInfo = "", "", strMsgInfo & vbCrLf) & strMsg_����
    End If
    
    If strMsg_������ <> "" Then
        strMsg_������ = "�����̣�" & vbCrLf & strMsg_������
        If intCount_������ > 5 Then strMsg_������ = strMsg_������ & vbCrLf & "....."
        strMsgInfo = IIf(strMsgInfo = "", "", strMsgInfo & vbCrLf) & strMsg_������
    End If
    
    If strMsg_��Ӧ�� <> "" Then
        strMsg_��Ӧ�� = "��Ӧ�̣�" & vbCrLf & strMsg_��Ӧ��
        If intCount_��Ӧ�� > 5 Then strMsg_��Ӧ�� = strMsg_��Ӧ�� & vbCrLf & "....."
        strMsgInfo = IIf(strMsgInfo = "", "", strMsgInfo & vbCrLf) & strMsg_��Ӧ��
    End If
    
    If strMsgInfo <> "" Then
        strMsgInfo = "������Ŀ����У��δͨ�������飺" & vbCrLf & strMsgInfo
        MsgBox strMsgInfo, vbOKOnly, gstrSysName
        CheckQualifications = False
        Exit Function
    End If
    
    CheckQualifications = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetBilCheckFlag(ByVal intRow As Integer, ByVal intCol As Integer, ByVal blnFlag As Boolean)
    '����У����
    'blnFlag��True-��intRow�У�intCol�д򹴣���ʾ������Ŀ��У��ͨ����False�����򹴣�����intRow�У�intCol���Ϻ�ɫ�����ʶ
    Dim i As Integer
    With mshBill
        If blnFlag = False Then
            .Row = intRow
            i = .ColData(intCol)
            .ColData(intCol) = 0
            .Col = intCol
            .MsfObj.CellForeColor = vbRed
            .MsfObj.CellFontBold = True
            .ColData(intCol) = i
        Else
            .TextMatrix(intRow, intCol) = "��"
        End If
    End With
End Sub

Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, _
        ByVal int�༭״̬ As Integer, Optional blnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1724)

    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True

    Set mfrmMain = frmMain
    mblnCostView = IsHavePrivs(mstrPrivs, "�鿴�ɱ���")

    If mint�༭״̬ = 1 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If

    End If

    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�

End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()

    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindData mshBill, mHeadCol.����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindData mshBill, mHeadCol.����, txtCode.Text, False
    ElseIf KeyCode = vbKeyEscape Then
        If Msf��Ӧ��ѡ��.Visible Then
            Msf��Ӧ��ѡ��.ZOrder 1
            Msf��Ӧ��ѡ��.Visible = False
            Exit Sub
        End If
        Call CmdCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        Call FrmBillPrint.ShowMe(Me, glngSys, "zl1_bill_1724", 0, mintUnit, 1724, "���Ĳɹ��ƻ���", txtNO.Tag)
        '�˳�
        Unload Me
        Exit Sub
    End If

    If mint�༭״̬ = 3 Then        '���
        '����У��
        If mblnFirstCheck = False Then
            mblnCheck = CheckQualifications
            mblnFirstCheck = True
            If mblnCheck = False Then
                Exit Sub
            End If
        End If
        
        If mblnCheck = False Then
            If mintУ�鷽ʽ = 1 Then
                If MsgBox("�������ģ������̣���Ӧ��δͨ��У�飬�Ƿ���ˣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                
                If SaveCheckCard = False Then Exit Sub
            ElseIf mintУ�鷽ʽ = 2 Then
                MsgBox "�������ģ������̣���Ӧ��δͨ��У�飬������ˣ�", vbOKOnly, gstrSysName
                Exit Sub
            End If
        End If
        
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    ReportOpen gcnOracle, glngSys, "zl1_bill_1724", Me, "���ݱ��=" & txtNO.Tag, "��λ=" & mintUnit, 2
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If

    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard

    If blnSuccess = True Then

        If IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                ReportOpen gcnOracle, glngSys, "zl1_bill_1724", Me, "���ݱ��=" & txtNO.Tag, "��λ=" & mintUnit, 2
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    txtժҪ.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
End Sub

Private Function GetDeptRequestDataBill(ByVal strNOIn As String, ByVal lng�ⷿID As Long, _
    ByVal strStartDate As String, ByVal strEndDate As String, ByVal str����IDIN As String) As Boolean
    '-----------------------------------------------------------------------------------------------------
    '����:�Ӳ����깺�л�ȡ����
    '����:strNOIn-���ݺ�
    '     �ⷿID-�ⷿ
    '    strStartDate-��ʼ����
    '    strEndDate-��������
    '    str����IDIN-����ID_IN
    '����,���óɹ�,����true,���򷵻�False:
    '-----------------------------------------------------------------------------------------------------

    Dim strSQL As String
    Dim rsplan As New Recordset
    Dim lngRecord As Long, lngProcess As Long
    Dim intLop As Integer
    strNOIn = Replace(strNOIn, "'", "")
    Me.MousePointer = vbHourglass
    mshBill.Redraw = False
    stbThis.Panels(2).Text = "���ڼ���"
    
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    Pic����.Enabled = False

    err = 0: On Error GoTo ErrHand:
    If str����IDIN <> "" Then
          strSQL = "" & _
                  "   Select  /*+ Rule*/ A.����id,('['|| q.���� || ']' || q.����) as ������Ϣ,b.�б����,q.���,nvl(max(A.�ϴ�������),max(q.����)) as ����,max(A.�ϴι�Ӧ��) as �ϴι�Ӧ��," & _
                  "      sum(nvl(A.�빺����,0)) as �빺����," & _
                  "      sum(nvl(A.�ƻ�����,0)) as ��������," & _
                          IIf(mintUnit = 0, "Q.���㵥λ", "B.��װ��λ") & " as ��λ, " & _
                          IIf(mintUnit = 0, "1", "B.����ϵ��") & " as ����ϵ��, " & _
                  "       max(A.����) as ����," & _
                  "       sum(nvl(A.���,0)) as ��� " & _
                  "   From ���ϼƻ����� A,�������� B,���ϲɹ��ƻ� c,�շ���ĿĿ¼ Q,������ĿĿ¼ M, " & _
                  "       Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) J, " & _
                  "       Table(Cast(f_Str2list([2]) As zlTools.t_StrList)) L" & _
                  "   Where A.����id=B.����id and A.����id=q.id And (q.վ��=[6] or q.վ�� is null) " & _
                  "         And a.�ƻ�id=c.id and c.����=1" & IIf(lng�ⷿID = 0, "", " And C.�ⷿid=[3]") & _
                  "         And (C.������� between [4] and [5]) And C.No =L.Column_Value" & _
                  "         And B.����id=M.id and M.����id=J.Column_Value " & _
                  "   Group by A.����id,q.���� ,q.����,q.���,q.����,b.�б����,B.����ϵ��," & IIf(mintUnit = 0, "Q.���㵥λ", "B.��װ��λ")
      Else
          strSQL = "" & _
                  "   Select /*+ Rule*/ A.����id, ('['|| q.���� || ']' || q.����) as ������Ϣ,B.�б����,q.��� ,nvl(max(A.�ϴ�������),q.����) as ����,max(A.�ϴι�Ӧ��) as �ϴι�Ӧ��," & _
                  "      sum(nvl(A.�빺����,0)) as �빺����," & _
                  "      sum(nvl(A.�ƻ�����,0)) as ��������," & _
                         IIf(mintUnit = 0, "Q.���㵥λ", "B.��װ��λ") & " as ��λ, " & _
                         IIf(mintUnit = 0, "1", "B.����ϵ��") & " as ����ϵ��, " & _
                  "      max(A.����) as ����," & _
                  "      sum(nvl(A.���,0)) as ��� " & _
                  "   From ���ϼƻ����� A,�������� B,���ϲɹ��ƻ� c,�շ���ĿĿ¼ Q," & _
                  "       Table(Cast(f_Str2list([2]) As zlTools.t_StrList)) L" & _
                  "   Where A.����id=B.����id and A.����id=q.id And (q.վ��=[6] or q.վ�� is null) " & _
                  "         And a.�ƻ�id=c.id and c.����=1 " & IIf(lng�ⷿID = 0, "", " And C.�ⷿid=[3]") & _
                  "         And (C.������� between [4] and [5]) And C.No =L.Column_Value" & _
                  "   Group by A.����id,q.���� ,q.����,q.���,q.����,b.�б����,B.����ϵ��," & IIf(mintUnit = 0, "Q.���㵥λ", "B.��װ��λ") & _
                  "   "
      End If
    
    strSQL = "" & _
    "   Select  A.����ID,A.������Ϣ,A.���,A.�б����,nvl(max(a.����),a.����) as ����,nvl(max(�ϴι�Ӧ��),A.�ϴι�Ӧ��) as �ϴι�Ӧ��," & _
    "           A.�빺����,A.��������,A.��λ,A.����ϵ��," & _
    "           nvl(max(b.�ϴβɹ���),a.����)  as ����," & _
    "           sum(nvl(B.ʵ������,0)) as �������" & _
    "   From (" & strSQL & ") A,ҩƷ��� B,��Ӧ�� C" & _
    "   Where a.����id=b.ҩƷid(+) and nvl(B.�ϴι�Ӧ��id,0)=C.id(+) and b.����(+)=1  " & IIf(lng�ⷿID = 0, "", " And B.�ⷿid(+)=[3]") & _
    "   Group by A.����ID,A.������Ϣ,A.���,A.�б����,a.����,A.�ϴι�Ӧ��,A.�빺����,A.��������,A.��λ,A.����ϵ��,A.����,A.���" & _
    "   order by ������Ϣ"
    
    Set rsplan = zlDatabase.OpenSQLRecord(strSQL, mstrCaption, str����IDIN, strNOIn, lng�ⷿID, CDate(strStartDate), CDate(strEndDate & " 23:59:59"), gstrNodeNo)
    With rsplan
        lngRecord = .RecordCount
        If lngRecord = 0 Then
            mshBill.Redraw = True
            Me.MousePointer = vbDefault
            CmdSave.Enabled = True
            CmdCancel.Enabled = True
            Pic����.Enabled = True
            Me.stbThis.Panels(2).Text = ""
            GetDeptRequestDataBill = True
            Exit Function
        End If
        lngProcess = 0
        If .RecordCount <> 0 Then .MoveFirst
        For intLop = 1 To .RecordCount
            mshBill.TextMatrix(intLop, 0) = Val(NVL(!����ID))
            mshBill.TextMatrix(intLop, mHeadCol.����) = NVL(!������Ϣ)
            mshBill.TextMatrix(intLop, mHeadCol.���) = NVL(!���)
            mshBill.TextMatrix(intLop, mHeadCol.����) = NVL(!����)
            mshBill.TextMatrix(intLop, mHeadCol.��λ) = NVL(!��λ)
            mshBill.TextMatrix(intLop, mHeadCol.ǰ������) = ""
            mshBill.TextMatrix(intLop, mHeadCol.��������) = ""
            mshBill.TextMatrix(intLop, mHeadCol.�������) = Format(Val(NVL(!�������)) / Val(NVL(!����ϵ��)), mFMT.FM_����)
            mshBill.TextMatrix(intLop, mHeadCol.�ƻ�����) = Format(Val(NVL(!��������)) / Val(NVL(!����ϵ��)), mFMT.FM_����)
            mshBill.TextMatrix(intLop, mHeadCol.����) = Format(Val(NVL(!����)) * Val(NVL(!����ϵ��)), mFMT.FM_�ɱ���)
            mshBill.TextMatrix(intLop, mHeadCol.���) = Format(Val(NVL(!����)) * Val(NVL(!��������)), mFMT.FM_���)
            mshBill.TextMatrix(intLop, mHeadCol.����ϵ��) = NVL(!����ϵ��)
            mshBill.TextMatrix(intLop, mHeadCol.�ϴι�Ӧ��) = NVL(!�ϴι�Ӧ��)
            mshBill.TextMatrix(intLop, mHeadCol.�б����) = NVL(!�б����)
            
            Call Calc����(Val(NVL(!����ID)), intLop)
            
            If intLop >= mshBill.Rows - 1 Then mshBill.Rows = mshBill.Rows + 1
            lngProcess = lngProcess + 1
            Call ShowPercent(lngProcess / lngRecord)
            .MoveNext
        Next
        Call ��ʾ�ϼƽ��
        .Close
    End With
    GetDeptRequestDataBill = True
    Call RefreshRowNO(mshBill, mHeadCol.���, 1)
    Call ��ʾ�ϼƽ��
    Me.MousePointer = vbDefault
    mshBill.Redraw = True
    CmdSave.Enabled = True
    Pic����.Enabled = True
    CmdCancel.Enabled = True
    mshBill.Col = mHeadCol.�ƻ�����
    Me.stbThis.Panels(2).Text = ""
    zlCommFun.StopFlash
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Activate()
    Dim intMonth As Integer

    If mblnFirst = False Then Exit Sub
    
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram stbThis, gSystem_Para.int���뷽ʽ
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    If mint�༭״̬ = 1 Then
        Dim str����ID As String, str���ͱ��� As String
        Dim lng�ⷿID As Long, int�ƻ����� As Integer, int���Ʒ��� As Integer
        
        If frmStuffPlanCondition.GetCondition(mfrmMain, str����ID, lng�ⷿID, int�ƻ�����, int���Ʒ���, mbln����, mint����, mint����, mstr������ID, mbln�б굥λ, mbln�ƻ�����) = True Then
            mlng�ⷿid = lng�ⷿID
            mint�ƻ����� = int�ƻ�����
            mint���Ʒ��� = int���Ʒ���
            
            Select Case mint�ƻ�����
                Case 1       '�¼ƻ�
                    mstr�ڼ� = Format(DateAdd("m", 1, zlDatabase.Currentdate), "yyyyMM")
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr�ڼ�, 1, 4) & "��" & Right(mstr�ڼ�, 2) & "��" & ") " & LblTitle.Tag & "�ɹ��ƻ�"
'                    mshBill.TextMatrix(0, mHeadCol.��������) = "��������"
'                    mshBill.TextMatrix(0, mHeadCol.��������) = "��������"
                Case 2       '���ƻ�
                    intMonth = Month(DateAdd("Q", 1, zlDatabase.Currentdate))
                    mstr�ڼ� = Format(DateAdd("Q", 1, zlDatabase.Currentdate), "yyyy") & IIf(intMonth <= 3, 1, IIf(intMonth >= 10, 4, IIf(intMonth <= 9 And intMonth >= 7, 3, 2)))
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr�ڼ�, 1, 4) & "��" & Right(mstr�ڼ�, 1) & "��" & ")" & LblTitle.Tag & "�ɹ��ƻ�"
'                    mshBill.TextMatrix(0, mHeadCol.��������) = "�ϼ�������"
'                    mshBill.TextMatrix(0, mHeadCol.��������) = "����������"
                Case 3    '��ƻ�
                    mstr�ڼ� = Format(DateAdd("yyyy", 1, zlDatabase.Currentdate), "yyyy")
                    LblTitle.Caption = GetUnitName & "(" & mstr�ڼ� & "��" & ")" & LblTitle.Tag & "�ɹ��ƻ�"
'                    mshBill.TextMatrix(0, mHeadCol.��������) = "��������"
'                    mshBill.TextMatrix(0, mHeadCol.��������) = "��������"
                Case 4      '�ܼƻ�
                    mstr�ڼ� = Format(DateAdd("ww", 1, sys.Currentdate), "yyyyWW")
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr�ڼ�, 1, 4) & "��" & Right(mstr�ڼ�, 2) & "��" & ") " & LblTitle.Tag & "�ɹ��ƻ�"
                   
            End Select
            If mint���Ʒ��� = 5 Then
                Dim strStartDate As String, strEndDate As String, strNOIn As String
                '�������깺���Ʋɹ��ƻ�
                 If FrmBillSelect.ShowCard(str����ID, mlng�ⷿid, mint�ƻ�����, strNOIn, strStartDate, strEndDate) = False Then Unload Me: Exit Sub
                 If GetDeptRequestDataBill(strNOIn, mlng�ⷿid, strStartDate, strEndDate, str����ID) = False Then Exit Sub
                 Exit Sub
                 
            Else
                ReFreshALLStuff str����ID, lng�ⷿID, int�ƻ�����, int���Ʒ���
            End If
        Else
            Unload Me
            Exit Sub
        End If
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    Else
        mblnChange = False
        Select Case mintParallelRecord
            Case 1
                '����
            Case 2
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
            Case 3
                '�޸ĵĵ����ѱ����
                MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
        End Select
    End If

End Sub

Private Sub ReFreshALLStuff(ByVal str����ID, _
    ByVal lng�ⷿID As Long, ByVal int�ƻ����� As Integer, ByVal int���Ʒ��� As Integer)
        '---------------------------------------------------
        '--����:������ҩƷ���мƻ�����
        '--����:
        '---------------------------------------------------
    Dim rsAllStuff As New ADODB.Recordset, rspurchase As New ADODB.Recordset
    Dim lngProcess  As Long, lngRecord As Long, lngRow As Long
    Dim blnOK As Boolean
    
    On Error GoTo errHandle
    Me.Refresh
    Me.MousePointer = vbHourglass
    mshBill.Redraw = False
    stbThis.Panels(2).Text = "���ڼ���"
    
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    Pic����.Enabled = False

    Dim str��λ As String
    
    Select Case mintUnit
    Case 0
        str��λ = ",F.���㵥λ ��λ,1 ����ϵ��"
    Case Else
        str��λ = ",A.��װ��λ ��λ,A.����ϵ�� ����ϵ��"
    End Select

    'ȡָ��������ҩƷ��Ϣ
    gstrSQL = "" & _
         " SELECT /*+ Rule*/ DISTINCT A.����id ҩƷID,A.�б����,F.����,NVL(B.����,F.����) AS ͨ������," & _
         "      F.���" & str��λ & ",DECODE(A.�ɱ���,NULL,NVL(A.ָ��������,0),0,NVL(A.ָ��������,0),NVL(A.�ɱ���,0)) AS ����,F.����" & _
         " FROM �������� A,�շ���Ŀ���� B,������ĿĿ¼ C,���Ʒ���Ŀ¼ L,�շ���ĿĿ¼ F " & _
         IIf(str����ID = "", "", ",Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) D ") & _
         " WHERE A.����ID=F.ID And (f.վ��=[3] or f.վ�� is null) And A.����ID=C.ID And C.����ID=L.ID and L.���� =7" & _
         "          And A.����ID = B.�շ�ϸĿID(+) And B.����(+)=3 " & _
         "          AND (F.����ʱ��>=TO_DATE('3000-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS') OR F.����ʱ�� IS NULL)" & _
                    IIf(str����ID = "", IIf(mstr������ID <> "", "", " And L.ID Is NULL"), " AND L.ID =D.Column_Value ")

    '��ҩƷ�������ȡ�п������ĵĹ�Ӧ���������Ϣ���޿�������ֻȡ���һ�����Ĺ�Ӧ���������Ϣ
    
    If lng�ⷿID = 0 Then
        '�����ȫ�ⷿ��ȡ���пⷿ��棬����ҩƷ�����ȡ�ϴι�Ӧ�̺��ϴβ���
        gstrSQL = "( " & gstrSQL & ") D," & _
                  " (Select a.ҩƷid, c.Id As �ϴι�Ӧ��id, c.���� As ��Ӧ��, b.�ϴβ���, a.�������, a.ƽ���ۼ� " & _
                " From (Select ҩƷid, Sum(ʵ������) As �������, " & _
                "              Decode(Sign(Sum(ʵ������)), 1, Decode(Sign(Sum(ʵ�ʽ��)), 1, Sum(ʵ�ʽ��), 0) / Sum(ʵ������), 0) ƽ���ۼ� " & _
                "       From ҩƷ��� " & _
                "       Where ���� = 1 " & _
                "       Group By ҩƷid) A, �������� B, (Select ID, ���� From ��Ӧ�� Where Substr(����, 5, 1) = 1) C " & _
                " Where a.ҩƷid = b.����id And b.�ϴι�Ӧ��id = c.Id(+)) E "
    Else
        'ȡ�����������������εĹ�Ӧ�̣��ϴβ���
        gstrSQL = "( " & gstrSQL & ") D," & _
                  " (   Select A.ҩƷID,C.ID �ϴι�Ӧ��ID, C.���� As ��Ӧ��, B.�ϴβ���, A.�������, A.ƽ���ۼ� " & _
                  "     From (  Select �ⷿid, ҩƷid, Sum(ʵ������) As �������, " & _
                  "                     Decode(Sign(Sum(ʵ������)), 1, Decode(Sign(Sum(ʵ�ʽ��)), 1, Sum(ʵ�ʽ��), 0) / Sum(ʵ������), 0) ƽ���ۼ� " & _
                  "             From ҩƷ��� " & _
                  "             Where ���� = 1 " & IIf(lng�ⷿID = 0, "", " AND �ⷿID= [1]") & _
                  "             Group By �ⷿid, ҩƷid) A, " & _
                  "          (  Select �ⷿid,ҩƷid,����,�ϴι�Ӧ��ID,�ϴβ��� From ҩƷ��� " & _
                  "             Where ���� = 1 " & IIf(lng�ⷿID = 0, "", " AND �ⷿID= [1]") & _
                  "                     And (ҩƷID,Nvl(����, 0)) in  (Select ҩƷid,Nvl(Max(Nvl(����, 0)), 0) ���� From ҩƷ��� Where ���� = 1 " & IIf(lng�ⷿID = 0, "", " AND �ⷿID=[1] ") & " group by ҩƷid )" & _
                  "                                   ) B, " & _
                  "          (SELECT ID,���� FROM ��Ӧ�� WHERE SUBSTR(����,5,1)=1 ) C " & _
                  "     Where A.�ⷿid = B.�ⷿid And A.ҩƷid = B.ҩƷid And B.�ϴι�Ӧ��id = C.ID(+) " & _
                  "     ) E "
    End If
    '������ȡ���ϴ����޶��SQL
    gstrSQL = gstrSQL & _
            " ,  (  Select ����id ҩƷID,sum(nvl(����,0)) ����,sum(nvl(����,0)) ���� " & _
            "       From ���ϴ����޶�  " & _
            "       " & IIf(lng�ⷿID = 0, "", " Where �ⷿID=[1]") & _
            "       Group By ����ID)     F"

    '�������У�����������ȡ���ϴ����޶�.���ޣ�
    gstrSQL = "" & _
        "   SELECT d.ҩƷid,D.�б����,e.�ϴι�Ӧ��ID, d.����, d.ͨ������, d.���, " & _
        "           DECODE (e.�ϴβ���, NULL, d.����, e.�ϴβ���) AS ����," & _
        "           d.��λ,nvl(e.�������,0)/d.����ϵ�� as ������� ,f.����/d.����ϵ�� ����,f.����/d.����ϵ�� ���� , d.���� as ���� , e.��Ӧ��,d.����ϵ�� from " & _
                gstrSQL & _
        " WHERE d.ҩƷid = e.ҩƷid (+) "
    '���ϴ����޶���жϣ����ڴ����޶��ҩƷ����ȡ�������ɹ��ƻ�
    gstrSQL = gstrSQL & " And d.ҩƷID=F.ҩƷID(+)"
    If mbln���� Then
        '���������ж�
        gstrSQL = "Select * From (" & gstrSQL & ") Where (�������<���� and ����<>0)"
    End If
    gstrSQL = gstrSQL & " Order by ����"
        
    Set rsAllStuff = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng�ⷿID, str����ID, gstrNodeNo)
    
    With rsAllStuff
        lngRecord = .RecordCount

        If lngRecord = 0 Then
            mshBill.Redraw = True
            Me.MousePointer = vbDefault
            CmdSave.Enabled = True
            CmdCancel.Enabled = True
            Pic����.Enabled = True
            Me.stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        .MoveFirst
        Me.Refresh
        DoEvents
        Dim str�ϴι�Ӧ�� As String
        
        lngRow = 0
        lngProcess = 1
        Do While Not .EOF
            blnOK = False
            str�ϴι�Ӧ�� = ""
            If mstr������ID = "" Then
                blnOK = True
            Else
                If Val(NVL(!�ϴι�Ӧ��id)) = 0 And mbln�б굥λ Then
                    gstrSQL = "Select b.���� from �����б굥λ a,��Ӧ�� b,Table(cast(f_Num2List([2]) as zlTools.t_NumList)) C " & _
                              "Where a.����ID=[1] and (b.վ��=[2] or b.վ�� is null) " & _
                              "    and a.��λid=b.id and a.��λID=c.Column_Value "
                    Set rspurchase = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(NVL(!ҩƷID)), mstr������ID, gstrNodeNo)
                    If rspurchase.RecordCount <> 0 Then
                        blnOK = True
                        str�ϴι�Ӧ�� = NVL(rspurchase!����)
                    End If
                Else
                     If "," & mstr������ID & "," Like "*," & Val(NVL(!�ϴι�Ӧ��id)) & ",*" Then
                         blnOK = True
                     End If
                End If
            End If
            If blnOK Then
                    lngRow = lngRow + 1
                    mshBill.TextMatrix(lngRow, 0) = !ҩƷID
                    mshBill.TextMatrix(lngRow, mHeadCol.����) = "[" & !���� & "]" & !ͨ������
                    mshBill.TextMatrix(lngRow, mHeadCol.���) = IIf(IsNull(!���), "", !���)
                    mshBill.TextMatrix(lngRow, mHeadCol.����) = IIf(IsNull(!����), "", !����)
                    mshBill.TextMatrix(lngRow, mHeadCol.��λ) = IIf(IsNull(!��λ), "", !��λ)
                    mshBill.TextMatrix(lngRow, mHeadCol.����) = Format(Val(NVL(!����)) * NVL(!����ϵ��, 1), mFMT.FM_�ɱ���)
                    
                    mshBill.TextMatrix(lngRow, mHeadCol.�ϴι�Ӧ��) = IIf(IsNull(!��Ӧ��), str�ϴι�Ӧ��, !��Ӧ��)
                    mshBill.TextMatrix(lngRow, mHeadCol.�������) = Format(Val(NVL(!�������)), mFMT.FM_����)
                    mshBill.TextMatrix(lngRow, mHeadCol.����ϵ��) = NVL(!����ϵ��, 1)
                    mshBill.TextMatrix(lngRow, mHeadCol.�б����) = IIf(NVL(!�б����) = 1, "��", "")
                    
                    mshBill.TextMatrix(lngRow, mHeadCol.�洢����) = Format(Val(NVL(!����)), mFMT.FM_����)
                    mshBill.TextMatrix(lngRow, mHeadCol.�洢����) = Format(Val(NVL(!����)), mFMT.FM_����)
                    
                    SetNumer !ҩƷID, lng�ⷿID, Val(NVL(!�������)), lngRow, int�ƻ�����, int���Ʒ���
                    If lngRow = mshBill.Rows - 1 Then mshBill.Rows = mshBill.Rows + 1
            End If
            lngProcess = lngProcess + 1
            Call ShowPercent(lngProcess / lngRecord)
            .MoveNext
        Loop
    End With
    Call RefreshRowNO(mshBill, mHeadCol.���, 1)
    Call ��ʾ�ϼƽ��
    Me.MousePointer = vbDefault
    mshBill.Redraw = True
    CmdSave.Enabled = True
    Pic����.Enabled = True
    CmdCancel.Enabled = True
    mshBill.Col = mHeadCol.�ƻ�����
    Me.stbThis.Panels(2).Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDate(ByVal intģʽ As Integer, ByVal datCurrent As Date, _
        ByRef strBegin As String, ByRef strEnd As String) As Boolean
    Dim rsdate As New Recordset

    'intģʽ=1,�¼ƻ���2�����ƻ�
    On Error GoTo errHandle
    GetDate = False
    If intģʽ = 1 Then
        strBegin = Year(datCurrent) & "-" & String(2 - Len(Month(datCurrent)), "0") & Month(datCurrent) & "-01"
        gstrSQL = "select last_day(to_date([1],'yyyy-mm-dd')) from dual"
        Set rsdate = zlDatabase.OpenSQLRecord(gstrSQL, "GetDate", Format(datCurrent, "yyyy-mm-dd"))
        
        strEnd = Format(rsdate.Fields(0), "yyyy-mm-dd")
        rsdate.Close
    Else
        Select Case DatePart("Q", datCurrent)
            Case 1
                strBegin = Year(datCurrent) & "-01-01"
                strEnd = Year(datCurrent) & "-03-31"
            Case 2
                strBegin = Year(datCurrent) & "-04-01"
                strEnd = Year(datCurrent) & "-06-30"
            Case 3
                strBegin = Year(datCurrent) & "-07-01"
                strEnd = Year(datCurrent) & "-09-30"
            Case 4
                strBegin = Year(datCurrent) & "-10-01"
                strEnd = Year(datCurrent) & "-12-31"
        End Select
    End If
    GetDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'����ǰ�������������������ƻ�����,����
Private Sub SetNumer(ByVal lngҩƷid As Long, ByVal lng�ⷿID As Long, _
        ByVal num������� As Double, ByVal intCurrentRow As Integer, _
        ByVal int�ƻ����� As Integer, ByVal int���Ʒ��� As Integer)
    '---------------------------------------------------------------------------
    '--����:ȷ�����������ͼƻ�����
    '   1 )����ͬ�����Բ��շ�������ȥǰ��ͬ��ҩƷ����������������Թ滮ԭ��Ԥ�����ģ��Աȿ������ɹ��ƻ����û��޸ĵ���
    '   2 )�ٽ��ڼ�ƽ�����շ�����ͬ���ٽ��ڼ�(ǰ�ڡ�����)��ƽ������Ԥ�����ĶԱȿ������ɹ��ƻ����û��޸ĵ�����
    '   3 )ҩƷ�������շ�������ҩƷ������������������õĲ��ΪҩƷ�ƻ��ɹ�����

    '--����:
    '       int�ƻ�����:1:�¶ȼƻ�,2.���ȼƻ�,3.��ȼƻ�
    '       int���Ʒ���:1 ��ʾ����ͬ�����Բ��շ�,2 �ٽ��ڼ�ƽ�����շ�,3.�����޶�;4.��������
    '--����:
    '---------------------------------------------------------------------------
    Dim numǰ������ As Double
    Dim num�������� As Double
    Dim num�ƻ����� As Double
    Dim num���� As Double, num���� As Double
    Dim lng���� As Long

    Dim datǰ�� As Date
    Dim dat���� As Date
    Dim strBegin As String
    Dim strEnd As String
    Dim rsNum As New Recordset
    
    On Error GoTo errHandle
    With mshBill
        Select Case int���Ʒ���
            Case 1      '����ͬ�����β���   ֻ���¶Ⱥͼ��ȼƻ�
                datǰ�� = DateAdd("m", Choose(int�ƻ�����, 1, 3), DateAdd("yyyy", -2, zlDatabase.Currentdate))
                dat���� = DateAdd("m", Choose(int�ƻ�����, 1, 3), DateAdd("yyyy", -1, zlDatabase.Currentdate))
    
    
                If lng�ⷿID = 0 Then
                    GetDate int�ƻ�����, datǰ��, strBegin, strEnd
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS ǰ������ " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "           and ���� <>19 and ����>=15 AND b.ϵ�� = -1 " & _
                        "           AND ҩƷid+0 = [3]" & _
                        "           AND ���� BETWEEN [1] and [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lngҩƷid)
                    
                    If rsNum.EOF Then
                        numǰ������ = 0
                    Else
                        numǰ������ = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    GetDate int�ƻ�����, dat����, strBegin, strEnd
                    
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS �������� " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "           and ���� <>19 and ����>=15 AND b.ϵ�� = -1 " & _
                        "           AND ҩƷid+0 = [3]" & _
                        "           AND ���� BETWEEN [1] and [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lngҩƷid)
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    GetDate int�ƻ�����, datǰ��, strBegin, strEnd
                    
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS ǰ������ " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "           AND b.ϵ�� = -1 " & _
                        "           and �ⷿid+0=[4]" & _
                        "           AND ҩƷid+0= [3] " & _
                        "           AND ���� BETWEEN [1] and [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lngҩƷid, lng�ⷿID)
    
                    If rsNum.EOF Then
                        numǰ������ = 0
                    Else
                        numǰ������ = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    
                    rsNum.Close
                    
                    GetDate int�ƻ�����, dat����, strBegin, strEnd
                    
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS �������� " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "       AND b.ϵ�� = -1 " & _
                        "       and �ⷿid+0=[4]" & _
                        "       AND ҩƷid+0= [3]" & _
                        "       AND ���� BETWEEN [1]  and  [2]"
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lngҩƷid, lng�ⷿID)
    
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num�������� = num�������� / .TextMatrix(intCurrentRow, mHeadCol.����ϵ��)
                numǰ������ = numǰ������ / .TextMatrix(intCurrentRow, mHeadCol.����ϵ��)
                '�ƻ�����=2������������ǰ���������������
                If mbln�ƻ����� = True Then
                    num�ƻ����� = 2 * num�������� - numǰ������ - num�������
                    If num�ƻ����� < 0 Then num�ƻ����� = 0
                End If
    
                .TextMatrix(intCurrentRow, mHeadCol.ǰ������) = Format(numǰ������, mFMT.FM_����)
                .TextMatrix(intCurrentRow, mHeadCol.��������) = Format(num��������, mFMT.FM_����)
                .TextMatrix(intCurrentRow, mHeadCol.�ƻ�����) = IIf(Format(num�ƻ�����, mFMT.FM_����) = 0, "", Format(num�ƻ�����, mFMT.FM_����))
                .TextMatrix(intCurrentRow, mHeadCol.���) = IIf(Format(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.����) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.����)), mFMT.FM_����) = 0, "", Format(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.����) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.����)), mFMT.FM_����))
            Case 2      '�ٽ��ڼ�ƽ�����շ�
                datǰ�� = Choose(int�ƻ�����, DateAdd("m", -2, zlDatabase.Currentdate), DateAdd("m", -6, zlDatabase.Currentdate), DateAdd("yyyy", -2, zlDatabase.Currentdate), DateAdd("d", -14, sys.Currentdate))
                dat���� = Choose(int�ƻ�����, DateAdd("m", -1, zlDatabase.Currentdate), DateAdd("m", -3, zlDatabase.Currentdate), DateAdd("yyyy", -1, zlDatabase.Currentdate), DateAdd("d", -7, sys.Currentdate))
    
                If lng�ⷿID = 0 Then
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS ǰ������ " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "           and ���� <>19 and ����>=15 AND b.ϵ�� = -1 " & _
                        "           AND ҩƷid+0= [3]" & _
                        "           AND ���� BETWEEN [1] and [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(Format(DateAdd(Choose(int�ƻ�����, "m", "m", "m", "d"), Choose(int�ƻ�����, -1, -3, -12, -7), datǰ��), "yyyy-mm-dd hh:mm:ss")), CDate(Format(datǰ��, "yyyy-mm-dd hh:mm:ss")), lngҩƷid)
                    
                    If rsNum.EOF Then
                        numǰ������ = 0
                    Else
                        numǰ������ = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS �������� " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "       and ���� <>19 and ����>=15 AND b.ϵ�� = -1 " & _
                        "       AND ҩƷid+0= [3]" & _
                        "       AND ���� BETWEEN [1] " & _
                        "       and [2]"
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(Format(DateAdd(Choose(int�ƻ�����, "m", "m", "m", "d"), Choose(int�ƻ�����, -1, -3, -12, -7), dat����), "yyyy-mm-dd hh:mm:ss")), CDate(Format(dat����, "yyyy-mm-dd hh:mm:ss")), lngҩƷid)
    
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS ǰ������ " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "       AND b.ϵ�� = -1 " & _
                        "       and a.�ⷿid+0=[4]" & _
                        "       AND ҩƷid+0=[3] " & _
                        "       AND ���� BETWEEN [1] and  [2] "
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(Format(DateAdd(Choose(int�ƻ�����, "m", "m", "m", "d"), Choose(int�ƻ�����, -1, -3, -12, -7), datǰ��), "yyyy-mm-dd hh:mm;ss")), CDate(Format(datǰ��, "yyyy-mm-dd hh:mm:ss")), lngҩƷid, lng�ⷿID)
    
                    If rsNum.EOF Then
                        numǰ������ = 0
                    Else
                        numǰ������ = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS �������� " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "           AND b.ϵ�� = -1 " & _
                        "           and a.�ⷿid+0=[4]" & _
                        "           AND ҩƷid+0=[3] " & _
                        "           AND ���� BETWEEN [1]  and [2]"
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(Format(DateAdd(Choose(int�ƻ�����, "m", "m", "m", "d"), Choose(int�ƻ�����, -1, -3, -12, -7), dat����), "yyyy-mm-dd hh:mm;ss")), CDate(Format(dat����, "yyyy-mm-dd hh:mm:ss")), lngҩƷid, lng�ⷿID)
    
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num�������� = num�������� / .TextMatrix(intCurrentRow, mHeadCol.����ϵ��)
                numǰ������ = numǰ������ / .TextMatrix(intCurrentRow, mHeadCol.����ϵ��)
                
                '�ƻ����� = (ǰ������ + ��������) / 2 - �������
                If mbln�ƻ����� = True Then
                    num�ƻ����� = (num�������� + numǰ������) / 2 - num�������
                    If num�ƻ����� < 0 Then num�ƻ����� = 0
                End If
                .TextMatrix(intCurrentRow, mHeadCol.ǰ������) = Format(numǰ������, mFMT.FM_����)
                .TextMatrix(intCurrentRow, mHeadCol.��������) = Format(num��������, mFMT.FM_����)
                .TextMatrix(intCurrentRow, mHeadCol.�ƻ�����) = IIf(Format(num�ƻ�����, mFMT.FM_����) = 0, "", Format(num�ƻ�����, mFMT.FM_����))
                .TextMatrix(intCurrentRow, mHeadCol.���) = IIf(Format(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.����) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.����)), mFMT.FM_����) = 0, "", Format(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.����) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.����)), mFMT.FM_����))
    
            Case 3      'ҩƷ����������շ�
                If lng�ⷿID = 0 Then
                    gstrSQL = "select sum(����) as  ���� from ���ϴ����޶�  where ����id=[1]"
                Else
                    gstrSQL = "select ���� from ���ϴ����޶�  where ����id=[1] and �ⷿid=[2]"
    
                End If
                Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngҩƷid, lng�ⷿID)
    
                If rsNum.EOF Then
                    num���� = 0
                Else
                    num���� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num���� = num���� / .TextMatrix(intCurrentRow, mHeadCol.����ϵ��)
                '�ƻ�����=�������ޣ��������
                If mbln�ƻ����� = True Then
                    num�ƻ����� = IIf(num���� > num�������, num���� - num�������, 0)
                End If
                .TextMatrix(intCurrentRow, mHeadCol.�ƻ�����) = IIf(Format(num�ƻ�����, mFMT.FM_����) = 0, "", Format(num�ƻ�����, mFMT.FM_����))
                .TextMatrix(intCurrentRow, mHeadCol.���) = IIf(Format(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.����) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.����)), mFMT.FM_����) = 0, "", Format(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.����) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.����)), mFMT.FM_����))
            Case 4  '��������
                datǰ�� = Choose(int�ƻ�����, DateAdd("m", -2, zlDatabase.Currentdate), DateAdd("m", -6, zlDatabase.Currentdate), DateAdd("yyyy", -2, zlDatabase.Currentdate), DateAdd("d", -14, sys.Currentdate))
                dat���� = Choose(int�ƻ�����, DateAdd("m", -1, zlDatabase.Currentdate), DateAdd("m", -3, zlDatabase.Currentdate), DateAdd("yyyy", -1, zlDatabase.Currentdate), DateAdd("d", -7, sys.Currentdate))
                GetDate int�ƻ�����, dat����, strBegin, strEnd
                lng���� = CDate(Format(strEnd, "yyyy-MM-DD")) - CDate(Format(strBegin, "yyyy-MM-DD")) + 1
                If lng���� <= 0 Then lng���� = 1
                
                If lng�ⷿID = 0 Then
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS �������� " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "           and ���� <>19 and ����>=15 AND b.ϵ�� = -1 " & _
                        "           AND ҩƷid+0 =[3] " & _
                        "           AND ���� BETWEEN [1]   and  [2] "
                
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lngҩƷid)
                            
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = "" & _
                        "   SELECT ABS(SUM(NVL(����, 0))) AS �������� " & _
                        "   FROM ҩƷ�շ����� a, ҩƷ������ b " & _
                        "   Where a.���id = b.id " & _
                        "       AND b.ϵ�� = -1 " & _
                        "       and �ⷿid+0=[4]" & _
                        "       AND ҩƷid+0=[3] " & _
                        "       AND ���� BETWEEN [1] and [2]"
                    
                    Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, CDate(strBegin), CDate(strEnd), lngҩƷid, lng�ⷿID)
    
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num�������� = num�������� / .TextMatrix(intCurrentRow, mHeadCol.����ϵ��)
                num���� = num�������� / lng���� * mint����
                num���� = num�������� / lng���� * mint����
                '�ƻ�����=2������������ǰ���������������
                If mbln�ƻ����� = True Then
                    If num������� < num���� Then
                        num�ƻ����� = num���� - num�������
                    Else
                        num�ƻ����� = 0
                    End If
                    If num�ƻ����� < 0 Then num�ƻ����� = 0
                End If
    
                .TextMatrix(intCurrentRow, mHeadCol.ǰ������) = Format(numǰ������, mFMT.FM_����)
                .TextMatrix(intCurrentRow, mHeadCol.��������) = Format(num��������, mFMT.FM_����)
                .TextMatrix(intCurrentRow, mHeadCol.�ƻ�����) = IIf(Format(num�ƻ�����, mFMT.FM_����) = 0, "", Format(num�ƻ�����, mFMT.FM_����))
                .TextMatrix(intCurrentRow, mHeadCol.���) = IIf(Format(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.����) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.����)), mFMT.FM_����) = 0, "", Format(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mHeadCol.����) = "", 0, mshBill.TextMatrix(intCurrentRow, mHeadCol.����)), mFMT.FM_����))
        End Select
    
        Call Calc����(lngҩƷid, intCurrentRow)
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Calc����(ByVal lngҩƷid As Long, ByVal intCurrentRow As Integer)
    '�ֱ�������ںͱ��ڵ�������
    'ȡ���ڵ����䷶Χ
    Dim strBegin As String
    Dim strEnd As String
    Dim rsNum As ADODB.Recordset
    
    On Error GoTo errHandle
    With mshBill
        Select Case mint�ƻ�����
            '1:�¶ȼƻ�,2.���ȼƻ�,3.��ȼƻ�
            Case 1
                '����ʱ�䷶Χ
                strBegin = Format(DateAdd("m", -1, CDate(mstrNow)), "YYYY-MM") & "-01"
                strEnd = Format(DateAdd("d", -1, CDate(Format(CDate(mstrNow), "YYYY-MM") & "-01")), "YYYY-MM-DD") & " 23:59:59"
            Case 2
                '�ϼ���ʱ�䷶Χ
                Select Case DatePart("Q", CDate(mstrNow))
                    Case 1
                        strBegin = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-10-01"
                        strEnd = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-12-31 23:59:59"
                    Case 2
                        strBegin = Format(mstrNow, "YYYY") & "-01-01"
                        strEnd = Format(mstrNow, "YYYY") & "-03-31 23:59:59"
                     Case 3
                        strBegin = Format(mstrNow, "YYYY") & "-04-01"
                        strEnd = Format(mstrNow, "YYYY") & "-06-30 23:59:59"
                    Case 4
                        strBegin = Format(mstrNow, "YYYY") & "-07-01"
                        strEnd = Format(mstrNow, "YYYY") & "-09-30 23:59:59"
                End Select
            Case 3
                '�����ʱ�䷶Χ
                strBegin = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-01-01"
                strEnd = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-12-31 23:59:59"
            Case 4
                '����ʱ�䷶Χ
                strBegin = Format(DateAdd("d", 2 - Weekday(CDate(mstrNow)) - 7, CDate(mstrNow)), "YYYY-mm-dd")
                strEnd = Format(DateAdd("d", 8 - Weekday(CDate(mstrNow)) - 7, CDate(mstrNow)), "YYYY-mm-dd") & " 23:59:59"
                
        End Select
            
        '������������������Ҫ��ȷֵ����ҩƷ�շ�����ͳ�ƣ�
        gstrSQL = "Select -Sum(Nvl(����, 0)) As �������� " & _
            " From ҩƷ�շ�����" & _
            " Where ���id + 0 In (19,20,21) And ҩƷid+0=[1] And ���� Between [2] And [3] "
        Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngҩƷid, CDate(strBegin), CDate(strEnd))
        If rsNum.RecordCount > 0 Then
            .TextMatrix(intCurrentRow, mHeadCol.��������) = Format(NVL(rsNum!��������, 0) / Val(.TextMatrix(intCurrentRow, mHeadCol.����ϵ��)), mFMT.FM_����)
        End If
        
        'ȡ���ڵ����䷶Χ
        Select Case mint�ƻ�����
            '1:�¶ȼƻ�,2.���ȼƻ�,3.��ȼƻ�
            Case 1
                '����ʱ�䷶Χ
                strBegin = Format(mstrNow, "YYYY-MM") & "-01"
            Case 2
                '������ʱ�䷶Χ
                Select Case DatePart("Q", CDate(mstrNow))
                    Case 1
                        strBegin = Format(mstrNow, "YYYY") & "-01-01"
                    Case 2
                        strBegin = Format(mstrNow, "YYYY") & "-04-01"
                    Case 3
                        strBegin = Format(mstrNow, "YYYY") & "-07-01"
                    Case 4
                        strBegin = Format(mstrNow, "YYYY") & "-10-01"
                End Select
            Case 3
                '�����ʱ�䷶Χ
                strBegin = Format(mstrNow, "YYYY") & "-01-01"
            Case 4
                '����ʱ�䷶Χ
                strBegin = Format(DateAdd("d", 2 - Weekday(CDate(mstrNow)), CDate(mstrNow)), "YYYY-mm-dd")
        End Select
        
        '���ڽ���ʱ���ֹ������
        strEnd = Format(mstrNow, "YYYY-MM-DD") & " 23:59:59"
            
        '���㱾������������Ҫ��ȷֵ����ҩƷ�շ�����ͳ�ƣ�
        gstrSQL = "Select -Sum(Nvl(����, 0)) As �������� " & _
            " From ҩƷ�շ�����" & _
            " Where ���id + 0 In (19,20,21) And ҩƷid+0=[1] And ���� Between [2] And [3] "
        Set rsNum = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngҩƷid, CDate(strBegin), CDate(strEnd))
        If rsNum.RecordCount > 0 Then
            .TextMatrix(intCurrentRow, mHeadCol.��������) = Format(NVL(rsNum!��������, 0) / Val(.TextMatrix(intCurrentRow, mHeadCol.����ϵ��)), mFMT.FM_����)
        End If
    End With
        
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowPercent(sngPercent As Single)
'����:��״̬���ϸ��ݰٷֱ���ʾ��ǰ��������(��)
    Dim intAll As Integer
    intAll = stbThis.Panels(2).Width / TextWidth("��") - 4
    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "��")
End Sub

Private Sub Form_Load()
    Dim strReg As String
    mFMT.FM_��� = GetDigit
    mblnSelectStock = IIf(Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModule, "0")) = 1, 1, 0)
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mstrNow = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    
    mintУ�鷽ʽ = Val(Mid(zlDatabase.GetPara("����У��", glngSys, mlngModule, ""), 1, 1))
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
     
'    mintUnit = GetUnit()
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Call initCard
    
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(mHeadCol.����) = IIf(mblnCostView = True, 1000, 0)
        .ColWidth(mHeadCol.���) = IIf(mblnCostView = True, 1000, 0)
    End With
    
    mshBill.ColWidth(mHeadCol.У��) = IIf(mint�༭״̬ = 3 And mintУ�鷽ʽ = 1, 500, 0)
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim intRecordCount As Integer
    Dim str��λ As String
    Dim strOrder As String, strCompare As String
    
    On Error GoTo errHandle
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strCompare = Mid(strOrder, 1, 1)

    '�ⷿ
    Select Case mint�༭״̬
        Case 1
            Txt������ = gstrUserName
            Txt�������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
            strUnit = "��װ��λ"
            Select Case mintUnit
            Case 0
                str��λ = ",j.���㵥λ ��λ,1 ����ϵ��"
            Case Else
                str��λ = ",m.��װ��λ ��λ,m.����ϵ�� ����ϵ��"
            End Select
            
            initGrid
            
            gstrSQL = "Select �ⷿid From  ���ϲɹ��ƻ� where nvl(����,0)=0 and NO=[1] and rownum=1 "
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            mlng�ⷿid = Val(NVL(rsInitCard!�ⷿID))
            
            gstrSQL = "" & _
                "   SELECT a.id,nvl(a.�ⷿid,0) as �ⷿid,nvl(c.����,'ȫԺ') AS �ⷿ,a.no, a.�ƻ�����,a.�ڼ�, a.���Ʒ���, a.������," & _
                "           TO_CHAR (a.��������, 'yyyy-mm-dd HH24:MI:SS') AS ��������, a.�����," & _
                "           TO_CHAR (a.�������, 'yyyy-mm-dd HH24:MI:SS') AS �������,a.����˵��," & _
                "           b.���,b.����id ҩƷid,m.�б���� ,F.����,F.����,J.����,J.���� ͨ������, J.���" & str��λ & ", b.ǰ������, b.��������,b.��������,b.��������, b.�������, b.�ƻ�����, b.����, b.���, b.�ϴι�Ӧ��,b.�ϴ������� " & _
                "   FROM ���ϲɹ��ƻ� a, ���ϼƻ����� b,���ű� c,�������� M,�շ���ĿĿ¼ J," & _
                "       ( Select ����id ,sum(nvl(����,0)) ����,sum(nvl(����,0)) ���� " & _
                "               From ���ϴ����޶�  " & _
                "               " & IIf(mlng�ⷿid = 0, "", " Where �ⷿID=[2]") & _
                "               Group By ����ID ) F " & _
                "   Where a.id = b.�ƻ�id and b.����ID=f.����id(+) and nvl(a.�ⷿid,0)=c.id(+) " & _
                "          and b.����id=m.����id and m.����id=J.id And (j.վ��=[3] or j.վ�� is null) And nvl(a.����,0)=0 AND a.no = [1]" & _
                "   Order by " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "����", "ͨ������")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, mlng�ⷿid, gstrNodeNo)
                '"       (   SELECT DISTINCT a.����id as ҩƷid,c.����,C.����  AS ͨ������,c.���,c.���㵥λ as ɢװ��λ,A.��װ��λ,a.����ϵ�� " & _
                "           FROM �������� a, �շ���Ŀ���� b, �շ���ĿĿ¼ c " & _
                "           WHERE a.����id = b.�շ�ϸĿID(+) and B.����(+)=3  AND a.����id = c.ID" & _
                "        ) d
                
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If

            intRecordCount = rsInitCard.RecordCount

            Txt������ = rsInitCard!������
            If mint�༭״̬ = 2 Then
                Txt������ = gstrUserName
            End If
            Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")

            Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
            Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsInitCard!����˵��), "", rsInitCard!����˵��)
            txt�ƻ����� = Choose(rsInitCard!�ƻ�����, "�¶ȼƻ�", "���ȼƻ�", "��ȼƻ�", "�ܶȼƻ�")
            txt���Ʒ��� = Choose(rsInitCard!���Ʒ���, "����ͬ�����β��շ�", "�ٽ��ڼ�ƽ�����շ�", "���ϴ���������շ�", "���������������շ�", "�����깺���շ�")
            mint�ƻ����� = rsInitCard!�ƻ�����
            mint���Ʒ��� = rsInitCard!���Ʒ���
            mlng�ⷿid = rsInitCard!�ⷿID
            mlng�ƻ�ID = rsInitCard!Id

            mstr�ڼ� = rsInitCard!�ڼ�
            Select Case mint�ƻ�����
                Case 1       '�¼ƻ�
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr�ڼ�, 1, 4) & "��" & Right(mstr�ڼ�, 2) & "��" & ") " & rsInitCard!�ⷿ & "�ɹ��ƻ�"
                Case 2       '���ƻ�
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr�ڼ�, 1, 4) & "��" & Right(mstr�ڼ�, 1) & "��" & ")" & rsInitCard!�ⷿ & "�ɹ��ƻ�"
                Case 3    '��ƻ�
                    LblTitle.Caption = GetUnitName & "(" & mstr�ڼ� & "��" & ")" & rsInitCard!�ⷿ & "�ɹ��ƻ�"
                Case 4       '�ܼƻ�
                    LblTitle.Caption = GetUnitName & "(" & Mid(mstr�ڼ�, 1, 4) & "��" & Right(mstr�ڼ�, 2) & "��" & ") " & rsInitCard!�ⷿ & "�ɹ��ƻ�"
            End Select

            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If

            With mshBill
'                Select Case mint�ƻ�����
'                    Case 1
'                        .TextMatrix(0, mHeadCol.��������) = "��������"
'                        .TextMatrix(0, mHeadCol.��������) = "��������"
'                    Case 2
'                        .TextMatrix(0, mHeadCol.��������) = "�ϼ�������"
'                        .TextMatrix(0, mHeadCol.��������) = "����������"
'                    Case Else
'                        .TextMatrix(0, mHeadCol.��������) = "��������"
'                        .TextMatrix(0, mHeadCol.��������) = "��������"
'                End Select
                For intRow = 1 To intRecordCount

                    .TextMatrix(intRow, 0) = rsInitCard!ҩƷID
                    .TextMatrix(intRow, mHeadCol.����) = "[" & rsInitCard!���� & "]" & rsInitCard!ͨ������
                    .TextMatrix(intRow, mHeadCol.���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) = IIf(IsNull(rsInitCard!�ϴι�Ӧ��), "", rsInitCard!�ϴι�Ӧ��)
                    .TextMatrix(intRow, mHeadCol.����) = IIf(IsNull(rsInitCard!�ϴ�������), "", rsInitCard!�ϴ�������)
                    .TextMatrix(intRow, mHeadCol.��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mHeadCol.����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mHeadCol.�б����) = IIf(Val(NVL(rsInitCard!�б����)) = 1, "��", "")
                    .TextMatrix(intRow, mHeadCol.�洢����) = Format(Val(NVL(rsInitCard!����)) / rsInitCard!����ϵ��, mFMT.FM_����)
                    .TextMatrix(intRow, mHeadCol.�洢����) = Format(Val(NVL(rsInitCard!����)) / rsInitCard!����ϵ��, mFMT.FM_����)
                    .TextMatrix(intRow, mHeadCol.ǰ������) = Format(Val(NVL(rsInitCard!ǰ������)) / rsInitCard!����ϵ��, mFMT.FM_����)
                    .TextMatrix(intRow, mHeadCol.��������) = Format(Val(NVL(rsInitCard!��������)) / rsInitCard!����ϵ��, mFMT.FM_����)
                    .TextMatrix(intRow, mHeadCol.�������) = Format(Val(NVL(rsInitCard!�������)) / rsInitCard!����ϵ��, mFMT.FM_����)
                    
                    .TextMatrix(intRow, mHeadCol.��������) = Format(Val(NVL(rsInitCard!��������)) / rsInitCard!����ϵ��, mFMT.FM_����)
                    .TextMatrix(intRow, mHeadCol.��������) = Format(Val(NVL(rsInitCard!��������)) / rsInitCard!����ϵ��, mFMT.FM_����)
                    
                    .TextMatrix(intRow, mHeadCol.�ƻ�����) = IIf(Format(Val(NVL(rsInitCard!�ƻ�����)), mFMT.FM_����) = 0, "", Format(rsInitCard!�ƻ����� / rsInitCard!����ϵ��, mFMT.FM_����))
                    .TextMatrix(intRow, mHeadCol.����) = Format(Val(NVL(rsInitCard!����)) * rsInitCard!����ϵ��, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mHeadCol.���) = IIf(Format(Val(NVL(rsInitCard!���)), mFMT.FM_���) = 0, "", Format(Val(NVL(rsInitCard!���)), mFMT.FM_���))
                    If intRow = .Rows - 1 Then .Rows = .Rows + 1
                    rsInitCard.MoveNext
                Next
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mHeadCol.���, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ʼ���༭�ؼ�
Private Sub initGrid()
    Dim intCol As Integer

    With mshBill
        .Active = True
        .Cols = mconIntColS
        .MsfObj.FixedCols = 2

        .TextMatrix(0, mHeadCol.���) = "���"
        .TextMatrix(0, mHeadCol.У��) = "У��"
        .TextMatrix(0, mHeadCol.����) = "�������������"
        .TextMatrix(0, mHeadCol.���) = "���"
        .TextMatrix(0, mHeadCol.����) = "����"
        .TextMatrix(0, mHeadCol.��λ) = "��λ"
        .TextMatrix(0, mHeadCol.����ϵ��) = "����ϵ��"
        .TextMatrix(0, mHeadCol.�б����) = "�б����"
        .TextMatrix(0, mHeadCol.�洢����) = "�洢����"
        .TextMatrix(0, mHeadCol.�洢����) = "�洢����"

        .TextMatrix(0, mHeadCol.ǰ������) = "ǰ������"
        .TextMatrix(0, mHeadCol.��������) = "��������"
        .TextMatrix(0, mHeadCol.�������) = "�������"
        .TextMatrix(0, mHeadCol.��������) = "��������"
        .TextMatrix(0, mHeadCol.��������) = "��������"
        
        .TextMatrix(0, mHeadCol.�ƻ�����) = "�ƻ�����"
        .TextMatrix(0, mHeadCol.����) = "�ɱ���"
        .TextMatrix(0, mHeadCol.���) = "�ɱ����"
        .TextMatrix(0, mHeadCol.�ϴι�Ӧ��) = "��Ӧ��"

        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mHeadCol.���) = "1"

        .ColWidth(mHeadCol.���) = 500
        .ColWidth(mHeadCol.У��) = IIf(mint�༭״̬ = 3 And mintУ�鷽ʽ = 1, 500, 0)
        .ColWidth(mHeadCol.����) = 2000
        .ColWidth(mHeadCol.���) = 900
        .ColWidth(mHeadCol.����) = 800
        .ColWidth(mHeadCol.��λ) = 500
        .ColWidth(mHeadCol.ǰ������) = 1100
        .ColWidth(mHeadCol.��������) = 1100
        .ColWidth(mHeadCol.�������) = 1100
        .ColWidth(mHeadCol.��������) = 1100
        .ColWidth(mHeadCol.��������) = 1100
        .ColWidth(mHeadCol.�ƻ�����) = 1100
        .ColWidth(mHeadCol.�б����) = 800
        .ColWidth(mHeadCol.�洢����) = 1000
        .ColWidth(mHeadCol.�洢����) = 1000
        
        .ColWidth(mHeadCol.����) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(mHeadCol.���) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mHeadCol.�ϴι�Ӧ��) = 900
        .ColWidth(mHeadCol.����ϵ��) = 0
        .ColWidth(0) = 0

        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
        For intCol = 0 To .Cols - 1
            .ColData(intCol) = 5
        Next

        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtժҪ.Enabled = True
            .ColData(mHeadCol.����) = 1
            .ColData(mHeadCol.�ƻ�����) = 4
            .ColData(mHeadCol.����) = 4

            .ColData(mHeadCol.����) = 1
            .ColData(mHeadCol.�ϴι�Ӧ��) = 1
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
            txtժҪ.Enabled = False
            .ColData(mHeadCol.�ƻ�����) = 0
        End If
        
        .ColData(mHeadCol.У��) = IIf(mint�༭״̬ = 3 And mintУ�鷽ʽ = 1, 0, 5)
        
        .ColAlignment(mHeadCol.У��) = flexAlignCenterCenter
        .ColAlignment(mHeadCol.����) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.���) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.����) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.��λ) = flexAlignCenterCenter
        .ColAlignment(mHeadCol.ǰ������) = flexAlignRightCenter
        .ColAlignment(mHeadCol.��������) = flexAlignRightCenter
        .ColAlignment(mHeadCol.�������) = flexAlignRightCenter
        .ColAlignment(mHeadCol.��������) = flexAlignRightCenter
        .ColAlignment(mHeadCol.��������) = flexAlignRightCenter
        .ColAlignment(mHeadCol.�ƻ�����) = flexAlignRightCenter
        .ColAlignment(mHeadCol.����) = flexAlignRightCenter
        .ColAlignment(mHeadCol.���) = flexAlignRightCenter
        .ColAlignment(mHeadCol.�ϴι�Ӧ��) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.�б����) = 4
        .ColAlignment(mHeadCol.�洢����) = 7
        .ColAlignment(mHeadCol.�洢����) = 7

        .PrimaryCol = mHeadCol.����
        .LocateCol = mHeadCol.����
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mHeadCol.����) = 0
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub

    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With

    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With


    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With

    txt���Ʒ���.Left = mshBill.Left + mshBill.Width - txt���Ʒ���.Width
    lbl���Ʒ���.Left = txt���Ʒ���.Left - lbl���Ʒ���.Width - 100


    Lbl�ƻ�����.Left = mshBill.Left

    txt�ƻ�����.Left = Lbl�ƻ�����.Left + Lbl�ƻ�����.Width + 100

    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With

    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With

    With Lbl��������
        .Top = Lbl������.Top
        .Left = Txt������.Left + Txt������.Width + 250
    End With

    With Txt��������
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With

    With Txt�������
        .Top = Lbl������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With

    With Lbl�������
        .Top = Lbl������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With

    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = Lbl�������.Left - 200 - .Width
    End With

    With Lbl�����
        .Top = Lbl������.Top
        .Left = Txt�����.Left - 100 - .Width
    End With

    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With

    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With

    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If

    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With

    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With

    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With

    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With

    With cmdFind
        .Top = CmdCancel.Top
    End With

    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mblnFirstCheck = False
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If

End Sub

Private Function SaveCheck() As Boolean
    Dim str����� As String

    mblnSave = False
    SaveCheck = False

    str����� = gstrUserName

    On Error GoTo errHandle
    'zl_���ϼƻ�����_VERIFY( /*ID_IN*/, /*�����_IN*/ );
    gstrSQL = "zl_���ϼƻ�����_VERIFY('" & mlng�ƻ�ID & "','" & str����� & "')"
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption

    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    'MsgBox "���ʧ�ܣ�", vbInformation, gstrSysName
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub Msf��Ӧ��ѡ��_DblClick()
    Dim blnCancel As Boolean
    With mshBill
        .Text = Msf��Ӧ��ѡ��.TextMatrix(Msf��Ӧ��ѡ��.Row, 2)
        .TextMatrix(.Row, mHeadCol.�ϴι�Ӧ��) = Msf��Ӧ��ѡ��.TextMatrix(Msf��Ӧ��ѡ��.Row, 2)
    End With
    Msf��Ӧ��ѡ��.Visible = False
    mshBill.SetFocus
    Call SendKeys("{ENTER}")
End Sub

Private Sub Msf��Ӧ��ѡ��_GotFocus()
    If Msf��Ӧ��ѡ��.Rows - 1 = 1 Then Call Msf��Ӧ��ѡ��_DblClick
End Sub

Private Sub Msf��Ӧ��ѡ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Msf��Ӧ��ѡ��_DblClick
    End If
End Sub

Private Sub Msf��Ӧ��ѡ��_LostFocus()
    Msf��Ӧ��ѡ��.ZOrder 1
    Msf��Ӧ��ѡ��.Visible = False
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mHeadCol.���, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mHeadCol.���, mshBill.Row)
    Call ��ʾ�ϼƽ��
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mHeadCol.����) = 0 Then
        Exit Sub
    End If
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ����������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim sngLeft As Single, sngTop As Single
    Dim RecReturn As Recordset
    Dim strUnit As String
    
    On Error GoTo errHandle
    If mshBill.Col = mHeadCol.���� Then
        Set RecReturn = Frm����ѡ����.ShowMe(Me, 1, , mlng�ⷿid, , , , , , , , , , , , , , mstrPrivs)
        If RecReturn.RecordCount > 0 Then
            If RecReturn.RecordCount = 1 Then
                If mintUnit = 0 Then
                    strUnit = "ɢװ��λ"
                Else
                    strUnit = "��װ��λ"
                End If
                SetStuffRows RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                            IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                            Switch(strUnit = "ɢװ��λ", RecReturn!ɢװ��λ, strUnit = "��װ��λ", RecReturn!��װ��λ), RecReturn!ָ��������, _
                            Switch(strUnit = "ɢװ��λ", 1, strUnit = "��װ��λ", RecReturn!����ϵ��)
            End If
            RecReturn.Close
        End If
    ElseIf mshBill.Col = mHeadCol.�ϴι�Ӧ�� Then
        'ҩƷ��Ӧ�̵�ѡ��
        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
        If sngLeft + Msf��Ӧ��ѡ��.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf��Ӧ��ѡ��.Width - 100

        Set RecReturn = New ADODB.Recordset
        gstrSQL = "Select ID,����,����,���� From ��Ӧ�� " & _
                  "Where ĩ��=1 And (վ��=[1] or վ�� is null) And (substr(����,5,1)=1  Or Nvl(ĩ��,0)=0) " & _
                  "  And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) Order By ���� "
        Set RecReturn = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "-��Ӧ��", gstrNodeNo)
        If RecReturn.RecordCount = 0 Then
            MsgBox "���ȳ�ʼ���������Ϲ�Ӧ�̣�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        With Msf��Ӧ��ѡ��
            .Clear
            Set .DataSource = RecReturn
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800

            .Row = 1
            .ColSel = .Cols - 1
        End With
        With Msf��Ӧ��ѡ��
            .Left = sngLeft
            .Top = sngTop
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    ElseIf mshBill.Col = mHeadCol.���� Then
        '�����̵�ѡ��
        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
        If sngLeft + msh������.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - msh������.Width - 100

        Set RecReturn = New ADODB.Recordset
        gstrSQL = "Select ����,����,����,������ҵ����֤,������ҵ����֤Ч�� From ���������� Order By ���� "
        zlDatabase.OpenRecordset RecReturn, gstrSQL, "��ȡ����������"
        If RecReturn.RecordCount = 0 Then
            MsgBox "���ȳ�ʼ���������Ϲ�Ӧ�̣�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        With msh������
            .Clear
            Set .DataSource = RecReturn
            .ColWidth(0) = 800
            .ColWidth(1) = 2000
            .ColWidth(2) = 800
            .ColWidth(3) = 1000
            .ColWidth(4) = 1000

            .Row = 1
            .ColSel = .Cols - 1
        End With
        With msh������
            .Left = sngLeft
            .Top = sngTop
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_DblClick(Cancel As Boolean)
    Dim blnAllowChange As Boolean
    Dim lngColor As Long
    Dim i As Integer
    
    With mshBill
        If mblnFirstCheck = False Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .Col <> mHeadCol.У�� Then Exit Sub
        If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
        
        i = .ColData(mHeadCol.����)
        .ColData(mHeadCol.����) = 0
        .Col = mHeadCol.����
        lngColor = .MsfObj.CellForeColor
        If lngColor = vbRed Then
            blnAllowChange = True
        End If
        .ColData(mHeadCol.����) = i
        
        i = .ColData(mHeadCol.����)
        .ColData(mHeadCol.����) = 0
        .Col = mHeadCol.����
        lngColor = .MsfObj.CellForeColor
        If lngColor = vbRed Then
            blnAllowChange = True
        End If
        .ColData(mHeadCol.����) = i
        
        i = .ColData(mHeadCol.�ϴι�Ӧ��)
        .ColData(mHeadCol.�ϴι�Ӧ��) = 0
        .Col = mHeadCol.�ϴι�Ӧ��
        lngColor = .MsfObj.CellForeColor
        If lngColor = vbRed Then
           blnAllowChange = True
        End If
        .ColData(mHeadCol.�ϴι�Ӧ��) = i
        
        .Col = mHeadCol.У��
        If blnAllowChange = True Then
            If .TextMatrix(.Row, .Col) = "��" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "��"
            End If
        End If
    End With
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mHeadCol.�ƻ����� Or .Col = mHeadCol.���� Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mHeadCol.�ƻ�����
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.����С��, g_С��λ��.obj_ɢװС��.����С��)
                Case mHeadCol.����
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.���ۼ�С��, g_С��λ��.obj_ɢװС��.���ۼ�С��)
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                KeyAscii = 0
                Exit Sub
            End If
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If

        Select Case .Col
            Case mHeadCol.����
                .TxtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
            Case mHeadCol.����
                .TxtCheck = False
                .MaxLength = 40
            Case mHeadCol.�ϴι�Ӧ��
                .MaxLength = 40
                .TxtCheck = False
            Case mHeadCol.�ƻ�����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mHeadCol.����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"

        End Select

    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsStuff As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    
    Dim rsTemp As Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        If .Col = mHeadCol.���� Then
            .Text = UCase(Trim(.Text))
        Else
            .Text = Trim(.Text)
        End If
        strKey = .Text

        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col

            Case mHeadCol.����
                If strKey <> "" Then

                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If

                    Set rsTemp = FrmMulitSel.ShowSelect(Me, 1, , mlng�ⷿid, , strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, , , , , , , , , , , , mstrPrivs)

                    If rsTemp.RecordCount = 1 Then
                        If mintUnit = 0 Then
                            strUnit = "ɢװ��λ"
                        Else
                            strUnit = "��װ��λ"
                        End If
                        If SetStuffRows(rsTemp!����ID, "[" & rsTemp!���� & "]" & rsTemp!����, _
                            IIf(IsNull(rsTemp!���), "", rsTemp!���), IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                            Switch(strUnit = "ɢװ��λ", rsTemp!ɢװ��λ, strUnit = "��װ��λ", rsTemp!��װ��λ), rsTemp!ָ��������, _
                            Switch(strUnit = "ɢװ��λ", 1, strUnit = "��װ��λ", rsTemp!����ϵ��)) = False Then
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = .TextMatrix(.Row, .Col)
                    Else
                       
                        Cancel = True
                    End If
                End If
            Case mHeadCol.�ƻ�����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬼ƻ���������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "����������(0~99999999)��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
'                    If .TxtVisible = True Then
'                        .TextMatrix(.Row, mHeadCol.�ƻ�����) = ""
'                    End If
'                    .Col = mHeadCol.����
'                    If .Row < .Rows - 1 Then
'                        .Row = .Row + 1
'                    Else
'                        If .TextMatrix(.Row, 0) <> "" Then
'                            .Rows = .Rows + 1
'                            .Row = .Row + 1
'                        End If
'                    End If
                    Cancel = True

                    Exit Sub
                End If


                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    If .TextMatrix(.Row, mHeadCol.����) <> "" Then
                        .TextMatrix(.Row, mHeadCol.���) = Format(.TextMatrix(.Row, mHeadCol.����) * strKey, mFMT.FM_���)
                    End If

                End If
                Call ��ʾ�ϼƽ��
            Case mHeadCol.����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "���۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "���۱�����(0~99999999)��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mHeadCol.����) = " "
                        .Text = " "
                    End If
'                    .Col = mHeadCol.����
'                    If .Row < .Rows - 1 Then
'                        .Row = .Row + 1
'                    Else
'                        If .TextMatrix(.Row, 0) <> "" Then
'                            .Rows = .Rows + 1
'                            .Row = .Row + 1
'                        End If
'                    End If
'                    .TextMatrix(.Row, mHeadCol.���) = format(Val(.TextMatrix(.Row, mHeadCol.����)) * Val(.TextMatrix(.Row, mHeadCol.�ƻ�����)), mFMT.FM_�ɱ���)
                                 
'                    Cancel = True
'                    Exit Sub
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.����) = strKey
                End If
                .TextMatrix(.Row, mHeadCol.���) = Format(Val(.TextMatrix(.Row, mHeadCol.����)) * Val(.TextMatrix(.Row, mHeadCol.�ƻ�����)), mFMT.FM_���)
                Call ��ʾ�ϼƽ��
                
            Case mHeadCol.�ϴι�Ӧ��
                If .TxtVisible = False Then Exit Sub
                If strKey = "" And .TextMatrix(.Row, mHeadCol.�ϴι�Ӧ��) = "" Then
                    strKey = " "
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.�ϴι�Ӧ��) = strKey
                Else
                    If StrIsValid(strKey, 40) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = UCase(strKey)
                    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
                    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngLeft + Msf��Ӧ��ѡ��.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf��Ӧ��ѡ��.Width - 100
            
                    Set rsTemp = New ADODB.Recordset
                    gstrSQL = "" & _
                        "   Select ID,����,����,���� " & _
                        "   From ��Ӧ�� " & _
                        "   Where ĩ��=1 And (վ��=[2] or վ�� is null) And (substr(����,5,1)=1 Or Nvl(ĩ��,0)=0) " & _
                        "           And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                        "           And (upper(����) Like [1] Or Upper(����) Like [1] Or Upper(����) Like [1])" & _
                        "   Order By ���� "
                    
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Ϲ�Ӧ��", strKey & "%", gstrNodeNo)
                    
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "û���ҵ����������Ĺ�Ӧ�̣�", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    ElseIf rsTemp.RecordCount = 1 Then
                        .Text = rsTemp!����
                        Exit Sub
                    End If
                    
                    With Msf��Ӧ��ѡ��
                        .Clear
                        Set .DataSource = rsTemp
                        .ColWidth(0) = 0
                        .ColWidth(1) = 800
                        .ColWidth(2) = 3000
                        .ColWidth(3) = 800
            
                        .Row = 1
                        .ColSel = .Cols - 1
                    End With
                    With Msf��Ӧ��ѡ��
                        .Left = sngLeft
                        .Top = sngTop
                        .Visible = True
                        .ZOrder 0
                        .SetFocus
                    End With
                    Cancel = True
                End If
            Case mHeadCol.����
'                If .TxtVisible = False Then Exit Sub
                If strKey = "" And .TextMatrix(.Row, mHeadCol.����) = "" Then
                    strKey = " "
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.����) = strKey
                Else
                    If strKey <> "" Then
                        If StrIsValid(strKey, 40) = False Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        strKey = UCase(strKey)
                        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
                        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                        If sngLeft + msh������.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - msh������.Width - 100
                
                        Set rsTemp = New ADODB.Recordset
                        gstrSQL = "" & _
                            "   Select ����,����,����,������ҵ����֤,������ҵ����֤Ч�� " & _
                            "   From ���������� " & _
                            "   Where (upper(����) Like [1] Or Upper(����) Like [1] Or Upper(����) Like [1])" & _
                            "   Order By ���� "
                        
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����������", strKey & "%")
                        
                        If rsTemp.RecordCount = 0 Then
                            MsgBox "û���ҵ����������������̣�", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        ElseIf rsTemp.RecordCount = 1 Then
                            .Text = rsTemp!����
                            Exit Sub
                        End If
                        
                        With msh������
                            .Clear
                            Set .DataSource = rsTemp
                            .ColWidth(0) = 800
                            .ColWidth(1) = 2000
                            .ColWidth(2) = 800
                            .ColWidth(3) = 1000
                            .ColWidth(4) = 1000
                            
                            .Row = 1
                            .ColSel = .Cols - 1
                        End With
                        With msh������
                            .Left = sngLeft
                            .Top = sngTop
                            .Visible = True
                            .ZOrder 0
                            .SetFocus
                        End With
                        Cancel = True
                    End If
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh������_DblClick()
    Dim blnCancel As Boolean
    With mshBill
        .Text = msh������.TextMatrix(msh������.Row, 1)
        .TextMatrix(.Row, mHeadCol.����) = msh������.TextMatrix(msh������.Row, 1)
    End With
    msh������.Visible = False
    mshBill.SetFocus
    Call SendKeys("{ENTER}")
End Sub


Private Sub msh������_GotFocus()
    If msh������.Rows - 1 = 1 Then Call msh������_DblClick
End Sub

Private Sub msh������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call msh������_DblClick
    End If
End Sub

Private Sub msh������_LostFocus()
    msh������.ZOrder 1
    msh������.Visible = False
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer

    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����

            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > 40 Then
                MsgBox "ժҪ����,���������20�����ֻ�40���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If

            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mHeadCol.����)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mHeadCol.�ƻ�����))) <> "" Then
                        If Not IsNumeric(.TextMatrix(intLop, mHeadCol.�ƻ�����)) Then
                            MsgBox "��" & intLop & "���������ϵļƻ�������Ϊ�����ͣ����飡", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mHeadCol.�ƻ�����
                            Exit Function
                        End If

                    End If
                    
                    If Val(.TextMatrix(intLop, mHeadCol.�ƻ�����)) > 9999999999# Then
                        MsgBox "��" & intLop & "���������ϵļƻ��������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.�ƻ�����
                        Exit Function
                    End If

                    If Val(.TextMatrix(intLop, mHeadCol.����)) > 9999999999# Then
                        MsgBox "��" & intLop & "���������ϵĵ��۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.����
                        Exit Function
                    End If

                    If Val(.TextMatrix(intLop, mHeadCol.���)) > 9999999999999# Then
                        MsgBox "��" & intLop & "���������ϵĽ����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.�ƻ�����
                        Exit Function
                    End If
                    
                    If Trim(.TextMatrix(intLop, mHeadCol.�ϴι�Ӧ��)) = "" Then
                        MsgBox "��" & intLop & "����������δѡ��Ӧ�̣����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.�ϴι�Ӧ��
                        Exit Function
                    End If

                
                End If
            Next
        Else
            Exit Function
        End If
    End With

    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim lng��� As Long
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim �ƻ�����_IN As Integer
    Dim �ڼ�_IN As String
    Dim �ⷿID_IN As Long
    Dim ���Ʒ���_IN As Integer
    Dim ������_IN As String
    Dim ��������_IN As String
    Dim ����˵��_IN As String

    Dim ����ID_IN As Long
    Dim �ƻ�����_IN As Double
    Dim ����_IN As Double
    Dim ���_IN As Double
    Dim ǰ������_IN As Double
    Dim ��������_IN As Double
    Dim �������_IN As Double
    Dim ��������_IN As Double
    Dim ��������_IN As Double
    Dim �ϴι�Ӧ��_IN As String
    Dim �ϴ�������_IN As String
    Dim intRow As Integer
    Dim cllTemp As New Collection
    SaveCard = False
    With mshBill
        ID_IN = zlDatabase.GetNextId("���ϲɹ��ƻ�")
        NO_IN = Trim(txtNO)
        
        If NO_IN = "" Then NO_IN = zlDatabase.GetNextNo(77, mlng�ⷿid)
        If IsNull(NO_IN) Then Exit Function
        Me.txtNO.Tag = NO_IN
        
        �ƻ�����_IN = mint�ƻ�����
        ���Ʒ���_IN = mint���Ʒ���
        �ⷿID_IN = mlng�ⷿid
        ������_IN = gstrUserName
        ��������_IN = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ����˵��_IN = Trim(txtժҪ.Text)
        �ڼ�_IN = mstr�ڼ�

        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_���ϼƻ�����_DELETE('" & mlng�ƻ�ID & "')"
            cllTemp.Add gstrSQL
        End If
        'Zl_���ϼƻ���������_Insert
        gstrSQL = "Zl_���ϼƻ���������_Insert("
        '  Id_In       In ���ϲɹ��ƻ�.ID%Type,
        gstrSQL = gstrSQL & "" & ID_IN & ","
        '  ����_In     In ���ϲɹ��ƻ�.����%Type,
        gstrSQL = gstrSQL & "" & 0 & ","
        '  No_In       In ���ϲɹ��ƻ�.NO%Type,
        gstrSQL = gstrSQL & "'" & NO_IN & "',"
        '  �ƻ�����_In In ���ϲɹ��ƻ�.�ƻ�����%Type,
        gstrSQL = gstrSQL & "" & �ƻ�����_IN & ","
        '  �ڼ�_In     In ���ϲɹ��ƻ�.�ڼ�%Type,
        gstrSQL = gstrSQL & "'" & �ڼ�_IN & "',"
        '  �ⷿid_In   In ���ϲɹ��ƻ�.�ⷿid%Type,
        gstrSQL = gstrSQL & "" & IIf(�ⷿID_IN = 0, "NULL", �ⷿID_IN) & ","
        '  ����id_In   In ���ϲɹ��ƻ�.����id%Type,
        gstrSQL = gstrSQL & "NULL,"
        '  ���Ʒ���_In In ���ϲɹ��ƻ�.���Ʒ���%Type,
        gstrSQL = gstrSQL & "" & ���Ʒ���_IN & ","
        '  ������_In   In ���ϲɹ��ƻ�.������%Type,
        gstrSQL = gstrSQL & "'" & ������_IN & "',"
        '  ��������_In In ���ϲɹ��ƻ�.��������%Type,
        gstrSQL = gstrSQL & "to_date('" & ��������_IN & "','yyyy-mm-dd HH24:MI:SS'),"
        '  ����˵��_In In ���ϲɹ��ƻ�.����˵��%Type := Null
        gstrSQL = gstrSQL & "'" & ����˵��_IN & "')"
        cllTemp.Add gstrSQL
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng��� = .TextMatrix(intRow, mHeadCol.���)
                ����ID_IN = .TextMatrix(intRow, 0)
                ����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.����)) / Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.�ɱ���С��)
                ���_IN = Round(Val(.TextMatrix(intRow, mHeadCol.���)), g_С��λ��.obj_���С��.���С��)
                ǰ������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.ǰ������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.����С��)
                ��������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.��������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.����С��)
                �������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.�������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.����С��)
                �ƻ�����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.�ƻ�����)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.����С��)
                ��������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.��������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.����С��)
                ��������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.��������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.����С��)
                �ϴι�Ӧ��_IN = .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��)
                �ϴ�������_IN = .TextMatrix(intRow, mHeadCol.����)
                'zl_ҩƷ�ƻ������α�_INSERT( /*�ƻ�ID_IN*/, /*����ID_IN*/,/�빺����_IN /*�ƻ�����_IN*/,
                    '/*����_IN*/, /*���_IN*/, /*ǰ������_IN*/, /*��������_IN*/, /*�������_IN*/,
                    '/*�ϴι�Ӧ��_IN*/, /*�ϴ�������_IN*/ );

                gstrSQL = "zl_���ϼƻ������α�_INSERT(" & ID_IN & "," & ����ID_IN & "," & lng��� & ",0," & �ƻ�����_IN _
                    & "," & ����_IN & "," & ���_IN & "," & ǰ������_IN & "," & ��������_IN _
                    & "," & �������_IN & ",'" & �ϴι�Ӧ��_IN & "','" & �ϴ�������_IN & "'," & ��������_IN & "," & ��������_IN & ")"
                cllTemp.Add gstrSQL
            End If
        Next
    End With
    On Error GoTo errHandle
    ExecuteProcedureArrAy cllTemp, mstrCaption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveCheckCard() As Boolean
    '���ʱ������ͨ����˵ĵ��ݣ�У���д򹴵��У�
    Dim lng��� As Long
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim �ƻ�����_IN As Integer
    Dim �ڼ�_IN As String
    Dim �ⷿID_IN As Long
    Dim ���Ʒ���_IN As Integer
    Dim ������_IN As String
    Dim ��������_IN As String
    Dim ����˵��_IN As String

    Dim ����ID_IN As Long
    Dim �ƻ�����_IN As Double
    Dim ����_IN As Double
    Dim ���_IN As Double
    Dim ǰ������_IN As Double
    Dim ��������_IN As Double
    Dim �������_IN As Double
    Dim ��������_IN As Double
    Dim ��������_IN As Double
    Dim �ϴι�Ӧ��_IN As String
    Dim �ϴ�������_IN As String
    Dim intRow As Integer
    Dim cllTemp As New Collection
    Dim blnNoRecord As Boolean
    
    blnNoRecord = True
    
    SaveCheckCard = False
    
    With mshBill
        ID_IN = zlDatabase.GetNextId("���ϲɹ��ƻ�")
        NO_IN = Trim(txtNO)
        
        If NO_IN = "" Then NO_IN = zlDatabase.GetNextNo(77, mlng�ⷿid)
        If IsNull(NO_IN) Then Exit Function
        Me.txtNO.Tag = NO_IN
        
        �ƻ�����_IN = mint�ƻ�����
        ���Ʒ���_IN = mint���Ʒ���
        �ⷿID_IN = mlng�ⷿid
        ������_IN = IIf(Txt������.Caption <> "", Txt������, gstrUserName)
        ��������_IN = IIf(Txt��������.Caption <> "", Format(Txt��������.Caption, "yyyy-mm-dd hh:mm:ss"), Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss"))
        ����˵��_IN = Trim(txtժҪ.Text)
        �ڼ�_IN = mstr�ڼ�

        'ɾ��ԭ���ĵ���
        gstrSQL = "zl_���ϼƻ�����_DELETE('" & mlng�ƻ�ID & "')"
        cllTemp.Add gstrSQL
        
        'Zl_���ϼƻ���������_Insert
        gstrSQL = "Zl_���ϼƻ���������_Insert("
        '  Id_In       In ���ϲɹ��ƻ�.ID%Type,
        gstrSQL = gstrSQL & "" & ID_IN & ","
        '  ����_In     In ���ϲɹ��ƻ�.����%Type,
        gstrSQL = gstrSQL & "" & 0 & ","
        '  No_In       In ���ϲɹ��ƻ�.NO%Type,
        gstrSQL = gstrSQL & "'" & NO_IN & "',"
        '  �ƻ�����_In In ���ϲɹ��ƻ�.�ƻ�����%Type,
        gstrSQL = gstrSQL & "" & �ƻ�����_IN & ","
        '  �ڼ�_In     In ���ϲɹ��ƻ�.�ڼ�%Type,
        gstrSQL = gstrSQL & "'" & �ڼ�_IN & "',"
        '  �ⷿid_In   In ���ϲɹ��ƻ�.�ⷿid%Type,
        gstrSQL = gstrSQL & "" & IIf(�ⷿID_IN = 0, "NULL", �ⷿID_IN) & ","
        '  ����id_In   In ���ϲɹ��ƻ�.����id%Type,
        gstrSQL = gstrSQL & "NULL,"
        '  ���Ʒ���_In In ���ϲɹ��ƻ�.���Ʒ���%Type,
        gstrSQL = gstrSQL & "" & ���Ʒ���_IN & ","
        '  ������_In   In ���ϲɹ��ƻ�.������%Type,
        gstrSQL = gstrSQL & "'" & ������_IN & "',"
        '  ��������_In In ���ϲɹ��ƻ�.��������%Type,
        gstrSQL = gstrSQL & "to_date('" & ��������_IN & "','yyyy-mm-dd HH24:MI:SS'),"
        '  ����˵��_In In ���ϲɹ��ƻ�.����˵��%Type := Null
        gstrSQL = gstrSQL & "'" & ����˵��_IN & "')"
        cllTemp.Add gstrSQL
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And .TextMatrix(intRow, mHeadCol.У��) = "��" Then
                lng��� = lng��� + 1
                ����ID_IN = .TextMatrix(intRow, 0)
                ����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.����)) / Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.�ɱ���С��)
                ���_IN = Round(Val(.TextMatrix(intRow, mHeadCol.���)), g_С��λ��.obj_ɢװС��.���С��)
                ǰ������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.ǰ������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                ��������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.��������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                �������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.�������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                �ƻ�����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.�ƻ�����)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                ��������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.��������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                ��������_IN = Round(Val(.TextMatrix(intRow, mHeadCol.��������)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                �ϴι�Ӧ��_IN = .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��)
                �ϴ�������_IN = .TextMatrix(intRow, mHeadCol.����)
                'zl_ҩƷ�ƻ������α�_INSERT( /*�ƻ�ID_IN*/, /*����ID_IN*/,/�빺����_IN /*�ƻ�����_IN*/,
                    '/*����_IN*/, /*���_IN*/, /*ǰ������_IN*/, /*��������_IN*/, /*�������_IN*/,
                    '/*�ϴι�Ӧ��_IN*/, /*�ϴ�������_IN*/ );

                gstrSQL = "zl_���ϼƻ������α�_INSERT(" & ID_IN & "," & ����ID_IN & "," & lng��� & ",0," & �ƻ�����_IN _
                    & "," & ����_IN & "," & ���_IN & "," & ǰ������_IN & "," & ��������_IN _
                    & "," & �������_IN & ",'" & �ϴι�Ӧ��_IN & "','" & �ϴ�������_IN & "'," & ��������_IN & "," & ��������_IN & ")"
                cllTemp.Add gstrSQL
                
                blnNoRecord = False
            End If
        Next
    End With
    
    If blnNoRecord = True Then
        MsgBox "��ѡ���������ͨ���ļ�¼��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mlng�ƻ�ID = ID_IN

    On Error GoTo errHandle
    
    ExecuteProcedureArrAy cllTemp, mstrCaption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCheckCard = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim Dbl��� As Double
    Dim intLop As Integer

    Dbl��� = 0

    With mshBill
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                Dbl��� = Dbl��� + Val(.TextMatrix(intLop, mHeadCol.���))
            End If
        Next
    End With

    lblPurchasePrice.Caption = "���ϼƣ�" & Format(Dbl���, mFMT.FM_���)
End Sub

Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    zlCommFun.OpenIme (True)
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Function SetStuffRows(ByVal lng����ID As Long, ByVal str���� As String, _
        ByVal str��� As String, ByVal str���� As String, ByVal str��λ As String, _
        ByVal dblָ�������� As Double, ByVal dbl����ϵ�� As Double) As Boolean
    Dim rsData As New Recordset
    Dim intCount As Integer
    Dim intRow As Integer
    Dim intCol As Integer

    Dim lng���� As Long
    Dim dbl������� As Double
    Dim dbl�ɱ��� As Double

    On Error GoTo errH
    SetStuffRows = False

    With mshBill
        intRow = .Row
        For intCount = 1 To .Rows - 1
            If intCount <> intRow And .TextMatrix(intCount, 0) <> "" Then
                If .TextMatrix(intCount, 0) = lng����ID Then
                    MsgBox "�Բ��𣬸��������������ˣ��������䣡", vbOKOnly + vbExclamation, gstrSysName
                    Exit Function
                End If
            End If
        Next

        For intCol = 0 To .Cols - 1
            .TextMatrix(intRow, intCol) = ""
        Next
    End With

    With mshBill
        .TextMatrix(.Row, mHeadCol.���) = .Row
        .TextMatrix(.Row, mHeadCol.����) = str����
        .TextMatrix(.Row, 0) = lng����ID
        .TextMatrix(.Row, mHeadCol.����ϵ��) = dbl����ϵ��
        
        'ȡƽ���ɱ��ۣ����û�����ã���ȡָ�������ۣ�
        gstrSQL = "Select �ɱ���,ָ��������,�б���� From  �������� Where ����ID=[1]"
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ɱ���", lng����ID)
        
        dbl�ɱ��� = NVL(rsData!�ɱ���, 0)
        If dbl�ɱ��� = 0 Then dbl�ɱ��� = NVL(rsData!ָ��������, 0)
        .TextMatrix(.Row, mHeadCol.�б����) = IIf(Val(NVL(rsData!�б����)) = "1", "��", "")
        
        gstrSQL = "" & _
            " SELECT MIN (B.����) AS ��Ӧ��, MIN (�ϴβ���) AS �ϴβ���,SUM(ʵ������) AS ������� " & _
            " FROM ҩƷ��� A, (SELECT ID,���� FROM ��Ӧ�� WHERE SUBSTR(����,5,1)=1) B  " & _
            " WHERE A.����=1 AND A.�ϴι�Ӧ��ID = B.ID(+) And A.ҩƷID=[1] " & _
            IIf(mlng�ⷿid = 0, "", " AND A.�ⷿID=[2]")
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ϴι�Ӧ�̼�������Ϣ", lng����ID, mlng�ⷿid)
        
        If NVL(rsData!�������, 0) = 0 And NVL(rsData!��Ӧ��) = "" And NVL(rsData!�ϴβ���) = "" Then
            gstrSQL = "Select c.���� As ��Ӧ��, Decode(a.�ϴβ���, Null, b.����, a.�ϴβ���) As �ϴβ���, 0 As �������" & _
                       " From �������� A, �շ���ĿĿ¼ B, (Select ID, ���� From ��Ӧ�� Where Substr(����, 5, 1) = 1 And (վ�� = [3] Or վ�� Is Null)) C" & _
                       " Where a.����id = b.Id And a.�ϴι�Ӧ��id = c.Id And a.����id = [1]"
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ϴι�Ӧ�̼�������Ϣ", lng����ID, mlng�ⷿid, gstrNodeNo)
        End If
            
        If Not rsData.EOF Then
            .TextMatrix(.Row, mHeadCol.�������) = Format(IIf(IsNull(rsData!�������), 0, rsData!�������) / dbl����ϵ��, mFMT.FM_����)
            .TextMatrix(.Row, mHeadCol.�ϴι�Ӧ��) = IIf(IsNull(rsData!��Ӧ��), "", rsData!��Ӧ��)
            .TextMatrix(.Row, mHeadCol.����) = IIf(IsNull(rsData!�ϴβ���), str����, rsData!�ϴβ���)
            SetNumer lng����ID, mlng�ⷿid, .TextMatrix(.Row, mHeadCol.�������), .Row, mint�ƻ�����, mint���Ʒ���
        End If
        .TextMatrix(.Row, mHeadCol.����) = str����
        .TextMatrix(.Row, mHeadCol.���) = str���
        .TextMatrix(.Row, mHeadCol.��λ) = str��λ
        .TextMatrix(.Row, mHeadCol.����) = Format(dbl�ɱ��� * dbl����ϵ��, mFMT.FM_�ɱ���)
        
        
        gstrSQL = "" & _
            "   Select sum(nvl(����,0)) ����,sum(nvl(����,0)) ���� " & _
            "   From ���ϴ����޶�  " & _
            "   where ����ID=[1] " & IIf(mlng�ⷿid = 0, "", " and �ⷿID=[2]") & _
            "   Group By ����ID"
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������", lng����ID, mlng�ⷿid)
        If Not rsData.EOF Then
            .TextMatrix(.Row, mHeadCol.�洢����) = Format(Val(NVL(rsData!����)) / dbl����ϵ��, mFMT.FM_����)
            .TextMatrix(.Row, mHeadCol.�洢����) = Format(Val(NVL(rsData!����)) / dbl����ϵ��, mFMT.FM_����)
        End If
        
    End With
    rsData.Close
    SetStuffRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

'�����룬���ƣ���������ĳһ��
Private Function FindData(ByVal mshBill As BillEdit, ByVal int�Ƚ��� As Integer, _
    ByVal str�Ƚ�ֵ As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    Dim strKey As String
    FindData = True
    
    On Error GoTo errHandle
    With mshBill
        If .Rows = 2 Then Exit Function
        If str�Ƚ�ֵ = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                If InStr(1, UCase(strCode), UCase(str�Ƚ�ֵ)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int�Ƚ���
                    .MsfObj.TopRow = .Row
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = " SELECT DISTINCT b.���� " & _
                  " FROM " & _
                  "    (SELECT DISTINCT A.�շ�ϸĿid " & _
                  "    FROM �շ���Ŀ���� A" & _
                  "    Where A.���� LIKE [1]) a," & _
                  " �շ���ĿĿ¼ B " & _
                  " Where a.�շ�ϸĿid = b.ID"
        
        strKey = IIf(gstrMatchMethod = "0", "%", "") & str�Ƚ�ֵ & "%"
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strKey)
                  
        If rsCode.EOF Then
            FindData = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!����)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int�Ƚ���
                        .MsfObj.TopRow = .Row
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    FindData = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
