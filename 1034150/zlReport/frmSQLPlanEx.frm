VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSQLPlanEx 
   Caption         =   "�鿴ִ�мƻ�"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15180
   Icon            =   "frmSQLPlanEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   15180
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMiddle 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8535
      TabIndex        =   3
      Top             =   60
      Width           =   4920
      Begin VB.TextBox txtEnd 
         Height          =   300
         Left            =   4080
         MaxLength       =   7
         TabIndex        =   7
         Text            =   "1000000"
         Top             =   217
         Width           =   735
      End
      Begin VB.TextBox txtBegin 
         Height          =   300
         Left            =   3000
         MaxLength       =   7
         TabIndex        =   5
         Text            =   "3000"
         Top             =   217
         Width           =   735
      End
      Begin VB.CheckBox chkMiddle 
         Caption         =   "������ͱ�"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "���ͱ���¼����Χ"
         Height          =   255
         Left            =   1530
         TabIndex        =   8
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "~"
         Height          =   135
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   13440
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RtbRemark 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10821
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      FileName        =   "D:\�����Ż���ƹ淶2014������.rtf"
      TextRTF         =   $"frmSQLPlanEx.frx":6852
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   1270
      ButtonWidth     =   1376
      ButtonHeight    =   1270
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "img��ɫ"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "Copy"
            Description     =   "����"
            Object.ToolTipText     =   "ִ�б���"
            Object.Tag             =   "ִ��"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����Ϊ������"
            Key             =   "Save"
            Description     =   "����Ϊ"
            Object.ToolTipText     =   "����Ϊtxt"
            Object.Tag             =   "����Ϊ"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ˢ��"
            Key             =   "Review"
            ImageKey        =   "View"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
            Key             =   "Quit"
            Description     =   "�˳�"
            Object.ToolTipText     =   "�˳�"
            Object.Tag             =   "�˳�"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img��ɫ 
      Left            =   11160
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQLPlanEx.frx":11C06
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQLPlanEx.frx":12300
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQLPlanEx.frx":1251A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQLPlanEx.frx":12734
            Key             =   "View"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPlan 
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   13020
      _cx             =   22966
      _cy             =   12515
      Appearance      =   2
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
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   235
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSQLPlanEx.frx":1294E
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   4
      OutlineCol      =   1
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
End
Attribute VB_Name = "frmSQLPlanEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQLCheck As String
Private mblnPro As Boolean   '�Ƿ�����������
Private mintType As Integer  '0-�鿴ִ�мƻ���1-�鿴�Ż�����
Private mstrDataName As String
Private mintConnect As Integer

Public Function ShowMe(frmParent As Object, ByVal intConnect As Integer, ByVal strSQLCheck As String, _
    Optional ByVal intType As Integer, Optional ByVal strDataName As String) As Boolean
    mstrSQLCheck = strSQLCheck
    mintType = intType
    mstrDataName = strDataName
    mintConnect = intConnect
    
    Me.Show 1, frmParent
    ShowMe = mblnPro
End Function

Private Sub Form_Activate()
    If Me.Visible And Val(Me.Tag) = Val("-1-�쳣") Then
        Unload Me
    End If
End Sub

Private Sub chkMiddle_Click()
    If chkMiddle.Value Then
        txtBegin.Enabled = True
        txtEnd.Enabled = True
        txtBegin.BackColor = vsPlan.BackColor
        txtEnd.BackColor = vsPlan.BackColor
    Else
        txtBegin.Enabled = False
        txtEnd.Enabled = False
        txtBegin.BackColor = Me.BackColor
        txtEnd.BackColor = Me.BackColor
    End If
    If Me.Visible Then Call CheckSQLPlan(mstrSQLCheck, vsPlan)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLPlanEx", "MiddleTable", chkMiddle.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long, StrPar As String, blnSuccess As Boolean
    
    mblnPro = CheckSQLPlan(mstrSQLCheck, vsPlan)
    If mintType = 0 Then
        mblnPro = CheckSQLPlan(mstrSQLCheck, vsPlan, mintConnect, blnSuccess)
        If blnSuccess = False Then
            Me.Tag = "-1"
        End If
        
        RtbRemark.Visible = False
        Me.Caption = "�鿴ִ�мƻ�"
        tbrMain.Buttons("Review").Visible = True
        
        fraMiddle.Visible = True
        StrPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLPlanEx", "MiddleTable", "1")
        chkMiddle.Value = Val(StrPar)
        Call chkMiddle_Click
        StrPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLPlanEx", "MiddleTableRows", "3000|1000000")
        txtBegin.Text = Split(StrPar, "|")(0)
        txtEnd.Text = Split(StrPar, "|")(1)
        txtBegin.Tag = txtBegin.Text: txtEnd.Tag = txtEnd.Text
    Else
        vsPlan.Visible = False
        Me.Caption = "�鿴�Ż�����"
        tbrMain.Buttons("Review").Visible = False
        fraMiddle.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsPlan.Top = 0: vsPlan.Left = 0
    vsPlan.Width = Me.ScaleWidth - vsPlan.Left
    vsPlan.Height = Me.ScaleHeight - vsPlan.Top - 60
    
    RtbRemark.Top = 0: RtbRemark.Left = 0
    RtbRemark.Width = Me.ScaleWidth - vsPlan.Left
    RtbRemark.Height = Me.ScaleHeight - vsPlan.Top - 60
    fraMiddle.Left = Me.ScaleWidth - fraMiddle.Width - 200
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long, strText As String
    Dim strFormat As String * 4
    Dim strSpace As String * 100
    
    Select Case Button.Key
    Case "Copy"
        If mintType = 1 Then
            Clipboard.Clear
            Call Clipboard.SetText(RtbRemark.Text)
        Else
            With vsPlan
                strSpace = " "
                For i = .FixedRows To .Rows - 1
                    strFormat = .TextMatrix(i, 0)
                    strText = strText & IIF(strText = "", "", vbCrLf) & strFormat & " " & Mid(strSpace, 100 - Val(.RowOutlineLevel(i))) & .TextMatrix(i, 1)
                Next
                If strText <> "" Then
                    Clipboard.Clear
                    Call Clipboard.SetText(strText)
                End If
            End With
        End If
    Case "Save"
        If mintType = 1 Then
            With CommonDialog
                .DialogTitle = "�����ļ�"
                .Filter = "RTF Files|*.rtf"
                .Flags = &H200000 + &H2000 + &H2 + &H800
                .InitDir = App.Path
                .FileName = "SQL��ѯ�Ż�"
                .CancelError = True
                On Error Resume Next
                .ShowSave
                If Err.Number = 0 Then
                    RtbRemark.SaveFile .FileName
                End If
            End With
        Else
            With CommonDialog
                .DialogTitle = "�����ļ�"
                .Filter = "RTF Files|*.RTF"
                .Flags = &H200000 + &H2000 + &H2 + &H800
                .InitDir = App.Path
                .FileName = mstrDataName
                .CancelError = True
                On Error Resume Next
                .ShowSave
            End With
            If Err.Number = 0 Then
                With vsPlan
                    strSpace = " "
                    strText = "--------------" & mstrDataName & "-ִ�мƻ�" & "--------------"
                    strText = strText & vbCrLf & "����  ����"
                    For i = .FixedRows To .Rows - 1
                        strFormat = .TextMatrix(i, 0)
                        strText = strText & vbCrLf & strFormat & " " & Mid(strSpace, 100 - Val(.RowOutlineLevel(i))) & .TextMatrix(i, 1)
                    Next
                    If strText <> "" Then
                        RtbRemark.Text = strText
                        RtbRemark.SelStart = 1: RtbRemark.SelLength = Len(RtbRemark.Text)
                        RtbRemark.SelFontSize = 5.5
                        RtbRemark.SaveFile CommonDialog.FileName
                    End If
                End With
            End If
        End If
    Case "Review"
        mblnPro = CheckSQLPlan(mstrSQLCheck, vsPlan)
    Case "Quit"
        Unload Me
    End Select
End Sub

Private Sub txtBegin_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtBegin_Validate(Cancel As Boolean)
    If Val(txtEnd.Text) <= Val(txtBegin.Text) Then
        MsgBox "��С��¼��Ӧ�ñ�����¼��С�����顣", vbInformation, App.Title
        Cancel = True
    End If
    If Val(txtBegin.Text) < 1000 Then
        MsgBox "���ͱ���¼��Ӧ�ô��ڵ���1000����¼����", vbInformation, App.Title
        Cancel = True
    End If
    If Cancel = False Then
        txtBegin.Tag = txtBegin.Text
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLPlanEx", "MiddleTableRows", txtBegin.Text & "|" & txtEnd.Text
    Else
        txtBegin.Text = txtBegin.Tag
    End If
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtEnd_Validate(Cancel As Boolean)
    If Val(txtEnd.Text) <= Val(txtBegin.Text) Then
        MsgBox "��С��¼��Ӧ�ñ�����¼��С�����顣", vbInformation, App.Title
        Cancel = True
        txtEnd.Text = txtEnd.Tag
    End If
    If Cancel = False Then
        txtEnd.Tag = txtEnd.Text
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLPlanEx", "MiddleTableRows", txtBegin.Text & "|" & txtEnd.Text
    End If
End Sub

Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsPlan.ForeColorSel = vsPlan.Cell(flexcpForeColor, NewRow, NewCol)
End Sub