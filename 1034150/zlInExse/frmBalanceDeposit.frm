VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBalanceDeposit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������Ԥ���˿�"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   Icon            =   "frmBalanceDeposit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   825
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   4335
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4965
      TabIndex        =   1
      ToolTipText     =   "�ȼ���F2"
      Top             =   4305
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6405
      TabIndex        =   2
      ToolTipText     =   "�ȼ�:Esc"
      Top             =   4305
      Width           =   1410
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDeposit 
      Height          =   3600
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   7980
      _cx             =   14076
      _cy             =   6350
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      BackColorBkg    =   -2147483633
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceDeposit.frx":06EA
      ScrollTrack     =   0   'False
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
   Begin VB.Label lbl����� 
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3120
      TabIndex        =   7
      Top             =   4350
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2175
      TabIndex        =   6
      Top             =   4350
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   330
      Picture         =   "frmBalanceDeposit.frx":0829
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "�����Ǳ��ν��ʲ��˵�������Ԥ�����,  �������Ҫ�����˿�    "
      Height          =   360
      Left            =   885
      TabIndex        =   5
      Top             =   135
      Width           =   3420
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   3
      Top             =   4395
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmBalanceDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsDeposit As ADODB.Recordset, mrsInfo As ADODB.Recordset
Private mblnUnload As Boolean
Private mlng����ID As Long, mlng����ID As Long
Private mlngModul As Long, mblnAll As Boolean
Private mblnDateMoved As Boolean
Private mstrסԺ���� As String
Private mstrDepositDate    As String
Private mintԤ�����    As Integer
Private mstrCardPrivs As String, mstrForceNote As String
Private mstrǿ�����ֲ���Ա As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveData
End Sub

Private Sub SaveData()
    Dim i As Integer, cllSQL As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSql As String, strFailNo As String, strXMLExpend As String, dblMoney As Double
    Dim strCardNo As String, strPassWord As String, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllSquareBalance As Collection, strIDs As String, strNos As String, dbl��� As Double
    Dim rsTmp As ADODB.Recordset, lngRow As Long, j As Integer, strValue As String
    Dim strInXML As String, strOutXML As String, strExpend As String, strBalanceIDs As String
    If lbl�����.Visible Then
        dbl��� = Val(lbl�����.Caption)
    End If
    For i = 1 To vsfDeposit.Rows - 1
        Set cllSQL = New Collection
        Set cllSquareBalance = New Collection
        Set cllThreeSwap = New Collection
        With vsfDeposit
            If Val(.TextMatrix(i, .ColIndex("��Ԥ��"))) <> 0 Then
                If Val(.TextMatrix(i, .ColIndex("����"))) = 0 Then
                    If .TextMatrix(i, .ColIndex("ת��")) = 1 Then
                        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
                                Val(.RowData(i)), False, _
                            mrsInfo!����, mrsInfo!�Ա�, mrsInfo!����, Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), strCardNo, strPassWord, _
                            False, True, False, False, cllSquareBalance) = False Then
                            strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                        Else
                            zlXML.ClearXmlText
                            zlXML.AppendNode "IN"
                            zlXML.appendData "CZLX", "2"
                            zlXML.AppendNode "IN", True
                            strXMLExpend = zlXML.XmlText
                            zlXML.ClearXmlText
                            If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, Val(.RowData(i)), _
                                strCardNo, Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), "", strXMLExpend) = False Then
                                strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                            Else
                                mrsDeposit.Filter = "���㿨����='" & .TextMatrix(i, .ColIndex("���㿨����")) & "'"
                                dblMoney = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                                Do While Not mrsDeposit.EOF
                                    If dblMoney > 0 Then
                                        If dblMoney > Val(mrsDeposit!���) Then
                                            strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    Val(mrsDeposit!���) & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "')"
                                            dblMoney = dblMoney - Val(mrsDeposit!���)
                                        Else
                                            strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    dblMoney & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "')"
                                            dblMoney = 0
                                        End If
                                        zlAddArray cllSQL, strSql
                                    End If
                                    mrsDeposit.MoveNext
                                Loop
                                zlExecuteProcedureArrAy cllSQL, Me.Caption, True
                                zlXML.ClearXmlText
                                zlXML.AppendNode "IN"
                                zlXML.appendData "CZLX", "2"
                                zlXML.AppendNode "IN", True
                                strXMLExpend = zlXML.XmlText
                                zlXML.ClearXmlText
                                If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, Val(.RowData(i)), _
                                    strCardNo, Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), "", strXMLExpend) = False Then
                                    gcnOracle.RollbackTrans
                                    strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                                Else
                                    zlXML.ClearXmlText
                                    zlXML.AppendNode "IN"
                                        zlXML.appendData "CZLX", "2"
                                    zlXML.AppendNode "IN", True
                                    strXMLExpend = zlXML.XmlText
                                    zlXML.ClearXmlText
                                    If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, Val(.RowData(i)), strCardNo, _
                                        mlng����ID, Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
                                        gcnOracle.RollbackTrans
                                        strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                                    Else
                                        Set cllUpdate = New Collection
                                        Set cllThreeSwap = New Collection
    '                                    Call zlAddUpdateSwapSQL(False, mlng����ID, Val(.RowData(i)), False, strCardNo, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
                                        Call zlAddThreeSwapSQLToCollection(False, mlng����ID, Val(.RowData(i)), False, strCardNo, strSwapExtendInfor, cllThreeSwap)
                                        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
                                        zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, False, True
                                    End If
                                End If
                            End If
                        End If
                    Else
                        '����˿�
                        strBalanceIDs = ""
                        zlXML.ClearXmlText
                        mrsDeposit.Filter = "���㿨����='" & .TextMatrix(i, .ColIndex("���㿨����")) & "'"
                        dblMoney = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                        Call zlXML.AppendNode("JSLIST")
                        Do While Not mrsDeposit.EOF
                            If dblMoney > 0 Then
                                If dblMoney > Val(mrsDeposit!���) Then
                                    strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    Val(mrsDeposit!���) & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "')"
                                    zlAddArray cllSQL, strSql
                                    strSql = "Select ID,����,������ˮ��,����˵�� From ����Ԥ����¼ Where ID = [1]"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsDeposit!Ԥ��ID))
                                    If Not rsTmp.EOF Then
                                        Call zlXML.AppendNode("JS")
                                            Call zlXML.appendData("KH", Nvl(rsTmp!����))
                                            Call zlXML.appendData("JYLSH", Nvl(rsTmp!������ˮ��))
                                            Call zlXML.appendData("JYSM", Nvl(rsTmp!����˵��))
                                            Call zlXML.appendData("ZFJE", Val(mrsDeposit!���))
                                            Call zlXML.appendData("JSLX", 1)
                                            Call zlXML.appendData("ID", Nvl(rsTmp!ID))
                                        Call zlXML.AppendNode("JS", True)
                                        strSql = "Zl_�����˿���Ϣ_Insert("
                                        strSql = strSql & mlng����ID & ","
                                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                                        strSql = strSql & Val(mrsDeposit!���) & ",'"
                                        strSql = strSql & Nvl(rsTmp!����) & "','"
                                        strSql = strSql & Nvl(rsTmp!������ˮ��) & "','"
                                        strSql = strSql & Nvl(rsTmp!����˵��) & "')"
                                        zlAddArray cllThreeSwap, strSql
                                        strBalanceIDs = strBalanceIDs & "," & Val(Nvl(rsTmp!ID))
                                    End If
                                    dblMoney = dblMoney - Val(mrsDeposit!���)
                                Else
                                    strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                                    "'" & mrsDeposit!NO & "'" & ",0," & _
                                                    dblMoney & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "')"
                                    zlAddArray cllSQL, strSql
                                    strSql = "Select ID,����,������ˮ��,����˵�� From ����Ԥ����¼ Where ID = [1]"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsDeposit!Ԥ��ID))
                                    If Not rsTmp.EOF Then
                                        Call zlXML.AppendNode("JS")
                                            Call zlXML.appendData("KH", Nvl(rsTmp!����))
                                            Call zlXML.appendData("JYLSH", Nvl(rsTmp!������ˮ��))
                                            Call zlXML.appendData("JYSM", Nvl(rsTmp!����˵��))
                                            Call zlXML.appendData("ZFJE", dblMoney)
                                            Call zlXML.appendData("JSLX", 1)
                                            Call zlXML.appendData("ID", Nvl(rsTmp!ID))
                                        Call zlXML.AppendNode("JS", True)
                                        strSql = "Zl_�����˿���Ϣ_Insert("
                                        strSql = strSql & mlng����ID & ","
                                        strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                                        strSql = strSql & dblMoney & ",'"
                                        strSql = strSql & Nvl(rsTmp!����) & "','"
                                        strSql = strSql & Nvl(rsTmp!������ˮ��) & "','"
                                        strSql = strSql & Nvl(rsTmp!����˵��) & "')"
                                        zlAddArray cllThreeSwap, strSql
                                        strBalanceIDs = strBalanceIDs & "," & Val(Nvl(rsTmp!ID))
                                    End If
                                    dblMoney = 0
                                End If
                            End If
                            mrsDeposit.MoveNext
                        Loop
                        Call zlXML.AppendNode("JSLIST", True)
                        strXMLExpend = zlXML.XmlText
                        strInXML = zlXML.XmlText
                        If strBalanceIDs <> "" Then strBalanceIDs = "1|" & Mid(strBalanceIDs, 2)
                        
                        If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, Val(.RowData(i)), False, strCardNo, _
                            strBalanceIDs, Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), strSwapGlideNO, strSwapMemo, strXMLExpend) = False Then
                            strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                        Else
                            zlExecuteProcedureArrAy cllSQL, Me.Caption, True
                            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, True, True
                            If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModul, Val(.RowData(i)), False, strInXML, _
                                 mlng����ID, strOutXML, strExpend) = False Then
                                gcnOracle.RollbackTrans:
                                strFailNo = strFailNo & "," & .TextMatrix(i, .ColIndex("���㿨����"))
                            Else
                                '�ύ
                                Set cllThreeSwap = New Collection
                                If zlXML_Init = True Then
                                    If strOutXML <> "" Then
                                        If zlXML_LoadXMLToDOMDocument(strOutXML, False) Then
                                            Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
                                            For j = 0 To lngRow - 1
                                                Call zlXML_GetNodeValue("ID", i, strValue)
                                                strSql = "Zl_�����˿���Ϣ_Insert("
                                                strSql = strSql & mlng����ID & ","
                                                strSql = strSql & Val(strValue) & ","
                                                strSql = strSql & 0 & ",'"
                                                Call zlXML_GetNodeValue("KH", i, strValue)
                                                strSql = strSql & strValue & "','"
                                                Call zlXML_GetNodeValue("TKLSH", i, strValue)
                                                strSql = strSql & strValue & "','"
                                                Call zlXML_GetNodeValue("TKSM", i, strValue)
                                                strSql = strSql & strValue & "',"
                                                strSql = strSql & 1 & ")"
                                                zlAddArray cllThreeSwap, strSql
                                            Next j
                                        End If
                                    End If
                                    
                                    If strExpend <> "" Then
                                        strSwapExtendInfor = ""
                                        If zlXML_LoadXMLToDOMDocument(strExpend, False) Then
                                            Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
                                            For j = 0 To lngRow - 1
                                                Call zlXML_GetNodeValue("XMMC", j, strValue)
                                                strSwapExtendInfor = strSwapExtendInfor & "||" & strValue
                                                Call zlXML_GetNodeValue("XMNR", j, strValue)
                                                strSwapExtendInfor = strSwapExtendInfor & "|" & strValue
                                            Next j
                                        End If
                                    End If
                                    If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
                                End If
                                strSql = "Select ���� From ����Ԥ����¼ Where ����ID= [1] And �����ID= [2]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, Val(.RowData(i)))
                                If Not rsTmp.EOF Then
                                    strCardNo = Nvl(rsTmp!����)
                                End If
    '                            Call zlAddUpdateSwapSQL(False, mlng����ID, Val(.RowData(i)), False, strCardNo, "", "", cllUpdate, 0)
                                Call zlAddThreeSwapSQLToCollection(False, mlng����ID, Val(.RowData(i)), False, strCardNo, strSwapExtendInfor, cllThreeSwap)
    '                            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
                                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, True, True
                                gcnOracle.CommitTrans
                            End If
                        End If
                    End If
                Else
                    '����
                    mrsDeposit.Filter = "���㿨����='" & .TextMatrix(i, .ColIndex("���㿨����")) & "'"
                    dblMoney = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                    
                    Do While Not mrsDeposit.EOF
                        If dblMoney > 0 Then
                            If dblMoney > Val(mrsDeposit!���) Then
                                strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                        "'" & mrsDeposit!NO & "'" & ",1," & _
                                        Val(mrsDeposit!���) & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "'" & IIf(lbl�����.Visible And dbl��� <> 0, "," & dbl��� & ",'", ",Null,'") & mstrForceNote & "')"
                                dblMoney = dblMoney - Val(mrsDeposit!���)
                                If dbl��� <> 0 And lbl���.Visible Then
                                    dbl��� = 0
                                End If
                            Else
                                strSql = "Zl_����Ԥ����¼_�����˿�(" & Val(mrsDeposit!ID) & "," & _
                                        "'" & mrsDeposit!NO & "'" & ",1," & _
                                        dblMoney & "," & mlng����ID & "," & mlng����ID & ",Null,Null,Null,'" & Nvl(mrsDeposit!���㷽ʽ) & "'" & IIf(lbl�����.Visible And dbl��� <> 0, "," & dbl��� & ",'", ",Null,'") & mstrForceNote & "')"
                                dblMoney = 0
                                If dbl��� <> 0 And lbl���.Visible Then
                                    dbl��� = 0
                                End If
                            End If
                            zlAddArray cllSQL, strSql
                        End If
                        mrsDeposit.MoveNext
                    Loop
                    zlExecuteProcedureArrAy cllSQL, Me.Caption
                End If
            End If
        End With
    Next i
    If strFailNo <> "" Then
        MsgBox "������������Ԥ�������˿�����г��ִ���,��ʹ������˿�ܶԸ���Ԥ��������˿�!" & vbCrLf & Mid(strFailNo, 2)
    End If
    mblnUnload = True
    Unload Me
End Sub

Public Sub ShowMe(frmMain As Object, lngModule As Long, lng����ID As Long, lng����ID As Long, blnAll As Boolean, _
                  Optional ByVal blnDateMoved As Boolean = False, Optional ByVal strסԺ���� As String = "", Optional ByVal strDepositDate As String = "", Optional ByVal intԤ����� As Integer)
    mlngModul = lngModule
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mblnAll = blnAll
    mblnDateMoved = blnDateMoved
    mstrסԺ���� = strסԺ����
    mstrDepositDate = strDepositDate
    mintԤ����� = intԤ�����
    On Error Resume Next
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim i As Integer
    Dim lngRow As Long
    
    mblnUnload = False
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    
    vsfDeposit.Clear 1
    vsfDeposit.Rows = 2
    Set mrsDeposit = GetThreeDeposit(mlng����ID, mblnDateMoved, mstrסԺ����, mstrDepositDate, mintԤ�����)
    Do While Not mrsDeposit.EOF
        With vsfDeposit
            lngRow = 0
            For i = 1 To .Rows - 1
                If .RowData(i) = Nvl(mrsDeposit!�����ID) Then
                    lngRow = i
                    Exit For
                End If
            Next i
            If lngRow = 0 Then
                .TextMatrix(.Rows - 1, 0) = Nvl(mrsDeposit!���㿨����)
                .TextMatrix(.Rows - 1, 1) = Nvl(mrsDeposit!���㷽ʽ)
                .TextMatrix(.Rows - 1, 2) = Format(Nvl(mrsDeposit!���), "0.00")
                If mblnAll Then
                    .TextMatrix(.Rows - 1, 3) = Format(Nvl(mrsDeposit!���), "0.00")
                Else
                    .TextMatrix(.Rows - 1, 3) = Format(0, "0.00")
                End If
                .TextMatrix(.Rows - 1, 4) = 0
                If Val(mrsDeposit!����) = 1 Then
                    '��������,�����޸�
                    .Cell(flexcpData, .Rows - 1, 4) = 1
                    .Cell(flexcpBackColor, .Rows - 1, 4) = vbWhite
                Else
                    .Cell(flexcpData, .Rows - 1, 4) = 0
                    .Cell(flexcpBackColor, .Rows - 1, 4) = &H8000000F
                End If
                .TextMatrix(.Rows - 1, 5) = Nvl(mrsDeposit!Ԥ��ID)
                .TextMatrix(.Rows - 1, 6) = Nvl(mrsDeposit!����)
                .TextMatrix(.Rows - 1, 7) = Nvl(mrsDeposit!ID)
                .TextMatrix(.Rows - 1, 8) = Nvl(mrsDeposit!��¼״̬)
                .RowData(.Rows - 1) = Nvl(mrsDeposit!�����ID)
                .Rows = .Rows + 1
            Else
                .TextMatrix(lngRow, 0) = Nvl(mrsDeposit!���㿨����)
                .TextMatrix(lngRow, 1) = Nvl(mrsDeposit!���㷽ʽ)
                .TextMatrix(lngRow, 2) = Format(Val(.TextMatrix(lngRow, 2)) + Val(Nvl(mrsDeposit!���)), "0.00")
                If mblnAll Then
                    .TextMatrix(lngRow, 3) = Format(Val(.TextMatrix(lngRow, 3)) + Val(Nvl(mrsDeposit!���)), "0.00")
                Else
                    .TextMatrix(lngRow, 3) = Format(0, "0.00")
                End If
                .TextMatrix(lngRow, 4) = 0
                If Val(mrsDeposit!����) = 1 Then
                    '��������,�����޸�
                    .Cell(flexcpData, lngRow, 4) = 1
                    .Cell(flexcpBackColor, lngRow, 4) = vbWhite
                Else
                    .Cell(flexcpData, lngRow, 4) = 0
                    .Cell(flexcpBackColor, lngRow, 4) = &H8000000F
                End If
                .TextMatrix(lngRow, 5) = Nvl(mrsDeposit!Ԥ��ID)
                .TextMatrix(lngRow, 6) = Nvl(mrsDeposit!����)
                .TextMatrix(lngRow, 7) = Nvl(mrsDeposit!ID)
                .TextMatrix(lngRow, 8) = Nvl(mrsDeposit!��¼״̬)
                .RowData(lngRow) = Nvl(mrsDeposit!�����ID)
            End If
        End With
        mrsDeposit.MoveNext
    Loop
    If mrsDeposit.RecordCount = 0 Then
        mblnUnload = True
        Unload Me: Exit Sub
    End If
    
    vsfDeposit.Rows = vsfDeposit.Rows - 1
    strSql = "Select ����,����,�Ա� From ������Ϣ Where ����ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnUnload = False Then
        If MsgBox("�Ƿ�ȷ��ȡ��Ԥ�����˿�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If
    mstrForceNote = ""
    mstrǿ�����ֲ���Ա = ""
    mblnUnload = False
End Sub

Private Sub vsfDeposit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    If Col = 3 Then
        If IsNumeric(vsfDeposit.TextMatrix(Row, 3)) = False Then
            MsgBox "��������ȷ���˿���!", vbInformation, gstrSysName
            vsfDeposit.TextMatrix(Row, 3) = "0.00"
        End If
        If Val(vsfDeposit.TextMatrix(Row, 3)) > Val(vsfDeposit.TextMatrix(Row, 2)) Then
            MsgBox "������˿������,����", vbInformation, gstrSysName
            vsfDeposit.TextMatrix(Row, 3) = "0.00"
        End If
        vsfDeposit.TextMatrix(Row, 3) = Format(vsfDeposit.TextMatrix(Row, 3), "0.00")
    End If
    Call RecalCash
    If Col = 4 And Val(vsfDeposit.Cell(flexcpData, Row, Col)) = 0 Then
        mstrForceNote = ""
        For i = 1 To vsfDeposit.Rows - 1
            If Abs(vsfDeposit.TextMatrix(i, vsfDeposit.ColIndex("����"))) = 1 Then
                mstrForceNote = mstrForceNote & IIf(mstrForceNote = "", mstrǿ�����ֲ���Ա & "ǿ������:", ";") & vsfDeposit.TextMatrix(i, 0) & "," & Format(vsfDeposit.TextMatrix(i, 3), "0.00") & "Ԫ"
            End If
        Next i
    End If
End Sub

Private Sub RecalCash()
    '�����ֽ���
    Dim i As Integer, dblSum As Double
    Dim dbl��� As Double, dblʵ�� As Double
    Dim cur��� As Currency
    dblSum = 0
    For i = 1 To vsfDeposit.Rows - 1
        If Abs(vsfDeposit.TextMatrix(i, 4)) = 1 Then
            dblSum = dblSum + Val(vsfDeposit.TextMatrix(i, 3))
        End If
    Next i
    If dblSum = 0 Then
        txtMoney.Visible = False
        lblMoney.Visible = False
        lbl���.Visible = False
        lbl�����.Visible = False
    Else
        txtMoney.Visible = True
        lblMoney.Visible = True
        dblʵ�� = CentMoney(dblSum)
        cur��� = Val(dblSum) - Val(dblʵ��)
        txtMoney.Text = Format(dblʵ��, "0.00")
        If cur��� <> 0 Then
            lbl���.Visible = True
            lbl�����.Visible = True
            lbl�����.Caption = Format(cur���, "0.######")
        Else
            lbl���.Visible = False
            lbl�����.Visible = False
        End If
    End If
End Sub

Private Sub vsfDeposit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 3 And Col <> 4 Then
        Cancel = True
    End If
End Sub

Private Function GetThreeDeposit(lng����ID As Long, _
    Optional blnDateMoved As Boolean, Optional strTime As String, _
    Optional ByVal strPepositDate As String = "", _
    Optional intԤ����� As Integer = 0) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ʣ��Ԥ������ϸ(������)
    '���:strTime-סԺ����,��:1,2,3
    '        bln����תסԺ-�Ƿ��������תסԺ(ֻ�ܳ�ָ����Ԥ��)
    '        strPepositDate-ָ����Ԥ������
    '       intԤ�����-0-�����סԺ;1-����;2- סԺ
    '����:
    '����:Ԥ����ϸ����
    '����:������
    '����:2016-2-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strSub1 As String
    Dim strWherePage As String, strTable As String
    Dim strWhere As String, strDate As String
    On Error GoTo errH
    
    If intԤ����� = 1 Then strTime = ""    '69500
    
    strWherePage = IIf(strTime = "", "", " And instr(','||[2]||',',','||Nvl(A.��ҳID,0)||',')>0")
    strTable = IIf(blnDateMoved, zlGetFullFieldsTable("����Ԥ����¼"), "����Ԥ����¼ A")
    strWhere = "": strDate = "1974-04-28 00:00:00"
    If strPepositDate <> "" Then
        If IsDate(strPepositDate) Then
            strDate = strPepositDate
            strWhere = " And A.�տ�ʱ��=[3]"
        End If
    End If
    If intԤ����� <> 0 Then
        strWhere = strWhere & " And A.Ԥ����� =[4]"
    End If
    
    '���Ӳ�ѯ��������Ԥ�����շѼ��˷�ʱ��һ��һ��,ע��ϵͳ�������ʵ�Ԥ�������Ԥ���˷�,��Ҫ���ϼ�¼״̬�ж�
    strSub1 = _
        "   Select NO,Sum(Nvl(A.���,0)) as ���  " & _
        "    From " & strTable & _
        "   Where A.����ID Is Null And Nvl(A.���, 0)<>0 And A.����ID=[1]" & _
        "   Group by NO Having Sum(Nvl(A.���,0))<>0"
    
    '����=5:���۷�
    strSql = _
        " Select A.ID,A.��¼״̬,A.ʵ��Ʊ�� As Ʊ�ݺ�,A.NO,A.�տ�ʱ�� as ����, " & _
        "       A.���㷽ʽ,Nvl(A.���,0) as ���,A.�����ID,A.ID As Ԥ��ID" & _
        " From " & strTable & " ,(" & strSub1 & ") B" & _
        " Where A.����ID Is Null And Nvl(A.���,0)<>0" & _
        "           And A.���㷽ʽ Not IN(Select ���� From ���㷽ʽ Where ����=5)" & _
        "           And A.NO=B.NO And A.����ID=[1] " & strWherePage & strWhere & _
        " Union All" & _
        " Select 0 as ID,��¼״̬,Min(ʵ��Ʊ��) As Ʊ�ݺ�,NO,Min(�տ�ʱ��) as ����, " & _
        "        ���㷽ʽ,Sum(Nvl(���,0)-Nvl(��Ԥ��,0)) as ���,min(�����ID) as �����ID,Min(ID) As Ԥ��ID" & _
        " From " & strTable & _
        " Where ��¼���� IN(1,11) And ����ID is Not NULL And Nvl(���,0)<>Nvl(��Ԥ��,0)  " & _
        "       And ����ID=[1]" & strWherePage & strWhere & _
        " Having Sum(Nvl(���,0)-Nvl(��Ԥ��,0))<>0" & _
        " Group by ��¼״̬,NO,���㷽ʽ"
        
    strSql = "" & _
    "   Select Max(A.ID) As ID,Max(A.��¼״̬) As ��¼״̬,Max(A.Ʊ�ݺ�) As Ʊ�ݺ�,A.NO,Max(A.����) As ����,A.���㷽ʽ,Sum(A.���) As ���, " & _
    "           A.�����ID,decode(nvl(B.�Ƿ�ת�ʼ�����,0),0,0,1) as ����,Min(A.Ԥ��ID) As Ԥ��ID,Nvl(B.�Ƿ�����,0) As ����,B.���� As ���㿨����" & _
    "   From (" & strSql & ") A,ҽ�ƿ���� B" & _
    "   Where A.�����ID=B.ID(+) And A.�����ID Is Not Null" & _
    "   Group By a.No, a.���㷽ʽ, a.�����id, Decode(Nvl(b.�Ƿ�ת�ʼ�����, 0), 0, 0, 1),Nvl(B.�Ƿ�����,0),B.����" & vbNewLine & _
    "   Having Sum(a.���) <> 0" & _
    "   Order by A.�����ID Desc,A.NO,A.���㷽ʽ"
    
    '��Ҫ������֧��������,���۱�־Ϊ1,��������10.35�汾��֧��(��ʹ��֧�����ɵ�Ԥ������)
    Set GetThreeDeposit = zlDatabase.OpenSQLRecord(strSql, "mdlInExse", lng����ID, strTime, strDate, intԤ�����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsfDeposit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblMoney As Double, lngRow As Long
    Dim str����Ա���� As String, strDBUser As String
    Dim strPrivs As String
    
    If Col = 4 Then
        If Val(vsfDeposit.Cell(flexcpData, Row, Col)) = 0 Then
            If InStr(";" & mstrCardPrivs & ";", ";�����˿�ǿ������;") = 0 Then
                If mstrǿ�����ֲ���Ա = "" Then
                    mstrǿ�����ֲ���Ա = zlDatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
                    If mstrǿ�����ֲ���Ա = "" Then
                        MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣�", vbInformation, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Else
                If mstrǿ�����ֲ���Ա = "" Then
                    If MsgBox("ѡ��Ľ��㿨��֧������,�Ƿ�ǿ�ƽ������֣�", _
                                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Cancel = True: Exit Sub
                    mstrǿ�����ֲ���Ա = UserInfo.����
                End If
            End If
        End If
    End If
End Sub
