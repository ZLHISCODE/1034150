VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#���ܣ�VSFlexGrid��MSHFlexGrid��صĴ���?�жϵȹ�������
'�������򣺺�����������Ϊͨ���׶���Ӣ�ļ��ɣ�VSFlexGrid�������ü�VSF�ؼ���,�������ͱ���ؼ���Ҫ������ǰ׺
'�ϼ��ࣺclsControl
'##############################################################################
'����ռ�������Ŀɼ��ԣ�VSFHScrollVisible,VSFVScrollVisible
Private Const SB_HORZ = &H0
Private Const SB_VERT = &H1
Private Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long

Public Sub SetColVisible(ByVal frmMain As Object, ByVal strTittle As String, vsGrid As Object, _
                    Optional ByVal WinLeft As Double = 0, Optional ByVal WinTop As Double = 0, _
                    Optional ByVal lngTxtHeight As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:VsGrid��������(��ʾ������)
    '���:frmMain-������
    '     strTittle-�б���������
    '     VsGrid��������(��ʾ������):
    '           ColData�洢��:����������,��ʽΪ:������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
    '     WinLeft-WinTop:��߾�;���߾�
    '     lngTxtHeight-�ı���߶�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Call frmVsColSel.ShowColSet(frmMain, strTittle, vsGrid, WinLeft, WinTop, lngTxtHeight)
    Exit Sub
Errhand:
     If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Public Sub SetFontSize(ByRef vsInput As Object, ByVal lngFontSize As Long, Optional ByVal lngCol As Long)
'���ܣ�����vsfflexgrid����Ĵ�С�����Զ������п����и�
'������lngFontSize�������ã�9�ż�С�壬12�ż�С��
'      lngCol,���Ҫ�����Զ������иߣ������и����ݵ��У�Ҫ��AutoSizeMode = flexAutoSizeRowHeight,WordWrap =True
    Call gobjComLib.zlControl.VSFSetFontSize(vsInput, lngFontSize, lngCol)
End Sub

Public Function IsLastVisible(ByRef vsInput As Object, ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����һ�ɼ���
    Dim i As Long
    
    With vsInput
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            IsLastVisible = lngRow = i
        End If
    End With
End Function

Public Sub Init(ByRef vsInput As Object, ByVal strHead As String, Optional ByVal strRowContent As String, Optional ByVal intFixedCols As Integer, Optional ByVal intFixedRows As Integer = 1)
'���ܣ���ʼ���������ݣ����ڴ�����Ի����ûָ�֮ǰ
'������vsInput=Ҫ���ø�ʽ�ı���
'          strHead=������и�ʽ����ʽΪ���б���1,�п�1,����1,��������1,��ʽ��1,����1;�б���2,�п�2,����2,��������2,��ʽ��2,����2.....
'          strRowContent=�����Ԥ����������,��ʽΪ����1,����1,��2,����2:��1;��1,����1,��2,����2:��2;
'˵����һ��ColKey������������Ҫ�������������б����ӿո񲢼��ϵ�����colKey,���ʽ�������� Fee,200,1��ʾ�������е�ColKey����ΪFee
'
    Dim i As Integer, lngRow As Long, j As Long
    Dim arrHead As Variant, arrCol As Variant, arrRow As Variant
    Dim arrTmp As Variant
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    '������
    With vsInput
        If strHead <> "" Then
            arrHead = Split(strHead, ";")
            .Clear: .Cols = 0: .Rows = 0
            .Rows = intFixedRows + 1: .Cols = UBound(arrHead) + 1
            .FixedRows = intFixedRows: .FixedCols = intFixedCols
            For i = LBound(arrHead) To UBound(arrHead)
                arrCol = Split(arrHead(i), ",")
                .FixedAlignment(i) = 4
                If intFixedRows <> 0 Then
                    If Not arrCol(0) Like "* ��*" Then
                        arrCol(0) = arrCol(0) & " " & arrCol(0)
                    End If
                    arrTmp = Split(arrCol(0), " ")
                    .TextMatrix(0, i) = arrTmp(0)
                    .ColKey(i) = arrTmp(1)
                End If
                
                If UBound(arrCol) > 0 Then
                    .ColWidth(i) = Val(arrCol(1))
                Else
                    .ColHidden(i) = True
                End If
                If UBound(arrCol) > 1 Then .ColAlignment(i) = Val(arrCol(2))
                If UBound(arrCol) > 2 Then .ColDataType(i) = Val(arrCol(3))
                If UBound(arrCol) > 3 Then .ColFormat(i) = arrCol(4)
                If UBound(arrCol) > 4 Then .ColHidden(i) = Val(arrCol(5))
            Next
        End If
        '���ý�����
        If strRowContent <> "" Then
            .Rows = .FixedRows
            lngRow = .FixedRows - 1: arrRow = Split(strRowContent, ";")
            For i = LBound(arrRow) To UBound(arrRow)
                arrTmp = Split(arrRow(i), ":")
                'ȷ���к�
                lngRow = lngRow + 1 'û���кŵĴ���һ�е��к��ϵ���
                If UBound(arrTmp) > 0 Then lngRow = Val(arrTmp(1))
                If .Rows - 1 < lngRow Then .Rows = lngRow + 1 '�����кŴ��ڵ�ǰ���ڵ��У����Զ�������
                '��������
                arrCol = Split(arrTmp(0), ",")
                For j = LBound(arrCol) To UBound(arrCol) Step 2
                    .TextMatrix(lngRow, Val(arrCol(j))) = arrCol(j + 1)
                Next
            Next
        End If
    End With
    Exit Sub
errH:
    Debug.Print Err.Source & "-Init:" & Err.Description
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Function HScrollVisible(vsInput As Object) As Boolean
'�ж�ˮƽ�������Ŀɼ���
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    HScrollVisible = False
    i = GetScrollRange(vsInput.hWnd, SB_HORZ, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then HScrollVisible = True
End Function

Public Function VScrollVisible(vsInput As Object) As Boolean
'�жϴ�ֱ�������Ŀɼ���
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    VScrollVisible = False
    i = GetScrollRange(vsInput.hWnd, SB_VERT, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then VScrollVisible = True
End Function

Public Function CopyTo(ByRef vsfSource As Object, Optional ByRef vsfInput As Object) As Object
'����: ��vsfSource�Ŀɼ����е����ݸ��Ƶ�vsfInput����һ������������ , ����Excel���
'����:
'     vsfInput-���ƺ�Ķ���
'     vsfSource-���Ƶ��Ķ���
'     strMsg -��ʾ��Ϣ
'���أ�һ�����ƶ��󣬸ú������Խ����ݸ��Ƶ��ƶ�����Ҳ�����Զ�����һ������
    Dim i As Long
    Dim j As Long
    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngTmp As Long
    Dim vsfReturn As VSFlexGrid
    Dim objTmp As frmPubTmp
    
    On Error GoTo errH:
    If vsfInput Is Nothing Then
        Set objTmp = New frmPubTmp
        Set vsfReturn = objTmp.vsfTmp
    End If
    With IIf(vsfInput Is Nothing, vsfReturn, vsfInput)
        .Rows = 0: .Cols = 0
        .Rows = vsfSource.Rows
        .FixedRows = vsfSource.FixedRows
        
        '����
        lngCol = 0
        For i = 0 To vsfSource.Cols - 1 '��
            If Not vsfSource.ColHidden(i) Then
                .Cols = .Cols + 1
                .ColWidth(lngCol) = vsfSource.ColWidth(i)
                lngRow = 0: lngTmp = 0
                
                For j = 0 To vsfSource.Rows - 1 '��
                    If Not vsfSource.RowHidden(j) Then
                        .ColAlignment(i) = vsfSource.ColAlignment(i)
                        .Cell(flexcpAlignment, lngRow, lngCol) = vsfSource.Cell(flexcpAlignment, j, i)  '���뷽ʽ
                        .TextMatrix(lngRow, lngCol) = vsfSource.TextMatrix(j, i)
                        lngRow = lngRow + 1
                    Else
                        lngTmp = lngTmp + 1  '��¼������
                    End If
                Next
                lngCol = lngCol + 1
            End If
        Next
        '
        .Rows = .Rows - lngTmp 'ɾ��������
        .FixedCols = vsfSource.FixedCols
        .RowHeight(-1) = vsfSource.RowHeightMin
    End With
    Set CopyTo = IIf(vsfInput Is Nothing, vsfReturn, vsfInput)
    Exit Function
errH:
    Set CopyTo = Nothing
    Err.Clear
    If 1 = 2 Then
        Resume
    End If
End Function

Public Sub CheckKeyPress(ByVal vsInput As Object, ByRef lngRow As Long, ByRef lngCol As Long, ByRef intKeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '����:ֻ���������ֺͻس����˸�
    '����:
    '   objctl:Vsgrid8.0�ؼ�
    '   intKeyascii:
    '           Keyascii:8 (�˸�)
    '   Row-��ǰ��
    '   Col-��ǰ��
    '   TextType:(0-�ı�ʽ;1-����ʽ;2-���ʽ)
    '����:һ��KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    
    If TextType = m�ı�ʽ Then
        If intKeyAscii = Asc("'") Then
            intKeyAscii = 0
        End If
        Exit Sub
    End If

    If intKeyAscii < Asc("0") Or intKeyAscii > Asc("9") Then
        Select Case intKeyAscii
        Case vbKeyReturn       '�س�
        
        Case 8                 '�˸�
        
        Case Asc(".")
            If TextType = m���ʽ Or TextType = m�����ʽ Then
                If InStr(vsInput.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                    intKeyAscii = 0
                End If
            Else
                intKeyAscii = 0
            End If
        Case Asc("-")          '����
            Dim iRow As Long
            Dim icol As Long
            If Trim(vsInput.EditText) = "" Then Exit Sub
            If TextType <> m�����ʽ Then intKeyAscii = 0: Exit Sub
            If vsInput.EditSelStart <> 0 Then intKeyAscii = 0: Exit Sub      '��겻���һλ,�������븺��
            If InStr(1, vsInput.EditText, "-") <> 0 Then   'ֻ�ܴ���һ������
                intKeyAscii = 0
            End If
        Case Else
            intKeyAscii = 0
        End Select
    End If
    Exit Sub
Errhand:
    Err.Clear
    intKeyAscii = 0
End Sub

Public Function BoolVal(ByVal vsGrid As Object, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '����:��ȡbool�е�ֵ
    '����:�Ǹõ�Ԫ��Ϊtrue,����true,���򷵻�False
    '------------------------------------------------------------------------------
    Dim strTemp As String
    Err = 0: On Error GoTo errH:
    With vsGrid
        strTemp = .TextMatrix(lngRow, lngCol)
    End With
    If UCase(strTemp) = UCase("True") Then
        BoolVal = True: Exit Function
    End If
    BoolVal = Val(strTemp) <> 0
    Exit Function
errH:
    Err.Clear
End Function

Public Sub BandRec(objGrid As Object, rsData As ADODB.Recordset, Optional ByVal blnBandMark As Boolean)
'���ܣ������ݰ󶨵�һ������ؼ���
'������blnBandMark=�Ƿ����ǩ����ǩ����RowData��
    Dim blnPre As Boolean, i As Long, j As Long
    
    objGrid.Clear: objGrid.Rows = 2: objGrid.Cols = 2
    objGrid.FixedRows = 1: objGrid.FixedCols = 0
    
    If rsData Is Nothing Then Exit Sub
    If rsData.State = adStateClosed Then Exit Sub
    
    blnPre = objGrid.Redraw
    objGrid.Redraw = False
    
    objGrid.Cols = rsData.Fields.count
    objGrid.Rows = IIf(rsData.RecordCount = 0, 2, rsData.RecordCount + 1)
    objGrid.FixedRows = 1
    
    For j = 0 To rsData.Fields.count - 1
        objGrid.TextMatrix(0, j) = rsData.Fields(j).Name
        If UCase(TypeName(objGrid)) <> UCase("MSHFlexGrid") Then objGrid.ColKey(j) = rsData.Fields(j).Name
    Next
    
    If rsData.RecordCount = 0 Then objGrid.Redraw = blnPre: Exit Sub
    
    rsData.MoveFirst
    For i = 1 To rsData.RecordCount
        For j = 0 To rsData.Fields.count - 1
            objGrid.TextMatrix(i, j) = "" & rsData.Fields(j).value
        Next
        If blnBandMark Then objGrid.RowData(i) = rsData.Bookmark '��¼��ǩ
        rsData.MoveNext
    Next
End Sub

Public Sub AdjustCols(ByRef vsTmp As Object, Optional ByVal strOtherCol As String)
'���ܣ����������п����ı�����������Ҫ�����п����е������ú������Ǵ������������
' ������vsTmp=�������
'          strOtherCol=��Ӧ������仯����
'          lngBaseFontSize=ԭʼ�����С
    Dim lngFixed As Long
    Dim dbAdd As Double, dbMinus As Double
    Dim arrAdd As Variant, arrMinus As Variant
    Dim lngCount As Long, lngTmp As Long
    Dim i As Long
    Dim lngLen As Long

    On Error GoTo errH
    lngFixed = gobjComLib.zlControl.OneCharWidth(vsTmp.Font) '��ȡһ���ַ���
    arrMinus = Array(): arrAdd = Array()
    With vsTmp
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) And InStr(strOtherCol, "," & i & ",") = 0 Then
                lngLen = gobjComLib.zlStr.ActualLen(.TextMatrix(0, i))
                If lngLen <= 2 Then lngLen = 2
                lngLen = lngLen + 1 '���ӱ߽����
                If .ColWidth(i) < lngLen * lngFixed Then  'С���������������С���ȣ�����Ҫ�ӿ�
                    ReDim Preserve arrAdd(UBound(arrAdd) + 1)
                    arrAdd(UBound(arrAdd)) = i
                    dbAdd = dbAdd + lngLen * lngFixed - .ColWidth(i) '��¼��Ҫ����������
                Else '�����������������С���ȣ�����Ҫ���ٿ���
                    ReDim Preserve arrMinus(UBound(arrMinus) + 1)
                    arrMinus(UBound(arrMinus)) = i
                    dbMinus = dbMinus + .ColWidth(i) - lngLen * lngFixed '��¼��Ҫ���ٵ�����
                End If
            End If
        Next
        If UBound(arrAdd) = -1 Then Exit Sub 'û�п��Լӿ����У���������
        If UBound(arrMinus) = -1 Then Exit Sub 'û�п��Լ��ٿ��ȵ��У���������
        For i = LBound(arrAdd) To UBound(arrAdd)
            lngLen = gobjComLib.zlStr.ActualLen(.TextMatrix(0, arrAdd(i)))
            If lngLen <= 2 Then lngLen = 2 '��С����Ӣ���ַ�����
            lngLen = lngLen + 1 '���ӱ߽����
            .ColWidth(arrAdd(i)) = lngLen * lngFixed '�ӿ���������Ϊ��С����
        Next
        dbAdd = IIf(dbAdd > dbMinus, dbMinus, dbAdd) '����ʵ�ʼӿ�������
        For i = LBound(arrMinus) To UBound(arrMinus)
            lngLen = gobjComLib.zlStr.ActualLen(.TextMatrix(0, arrMinus(i)))
            If lngLen <= 2 Then lngLen = 2 '��С����Ӣ���ַ�����
            lngLen = lngLen + 1 '���ӱ߽����
            '����ƽ̯��ÿһ���ɼ��ٿ��ȵ��У���ĳһ�м��ٺ��С����С���ȣ�������Ϊ��С����
            If .ColWidth(arrMinus(i)) - dbAdd / (UBound(arrMinus) - i + 1) < lngLen * lngFixed Then
                dbAdd = dbAdd - (lngLen * lngFixed - .ColWidth(arrMinus(i))) '����������ȥ���ٿ���
                .ColWidth(arrMinus(i)) = lngLen * lngFixed  '����Ϊ��С����
            Else
                dbAdd = dbAdd - dbAdd / (UBound(arrMinus) - i + 1) '����������ȥƽ̯��
                .ColWidth(arrMinus(i)) = .ColWidth(arrMinus(i)) - dbAdd / (UBound(arrMinus) - i + 1) '����Ϊ��ǰ���ȼ�ȥƽ̯����
            End If
        Next
    End With
    Exit Sub
errH:
    Debug.Print Err.Source & "-AdjustCols:" & Err.Description
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Sub MshSetFormat(ByVal mshGrd As Object, ByVal strColumns As String, ByVal strCaption As String, Optional ByVal rsData As ADODB.Recordset, Optional ByVal blnData As Boolean = False, Optional ByVal blnRestore As Boolean = False)
    '���ܣ����ñ���ؼ�����ͷ��������ݵ�����ؼ�
    '���������壺msfObj�Ǳ������strColumns��������ʽ��������|����|�ж��뷽ʽ����blnData˵�������ø�ʽ���ǰ����ݣ�blnRestore˵���Ƿ�ָ�ԭʼ����
    '������ʽ��NO|1000|4,����|1600|4,��鵥λ|2000|4,�����|800|4,��Ҫ��¼|2000|4,����|1000|4
    Call gobjComLib.zlControl.MshSetFormat(mshGrd, strColumns, strCaption, rsData, blnData, blnRestore)
End Sub

Public Sub MshSetColWidth(mshGrd As Object, frmParent As Object)
'���ܣ��Զ�����MSHFlexGrid����ĸ��п���,����Сƥ��Ϊ׼,��������ͷ���ֿ���ȫ����ʾ
'������mshGrd=Ҫ�����п���MSHFlexGrid�ؼ�
'      frmParent=������,���ڼ������ֿ���
    Call gobjComLib.zlControl.MshSetColWidth(mshGrd, frmParent)
End Sub

Public Function MshGrdFindRow(ByVal mshGrd As Object, ByVal strTmp As Variant, ByVal intCol As Integer) As Integer
'����:Ѱ����ĳһ��Ԫֵ��ȵ���
'����:
'       mshGrd-����ؼ�
'       strTmp-ƥ�䴮
'       intCol-ƥ����
'����:���ҵ�����,��δ�ҵ�,��λ��һ��
    Call gobjComLib.zlControl.MshGrdFindRow(mshGrd, strTmp, intCol)
End Function

Public Sub MshColumnSort(ByVal mshGrd As Object, ByVal IntSorCol As Integer, ByRef intPreCol As Integer, ByRef intPreSort As Integer, Optional blnNum As Boolean = False, Optional intMachingCol As Integer = 0)
'����:���н�������
'����:
'       mshGrd-MSHFlexGrid����ؼ�
'       intPreCol-ǰһ�ε���ͷ��������
'       intPreSort-ǰһ�ε���ͷ������
'       blnNum-�Ƿ�Ϊ������
'       intMachingCol-ƥ����,��Ҫ������������¶�λ��ָ������
    Call gobjComLib.zlControl.MshColumnSort(mshGrd, IntSorCol, intPreCol, intPreSort, blnNum, intMachingCol)
End Sub

Public Sub MshClearData(mshGrd As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '����: ���ָ�������ָ����ָ���е�����
    '����: mshGrd=Ҫ����������ؼ�
    '      intRow=Ҫ������к�
    '      intCol=Ҫ������к��б���Array(1,2,3),������������Ա�ʾΪArray()
    '���أ�
    Call gobjComLib.zlControl.MshClearData(mshGrd, intRow, intCol)
End Sub
