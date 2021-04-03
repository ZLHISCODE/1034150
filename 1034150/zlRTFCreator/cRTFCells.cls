VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFCells"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRTFCell"
Attribute VB_Ext_KEY = "Member0" ,"cRTFCell"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function LastCell() As cRTFCell
Attribute LastCell.VB_Description = "��ȡ���һ����Ԫ��Ԫ�ء�"
    Set LastCell = mCol(mCol.Count)
End Function

Public Function Add(Row As Long, Col As Long, _
Optional DataType As CellContentTypeEnum = cct0_�ı�, _
Optional Right As Long, _
Optional Width As Long = 0, _
Optional WidthUnit As CellLengthUnitEnum = clt0_Ĭ��, _
Optional sKey As String, _
Optional AlignmentV As AlignVTypeEnum = alv0_Ĭ�϶���, _
Optional Alignment As AlignTypeEnum = alm0_Ĭ�϶���, Optional BorderLeft As Boolean = True, _
Optional BorderRight As Boolean = True, Optional BorderTop As Boolean = True, _
Optional BorderBottom As Boolean = True, Optional BorderLeftDown As Boolean = False, _
Optional BorderLeftUp As Boolean = False, Optional BorderStyle As BorderStyleEnum = cbs00_Ĭ�ϱ߿���ʽ, _
Optional BorderColor As ColorEnum = clr00_Ĭ��, Optional CellPatType As PatTypeEnum = ptt00_��, _
Optional CellPatBackColor As ColorEnum = clr00_Ĭ��, Optional CellPatForeColor As ColorEnum = clr00_Ĭ��, _
Optional MergeVStart As Boolean = False, Optional MergeV As Boolean = False, _
Optional MergeStart As Boolean = False, Optional Merge As Boolean = False, _
Optional FitText As Boolean = True, Optional NoWrap As Boolean = False, _
Optional Nested As Boolean = False, Optional BorderWidth As Long = 0, _
Optional CellPatPercentage As Long = 100) As Long
Attribute Add.VB_Description = "����һ����Ԫ��Ԫ�ء�"
    '�����¶���
    Dim objNewMember As cRTFCell
    Set objNewMember = New cRTFCell


    '���ô��뷽��������
    objNewMember.AlignmentV = AlignmentV
    objNewMember.Alignment = Alignment
    objNewMember.BorderLeft = BorderLeft
    objNewMember.BorderRight = BorderRight
    objNewMember.BorderTop = BorderTop
    objNewMember.BorderBottom = BorderBottom
    objNewMember.BorderLeftDown = BorderLeftDown
    objNewMember.BorderLeftUp = BorderLeftUp
    objNewMember.BorderStyle = BorderStyle
    objNewMember.BorderColor = BorderColor
    objNewMember.CellPatType = CellPatType
    objNewMember.CellPatBackColor = CellPatBackColor
    objNewMember.CellPatForeColor = CellPatForeColor
    objNewMember.Width = Width
    objNewMember.WidthUnit = WidthUnit
    objNewMember.Right = Right
    objNewMember.MergeVStart = MergeVStart
    objNewMember.MergeV = MergeV
    objNewMember.MergeStart = MergeStart
    objNewMember.Merge = Merge
    objNewMember.Row = Row
    objNewMember.Col = Col
    objNewMember.DataType = DataType
    objNewMember.FitText = FitText
    objNewMember.NoWrap = NoWrap
    objNewMember.Nested = Nested
    objNewMember.BorderWidth = BorderWidth
    objNewMember.CellPatPercentage = CellPatPercentage
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '�����Ѵ����Ķ���
    Add = mCol.Count
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cRTFCell
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "ɾ��һ����Ԫ��Ԫ�ء�"
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "������������ For...Each �﷨ö�ٸü��ϡ�"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub
