VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFParagraph"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarIsInTable As Boolean
Private mvarSpaceBefore As Long
Private mvarSpaceAfter As Long
Private mvarSpaceLine As Long
Private mvarSpaceFirstLine As Long
Private mvarTexts As cRTFTexts
Private mvarCount As Long
Private mvarFontSizeDefault As Long

Public Property Let FontSizeDefault(ByVal vData As Long)
    Dim i As Long
    For i = 1 To mvarTexts.Count
        mvarTexts(i).FontSizeDefault = vData
    Next
    mvarFontSizeDefault = vData
End Property

Public Function GetTextRTF() As String
    Dim i As Long, strR As String
    '���һ���ı���Ҫ�Զ����Ӷ�����
    mvarTexts(mvarTexts.Count).Text = mvarTexts(mvarTexts.Count).Text + vbCrLf
    For i = 1 To mvarTexts.Count
        If mvarSpaceAfter <> 0 Then
            If mvarTexts(i).SpaceAfter = 0 Then mvarTexts(i).SpaceAfter = mvarSpaceAfter
        End If
        If mvarSpaceBefore <> 0 Then
            If mvarTexts(i).SpaceBefore = 0 Then mvarTexts(i).SpaceBefore = mvarSpaceBefore
        End If
        If mvarSpaceLine <> 0 Then
            If mvarTexts(i).SpaceLine = 0 Then mvarTexts(i).SpaceLine = mvarSpaceLine
        End If
        If mvarSpaceFirstLine <> 0 Then
            If mvarTexts(i).SpaceFirstLine = 0 Then mvarTexts(i).SpaceFirstLine = mvarSpaceFirstLine
        End If
        If mvarFontSizeDefault <> 0 Then
            If mvarTexts(i).FontSizeDefault = 0 Then mvarTexts(i).FontSizeDefault = mvarFontSizeDefault
        End If

        strR = strR & mvarTexts(i).GetTextRTF
    Next
    GetTextRTF = strR
End Function

Public Property Get FontSizeDefault() As Long
    FontSizeDefault = mvarFontSizeDefault
End Property

Public Property Let Count(ByVal vData As Long)
Attribute Count.VB_Description = "������Ŀ��"
    mvarCount = vData
End Property

Public Property Get Count() As Long
    Count = mvarCount
End Property

Public Property Set Texts(ByVal vData As cRTFTexts)
    Set mvarTexts = vData
End Property

Public Property Get Texts() As cRTFTexts
    Set Texts = mvarTexts
End Property

Public Function LastText() As cRTFText
    Set LastText = mvarTexts(mvarTexts.Count)
End Function

Public Function AddText(Optional Text As String, Optional sKey As String, _
Optional Alignment As AlignTypeEnum = alm0_Ĭ�϶���, _
Optional Bold As Boolean = False, _
Optional Italic As Boolean = False, Optional ForeColor As ColorEnum = clr00_Ĭ��, _
Optional FontSize As FontSizeEnum = fs00_Ĭ��, Optional FontStyle As FontStyleEnum = ft0_����, _
Optional CharStyle As CharStyleEnum = csĬ��, Optional AnimType As AnimTypeEnum = ani0_��, _
Optional ScaleX As Long = 100, Optional Embo As EmboEnum = emb0_��, _
Optional Expand As Long = 0, Optional OutLine As Boolean = False, _
Optional Shadow As Boolean = False, Optional Strike As Boolean = False, _
Optional StrikeD As Boolean = False, Optional SubChar As Boolean = False, _
Optional SuperChar As Boolean = False, Optional UnderLineType As UnderLineTypeEnum = udl00_��, _
Optional UnderLineColor As ColorEnum = clr00_Ĭ��, Optional Hided As Boolean = False, _
Optional WebHidden As Boolean = False, Optional CharBorder As Boolean = False, _
Optional CharShadow As Long = 100, Optional CharPatForeColor As ColorEnum = clr00_Ĭ��, _
Optional CharPatBackColor As ColorEnum = clr00_Ĭ��, Optional PatType As PatTypeEnum = ptt00_��, _
Optional Highlight As Boolean = False, Optional HighlightColor As ColorEnum = clr00_Ĭ��, _
Optional ListType As ListTypeEnum = ltp00_��, _
Optional IsInTable As Boolean = False) As Long

    Dim i As Long

    i = mvarTexts.Add(Text, sKey, Alignment, Bold, Italic, ForeColor, FontSize, FontStyle, CharStyle, AnimType, _
        ScaleX, Embo, Expand, OutLine, Shadow, Strike, StrikeD, SubChar, SuperChar, UnderLineType, _
        UnderLineColor, Hided, WebHidden, CharBorder, CharShadow, CharPatForeColor, CharPatBackColor, _
        PatType, Highlight, HighlightColor, ListType, IsInTable)
    mvarCount = mvarTexts.Count
    AddText = i
End Function


Public Property Let SpaceFirstLine(ByVal vData As Long)
Attribute SpaceFirstLine.VB_Description = "��������ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceFirstLine = 5
    Dim i As Long
    For i = 1 To mvarTexts.Count
        mvarTexts(i).SpaceFirstLine = vData
    Next
    mvarSpaceFirstLine = vData
End Property

Public Property Get SpaceFirstLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceFirstLine
    SpaceFirstLine = mvarSpaceFirstLine
End Property

Public Property Let SpaceLine(ByVal vData As Long)
Attribute SpaceLine.VB_Description = "�м�ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceLine = 5
    Dim i As Long
    For i = 1 To mvarTexts.Count
        mvarTexts(i).SpaceLine = vData
    Next
    mvarSpaceLine = vData
End Property

Public Property Get SpaceLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceLine
    SpaceLine = mvarSpaceLine
End Property

Public Property Let SpaceAfter(ByVal vData As Long)
Attribute SpaceAfter.VB_Description = "��ĩ��ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceAfter = 5
    Dim i As Long
    For i = 1 To mvarTexts.Count
        mvarTexts(i).SpaceAfter = vData
    Next
    mvarSpaceAfter = vData
End Property

Public Property Get SpaceAfter() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceAfter
    SpaceAfter = mvarSpaceAfter
End Property

Public Property Let SpaceBefore(ByVal vData As Long)
Attribute SpaceBefore.VB_Description = "��ǰ��ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceBefore = 5
    Dim i As Long
    For i = 1 To mvarTexts.Count
        mvarTexts(i).SpaceBefore = vData
    Next
    mvarSpaceBefore = vData
End Property

Public Property Get SpaceBefore() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceBefore
    SpaceBefore = mvarSpaceBefore
End Property

Public Property Let IsInTable(ByVal vData As Boolean)
Attribute IsInTable.VB_Description = "�Ƿ��ڱ����С�Ĭ�ϣ���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.IsInTable = 5
    Dim i As Long
    For i = 1 To mvarTexts.Count
        mvarTexts(i).IsInTable = vData
    Next
    mvarIsInTable = vData
End Property

Public Property Get IsInTable() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.IsInTable
    IsInTable = mvarIsInTable
End Property

Private Sub Class_Initialize()
    Set mvarTexts = New cRTFTexts
End Sub

Private Sub Class_Terminate()
    Set mvarTexts = Nothing
End Sub