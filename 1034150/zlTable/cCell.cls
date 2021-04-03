VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'################################################################################################################
'## ö�ٳ���
'################################################################################################################

'ˮƽ���뷽ʽ
Public Enum HAlignEnum
    HALignLeft = DT_LEFT                '����
    HALignCentre = DT_CENTER            '����
    HALignRight = DT_RIGHT              '����
End Enum

'��ֱ���뷽ʽ
Public Enum VAlignEnum
    VALignTop = DT_TOP                  '����
    VALignVCentre = DT_VCENTER          '��ֱ����
    VALignBottom = DT_BOTTOM            '����
End Enum

'################################################################################################################
'## �ڲ�����
'################################################################################################################

'��ʽ
Private mvarKey As String               '�ؼ���
Private mvarRow As Long                 '��
Private mvarCol As Long                 '��
Private mvarMargin As Long              '��Ԫ��߾࣬Ĭ��Ϊ30
Private mvarSingleLine As Boolean       '�Ƿ�����ʾ��Ĭ��ΪFalse
Private mvarForeColor As OLE_COLOR      'ǰ��ɫ��Ĭ��Ϊ��ɫ
Private mvarBackColor As OLE_COLOR      '����ɫ��Ĭ��Ϊ��ɫ
Private mvarGridLineColor As OLE_COLOR  '�߿���ɫ��Ĭ��Ϊ��ɫ
Private mvarGridLineWidth As Long       '�߿���ȣ�Ĭ��Ϊ1
Private mvarFixedWidth As Boolean       '�Ƿ��������ȣ�Ĭ��ΪFalse
Private mvarAutoHeight As Boolean       '�Զ��߶ȣ�Ĭ��ΪTrue
Private mvarFontName As String          '��������
Private mvarFontSize As Long            '�����С����λ����ֵ
Private mvarFontBold As Boolean         '����
Private mvarFontItalic As Boolean       'б��
Private mvarFontStrikeout As Boolean    'ɾ����
Private mvarFontUnderline As Boolean    '�»���
Private mvarFontWeight As Long          '��������
Private mvarFormatString As String      '�ı���ʽ
Private mvarIndent As Long              '����ֵ
Private mvarHAlignment As HAlignEnum    'ˮƽ���뷽ʽ
Private mvarVAlignment As VAlignEnum    '��ֱ���뷽ʽ
Private mvarProtected As Boolean        '�Ƿ񱣻�
Private mvarToolTipText As String       '��ʾ�ı�
Private mvarPicture As StdPicture       '��Ԫ��ͼƬ

'����
Private mvarText As String              '�ı����ݣ�Ĭ��Ϊ��
Private mvarTag As String               '��չ���
Private mvarIcon As Long                'ͼ������ֵ��Ĭ��Ϊ-1����ʾû��ͼ��

'��������
Private mvarMergeInfo As String         '�ϲ��ַ������ֱ���"��1:��1:��2:��2"��ɣ���16λ��Ϊ�ձ�ʾû�кϲ�
Private mvarSelected As Boolean         '�Ƿ�ѡ��
Private mvarHot As Boolean              '�Ƿ��ȸ���
Private mvarVisibled As Boolean         '�Ƿ�ɼ�
Private mvarWidth As Long               '����
Private mvarHeight As Long              '�߶�
Private mvarDirty As Boolean            '�Ƿ��Ѿ����޸Ĺ�����Ҫ�ػ�

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let Row(ByVal vData As Long)
    mvarRow = vData
End Property

Public Property Get Row() As Long
    Row = mvarRow
End Property

Public Property Let Col(ByVal vData As Long)
    mvarCol = vData
End Property

Public Property Get Col() As Long
    Col = mvarCol
End Property

Public Property Let Margin(ByVal vData As Long)
    mvarMargin = vData
    mvarDirty = True
End Property

Public Property Get Margin() As Long
    Margin = mvarMargin
End Property

Public Property Let SingleLine(ByVal vData As Boolean)
    mvarSingleLine = vData
    mvarDirty = True
End Property

Public Property Get SingleLine() As Boolean
    SingleLine = mvarSingleLine
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
    mvarDirty = True
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
    mvarDirty = True
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let GridLineColor(ByVal vData As OLE_COLOR)
    mvarGridLineColor = vData
    mvarDirty = True
End Property

Public Property Get GridLineColor() As OLE_COLOR
    GridLineColor = mvarGridLineColor
End Property

Public Property Let GridLineWidth(ByVal vData As Long)
    mvarGridLineWidth = vData
    mvarDirty = True
End Property

Public Property Get GridLineWidth() As Long
    GridLineWidth = mvarGridLineWidth
End Property

Public Property Let FixedWidth(ByVal vData As Boolean)
    mvarFixedWidth = vData
    mvarDirty = True
End Property

Public Property Get FixedWidth() As Boolean
    FixedWidth = mvarFixedWidth
End Property

Public Property Let AutoHeight(ByVal vData As Boolean)
    mvarAutoHeight = vData
    mvarDirty = True
End Property

Public Property Get AutoHeight() As Boolean
    AutoHeight = mvarAutoHeight
End Property

Public Property Let FontName(ByVal vData As String)
    mvarFontName = vData
    mvarDirty = True
End Property

Public Property Get FontName() As String
    FontName = mvarFontName
End Property

Public Property Let FontSize(ByVal vData As Long)
    mvarFontSize = vData
    mvarDirty = True
End Property

Public Property Get FontSize() As Long
    FontSize = mvarFontSize
End Property

Public Property Let FontBold(ByVal vData As Boolean)
    mvarFontBold = vData
    mvarDirty = True
End Property

Public Property Get FontBold() As Boolean
    FontBold = mvarFontBold
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
    mvarFontItalic = vData
    mvarDirty = True
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mvarFontItalic
End Property

Public Property Let FontStrikeout(ByVal vData As Boolean)
    mvarFontStrikeout = vData
    mvarDirty = True
End Property

Public Property Get FontStrikeout() As Boolean
    FontStrikeout = mvarFontStrikeout
End Property

Public Property Let FontUnderline(ByVal vData As Boolean)
    mvarFontUnderline = vData
    mvarDirty = True
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = mvarFontUnderline
End Property

Public Property Let FontWeight(ByVal vData As Long)
    mvarFontWeight = vData
    mvarDirty = True
End Property

Public Property Get FontWeight() As Long
    FontWeight = mvarFontWeight
End Property

Public Property Let FormatString(ByVal vData As String)
    mvarFormatString = vData
    mvarDirty = True
End Property

Public Property Get FormatString() As String
    FormatString = mvarFormatString
End Property

Public Property Let Indent(ByVal vData As Long)
    mvarIndent = vData
    mvarDirty = True
End Property

Public Property Get Indent() As Long
    Indent = mvarIndent
End Property

Public Property Let HAlignment(ByVal vData As HAlignEnum)
    mvarHAlignment = vData
    mvarDirty = True
End Property

Public Property Get HAlignment() As HAlignEnum
    HAlignment = mvarHAlignment
End Property

Public Property Let VAlignment(ByVal vData As VAlignEnum)
    mvarVAlignment = vData
    mvarDirty = True
End Property

Public Property Get VAlignment() As VAlignEnum
    VAlignment = mvarVAlignment
End Property

Public Property Let Protected(ByVal vData As Boolean)
    mvarProtected = vData
    mvarDirty = True
End Property

Public Property Get Protected() As Boolean
    Protected = mvarProtected
End Property

Public Property Let ToolTipText(ByVal vData As String)
    mvarToolTipText = vData
    mvarDirty = True
End Property

Public Property Get ToolTipText() As String
    ToolTipText = mvarToolTipText
End Property

Public Property Set Picture(ByVal vData As StdPicture)
    Set mvarPicture = vData
    mvarDirty = True
End Property

Public Property Let Picture(ByVal vData As StdPicture)
    Set mvarPicture = vData
    mvarDirty = True
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mvarPicture
End Property

Public Property Let Text(ByVal vData As String)
    mvarText = vData
    mvarDirty = True
End Property

Public Property Get Text() As String
    Text = mvarText
End Property

Public Property Let Tag(ByVal vData As String)
    mvarTag = vData
    mvarDirty = True
End Property

Public Property Get Tag() As String
    Tag = mvarTag
End Property

Public Property Let Icon(ByVal vData As Long)
    mvarIcon = vData
    mvarDirty = True
End Property

Public Property Get Icon() As Long
    Icon = mvarIcon
End Property

Public Property Let MergeInfo(ByVal vData As String)
    mvarMergeInfo = vData
    mvarDirty = True
End Property

Public Property Get MergeInfo() As String
    MergeInfo = mvarMergeInfo
End Property

Public Property Let Selected(ByVal vData As Boolean)
    mvarSelected = vData
    mvarDirty = True
End Property

Public Property Get Selected() As Boolean
    Selected = mvarSelected
End Property

Public Property Let Hot(ByVal vData As Boolean)
    mvarHot = vData
    mvarDirty = True
End Property

Public Property Get Hot() As Boolean
    Hot = mvarHot
End Property

Public Property Let Visibled(ByVal vData As Boolean)
    mvarVisibled = vData
    mvarDirty = True
End Property

Public Property Get Visibled() As Boolean
    Visibled = mvarVisibled
End Property

Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
    mvarDirty = True
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
    mvarDirty = True
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let Dirty(ByVal vData As Boolean)
    mvarDirty = vData
End Property

Public Property Get Dirty() As Boolean
    Dirty = mvarDirty
End Property

Public Property Get MergeStartRow() As Long
    If mvarMergeInfo <> "" And Len(mvarMergeInfo) = 16 Then
        MergeStartRow = Val(Mid(mvarMergeInfo, 1, 4))
    End If
End Property

Public Property Get MergeStartCol() As Long
    If mvarMergeInfo <> "" And Len(mvarMergeInfo) = 16 Then
        MergeStartCol = Val(Mid(mvarMergeInfo, 5, 4))
    End If
End Property

Public Property Get MergeEndRow() As Long
    If mvarMergeInfo <> "" And Len(mvarMergeInfo) = 16 Then
        MergeEndRow = Val(Mid(mvarMergeInfo, 9, 4))
    End If
End Property

Public Property Get MergeEndCol() As Long
    If mvarMergeInfo <> "" And Len(mvarMergeInfo) = 16 Then
        MergeEndCol = Val(Mid(mvarMergeInfo, 13, 4))
    End If
End Property

'#########################################################################################################
'## ���ܣ�  ���ڶ����ı��������ı��߶�
'## ���أ�  ���ظ߶�����ֵ
'#########################################################################################################
Public Property Get EvaluateTextHeight(ByRef oTable As Table, _
    Optional ByVal sNewText As String = "", _
    Optional ByVal bCalculPicHeight As Boolean = True) As Long
    
    Dim tLF As LOGFONT
    Dim hFnt As Long
    Dim hFntOld As Long
    Dim tR As RECT
    Dim sFont As String
    Dim iChar As Integer
    Dim temp() As Byte
    Dim sText As String
    Dim lW As Long, lH As Long
    
    With tLF
        sFont = mvarFontName & vbNullString
        temp = StrConv(sFont, vbFromUnicode)
        For iChar = 0 To UBound(temp)
            .lfFaceName(iChar) = temp(iChar)
        Next iChar
        .lfItalic = mvarFontItalic
        .lfHeight = GetPixcelHeightByPoint(oTable.hDC, mvarFontSize)
        .lfUnderline = mvarFontUnderline
        .lfWeight = IIf(mvarFontBold, FW_BOLD, FW_NORMAL)
        .lfStrikeOut = mvarFontStrikeout
        .lfQuality = IIf(oTable.FontQuality < 0, 0, oTable.FontQuality)
        .lfCharSet = DEFAULT_CHARSET
    End With
    
    tR.Right = (mvarWidth - 2 * mvarMargin * p_TPPX - mvarGridLineWidth * p_TPPX - IIf(mvarIcon > 0, p_lIconWidth + 30, 0)) / p_TPPX
    
    sText = IIf(mvarFormatString = "", IIf(sNewText = "", mvarText, sNewText), Format(IIf(sNewText = "", mvarText, sNewText), mvarFormatString))
    If sText = "" Then
        If p_lIconHeight > 0 Then EvaluateTextHeight = p_lIconHeight + mvarMargin * p_TPPY * 2 + mvarGridLineWidth * p_TPPX
        If oTable.MinRowHeight > EvaluateTextHeight Then EvaluateTextHeight = oTable.MinRowHeight
        GoTo LL
    End If
    
    hFnt = CreateFontIndirect(tLF) '��������
    'ѡ��ָ�����壬���������
    hFntOld = SelectObject(oTable.hDC, hFnt)
    
    '�����ı��������ı����Զ�����ʱ������ʾʡ�Ժţ������ı����Զ��߶�ʱ������ʾʡ�Ժţ�
    DrawText oTable.hDC, sText, -1, tR, DT_NOPREFIX Or DT_EXPANDTABS Or DT_EDITCONTROL Or mvarVAlignment Or mvarHAlignment Or _
        IIf(mvarSingleLine, DT_SINGLELINE, DT_WORDBREAK) Or _
        IIf(oTable.WordEllipsis, DT_WORD_ELLIPSIS, 0) Or DT_CALCRECT

    EvaluateTextHeight = (tR.Bottom - tR.Top) * Screen.TwipsPerPixelY + mvarMargin * p_TPPY * 2 + 2 * mvarGridLineWidth * p_TPPX
    If EvaluateTextHeight < p_lIconHeight Then
        EvaluateTextHeight = p_lIconHeight + mvarMargin * p_TPPY * 2 + mvarGridLineWidth * p_TPPX
    End If
    If EvaluateTextHeight < oTable.MinRowHeight Then
        EvaluateTextHeight = oTable.MinRowHeight
    End If
   
    '�ָ�������
    SelectObject oTable.hDC, hFntOld
    '�ͷ���ʱ����
    DeleteObject hFnt
    
LL:
    If Not mvarPicture Is Nothing And bCalculPicHeight Then
        'ͼƬ��Ԫ�񣬸߶Ⱦ���ͼƬ���Ÿ߶�
        lW = (mvarWidth - 2 * mvarMargin * p_TPPX - mvarGridLineWidth * p_TPPX - IIf(mvarIcon > 0, p_lIconWidth + 30, 0))
        lH = (lW * mvarPicture.Height) / mvarPicture.Width + mvarMargin * p_TPPY * 2 + 2 * mvarGridLineWidth * p_TPPX
        If EvaluateTextHeight < lH Then EvaluateTextHeight = lH
    End If
End Property

'################################################################################################################
'## ���ܣ�  ���Ƶ�Ԫ�����
'##
'## ������  blnWithKey  :�Ƿ���ͬKeyֵһ���ƣ�
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cCell
    Dim cNode As New cCell
    If blnWithKey Then cNode.Key = Me.Key
    cNode.Row = Me.Row
    cNode.Col = Me.Col
    cNode.Margin = Me.Margin
    cNode.SingleLine = Me.SingleLine
    cNode.MergeInfo = Me.MergeInfo
    cNode.Selected = Me.Selected
    cNode.Hot = Me.Hot
    cNode.Visibled = Me.Visibled
    cNode.Width = Me.Width
    cNode.Height = Me.Height
    cNode.FixedWidth = Me.FixedWidth
    cNode.AutoHeight = Me.AutoHeight
    cNode.Icon = Me.Icon
    cNode.Text = Me.Text
    cNode.Tag = Me.Tag
    cNode.FormatString = Me.FormatString
    cNode.Indent = Me.Indent
    cNode.HAlignment = Me.HAlignment
    cNode.VAlignment = Me.VAlignment
    cNode.ForeColor = Me.ForeColor
    cNode.BackColor = Me.BackColor
    cNode.GridLineColor = Me.GridLineColor
    cNode.GridLineWidth = Me.GridLineWidth
    cNode.FontName = Me.FontName
    cNode.FontSize = Me.FontSize
    cNode.FontBold = Me.FontBold
    cNode.FontItalic = Me.FontItalic
    cNode.FontStrikeout = Me.FontStrikeout
    cNode.FontUnderline = Me.FontUnderline
    cNode.FontWeight = Me.FontWeight
    cNode.Protected = Me.Protected
    cNode.ToolTipText = Me.ToolTipText
    cNode.Dirty = Me.Dirty
    cNode.Picture = Me.Picture
    
    Set Clone = cNode
End Function

Public Sub GetCellTextBorder(ByRef lLeft As Long, ByRef lTop As Long, ByRef lRight As Long, ByRef lBottom As Long)
    lLeft = (ColInfo(mvarCol).LeftX + mvarGridLineWidth * p_TPPX + mvarMargin * p_TPPX + IIf(mvarIcon > 0, p_lIconWidth + 30, 0)) / p_TPPX
    lTop = (RowInfo(mvarRow).TopY + mvarGridLineWidth * p_TPPY + mvarMargin * p_TPPY) / p_TPPY
    lRight = (ColInfo(mvarCol).LeftX + mvarWidth - mvarMargin * p_TPPX) / p_TPPX
    lBottom = (RowInfo(mvarRow).TopY + mvarHeight - mvarMargin * p_TPPY) / p_TPPY
End Sub

Public Sub GetCellPictureBorder(ByRef lLeft As Long, ByRef lTop As Long, ByRef lWidth As Long, ByRef lHeight As Long)
    lLeft = (ColInfo(mvarCol).LeftX)
    lTop = (RowInfo(mvarRow).TopY)
    Dim lW As Long, lH As Long
    lW = mvarWidth + mvarGridLineWidth * p_TPPX - 1 * p_TPPX
    lH = (lW * mvarPicture.Height) / mvarPicture.Width
    If mvarAutoHeight = False Then
        '������Զ�����ģʽ��ͼƬ����
        lH = mvarHeight / p_TPPY + mvarGridLineWidth - 1
    End If
    lWidth = lW
    lHeight = lH
End Sub

'################################################################################################################
'## ���ܣ�  ���Ƶ�Ԫ��
'##
'## ������  oTable:     �������
'##         lhDC:       ���Ƶ�ָ����DC�����ʡ�ԣ���ʹ��Table��DC
'################################################################################################################
Public Sub DrawCell(ByRef oTable As Table, Optional ByRef lhDC As Long = 0)
    Dim tLF As LOGFONT
    Dim hFnt As Long
    Dim hBr As Long
    Dim hBr2 As Long
    Dim hFntOld As Long
    Dim tR As RECT
    Dim sFont As String
    Dim iChar As Integer
    Dim temp() As Byte
    Dim sText As String
    Dim tRCopy As RECT
    Dim hPen As Long
    Dim hPenOld As Long
    Dim tJ As POINTAPI
    Dim lLeft As Long, lTop As Long, lRight As Long, lBottom As Long
    Dim i As Long, j As Long
    Dim NoFocusHighlightBackColor As OLE_COLOR, NoFocusHighlightForeColor As OLE_COLOR
    Dim bExportMode As Boolean, LL As Long, lT As Long, lW As Long, lH As Long
    
    If mvarVisibled = False Then Exit Sub
    
    If lhDC = 0 Then
        bExportMode = False
        lhDC = oTable.hDC
    Else
        bExportMode = True
    End If
    
    On Error Resume Next
    tR.Left = (ColInfo(mvarCol).LeftX) / p_TPPX
    tR.Top = (RowInfo(mvarRow).TopY) / p_TPPY
    tR.Right = (ColInfo(mvarCol).LeftX + mvarWidth) / p_TPPX + mvarGridLineWidth - 1
    tR.Bottom = (RowInfo(mvarRow).TopY + mvarHeight) / p_TPPY + mvarGridLineWidth - 1
    
    If oTable.HighlightMode = HMFilledRectSolid Then
        If oTable.AlternateRowBackColor <> -1 And (mvarRow Mod 2) = 0 And mvarBackColor = -1 Then
            NoFocusHighlightBackColor = BlendColor(oTable.HighlightBackColor, oTable.AlternateRowBackColor, 92)
        Else
            NoFocusHighlightBackColor = BlendColor(oTable.HighlightBackColor, IIf(mvarBackColor = -1, oTable.BackColor, mvarBackColor), 92)
        End If
    Else
        NoFocusHighlightBackColor = BlendColor(oTable.HighlightBackColor, vbWhite, 120)
    End If
    NoFocusHighlightForeColor = mvarForeColor
    
    If Not mvarPicture Is Nothing Then
        '����ͼƬ��Ԫ��
        LL = tR.Left + 1
        lT = tR.Top + 1
        lW = mvarWidth / p_TPPX + mvarGridLineWidth - 1
        lH = (lW * mvarPicture.Height) / mvarPicture.Width
        If oTable.AutoHeight = False Then
'            '������Զ�����ģʽ����ô������ͼƬ�߶ȳ�����Ԫ��߶�
'            If lH > mvarHeight / p_TPPY Then lH = mvarHeight / p_TPPY
            '������Զ�����ģʽ��ͼƬ����
            lH = mvarHeight / p_TPPY + mvarGridLineWidth - 1
        End If
        frmProperty.picTMP.Cls
        frmProperty.picTMP.Width = lW * p_TPPX
        frmProperty.picTMP.Height = lH * p_TPPY
        frmProperty.picTMP.PaintPicture mvarPicture, 0, 0, frmProperty.picTMP.Width, frmProperty.picTMP.Height
        
        '���ư�͸��ͼƬ
        If mvarSelected Then
            Dim lBlend As Long
            Dim bf As BLENDFUNCTION
            
            ' Draw the first picture:
            bf.BlendOp = AC_SRC_OVER
            bf.BlendFlags = 0
            bf.SourceConstantAlpha = 255
            bf.AlphaFormat = 0
            CopyMemory lBlend, bf, 4
            frmProperty.picBlank.Cls
            frmProperty.picBlank.BackColor = oTable.HighlightBackColor
'            BitBlt frmProperty.picBlank.hDC, 0, 0, lW, lH, frmProperty.picTMP.hDC, 0, 0, vbSrcCopy
            
'            SetStretchBltMode frmProperty.picTMP.hDC, HALFTONE
'            StretchBlt frmProperty.picTMP.hDC, tR.left, tR.top, lW, lH, frmProperty.picBlank.hDC, 0, 0, frmProperty.picBlank.Width, frmProperty.picBlank.Height, SRCCOPY
            
            bf.SourceConstantAlpha = 100
            CopyMemory lBlend, bf, 4
            AlphaBlend frmProperty.picTMP.hDC, 0, 0, _
                frmProperty.picTMP.Width \ Screen.TwipsPerPixelX, _
                frmProperty.picTMP.Height \ Screen.TwipsPerPixelY, _
                frmProperty.picBlank.hDC, 0, 0, _
                frmProperty.picBlank.ScaleWidth \ Screen.TwipsPerPixelX, _
                frmProperty.picBlank.ScaleHeight \ Screen.TwipsPerPixelY, _
                lBlend
            frmProperty.picTMP.Refresh
        End If
        BitBlt lhDC, tR.Left, tR.Top, lW, lH, frmProperty.picTMP.hDC, 0, 0, vbSrcCopy
    Else
        '��䱳��
        If (mvarBackColor <> -1 Or ((oTable.AlternateRowBackColor <> -1) And (mvarRow Mod 2) = 0)) Then
            If (oTable.AlternateRowBackColor <> -1) And (mvarRow Mod 2) = 0 Then
                '����ǽ������Ҿ��н���ɫ
                hBr = CreateSolidBrush(TranslateColor(IIf(mvarBackColor <> -1, mvarBackColor, oTable.AlternateRowBackColor))) '��������ɫ�Ĵ�ɫ��ˢ
            Else
                hBr = CreateSolidBrush(TranslateColor(mvarBackColor))      '��������ɫ�Ĵ�ɫ��ˢ
            End If
            FillRect lhDC, tR, hBr
            DeleteObject hBr
            hBr = 0
        End If
    End If
    
    '���Ƹ�������
    If mvarSelected And oTable.InEdit = False And (mvarPicture Is Nothing) Then
        Select Case oTable.HighlightMode
        Case HMFilledRectSolid, HMFilledRectAlpha
            If oTable.HighlightMode = HMFilledRectSolid Then
                'ʵ�Ļ�ˢ
                hBr = CreateSolidBrush(TranslateColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor)))
                FillRect lhDC, tR, hBr
                DeleteObject hBr
                hBr = 0
            Else
                'ʵ�İ�͸����ˢ
                If (oTable.AlternateRowBackColor <> -1) And (mvarRow Mod 2) = 0 Then
                    '����ǽ������Ҿ��н���ɫ
                    hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), IIf(mvarBackColor <> -1, mvarBackColor, oTable.AlternateRowBackColor), 92))
                Else
                    If mvarBackColor = -1 Then
                        hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), oTable.BackColor, 92))
                    Else
                        hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), mvarBackColor, 92))
                    End If
                End If
                FillRect lhDC, tR, hBr
                DeleteObject hBr
                hBr = 0
            End If
        Case HMOnlyBorderRectAlpha, HMOnlyBorderRectSolid
            If oTable.HighlightMode = HMOnlyBorderRectSolid Then
                'ʵ�Ļ�ˢ���������Ʊ߿�
                hBr = CreateSolidBrush(TranslateColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor)))   '�ø�������ɫ������ɫ��ˢ
                LSet tRCopy = tR                        '��ֵ����
                tRCopy.Right = tRCopy.Right + 1
                tRCopy.Bottom = tRCopy.Bottom + 1
                InflateRect tRCopy, -1, -1              '�����α߽���С1����λ
                FrameRect lhDC, tRCopy, hBr       '�������Ʊ߿򣬿���Ϊ1
                DeleteObject hBr                        '�ͷ���Դ
                hBr = 0
            Else
                'ʵ�İ�͸����ˢ���������Ʊ߿�
                If (oTable.AlternateRowBackColor <> -1) And (mvarRow Mod 2) = 0 Then
                    '����ǽ������Ҿ��н���ɫ
                    hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), IIf(mvarBackColor <> -1, mvarBackColor, oTable.AlternateRowBackColor), 92))
                Else
                    If mvarBackColor = -1 Then
                        hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), oTable.BackColor, 92))
                    Else
                        hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), mvarBackColor, 92))
                    End If
                End If
                LSet tRCopy = tR                        '��ֵ����
                tRCopy.Right = tRCopy.Right + 1
                tRCopy.Bottom = tRCopy.Bottom + 1
                InflateRect tRCopy, -1, -1              '�����α߽���С1����λ
                FrameRect lhDC, tRCopy, hBr       '�������Ʊ߿򣬿���Ϊ1
                DeleteObject hBr                        '�ͷ���Դ
                hBr = 0
            End If
        End Select
        If oTable.DrawFocusRect Then
            '�������ý���߿�
            SetTextColor lhDC, TranslateColor(vbBlack)    '����������ɫ
            LSet tRCopy = tR
            tRCopy.Left = tRCopy.Left + 1
            tRCopy.Top = tRCopy.Top + 1
            InflateRect tRCopy, -1, -1
            DrawFocusRect lhDC, tRCopy
        End If
    End If
    
    '���Ʊ߿�
    hPen = CreatePen(PS_SOLID, 1, mvarGridLineColor)                '���ñ߿���ɫ����
    hPenOld = SelectObject(lhDC, hPen)                        'ѡ�뻭�ʣ�����ɻ���
    For i = 0 To mvarGridLineWidth - 1
        MoveToEx lhDC, tR.Left + i, tR.Bottom - i, tJ
        LineTo lhDC, tR.Right - i, tR.Bottom - i
        LineTo lhDC, tR.Right - i, tR.Top + i
        LineTo lhDC, tR.Left + i, tR.Top + i
        LineTo lhDC, tR.Left + i, tR.Bottom - i
    Next
    SelectObject lhDC, hPenOld
    DeleteObject hPen
    hPen = 0
    
'    Debug.Assert Me.Key <> 7
    GetCellTextBorder lLeft, lTop, lRight, lBottom
    tR.Left = lLeft
    tR.Top = lTop
    tR.Right = lRight
    tR.Bottom = lBottom
    '��������˴�ֱ���䷽ʽ����ôҪ�����ı��߶ȣ���Ϊ��ֱ���䷽ʽֻ�Ե����ı���Ч��ҪӦ�õ������ı�����Ҫ�ֹ�������
    If mvarSingleLine = False Then
        If mvarPicture Is Nothing Then
            '���ı���Ԫ��
            If mvarVAlignment = VALignVCentre Then
                tR.Top = lTop + (lBottom - lTop - (Me.EvaluateTextHeight(oTable) - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY) / 2
            ElseIf mvarVAlignment = VALignBottom Then
                tR.Top = lBottom - (Me.EvaluateTextHeight(oTable) - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY
            End If
        Else
            '����ͼƬ�ĵ�Ԫ��
            Dim lHT As Long
            lHT = Me.EvaluateTextHeight(oTable, , False)    '���ı��߶�
            lH = Me.EvaluateTextHeight(oTable)              '����ͼƬ�ĸ߶�
            If mvarVAlignment = VALignVCentre Then
                If lHT < lH Then
                    tR.Top = lTop + (lBottom - lTop - (Me.EvaluateTextHeight(oTable) - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY) / 2 + (lH - lHT) / (2 * p_TPPY)
                Else
                    tR.Top = lTop + (lBottom - lTop - (Me.EvaluateTextHeight(oTable) - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY) / 2
                End If
            ElseIf mvarVAlignment = VALignBottom Then
                If lHT < lH Then
                    tR.Top = lBottom - (lH - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY + (lH - lHT) / p_TPPY
                Else
                    tR.Top = lBottom - (lH - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY
                End If
            End If
        End If
    End If
    
    '����ͼ��
    If mvarIcon > 0 Then
        DrawImageIcon oTable.PtrImageList, mvarIcon, lhDC, tR.Left - p_lIconWidth / p_TPPX, tR.Top, p_lIconWidth / p_TPPX, _
        p_lIconHeight / p_TPPY, mvarSelected And oTable.HighlightSelectedIcons, Not oTable.Enabled
    End If
    
    '�����ı�
    sText = IIf(mvarFormatString = "", mvarText, Format(mvarText, mvarFormatString))
    With tLF
        sFont = mvarFontName & vbNullString
        temp = StrConv(sFont, vbFromUnicode)
        For iChar = 0 To UBound(temp)
            .lfFaceName(iChar) = temp(iChar)
        Next iChar
        .lfItalic = mvarFontItalic
        .lfHeight = GetPixcelHeightByPoint(lhDC, mvarFontSize)
        .lfUnderline = mvarFontUnderline
        .lfWeight = IIf(mvarFontBold, FW_BOLD, FW_NORMAL)
        .lfStrikeOut = mvarFontStrikeout
        .lfQuality = IIf(oTable.FontQuality < 0, 0, oTable.FontQuality)
        .lfCharSet = DEFAULT_CHARSET
    End With
    hFnt = CreateFontIndirect(tLF) '��������
    '����������ɫ
    If oTable.Enabled Then
        If mvarSelected = False Or oTable.HighlightMode <> HMFilledRectSolid Or (Not mvarPicture Is Nothing) Then
            SetTextColor lhDC, TranslateColor(mvarForeColor)
        Else
            SetTextColor lhDC, TranslateColor(IIf(oTable.InFocus, oTable.HighlightForeColor, NoFocusHighlightForeColor))
        End If
    Else
        SetTextColor lhDC, TranslateColor(vbGrayText)
    End If
    'ѡ��ָ�����壬���������
    hFntOld = SelectObject(lhDC, hFnt)
    '�����ı��������ı����Զ�����ʱ������ʾʡ�Ժţ������ı����Զ��߶�ʱ������ʾʡ�Ժţ�
'    tR.bottom = tR.bottom + 9999
    
    DrawText lhDC, sText, -1, tR, DT_NOPREFIX Or DT_EXPANDTABS Or DT_EDITCONTROL Or mvarVAlignment Or mvarHAlignment Or _
        IIf(mvarSingleLine, DT_SINGLELINE, DT_WORDBREAK) Or _
        IIf(oTable.WordEllipsis, DT_WORD_ELLIPSIS, 0)
        
    '�ָ�������
    SelectObject lhDC, hFntOld
    '�ͷ���ʱ����
    DeleteObject hFnt
    
    tR.Left = (ColInfo(mvarCol).LeftX) / p_TPPX + mvarGridLineWidth - 1
    tR.Top = (RowInfo(mvarRow).TopY) / p_TPPY + mvarGridLineWidth - 1
    tR.Right = (ColInfo(mvarCol).LeftX + mvarWidth) / p_TPPX
    tR.Bottom = (RowInfo(mvarRow).TopY + mvarHeight) / p_TPPY
    
    '�ȸ���
    If oTable.Enabled And (mvarHot Or mvarSelected) Then
        hBr2 = CreateSolidBrush(TranslateColor(oTable.HighlightBackColor))   '�ø�������ɫ������ɫ��ˢ
        LSet tRCopy = tR                        '��ֵ����
        tRCopy.Left = tRCopy.Left + 1
        tRCopy.Top = tRCopy.Top + 1
        FrameRect lhDC, tRCopy, hBr2          '�����α߿򣬿���ʼ��Ϊ1
        If mvarHot Then
            'Hot״̬���ƴ��ߣ�
            InflateRect tRCopy, -1, -1              '�����α߽���С1����λ
            FrameRect lhDC, tRCopy, hBr2      '�������Ʊ߿򣬿���Ϊ1
        End If
        DeleteObject hBr2                       '�ͷ���Դ
    End If
    
    '�������ʾ�ı�����ô����һ����ɫ�����ε����Ͻ�
    If mvarToolTipText <> "" And oTable.ShowToolTipText And bExportMode = False Then
        Dim poly(1 To 3) As POINTAPI, NumCoords As Long, hBrush As Long, hRgn As Long
        NumCoords = 3
        poly(1).X = (ColInfo(mvarCol).LeftX + mvarWidth) / p_TPPX
        poly(1).Y = (RowInfo(mvarRow).TopY) / p_TPPY + mvarGridLineWidth
        poly(2).X = poly(1).X - 5
        poly(2).Y = poly(1).Y
        poly(3).X = poly(1).X
        poly(3).Y = poly(1).Y + 5
        hBrush = CreateSolidBrush(vbRed)
        hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
        If hRgn Then FillRgn lhDC, hRgn, hBrush
        DeleteObject hRgn
        DeleteObject hBrush
    End If
End Sub

Private Sub Class_Initialize()
    'Ĭ�����Գ�ʼ��
    Me.Margin = 2                   'Ĭ�ϱ߾�Ϊ2
    Me.SingleLine = False           'Ĭ��Ϊ�����ı�
    Me.MergeInfo = ""               'Ĭ�ϲ��ϲ�
    Me.Selected = False             'Ĭ��δѡ��
    Me.Hot = False                  'Ĭ��False
    Me.Visibled = True              'Ĭ����ʾ
    Me.Width = 1600                 'Ĭ���п�
    Me.Height = 300                 'Ĭ���и�
    Me.FixedWidth = False           'Ĭ�ϲ������п��������ֹ�����
    Me.AutoHeight = True            'Ĭ��Ϊ�Զ��и�
    Me.Icon = -1                    'Ĭ��û��ͼ�꣬-1
    Me.Text = ""                    'Ĭ���ı�Ϊ��
    Me.Tag = ""                     'Ĭ�ϱ��Ϊ��
    Me.FormatString = ""            'Ĭ�ϸ�ʽ��Ϊ��
    Me.Indent = 0                   'Ĭ������ֵΪ0
    Me.HAlignment = HALignLeft      'Ĭ�Ͽ���
    Me.VAlignment = VALignTop       'Ĭ�Ͽ���
    Me.ForeColor = vbBlack          'Ĭ��ǰ��ɫΪ��ɫ
    Me.BackColor = -1               'Ĭ�ϱ���ɫΪ͸��
    Me.GridLineColor = vbBlack      'Ĭ�ϱ߿���ɫΪ��ɫ
    Me.GridLineWidth = 1            'Ĭ�ϱ߿����Ϊ1
    Me.FontName = "����"            'Ĭ������Ϊ�գ�ȡϵͳĬ������
    Me.FontSize = 11                'Ĭ�������СΪ10��
    Me.FontBold = False             'Ĭ�ϷǴ���
    Me.FontItalic = False           'Ĭ�Ϸ�б��
    Me.FontStrikeout = False        'Ĭ����ɾ����
    Me.FontUnderline = False        'Ĭ�����»���
    Me.FontWeight = 400             'Ĭ������Ϊ400
    Me.Protected = False            'Ĭ�ϲ�����
    Me.Dirty = True                 'Ĭ����Ҫ�ػ�
End Sub