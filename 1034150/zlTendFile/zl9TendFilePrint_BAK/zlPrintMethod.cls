VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlPrintMethod"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
Public objOutTo As Object           '��ӡ�����Ŀ�����,������printer��PictureBox��һҳ
Private WithEvents frmView As frmPreView       'Ԥ������
Attribute frmView.VB_VarHelpID = -1
Private frmView2 As frmPreview2     '�ṩ���ⲿ��̵�Ԥ������
Private frmExcel As frmOutExcel
Private frmSet As frmPrintSet
Private mvarViewPage As Byte        'Ԥ��ҳ��
Private mvarSuccess As Boolean      'ִ�гɹ���־
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'zlPrintSet             ���д�ӡ���ĳ�������
'zlPrintAsk             ���д��������й���������
'zlPrintOrViewTends     ��ӡ��Ԥ��ָ��������¼��
'zlPrintCell            ��ָ�������ӡһ�����ݵ�Ԫ,������ǰ�����ƶ�����Ԫ���Ͻ�λ��
'PaperName              ���ݵ�ǰ��ӡ�������ã���ȡֽ������
'PaperSource            ���ݵ�ǰ��ӡ�������ã���ȡ��ֽ��ʽ����
'zlInitPrn              ��ʼ����ӡ�豸�������ⲿֱ�ӱ�д�����ӡ�豸�ĳ���ʱʹ��
'zlNewPage              ��ӡ�豸����һҳ�������ⲿֱ�ӱ�д�����ӡ�豸�ĳ���ʱʹ��
'zlPreview              ���Ѿ���ɵ�Ԥ���ṩ����Ч���������ⲿֱ�ӱ�д�����ӡ�豸�ĳ���ʱʹ��
'zlPrintEnd             ������ɱ����ӡ�ĵ�
'zlPrintReset           �ָ�ȱʡ��ӡ����
'zlGetColor             ͨ����ɫ���ȡRGB��ɫֵ
'zlGetPrinterSet        ��ȡ��ϵͳע����Ĵ�ӡȱʡ����
'zlPutPrinterSet        ��ϵͳע����б����ӡȱʡ����
'SwitchPrintSet         ��ȡ�û��Ե�ǰ���Ĵ�ӡ���ã���Ϊ��ǰȱʡ��ӡ����
'                       �򽫵�ǰ��ӡȱʡ���ñ���Ϊ�û��Ե�ǰ���Ĵ�ӡ����
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Event zlAfterPrint()

Public Sub InitPrint(ByVal cnMain As ADODB.Connection, ByVal strDBUser As String)
'���ܣ������ݿ����ӽ��г�ʼ��
    Set gcnOracle = cnMain
    gstrDBUser = strDBUser
    gstrPrivs = GetPrivFunc(0, 16)
End Sub

Public Property Let Success(ByVal vData As Boolean)
    mvarSuccess = vData
End Property

Public Property Get OutTo() As Object
    Set OutTo = objOutTo
End Property

Public Property Let ViewPage(ByVal vData As Byte)
    If vData > 0 Then
        mvarViewPage = vData
    End If
End Property

Public Property Get ViewPage() As Byte
    ViewPage = mvarViewPage
End Property

Private Sub Class_Initialize()
    Set objOutTo = Printer
    Set frmAsk = New frmTendPrintAsk
    Set frmView = New frmPreView
    Set frmView2 = New frmPreview2
    Set frmSet = New frmPrintSet
    mvarViewPage = 6
    mvarSuccess = True
    
    gstrGrant = GetSetting("ZLSOFT", "ע����Ϣ", "Kind", "")
    If gstrGrant = "��ʽ" Then gstrGrant = ""
    Call GetOEM
End Sub

Private Sub Class_Terminate()
    Set frmAsk = Nothing
'    Unload frmView
    Set frmView = Nothing
    Set frmView2 = Nothing
    Set frmSet = Nothing
    Set gobjOutTo = Nothing
    Set gobjSend = Nothing
    Set objOutTo = Nothing
    Set gcnOracle = Nothing
End Sub


Public Function zlPrintSet() As Boolean
    '------------------------------------------------
    '���ܣ� ���д�ӡ���ĳ�������
    '������
    '���أ� �ɹ����ã�����true��ϵͳδ��װ��ӡ��������false
    '------------------------------------------------
    If Printers.Count = 0 Then
        MsgBox "ע�⣺" & Chr(13) _
            & "    δ��װ��ӡ������ͨ��ϵͳ���õĴ�ӡ��" & Chr(13) _
            & "�������Ӱ�װ��ӡ����", vbCritical + vbOKOnly, gstrSysName
        zlPrintSet = False
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    zlGetPrinterSet
    frmSet.Show 1
    zlPutPrinterSet
    
    zlPrintSet = True
    Exit Function

errHand:
    MsgBox "ϵͳ���ֲ���Ԥ֪�Ĵ���" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    zlPrintSet = False

End Function

Public Function zlPrintAsk(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, ByVal lng�ļ�ID As Long, Optional ByVal blnBatch As Boolean = False) As Byte
    '------------------------------------------------
    '���ܣ� ���д��������й���������
    '������
    '   objSend:��ӡ����
    '���أ� ������������=0����ӡ=1��Ԥ��=2
    '------------------------------------------------
    If Printers.Count = 0 Then
        MsgBox "ע�⣺" & Chr(13) _
            & "    δ��װ��ӡ������ͨ��ϵͳ���õĴ�ӡ��" & Chr(13) _
            & "�������Ӱ�װ��ӡ����", vbCritical + vbOKOnly, gstrSysName
        zlPrintAsk = 0
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    glng�ļ�ID = lng�ļ�ID
    glng����ID = lng����ID
    glng��ҳID = lng��ҳID
    gintӤ�� = intӤ��
    '��ӡ����ǰ�ر�ѯ�ʴ���
    If blnBatch Then
        Load frmAsk
    Else
        frmAsk.Show vbModal
    End If
    If frmAsk.byRunMode = 0 Then
        zlPrintAsk = 0
        Exit Function
    End If
    
    zlPrintAsk = frmAsk.byRunMode
    Exit Function
errHand:
    MsgBox "ϵͳ���ֲ���Ԥ֪�Ĵ���" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    zlPrintAsk = 0
    
End Function
'
'Public Function zlPrintOrViewTends(objSend As zlPrintTends, Optional byRunMode As Byte = 1) As Boolean
'    '------------------------------------------------
'    '���ܣ� ��ӡ��Ԥ��ָ��msFlexGrid��ӡ����
'    '������
'    '   objSend:��ӡ����
'    '   byRunMode:���ģʽ��1-��ӡ��2-Ԥ��
'    '���أ� �ɹ�����true �����󷵻�false
'    '------------------------------------------------
'    Dim GridRedraw As Boolean
'
'    If byRunMode = 3 Then
'        If gstrGrant <> "" Then
'            MsgBox "���û���԰汾����ʹ�øù��ܡ�", vbInformation, gstrSysName
'            Exit Function
'        End If
'        If Not HaveExcel() Then
'            MsgBox "�밲װ��Excel��WPS(ET)����ʹ�ñ����ܡ�", vbCritical + vbOKOnly, gstrSysName
'            zlPrintOrViewTends = False
'            Exit Function
'        End If
'        If objSend.Privileged = False Then
'            If InStr(1, ";" & gstrPrivs & ";", ";Excel���;") = 0 Then
'                MsgBox "��û��Excel�����Ȩ�ޣ�����ʹ�øù��ܡ�", vbInformation + vbOKOnly, gstrSysName
'                zlPrintOrViewTends = False
'                Exit Function
'            End If
'        End If
'
'        Set gobjSend = objSend
'        Set frmExcel = New frmOutExcel
'
'        GridRedraw = objSend.Body.Redraw
'        objSend.Body.Redraw = False
'        frmExcel.Show 1
'        objSend.Body.Redraw = GridRedraw
'
'        Set frmExcel = Nothing
'        Exit Function
'    End If
'
'    If Printers.Count = 0 Then
'        MsgBox "ע�⣺" & Chr(13) _
'            & "    δ��װ��ӡ������ͨ��ϵͳ���õĴ�ӡ��" & Chr(13) _
'            & "�������Ӱ�װ��ӡ����", vbCritical + vbOKOnly, gstrSysName
'            zlPrintOrViewTends = False
'        Exit Function
'    End If
'    zlGetPrinterSet
'
'    Err = 0
'    On Error GoTo errHand
'    '----------------------------------------------------
'    '   �������Ա���
'    '----------------------------------------------------
'    GridRedraw = objSend.Body.Redraw
'    objSend.Body.Redraw = False
'
'    '----------------------------------------------------
'    '��ʼ�����Ŀ�����
'    '----------------------------------------------------
'    gintObjType = gintTends
'    Set gobjSend = objSend
'    If Not ReadVar Then
'        objSend.Body.Redraw = GridRedraw
'        zlPrintOrViewTends = False
'        Exit Function
'    End If
'
'    '��಻��һ�о���һҳ��
'    Call CalculateHeight
'    Call CalculateRC
'
'    If byRunMode = 1 Then
'        If objSend.Privileged = False Then
'            If InStr(1, ";" & gstrPrivs & ";", ";��ӡ;") = 0 Then
'                MsgBox "��û�д�ӡȨ�ޣ�����ʹ�øù��ܡ�", vbInformation + vbOKOnly, gstrSysName
'                zlPrintOrViewTends = False
'                Exit Function
'            End If
'        End If
'        RealPrint 1, gintRowTotal * gintColTotal
'    Else
'        Set frmView = New frmPreView
'        frmView.Show 1
'        Set frmView = Nothing
'    End If
'
'
'    If Not mvarSuccess Then zlPrintOrViewTends = False: Exit Function
'
'    '�������Իָ�
'    objSend.Body.Redraw = GridRedraw
'    zlPrintOrViewTends = True
'    Exit Function
'
'errHand:
'    MsgBox "ϵͳ���ֲ���Ԥ֪�Ĵ���" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
'    objSend.Body.Redraw = GridRedraw
'    zlPrintOrViewTends = False
'End Function

Public Function zlPrintOrViewTends(ByVal blnBatch As Boolean, Optional byRunMode As Byte = 1) As Boolean
    '------------------------------------------------
    '���ܣ� ��ӡ��Ԥ��ָ��msFlexGrid��ӡ����
    '������
    '   byRunMode:���ģʽ��1-��ӡ��2-Ԥ��
    '���أ� �ɹ�����true �����󷵻�false
    '------------------------------------------------
    Dim lngMaxPage As Long, lngStartPage As Long, lngPages As Long
    gblnPrintMode = (byRunMode = 1)
    If byRunMode = 3 Then
        If gstrGrant <> "" Then
            MsgBox "���û���԰汾����ʹ�øù��ܡ�", vbInformation, gstrSysName
            GoTo ex_d
        End If
        If Not HaveExcel() Then
            MsgBox "�밲װ��Excel����ʹ�ñ����ܡ�", vbCritical + vbOKOnly, gstrSysName
            GoTo ex_d
        End If

        Load frmTendFileReader
        lngMaxPage = frmTendFileReader.GetPages
        lngStartPage = frmTendFileReader.GetStartPage
        lngPages = lngMaxPage - lngStartPage + 1

        If gobjSend.Privileged = False Then
            If InStr(1, ";" & gstrPrivs & ";", ";Excel���;") = 0 Then
                MsgBox "��û��Excel�����Ȩ�ޣ�����ʹ�øù��ܡ�", vbInformation + vbOKOnly, gstrSysName
                GoTo ex_d
            End If
        End If

        Set frmExcel = New frmOutExcel
        frmExcel.Show 1

        Set frmExcel = Nothing
        GoTo ex_d
    End If
    
    If Printers.Count = 0 Then
        MsgBox "ע�⣺" & Chr(13) _
            & "    δ��װ��ӡ������ͨ��ϵͳ���õĴ�ӡ��" & Chr(13) _
            & "�������Ӱ�װ��ӡ����", vbCritical + vbOKOnly, gstrSysName
            zlPrintOrViewTends = False
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand
    Load frmTendFileReader
    lngMaxPage = frmTendFileReader.GetPages
    lngStartPage = frmTendFileReader.GetStartPage
    lngPages = lngMaxPage - lngStartPage + 1
'    If lngPages = 0 Then
'        MsgBox "û�д�ӡ���ݣ�", vbInformation, gstrSysName
'        GoTo ex_d
'    End If
    If Not ReadVar Then GoTo ex_d
    
    If byRunMode = 1 Then
        If InStr(1, ";" & gstrPrivs & ";", ";��ӡ;") = 0 Then
            MsgBox "��û�д�ӡȨ�ޣ�����ʹ�øù��ܡ�", vbInformation + vbOKOnly, gstrSysName
            GoTo ex_d
        End If
        RealPrint 1, lngPages
    Else
        Set frmView = New frmPreView
        frmView.Show 1
        Set frmView = Nothing
    End If
    If Not mvarSuccess Then GoTo ex_d
    
    zlPrintOrViewTends = True
    GoTo ex_d
errHand:
    MsgBox "ϵͳ���ֲ���Ԥ֪�Ĵ���" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    zlPrintOrViewTends = False
ex_d:
    Unload frmTendFileReader
    Unload frmAsk
End Function

Public Sub zlPrintCell(ByVal Text As String, _
    ByVal x As Long, ByVal Y As Long, _
    Optional ByVal Wide, _
    Optional ByVal High, _
    Optional Alignment As Byte = 0, _
    Optional ForeColor As Long = 0, _
    Optional GridColor As Long = 0, _
    Optional FillColor As Long = 0, _
    Optional LineStyle As String = "1111", _
    Optional FontName, Optional FontSize, _
    Optional FontBold, Optional FontItalic)
    '------------------------------------------------
    '���ܣ� ��ָ�������ӡһ�����ݵ�Ԫ,������ǰ�����ƶ�����Ԫ���Ͻ�λ��
    '������
    '   Text:    ������ַ���,���в������س����з�
    '   X:       ���Ͻ�X����
    '   Y:       ���Ͻ�Y����
    '   Wide:    �������
    '   High:    ����߶�
    '   Alignment:    ����ģʽ��0-�����(ȱʡ),1-�Ҷ���,2-����
    '   ForeColorǰ��ɫ,ȱʡΪ��ɫ
    '   GridColor����ɫ,ȱʡΪ��ɫ
    '   FillColor���ɫ,ȱʡΪ�豸����ɫ,����ϵͳ�����˺�ɫ��ɫ�룬���Խ�����������ɫ
    '   LineStyle:����ֱ�Ϊ�������µ���������
    '           0-���ߣ�1-9����Ӵ֣�1Ϊȱʡ
    '   FontName,FontSize,FontBold,FontItalic:��������
    '���أ�
    '------------------------------------------------
    Dim aryString() As String       '�س��ָ���ַ���
    Dim lngOldForeColor As Long     '����豸ȱʡǰ��ɫ
    Dim intRow As Long, intAllRow As Long
    Dim strRest As String, lngYMove As Long
    Dim oldFontName, oldFontSize, oldFontBold, oldFontItalic
    lngOldForeColor = objOutTo.ForeColor
    
    On Error Resume Next
    With objOutTo
        If Not IsMissing(FontName) Then
            oldFontName = objOutTo.FontName
            .FontName = FontName
        End If
        If Not IsMissing(FontSize) Then
            .FontSize = FontSize
            oldFontSize = objOutTo.FontSize
        End If
        If Not IsMissing(FontBold) Then
            .FontBold = FontBold
            oldFontBold = objOutTo.FontBold
        End If
        If Not IsMissing(FontItalic) Then
            .FontItalic = FontItalic
            oldFontItalic = objOutTo.FontItalic
        End If
    End With
    
    If IsMissing(Wide) Then Wide = objOutTo.TextWidth(Text) + 2 * conLineWide
    If IsMissing(High) Then High = objOutTo.TextHeight(Text) + 2 * conLineHigh
    
    If Wide * High = 0 Then Exit Sub
    
    If UCase(TypeName(LineStyle)) <> "STRING" Then LineStyle = CStr(LineStyle)
    If Len(LineStyle) < 4 Then
        LineStyle = Left(LineStyle & "1111", 4)
    End If
    
    '------------------------------------------
    '   ���ߴ�ӡ
    '------------------------------------------
    If Mid(LineStyle, 1, 1) <> 0 Then
        objOutTo.DrawWidth = Mid(LineStyle, 1, 1)
        objOutTo.Line (x, Y)-(x + Wide, Y), GridColor
    End If
    
    If Mid(LineStyle, 2, 1) <> 0 Then
        objOutTo.DrawWidth = Mid(LineStyle, 2, 1)
        objOutTo.Line (x, Y)-(x, Y + High), GridColor
    End If
    
    If Mid(LineStyle, 3, 1) <> 0 Then
        objOutTo.DrawWidth = Mid(LineStyle, 3, 1)
        objOutTo.Line (x + Wide, Y)-(x + Wide, Y + High), GridColor
    End If
    
    If Mid(LineStyle, 4, 1) <> 0 Then
        objOutTo.DrawWidth = Mid(LineStyle, 4, 1)
        objOutTo.Line (x, Y + High)-(x + Wide, Y + High), GridColor
    End If
    
    If Wide > conLineWide And High > conLineHigh Then
        '------------------------------------------
        '   ��ɫ���
        '------------------------------------------
'        If FillColor <> 0 Then
'            Printer.FillStyle = 1
'            objOutTo.Line (X + conLineWide / 2, Y + conLineHigh / 2)- _
'                (X + Wide - conLineWide / 2, Y + High - conLineHigh / 2), _
'                FillColor, BF
'        End If
        
        '------------------------------------------
        '   ���ִ�ӡ
        '------------------------------------------
        objOutTo.ForeColor = ForeColor
    
        If InStr(1, Text, Chr(13)) = 0 Then
            If Wide - conLineWide < objOutTo.TextWidth("1") Then    'С��һ���ַ�
                intAllRow = 1
            Else
                If objOutTo.TextWidth(Text) Mod (Wide - conLineWide) = 0 Then
                    intAllRow = objOutTo.TextWidth(Text) \ (Wide - conLineWide)
                Else
                    intAllRow = objOutTo.TextWidth(Text) \ (Wide - conLineWide) + 1
                End If
            End If
            For intRow = intAllRow To 1 Step -1
                If High >= objOutTo.TextHeight(Text) * intRow Then
                    Exit For
                End If
            Next
            intAllRow = intRow
            lngYMove = (High - conLineHigh - objOutTo.TextHeight(Text) * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow - 1
                Do While objOutTo.TextWidth(Text) > Wide - conLineWide
                    If Len(Trim(Text)) <= 1 Then Exit Do
                    Text = Left(Text, Len(Text) - 1)
                Loop
                strRest = Mid(strRest, Len(Text) + 1)
                Select Case Alignment
                Case 2
                    objOutTo.CurrentX = x + (Wide - objOutTo.TextWidth(Text)) / 2
                Case 1
                    objOutTo.CurrentX = x - conLineWide / 2 + Wide - objOutTo.TextWidth(Text)
                Case Else
                    objOutTo.CurrentX = x + conLineWide / 2
                End Select
                objOutTo.CurrentY = Y + conLineHigh / 2 + lngYMove + intRow * objOutTo.TextHeight(Text)
                objOutTo.Print Text
                Text = strRest
            Next
        Else
            intAllRow = 1
            aryString = Split(Trim(Text), Chr(13))
            intAllRow = UBound(aryString)
            lngYMove = (High - conLineHigh - objOutTo.TextHeight("ZYL") * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow
                Text = aryString(intRow)
                Select Case Alignment
                Case 2
                    objOutTo.CurrentX = x + (Wide - objOutTo.TextWidth(Text)) / 2
                Case 1
                    objOutTo.CurrentX = x - conLineWide / 2 + Wide - objOutTo.TextWidth(Text)
                Case Else
                    objOutTo.CurrentX = x + conLineWide / 2
                End Select
                objOutTo.CurrentY = Y + conLineHigh / 2 + lngYMove + intRow * objOutTo.TextHeight(Text)
                objOutTo.Print Text
                Text = strRest
            Next
        
        End If
    End If
    objOutTo.CurrentX = x + Wide
    objOutTo.CurrentY = Y
    objOutTo.DrawStyle = 0
    objOutTo.DrawWidth = 1
    objOutTo.ForeColor = lngOldForeColor

    If Not IsMissing(FontName) Then objOutTo.FontName = oldFontName
    If Not IsMissing(FontSize) Then objOutTo.FontSize = oldFontSize
    If Not IsMissing(FontBold) Then objOutTo.FontBold = oldFontBold
    If Not IsMissing(FontItalic) Then objOutTo.FontItalic = oldFontItalic
End Sub


Public Function PaperName() As String
    '------------------------------------------------
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
    '������
    '���أ� ֽ������
    '------------------------------------------------
    Dim mSize As Integer
    Err = 0
    On Error GoTo errHand
    
    If Printer.PaperSize = 256 Then
        PaperName = "�û��Զ��壬" _
            & Printer.Width / 56.6857142857143 & "x" _
            & Printer.Height / 56.6857142857143 & "����"
        Exit Function
    End If
    If Printer.PaperSize >= 1 And Printer.PaperSize <= 41 Then
        mSize = Printer.PaperSize
        PaperName = IIf(Printer.Orientation = 1, "����", "����") & Space(2) _
            & Switch( _
            mSize = 1, conSize1, mSize = 2, conSize2, mSize = 3, conSize3, mSize = 4, conSize4, mSize = 5, conSize5, _
            mSize = 6, conSize6, mSize = 7, conSize7, mSize = 8, conSize8, mSize = 9, conSize9, mSize = 10, conSize10, _
            mSize = 11, conSize11, mSize = 12, conSize12, mSize = 13, conSize13, mSize = 14, conSize14, mSize = 15, conSize15, _
            mSize = 16, conSize16, mSize = 17, conSize17, mSize = 18, conSize18, mSize = 19, conSize19, mSize = 20, conSize20, _
            mSize = 21, conSize21, mSize = 22, conSize22, mSize = 23, conSize23, mSize = 24, conSize24, mSize = 25, conSize25, _
            mSize = 26, conSize26, mSize = 27, conSize27, mSize = 28, conSize28, mSize = 29, conSize29, mSize = 30, conSize30, _
            mSize = 31, conSize31, mSize = 32, conSize32, mSize = 33, conSize33, mSize = 34, conSize34, mSize = 35, conSize35, _
            mSize = 36, conSize36, mSize = 37, conSize37, mSize = 38, conSize38, mSize = 39, conSize39, mSize = 40, conSize40, _
            mSize = 41, conSize41)
        Exit Function
    End If
errHand:
    PaperName = "���ɲ��ֽ��"
    
End Function

Public Function PaperSource() As String
    '------------------------------------------------
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡ��ֽ��ʽ����
    '������
    '���أ� ��ֽ��ʽ�ַ���
    '------------------------------------------------
    Dim mBin As Integer
    
    Err = 0
    On Error GoTo errHand
    
    If Printer.PaperBin = 14 Then
        PaperSource = "���ӵĿ�ʽֽ�н�ֽ"
        Exit Function
    End If
    If Printer.PaperBin >= 1 And Printer.PaperBin <= 11 Then
        PaperSource = Switch( _
            mBin = 1, conBin1, mBin = 2, conBin2, mBin = 3, conBin3, mBin = 4, conBin4, mBin = 5, conBin5, _
            mBin = 6, conBin6, mBin = 7, conBin7, mBin = 8, conBin8, mBin = 9, conBin9, mBin = 10, conBin10, _
            mBin = 11, conBin11)
        Exit Function
    End If
errHand:
    PaperSource = "���ɲ�Ľ�ֽ��ʽ"

End Function

Public Function zlInitPrn(Optional byRunMode As Byte = 1, _
    Optional lngHeight As Long = 0, _
    Optional lngWidth As Long = 0) As Boolean
    '------------------------------------------------
    '���ܣ� ��ʼ����ӡ�豸�������ⲿֱ�ӱ�д�����ӡ�豸�ĳ���ʱʹ��
    '������
    '   byRunMode:��ӡģʽ��ȱʡ��1���豸ָ���ӡ����2-ָ��Ԥ��ҳ��
    '���أ��ɹ�������true;ʧ�ܣ�����false,�Ҵ�ӡ�豸Ϊ��
    '------------------------------------------------
    Dim iPage As Integer
    If Printers.Count = 0 Then
        MsgBox "ע�⣺" & Chr(13) _
            & "    δ��װ��ӡ������ͨ��ϵͳ���õĴ�ӡ��" & Chr(13) _
            & "�������Ӱ�װ��ӡ����", vbCritical + vbOKOnly, gstrSysName
        Set objOutTo = Nothing
        zlInitPrn = False
        Exit Function
    End If
    
    Err = 0
    On Error Resume Next
    If Printer.Orientation = 1 Then
        If lngHeight <> 0 Then Printer.Height = lngHeight
        If lngWidth <> 0 Then Printer.Width = lngWidth
    Else
        If lngHeight <> 0 Then Printer.Width = lngHeight
        If lngWidth <> 0 Then Printer.Height = lngWidth
    End If
    Err = 0
    On Error GoTo errHand
    If byRunMode = 1 Then
        Set objOutTo = Printer
    Else
        If frmView2.PctPage.Count > 1 Then
            Set frmView2 = Nothing
            Set frmView2 = New frmPreView
        End If
        frmView2.PctPage(0).Height = Printer.Height
        frmView2.PctPage(0).Width = Printer.Width
        Load frmView2.PctPage(1)
        Set objOutTo = frmView2.PctPage(1)
    End If
    objOutTo.FillStyle = 1          '�������͸��
    
    zlInitPrn = True
    Exit Function

errHand:
    MsgBox "ϵͳ���ֲ���Ԥ֪�Ĵ���" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    Set objOutTo = Nothing
    zlInitPrn = False

End Function


Public Function zlNewPage() As Boolean
    '------------------------------------------------
    '���ܣ� ��ӡ�豸����һҳ�������ⲿֱ�ӱ�д�����ӡ�豸�ĳ���ʱʹ��
    '������
    '���أ��ɹ�������true;ʧ�ܣ�����false
    '------------------------------------------------
    If Printers.Count = 0 Then
        MsgBox "ע�⣺" & Chr(13) _
            & "    δ��װ��ӡ������ͨ��ϵͳ���õĴ�ӡ��" & Chr(13) _
            & "�������Ӱ�װ��ӡ����", vbCritical + vbOKOnly, gstrSysName
        Set objOutTo = Nothing
        zlNewPage = False
        Exit Function
    End If
    
    If objOutTo Is Nothing Then
        MsgBox "ע�⣺" & Chr(13) _
            & "û�г�ʼ����ӡ�豸��", vbCritical + vbOKOnly, gstrSysName
        Set objOutTo = Nothing
        zlNewPage = False
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand
    If objOutTo Is Printer Then
        objOutTo.NewPage
    Else
        Dim intPageNo As Integer
        intPageNo = frmView2.PctPage.Count
        Load frmView2.PctPage(intPageNo)
        frmView2.PctPage(intPageNo).Height = Printer.Height
        frmView2.PctPage(intPageNo).Width = Printer.Width
        Set objOutTo = frmView2.PctPage(intPageNo)
    End If
    zlNewPage = True
    Exit Function

errHand:
    MsgBox "ϵͳ���ֲ���Ԥ֪�Ĵ���" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    Set objOutTo = Nothing
    zlNewPage = False

End Function

Public Function zlPreview() As Boolean
    '------------------------------------------------
    '���ܣ����Ѿ���ɵ�Ԥ���ṩ����Ч���������ⲿֱ��
    '   ��д�����ӡ�豸�ĳ���ʱʹ��
    '������
    '���أ��ɹ�������true;ʧ�ܣ�����false
    '------------------------------------------------
    If frmView2.PctPage.Count < 2 Then
        MsgBox "ע�⣺" & Chr(13) _
            & "��������˴�ӡ����", vbCritical + vbOKOnly, gstrSysName
        Set objOutTo = Nothing
        zlPreview = False
        Exit Function
    End If
    
    frmView2.Toolbar.Buttons("print").Enabled = False
    frmView2.Show 1
    Unload frmView2
    zlPreview = True
    Exit Function
    
errHand:
    MsgBox "ϵͳ���ֲ���Ԥ֪�Ĵ���" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    zlPreview = False

End Function

Public Sub zlPrintEnd(Optional EndMode As Byte = 0)
    '----------------------------
    '���ܣ�������ɱ����ӡ�ĵ�
    '----------------------------
    If EndMode <> 0 Then
        Printer.KillDoc
    Else
        Printer.EndDoc
    End If
End Sub

Public Sub zlPrintReset()
    '----------------------------
    '���ܣ��ָ�ȱʡ��ӡ����
    '----------------------------
    Class_Terminate
    Class_Initialize
End Sub


Public Function zlGetColor(Optional Color As Long) As Long
    '------------------------------------------------
    '���ܣ�ͨ����ɫ���ȡRGB��ɫֵ
    '������
    '   color:ȱʡ��ɫ
    '���أ�ȷ����������ɫ;ȡ��������ȱʡ��ɫ���ɫ
    '------------------------------------------------
    Dim blnDefault As Boolean
    
    Err = 0
    On Error GoTo errHand
    If Not IsMissing(Color) Then
        blnDefault = True
        If LCase(TypeName(Color)) <> "long" And _
            LCase(TypeName(Color)) <> "integer" And _
            LCase(TypeName(Color)) <> "byte" Then
            MsgBox "ȱʡ��ɫ����!", vbExclamation, gstrSysName
            blnDefault = False
        End If
    Else
        blnDefault = False
    End If
    
    With frmAsk
        If blnDefault Then .comDlg.Color = Color
        .comDlg.CancelError = True
        .comDlg.Flags = cdlCCRGBInit
        .comDlg.ShowColor
        zlGetColor = .comDlg.Color
    End With
    Exit Function
        
errHand:
    If blnDefault Then
        zlGetColor = Color
    Else
        zlGetColor = 0
    End If

End Function

Public Function zlGetPrinterSet() As Boolean
    
    zlGetPrinterSet = GetPrinterSet
End Function

Public Function zlPutPrinterSet() As Boolean
    '------------------------------------------------
    '���ܣ���ϵͳע����б����ӡȱʡ����
    '------------------------------------------------
    If Printers.Count = 0 Then
        zlPutPrinterSet = False
        Exit Function
    End If
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperSize", Printer.PaperSize
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperBin", Printer.PaperBin
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Orientation", Printer.Orientation
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Width", Printer.Width
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Height", Printer.Height
    zlPutPrinterSet = True
End Function

Public Function SwitchPrintSet(strSection As String, Optional blnSave As Boolean = False) As Boolean
    '------------------------------------------------
    '���ܣ���ȡ�û��Ե�ǰ���Ĵ�ӡ���ã���Ϊ��ǰȱʡ��ӡ����
    '   �򽫵�ǰ��ӡȱʡ���ñ���Ϊ�û��Ե�ǰ���Ĵ�ӡ����
    '������
    '   strSection:��ǰ����ӡ���ö�����
    '   blnSave:true,����;false,��ȡ
    '------------------------------------------------
    Dim strDeviceName As String
    Dim intPaperSize As Integer
    Dim intPaperBin As Integer
    Dim intOrientation As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim strSource As String, strObject As String
    
    If Printers.Count = 0 Then
        SwitchPrintSet = False
        Exit Function
    End If
    
    If blnSave Then
        strSource = "����ģ��\" & App.ProductName & "\Default"
        strObject = "����ģ��\" & App.ProductName & "\" & strSection
    Else
        strSource = "����ģ��\" & App.ProductName & "\" & strSection
        strObject = "����ģ��\" & App.ProductName & "\Default"
    End If
    
    strDeviceName = GetSetting("ZLSOFT", strSource, "DeviceName", Printer.DeviceName)
    intPaperSize = GetSetting("ZLSOFT", strSource, "PaperSize", Printer.PaperSize)
    lngWidth = GetSetting("ZLSOFT", strSource, "Width", Printer.Width)
    lngHeight = GetSetting("ZLSOFT", strSource, "Height", Printer.Height)
    intPaperBin = GetSetting("ZLSOFT", strSource, "PaperBin", Printer.PaperBin)
    intOrientation = GetSetting("ZLSOFT", strSource, "Orientation", Printer.Orientation)
    
    SaveSetting "ZLSOFT", strObject, "DeviceName", strDeviceName
    SaveSetting "ZLSOFT", strObject, "PaperSize", intPaperSize
    SaveSetting "ZLSOFT", strObject, "PaperBin", intPaperBin
    SaveSetting "ZLSOFT", strObject, "Orientation", intOrientation
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Width", lngWidth
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Height", lngHeight
    
    SwitchPrintSet = True
End Function

Private Sub GetOEM()
'��ע����еõ���Ʒע����Ϣ
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), "")
    If gstrSysName = "" Then
        gstrSysName = "��������"
    End If
End Sub

Private Sub frmView_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub

