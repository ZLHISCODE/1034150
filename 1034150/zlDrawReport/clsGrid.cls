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
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'######################################################################################################################

Public Enum CellProperty
    AligmentV = 1                           '��ֱ����
    AligmentH = 2                           'ˮƽ����
    FontName = 3
    FontSize = 4
    ForeColor = 5
    WordWarp = 6
    MuliLine = 7
    AutoFit = 8
    Data = 9
    OutPutV = 10
End Enum

Private mrsGridFormat As ADODB.Recordset
Private mrsGridData As ADODB.Recordset
Private mrsGridRowFormat As ADODB.Recordset
Private mrsGridColFormat As ADODB.Recordset

Private mclsCommonLib As clsCommonLib

'######################################################################################################################
Public Function InitGrid(ByVal bytAligment As Byte, ByVal lngForeColor As Long, ByVal lngBackColor As Long, ByVal strFontName As String, ByVal sglFontNameSize As Single, Optional ByVal strGridLine As String = "111111", Optional ByVal bytAligmentV As Byte = 2) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʼ�����������
    '������ rsData��- ��Ԫ�����ݣ�rsRowFormat -  �и�ʽ˵����rsColFormat - �и�ʽ˵����
    '���أ�
    '******************************************************************************************************************
    
    Set mrsGridFormat = New ADODB.Recordset
    With mrsGridFormat
        .Fields.Append "������", adVarChar, 12
        .Fields.Append "ǰ��ɫ", adVarChar, 20
        .Fields.Append "����ɫ", adVarChar, 20
        .Fields.Append "����", adVarChar, 20
        .Fields.Append "��С", adVarChar, 10
        .Fields.Append "ˮƽ����", adVarChar, 2
        .Fields.Append "��ֱ����", adVarChar, 2
        .Open
    End With
        
    Set mrsGridColFormat = New ADODB.Recordset
    With mrsGridColFormat
        .Fields.Append "�к�", adTinyInt
        .Fields.Append "�п�", adVarChar, 30
        .Fields.Append "����", adTinyInt
        .Fields.Append "�߿�", adTinyInt
        .Open
    End With
    
    Set mrsGridRowFormat = New ADODB.Recordset
    With mrsGridRowFormat
        .Fields.Append "�к�", adTinyInt
        .Fields.Append "�и�", adVarChar, 30
        .Fields.Append "����", adTinyInt
        .Fields.Append "�߿�", adTinyInt
        .Open
    End With
    
    Set mrsGridData = New ADODB.Recordset
    With mrsGridData
        .Fields.Append "�к�", adTinyInt
        .Fields.Append "�к�", adTinyInt
        .Fields.Append "����", adVarChar, 1000
        .Fields.Append "ǰ��ɫ", adVarChar, 20
        .Fields.Append "����ɫ", adVarChar, 20
        .Fields.Append "����", adVarChar, 20
        .Fields.Append "��С", adVarChar, 10
        .Fields.Append "ˮƽ����", adVarChar, 2
        .Fields.Append "��ֱ����", adVarChar, 2
        .Fields.Append "�Զ�����", adTinyInt
        .Fields.Append "�������", adTinyInt
        .Fields.Append "�Զ���Ӧ", adTinyInt
        .Fields.Append "��ת�Ƕ�", adTinyInt   '0-����;1-���ϵ�������;2-���ϵ���������ת90��;3-���µ�������;4-���µ�����ת90��
        .Open
    End With

    mrsGridFormat.AddNew
    mrsGridFormat("ˮƽ����").Value = bytAligment
    mrsGridFormat("��ֱ����").Value = bytAligmentV
    mrsGridFormat("ǰ��ɫ").Value = lngForeColor
    mrsGridFormat("����ɫ").Value = lngBackColor
    mrsGridFormat("����").Value = strFontName
    mrsGridFormat("��С").Value = sglFontNameSize
    mrsGridFormat("������").Value = Mid(strGridLine, 1, 1) & ";" & Mid(strGridLine, 2, 1) & ";" & Mid(strGridLine, 3, 1) & ";" & Mid(strGridLine, 4, 1) & ";" & Mid(strGridLine, 5, 1) & ";" & Mid(strGridLine, 6, 1)
    
    InitGrid = True
    
End Function

Public Property Let Cell(ByVal Property As CellProperty, ByVal intRow As Integer, ByVal intCol As Integer, ByVal vData As Variant)

    mrsGridData.Filter = ""
    mrsGridData.Filter = "�к�=" & intRow & " And �к�=" & intCol
    
    If mrsGridData.RecordCount > 0 Then
        Select Case Property
        Case CellProperty.AligmentH
            mrsGridData("ˮƽ����").Value = Val(vData)
        Case CellProperty.AligmentV
            mrsGridData("��ֱ����").Value = Val(vData)
        Case CellProperty.FontName
            mrsGridData("����").Value = Trim(vData)
        Case CellProperty.FontSize
            mrsGridData("��С").Value = Val(vData)
        Case CellProperty.ForeColor
            mrsGridData("ǰ��ɫ").Value = Val(vData)
        Case CellProperty.WordWarp
            mrsGridData("�Զ�����").Value = Val(vData)
        Case CellProperty.MuliLine
            mrsGridData("�������").Value = Val(vData)
        Case CellProperty.AutoFit
            mrsGridData("�Զ���Ӧ").Value = Val(vData)
        Case CellProperty.Data
            mrsGridData("����").Value = vData
        Case CellProperty.OutPutV
            mrsGridData("��ת�Ƕ�").Value = Val(vData)
        End Select
    Else
        Select Case Property
        Case CellProperty.Data
    
            mrsGridData.AddNew
            mrsGridData("�к�").Value = intRow
            mrsGridData("�к�").Value = intCol
            mrsGridData("����").Value = vData
            mrsGridData("����").Value = mrsGridFormat("����").Value
            mrsGridData("��С").Value = mrsGridFormat("��С").Value
            mrsGridData("ǰ��ɫ").Value = mrsGridFormat("ǰ��ɫ").Value
            mrsGridData("����ɫ").Value = mrsGridFormat("����ɫ").Value
            mrsGridData("ˮƽ����").Value = mrsGridFormat("ˮƽ����").Value
            mrsGridData("��ֱ����").Value = mrsGridFormat("��ֱ����").Value
            mrsGridData("�Զ�����").Value = 0
            mrsGridData("�������").Value = 0
            mrsGridData("�Զ���Ӧ").Value = 0
            mrsGridData("��ת�Ƕ�").Value = 0
        End Select
    End If
    

        
End Property

Public Property Let RowCell(ByVal Property As CellProperty, ByVal intRow1 As Integer, ByVal intRow2 As Integer, ByVal vData As Variant)

    mrsGridData.Filter = ""
    mrsGridData.Filter = "�к�>=" & intRow1 & " And �к�<=" & intRow2
    
    If mrsGridData.RecordCount > 0 Then
        mrsGridData.MoveFirst
        Do While Not mrsGridData.EOF
        
            Select Case Property
            Case CellProperty.AligmentH
                mrsGridData("ˮƽ����").Value = Val(vData)
            Case CellProperty.AligmentV
                mrsGridData("��ֱ����").Value = Val(vData)
            Case CellProperty.FontName
                mrsGridData("����").Value = Trim(vData)
            Case CellProperty.FontSize
                mrsGridData("��С").Value = Val(vData)
            Case CellProperty.ForeColor
                mrsGridData("ǰ��ɫ").Value = Val(vData)
            End Select
            
            mrsGridData.MoveNext
        Loop
    End If
    
End Property

Public Property Let ColCell(ByVal Property As CellProperty, ByVal intCol1 As Integer, ByVal intCol2 As Integer, ByVal vData As Variant)

    mrsGridData.Filter = ""
    mrsGridData.Filter = "�к�>=" & intCol1 & " And �к�<=" & intCol2
    
    If mrsGridData.RecordCount > 0 Then
        mrsGridData.MoveFirst
        Do While Not mrsGridData.EOF
        
            Select Case Property
            Case CellProperty.AligmentH
                mrsGridData("ˮƽ����").Value = Val(vData)
            Case CellProperty.AligmentV
                mrsGridData("��ֱ����").Value = Val(vData)
            Case CellProperty.FontName
                mrsGridData("����").Value = Trim(vData)
            Case CellProperty.FontSize
                mrsGridData("��С").Value = Val(vData)
            Case CellProperty.ForeColor
                mrsGridData("ǰ��ɫ").Value = Val(vData)
            End Select
            
            mrsGridData.MoveNext
        Loop
    End If
    
End Property


Public Function SetGridRowFormat(ByVal strRowHeight As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strRowHeight, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGridRowFormat.AddNew
        mrsGridRowFormat("�к�").Value = intLoop + 1
        mrsGridRowFormat("�и�").Value = mclsCommonLib.GetTwipsY(Val(varTmp(intLoop)))
        mrsGridRowFormat("�߿�").Value = 1
        mrsGridRowFormat("����").Value = 0
    Next
    
    SetGridRowFormat = True
End Function

Public Function SetGridRowHeight(ByVal intRow As Integer, ByVal sglHeight As Single) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mrsGridRowFormat.Filter = ""
    mrsGridRowFormat.Filter = "�к�=" & intRow + 1
    If mrsGridRowFormat.RecordCount > 0 Then
        mrsGridRowFormat("�и�").Value = mclsCommonLib.GetTwipsY(sglHeight)
    End If
    mrsGridRowFormat.Filter = ""

    SetGridRowHeight = True
    
End Function

Public Function SetGridRowLineType(ByVal intRow As Integer, ByVal bytLineType As Byte) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mrsGridRowFormat.Filter = ""
    mrsGridRowFormat.Filter = "�к�=" & intRow
    If mrsGridRowFormat.RecordCount > 0 Then
        mrsGridRowFormat("�߿�").Value = bytLineType
    End If
    
    SetGridRowLineType = True
End Function

Public Function SetGridRowLineStyle(ByVal intRow As Integer, ByVal bytLineStyle As Byte) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mrsGridRowFormat.Filter = ""
    mrsGridRowFormat.Filter = "�к�=" & intRow
    If mrsGridRowFormat.RecordCount > 0 Then
        mrsGridRowFormat("����").Value = bytLineStyle
    End If
    
    SetGridRowLineStyle = True
End Function

Public Function SetGridColFormat(ByVal strColWidth As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strColWidth, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGridColFormat.AddNew
        mrsGridColFormat("�к�").Value = intLoop + 1
        mrsGridColFormat("�п�").Value = mclsCommonLib.GetTwipsY(Val(varTmp(intLoop)))
        mrsGridColFormat("�߿�").Value = 1
        mrsGridRowFormat("����").Value = 0
    Next
    
    SetGridColFormat = True
    
End Function

Public Function SetGridColLineType(ByVal intCol As Integer, ByVal bytLineType As Byte) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mrsGridColFormat.Filter = ""
    mrsGridColFormat.Filter = "�к�=" & intCol
    If mrsGridColFormat.RecordCount > 0 Then
        mrsGridColFormat("�߿�").Value = bytLineType
    End If
    
    SetGridColLineType = True
End Function


Public Function Width() As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglWidth As Single
    Dim intCols As Integer
    
    sglWidth = 0
    mrsGridColFormat.Filter = ""
    intCols = mrsGridColFormat.RecordCount
    For intLoop = 1 To intCols

        mrsGridColFormat.Filter = ""
        mrsGridColFormat.Filter = "�к�=" & intLoop
        If mrsGridColFormat.RecordCount > 0 Then
'
            sglWidth = sglWidth + Val(mrsGridColFormat("�п�").Value)

        End If
    Next
    
    Width = mclsCommonLib.GetCentimetersX(sglWidth)
End Function

Public Function SumGridColWidth(Optional ByVal intCol As Integer) As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglWidth As Single
    Dim intCols As Integer
    
    sglWidth = 0
    mrsGridColFormat.Filter = ""
    intCols = mrsGridColFormat.RecordCount
    For intLoop = 1 To intCols

        mrsGridColFormat.Filter = ""
        mrsGridColFormat.Filter = "�к�=" & intLoop
        If mrsGridColFormat.RecordCount > 0 Then
'
            sglWidth = sglWidth + Val(mrsGridColFormat("�п�").Value)
            
            If intLoop = intCol Then
                SumGridColWidth = sglWidth
                Exit Function
            End If

        End If
    Next
    
    SumGridColWidth = sglWidth
End Function

Public Function Height(Optional ByVal intRow As Integer) As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglHeight As Single
    Dim intRows As Integer
    
    sglHeight = 0
    
    mrsGridRowFormat.Filter = ""
    intRows = mrsGridRowFormat.RecordCount
    For intLoop = 1 To intRows

        mrsGridRowFormat.Filter = ""
        mrsGridRowFormat.Filter = "�к�=" & intLoop
        If mrsGridRowFormat.RecordCount > 0 Then
            
            sglHeight = sglHeight + mclsCommonLib.GetCentimetersY(Val(mrsGridRowFormat("�и�").Value))
            
            If intLoop = intRow Then
                Height = sglHeight
                Exit Function
            End If
            
        End If
    Next
    
    Height = sglHeight
End Function

Public Function SumGridRowHeight(Optional ByVal intRow As Integer) As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglHeight As Single
    Dim intRows As Integer
    
    sglHeight = 0
    
    mrsGridRowFormat.Filter = ""
    intRows = mrsGridRowFormat.RecordCount
    For intLoop = 1 To intRows

        mrsGridRowFormat.Filter = ""
        mrsGridRowFormat.Filter = "�к�=" & intLoop
        If mrsGridRowFormat.RecordCount > 0 Then
'
            sglHeight = sglHeight + Val(mrsGridRowFormat("�и�").Value)
            
            If intLoop = intRow Then
                SumGridRowHeight = sglHeight
                Exit Function
            End If

        End If
    Next
    
    SumGridRowHeight = sglHeight
End Function

Public Function SetGridColData(ByVal intCol As Integer, ByVal strData As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strData, ";")
    
    
    For intLoop = 0 To UBound(varTmp)
        mrsGridData.AddNew
        mrsGridData("�к�").Value = intLoop + 1
        mrsGridData("�к�").Value = intCol
        mrsGridData("����").Value = CStr(varTmp(intLoop))
        mrsGridData("����").Value = mrsGridFormat("����").Value
        mrsGridData("��С").Value = mrsGridFormat("��С").Value
        mrsGridData("ǰ��ɫ").Value = mrsGridFormat("ǰ��ɫ").Value
        mrsGridData("����ɫ").Value = mrsGridFormat("����ɫ").Value
        mrsGridData("ˮƽ����").Value = mrsGridFormat("ˮƽ����").Value
        mrsGridData("��ֱ����").Value = mrsGridFormat("��ֱ����").Value
        mrsGridData("�Զ�����").Value = 0
        mrsGridData("�������").Value = 0
        mrsGridData("�Զ���Ӧ").Value = 0
        mrsGridData("��ת�Ƕ�").Value = 0
'        mrsGridData("ͼ���ļ�").Value = ""
    Next
    
    SetGridColData = True
    
End Function

Public Function SetGridRowData(ByVal intRow As Integer, ByVal strData As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strData, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGridData.AddNew
        mrsGridData("�к�").Value = intRow
        mrsGridData("�к�").Value = intLoop + 1
        mrsGridData("����").Value = CStr(varTmp(intLoop))
        mrsGridData("����").Value = mrsGridFormat("����").Value
        mrsGridData("��С").Value = mrsGridFormat("��С").Value
        mrsGridData("ǰ��ɫ").Value = mrsGridFormat("ǰ��ɫ").Value
        mrsGridData("����ɫ").Value = mrsGridFormat("����ɫ").Value
        mrsGridData("ˮƽ����").Value = mrsGridFormat("ˮƽ����").Value
        mrsGridData("��ֱ����").Value = mrsGridFormat("��ֱ����").Value
        mrsGridData("�Զ�����").Value = 0
        mrsGridData("�������").Value = 0
        mrsGridData("�Զ���Ӧ").Value = 0
        mrsGridData("��ת�Ƕ�").Value = 0
'        mrsGridData("ͼ���ļ�").Value = ""
    Next
    
    SetGridRowData = True
    
End Function

Friend Function FormatData(Optional ByVal strKind As String = "����") As Boolean
    '******************************************************************************************************************
    '���ܣ� ������������
    '������ mrsGridData��- ��Ԫ�����ݣ�mrsGridRowFormat -  �и�ʽ˵����mrsGridColFormat - �и�ʽ˵����
    '       gobjRect - ������������Ͻ���������½�����
    '���أ�
    '******************************************************************************************************************
    Dim strTmp As String
    Dim intRows As Integer
    Dim intCols As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim X0 As Single
    Dim Y0 As Single
    Dim X1 As Single
    Dim Y1 As Single
    Dim varGridLine As Variant
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim objSvrFont As USERFONT
    Dim objSvrRect As USERRECT
    Dim sngOldY0 As Single
    
    On Error GoTo errHand
    
    Call SaveFont(gobjFont, objSvrFont)
    Call SaveRect(gobjRect, objSvrRect)
    
    X0 = gobjRect.X0
    Y0 = gobjRect.Y0
    X1 = gobjRect.X1
    Y1 = gobjRect.Y1
    
    mrsGridRowFormat.Filter = ""
    mrsGridColFormat.Filter = ""
    
    intRows = mrsGridRowFormat.RecordCount
    intCols = mrsGridColFormat.RecordCount
    
    '�Ȼ�������
    '------------------------------------------------------------------------------------------------------------------
    varGridLine = Split(mrsGridFormat("������").Value, ";")
    
    '����Ϻ���
    If Val(varGridLine(0)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.X1 = X1
        gobjRect.Y1 = Y0
        Call AppendPrintData(strKind, "����")
    End If
    
    '���������
    If Val(varGridLine(1)) > 0 Then
        gobjRect.X0 = X1
        gobjRect.Y0 = Y0
        gobjRect.X1 = X1
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "����")
    
    End If
    
    '����º���
    If Val(varGridLine(2)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y1
        gobjRect.X1 = X1
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "����")
    
    End If
    
    '���������
    If Val(varGridLine(3)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.X1 = X0
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "����")
            
    End If
    
    
    '�ڿ򣺺���
    gobjFont.LineWidth = 1
    If Val(varGridLine(4)) > 0 Then
        gobjRect.Y0 = Y0
        gobjRect.X0 = X0
        gobjRect.X1 = X1
        For intRow = 1 To intRows - 1
            mrsGridRowFormat.Filter = ""
            mrsGridRowFormat.Filter = "�к�=" & intRow
            If mrsGridRowFormat.RecordCount > 0 Then
                gobjRect.Y0 = gobjRect.Y0 + Val(mrsGridRowFormat("�и�").Value)
                gobjRect.Y1 = gobjRect.Y0
                If Val(mrsGridRowFormat("�߿�").Value) <> 0 Then
                    gobjFont.LineWidth = Val(mrsGridRowFormat("�߿�").Value)
                    gobjFont.LineStyle = Val(mrsGridRowFormat("����").Value)
                    Call AppendPrintData(strKind, "����")
                End If
            End If
        Next
    End If
        
    '�ڿ�����
    If Val(varGridLine(5)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.Y1 = Y1
        For intCol = 1 To intCols - 1
            mrsGridColFormat.Filter = ""
            mrsGridColFormat.Filter = "�к�=" & intCol
            If mrsGridColFormat.RecordCount > 0 Then
                gobjRect.X0 = gobjRect.X0 + Val(mrsGridColFormat("�п�").Value)
                gobjRect.X1 = gobjRect.X0
                gobjFont.LineWidth = Val(mrsGridColFormat("�߿�").Value)
                gobjFont.LineStyle = Val(mrsGridRowFormat("����").Value)
                Call AppendPrintData(strKind, "����")
            End If
        Next
    End If
    gobjFont.LineWidth = 1
    
    '��д����
    '------------------------------------------------------------------------------------------------------------------
    gobjRect.Y1 = Y0
    For intRow = 1 To intRows
        
        gobjRect.Y0 = gobjRect.Y1
        gobjRect.X0 = X0
        
        mrsGridRowFormat.Filter = ""
        mrsGridRowFormat.Filter = "�к�=" & intRow
        
        If mrsGridRowFormat.RecordCount > 0 Then

            gobjRect.Y1 = gobjRect.Y0 + Val(mrsGridRowFormat("�и�").Value)
            gobjRect.X1 = X0
            
            For intCol = 1 To intCols
                
                mrsGridColFormat.Filter = ""
                mrsGridColFormat.Filter = "�к�=" & intCol
                    
                If mrsGridColFormat.RecordCount > 0 Then
                
                    gobjRect.X0 = gobjRect.X1
                    gobjRect.X1 = gobjRect.X0 + Val(mrsGridColFormat("�п�").Value)
                                        
                    mrsGridData.Filter = ""
                    mrsGridData.Filter = "�к�=" & intRow & " And �к�=" & intCol
                    If mrsGridData.RecordCount > 0 Then
     
                        strTmp = mrsGridData("����").Value
                    
                        gobjFont.Name = mrsGridData("����").Value
                        gobjFont.Size = mrsGridData("��С").Value
                        gobjFont.BackColor = Val(mrsGridData("����ɫ").Value)
                        gobjFont.ForeColor = Val(mrsGridData("ǰ��ɫ").Value)

'                        gobjFont.Bold = False
                        gobjFont.Underline = False
                        Call SetDraw(gobjFont)
                        
                        '<Image>c:\a<Image>"
                        intPos1 = InStr(strTmp, "<Image>")
                        intPos2 = InStr(strTmp, "</Image>")
                        If intPos1 > 0 And intPos2 > 0 And intPos1 < intPos2 Then
                            strTmp = Mid(strTmp, intPos1 + 7, intPos2 - intPos1 - 7)
                            If Dir(strTmp) <> "" And strTmp <> "" Then
                                Call AppendPrintData(strKind, "ͼ��", Val(mrsGridData("ˮƽ����").Value), False, strTmp)
                            End If
                        Else
                        
                            If Val(mrsGridData("ˮƽ����").Value) = 1 Then gobjRect.X0 = gobjRect.X0 + mclsCommonLib.GetTwipsX(0.05)
                            If Val(mrsGridData("��ֱ����").Value) = 1 Then gobjRect.Y0 = gobjRect.Y0 + mclsCommonLib.GetTwipsY(0.05)
                        
                            If (Val(mrsGridData("�������").Value) = 1) = True Then sngOldY0 = gobjRect.Y0
                            
                            Call AppendPrintData(strKind, "�ı�", Val(mrsGridData("ˮƽ����").Value), (Val(mrsGridData("�Զ�����").Value) = 1), strTmp, Val(mrsGridData("��ֱ����").Value), (Val(mrsGridData("�������").Value) = 1), gobjFont.Rows, (Val(mrsGridData("�Զ���Ӧ").Value) = 1), , , Val(mrsGridData("��ת�Ƕ�").Value))
                            
                            If (Val(mrsGridData("�������").Value) = 1) = True Then gobjRect.Y0 = sngOldY0
                        
                            If Val(mrsGridData("ˮƽ����").Value) = 1 Then gobjRect.X0 = gobjRect.X0 - mclsCommonLib.GetTwipsX(0.05)
                            If Val(mrsGridData("��ֱ����").Value) = 1 Then gobjRect.Y0 = gobjRect.Y0 - mclsCommonLib.GetTwipsY(0.05)
                        
                        End If

                        
                    End If
                End If
            Next
        End If
    Next
    
    Call SaveFont(objSvrFont, gobjFont)
    Call SaveRect(objSvrRect, gobjRect)
    
    FormatData = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveFont(objSvrFont, gobjFont)
End Function

Private Sub Class_Initialize()
    Set mclsCommonLib = New clsCommonLib
End Sub

Private Sub Class_Terminate()
    Set mclsCommonLib = Nothing
End Sub