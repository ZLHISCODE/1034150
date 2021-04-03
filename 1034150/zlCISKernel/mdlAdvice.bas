Attribute VB_Name = "mdlAdvice"
Option Explicit

'ˢ������ʱ����Ĳ���״̬
Public Enum TYPE_PATI_State
    ps��Ժ = 0
    psԤ�� = 1
    ps��Ժ = 2
    ps���� = 3          'ҽ��վ:�����ﲡ��(��Ժ)
    ps���� = 4          'ҽ��վ:�ѻ��ﲡ��
    ps���ת�� = 5      'ҽ��վ:���ת�ƻ�ת�����Ĳ���(��Ժ)
    ps��ת�� = 6        'ҽ��վ:��ƴ���ס��ת��������������
End Enum

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function StringMask(ByVal strText As String, ByVal strMask As String) As Boolean
'���ܣ�����ַ����Ƿ�ֻ����ָ�����ַ�
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strMask, Mid(strText, i, 1)) = 0 Then Exit Function
    Next
    StringMask = True
End Function

Public Function BillExpend(ByVal strNO As String) As Boolean
'���ܣ��жϹҺŵ��Ƿ��Ѿ�������Ч�Һ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    '����������������������Һ���Ч�����Ĳ��˲��������ʾ�������
    If Val(zldatabase.GetPara(210, glngSys)) = 1 Then Exit Function
    '��ʱ����
    strSql = "Select  Sysdate-����ʱ�� as ���,���� From ���˹Һż�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=1"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISWork", strNO)
    If Not rsTmp.EOF Then
        BillExpend = Val(rsTmp!���) > IIF(Val("" & rsTmp!����) = 1, IIF(gint����Һ����� = 0, 1, gint����Һ�����), IIF(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckOutAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ���鲡���Ƿ����´��˳�Ժҽ��
    Dim strSql As String, rsTmp As Recordset
    
    strSql = "Select 1 from ����ҽ����¼ A,������ĿĿ¼ B Where a.������ĿID=b.ID And a.����ID=[1] And a.��ҳID=[2] And b.���='Z' And b.��������='5' And a.ҽ��״̬<>4  and nvl(A.Ӥ��,0)=0"
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, gstrSysName, lng����ID, lng��ҳID)
    CheckOutAdvice = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExeTimeValid(ByVal strTime As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, Optional ByVal bln���� As Boolean) As Boolean
'���ܣ����ָ����ִ��ʱ���Ƿ�Ϸ�
    Dim arrTime() As String, strTmp As String, i As Integer
    Dim strPreTime As String, intPreDay As Long, intCurDay As Long
    
    If strTime = "" Then
        If str�����λ = "����" Then ExeTimeValid = True
        Exit Function
    End If
    
    If str�����λ = "��" Then
        '1/8:00-3/15:00-5/9:00��1/8:00-3/15-5/9:00
        If Not StringMask(strTime, "0123456789:-/") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If bln���� Then
            If Not Between(UBound(arrTime) + 1, 1, intƵ�ʴ���) Then Exit Function
        Else
            If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
        End If
        
        For i = 0 To UBound(arrTime)
            If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
            '���ڲ���
            strTmp = Split(arrTime(i), "/")(0)
            If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
            intCurDay = Val(strTmp)
            If intCurDay < 1 Or intCurDay > 7 Then Exit Function
            If intPreDay <> 0 Then
                If intCurDay < intPreDay Then Exit Function
            End If
            
            '����ʱ�䲿��
            strTmp = Split(arrTime(i), "/")(1)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
            If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            
            strPreTime = Format(strTmp, "HH:mm")
            intPreDay = intCurDay
        Next
    ElseIf str�����λ = "��" Then
        If intƵ�ʼ�� = 1 Then
            '8:00-12:00-14:00��8:00-12-14:00
            If Not StringMask(strTime, "0123456789:-") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If bln���� Then
                If Not Between(UBound(arrTime) + 1, 1, intƵ�ʴ���) Then Exit Function
            Else
                If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                strTmp = arrTime(i)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
                If strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                strPreTime = Format(strTmp, "HH:mm")
            Next
        Else
            '1/8:00-1/15:00-2/9:00��1/8:00-1/15-2/9:00
            If Not StringMask(strTime, "0123456789:-/") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If bln���� Then
                If Not Between(UBound(arrTime) + 1, 1, intƵ�ʴ���) Then Exit Function
            Else
                If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
                '�����������
                strTmp = Split(arrTime(i), "/")(0)
                If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
                intCurDay = Val(strTmp)
                If intCurDay < 1 Or intCurDay > intƵ�ʼ�� Then Exit Function
                If intPreDay <> 0 Then
                    If intCurDay < intPreDay Then Exit Function
                End If
                
                '����ʱ�䲿��
                strTmp = Split(arrTime(i), "/")(1)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
                If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                
                strPreTime = Format(strTmp, "HH:mm")
                intPreDay = intCurDay
            Next
        End If
    ElseIf str�����λ = "Сʱ" Then
        '1:30-2-3:30
        If Not StringMask(strTime, "0123456789:-") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
        
        For i = 0 To UBound(arrTime)
            strTmp = arrTime(i)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) < 1 Or Val(Split(strTmp, ":")(0)) > intƵ�ʼ�� Or Split(strTmp, ":")(0) = "" Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
            If strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            strPreTime = Format(strTmp, "HH:mm")
        Next
    End If
    
    ExeTimeValid = True
End Function

Public Function GetWeekBase(ByVal datDate As Date) As Date
'���ܣ���ȡָ��ʱ���������ڵ�����һ��ʱ��
    'Oracle:Select Sysdate-(To_Number(To_Char(Sysdate,'D'))-1)+1 From Dual
    GetWeekBase = Format(datDate - (Weekday(datDate, vbMonday) - 1), "yyyy-MM-dd 00:00:00")
End Function

Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ��ʱ���Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '������δ���û���ͣ��ʱ��ֹͣ
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function


Public Function GetMaxBedLen(Optional lng����ID As Long, Optional bln���� As Boolean) As Integer
'���ܣ���ȡָ�����ŵĴ�λ�ŵ���󳤶�
'������lng����ID=����ID�����ID,Ϊ0��ʾ���в��������
'      blnռ��=�Ƿ�ֻ�ܱ�ռ�õĴ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    If Not bln���� Or lng����ID = 0 Then
        strSql = "Select Max(LengthB(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIF(lng����ID = 0, " is Not NULL", "= [1] ")
    Else
        strSql = "Select Max(LengthB(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIF(lng����ID = 0, " is Not NULL", "= [1] ")
    End If
    
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng����ID)
    
    If Not rsTmp.EOF Then GetMaxBedLen = IIF(IsNull(rsTmp!����), 0, rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DateIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ�������Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
'˵��������ʱ���ж�,����ͣ���ڰ���ʼ����ֹ�����ж�
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Format(Split(arrPause(i), ",")(0), "yyyy-MM-dd")
        strEnd = Format(Split(arrPause(i), ",")(1), "yyyy-MM-dd")
        If strEnd = "" Then strEnd = "3000-01-01" '������δ���û���ͣ��ʱ��ֹͣ
        If strEnd > strBegin Then
            If Between(Format(vDate, "yyyy-MM-dd"), strBegin, _
                Format(DateAdd("d", -1, CDate(strEnd)), "yyyy-MM-dd")) Then
                DateIsPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function TimeisLastPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ��ʱ���Ƿ������һ����ͣ��ʱ����,�����һ����ͣû������
'˵������Ϊ���������,�������û����ֹʱ��,ĳЩ�������ѭ��
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    
    For i = UBound(arrPause) To 0 Step -1
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then
            strEnd = "3000-01-01 00:00:00"
            If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
                TimeisLastPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function Calc�����ֽ�ʱ��(lng���� As Long, ByVal dat��ʼʱ�� As Date, dat��ֹʱ�� As Date, strPause As String, _
    ByVal strִ��ʱ�� As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal dat�������� As Date) As String
'���ܣ�������������εķֽ�ִ��ʱ��,Ҫ��<=��ֹʱ�估������ͣʱ�����
'������dat��ʼʱ��=ҽ���Ŀ�ʼִ��ʱ��
'      dat��ֹʱ��=ҽ����ִ����ֹʱ��,û��ʱ����"3000-01-01"
'      strPause=ҽ������ͣʱ���
'      dat��������=��������ʱ��������
'���أ�1."ʱ��1,ʱ��2,...."(yyyy-MM-dd HH:mm:ss)
'      2.lng����=ʵ���ܹ��ֽ�Ĵ���
'˵����1.��Ϊ��ֹʱ�������,��˷ֽ������ʱ���������С��Ҫ�ֽ�Ĵ���
'      2.�������Ǽٶ���ִ��ʱ�估Ƶ��������ȫ��ȷ������¼��㡣
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrFirst As Variant, arrNormal As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(strִ��ʱ��, ",") > 0 Then
        arrNormal = Split(Split(strִ��ʱ��, ",")(1), "-")
        arrFirst = Split(Split(strִ��ʱ��, ",")(0), "-")
    Else
        arrNormal = Split(strִ��ʱ��, "-")
        arrFirst = Array()
    End If
    
    vCurTime = dat��ʼʱ��
    
    If str�����λ = "��" Then
        vCurTime = GetWeekBase(dat��ʼʱ��)
        
        Do While lng���� > 0
            blnFirst = (GetWeekBase(vCurTime) = GetWeekBase(dat��������)) And dat�������� <> Empty And UBound(arrFirst) <> -1
            arrTime = IIF(blnFirst, arrFirst, arrNormal)

            '1/8:00-3/15:00-5/9:00
            For i = 1 To intƵ�ʴ���
                If i - 1 <= UBound(arrTime) Then '���ܿ��ܴ�������
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime > dat��ֹʱ�� Then
                        Exit Do
                    ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                        Exit Do
                    ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        lng���� = lng���� - 1
                        If lng���� = 0 Then Exit Do
                    End If
                End If
            Next
            vCurTime = vCurTime + 7
        Loop
    ElseIf str�����λ = "��" Then
        Do While lng���� > 0
            blnFirst = (Int(vCurTime) = Int(dat��������)) And dat�������� <> Empty And UBound(arrFirst) <> -1
            arrTime = IIF(blnFirst, arrFirst, arrNormal)
        
            If intƵ�ʼ�� = 1 Then
                '8:00-12:00-14:00��8-12-14
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        
                        If vTmpTime > dat��ֹʱ�� Then
                            Exit Do
                        ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                            Exit Do
                        ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            lng���� = lng���� - 1
                            If lng���� = 0 Then Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime > dat��ֹʱ�� Then
                            Exit Do
                        ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                            Exit Do
                        ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            lng���� = lng���� - 1
                            If lng���� = 0 Then Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = vCurTime + intƵ�ʼ��
        Loop
    ElseIf str�����λ = "Сʱ" Then
        '10:00-20:00-40:00��10-20-40��02:30
        arrTime = arrNormal
        Do While lng���� > 0
            For i = 1 To intƵ�ʴ���
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime > dat��ֹʱ�� Then
                    Exit Do
                ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                    Exit Do
                ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    lng���� = lng���� - 1
                    If lng���� = 0 Then Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + intƵ�ʼ�� / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str�����λ = "����" Then
        '��ִ��ʱ��
        Do While lng���� > 0
            vTmpTime = vCurTime
            
            If vTmpTime > dat��ֹʱ�� Then
                Exit Do
            ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                Exit Do
            ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                lng���� = lng���� - 1
                If lng���� = 0 Then Exit Do
            End If

            vCurTime = Format(vCurTime + intƵ�ʼ�� / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If

    lng���� = UBound(Split(Mid(strDetailTime, 2), ",")) + 1
    Calc�����ֽ�ʱ�� = Mid(strDetailTime, 2)
End Function

Public Function Calc���ڷֽ�ʱ��(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal strִ��ʱ�� As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal dat�������� As Date) As String
'���ܣ���ʱ��μ�����εķֽ�ִ��ʱ�估����
'������datBegin-datEnd=Ҫ�����ʱ���,����datBeginӦΪÿ�����ڵĿ�ʼ��׼ʱ��
'      strPause=��ͣ��ʱ���
'      dat��������=��������ʱ��������
'���أ�"ʱ��1,ʱ��2,...."(yyyy-MM-dd HH:mm:ss),ʱ�������Ϊ����
'˵����1.ʱ�����Ҫ�ų���ͣ��ʱ���,����������˶�����
'      2.�������Ǽٶ���ִ��ʱ�估Ƶ��������ȫ��ȷ������¼��㡣
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrNormal As Variant, arrFirst As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(strִ��ʱ��, ",") > 0 Then
        arrNormal = Split(Split(strִ��ʱ��, ",")(1), "-")
        arrFirst = Split(Split(strִ��ʱ��, ",")(0), "-")
    Else
        arrNormal = Split(strִ��ʱ��, "-")
        arrFirst = Array()
    End If
        
    vCurTime = datBegin
    
    If str�����λ = "��" Then
        vCurTime = GetWeekBase(datBegin)
        If dat�������� <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (vCurTime = GetWeekBase(dat��������))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIF(blnFirst, arrFirst, arrNormal)
            blnFirst = False
                        
            '1/8:00-3/15:00-5/9:00
            For i = 1 To intƵ�ʴ���
                If i - 1 <= UBound(arrTime) Then '���ܿ��ܴ�������
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "��" Then
        If dat�������� <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (Int(vCurTime) = Int(dat��������))
        Else
            blnFirst = False
        End If
        
        Do While vCurTime <= datEnd
            arrTime = IIF(blnFirst, arrFirst, arrNormal)
            blnFirst = False
            
            If intƵ�ʼ�� = 1 Then
                '8:00-12:00-14:00��8-12-14
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + intƵ�ʼ��, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "Сʱ" Then
        '10:00-20:00-40:00��10-20-40��02:30
        arrTime = arrNormal
        Do While vCurTime <= datEnd
            For i = 1 To intƵ�ʴ���
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime >= Format(datBegin, "yyyy-MM-dd HH:mm:ss") And vTmpTime <= Format(datEnd, "yyyy-MM-dd HH:mm:ss") Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + intƵ�ʼ�� / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str�����λ = "����" Then
        '��ִ��ʱ��
        Do While vCurTime <= datEnd
            vTmpTime = vCurTime
            
            If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                If Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                End If
            ElseIf vTmpTime > datEnd Then
                Exit Do
            End If

            vCurTime = Format(vCurTime + intƵ�ʼ�� / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If
    
    Calc���ڷֽ�ʱ�� = Mid(strDetailTime, 2)
End Function

Public Function CalcȱʡҩƷ����(ByVal dbl���� As Double, ByVal int�Ƴ� As Integer, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, Optional ByVal strִ��ʱ�� As String, _
    Optional ByVal dbl����ϵ�� As Double, Optional ByVal dbl��װϵ�� As Double, Optional ByVal int���� As Integer, Optional ByVal dbl�״����� As Double) As Double
'���ܣ����Ƴ̼��������Լ���ҩƷ������ȱʡ����(���䷽ȱʡ����)
'������dbl����=��������λ��һ������
'      int�Ƴ�=һ���Ƴ̵�����
'      int����=0-�ɷ���,1-������,2-һ����(��ʱʧЧ),-N-N���ڷ���ʹ����Ч
'      dbl��װϵ��=�����װ��סԺ��װ
'���أ���סԺ��λ�����ҩƷ����
'˵����
'     1.ҩƷ������������������סԺ��װ���Եġ�
'     2.dbl����ϵ��,dbl��װϵ��,int����=��ҩ������,ֻ���㸶��
    Dim dbl��� As Double, dbl���� As Double
    Dim dblʣ�� As Double, dblOne As Double
    Dim intStep As Integer, dblEnd As Double
    Dim arrTime() As String, strBegin As String
    Dim strTime As String, i As Integer, j As Integer
    Dim dblһ������ As Double
    
    '�Ƴ̲���һ��Ƶ������ʱ�Ͳ����Ƴ�
    If str�����λ = "��" Then
        If int�Ƴ� < 7 Then int�Ƴ� = 1
    ElseIf str�����λ = "��" Then
        If int�Ƴ� < intƵ�ʼ�� Then int�Ƴ� = 1
    ElseIf str�����λ = "Сʱ" Then
        If int�Ƴ� < intƵ�ʼ�� / 24 Then int�Ƴ� = 1
    ElseIf str�����λ = "����" Then
        If int�Ƴ� < intƵ�ʼ�� / (24 * 60) Then int�Ƴ� = 1
    End If
    
    'һ��Ƶ�����ڵĴ���(����)
    If str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / 7
    ElseIf str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / intƵ�ʼ��
    ElseIf str�����λ = "Сʱ" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * 24
    ElseIf str�����λ = "����" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * (24 * 60)
    End If
    
    If dbl����ϵ�� = 0 And dbl��װϵ�� = 0 Then
        '��ҩ����(����) = ����*�Ƴ�*(Ƶ�ʴ���/Ƶ�ʼ��)
        dbl���� = IntEx(int�Ƴ� * dbl���)
    Else
        'ҩƷ�������� = ����/סԺ��װ(����*�Ƴ�*(Ƶ�ʴ���/Ƶ�ʼ��))
        If int���� = 0 Then
            '�ɷ���
            dbl���� = dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��
        ElseIf int���� = 1 Then
            '������
            dbl���� = IntEx(dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��)
        ElseIf int���� = 2 Then
            'һ����(��ʱʧЧ)
            dbl���� = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * IntEx(int�Ƴ� * dbl���)
        ElseIf int���� < 0 Then
            'ABS(int����)���ڷ���ʹ����Ч(�����������)
            If strִ��ʱ�� <> "" Then
                'һ������/סԺ��װ�ļ���
                dblOne = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��)
                dblһ������ = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��)
                'ȱʡִ�еĴ�����ʱ��ֽ�
                strTime = Calc�����ֽ�ʱ��(IntEx(int�Ƴ� * dbl���), Date, CDate("3000-01-01"), "", strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                If strTime <> "" Then
                    arrTime = Split(strTime, ",")
                    dblʣ�� = dblOne: dbl���� = dblһ������
                    strBegin = arrTime(0)
                    
                    '��������
                    For i = 0 To UBound(arrTime)
                        If dblʣ�� < dbl���� Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                            If CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                                dblʣ�� = dblOne
                            Else
                                dblʣ�� = dblʣ�� + dblOne
                            End If
                            dbl���� = dbl���� + dblһ������
                            strBegin = arrTime(i)
                        End If
                        dblʣ�� = dblʣ�� - dbl����
                        If dblʣ�� >= dbl����ϵ�� * dbl��װϵ�� Then
                            '���ʣ����Ǵ�����ƿ�����ȥ����
                            dbl���� = dbl���� - Int(dblʣ�� / dbl����ϵ�� / dbl��װϵ��)
                            dblʣ�� = dblʣ�� Mod (dbl����ϵ�� * dbl��װϵ��)
                        End If
                    Next
                End If
            End If
        End If
    End If
    If dbl���� > 0 And dbl�״����� > 0 Then
        dbl���� = dbl���� + (dbl�״����� - dbl����) / dbl����ϵ�� / dbl��װϵ��
    End If
    CalcȱʡҩƷ���� = dbl����
End Function

Public Function CalcȱʡҩƷ����(ByVal dbl���� As Double, ByVal dbl���� As Double, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal dbl����ϵ�� As Double, Optional ByVal dbl��װϵ�� As Double, Optional ByVal int���� As Integer) As Long
'���ܣ�������������������ҩƷ���Լ�����ҩ����
'������dbl����=�û����������
'      dbl����=��������λ��һ������
'      int����=0-�ɷ���,1-������,2-һ����(��ʱʧЧ),-N-N���ڷ���ʹ����Ч
'      dbl��װϵ��=�����װ��סԺ��װ
'���أ���ҩ����(��ҩ����������)
    Dim dbl��� As Double
    Dim lng���� As Long
    
    'һ��Ƶ�����ڵĴ���(����)
    If str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / 7
    ElseIf str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / intƵ�ʼ��
    ElseIf str�����λ = "Сʱ" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * 24
    ElseIf str�����λ = "����" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * (24 * 60)
    End If
    
    If int���� = 0 Then
        '�ɷ���
        'dbl���� = dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��
        lng���� = Format(dbl���� * dbl��װϵ�� * dbl����ϵ�� / dbl���� / dbl���, "0")
    ElseIf int���� = 1 Then
        '������
        'dbl���� = IntEx(dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��)
        lng���� = Format(dbl���� * dbl��װϵ�� * dbl����ϵ�� / dbl���� / dbl���, "0")
    ElseIf int���� = 2 Then
        'һ����(��ʱʧЧ)
        'dbl���� = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * IntEx(int�Ƴ� * dbl���)
        lng���� = Format(dbl���� / IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) / dbl���, "0")
    ElseIf int���� < 0 Then
        'ABS(int����)���ڷ���ʹ����Ч(�����������)
        lng���� = Format(dbl���� * dbl��װϵ�� * dbl����ϵ�� / dbl���� / dbl���, "0")
    End If

    CalcȱʡҩƷ���� = lng����
End Function

Public Function Calc����ҩƷ����(ByVal dat��ʼִ��ʱ�� As Date, lng���� As Long, str�ֽ�ʱ�� As String, _
    ByVal dbl���� As Double, ByVal dbl����ϵ�� As Double, ByVal dbl��װϵ�� As Double, _
    ByVal int���� As Integer, ByVal dat��ֹʱ�� As Date, ByVal strPause As String, ByVal strִ��ʱ�� As String, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal blnLimit As Boolean, Optional ByVal dbl�״����� As Double, Optional ByVal dat�ϴ�ִ��ʱ�� As Date) As Double
'���ܣ������ʹ������������Լ����ҩ����
'������dat��ʼִ��ʱ��=ҽ���Ŀ�ʼִ��ʱ��,���ڼ�����һִ�����ڿ�ʼ��׼ʱ��
'      lng����=���μƻ�Ҫ���͵Ĵ���
'      dbl����=��������λ��һ������
'      int����=0-�ɷ���,1-������,2-һ����(��ʱʧЧ),-N-N���ڷ���ʹ����Ч(��24Сʱ����)
'      dbl��װϵ��=�����װ��סԺ��װ
'      blnLimit=�Ƿ�ʱ�����Ƽ����ҩ;��������ʣ�ಿ��
'���в������ڲ�����ҩƷ����(����-N��)��
'      str�ֽ�ʱ��=���η��ͼƻ�ִ�еķֽ�ʱ��,�������Ӧ
'      strPause=ҽ������ͣʱ���
'      dat��ֹʱ��=ҽ����ִ����ֹʱ��,û��ʱ����"3000-01-01"
'���أ�1.������/סԺ��λ�����ҩƷ����
'      2.lng����=������ҩƷ(����-N�ͷ���ҩƷ)������ʵ��ִ�д���(����)
'      3.str�ֽ�ʱ��=������ҩƷ(����-N�ͷ���ҩƷ)�����ķֽ�ʱ��(����)
'˵����ҩƷ������������������סԺ��װ���Եġ�
    Dim dbl���� As Double, dblʣ�� As Double
    Dim arrTime() As String, dblOne As Double
    Dim strBegin As String, datBase As Date
    Dim strTmp As String, i As Long
    Dim blnIsFirst As Boolean
    
    'ע��һЩ�ط���Val����Ϊ��������Double��ĳЩ�ط��ж�ʱ���ڲ����������⣬���±���0.9<>0.9
    If int���� = 0 Then
        '�ɷ���
        dbl���� = Val(dbl���� * lng���� / dbl����ϵ�� / dbl��װϵ��)
        '����ϴ�ִ��ʱ��ΪNULL��˵�������״�
        If dat�ϴ�ִ��ʱ�� = CDate(0) And dbl�״����� > 0 Then
            dbl���� = Val(dbl���� + (dbl�״����� - dbl����) / dbl����ϵ�� / dbl��װϵ��)
        End If
    ElseIf int���� = 1 Then
        '������
        dbl���� = Val(dbl���� * lng���� / dbl����ϵ�� / dbl��װϵ��)
        '����ϴ�ִ��ʱ��ΪNULL��˵�������״�
        If dat�ϴ�ִ��ʱ�� = CDate(0) And dbl�״����� > 0 Then
            dbl���� = Val(dbl���� + (dbl�״����� - dbl����) / dbl����ϵ�� / dbl��װϵ��)
        End If
        dbl���� = Val(IntEx(dbl����))
        '�����������ʱ,����ľ�����ʹ��,�Ӷ�ʹ���ʹ�������
        If Not blnLimit Then
            dblʣ�� = Val(dbl���� * dbl��װϵ�� * dbl����ϵ�� - dbl���� * lng����)
            If dblʣ�� >= dbl���� And dbl���� <> 0 Then
                'ʣ�����ۿ���ִ�еĴ���
                i = Int(Val(dblʣ�� / dbl����))
                'ʣ��ʵ�ʿ���ִ�еĴ�����ʱ��ֽ�(����ֹʱ������)
                arrTime = Split(str�ֽ�ʱ��, ",")
                datBase = Calc�����ڿ�ʼʱ��(dat��ʼִ��ʱ��, CDate(arrTime(UBound(arrTime))), intƵ�ʼ��, str�����λ)
                
                '��������չʱ��ʱ,���һ����������ִ�е�ʱ�䲻�ټ���,����ͣ����
                strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
                If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                
                strTmp = Calc�����ֽ�ʱ��(i, datBase, dat��ֹʱ��, strPause, strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ, dat��ʼִ��ʱ��)
                If strTmp <> "" Then
                    lng���� = lng���� + i
                    str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & strTmp
                End If
            End If
        End If
    ElseIf int���� = 2 Then
        'һ����(��ʱʧЧ)
        dbl���� = Val(IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * lng����)
        '����ϴ�ִ��ʱ��ΪNULL��˵�������״�
        If dat�ϴ�ִ��ʱ�� = CDate(0) And dbl�״����� > 0 Then
            dbl���� = Val(dbl���� + IntEx(dbl�״����� / dbl����ϵ�� / dbl��װϵ��) - IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��))
        End If
    ElseIf int���� < 0 Then
        'ABS(int����)���ڷ���ʹ����Ч(�����������)
        arrTime = Split(str�ֽ�ʱ��, ",")
        strBegin = arrTime(0)
        
        'һ������/סԺ��װ�ļ���(������λ)
        dblOne = Val(IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��))
        'һ������/סԺ��װ�ļ���(��װ��λ)
        dbl���� = Val(IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��))
        '����ϴ�ִ��ʱ��ΪNULL��˵�������״�
        If dat�ϴ�ִ��ʱ�� = CDate(0) And dbl�״����� > 0 Then
            dbl���� = Val(IntEx(dbl�״����� / dbl����ϵ�� / dbl��װϵ��))
            dblOne = Val(IntEx(dbl�״����� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��))
            blnIsFirst = True
        End If
         '��������
        dblʣ�� = dblOne
        For i = 0 To UBound(arrTime)
            '��һ��ѭ���϶���,���Բ���������
            If dblʣ�� < IIF(blnIsFirst, dbl�״�����, dbl����) Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                If CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                    dblʣ�� = dblOne
                    dbl���� = dbl���� + IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��)
                Else
                    If dblʣ�� + dbl����ϵ�� * dbl��װϵ�� >= IIF(blnIsFirst, dbl�״�����, dbl����) Then
                        'ֻ��ʣ���һ����װ��λ����
                        dblʣ�� = dblʣ�� + dbl����ϵ�� * dbl��װϵ��
                        dbl���� = dbl���� + 1
                    Else
                        '��Ҫʣ���һ�ΰ�װ��λ�Ź�
                        dblʣ�� = dblʣ�� + dblOne
                        dbl���� = dbl���� + IntEx(IIF(blnIsFirst, dbl�״�����, dbl����) / dbl����ϵ�� / dbl��װϵ��)
                    End If
                End If
                strBegin = arrTime(i)
            End If
            dblʣ�� = dblʣ�� - IIF(blnIsFirst, dbl�״�����, dbl����)
            If blnIsFirst Then
                blnIsFirst = False
                dblOne = Val(IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��))
            End If
        Next
        
        'ʣ�ಿ�ּ�������Ч���ڰ����������,�Ӷ�ʹ���ʹ�������
        If Not blnLimit Then
            If dblʣ�� >= dbl���� And dbl���� <> 0 Then
                'ʣ�����ۿ���ִ�еĴ���
                i = Int(Val(dblʣ�� / dbl����))
                'ʣ��ʵ�ʿ���ִ�еĴ�����ʱ��ֽ�(����ֹʱ������)
                datBase = Calc�����ڿ�ʼʱ��(dat��ʼִ��ʱ��, CDate(arrTime(UBound(arrTime))), intƵ�ʼ��, str�����λ)
                
                '��������չʱ��ʱ,���һ����������ִ�е�ʱ�䲻�ټ���,����ͣ����
                strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
                If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                
                strTmp = Calc�����ֽ�ʱ��(i, datBase, dat��ֹʱ��, strPause, strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ, dat��ʼִ��ʱ��)
                If strTmp <> "" Then
                    arrTime = Split(strTmp, ",")
                    For i = 0 To UBound(arrTime)
                        If dblʣ�� < dbl���� Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                            Exit For
                        End If
                        lng���� = lng���� + 1
                        str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & arrTime(i)
                        dblʣ�� = dblʣ�� - dbl����
                    Next
                End If
            End If
        End If
    End If
    
    Calc����ҩƷ���� = dbl����
End Function

Public Function Calc�����ڿ�ʼʱ��(ByVal dat��ʼִ��ʱ�� As Date, ByVal datĳ��ִ��ʱ�� As Date, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As Date
'���ܣ����ݳ�����ĳ��ִ��ʱ�䣬�õ����ڸ������ڵĿ�ʼ��׼ʱ��
    Dim datBegin As Date, datCurr As Date
    
    datCurr = dat��ʼִ��ʱ��
    datBegin = datCurr
    If str�����λ = "��" Then datCurr = GetWeekBase(datCurr)
    
    If str�����λ = "" Then Exit Function
    Do While datCurr <= datĳ��ִ��ʱ��
        datBegin = datCurr
        If str�����λ = "��" Then
            datCurr = datCurr + 7
        ElseIf str�����λ = "��" Then
            datCurr = datCurr + intƵ�ʼ��
        ElseIf str�����λ = "Сʱ" Then
            datCurr = DateAdd("h", intƵ�ʼ��, datCurr)
        ElseIf str�����λ = "����" Then
            datCurr = DateAdd("n", intƵ�ʼ��, datCurr)
        End If
    Loop
    Calc�����ڿ�ʼʱ�� = datBegin
End Function

Public Function Trim�ֽ�ʱ��(ByVal lng���� As Long, ByVal str�ֽ�ʱ�� As String) As String
'���ܣ���ҽ��ִ�еķֽ�ʱ�䰴�������нض�
    Dim arrTime() As String, strTmp As String, i As Long
    
    arrTime = Split(str�ֽ�ʱ��, ",")
    For i = 0 To lng���� - 1
        strTmp = strTmp & "," & arrTime(i)
    Next
    Trim�ֽ�ʱ�� = Mid(strTmp, 2)
End Function

Public Function Calc�����Գ�������(ByVal datBegin As Date, ByVal datEnd As Date, _
    ByVal str�ϴ�ִ��ʱ�� As String, ByVal strִ����ֹʱ�� As String, _
    ByVal strPause As String, Optional str�״�ʱ�� As String, _
    Optional strĩ��ʱ�� As String, Optional str�ֽ�ʱ�� As String) As Long
'���ܣ��Գ����Է�ҩ��������������Ӧ�÷��͵Ĵ���,����ĩʱ��
'������str�ϴ�ִ��ʱ��=��һ�����ڱ��η��͵Ŀ�ʼʱ��
'      strִ����ֹʱ��=
'���أ����θ�ҽ�����͵Ĵ���
'      str�״�ʱ��,strĩ��ʱ��=����yyyy-MM-dd HH:mm:ss
'˵���������Գ���������ÿ�췢��һ�δ���,���������봲λ������(��ͣʱ����ʼ����ֹ)
    Dim curDate As Date, lng���� As Long, blnSend As Boolean
    
    str�״�ʱ�� = "": strĩ��ʱ�� = "": str�ֽ�ʱ�� = ""
    curDate = CDate(Format(datBegin, "yyyy-MM-dd"))
    Do While curDate <= CDate(Format(datEnd, "yyyy-MM-dd"))
        If Not DateIsPause(curDate, strPause) Then
            blnSend = True
            If str�ϴ�ִ��ʱ�� <> "" Then
                If Format(curDate, "yyyy-MM-dd") <= Format(str�ϴ�ִ��ʱ��, "yyyy-MM-dd") Then
                    blnSend = False 'Ӧ�����ϴ�ִ��ʱ���ִ��
                End If
            End If
            If strִ����ֹʱ�� <> "" Then
                If Format(curDate, "yyyy-MM-dd") > Format(strִ����ֹʱ��, "yyyy-MM-dd") Then
                    blnSend = False 'ӦС�ڵ���ִ����ֹʱ���ִ��
                End If
            End If
            If blnSend Then
                lng���� = lng���� + 1
                If lng���� = 1 Then
                    str�״�ʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00") '��Ϊ���ִ��
                    If str�״�ʱ�� < Format(datBegin, "yyyy-MM-dd HH:mm:ss") Then
                        str�״�ʱ�� = Format(datBegin, "yyyy-MM-dd HH:mm:ss")
                    End If
                    strĩ��ʱ�� = str�״�ʱ��
                    str�ֽ�ʱ�� = str�״�ʱ��
                Else
                    strĩ��ʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00")
                    str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & strĩ��ʱ��
                End If
            End If
        End If
        curDate = curDate + 1
    Loop
    
    Calc�����Գ������� = lng����
End Function

 Public Function Calc������������(ByVal dbl���� As Double, ByVal dbl���� As Double, ByVal dbl����ϵ�� As Double, ByVal dbl��װϵ�� As Double, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As Double
'���ܣ�����ָ����������������Ƶ�ʼ���ҩƷ����ʹ�õ�����
    Dim dbl��� As Double
    Dim dbl�ܵ��� As Double
    
    'һ��Ƶ�����ڵĴ���(����)
    If str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / 7
    ElseIf str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / intƵ�ʼ��
    ElseIf str�����λ = "Сʱ" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * 24
    ElseIf str�����λ = "����" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * (24 * 60)
    End If
    
    dbl�ܵ��� = dbl���� * dbl��װϵ�� * dbl����ϵ��
    
    Calc������������ = dbl�ܵ��� / dbl���� / dbl���
End Function

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str���� As String, ByVal curʣ���� As Currency, _
    ByVal cur���ս�� As Currency, ByVal cur���ʽ�� As Currency, ByVal cur������� As Currency, _
    ByVal str�շ���� As String, ByVal str������� As String, str�ѱ���� As String, _
    intWarn As Integer, Optional ByVal bln���� As Boolean, _
    Optional blnNotCheck��� As Boolean = False) As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:rsWarn=���������������õļ�¼��(�ò��˲���,�����ֺ���ҽ��)
'     str�շ����=��ǰҪ�������,���ڷ��౨��
'     str�������=�������,������ʾ
'     bln����=���ɻ��۷���ʱ�ı��������ƾ���Ƿ��ǿ�Ƽ���Ȩ��ʱ�Ĵ���
'     intWarn=�Ƿ���ʾѯ���Ե���ʾ,-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��
'     blnNotCheck���:���������м��(��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
'����:str�ѱ����="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
'     intWarn=����ѯ������ʾ�е�ѡ����,0=Ϊ��,1-Ϊ��
'     0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
    Dim bln�ѱ��� As Boolean, byt��־ As Byte
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str���� As String, i As Long
    
    BillingWarn = 0
    
    '�����������:NULL��û������,0�������˵�
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str�շ����) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־1 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str�շ����) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־2 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str�շ����) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str������� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־3 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ�����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    'ʾ����"-" �� ",ABC,567,DEF"
    '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
    bln�ѱ��� = InStr(str�ѱ����, str�շ����) > 0 Or str�ѱ���� Like "-*"
    
    If bln�ѱ��� Then '��intWarn = -1ʱ,Ҳ��ǿ���ٱ���
        If byt��־ = 2 Then
            If str�ѱ���� Like "-*" Then
                byt�ѱ���ʽ = IIF(Right(str�ѱ����, 1) = "��", 2, 1)
            Else
                arrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str�շ����) > 0 Then
                        byt�ѱ���ʽ = IIF(Right(arrTmp(i), 1) = "��", 2, 1)
                        'Exit For 'ȡ��˵����סԺ����ģ��
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str������� <> "" Then str������� = """" & str������� & """����"
    str���� = IIF(cur������� = 0, "", "(��������:" & Format(cur�������, "0.00") & ")")
    curʣ���� = curʣ���� + cur������� - cur���ʽ��
    cur���ս�� = cur���ս�� + cur���ʽ��
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If curʣ���� < 0 Then
                        byt��ʽ = 2
                        If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�,�����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    ElseIf curʣ���� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ������ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If curʣ���� < 0 Then
                            byt��ʽ = 2
                            If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�,�����ò��˼�����", frmParent)
                                    If vMsg = vbNo Or vMsg = vbCancel Then
                                        If vMsg = vbCancel Then intWarn = 0
                                        BillingWarn = 2
                                    ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                        If vMsg = vbIgnore Then intWarn = 1
                                        BillingWarn = 4
                                    End If
                                Else
                                    If intWarn = 0 Then
                                        BillingWarn = 2
                                    ElseIf intWarn = 1 Then
                                        BillingWarn = 4
                                    End If
                                End If
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If curʣ���� < rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ս�� > rsWarn!����ֵ Then
                    If Not (InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") > 0 Or bln����) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIF(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־3
            End If
        End If
    End If
End Function

Public Sub GetPatiLastChange(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByRef lng����ID As Long, ByRef lng����id As Long, Optional ByVal int���� As Integer = -1, Optional ByRef strTurnDate As String)
'���ܣ���ȡ���������ת�ƻ�ת������Ϣ
'������int���� -1-����������0-ҽ��վ��1-��ʿվ
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    If int���� = -1 Or int���� = 1 Then
        strSql = " And (��ֹԭ�� = 3 Or ��ֹԭ�� = 15)"
    ElseIf int���� = 0 Or int���� = 2 Then
        strSql = " And (��ֹԭ�� = 3 )"
    End If
    
    strSql = "Select ����id, ����id,��ֹʱ��" & vbNewLine & _
        "From (Select ����id, ����id,��ֹʱ��" & vbNewLine & _
        "       From ���˱䶯��¼" & vbNewLine & _
        "       Where ����id = [1] And ��ҳid = [2]  And ��ֹʱ�� Is Not Null" & strSql & vbNewLine & _
        "       Order By ��ֹʱ�� Desc)" & vbNewLine & _
        "Where Rownum = 1"
        
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "GetPatiLastChange", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        lng����ID = Val("" & rsTmp!����ID)
        lng����id = Val("" & rsTmp!����ID)
        strTurnDate = Format(rsTmp!��ֹʱ��, "yyyy-MM-dd HH:mm:ss")
    Else
        lng����ID = 0
        lng����id = 0
        strTurnDate = ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiCanBilling(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String, Optional ByVal lngModual As Long) As Boolean
'���ܣ����ָ�������Ƿ�������Ȩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(strPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
        Exit Function
    End If
    On Error GoTo errH
    strSql = "Select NVL(B.����,A.����) ����,B.��Ժ����,B.״̬,X.�������" & _
        " From ������Ϣ A,������ҳ B,������� X" & _
        " Where A.����ID=B.����ID And A.����ID=X.����ID(+) And X.����(+) = 2" & _
        " And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlExpense", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!��Ժ����) And Nvl(rsTmp!״̬, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") = 0 Then
            If Nvl(rsTmp!�������, 0) <> 0 Then
                strMsg = """" & rsTmp!���� & """�ķ���δ���壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If InStr(strPrivs, "��Ժ����ǿ�Ƽ���") = 0 Then
            If Nvl(rsTmp!�������, 0) = 0 Then
                strMsg = """" & rsTmp!���� & """�ķ����ѽ��壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If lngModual = pҽ�����ѹ��� Or lngModual = pסԺҽ������ Or lngModual = pסԺҽ���´� Then
            '68081��������Ժ���˴���ҽ������
            strMsg = """" & rsTmp!���� & """�Ѿ���Ժ(��Ԥ��Ժ)�����ܶԸò��˵�ҽ�����з��͡������ջء�ִ�С����ˡ�"
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckEPRReport(ByVal lngҽ��ID As Long, Optional lng����ID As Long, Optional blnBySign As Boolean, Optional ByVal intִ��״̬ As Integer = -999) As Integer
'���ܣ�����Ӧ��Ŀ�ı�����д���
'������lngҽ��ID=�ɼ��е�ҽ��ID
'      lng����ID=���Դ��룬��Ҫ���ڷ��ر��没��ID
'      intִ��״̬=���ڼ������ʱ�������ۺϵ�ִ��״̬
'������blnBySign=�����Ƿ����ͨ��ǩ�������ж�(����ҽ������վ)
'���أ�0-���滹û����д
'      1-��������д���(��ǩ��,�����޶���ǩ��,����ִ�����)
'      2-����δ��д���(δǩ��,���޶���δǩ��,��δִ�����)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, str��鱨��ID As String
    
    On Error GoTo errH
    
    '��鱨���Ƿ�����д
    If lng����ID = 0 Then
        strSql = "Select ����ID,��鱨��ID || '' as ��鱨��ID From ����ҽ������ Where ҽ��ID=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID)
        If Not rsTmp.EOF Then lng����ID = Val(rsTmp!����id & ""): str��鱨��ID = rsTmp!��鱨��ID & ""
    End If
    If lng����ID = 0 And str��鱨��ID = "" Then
        CheckEPRReport = 0: Exit Function
    End If
    
    If Not blnBySign Then
        '��鱨��ִ�й���(5-���;6-�������)��״̬(1-���)
        '���鱨���ǹ������ɼ���ʽ����ģ����ɼ���ʽ����Ϊ����δ�������ͼ�¼
        strSql = _
            " Select 2 as ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
            " Union ALL" & _
            " Select ����,ҽ��ID,ִ�й���,Decode([2],-999,ִ��״̬,[2]) as ִ��״̬,����ʱ��" & _
            " From (" & _
                " Select 1 as ����,B.ҽ��ID,B.ִ�й���,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                    " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                " Order by A.���" & _
            " ) Where Rownum=1" & _
            " Order by ����,����ʱ�� Desc"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID, intִ��״̬)
        If Nvl(rsTmp!ִ�й���, 0) >= 5 Or Nvl(rsTmp!ִ��״̬, 0) = 1 Then
            CheckEPRReport = 1
        Else
            CheckEPRReport = 2
        End If
    Else
        'ͨ��ǩ���汾�жϱ�����ɵķ�ʽ
        strSql = "Select B.�ļ�ID,Max(B.��ʼ��) as ǩ���汾 From ���Ӳ������� B Where B.�ļ�ID=[1] And B.��������=8 Group by B.�ļ�ID"
        strSql = "Select B.���ʱ��,B.���汾,C.ǩ���汾 From ���Ӳ�����¼ B,(" & strSql & ") C Where B.ID=[1] And B.ID=C.�ļ�ID(+)"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "CheckEPRReport", lng����ID)
            
        '(ǩ������ֱ���޸ģ������޶������ǩ�������汾Ӧ��ǩ���汾һ��)
        If IsNull(rsTmp!���ʱ��) Or Nvl(rsTmp!���汾, 0) <> Nvl(rsTmp!ǩ���汾, 0) Then
            '���ҽ�������Ѿ�ִ��,��ʹû��ǩ���򲻷�Ҳ��ͬ���
            strSql = _
                " Select 2 as ����,ҽ��ID,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
                " Union ALL" & _
                " Select ����,ҽ��ID,Decode([2],-999,ִ��״̬,[2]) as ִ��״̬,����ʱ��" & _
                " From (" & _
                    " Select 1 as ����,B.ҽ��ID,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                    " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                        " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                    " Order by A.���" & _
                " ) Where Rownum=1" & _
                " Order by ����,����ʱ�� Desc"
            Set rsTmp = zldatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID, intִ��״̬)
            If Nvl(rsTmp!ִ��״̬, 0) = 1 Then
                CheckEPRReport = 1
            Else
                CheckEPRReport = 2
            End If
        Else
            CheckEPRReport = 1
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetTestLabel(ByVal strScript As String, ByVal strSelect As String, strLabel As String, intResult As Integer)
'���ܣ���ȡƤ�Ա�ע�ͽ��
'������strScript=Ƥ�Խ������������"����(+),������(++);����(-)"
'      strSelect=��ѡ���Ƥ�Խ������������"����"
'���أ�strLabel = Ƥ�Խ����ע����"(+)"
'      intResult=Ƥ�Խ����0-���ԣ�1-����
    Dim arr���� As Variant, arr���� As Variant
    Dim i As Integer
    
    strLabel = "": intResult = 0
    
    arr���� = Split(Split(strScript, ";")(0), ",")
    arr���� = Split(Split(strScript, ";")(1), ",")
    
    For i = 0 To UBound(arr����)
        If arr����(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr����(i), Len(strSelect) + 1)
            intResult = 1: Exit Sub
        End If
    Next
    For i = 0 To UBound(arr����)
        If arr����(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr����(i), Len(strSelect) + 1)
            intResult = 0: Exit Sub
        End If
    Next
End Sub


Public Function ItemHaveCash(ByVal int������Դ As Integer, ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, _
    ByVal lng���ͺ� As Long, ByVal str��� As String, ByVal str���ݺ� As String, ByVal int��¼���� As Integer, ByVal int������� As Integer, ByVal int��ʽ As Integer, _
    Optional ByVal blnMove As Boolean, Optional ByVal dat����ʱ�� As Date, Optional ByRef strҽ��IDs As String, Optional ByRef strNos As String, Optional ByRef blnIsAbnormal As Boolean) As Boolean
'���ܣ��жϵ�ǰ��ִ��ҽ���Ƿ����շѻ���ʻ��۵��Ƿ������
'������int������Դ=1-����,2-סԺ
'      str���=����������ڴ�һ��ҽ�������ַֿ�ִ�е�����
'      int��ʽ=0-����Ƿ����δ�շѼ�¼
'              1-����Ƿ�������շѼ�¼
'      int�������=1=סԺ���͵��������
'      ���أ�strҽ��IDs=��ҽ������ص�ҽ��ID,NOs=ҽ�����͵ĵ��ݺźͲ��ĸ����еĵ��ݺ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTab As String
    
    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
        strTab = "סԺ���ü�¼"
    Else
        strTab = "������ü�¼"
    End If
    ItemHaveCash = True
    strҽ��IDs = ""
    strNos = ""
    
    '��Ӧ�ķ������Ƿ����δ�շ�[��������]������
    '���嵥ֻ��ʾ���շѲ�ͬ��
    '1.�����ҽ������(���Ӽ�¼���ʵ���������Ϊ���ܲ��շѵ�����ʵ�)
    '2.���ʻ���Ҳ��ʾΪδ��(�嵥��Ҫ���Գ���ִ�к����)
    '3.��NO��Ӧ�����ҽ���ķ��ü��(�嵥�ǰ���ʾ��ҽ��ID)
    strSql = _
        " Select A.��¼״̬,Nvl(B.���ID,B.ID) as ҽ��ID,B.�������,A.ִ��״̬,A.NO" & IIF(strTab = "סԺ���ü�¼", ",0 as ����״̬", ",NVL(A.����״̬,0) as ����״̬") & _
        " From " & strTab & " A,����ҽ����¼ B" & _
        " Where A.NO=[4] And A.��¼״̬ IN(0,1,3) And A.ҽ�����+0=B.ID And MOD(A.��¼����,10)=[5]" & IIF(bln����ִ��, " And B.ID=[2]", "") & _
        " Union ALL " & _
        " Select B.��¼״̬,Nvl(C.���ID,C.ID) as ҽ��ID,C.�������,B.ִ��״̬,A.NO" & IIF(strTab = "סԺ���ü�¼", ",0 as ����״̬", ",NVL(b.����״̬,0) as ����״̬") & _
        " From ����ҽ����¼ C," & strTab & " B,����ҽ������ A" & _
        " Where A.NO=B.NO And A.��¼����=MOD(B.��¼����,10) And A.ҽ��ID=B.ҽ�����+0" & IIF(bln����ִ��, " And A.ҽ��ID=[2]", _
            " And A.ҽ��ID IN (Select ID From ����ҽ����¼ Where (ID=[1] Or ���ID=[1]) And �������=[6])") & _
        " And A.���ͺ�=[3] And B.��¼״̬ IN(0,1,3) And A.ҽ��ID=C.ID And A.��¼����=[5]"
    If blnMove Then
        strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
        strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
        strSql = Replace(strSql, strTab, "H" & strTab)
    ElseIf zldatabase.DateMoved(dat����ʱ��) Then
        strSql = strSql & " Union ALL " & Replace(strSql, strTab, "H" & strTab)
    End If
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "ItemHaveCash", IIF(lng���ID <> 0, lng���ID, lngҽ��ID), lngҽ��ID, lng���ͺ�, str���ݺ�, int��¼����, str���)
    If Not rsTmp.EOF Then
        If int��ʽ = 0 Then
            rsTmp.Filter = "ҽ��ID=" & IIF(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ����״̬=1"
            If Not rsTmp.EOF Then
                blnIsAbnormal = True
                ItemHaveCash = False
            Else
                rsTmp.Filter = "ҽ��ID=" & IIF(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬=0"
                If Not rsTmp.EOF Then ItemHaveCash = False
            End If
            
            While Not rsTmp.EOF
                If InStr("," & strҽ��IDs & ",", "," & rsTmp!ҽ��ID & ",") = 0 Then
                    strҽ��IDs = strҽ��IDs & "," & rsTmp!ҽ��ID
                End If
                If InStr("," & strNos & ",", "," & rsTmp!NO & ",") = 0 Then
                    strNos = strNos & "," & rsTmp!NO
                End If
                rsTmp.MoveNext
            Wend
            strNos = Mid(strNos, 2)
            strҽ��IDs = Mid(strҽ��IDs, 2)
        ElseIf int��ʽ = 1 Then
            rsTmp.Filter = "ҽ��ID=" & IIF(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬<>1 And ����״̬<>1"
            If Not rsTmp.EOF Then ItemHaveCash = False
        End If
    ElseIf int��ʽ = 1 Then
        ItemHaveCash = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceMoney(ByVal str��ID As String, ByVal strҽ��ID As String, ByVal str���ͺ� As String, _
    str��� As String, str����� As String, ByVal bln����ִ�� As Boolean, ByVal byt��Դ As Byte) As Currency
'���ܣ�����ָ����ҽ��ID������ȡҽ����Ӧδ��˵ļ��ʷ��úϼ�
'������str��ID,strҽ��ID,str���ͺ�="ID1,ID2,..."
'      bln����ִ��=������Ŀ����ִ�У���ʱֻ��һ��ҽ��ID
'      byt��Դ��1:���2-סԺ
'���أ�str���,str�����=���ڱ�����ʾ
'˵������ϵͳ����Ϊִ�к���˷���ʱ�ŷ��ء�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, curMoney As Currency
    Dim strTab As String
    
    str��� = "": str����� = ""
    
    On Error GoTo errH
    
    If zldatabase.GetPara(81, glngSys) <> "1" Then Exit Function
    strTab = IIF(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
    
    If bln����ִ�� Then
        strSql = _
            " Select B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�շ���Ŀ��� B" & _
            " Where A.ҽ����� + 0 = [2] And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3]" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3])" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
            " Group by B.����,B.����"
    Else
        strSql = _
            " Select /*+ RULE */ B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�շ���Ŀ��� B" & _
            " Where A.ҽ����� + 0 In" & _
            "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
            "       Union All" & _
            "       Select ID From ����ҽ����¼" & _
            "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "  And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In" & _
                "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
                "       Union All" & _
                "       Select ID From ����ҽ����¼" & _
                "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)))" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In (Select Column_Value From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))))" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
            " Group by B.����,B.����"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "GetAdviceMoney", str��ID, strҽ��ID, str���ͺ�, glngSys)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!���, 0)
        str��� = str��� & rsTmp!����
        str����� = str����� & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    str����� = Mid(str�����, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceStuffMoney(ByVal str��ID As String, ByVal strҽ��ID As String, _
    ByVal str���ͺ� As String, ByVal bln����ִ�� As Boolean, ByVal int������Դ As Integer, ByVal int��¼���� As Integer, ByVal int������� As Integer) As Currency
'���ܣ�����ָ����ҽ��ID������ȡҽ����Ӧδ��˵ĸ����������ļ��ʷ��úϼ�
'������str��ID,strҽ��ID,str���ͺ�="ID1,ID2,..."
'      bln����ִ��=������Ŀ����ִ�У���ʱֻ��һ��ҽ��ID
'      int������Դ��1:���2-סԺ
'      int�������=1=סԺ���͵��������
'���أ�str���,str�����=���ڱ�����ʾ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strTab As String
    
    On Error GoTo errH
    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
        strTab = "סԺ���ü�¼"
    Else
        strTab = "������ü�¼"
    End If
    
    If bln����ִ�� Then
        strSql = _
            " Select Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�������� B" & _
            " Where A.ҽ����� + 0 = [2] And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3]" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3])" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����='4' And A.�շ�ϸĿID=B.����ID And B.��������=1"
    Else
        strSql = _
            " Select /*+ RULE */ Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�������� B" & _
            " Where A.ҽ����� + 0 In" & _
            "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
            "       Union All" & _
            "       Select ID From ����ҽ����¼" & _
            "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "  And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In" & _
                "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
                "       Union All" & _
                "       Select ID From ����ҽ����¼" & _
                "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)))" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In (Select Column_Value From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))))" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����='4' And A.�շ�ϸĿID=B.����ID And B.��������=1"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "GetAdviceStuffMoney", str��ID, strҽ��ID, str���ͺ�, glngSys)
    If Not rsTmp.EOF Then GetAdviceStuffMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng����ID As Long, _
    ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal cur��� As Currency, ByVal str��� As String, ByVal str����� As String) As Boolean
'���ܣ���ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
'������str���="CDE..."����������漰�����շ����
'      str�����="���,����,..."����Ӧ�������������ʾ
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSql As String, intR As Integer, i As Long
    Dim cur���� As Currency
    
    On Error GoTo errH
    
    If lng��ҳID <> 0 Then
        'סԺ���˱���
        strSql = _
            " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1] And ���� = 2" & _
            " Union ALL" & _
            " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
        strSql = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSql & ") Group by ����ID"
        
        strSql = "Select NVL(B.����,A.����) ����,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���,C.ʣ���," & _
            " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
            " From ������Ϣ A,������ҳ B,(" & strSql & ") C" & _
            " Where A.����ID=B.����ID And A.����ID=C.����ID(+)" & _
            " And A.����ID=[1] And B.��ҳID=[2]"
        Set rsPati = zldatabase.OpenSQLRecord(strSql, "FinishBillingWarn", lng����ID, lng��ҳID)
    Else
        '���������ﱨ��
        strSql = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ����ID=[1] And ���� = 1"
        strSql = "Select A.����,zl_PatiWarnScheme(A.����ID) as ���ò���,A.������," & _
            " Nvl(B.Ԥ�����,0)-Nvl(B.�������,0) as ʣ���" & _
            " From ������Ϣ A,(" & strSql & ") B" & _
            " Where A.����ID=B.����ID(+)" & _
            " And A.����ID=[1]"
        Set rsPati = zldatabase.OpenSQLRecord(strSql, "FinishBillingWarn", lng����ID)
    End If
    
    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ
    'ִ�б���:���ﲡ�˲���ID=0
    strSql = "Select Nvl(��������,1) as ��������,����ֵ,������־1,������־2,������־3 From ���ʱ����� Where Nvl(����ID,0)=[1] And ���ò���=[2]"
    Set rsWarn = zldatabase.OpenSQLRecord(strSql, "FinishBillingWarn", lng����ID, CStr(Nvl(rsPati!���ò���)))
    If Not rsWarn.EOF Then
        If rsWarn!�������� = 2 Then cur���� = GetPatiDayMoney(lng����ID)
        str����� = Mid(str�����, 2)
        For i = 1 To Len(str���)
            intR = BillingWarn(frmParent, strPrivs, rsWarn, Nvl(rsPati!����), Nvl(rsPati!ʣ���, 0), cur����, cur���, Nvl(rsPati!������, 0), Mid(str���, i, 1), Split(str�����, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemCanCancel(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal lng��ID As Long, str������� As String, _
    ByVal bln����ִ�� As Boolean, ByVal blnMove As Boolean, ByVal byt��Դ As Byte) As Boolean
'���ܣ��ж�ָ����Ŀ�Ƿ����ȡ��ִ��
'������byt��Դ=1:���2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    If gbytBillOpt = 0 Then ItemCanCancel = True: Exit Function
    
    On Error GoTo errH
    
    If bln����ִ�� Then
        strSql = _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]"
    Else
        strSql = _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ���ͺ�=[2]" & _
            " And ҽ��ID IN(Select ID From ����ҽ����¼ Where (ID=[3] Or ���ID=[3]) And �������=[4])"
    End If
    If blnMove Then
        strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
        strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "ItemCanCancel", lngҽ��ID, lng���ͺ�, lng��ID, str�������)
    
    Do While Not rsTmp.EOF
        '�������ſ��˽��ʽ��Ϊ0�ģ�����ķ��õǼ�
        If HaveBilling(rsTmp!NO, True, "", IIF(bln����ִ��, lngҽ��ID, 0), byt��Դ) <> 0 Then
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("����Ŀ�����Ѿ����ʵķ���,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
                    MsgBox "����Ŀ�����Ѿ����ʵķ���,�������ܼ�����", vbExclamation, gstrSysName
                    Exit Function
            End Select
        End If
        rsTmp.MoveNext
    Loop
    ItemCanCancel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��Ӧ����IDs(ByVal lng����ID As Long) As Recordset
'���ܣ���ȡ������Ӧ�Ŀ���
    Dim strSql As String, i As Long

    strSql = " Select B.����ID From �������Ҷ�Ӧ B" & _
            " Where B.����ID=[1]"
    On Error GoTo errH
    Set Get��Ӧ����IDs = zldatabase.OpenSQLRecord(strSql, "Get��Ӧ����IDs", lng����ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
'���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
'�������Ƿ�ȡ���������µĿ���
    Static rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, blnNew As Boolean
    
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    'û��ǿ�������ٴ�,����ҽ��������
    If blnNew Then
        strSql = "Select 1 as ���,����ID From ������Ա Where ��ԱID=[1] Union" & _
                " Select Distinct 2 as ���,B.����ID From ������Ա A,�������Ҷ�Ӧ B" & _
                " Where A.����ID=B.����ID And A.��ԱID=[1]"
        On Error GoTo errH
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISJob", UserInfo.ID)
    End If
    If bln���� = False Then
        rsTmp.Filter = "��� = 1"
    Else
        rsTmp.Filter = ""
    End If
    
    For i = 1 To rsTmp.RecordCount
        If InStr("," & GetUser����IDs & ",", "," & rsTmp!����ID & ",") = 0 Then
            GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser����IDs() As String
'���ܣ���ȡ����Ա�����Ĳ���(ֱ�����ڲ��������ڿ��������Ĳ���),�����ж��
    Static rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, blnNew As Boolean
        
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    If blnNew Then
        strSql = _
            "Select Distinct ����ID From (" & _
            " Select A.����ID as ����ID" & _
            " From ��������˵�� A,������Ա B" & _
            " Where A.����ID=B.����ID And B.��ԱID=[1]" & _
            " And A.������� in(1,2,3) And A.��������='����'" & _
            " Union" & _
            " Select A.����ID From �������Ҷ�Ӧ A,������Ա B" & _
            " Where A.����ID=B.����ID And B.��ԱID=[1])"
        
        On Error GoTo errH
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISWork", UserInfo.ID)
    ElseIf rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
    End If
    For i = 1 To rsTmp.RecordCount
        GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Have��������(ByVal lng����id As Long, ByVal str���� As String, Optional ByRef blnOutDept As Boolean) As Boolean
'���ܣ����ָ�������Ƿ����ָ����������
'˵������Ϊ��������һ�㲻�䶯���ִ���ʹ�ã����û����ȡ
'���أ�blnOutDept=�Ƿ�Ϊ������������Ĳ���
    Static rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, blnNew As Boolean
    
    blnOutDept = False
     
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    On Error GoTo errH
    If blnNew Then
        strSql = "Select ����ID,��������,������� From ��������˵��"
        Set rsTmp = New ADODB.Recordset
        Call zldatabase.OpenRecordset(rsTmp, strSql, "Have��������")
    End If
    rsTmp.Filter = "����ID=" & lng����id & " And ��������='" & str���� & "'"
    Have�������� = Not rsTmp.EOF
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "����ID=" & lng����id & " And ��������='" & str���� & "' And �������<>1"
        blnOutDept = rsTmp.RecordCount = 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveBilling(ByVal strNO As String, ByVal blnALL As Boolean, _
     ByVal strTime As String, ByVal lngҽ��ID As Long, ByVal byt��Դ As Byte) As Integer
'���ܣ��ж�һ�ż��ʵ�/���Ƿ��Ѿ�����
'������strNO=���ʵ��ݺ�,�������ＰסԺ
'      blnALL=�Ƿ�����ŵ������ݽ����ж�,����ֻ��δ���ʲ��ֽ����ж�(����ʱ)
'      byt��Դ=1:���2-סԺ
'���أ�0-δ����,1=��ȫ������,2-�Ѳ��ֽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngTmp As Long
    Dim strTab As String
    
    On Error GoTo errH
    strTab = IIF(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
        
    '��δ���ϵķ�����
    strSql = _
        " Select ��� From (" & _
        " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���," & _
        " Avg(Nvl(����, 1) * ����) As ����" & _
        " From " & strTab & "" & _
        " Where NO=[1] And ��¼����=2" & _
        " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    
    '��ÿ�еĽ������
    strSql = _
        "Select Nvl(�۸񸸺�,���) as ���,Sum(Nvl(���ʽ��,0)) as ���ʽ��" & _
        " From " & strTab & "" & _
        " Where NO=[1] And ��¼���� IN(2,12)" & _
        IIF(Not blnALL, " And Nvl(�۸񸸺�,���) IN(" & strSql & ")", "") & _
        IIF(strTime <> "", " And �Ǽ�ʱ��=[2]", "") & _
        IIF(lngҽ��ID <> 0, " And ҽ�����+0=[3]", "") & _
        " Group by Nvl(�۸񸸺�,���)"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "HaveBilling", strNO, CDate(IIF(strTime = "", "1990-01-01", strTime)), lngҽ��ID)
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '��������
        rsTmp.Filter = "���ʽ��<>0"
        If rsTmp.EOF Then
            HaveBilling = 0 '�޽�����
        ElseIf rsTmp.RecordCount = lngTmp Then
            HaveBilling = 1 'ȫ�����ѽ���
        ElseIf rsTmp.RecordCount > 0 Then
            HaveBilling = 2 '�������ѽ���
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckKSSPrivilege(ByVal int���� As Integer) As Boolean
'���ܣ����ϵͳ�Ƿ���ڿ���ҩ����Ȩ����Ա���������õ�ǰ����Ա����ҩ����UserInfo����
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    UserInfo.��ҩ���� = 0
    
    On Error GoTo errH
    strSql = "Select ���� From ��Ա����ҩ��Ȩ�� Where ��¼״̬=1 and ��ԱID = [1] And ����=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", UserInfo.ID, int����)
    If rsTmp.RecordCount > 0 Then
        UserInfo.��ҩ���� = Val("" & rsTmp!����)
        CheckKSSPrivilege = True
    Else
        strSql = "Select 1 From ��Ա����ҩ��Ȩ�� Where ��¼״̬=1 and Rownum<2 And ����=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", int����)
        CheckKSSPrivilege = rsTmp.RecordCount > 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function CheckLISShowVer(lng���ID As Long) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����˵��:          ͨ�����ID����ϰ�LIS���Ƿ��м�¼������м�¼��ʾʹ���ϰ�򿪣�����ʹ���°�򿪡�
    '����:              lng���ID       ҽ�����ID
    '����:              True �ϰ����ҵ���¼  False û���ҵ���¼
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    CheckLISShowVer = False
    If lng���ID = 0 Then
        Exit Function
    End If
    strSql = "select id from ������Ŀ�ֲ� where ҽ��id = [1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng���ID)
    If rsTmp.RecordCount > 0 Then
        CheckLISShowVer = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SetPublicFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'���ܣ����ô��弰���пؼ��������С
'������frmMe=��Ҫ��������Ĵ������
'      bytSize:����Ϊ9������,0:����Ϊ9������,1,����Ϊ12������
'      strOther:�������������õĿؼ��������ļ���,��ʽΪ����������1,��������2,��������3,....
'˵����1.����漰��VsFlexGrid�ȱ���ؼ�����Ҫ�������ڵĻ������µ����п����и�
'      2.�������δ�г��������ؼ����Զ���ؼ�,��Ҫ���ض�����ָ�������С����ش����ģ������ⵥ������

    Dim objCtrol As Control, objrptCol As ReportColumn
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIF(bytSize = 0, 9, IIF(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "ReportControl", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKind", "PatiIdentify", "VSFlexGrid"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '����CommandBars�û��Զ���ؼ���ȡobjCtrol.Container�����
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.Name
            err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Height = frmMe.TextHeight("��") + 20
                        'Label������Ҫ���е���
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.Count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("����" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                        objCtrol.Height = frmMe.TextHeight("��") + IIF(bytSize = 0, 100, 120)
                Case "TextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("��") + 60
                Case "MaskEdBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("��") + 90
                Case "ReportControl"
                        lngOldSize = objCtrol.PaintManager.TextFont.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        Set CtlFont = objCtrol.PaintManager.TextFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.TextFont = CtlFont
                        For Each objrptCol In objCtrol.Columns
                            objrptCol.Width = objrptCol.Width * dblRate
                        Next
                        objCtrol.Redraw
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        lngOldSize = objCtrol.FontSize
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                Case "IDKind"
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "PatiIdentify"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        objCtrol.IDKindFont.Size = lngFontSize
                        objCtrol.objIDKind.Refrash
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "VSFlexGrid"
                    Call VSFSetFontSize(objCtrol, lngFontSize)
                Case "RichTextBox"
                    Call RTFSetFontSize(objCtrol, lngFontSize)
            End Select
        End If
    Next
End Function

Public Sub VSFSetFontSize(ByRef vsf As Object, ByVal lngFontSize As Long, Optional ByVal lngCol As Long)
'���ܣ�����vsfflexgrid����Ĵ�С�����Զ������п����и�
'������lngFontSize�������ã�9�ż�С�壬12�ż�С��
'      lngCol,���Ҫ�����Զ������иߣ������и����ݵ��У�Ҫ��AutoSizeMode = flexAutoSizeRowHeight,WordWrap =True
    Dim i As Long, lngRate As Double, lngTmp As Long
    
    If lngFontSize < 5 Or lngFontSize > 50 Then Exit Sub
    With vsf
        lngRate = lngFontSize / .FontSize
        lngTmp = .Redraw
        
        .Redraw = flexRDNone
        .FontSize = lngFontSize
        
        For i = 0 To .Cols - 1
           If .ColWidth(i) > 0 Then
             .ColWidth(i) = .ColWidth(i) * lngRate
           End If
        Next
        
        If .AutoSizeMode = flexAutoSizeRowHeight And .WordWrap And lngCol > 0 Then
            .AutoSize lngCol
        Else
            For i = 0 To .Rows - 1
                .RowHeight(i) = .RowHeight(i) * lngRate
            Next
        End If
        .Redraw = lngTmp
    End With
End Sub

Public Sub RTFSetFontSize(ByRef objRTF As Object, ByVal lngFontSize As Long)
'���ܣ���RichTextBox������������
'������objRTF RichTextBox����
'      bytSize 0-С����,1-������
    With objRTF
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = lngFontSize
        .SelLength = 0
    End With
End Sub

Public Sub SetCtrlPosOnLine(ByVal blnvertical As Boolean, ByVal intAligType As Integer, ParamArray arrControls() As Variant)
'����:��ͬһ�еĿؼ�����λ������
'������
'blnvertical  true ,��ֱ�������ÿؼ�λ�ã�false,ˮƽ�������ÿؼ�λ��
'blnvertical=false :intAligType=-1,���˶��룬0-�м���룬1-�׶˶���,blnvertical=true,intAligType=-1,����룬0-ˮƽ���Ķ��룬1-�Ҷ���
'   arrControls��ʽΪ�ؼ�1,���1,�ؼ�2,���2,�ؼ�3,...
    Dim i As Long
    Dim lngPos As Long '��һ���ؼ���ĳһλ��
    Dim dblRate As Double
    If UBound(arrControls) = -1 Then Exit Sub
    If blnvertical Then
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Left
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Left + 0.5 * arrControls(0).Width
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Left + arrControls(0).Width
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Top = arrControls(i - 2).Top + arrControls(i - 2).Height + arrControls(i - 1)
                arrControls(i).Left = lngPos - arrControls(i).Width * dblRate
            End If
        Next
    Else
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Top
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Top + 0.5 * arrControls(0).Height
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Top + arrControls(0).Height
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Left = arrControls(i - 2).Left + arrControls(i - 2).Width + arrControls(i - 1)
                arrControls(i).Top = lngPos - arrControls(i).Height * dblRate
            End If
        Next
    End If
End Sub

Public Function InitAdviceDefine() As Recordset
'���ܣ���ȡҽ�����ݶ����¼��
'������blnNew-�Ƿ񴴽�objVBA��objScript����
'˵����
    Dim strSql As String
    Dim rsDefine As Recordset
    

    On Error GoTo errH
    strSql = "Select �������,ҽ������ From ҽ�����ݶ��� Order by �������"
    Set rsDefine = New ADODB.Recordset
    Call zldatabase.OpenRecordset(rsDefine, strSql, "InitAdviceDefine")
    Set InitAdviceDefine = rsDefine
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckSign(ByVal intǩ������ As Long, ByVal lng��������ID As Long, Optional ByVal lngҽ������ID As Long, Optional ByVal lng���˿���ID As Long, _
    Optional ByVal int���˷�Χ As Integer = 2, Optional ByVal blnCheckCert As Boolean = True, Optional ByRef objESign As Object, Optional ByVal str����ҽ�� As String) As Boolean
'���ܣ��ж�һ�����Ż���һ�鲿�����Ƿ���������˵���ǩ�����Ƶ�
'������int���˷�Χ=1-����,2-סԺ(ȱʡ)
'     intǩ������:0-����ҽ���Ͳ�����1-סԺҽ��ҽ���Ͳ�����2-סԺ��ʿҽ����3-ҽ��ҽ���ͱ��棻4-������¼�ͻ���������5-ҩƷ��ҩ��6-LIS;7-PACS;
'     lng��������ID=���lng��������ID=0������Ҫ���ݴ����ҽ�����ң����˿���ID���Ӧ��Ĭ�Ͽ�������
'                   ��ʿվУ�Ժ�ȷ��ֹͣʱ������Ĳ���ID�����жϲ����Ƿ������˵���ǩ��
'                   ����-1������ҩ�����ʱ������ж��Ƿ�ֿ������ã�
'     blnCheckCert=true ���֤���Ƿ�ͣ�ã�=false��ʾ�����
    Dim strSql As String, intTmp As Integer
    Dim rsTmp As Recordset
    
    '������϶�δ���ã��򷵻�false
    If intǩ������ = 0 Or intǩ������ = 1 Then
        intTmp = intǩ������ + 1
    ElseIf intǩ������ > 1 And intǩ������ <= 7 Then
        intTmp = intǩ������
    End If
    If Mid(gstrESign, intTmp, 1) <> "1" Then Exit Function
    If lng��������ID = 0 And (lng���˿���ID <> 0 Or lngҽ������ID <> 0) Then
        'ȡ��������
        lng��������ID = Get��������ID(UserInfo.ID, lngҽ������ID, lng���˿���ID, int���˷�Χ)
        If lng��������ID = 0 Then Exit Function
    End If
    grsSign.Filter = "����ID=" & lng��������ID & " and ����=" & intǩ������
    If grsSign.RecordCount = 0 Then
        strSql = "Select Zl_Fun_Getsignpar([1],[2]) as �Ƿ����� From dual"
        On Error GoTo errH
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlAdvice", intǩ������, lng��������ID)
        If rsTmp.RecordCount > 0 Then
            CheckSign = Val(rsTmp!�Ƿ����� & "") = 1
            grsSign.AddNew
            grsSign!����ID = lng��������ID
            grsSign!���� = intǩ������
            grsSign!�Ƿ����� = Val(rsTmp!�Ƿ����� & "")
        End If
    Else
        grsSign.MoveFirst
        CheckSign = Val(grsSign!�Ƿ����� & "") = 1
    End If
    If CheckSign = True And blnCheckCert Then
        If objESign Is Nothing Then
            On Error Resume Next
            Set objESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not objESign Is Nothing Then
                Call objESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        '���֤���Ƿ�ͣ��
        If objESign.CertificateStoped(UserInfo.����) Then CheckSign = False
                If str����ҽ�� <> "" Then If objESign.CertificateStoped(str����ҽ��) Then CheckSign = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��������ID(ByVal lngҽ��ID As Long, ByVal lngҽ������ID As Long, ByVal lng���˿���ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal lngִ�п���ID As Long, Optional ByVal lng�������ID As Long) As Long
'���ܣ���ҽ��ȷ����������
'������int��Χ=1-����,2-סԺ(ȱʡ)
'˵������ҽ���������ҷ�Χ��,����˳�����£�
'      1��ҽ������(ҽ������)
'      2���������
'      3�����˿���
'      4������������/סԺ���˵�ĳЩ����ҽ����ִ�п���
'      5������������/סԺ���˵Ŀ�����ΪĬ�Ͽ���
'      6������������/סԺ���˵Ŀ���
'      7��Ĭ�Ͽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Integer
    Dim arr����ID(1 To 7) As Long
    
    '�������ű������ٴ���ҽ��
    strSql = "Select Distinct A.����ID,Nvl(A.ȱʡ,0) as ȱʡ" & _
        " From ������Ա A,��������˵�� B,���ű� C" & _
        " Where A.����ID=C.ID And A.����ID=B.����ID" & _
        " And B.������� IN([2],3) And A.��ԱID=[1]" & _
        " And B.�������� IN('�ٴ�','���','����','����','����','Ӫ��','����')" & _
        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)"
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", lngҽ��ID, int��Χ)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!����ID = lngҽ������ID Then
            arr����ID(1) = rsTmp!����ID
        ElseIf rsTmp!����ID = lng�������ID Then
            arr����ID(2) = rsTmp!����ID
        ElseIf rsTmp!����ID = lng���˿���ID Then
            arr����ID(3) = rsTmp!����ID
        ElseIf rsTmp!����ID = lngִ�п���ID Then
            arr����ID(4) = rsTmp!����ID
        ElseIf rsTmp!ȱʡ = 1 Then
            arr����ID(5) = rsTmp!����ID
        ElseIf arr����ID(5) = 0 Then
            arr����ID(6) = rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    arr����ID(7) = UserInfo.����ID
    
    For i = LBound(arr����ID) To UBound(arr����ID)
        If arr����ID(i) <> 0 Then
            Get��������ID = arr����ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ�������������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ������" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function CreatePlugInOK(ByVal lngMod As Long, Optional ByVal int���� As Integer) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod, int����)
        Call zlPlugInErrH(err, "Initialize")
        err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Function Get��Һ��������() As String
'���ܣ���ȡ��Һ�������ĵĿ���IDs
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Integer
    Dim strReturn As String
    
    On Error GoTo errH

    strSql = "Select ����id From ��������˵�� Where �������� = '��������' Order by ����id"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "Get��Һ��������")
    
    For i = 1 To rsTmp.RecordCount
        strReturn = strReturn & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    Get��Һ�������� = Mid(strReturn, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HavePath(ByVal lng����ID As Long) As Boolean
'���ܣ����ָ�����һ����Ƿ��п��õ��ٴ�·��
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String

    strSql = "Select a.Id" & vbNewLine & _
            "From �ٴ�·��Ŀ¼ A, �ٴ�·���汾 B, �ٴ�·������ C," & vbNewLine & _
            "     (Select ����id From �������Ҷ�Ӧ Where ����id = [1]" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select ID From ���ű� Where ID = [1]) D" & vbNewLine & _
            "Where a.Id = b.·��id And a.���°汾 = b.�汾�� And a.Id = c.·��id(+) And (c.����id = d.����id or c.����id is null) And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlPublic", lng����ID)
    HavePath = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ID(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����id As Long, Optional ByRef bln��ҽ As Boolean = False) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    bln��ҽ = Have��������(lng����id, "��ҽ��")
    
    If bln��ҽ Then
        strSql = "Select ����id, ���id, �������,������� " & vbNewLine & _
                "From ������ϼ�¼" & vbNewLine & _
                "Where ��¼��Դ In (1, 2, 3) And ������� In (1, 2, 11, 12) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� = 1 And" & vbNewLine & _
                "      Nvl(�Ƿ�����, 0) = 0" & vbNewLine & _
                "Order By Decode(�������, 12, 1, 2, 2, 11, 3, 1, 4), Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc"
    Else
        strSql = "Select ����id, ���id, �������,������� " & vbNewLine & _
            "From ������ϼ�¼" & vbNewLine & _
            "Where ��¼��Դ In (1, 2, 3) And ������� In (1, 2, 11, 12) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� = 1 And Nvl(�Ƿ�����,0) = 0" & vbNewLine & _
            "Order By Sign(�������-10),������� Desc, Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc"
    End If
    '��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����
    '�������:1-��ҽ�������;2-��ҽ��Ժ���;11-��ҽ�������;12-��ҽ��Ժ���
    '�ж����ϵ�����£�������ϴ���ֻȡ��һ����Ҫ���
    '���������������ȣ���Ҫ��Ϊ��֧��������ϡ�
    
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ����", lng����ID, lng��ҳID)
    Set Get����ID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPathTable(lng����ID As Long, lng���ID As Long, lng����id As Long) As ADODB.Recordset
    Dim strSql As String
 
    strSql = "Select a.Id, a.����, a.����, a.����, a.˵��, Nvl(a.���ò���,'ͨ��') ���ò���, a.�����Ա�, a.��������, a.���°汾, c.��׼סԺ��,Nvl(a.��������,'��') as ��������,Nvl(a.ȷ������,0) as ȷ������" & vbNewLine & _
            "From �ٴ�·��Ŀ¼ A, �ٴ�·������ B,�ٴ�·���汾 C" & vbNewLine & _
            "Where a.Id = b.·��id And (b.����id = [1] Or b.���id = [2]) And a.���°汾 is not null And a.id = b.·��ID And a.���°汾 = c.�汾��" & vbNewLine & _
            "And a.Id = c.·��id And b.����=0 And (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ D Where a.Id = D.·��id And d.����id = [3]))"
    On Error GoTo errH
    Set GetPathTable = zldatabase.OpenSQLRecord(strSql, "��ȡ·��Ŀ¼", lng����ID, lng���ID, lng����id)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.�û���)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS��������δ�����ɹ���", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function

Public Function PatiFeeUsable(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����˵ĵ�ǰ�����Ƿ���Ч������true������ǰ�ѱ����
    Dim rsTmp As ADODB.Recordset, strSql As String
    On Error GoTo errH
    strSql = "Select 1 From ������ҳ A, �ѱ� B Where a.�ѱ�=b.���� And a.����id=[1] And a.��ҳid=[2] And Sysdate Between Nvl(b.��Ч��ʼ,Sysdate) And Nvl(b.��Ч����,Sysdate)"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "PatiFeeUsable", lng����ID, lng��ҳID)
    PatiFeeUsable = True
    If rsTmp.EOF Then
        MsgBox "�ò��˵ĵ�ǰ�ѱ��Ѿ�ʧЧ�����ܷ���ҽ�������ڲ�����Ϣ�е������˷ѱ�", vbInformation, gstrSysName
        PatiFeeUsable = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcÿСʱ����ӵ��췢�Ϳ�ʼʱ��(ByVal dat��ʼִ��ʱ�� As Date, ByVal datCurr As Date, ByVal intƵ�ʼ�� As Integer) As Date
'���ܣ���Ҫ����ÿNСʱһ�ε���������ҹ�ѡ�˳����ӵ��쿪ʼ���͵ļ������⣻89561
    Dim datBegin As Date

    datBegin = dat��ʼִ��ʱ��
    Do While CDate(Format(datBegin, "yyyy-mm-dd")) < CDate(Format(datCurr, "yyyy-mm-dd"))
        datBegin = DateAdd("h", intƵ�ʼ��, datBegin)
    Loop
    CalcÿСʱ����ӵ��췢�Ϳ�ʼʱ�� = datBegin
End Function

Public Function HaveRIS(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��ж� RIS�ӿڲ��� �Ƿ���ڣ�������
'������blnMsg������ʧ��ʱ�Ƿ���ʾ
    If Not gbln����Ӱ����Ϣϵͳ�ӿ� Then Exit Function
    If gobjRis Is Nothing Then
        On Error Resume Next
        Set gobjRis = CreateObject("zl9XWInterface.clsHISInner")
        err.Clear: On Error GoTo 0
        If Not gobjRis Is Nothing Then
            gbln����Ӱ����ϢϵͳԤԼ = gobjRis.HISSchedulingjudge = 0
        End If
    End If
    If gobjRis Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    HaveRIS = True
End Function