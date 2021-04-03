----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.34.150������ v10.34.150
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------


------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--131243:��ΰ��,2018-09-17,���֤У������̨�������
Create Or Replace Function Zl_Fun_Checkidcard
(
  Idcard_In   In Varchar2,
  Calcdate_In In Date := Null
) Return Varchar2 Is
  -------------------------------------------------------------------------------
  --���ܣ����֤����Ϸ���У��,���������֤�ŵĳ������ڡ��Ա�����
  --����˵��:
  -- ��� IDcard_In:���֤����
  --    Calcdate_In:��������,ȱʡʱ��ϵͳʱ��
  -- ����ֵ���̶���ʽXML��
  --<OUTPUT>
  --       <BIRTHDAY></BIRTHDAY>                //��������
  --       <SEX></SEX>                  //�Ա�
  --       <AGE></AGE>                //����
  --     <MSG></MSG>         //�մ�-���֤����Ч(�ɴ����֤���л�ȡ�������ں��Ա�)���ǿմ�-���ش�����Ϣ
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Count     Number(5);
  n_Sum       Number(5);
  v_У��λ    Varchar2(50);
  v_Pattern   Varchar2(500);
  v_Err_Msg   Varchar2(2000);
  v_�Ա�      Varchar2(100);
  v_����      Varchar2(100);
  d_Curr_Time Date;
  d_��������  Date;
  v_Temp      Varchar2(20);

Begin
  Select Sysdate Into d_Curr_Time From Dual;

  If Idcard_In Is Null Then
    v_Err_Msg := '�������֤��Ϊ��!';
    Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
  Else
    --���֤�Ϸ���֤
    v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    --��������
    If Instr(v_Pattern, Substr(Idcard_In, 1, 2)) = 0 Then
      v_Err_Msg := '���֤ǰ��λ�����벻��ȷ!';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    --���֤���ȼ��
    If Length(Idcard_In) = 15 Then
      --������֤��:15λ���֤��Ҫ��ȫ��Ϊ����
      v_Pattern := '^\d{15}$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '���֤�а����Ƿ��ַ�������!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --��ȡ�Ա�
      If Mod(To_Number(Substr(Idcard_In, 15, 1)), 2) = 1 Then
        v_�Ա� := '��';
      Else
        v_�Ա� := 'Ů';
      End If;
      --�������ڵĺϷ��Լ��
      v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(Idcard_In, 7, 6), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '���֤�еĳ���������Ч������!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        --��ǰ�������֤û�����������������ݴ����������ڸ�Ϊ2��28�ţ��磺19470229�������
        If Instr(',0229,0230,', ',' || Substr(Idcard_In, 9, 4) || ',') > 0 Then
          v_Temp     := '19' || Substr(Idcard_In, 7, 2) || '0301';
          d_�������� := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
        Else
          d_�������� := To_Date('19' || Substr(Idcard_In, 7, 6), 'yyyy-mm-dd');
        End If;
        If d_�������� > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '���֤�еĳ���������Ч������!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Elsif Length(Idcard_In) = 18 Then
      -- 18 λ���֤��ǰ17 λȫ��Ϊ���֣����1λ��Ϊ���ֻ�x
      v_Pattern := '^\d{17}[0-9Xx]$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '���֤�а����Ƿ��ַ�!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --��ȡ�Ա�
      If Mod(To_Number(Substr(Idcard_In, 17, 1)), 2) = 1 Then
        v_�Ա� := '��';
      Else
        v_�Ա� := 'Ů';
      End If;
      --�������ڵĺϷ��Լ��
      v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(Idcard_In, 7, 8), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '���֤�еĳ���������Ч������!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        --��ǰ�������֤û�����������������ݴ����������ڸ�Ϊ2��28�ţ��磺19470229�������
        If Instr(',0229,0230,', ',' || Substr(Idcard_In, 11, 4) || ',') > 0 Then
          v_Temp     := Substr(Idcard_In, 7, 4) || '0301';
          d_�������� := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
        Else
          d_�������� := To_Date(Substr(Idcard_In, 7, 8), 'yyyy-mm-dd');
        End If;
        If d_�������� > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '���֤�еĳ���������Ч������!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
        --����У��λ
        n_Sum     := (To_Number(Substr(Idcard_In, 1, 1)) + To_Number(Substr(Idcard_In, 11, 1))) * 7 +
                     (To_Number(Substr(Idcard_In, 2, 1)) + To_Number(Substr(Idcard_In, 12, 1))) * 9 +
                     (To_Number(Substr(Idcard_In, 3, 1)) + To_Number(Substr(Idcard_In, 13, 1))) * 10 +
                     (To_Number(Substr(Idcard_In, 4, 1)) + To_Number(Substr(Idcard_In, 14, 1))) * 5 +
                     (To_Number(Substr(Idcard_In, 5, 1)) + To_Number(Substr(Idcard_In, 15, 1))) * 8 +
                     (To_Number(Substr(Idcard_In, 6, 1)) + To_Number(Substr(Idcard_In, 16, 1))) * 4 +
                     (To_Number(Substr(Idcard_In, 7, 1)) + To_Number(Substr(Idcard_In, 17, 1))) * 2 +
                     To_Number(Substr(Idcard_In, 8, 1)) * 1 + To_Number(Substr(Idcard_In, 9, 1)) * 6 +
                     To_Number(Substr(Idcard_In, 10, 1)) * 3;
        n_Count   := Mod(n_Sum, 11);
        v_Pattern := '10X98765432';
        v_У��λ  := Substr(v_Pattern, n_Count + 1, 1);
        If v_У��λ <> Upper(Substr(Idcard_In, 18, 1)) Then
          v_Err_Msg := '���֤���벻��ȷ�����顣';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Else
      v_Err_Msg := '���֤���Ȳ���,���顣';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    v_���� := Zl_Age_Calc(0, d_��������, Calcdate_In);
  End If;

  Return '<OUTPUT><BIRTHDAY>' || To_Char(d_��������, 'YYYY-MM-DD') || '</BIRTHDAY><SEX>' || v_�Ա� || '</SEX><AGE>' || v_���� || '</AGE><MSG></MSG></OUTPUT>';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Checkidcard;
/
--131243:��ΰ��,2018-09-17,���֤У������̨�������

Create Or Replace Procedure Zl_Third_Buildpatient
(
  Patiinfo_In  In Xmltype,
  Patiinfo_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------
  --����˵��:
  -- ��� Patiinfo_In:
  --<IN>
  --  <ZJH></ZJH>                 //֤���ţ�Ŀǰ��֧�����֤��
  --  <ZJLX></ZJLX>                       //֤������(Ŀǰ��֧�����֤,Ϊ��ʱĬ��Ϊ���֤)
  --  <XM></XM>                       //����
  --  <SJH></SJH>                      //�ֻ���
  --</IN>

  --���� Patiinfo_Out��
  --<OUTPUT>
  --       <BRID></BRID>                //����ID
  --       <MZH></MZH>                  //�����
  --     <ERROR></ERROR>         //����д��󷵻ظýڵ�
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Pati_Id      ������Ϣ.����id%Type;
  n_Card_Type_Id ҽ�ƿ����.Id%Type;
  n_Count        Number(5);
  n_Sum          Number(5);
  v_У��λ       Varchar2(50);

  v_����         ������Ϣ.����%Type;
  v_���֤��     ������Ϣ.���֤��%Type;
  v_�ֻ���       ������Ϣ.��ͥ�绰%Type;
  v_�Ա�         ������Ϣ.�Ա�%Type;
  v_����         ������Ϣ.����%Type;
  v_����Ա       ��Ա��.����%Type;
  v_ҽ�Ƹ��ʽ ������Ϣ.ҽ�Ƹ��ʽ%Type;
  n_�����       ������Ϣ.�����%Type;
  v_֤������     ҽ�ƿ����.����%Type;
  v_֤����       ����ҽ�ƿ���Ϣ.����%Type;

  v_Pattern Varchar2(500);
  v_Temp    Varchar2(32767); --��ʱXML
  v_Err_Msg Varchar2(2000);
  n_����    Number(2);

  d_��������  ������Ϣ.��������%Type;
  d_Curr_Time Date;

  Err_Item Exception;
Begin
  Patiinfo_Out := Xmltype('<OUTPUT></OUTPUT>');
  Select Sysdate Into d_Curr_Time From Dual;

  --�½����ˣ����������֤�š��ֻ��ţ����ڼ�ͥ�绰�У����������ڡ��Ա�����(��������ɴ����֤�л�ȡ)��
  Select Extractvalue(Value(I), 'IN/XM'), Extractvalue(Value(I), 'IN/ZJH'), Extractvalue(Value(I), 'IN/SJH'),
         Extractvalue(Value(I), 'IN/ZJLX')
  Into v_����, v_֤����, v_�ֻ���, v_֤������
  From Table(Xmlsequence(Extract(Patiinfo_In, 'IN'))) I;

  Begin
    If v_֤������ Is Null Then
      Select ����id
      Into n_Pati_Id
      From ����ҽ�ƿ���Ϣ
      Where ���� = v_֤���� And �����id In (Select ID From ҽ�ƿ���� Where ���� Like '%���֤%') And Rownum < 2;
    Else
      Select ����id
      Into n_Pati_Id
      From ����ҽ�ƿ���Ϣ
      Where ���� = v_֤���� And �����id In (Select ID From ҽ�ƿ���� Where ���� = v_֤������) And Rownum < 2;
    End If;
    n_���� := 1;
  Exception
    When Others Then
      n_���� := 0;
  End;

  If Nvl(n_����, 0) = 1 Then
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    Select ����� Into n_����� From ������Ϣ Where ����id = n_Pati_Id;
    If n_����� Is Null Then
      n_����� := Nextno(3);
      Update ������Ϣ Set ����� = n_����� Where ����id = n_Pati_Id;
    End If;
    v_Temp := '<MZH>' || n_����� || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  Else
    If v_���� Is Null Then
      v_Err_Msg := '��������Ϊ��!';
      Raise Err_Item;
    End If;
    If v_֤������ Like '%���֤%' Or v_֤������ Is Null Then
      v_���֤�� := v_֤����;
    Else
      v_Err_Msg := 'Ŀǰ��֧�����֤����ķ�ʽ������';
      Raise Err_Item;
    End If;
  
    If v_���֤�� Is Null Then
      v_Err_Msg := '�������֤��Ϊ��!';
      Raise Err_Item;
    Else
      --���֤�Ϸ���֤
      v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    
      --��������
      If Instr(v_Pattern, Substr(v_���֤��, 1, 2)) = 0 Then
        v_Err_Msg := '���֤ǰ��λ�����벻��ȷ!';
        Raise Err_Item;
      End If;
      --���֤���ȼ��
      If Length(v_���֤��) = 15 Then
        --������֤��:15λ���֤��Ҫ��ȫ��Ϊ����
        v_Pattern := '^\d{15}$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_���֤��, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '���֤�а����Ƿ��ַ�������!';
          Raise Err_Item;
        End If;
        --��ȡ�Ա�
        If Mod(To_Number(Substr(v_���֤��, 15, 1)), 2) = 1 Then
          v_�Ա� := '��';
        Else
          v_�Ա� := 'Ů';
        End If;
        --�������ڵĺϷ��Լ��
      
        v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(v_���֤��, 7, 6), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '���֤�еĳ���������Ч������!';
          Raise Err_Item;
        Else
          --��ǰ�������֤û�����������������ݴ����������ڸ�Ϊ2��28�ţ��磺19470229�������
          If Instr(',0229,0230,', ',' || Substr(v_���֤��, 9, 4) || ',') > 0 Then
            v_Temp     := '19' || Substr(v_���֤��, 7, 2) || '0301';
            d_�������� := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_�������� := To_Date('19' || Substr(v_���֤��, 7, 6), 'yyyy-mm-dd');
          End If;
          If d_�������� > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '���֤�еĳ���������Ч������!';
            Raise Err_Item;
          End If;
        End If;
      Elsif Length(v_���֤��) = 18 Then
        -- 18 λ���֤��ǰ17 λȫ��Ϊ���֣����1λ��Ϊ���ֻ�x
        v_Pattern := '^\d{17}[0-9Xx]$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_���֤��, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '���֤�а����Ƿ��ַ�!';
          Raise Err_Item;
        End If;
        --��ȡ�Ա�
        If Mod(To_Number(Substr(v_���֤��, 17, 1)), 2) = 1 Then
          v_�Ա� := '��';
        Else
          v_�Ա� := 'Ů';
        End If;
        --�������ڵĺϷ��Լ��
        v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(v_���֤��, 7, 8), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '���֤�еĳ���������Ч������!';
          Raise Err_Item;
        Else
          --��ǰ�������֤û�����������������ݴ����������ڸ�Ϊ2��28�ţ��磺19470229�������
          If Instr(',0229,0230,', ',' || Substr(v_���֤��, 11, 4) || ',') > 0 Then
            v_Temp     := Substr(v_���֤��, 7, 4) || '0301';
            d_�������� := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_�������� := To_Date(Substr(v_���֤��, 7, 8), 'yyyy-mm-dd');
          End If;
          If d_�������� > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '���֤�еĳ���������Ч������!';
            Raise Err_Item;
          End If;
          --����У��λ
          n_Sum     := (To_Number(Substr(v_���֤��, 1, 1)) + To_Number(Substr(v_���֤��, 11, 1))) * 7 +
                       (To_Number(Substr(v_���֤��, 2, 1)) + To_Number(Substr(v_���֤��, 12, 1))) * 9 +
                       (To_Number(Substr(v_���֤��, 3, 1)) + To_Number(Substr(v_���֤��, 13, 1))) * 10 +
                       (To_Number(Substr(v_���֤��, 4, 1)) + To_Number(Substr(v_���֤��, 14, 1))) * 5 +
                       (To_Number(Substr(v_���֤��, 5, 1)) + To_Number(Substr(v_���֤��, 15, 1))) * 8 +
                       (To_Number(Substr(v_���֤��, 6, 1)) + To_Number(Substr(v_���֤��, 16, 1))) * 4 +
                       (To_Number(Substr(v_���֤��, 7, 1)) + To_Number(Substr(v_���֤��, 17, 1))) * 2 +
                       To_Number(Substr(v_���֤��, 8, 1)) * 1 + To_Number(Substr(v_���֤��, 9, 1)) * 6 +
                       To_Number(Substr(v_���֤��, 10, 1)) * 3;
          n_Count   := Mod(n_Sum, 11);
          v_Pattern := '10X98765432';
          v_У��λ  := Substr(v_Pattern, n_Count + 1, 1);
          If v_У��λ <> Upper(Substr(v_���֤��, 18, 1)) Then
            v_Err_Msg := '���֤���벻��ȷ�����顣';
            Raise Err_Item;
          End If;
        End If;
      Else
        v_Err_Msg := '���֤���Ȳ���,���顣';
        Raise Err_Item;
      End If;
    
      If Nvl(v_����, '_') = '_' Then
        v_���� := Zl_Age_Calc(0, d_��������, d_Curr_Time);
      End If;
    End If;
  
    Select ���� Into v_ҽ�Ƹ��ʽ From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1;
    n_Pati_Id := Nextno(1);
    n_�����  := Nextno(3);
    Insert Into ������Ϣ
      (����id, ����, ���֤��, ��ͥ�绰, ��������, �Ա�, ����, �Ǽ�ʱ��, �����, ҽ�Ƹ��ʽ, �ֻ���)
      Select n_Pati_Id, v_����, v_���֤��, v_�ֻ���, d_��������, v_�Ա�, v_����, d_Curr_Time, n_�����, v_ҽ�Ƹ��ʽ, v_�ֻ���



      From Dual;
    --������Ϣ����������ҽ�ƿ��󶨣��������֤�����İ󶨣�
    Begin
      If v_֤������ Is Null Then
        Select ID Into n_Card_Type_Id From ҽ�ƿ���� Where ���� Like '%���֤%' And Rownum < 2;
      Else
        Select ID Into n_Card_Type_Id From ҽ�ƿ���� Where ���� = v_֤������ And Rownum < 2;
      End If;
    Exception
      When No_Data_Found Then
        v_Err_Msg := '���֤����𲻴��ڣ�';
        Raise Err_Item;
    End;
    Select b.���� Into v_����Ա From �ϻ���Ա�� A, ��Ա�� B Where a.��Աid = b.Id And a.�û��� = User;
  
    Zl_ҽ�ƿ��䶯_Insert(11, n_Pati_Id, n_Card_Type_Id, Null, v_���֤��, '�������⿨', Null, v_����Ա, d_Curr_Time);
  
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    v_Temp := '<MZH>' || n_����� || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  End If;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Buildpatient;
/


------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.150.0014' Where ���=&n_System;
Commit;