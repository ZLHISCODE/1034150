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
--126591:���˺�,2018-06-07,�������ѵĴ���.

Create Or Replace Procedure Zl_Third_Settlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --����:�����ӿڽ���
  --���:Xml_In:
  --<IN>
  --        <BRID>����ID</BRID>         //����ID
  --        <XM>����</XM>               //����
  --        <SFZH>���֤��</SFZH>       //���֤��
  --        <ZYID>��ҳID</ZYID>         //��ҳID
  --        <JSLX>2</JSLX>         //��������,1-����,2-סԺ.Ŀǰ�̶���2
  --        <JE></JE>         //���ν����ܽ��
  --       <JSLIST>
  --         <JS>
  --              <JSKLB>֧�������</JSKLB >
  --              <JSKH>֧������</ JSKH >
  --              <JSFS>֧����ʽ</JSFS> //֧����ʽ:�ֽ�;֧Ʊ,�����������,���Դ���
  --              <JSJE>������</JSJE> //������(�������˲������ҽԺ�˿�)<SFCYJ>Ϊ1ʱΪ��Ԥ�����
  --              <JYLSH>������ˮ��</JYLSH>
  --              <ZY>ժҪ</ZY>
  --              <SFCYJ>�Ƿ��Ԥ��</SFCYJ>  //�Ƿ��Ԥ����0-���㣬1-��Ԥ��.�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�
  --              <SFXFK>�Ƿ����ѿ�</SFXFK>  //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ�
  --              <EXPENDLIST>  //��չ������Ϣ
  --                  <EXPEND>
  --                        <JYMC>��������</JYMC> //��������   �˿�ʱ,�����Ԥ������ˮ��
  --                        <JYLR>��������</JYLR> //��������   �˿�ʱ,�����Ԥ���Ľ��
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --����:Xml_Out
  --  <OUT>
  --       <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --    �D�D�������д�������˵����ȷִ��
  --    <ERROR>
  --      <MSG>������Ϣ</MSG>
  --    </ERROR>
  --  </OUT>
  --------------------------------------------------------------------------------------------------
  n_��ҳid     ������ҳ.��ҳid%Type;
  n_����id     ������ҳ.����id%Type;
  v_����       ������Ϣ.����%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  n_�����ܶ�   ����Ԥ����¼.��Ԥ��%Type;
  n_�����ʽ�� ����Ԥ����¼.��Ԥ��%Type;
  n_��������   Number(3);
  v_����Ա���� ���˽��ʼ�¼.����Ա���%Type;
  v_����Ա���� ���˽��ʼ�¼.����Ա����%Type;
  n_����id     ���˽��ʼ�¼.Id%Type;
  n_��Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  d_����ʱ��   Date;
  n_Ԥ����ֵ   ����Ԥ����¼.���%Type;
  d_��ʼ����   Date;
  d_��������   Date;
  d_��С����   Date;
  d_�������   Date;

  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_����id     ������ҳ.��Ժ����id%Type;
  n_���㿨��� �����ѽӿ�Ŀ¼.���%Type;
  n_ʱ������   Number(3);
  v_Ids        Varchar2(20000);
  v_No         ���˽��ʼ�¼.No%Type;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_���㿨��   ����Ԥ����¼.����%Type;
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  v_Temp       Varchar2(500);
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;

  n_Count    Number(18);
  n_Number   Number(2);
  n_����id   ������ü�¼.Id%Type;
  n_��¼���� ������ü�¼.��¼����%Type;
  v_����no   ������ü�¼.No%Type;
  n_���     ������ü�¼.���%Type;
  n_��¼״̬ ������ü�¼.��¼״̬%Type;
  n_ִ��״̬ ������ü�¼.ִ��״̬%Type;
  n_δ���� ������ü�¼.ʵ�ս��%Type;
  n_���ʽ�� ������ü�¼.ʵ�ս��%Type;
  n_����   ������ü�¼.ʵ�ս��%Type;

  v_�����     �������׼�¼.���%Type;
  v_���ѿ����� Varchar2(20000);

  Type t_���ý�����ϸ Is Ref Cursor;
  c_���ý�����ϸ t_���ý�����ϸ;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/JSLX')),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_��ҳid, n_����id, n_�����ܶ�, n_��������, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_�������� := Nvl(n_��������, 2);
  If n_�������� = 1 And Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  --0.��ؼ��
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ�������,������ɷ�!';
    Raise Err_Item;
  End If;

  --��Աid,��Ա���,��Ա����
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := 'ϵͳ�����ϱ���Ч�Ĳ���Ա,���������!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := v_Temp;
  v_Err_Msg    := Null;

  For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Not (c_���׼�¼.���㿨��� Is Null Or Nvl(c_���׼�¼.�Ƿ����ѿ�, '0') = '1' Or Nvl(c_���׼�¼.�Ƿ��Ԥ��, 0) = 1) Then
    
      Select Decode(Translate(Nvl(c_���׼�¼.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
      Into n_Count
      From Dual;
    
      If Nvl(n_Count, 0) = 1 Then
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���׼�¼.���㿨���);
      Else
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���׼�¼.���㿨���;
      End If;
    
      If v_����� Is Null Then
        v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, 2) = 0 Then
        v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
        Raise Err_Special;
      End If;
    End If;
  End Loop;

  Select Max(��Ժ����id) Into n_����id From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
  Select ���˽��ʼ�¼_Id.Nextval, Sysdate, Nextno(15) Into n_����id, d_����ʱ��, v_No From Dual;

  n_ʱ������ := Zl_Getsysparameter('���ʷ���ʱ��', 1137);

  Select ���˽��ʼ�¼_Id.Nextval, Sysdate, Nextno(15) Into n_����id, d_����ʱ��, v_No From Dual;

  If n_�������� = 2 Then
    Open c_���ý�����ϸ For
      Select Max(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
             Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
             Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
      From סԺ���ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1
      Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
      Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, ���;
  Else
  
    Open c_���ý�����ϸ For
      Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
             Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
             Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
      From ������ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1
      Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
      Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
      Union All
      Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
             Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
             Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
      From סԺ���ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1 And Mod(��¼����, 10) = 5
      Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
      Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, ���;
  End If;

  n_�����ʽ�� := 0;
  Loop
    Fetch c_���ý�����ϸ
      Into n_����id, n_��¼����, v_����no, n_���, n_��¼״̬, n_ִ��״̬, d_��С����, d_�������, n_δ����, n_���ʽ��;
    Exit When c_���ý�����ϸ%NotFound;
  
    n_�����ʽ�� := n_�����ʽ�� + Nvl(n_δ����, 0);
    If d_��ʼ���� Is Null Then
      d_��ʼ���� := d_��С����;
    Elsif d_��ʼ���� > d_��С���� Then
      d_��ʼ���� := d_��С����;
    End If;
    If d_�������� Is Null Then
      d_�������� := d_�������;
    Elsif d_�������� < d_������� Then
      d_�������� := d_�������;
    End If;
  
    If Nvl(n_���ʽ��, 0) = 0 Then
      If n_����id Is Not Null Then
        If Length(v_Ids || ',' || n_����id) > 4000 Then
          v_Ids := Substr(v_Ids, 2);
          Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
          v_Ids := '';
        End If;
        v_Ids := v_Ids || ',' || n_����id;
      Else
        Zl_���ʷ��ü�¼_Insert(0, v_����no, n_��¼����, n_��¼״̬, n_ִ��״̬, n_���, n_δ����, n_����id);
      End If;
    Else
      Zl_���ʷ��ü�¼_Insert(0, v_����no, n_��¼����, n_��¼״̬, n_ִ��״̬, n_���, n_δ����, n_����id);
    End If;
  
  End Loop;

  If v_Ids Is Not Null Then
    v_Ids := Substr(v_Ids, 2);
    Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
  End If;
  n_�����ʽ�� := Round(n_�����ʽ��, 6);

  If n_�����ʽ�� <> Nvl(n_�����ܶ�, 0) Then
    v_Err_Msg := '����Ľ��ʽ����ʵ�ʽ��ʽ���,���������!';
    Raise Err_Item;
  End If;

  Zl_���˽��ʼ�¼_Insert(n_����id, v_No, n_����id, d_����ʱ��, d_��ʼ����, d_��������, 0, 0, n_��ҳid, Null, 2, Null, n_��������);

  n_���ʽ�� := 0;
  n_Count    := 0;
  For r_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_�����   := r_���㷽ʽ.���㷽ʽ;
    n_���ʽ�� := n_���ʽ�� + Nvl(r_���㷽ʽ.������, 0);
  
    If Nvl(r_���㷽ʽ.�Ƿ��Ԥ��, 0) = 0 Then
      --����
      If n_Count = 1 Then
        v_Err_Msg := '���ʽ����ݲ�֧�ֶ��ֽ��㷽ʽ!';
        Raise Err_Item;
      End If;
      n_�����id := Null;
      If r_���㷽ʽ.���㿨��� Is Not Null Then
        Select Decode(Translate(Nvl(r_���㷽ʽ.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
      
        If Nvl(r_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
          If Nvl(n_Number, 0) = 1 Then
            Select Max(���), Max(���㷽ʽ), Max(����)
            Into n_���㿨���, v_���㷽ʽ, v_�����
            From �����ѽӿ�Ŀ¼
            Where ��� = n_�����id And Nvl(����, 0) = 1;
          Else
            Select Max(���), Max(���㷽ʽ), Max(����)
            Into n_���㿨���, v_���㷽ʽ, v_�����
            From �����ѽӿ�Ŀ¼
            Where ���� = r_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
          
          End If;
        
          If n_���㿨��� Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ�����ѿ���Ϣ';
            Raise Err_Item;
          
          End If;
          n_�����id := Null;
        
        Else
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID), Max(���㷽ʽ), Max(����)
            Into n_�����id, v_���㷽ʽ, v_�����
            From ҽ�ƿ����
            Where ID = n_�����id And Nvl(�Ƿ�����, 0) = 1;
          Else
            Select Max(ID), Max(���㷽ʽ), Max(����)
            Into n_�����id, v_���㷽ʽ, v_�����
            From ҽ�ƿ����
            Where ���� = r_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
          End If;
        
          If n_�����id Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ��ҽ�ƿ���Ϣ!';
            Raise Err_Item;
          End If;
        End If;
      End If;
    
      If n_�����id Is Not Null Then
        --������,����סԺԤ����
        v_���㿨�� := r_���㷽ʽ.���㿨��;
        If r_���㷽ʽ.������ > 0 Then
          --��ֵ���ֲ�Ӧ���������ν��� 
          Select ����Ԥ����¼_Id.Nextval, Nextno(11) Into n_Ԥ��id, v_Ԥ��no From Dual;
          Zl_����Ԥ����¼_Insert(n_Ԥ��id, v_Ԥ��no, Null, n_����id, n_��ҳid, n_����id, r_���㷽ʽ.������, v_���㷽ʽ, '', '', '', '', '',
                           v_����Ա����, v_����Ա����, Null, n_��������, n_�����id, Null, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, Null,
                           d_����ʱ��, 0);
          For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                         From Table(Xmlsequence(Extract(r_���㷽ʽ.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
            Zl_�������㽻��_Insert(n_�����id, 0, r_���㷽ʽ.���㿨��, n_Ԥ��id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 1);
          End Loop;
	  --���ʴ���
	  Zl_����Ԥ����¼_Insert(n_Ԥ��id,v_Ԥ��no, 1, r_���㷽ʽ.������,  n_����id, n_����id);
        Else
        
          Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, r_���㷽ʽ.������, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��,
                           Null, Null, Null, Null, Null, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��);
        
          For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                         From Table(Xmlsequence(Extract(r_���㷽ʽ.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
            Zl_�������㽻��_Insert(n_�����id, 0, r_���㷽ʽ.���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
          End Loop;
        
        End If;
      
      Else
        If n_���㿨��� Is Not Null Then
          --���ѿ�
          v_���ѿ����� := Nvl(v_���ѿ�����, '') || '||' || n_���㿨��� || '|' || r_���㷽ʽ.���㿨�� || '|0|' || r_���㷽ʽ.������;
        Else
          --��������
          Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, r_���㷽ʽ.������, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��,
                           Null, Null, Null, Null, Null, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��);
        End If;
      End If;
      n_Count := 1;
    Else
      --��Ԥ��,ĿǰĬ��ȫ��
      n_��Ԥ����� := r_���㷽ʽ.������;
      For r_Ԥ�� In (Select Min(ID) As ID, NO, ���㷽ʽ, Sum(Nvl(���, 0)) - Sum(Nvl(��Ԥ��, 0)) As ���, ������ˮ��
                   From ����Ԥ����¼
                   Where ����id = n_����id And Mod(��¼����, 10) = 1 And Nvl(Ԥ�����, 2) = 2 And (��ҳid = n_��ҳid Or ��ҳid Is Null)
                   Group By NO, ���㷽ʽ, ������ˮ��
                   Having Sum(Nvl(���, 0)) - Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
        Zl_����Ԥ����¼_Insert(r_Ԥ��.Id, r_Ԥ��.No, 1, r_Ԥ��.���, n_����id, n_����id);
        n_��Ԥ����� := n_��Ԥ����� - Nvl(r_Ԥ��.���, 0);
      End Loop;
      If n_��Ԥ����� <> 0 Then
        v_Err_Msg := '�����Ԥ�����������ʵ�ʲ���,����!';
        Raise Err_Item;
      End If;
    End If;
  
    Update �������׼�¼
    Set ҵ�����id = n_����id
    Where ��ˮ�� = Nvl(r_���㷽ʽ.������ˮ��, '-') And ��� = v_����� And ҵ������ = 2;
  End Loop;

  --���ѿ�����
  If v_���ѿ����� Is Not Null Then
    v_���ѿ����� := Substr(v_���ѿ�����, 3);
  End If;

  n_����   := Round(Nvl(n_�����ܶ�, 0) - Nvl(n_���ʽ��, 0), 6);
  v_���㷽ʽ := Null;
  If Abs(Nvl(n_����, 0)) > 1 Then
    v_Err_Msg := '���������������1.00��С��-1.00Ԫ,��������ʲ���,����!';
    Raise Err_Item;
  End If;

  n_�����ܶ� := n_���ʽ��;

  n_���ʽ�� := 0;
  If Nvl(n_����, 0) <> 0 Then
    Select Nvl(Max(����), '����') Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 9;
    n_���ʽ�� := Nvl(n_����, 0);
  End If;
  If Nvl(n_����, 0) <> 0 Or v_���ѿ����� Is Not Null Then
    Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, n_���ʽ��, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��, Null, Null,
                     Null, Null, Null, Null, Null, Null, Null, v_���ѿ�����);
  End If;

  --��������Ϣ�ܶ�������ܶ��Ƿ���ȷ
  Select Sum(��Ԥ��) Into n_���ʽ�� From ����Ԥ����¼ Where ����id = n_����id;
  If Round(n_���ʽ��, 6) <> Round(n_�����ܶ�, 6) Then
  
    v_Err_Msg := '����Ľ���ϼƽ����ʵ�ʽ��ʽ��ϼƲ���,���������!';
    Raise Err_Item;
  End If;

  Update ����Ԥ����¼ Set У�Ա�־ = 0 Where ����id = n_����id And Nvl(У�Ա�־, 0) <> 0;
  v_Temp := '<CZSJ>' || To_Char(d_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Third_Settlement;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.150.0004' Where ���=&n_System;
Commit;
