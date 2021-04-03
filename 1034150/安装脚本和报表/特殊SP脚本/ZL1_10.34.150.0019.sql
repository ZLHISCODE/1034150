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
--135887:���˺�,2018-12-20,�����������ʱ�������˻���Ԥ��������
Create Or Replace Procedure Zl_���˽��ʼ�¼_Delete
(
  No_In           ���˽��ʼ�¼.No%Type,
  ����Ա���_In   ���˽��ʼ�¼.����Ա���%Type,
  ����Ա����_In   ���˽��ʼ�¼.����Ա����%Type,
  �����_In     ����Ԥ����¼.��Ԥ��%Type := 0, --ҽ����Ԥ�����ֽ���������
  �������Ͻ���_In Varchar2 := Null, --���㷽ʽ|������|�������||......
  Ԥ�����ֽ�_In   Number := 0, --��Ԥ�������ֽ�ʱ�����㷽ʽ�����ͨ�������������Ͻ���_In����
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  ����ʱ��_In     Date := Null,
  ��Ԥ��id_In     ����Ԥ����¼.Id%Type := Null, --������ʱ����صĽ���ֵ��Ԥ����ʱ��д
  Ʊ�ݺ�_In       ���˽��ʼ�¼.ʵ��Ʊ��%Type := Null,
  ����id_In       Ʊ�����ü�¼.Id%Type := Null,
  Ʊ��_In         Ʊ��ʹ����ϸ.Ʊ��%Type := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --���α�����Ԥ����¼�����Ϣ
  Cursor c_Deposit(v_Id ����Ԥ����¼.����id%Type) Is
    Select ����id, ��¼����, ���㷽ʽ, ��Ԥ��, Ԥ����� From ����Ԥ����¼ Where ����id = v_Id;
  r_Depositrow c_Deposit%RowType;

  --���α����ڴ��������ػ��ܱ�
  Cursor c_Money(v_Id ����Ԥ����¼.����id%Type) Is
    Select NO, ��������id, ���˿���id, ִ�в���id, ���˲���id, ����id, ��ҳid, ������Ŀid, �����־, ���ʽ��
    From סԺ���ü�¼
    Where ����id = v_Id
    Union All
    Select NO, ��������id, ���˿���id, ִ�в���id, 0 As ���˲���id, ����id, 0 As ��ҳid, ������Ŀid, �����־, ���ʽ��
    From ������ü�¼
    Where ����id = v_Id;

  r_Moneyrow c_Money%RowType;

  --���α�������˵������Ϣ
  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����, a.�Ա�, a.����, a.סԺ��, a.�����, b.��ҳid, b.��Ժ����, b.��ǰ����id, b.��Ժ����id, Nvl(b.�ѱ�, a.�ѱ�) As �ѱ�, a.����, c.���� As ���ʽ
    From ������Ϣ A, ������ҳ B, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.����id = b.����id(+) And Nvl(a.��ҳid, 0) = b.��ҳid(+) And a.ҽ�Ƹ��ʽ = c.����(+);
  r_Pati c_Pati%RowType;

  --���̱���
  v_�������� Varchar2(500);
  v_��ǰ���� Varchar2(50);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_ʵ��Ʊ�� ����Ԥ����¼.ʵ��Ʊ��%Type;
  v_���no   סԺ���ü�¼.No%Type;
  v_���     ���㷽ʽ.����%Type;
  n_����id   ������Ϣ.����id%Type;

  n_ԭid   ���˽��ʼ�¼.Id%Type;
  n_����id ���˽��ʼ�¼.Id%Type;
  n_��ӡid Ʊ�ݴ�ӡ����.Id%Type;

  n_��Դ     Number; --1-����;2-סԺ;3-�����סԺ
  n_����ֵ   �������.Ԥ�����%Type;
  n_��id     ����ɿ����.Id%Type;
  n_Ԥ����� Number;
  d_Date     Date;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_������id ���˽��ʼ�¼.Id%Type;
  v_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;

  n_Ԥ���ϼ� ����Ԥ����¼.��Ԥ��%Type;
  n_���ʺϼ� סԺ���ü�¼.���ʽ��%Type;

Begin
  n_��id := Zl_Get��id(����Ա����_In);

  Select ���� Into v_��� From ���㷽ʽ Where ���� = 9 And Rownum = 1;

  Begin
    Select ID, ����id, ʵ��Ʊ�� Into n_ԭid, n_����id, v_ʵ��Ʊ�� From ���˽��ʼ�¼ Where ��¼״̬ = 1 And NO = No_In;
    --���һ�δ�ӡ������
    Select Max(ID)
    Into n_��ӡid
    From (Select b.Id
           From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
           Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 3 And b.No = No_In
           Order By a.ʹ��ʱ�� Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      Begin
        v_Err_Msg := 'û�з���Ҫ���ϵĽ��ʵ���,�����Ѿ����ϣ�';
        Raise Err_Item;
      End;
  End;

  Open c_Pati(n_����id);
  Fetch c_Pati
    Into r_Pati; --���ϵͳ���ô˹���,�������ʱû�в�����Ϣ

  d_Date := ����ʱ��_In;
  If d_Date Is Null Then
    Select Sysdate Into d_Date From Dual;
  End If;

  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 3, No_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��_In, Ʊ�ݺ�_In, 1, 6, ����id_In, v_��ӡid, d_Date, ����Ա����_In);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  End If;

  n_����id := ����id_In;
  If Nvl(n_����id, 0) = 0 Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  End If;

  --���˽��ʼ�¼
  Insert Into ���˽��ʼ�¼
    (ID, NO, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, ��ʼ����, ��������, �շ�ʱ��, ��ע, ԭ��, �ɿ���id, ��������)
    Select n_����id, NO, ʵ��Ʊ��, 2, ��;����, ����id, ����Ա���_In, ����Ա����_In, ��ʼ����, ��������, d_Date, ��ע, ԭ��, n_��id, ��������
    From ���˽��ʼ�¼
    Where ID = n_ԭid;

  Update ���˽��ʼ�¼ Set ��¼״̬ = 3 Where ID = n_ԭid;

  --�����ջ�Ʊ��(������ǰû��ʹ��Ʊ��,�޷��ջ�)
  If n_��ӡid Is Not Null Then
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
      Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
      From Ʊ��ʹ����ϸ
      Where ��ӡid = n_��ӡid And Ʊ�� In (1, 3) And ���� = 1;
  End If;

  --����Ԥ����¼(��Ԥ�����ɿ�)
  If �������Ͻ���_In Is Null Then
    --������ͨ�Ľ�����Ϣ(�����������˿�
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
       �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
      Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ����id, ��ҳid, ����id, Null,
             ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date, ����Ա����_In, ����Ա���_In, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, ������λ, 2
      From ����Ԥ����¼
      Where ����id = n_ԭid And Mod(��¼����, 10) <> 1 And
            ((�����id Is Not Null And Nvl(��Ԥ��, 0) > 0) Or (Mod(��¼����, 10) <> 1 And �����id Is Null));
  
    --����Ԥ�������������Ѿ��˵�Ԥ��
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
       �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
      Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11 As ��¼����, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
             d_Date, ����Ա����_In, ����Ա���_In, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 2
      From (Select a.No, Sum(a.��Ԥ��) As ��Ԥ��, Max(a.ʵ��Ʊ��) As ʵ��Ʊ��, Max(a.��¼״̬) As ��¼״̬, Max(a.����id) As ����id,
                    Max(a.��ҳid) As ��ҳid, Max(a.����id) As ����id, Max(a.���㷽ʽ) As ���㷽ʽ, Max(a.�������) As �������, Max(a.ժҪ) As ժҪ,
                    Max(a.�ɿλ) As �ɿλ, Max(a.��λ������) As ��λ������, Max(a.��λ�ʺ�) As ��λ�ʺ�, Max(a.Ԥ�����) As Ԥ�����,
                    Max(a.�����id) As �����id, Max(a.���㿨���) As ���㿨���, Max(a.����) As ����, Max(a.������ˮ��) As ������ˮ��,
                    Max(a.����˵��) As ����˵��, Max(a.������λ) As ������λ
             From (Select a.No, a.��Ԥ��, a.ʵ��Ʊ��, a.��¼״̬, a.����id, a.��ҳid, a.����id, a.���㷽ʽ, a.�������, a.ժҪ, a.�ɿλ, a.��λ������,
                           a.��λ�ʺ�, a.Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ��, a.����˵��, a.������λ
                    From ����Ԥ����¼ A
                    Where a.����id = n_ԭid And Mod(��¼����, 10) = 1
                    Union All
                    Select b.No, -1 * a.��� As ��Ԥ��, b.ʵ��Ʊ��, b.��¼״̬, b.����id, b.��ҳid, b.����id, b.���㷽ʽ, b.�������, b.ժҪ, b.�ɿλ,
                           b.��λ������, b.��λ�ʺ�, b.Ԥ�����, b.�����id, b.���㿨���, b.����, b.������ˮ��, b.����˵��, b.������λ

                    From �����˿���Ϣ A, ����Ԥ����¼ B
                    Where a.����id = n_ԭid And a.��¼id = b.Id) A
             Group By a.No) A
      Where Nvl(a.��Ԥ��, 0) <> 0;
  
    --���ѿ�����
    For c_���ѿ����� In (Select a.Id, a.���㿨���, Nvl(b.����, '���ѿ�') As ������
                    From ����Ԥ����¼ A, �����ѽӿ�Ŀ¼ B
                    Where a.���㿨��� = b.���(+) And a.����id = n_ԭid And Nvl(a.���㿨���, 0) <> 0) Loop
    
      For c_���ѿ� In (Select a.Id, a.�ӿڱ��, a.���ѿ�id, a.���, a.��¼״̬, a.���㷽ʽ, a.������, a.����, a.������ˮ��, a.����ʱ��, a.��ע, a.�����־,
                           b.ͣ������, b.����ʱ��
                    From ���˿������¼ A, ���ѿ�Ŀ¼ B
                    Where a.���ѿ�id = b.Id(+) And
                          a.Id In (Select ������id From ���˿�������� Where Ԥ��id = c_���ѿ�����.Id)) Loop
      
        If Nvl(c_���ѿ�.ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '����Ϊ"' || c_���ѿ�.���� || '"��' || c_���ѿ�����.������ || '�Ѿ�������ͣ�ã������ٽ��н�������,���飡';
          Raise Err_Item;
        End If;
      
        If c_���ѿ�.����ʱ�� < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '����Ϊ"' || c_���ѿ�.���� || '"��' || c_���ѿ�����.������ || '�Ѿ����գ������˷�,���飡';
          Raise Err_Item;
        End If;
        Select ���˿������¼_Id.Nextval Into n_������id From Dual;
      
        Insert Into ���˿������¼
          (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
        Values
          (n_������id, c_���ѿ�����.���㿨���, c_���ѿ�.���ѿ�id, c_���ѿ�.���, 2, c_���ѿ�.���㷽ʽ, -1 * c_���ѿ�.������, c_���ѿ�.����, c_���ѿ�.������ˮ��, d_Date,
           Null, 0);
      
        Insert Into ���˿��������
          (Ԥ��id, ������id)
          Select ID, n_������id
          From ����Ԥ����¼
          Where ����id = n_����id And ���㿨��� = Nvl(c_���ѿ�����.���㿨���, 0);
      
        Update ���ѿ�Ŀ¼ Set ��� = ��� + c_���ѿ�.������ Where ID = c_���ѿ�.���ѿ�id;
        If Sql%NotFound Then
          v_Err_Msg := '����Ϊ' || c_���ѿ�.���� || '��' || c_���ѿ�����.������ || 'δ�ҵ�!';
          Raise Err_Item;
        End If;
      End Loop;
    End Loop;
  
  Else
    --1.�ȴ����Ԥ������
    If Ԥ�����ֽ�_In = 0 Then
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ����id, ��ҳid, ����id,
               Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date, ����Ա����_In, ����Ա���_In, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 2
        From ����Ԥ����¼
        Where ����id = n_ԭid And ��¼���� In (1, 11) And Nvl(��Ԥ��, 0) <> 0;
    End If;
  
    --2.�ٴ�����ʽ���,����ҽ���ͷ�ҽ��
    v_�������� := �������Ͻ���_In || ' ||'; --�Կո�ֿ���|��β,û�н�������
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_������� := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, No_In, v_ʵ��Ʊ��, 12, 1, n_����id, r_Pati.��ҳid, r_Pati.��Ժ����id, Null, v_���㷽ʽ, v_�������, '���������˿�',
         Null, Null, Null, d_Date, ����Ա����_In, ����Ա���_In, -1 * n_������, n_����id, n_��id, 2);
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;
  --ȷ�����ʵķ��ü�¼��Դ
  Begin
    Select Case
             When Nvl(Max(סԺ), 0) = 1 And Nvl(Max(����), 0) = 1 Then
              3
             When Nvl(Max(סԺ), 0) = 1 Then
              2
             Else
              1
           End
    Into n_��Դ
    From (Select 1 As סԺ, 0 As ����
           From סԺ���ü�¼
           Where ����id = n_ԭid And Rownum = 1
           Union All
           Select 0 As סԺ, 1 As ����
           From ������ü�¼
           Where ����id = n_ԭid And Rownum = 1);
  
  Exception
    When Others Then
      n_��Դ := 3;
  End;

  If �����_In <> 0 Then
    Update ����Ԥ����¼
    Set ��Ԥ�� = ��Ԥ�� + �����_In
    Where NO = No_In And ��¼���� = 12 And ��¼״̬ = 1 And ����id = n_����id And ���㷽ʽ = v_���;
    If Sql%RowCount = 0 Then
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, No_In, v_ʵ��Ʊ��, 12, 1, n_����id, r_Pati.��ҳid, r_Pati.��Ժ����id, Null, v_���, Null, '���������˿�', Null,
         Null, Null, d_Date, ����Ա����_In, ����Ա���_In, �����_In, n_����id, n_��id, 2);
    End If;
  End If;

  If n_��Դ = 2 Or n_��Դ = 3 Then
    --���Ͻ��ʶ�Ӧ�ķ��ü�¼:������ԭʼ���ʲ����������Ŀ
    Insert Into סԺ���ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id,
       ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id,
       ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��, ����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ,
       �ɿ���id, ҽ��С��id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�,
             ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����,
             �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Null, Null, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����,
             ִ��ʱ��, ����Ա����, ����Ա���, -1 * ���ʽ��, n_����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id, ҽ��С��id
      From סԺ���ü�¼
      Where ����id = n_ԭid;
  End If;

  If n_��Դ = 1 Or n_��Դ = 3 Then
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����,
       �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��,
       ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��, ����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id,
             ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����,
             ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Null, Null, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���,
             -1 * ���ʽ��, n_����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id
      From ������ü�¼
      Where ����id = n_ԭid;
  End If;
  --��ػ��ܱ���
  For r_Depositrow In c_Deposit(n_����id) Loop
    If r_Depositrow.��¼���� In (1, 11) Then
    
      --�������(Ԥ��)
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - r_Depositrow.��Ԥ�� --ע:�µĽ���ID�������Ǹ������
      Where ����id = r_Depositrow.����id And ���� = Nvl(r_Depositrow.Ԥ�����, 2) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (r_Depositrow.����id, 1, Nvl(r_Depositrow.Ԥ�����, 2), -1 * r_Depositrow.��Ԥ��, 0);
        n_����ֵ := -1 * r_Depositrow.��Ԥ��;
      End If;
    
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete �������
        Where ���� = 1 And ����id = r_Depositrow.����id And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    
    Else
      --��Ա�ɿ����,ҽ����֧�����ϵĽ��㷽ʽ���µ�Ԥ���������ѱ�����Ϊ�����ֽ�,
      --�˴��ü�,��ʾ�ջ��˸����˵��ֽ�(����ʱ,�˿��Ǹ�,����ʱ����)
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + r_Depositrow.��Ԥ��
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Depositrow.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Depositrow.���㷽ʽ, 1, r_Depositrow.��Ԥ��);
        n_����ֵ := -1 * r_Depositrow.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Depositrow.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End If;
  End Loop;

  For r_Moneyrow In c_Money(n_����id) Loop
    --������� ,������ѽ���,���Բ���Ҫ�������������ܱ�
    If Nvl(v_���no, 'sc') <> Nvl(r_Moneyrow.No, 'sc') Then
      If Nvl(r_Moneyrow.�����־, 0) = 1 Or Nvl(r_Moneyrow.�����־, 0) = 2 Then
        n_Ԥ����� := r_Moneyrow.�����־;
      Elsif Nvl(r_Moneyrow.��ҳid, 0) = 0 Or Nvl(r_Moneyrow.�����־, 0) = 4 Then
        --���:���ﲡ��
        n_Ԥ����� := 1;
      Else
        n_Ԥ����� := 2;
      End If;
    
      Update �������
      Set ������� = Nvl(�������, 0) - r_Moneyrow.���ʽ�� --ע:�µĽ���ID�������Ǹ������
      Where ����id = r_Moneyrow.����id And ���� = n_Ԥ����� And ���� = 1
      Returning ������� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (r_Moneyrow.����id, 1, n_Ԥ�����, 0, -1 * r_Moneyrow.���ʽ��);
        n_����ֵ := -1 * r_Moneyrow.���ʽ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete �������
        Where ����id = r_Moneyrow.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - r_Moneyrow.���ʽ��
      Where ����id = r_Moneyrow.����id And Nvl(��ҳid, 0) = Nvl(r_Moneyrow.��ҳid, 0) And
            Nvl(���˲���id, 0) = Nvl(r_Moneyrow.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Moneyrow.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(r_Moneyrow.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(r_Moneyrow.ִ�в���id, 0) And
            ������Ŀid + 0 = r_Moneyrow.������Ŀid And ��Դ;�� + 0 = r_Moneyrow.�����־;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (r_Moneyrow.����id, Decode(r_Moneyrow.��ҳid, Null, Null, 0, Null, r_Moneyrow.��ҳid),
           Decode(r_Moneyrow.���˲���id, Null, Null, 0, Null, r_Moneyrow.���˲���id), r_Moneyrow.���˿���id, r_Moneyrow.��������id,
           r_Moneyrow.ִ�в���id, r_Moneyrow.������Ŀid, r_Moneyrow.�����־, -1 * r_Moneyrow.���ʽ��);
      End If;
    End If;
  End Loop;

  If Nvl(��Ԥ��id_In, 0) <> 0 Then
    --����ʱ���˿����ֵ��Ԥ�����ʻ�,��������Ǳ��ν��ʽɴ��
    Update ����Ԥ����¼ Set ����id = ����id_In Where ID = ��Ԥ��id_In And ����id Is Null;
    If Sql%NotFound Then
      v_Err_Msg := 'δ�ҵ���Ӧ��Ԥ�����¼��';
      Raise Err_Item;
    End If;
  End If;

  --��Ҫ����˿�ϼ��Ƿ����
  Select Round(Sum(��Ԥ��), 5), Round(Sum(���ʺϼ�), 5)
  Into n_Ԥ���ϼ�, n_���ʺϼ�
  From (Select Sum(��Ԥ��) As ��Ԥ��, 0 As ���ʺϼ�
         From ����Ԥ����¼
         Where ����id = n_����id
         Union All
         Select 0 As ��Ԥ��, Sum(���ʽ��) As ���ʺϼ�
         From סԺ���ü�¼
         Where ����id = n_����id
         Union All
         Select 0 As ��Ԥ��, Sum(���ʽ��) As ���ʺϼ�
         From ������ü�¼
         Where ����id = n_����id);

  If Nvl(n_Ԥ���ϼ�, 0) <> Nvl(n_���ʺϼ�, 0) Then
    v_Err_Msg := '���ν������Ϻϼ�(' || Trim(To_Char(Nvl(n_Ԥ���ϼ�, 0), '99999999999.99999')) || ')�뱾�ν������Ϸ��úϼ�(' ||
                 Trim(To_Char(Nvl(n_���ʺϼ�, 0), '99999999999.99999')) || ')���ȣ����鵱ǰ�����������ϸ�ϼ��Ƿ�һ�£�';
    Raise Err_Item;
  End If;

  Close c_Pati;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽��ʼ�¼_Delete;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.150.0019' Where ���=&n_System;
Commit;
