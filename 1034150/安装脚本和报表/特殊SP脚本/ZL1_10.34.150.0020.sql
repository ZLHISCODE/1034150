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
--134969:���ϴ�,2019-01-03,ʹ��Ԥ��֧��ʱ����Ƿ����
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 0, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
  ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;

  v_Err_Msg Varchar2(255);
  Err_Item    Exception;
  Err_Special Exception;

  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;
  v_�ű�     ������ü�¼.���㵥λ%Type;
  v_����     ������ü�¼.��ҩ����%Type;
  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_��ӡid        Ʊ�ݴ�ӡ����.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;
  n_���ѿ�id       ���ѿ�Ŀ¼.Id%Type;
  n_���ƿ�         Number;

  d_Date       Date;
  d_ԤԼʱ��   ������ü�¼.����ʱ��%Type;
  d_����ʱ��   Date;
  d_�Ŷ�ʱ��   Date;
  n_ʱ��       Number := 0;
  n_����       Number := 0;
  v_�Ŷ����   �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ   ������Ϣ.����ģʽ%Type;
  n_Ʊ��       Ʊ��ʹ����ϸ.Ʊ��%Type;
  v_���ʽ   ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_����Ա���� ���˹Һż�¼.������%Type;
  n_����ģʽ   Number := 0;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);

  --��ȡ���㷽ʽ����
  Begin
    Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�ֽ� := '�ֽ�';
  End;
  Begin
    Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
  Exception
    When Others Then
      v_�����ʻ� := '�����ʻ�';
  End;
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(������) Into v_����Ա���� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ In (1, 3) And NO = No_In;
      If v_����Ա���� Is Null Then
        v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�ȡ��';
        Raise Err_Item;
      Else
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�����';
          Raise Err_Special;
        Else
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
          Raise Err_Item;
        End If;
      End If;
  End;

  --�ж��Ƿ��ʱ��
  Begin
    Select 1
    Into n_ʱ��
    From Dual
    Where Exists (Select 1
           From �ҺŰ���ʱ�� A, �ҺŰ��� B
           Where a.����id = b.Id And b.���� = v_�ű� And Rownum < 2
           Union All
           Select 1
           From �Һżƻ�ʱ�� C, �ҺŰ��żƻ� D ��
           Where c.�ƻ�id = d.Id And d.���� = v_�ű� And d.��Чʱ�� > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_ʱ�� := 0;
  End;
  --��ʱ�εĺű�ֻ�ܵ������
  If n_ʱ�� = 1 And ��������_In = 0 And n_����ģʽ = 0 Then
    If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
      Raise Err_Item;
    End If;
  End If;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;
  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
      
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
          Begin
            Select 1 Into n_���� From �Һ����״̬ Where ���� = v_�ű� And ���� = Trunc(Sysdate) And ��� = v_����;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If n_���� = 0 Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
          Else
            --�����ѱ�ʹ�õ����
            Begin
              v_���� := 1;
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                Select Min(��� + 1)
                Into v_����
                From �Һ����״̬ A
                Where ���� = v_�ű� And ���� = Trunc(Sysdate) And Not Exists
                 (Select 1 From �Һ����״̬ Where ���� = a.���� And ���� = a.���� And ��� = a.��� + 1);
                Insert Into �Һ����״̬
                  (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
                Values
                  (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            End;
          End If;
        Else
          Update �Һ����״̬
          Set ״̬ = 1, �Ǽ�ʱ�� = Sysdate
          Where Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_���� And ���� = v_�ű� And ״̬ = 2;
          If Sql% NotFound Then
            Begin
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update �Һ����״̬
        Set ��� = ����_In, ״̬ = 1, �Ǽ�ʱ�� = Sysdate
        Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(d_����ʱ��), v_����, 1, ����Ա����_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      Begin
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
        Values
          (v_�ű�, Trunc(Sysdate), ����_In, 1, ����Ա����_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '���' || ����_In || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
      End;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ժҪ = Nvl(ժҪ_In, ժҪ)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      ժҪ = Nvl(ժҪ_In, ժҪ)
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, Nvl(ժҪ_In, ժҪ), v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In,
               Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��, Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
    If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      
        Begin
          Select 1,
                 Case
                   When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_�Ŷ�, n_�����Ŷ�
          From �ŶӽкŶ���
          Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
        Exception
          When Others Then
            n_�Ŷ� := 0;
        End;
        If n_�Ŷ� = 0 Then
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
        --ԤԼ����ʱ���ı��¼��־
        Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
      End Loop;
    End If;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And
     Nvl(���ʷ���_In, 0) = 0 Then
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������,
       ��������)
    Values
      (n_Ԥ��id, 4, 1, No_In, ����id_In, Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), d_Date, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
       n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ����id_In, 4);
  
    If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
    
      n_���ѿ�id := Null;
      Begin
        Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count = 0 Then
        v_Err_Msg := '[ZLSOFT]û�з���ԭ���㿨����Ӧ���,���ܼ���������[ZLSOFT]';
        Raise Err_Item;
      End If;
      If n_���ƿ� = 1 Then
        Select ID
        Into n_���ѿ�id
        From ���ѿ�Ŀ¼
        Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
              ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
      End If;
      Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, ���㷽ʽ_In, �ֽ�֧��_In, ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In, n_Ԥ��id);
    End If;
  
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    n_Ԥ����� := Ԥ��֧��_In;
    For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
      n_��ǰ��� := Case
                  When r_Deposit.��� - n_Ԥ����� < 0 Then
                   r_Deposit.���
                  Else
                   n_Ԥ�����
                End;
      If r_Deposit.����id = 0 Then
        --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
        Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, Nvl(1, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ƿ��Ѿ�������
      If r_Deposit.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
      Else
        n_Ԥ����� := 0;
      End If;
    
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    End Loop;
    IF n_Ԥ����� > 0 Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����С�ڱ���֧����� ' || Ltrim(To_Char(Ԥ��֧��_In, '9999999990.00')) || '�����ܼ���������';
      Raise Err_Item;
    End IF;
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       Null, Null, Null, Null, Null, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(�ֽ�֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 0) = 0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + �ֽ�֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�)
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (����Ա����_In, Nvl(���㷽ʽ_In, v_�ֽ�), 1, �ֽ�֧��_In);
      n_����ֵ := �ֽ�֧��_In;
    
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 0) = 0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) >= d_Date;
    End If;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_Insert;
/

--134969:���ϴ�,2019-01-03,ʹ��Ԥ��֧��ʱ����Ƿ����
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_����_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      Varchar2, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 0, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
  ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;

  v_Err_Msg Varchar2(255);
  Err_Item    Exception;
  Err_Special Exception;
  v_����Ա���� ���˹Һż�¼.������%Type;
  v_�ֽ�       ���㷽ʽ.����%Type;
  v_�����ʻ�   ���㷽ʽ.����%Type;
  v_��������   �ŶӽкŶ���.��������%Type;
  v_�ű�       ������ü�¼.���㵥λ%Type;
  v_����       ������ü�¼.��ҩ����%Type;
  v_�ŶӺ���   �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_��ӡid        Ʊ�ݴ�ӡ����.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;
  n_���ѿ�id       ���ѿ�Ŀ¼.Id%Type;
  n_���ƿ�         Number;

  d_Date         Date;
  d_ԤԼʱ��     ������ü�¼.����ʱ��%Type;
  d_����ʱ��     Date;
  d_�Ŷ�ʱ��     Date;
  n_ʱ��         Number := 0;
  n_����         Number := 0;
  v_��������     Varchar2(2000);
  v_��ǰ����     Varchar2(500);
  n_������     ����Ԥ����¼.��Ԥ��%Type;
  v_�������     ����Ԥ����¼.�������%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־   Number(3);
  v_�Ŷ����     �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ     ������Ϣ.����ģʽ%Type;
  n_Ʊ��         Ʊ��ʹ����ϸ.Ʊ��%Type;
  v_���ʽ     ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_����ģʽ     Number := 0;
  n_�����¼id   ���˹Һż�¼.�����¼id%Type;
  n_�³����¼id ���˹Һż�¼.�����¼id%Type;
  n_��Դid       �ٴ������¼.��Դid%Type;
  n_ԤԼ˳���   �ٴ�������ſ���.ԤԼ˳���%Type;
  n_�ɷ�ʱ��     �ٴ������¼.�Ƿ��ʱ��%Type;
  n_����ſ���   �ٴ������¼.�Ƿ���ſ���%Type;
  n_�ɿ���id     �ٴ������¼.����id%Type;
  n_����Ŀid     �ٴ������¼.��Ŀid%Type;
  n_��ҽ��id     �ٴ������¼.ҽ��id%Type;
  n_�Һ�ģʽ     Number(3);
  d_����ʱ��     Date;
  v_Paratemp     Varchar2(500);
  v_Registtemp   Varchar2(500);
  n_���         Number(3);
  n_��ſ���     �ٴ������¼.�Ƿ���ſ���%Type;
  v_���ϰ�ʱ��   �ٴ������¼.�ϰ�ʱ��%Type;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  v_Paratemp      := Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);
  n_�Һ�ģʽ      := To_Number(Substr(v_Paratemp, 1, 1));
  If n_�Һ�ģʽ = 1 Then
    Begin
      d_����ʱ�� := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  End If;

  --��ȡ���㷽ʽ����
  Begin
    Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�ֽ� := '�ֽ�';
  End;
  Begin
    Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
  Exception
    When Others Then
      v_�����ʻ� := '�����ʻ�';
  End;
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ, �����¼id
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ, n_�����¼id
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(������) Into v_����Ա���� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ In (1, 3) And NO = No_In;
      If v_����Ա���� Is Null Then
        v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�ȡ��';
        Raise Err_Item;
      Else
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�����';
          Raise Err_Special;
        Else
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
          Raise Err_Item;
        End If;
      End If;
  End;

  --�ж��Ƿ��ʱ��
  Select Nvl(�Ƿ��ʱ��, 0), ��Դid, Nvl(�Ƿ���ſ���, 0)
  Into n_ʱ��, n_��Դid, n_��ſ���
  From �ٴ������¼
  Where ID = n_�����¼id;

  If n_ʱ�� = 1 And ��������_In = 0 And n_����ģʽ = 0 Then
    If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
      Raise Err_Item;
    End If;
  End If;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;

  If d_����ʱ�� Is Not Null Then
    If d_����ʱ�� < d_����ʱ�� Then
      v_Err_Msg := '��ǰԤԼ�Һŵ����ڳ�����Ű�ģʽ���ţ�������' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '֮ǰ����!';
      Raise Err_Item;
    End If;
  End If;

  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where ��� = v_���� And ��¼id = n_�����¼id;
        
          Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
          Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
          From �ٴ������¼
          Where ID = n_�����¼id;
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                  Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
        
          Begin
            Select 1
            Into n_����
            From �ٴ�������ſ���
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Exception
            When Others Then
              n_���� := 0;
          End;
        
          If n_���� = 1 Then
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Else
            --�����ѱ�ʹ�õ����
            Select Min(���) Into v_���� From �ٴ�������ſ��� Where ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
            If v_���� Is Null Then
              v_Err_Msg := '���յ���û�п������,�޷�����!';
              Raise Err_Item;
            End If;
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          End If;
        Else
          Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
          Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
          From �ٴ������¼
          Where ID = n_�����¼id;
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                  Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
          Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
          Returning ԤԼ˳��� Into n_ԤԼ˳���;
        
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
          Where ��� = v_���� And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
          If Sql% RowCount = 0 Then
            v_Err_Msg := '���յ������' || v_���� || '�ѱ�������ʹ��,�޷�����.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
        Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
        From �ٴ������¼
        Where ID = n_�����¼id;
        Begin
          Select ID
          Into n_�³����¼id
          From �ٴ������¼
          Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
            Raise Err_Item;
        End;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
        Returning ԤԼ˳��� Into n_ԤԼ˳���;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
        Where ��� = ����_In And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���յ������' || ����_In || '�ѱ�������ʹ��,�޷�����.';
          Raise Err_Item;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id;
      
      End If;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ժҪ = Nvl(ժҪ_In, ժҪ)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('�Һ��Ű�ģʽ');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_����ʱ�� Then
        v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || 'δ���ó�����Ű�ģʽ,Ŀǰ�޷�����!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_���
      From �ٴ������¼
      Where ID = Nvl(n_�³����¼id, n_�����¼id) And d_����ʱ�� Between ͣ�￪ʼʱ�� And ͣ����ֹʱ��;
    Exception
      When Others Then
        n_��� := 0;
    End;
    If n_��� = 1 And Not (n_ʱ�� = 1 And n_��ſ��� = 1) Then
      v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '�İ����Ѿ���ͣ��,�޷�����!';
      Raise Err_Item;
    End If;
  End If;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      �����¼id = Nvl(n_�³����¼id, n_�����¼id), ժҪ = Nvl(ժҪ_In, ժҪ)
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ, �����¼id)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, Nvl(ժҪ_In, ժҪ), v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In,
               Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��, Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, Nvl(n_�³����¼id, n_�����¼id)
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
    If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      
        Begin
          Select 1,
                 Case
                   When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_�Ŷ�, n_�����Ŷ�
          From �ŶӽкŶ���
          Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
        Exception
          When Others Then
            n_�Ŷ� := 0;
        End;
        If n_�Ŷ� = 0 Then
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
      End Loop;
    End If;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 Then
      v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, Null, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4, v_�������);
          If Nvl(���㿨���_In, 0) <> 0 Then
            n_���ѿ�id := Null;
            Begin
              Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 0 Then
              v_Err_Msg := 'û�з���ԭ���㿨����Ӧ���,���ܼ���������';
              Raise Err_Item;
            End If;
            If n_���ƿ� = 1 Then
              Select ID
              Into n_���ѿ�id
              From ���ѿ�Ŀ¼
              Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
                    ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
            End If;
            Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, v_���㷽ʽ, n_������, ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In, n_Ԥ��id);
          End If;
        End If;
      
        If Nvl(���½������_In, 0) = 0 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + n_������
          Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�)
          Returning ��� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, Nvl(v_���㷽ʽ, v_�ֽ�), 1, n_������);
            n_����ֵ := n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
    End If;
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    n_Ԥ����� := Ԥ��֧��_In;
    For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
      n_��ǰ��� := Case
                  When r_Deposit.��� - n_Ԥ����� < 0 Then
                   r_Deposit.���
                  Else
                   n_Ԥ�����
                End;
      If r_Deposit.����id = 0 Then
        --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
        Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, Nvl(1, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ƿ��Ѿ�������
      If r_Deposit.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
      Else
        n_Ԥ����� := 0;
      End If;
    
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    End Loop;
    IF n_Ԥ����� > 0 Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����С�ڱ���֧����� ' || Ltrim(To_Char(Ԥ��֧��_In, '9999999990.00')) || '�����ܼ���������';
      Raise Err_Item;
    End IF;
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       Null, Null, Null, Null, Null, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 0) = 0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) >= d_Date;
    End If;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_����_Insert;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.150.0020' Where ���=&n_System;
Commit;
