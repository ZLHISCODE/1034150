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
--134222:����,2018-11-23,����Oracle����Zl_Third_Getregistalter,�ֶδ������ص�XML�ַ���
Create Or Replace Procedure Zl_Third_Getregistalter
(
  Xml_In  Xmltype,
  Xml_Out Out Xmltype
) Is
  -----------------------------------------------
  --���ܣ���ȡ���������ͣ���ﰲ��
  --��Σ�XML_IN
  --<IN>
  --  <JSKLB>���㿨���</JSKLB>
  --  <RQ>����</RQ>
  --</IN>
  --����:XML_OUT
  --<OUTPUT>
  --  <TZLISTS>          //ͣ���б�
  --    <ITEM>
  --      <HM>����</HM>
  --      <YSID>ҽ��ID</YSID>
  --      <YS>ҽ������</YS>
  --      <KSSJ>ͣ�￪ʼʱ��</KSSJ>
  --      <JSSJ>ͣ�����ʱ��</JSSJ>
  --      <BRLIST>
  --        <INFO>
  --          <YYNO>ԤԼ���ݺ�</YYNO>
  --          <BRID>����ID</BRID>
  --          <YYSJ>ԤԼʱ��</YYSJ>
  --          <CZSJ>����ʱ��</CZSJ>
  --          <YYKS>ԤԼ����</YYKS>
  --          <GHLX>����</GHLX>
  --          <YSXM>ҽ������</YSXM>
  --        </INFO>
  --      </BRLIST>
  --    </ITEM>
  --  </TZLISTS>
  --  <HZLISTS>          //�����б�
  --    <ITEM>
  --      <BRID>����ID</BRID>
  --      <YYSJ>ԤԼ�Ĳ���ʱ��</YYSJ>
  --      <YSJ>ԭԤԼʱ��</YSJ>
  --      <YHM>ԭ����</YHM>
  --      <YYS>ԭҽ��</YYS>
  --      <YZC>ԭҽ����ְ��</YZC>
  --      <XSJ>��ԤԼʱ��</XSJ>
  --      <XHM>�ֺ���</XHM>
  --      <XYS>��ҽ��</XYS>
  --      <XZC>��ҽ����ְ��</XZC>
  --    </ITEM>
  --  </HZLIST>
  --</OUTPUT>
  -----------------------------------------------------

  d_Date     Date;
  v_Jsklb    Varchar2(100);
  n_�����id ҽ�ƿ����.Id%Type;
  n_Cnt      Number(3);
  v_Temp     Clob;
  v_Brinfo   Varchar2(4000);
  d_����ʱ�� Date;
  v_Para     Varchar2(2000);
  n_Exists   Number(3);
  n_�Һ�ģʽ Number(3);
  x_Templet  Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/JSKLB') Into v_Jsklb From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd')
  Into d_Date
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select b.Id Into n_�����id From ҽ�ƿ���� B Where b.���� = v_Jsklb And Rownum < 2;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  If n_�Һ�ģʽ = 1 And Nvl(d_Date, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) Then
    --������Ű�ģʽ
    --��ȡͣ�ﰲ��
    For r_ͣ�� In (Select a.Id As ��¼id, b.����, a.ҽ��id, a.ҽ������, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��
                 From �ٴ������¼ A, �ٴ������Դ B, �ٴ�����ͣ���¼ C
                 Where a.Id = c.��¼id And a.��Դid = b.Id And a.ͣ�￪ʼʱ�� Is Not Null And c.����ʱ�� Between d_Date And
                       d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || r_ͣ��.���� || '</HM><YSID>' || r_ͣ��.ҽ��id || '</YSID><YS>' || r_ͣ��.ҽ������ ||
                '</YS><KSSJ>' || r_ͣ��.ͣ�￪ʼʱ�� || '</KSSJ><JSSJ>' || r_ͣ��.ͣ����ֹʱ�� || '</JSSJ><BRLIST>';
      For r_ͣ�ﲡ�� In (Select a.��¼����, a.No, a.����id, To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��,
                            To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, b.����, d.����, c.ҽ������ As ҽ������
                     From ���˹Һż�¼ A, ���ű� B, �ٴ������¼ C, �ٴ������Դ D
                     Where a.ִ�в���id = b.Id And a.�����¼id = c.Id And c.��Դid = d.Id And ��¼״̬ = 1 And
                           ����ʱ�� Between r_ͣ��.ͣ�￪ʼʱ�� And r_ͣ��.ͣ����ֹʱ�� And a.�����¼id = r_ͣ��.��¼id And Not Exists
                      (Select 1 From ����䶯��¼ Where �Һŵ� = a.No)) Loop
        --ͣ�ﲡ���б��������Ѿ������ȡ���˵Ĳ���
        If r_ͣ�ﲡ��.��¼���� = 2 Then
          v_Brinfo := '<INFO><YYNO>' || r_ͣ�ﲡ��.No || '</YYNO><BRID>' || r_ͣ�ﲡ��.����id || '</BRID><YYSJ>' || r_ͣ�ﲡ��.����ʱ�� ||
                      '</YYSJ><CZSJ>' || r_ͣ�ﲡ��.�Ǽ�ʱ�� || '</CZSJ>' || '<YYKS>' || r_ͣ�ﲡ��.���� || '</YYKS><GHLX>' ||
                      r_ͣ�ﲡ��.���� || '</GHLX><YSXM>' || r_ͣ�ﲡ��.ҽ������ || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        Else
          Begin
            Select 1
            Into n_Exists
            From ����Ԥ����¼
            Where NO = r_ͣ�ﲡ��.No And ��¼���� = 4 And �����id = n_�����id;
          Exception
            When Others Then
              n_Exists := 0;
          End;
          If n_Exists = 1 Then
            v_Brinfo := '<INFO><YYNO>' || r_ͣ�ﲡ��.No || '</YYNO><BRID>' || r_ͣ�ﲡ��.����id || '</BRID><YYSJ>' || r_ͣ�ﲡ��.����ʱ�� ||
                        '</YYSJ><CZSJ>' || r_ͣ�ﲡ��.�Ǽ�ʱ�� || '</CZSJ>' || '<YYKS>' || r_ͣ�ﲡ��.���� || '</YYKS><GHLX>' ||
                        r_ͣ�ﲡ��.���� || '</GHLX><YSXM>' || r_ͣ�ﲡ��.ҽ������ || '</YSXM></INFO>';
            v_Temp   := v_Temp || v_Brinfo;
          End If;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    --��ȡ�����б�
    v_Temp := '';
    For r_���� In (Select d.��¼����, d.No, a.����id, To_Char(d.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ԤԼʱ��,
                        To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, a.ԭ����, a.ԭҽ������, b.רҵ����ְ�� As ԭְ��, a.�ֺ���, a.��ҽ������,
                        c.רҵ����ְ�� As ��ְ��
                 From ����䶯��¼ A, ��Ա�� B, ��Ա�� C, ���˹Һż�¼ D
                 Where a.�Ǽ�ʱ�� Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.ԭҽ��id = b.Id And a.��ҽ��id = c.Id And
                       a.�Һŵ� = d.No) Loop
      --ֻ���ظÿ����ҺŵĲ���         
      If r_����.��¼���� = 2 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || r_����.����id || '</BRID><YYSJ>' || r_����.�Ǽ�ʱ�� || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || r_����.ԤԼʱ�� || '</YSJ><YHM>' || r_����.ԭ���� || '</YHM><YYS>' || r_����.ԭҽ������ ||
                  '</YYS><YZC>' || r_����.ԭְ�� || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || r_����.ԤԼʱ�� || '</XSJ><XHM>' || r_����.�ֺ��� || '</XHM><XYS>' || r_����.��ҽ������ ||
                  '</XYS><XZC>' || r_����.��ְ�� || '</XZC></ITEM>';
      Else
        Begin
          Select 1 Into n_Exists From ����Ԥ����¼ Where NO = r_����.No And ��¼���� = 4 And �����id = n_�����id;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists = 1 Then
          v_Temp := v_Temp || '<ITEM><BRID>' || r_����.����id || '</BRID><YYSJ>' || r_����.�Ǽ�ʱ�� || '</YYSJ>';
          v_Temp := v_Temp || '<YSJ>' || r_����.ԤԼʱ�� || '</YSJ><YHM>' || r_����.ԭ���� || '</YHM><YYS>' || r_����.ԭҽ������ ||
                    '</YYS><YZC>' || r_����.ԭְ�� || '</YZC>';
          v_Temp := v_Temp || '<XSJ>' || r_����.ԤԼʱ�� || '</XSJ><XHM>' || r_����.�ֺ��� || '</XHM><XYS>' || r_����.��ҽ������ ||
                    '</XYS><XZC>' || r_����.��ְ�� || '</XZC></ITEM>';
        End If;
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --�ƻ��Ű�ģʽ
    --��ȡͣ�ﰲ��
    For Rs In (Select b.����, b.ҽ��id, b.ҽ������, To_Char(a.��ʼֹͣʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ʼֹͣʱ��,
                      To_Char(a.����ֹͣʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ֹͣʱ��
               From �ҺŰ���ͣ��״̬ A, �ҺŰ��� B
               Where a.����id = b.Id And a.�ƶ����� Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || Rs.���� || '</HM><YSID>' || Rs.ҽ��id || '</YSID><YS>' || Rs.ҽ������ ||
                '</YS><KSSJ>' || Rs.��ʼֹͣʱ�� || '</KSSJ><JSSJ>' || Rs.����ֹͣʱ�� || '</JSSJ><BRLIST>';
      ----2015/7/28
      For Rs_Br In (Select a.No, a.����id, To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��,
                           To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, b.����, c.����, a.ִ���� As ҽ������
                    From ���˹Һż�¼ A, ���ű� B, �ҺŰ��� C
                    Where a.�ű� = Rs.���� And a.ִ��״̬ = 0 And a.ִ�в���id = b.Id And b.Id = c.����id And a.�ű� = c.���� And
                          Trunc(����ʱ��) Between Trunc(To_Date(Rs.��ʼֹͣʱ��, 'yyyy-mm-dd hh24:mi:ss')) And
                          Trunc(To_Date(Rs.����ֹͣʱ��, 'yyyy-mm-dd hh24:mi:ss'))) Loop
        --ֻ���ظÿ����ҺŵĲ���
        Select Count(*)
        Into n_Cnt
        From (Select 1
               From ����Ԥ����¼ A
               Where a.No = Rs_Br.No And a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = Rs_Br.����id And �����id = n_�����id
               Union All
               Select 1 From ���˹Һż�¼ Where NO = Rs_Br.No And ��¼״̬ = 1 And ����˵�� = v_Jsklb);
        If n_Cnt > 0 Then
          v_Brinfo := '<INFO><YYNO>' || Rs_Br.No || '</YYNO><BRID>' || Rs_Br.����id || '</BRID><YYSJ>' || Rs_Br.����ʱ�� ||
                      '</YYSJ><CZSJ>' || Rs_Br.�Ǽ�ʱ�� || '</CZSJ>' || '<YYKS>' || Rs_Br.���� || '</YYKS><GHLX>' || Rs_Br.���� ||
                      '</GHLX><YSXM>' || Rs_Br.ҽ������ || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    --��ȡ�����¼
    v_Temp := '';
    For Rs In (Select d.No, a.����id, To_Char(d.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ԤԼʱ��,
                      To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, a.ԭ����, a.ԭҽ������, b.רҵ����ְ�� As ԭְ��, a.�ֺ���, a.��ҽ������,
                      c.רҵ����ְ�� As ��ְ��
               From ����䶯��¼ A, ��Ա�� B, ��Ա�� C, ���˹Һż�¼ D
               Where a.�Ǽ�ʱ�� Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.ԭҽ��id = b.Id And a.��ҽ��id = c.Id And
                     a.�Һŵ� = d.No) Loop
      --ֻ���ظÿ����ҺŵĲ���         
      Select Count(*)
      Into n_Cnt
      From (Select 1
             From ����Ԥ����¼ A
             Where a.No = Rs.No And a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = Rs.����id And �����id = n_�����id
             Union All
             Select 1 From ���˹Һż�¼ Where NO = Rs.No And ��¼״̬ = 1 And ����˵�� = v_Jsklb);
      If n_Cnt > 0 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || Rs.����id || '</BRID><YYSJ>' || Rs.�Ǽ�ʱ�� || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || Rs.ԤԼʱ�� || '</YSJ><YHM>' || Rs.ԭ���� || '</YHM><YYS>' || Rs.ԭҽ������ || '</YYS><YZC>' ||
                  Rs.ԭְ�� || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || Rs.ԤԼʱ�� || '</XSJ><XHM>' || Rs.�ֺ��� || '</XHM><XYS>' || Rs.��ҽ������ || '</XYS><XZC>' ||
                  Rs.��ְ�� || '</XZC></ITEM>';
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getregistalter;
/

--126662:����,2018-11-23,��ʷ�޸İ�֮ǰ�Ĺ��̸�����
Create Or Replace Procedure Zl_������Ŀ_Insert
(
  ���_In             In ������ĿĿ¼.���%Type := Null,
  ����id_In           In ������ĿĿ¼.����id%Type := Null,
  Id_In               In ������ĿĿ¼.Id%Type,
  ����_In             In ������ĿĿ¼.����%Type := Null,
  ����_In             In ������ĿĿ¼.����%Type := Null,
  ����ƴ��_In         In ������Ŀ����.����%Type := Null,
  �������_In         In ������Ŀ����.����%Type := Null,
  ����_In             ������ĿĿ¼.����%Type := Null,
  ����ƴ��_In         ������Ŀ����.����%Type := Null,
  �������_In         ������Ŀ����.����%Type := Null,
  ��������_In         In ������ĿĿ¼.��������%Type := Null,
  ִ��Ƶ��_In         In ������ĿĿ¼.ִ��Ƶ��%Type := Null,
  ����Ӧ��_In         In ������ĿĿ¼.����Ӧ��%Type := Null,
  ���㷽ʽ_In         In ������ĿĿ¼.���㷽ʽ%Type := Null,
  ���㵥λ_In         In ������ĿĿ¼.���㵥λ%Type := Null,
  �����Ա�_In         In ������ĿĿ¼.�����Ա�%Type := Null,
  ִ�а���_In         In ������ĿĿ¼.ִ�а���%Type := Null,
  �������_In         In ������ĿĿ¼.�������%Type := Null,
  �����Ŀ_In         In ������ĿĿ¼.�����Ŀ%Type := Null,
  �걾��λ_In         In ������ĿĿ¼.�걾��λ%Type := Null,
  ��������id_In       In ������϶���.����id%Type := Null,
  ִ�п���_In         In ������ĿĿ¼.ִ�п���%Type := Null,
  ����ִ��_In         In ����ִ�п���.ִ�п���id%Type := Null,
  סԺִ��_In         In ����ִ�п���.ִ�п���id%Type := Null,
  ����ִ��_In         In Varchar2, --�������Ҷ���ִ�е�˵��������'|'�ָÿ������'��������id^ִ�п���id'��ʽ��֯
  �ο�Ŀ¼id_In       In ������ĿĿ¼.�ο�Ŀ¼id%Type := Null,
  Ӧ�÷�Χ_In         In Number := 0,
  ¼������_In         In ������ĿĿ¼.¼������%Type := Null,
  ������Χ_In         In Number := 0,
  ִ�б��_In         In Number := 0,
  ִ�з���_In         In ������ĿĿ¼.ִ�з���%Type := 0,
  վ��_In             In ������ĿĿ¼.վ��%Type := Null,
  ��ĿƵ��_In         In Varchar2 := Null, --����Ŀ��Ƶ�����ô�������|����......
  �������_In         In ������ĿĿ¼.�������%Type := Null,
  ʹ�ÿ���_In         In Varchar2 := Null, --ʹ�ÿ��ҵ�IDs,�ö��ŷָ�
  ʹ�ÿ���Ӧ�÷�Χ_In In Number := 0, --ʹ�ÿ���Ӧ�õķ�Χ  0-���1-Ӧ����ͬ����2-���������У�3-Ӧ���ڵ�ǰ���
  First_In            In Number := 1, --First��1-��Ҫɾ��ִ�п��ң���������0-��ɾ��ִ�п��ң�ֱ������
  ����ϵ��_In         In ������ĿĿ¼.����ϵ��%Type := Null,
  ��Ѫ�������_In     In Varchar2 :=Null,
  ԭʼid_IN           In ������ĿĿ¼.Id%Type:=0,
  �Թܱ���_In         In ������ĿĿ¼.�Թܱ���%Type := Null  
) Is
  Type t_������Ŀ Is Ref Cursor;
  c_������Ŀ   t_������Ŀ;
  t_Id         t_Numlist;
  v_Id         ������ĿĿ¼.Id%Type;
  v_Records    Varchar2(4000); --��ʱ��¼�������Ҷ���ִ�п��ҵ��ַ���
  v_Currrec    Varchar2(1000); --�����ڶ���ִ�п����ַ����е�һ������
  v_Fields     Varchar2(1000);
  v_��������id ����ִ�п���.��������id%Type := Null;
  v_ִ�п���id ����ִ�п���.ִ�п���id%Type := Null;
  n_���       Number;
  v_���       Varchar2(1000);
  v_Strtmp     Varchar2(1000);
  v_Strinput   Varchar2(1000);
Begin
  If First_In = 1 Then
    Insert Into ������ĿĿ¼
      (���, ����id, ID, ����, ����, ��������, ִ��Ƶ��, ����Ӧ��, ���㷽ʽ, ���㵥λ, �����Ա�, ִ�а���, �������, ִ�п���, �����Ŀ, �걾��λ, ����ʱ��, ����ʱ��, �ο�Ŀ¼id, ¼������,
       ִ�б��, ִ�з���, �������, վ��, ����ϵ��,�Թܱ���)
    Values
      (���_In, ����id_In, Id_In, ����_In, ����_In, ��������_In, ִ��Ƶ��_In, ����Ӧ��_In, ���㷽ʽ_In, ���㵥λ_In, �����Ա�_In, ִ�а���_In, �������_In,
       ִ�п���_In, �����Ŀ_In, Decode(���_In, 'D', Decode(�����Ŀ_In, 1, '', �걾��λ_In), �걾��λ_In), Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), �ο�Ŀ¼id_In, ¼������_In, ִ�б��_In, ִ�з���_In, �������_In, վ��_In, ����ϵ��_In,�Թܱ���_In);
    If ��������id_In Is Not Null Then
      Insert Into ������϶��� (����id, ���id, ����id) Values (��������id_In, Null, Id_In);
    End If;
    If ����ƴ��_In Is Not Null Then
      Insert Into ������Ŀ���� (������Ŀid, ����, ����, ����, ����) Values (Id_In, ����_In, 1, ����ƴ��_In, 1);
    End If;
    If �������_In Is Not Null Then
      Insert Into ������Ŀ���� (������Ŀid, ����, ����, ����, ����) Values (Id_In, ����_In, 1, �������_In, 2);
    End If;
    If ����_In Is Not Null And ����ƴ��_In Is Not Null Then
      Insert Into ������Ŀ���� (������Ŀid, ����, ����, ����, ����) Values (Id_In, ����_In, 9, ����ƴ��_In, 1);
    End If;
    If ����_In Is Not Null And �������_In Is Not Null Then
      Insert Into ������Ŀ���� (������Ŀid, ����, ����, ����, ����) Values (Id_In, ����_In, 9, �������_In, 2);
    End If;
  End If;
  If Ӧ�÷�Χ_In = 1 Then
    If ����id_In Is Null Then
      Open c_������Ŀ For
        Select ID From ������ĿĿ¼ Where ����id Is Null Order By ����;
    Else
      Open c_������Ŀ For
        Select ID From ������ĿĿ¼ Where ����id = ����id_In Order By ����;
    End If;
  Elsif Ӧ�÷�Χ_In = 2 Then
    If ����id_In Is Null Then
      Open c_������Ŀ For
        Select c.Id
        From ������ĿĿ¼ C, (Select ID From ���Ʒ���Ŀ¼ Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id) D
        Where d.Id = c.����id
        Order By ����;
    Else
      Open c_������Ŀ For
        Select c.Id
        From ������ĿĿ¼ C, (Select ID From ���Ʒ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id) D
        Where d.Id = c.����id
        Order By ����;
    End If;
  Elsif Ӧ�÷�Χ_In = 3 Then
    Open c_������Ŀ For
      Select ID From ������ĿĿ¼ Where ��� = ���_In Order By ����;
  Else
    Open c_������Ŀ For
      Select ID From ������ĿĿ¼ Where ID = Id_In;
  End If;

  Loop
    Fetch c_������Ŀ
      Into v_Id;
    Exit When c_������Ŀ%NotFound;
  
    If First_In = 1 Then
      Delete From ����ִ�п��� Where ������Ŀid = v_Id;
      If ִ�п���_In = 4 And ����ִ��_In Is Not Null Then
        Insert Into ����ִ�п��� (������Ŀid, ������Դ, ��������id, ִ�п���id) Values (v_Id, 1, Null, ����ִ��_In);
      End If;
      If ִ�п���_In = 4 And סԺִ��_In Is Not Null Then
        Insert Into ����ִ�п��� (������Ŀid, ������Դ, ��������id, ִ�п���id) Values (v_Id, 2, Null, סԺִ��_In);
      End If;
    End If;
    If ִ�п���_In <> 4 Or ����ִ��_In Is Null Then
      v_Records := Null;
    Else
      v_Records := ����ִ��_In || '|';
    End If;
  
    While v_Records Is Not Null Loop
      v_Currrec    := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields     := v_Currrec;
      v_��������id := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_ִ�п���id := To_Number(v_Fields);
      Insert Into ����ִ�п���
        (������Ŀid, ������Դ, ��������id, ִ�п���id)
      Values
        (v_Id, Null, Decode(v_��������id, 0, Null, v_��������id), v_ִ�п���id);
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
    If Ӧ�÷�Χ_In <> 0 Then
      Update ������ĿĿ¼ Set ִ�п��� = ִ�п���_In Where ID = v_Id;
    End If;
  End Loop;
  Close c_������Ŀ;

  If First_In = 1 Then
    If ���_In = 'C' Or ���_In = 'F' Or ���_In = 'K' Then
      Insert Into ��������Ӧ��
        (�����ļ�id, Ӧ�ó���, ������Ŀid)
        Select a.�����ļ�id, 1, Id_In
        From ��������Ӧ�� A, ������ĿĿ¼ I
        Where a.������Ŀid = i.Id And i.��� = ���_In And Ӧ�ó��� = 1 And (�������_In = 0 Or �������_In = 1) And Rownum < 2;
      Insert Into ��������Ӧ��
        (�����ļ�id, Ӧ�ó���, ������Ŀid)
        Select a.�����ļ�id, 2, Id_In
        From ��������Ӧ�� A, ������ĿĿ¼ I
        Where a.������Ŀid = i.Id And i.��� = ���_In And Ӧ�ó��� = 2 And (�������_In = 0 Or �������_In = 2) And Rownum < 2;
    Elsif ���_In = 'D' Or ���_In = 'E' Then
      Insert Into ��������Ӧ��
        (�����ļ�id, Ӧ�ó���, ������Ŀid)
        Select a.�����ļ�id, 1, Id_In
        From ��������Ӧ�� A, ������ĿĿ¼ I
        Where a.������Ŀid = i.Id And i.��� = ���_In And �������� = ��������_In And Ӧ�ó��� = 1 And (�������_In = 0 Or �������_In = 1) And
              Rownum < 2;
      Insert Into ��������Ӧ��
        (�����ļ�id, Ӧ�ó���, ������Ŀid)
        Select a.�����ļ�id, 2, Id_In
        From ��������Ӧ�� A, ������ĿĿ¼ I
        Where a.������Ŀid = i.Id And i.��� = ���_In And �������� = ��������_In And Ӧ�ó��� = 2 And (�������_In = 0 Or �������_In = 2) And
              Rownum < 2;
    End If;
  End If;

  If ������Χ_In = 1 Then
    If ����id_In Is Null Then
      Update ������ĿĿ¼ Set ¼������ = ¼������_In Where ����id Is Null;
    Else
      Update ������ĿĿ¼ Set ¼������ = ¼������_In Where ����id = ����id_In;
    End If;
  Elsif ������Χ_In = 2 Then
    If ����id_In Is Null Then
      Update ������ĿĿ¼
      Set ¼������ = ¼������_In
      Where ����id In (Select ID From ���Ʒ���Ŀ¼ Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id);
    Else
      Update ������ĿĿ¼
      Set ¼������ = ¼������_In
      Where ����id In (Select ID From ���Ʒ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id);
    End If;
  Elsif ������Χ_In = 3 Then
    Update ������ĿĿ¼ Set ¼������ = ¼������_In Where ��� = ���_In;
  Elsif ������Χ_In = 4 Then
    Update ������ĿĿ¼ Set ¼������ = ¼������_In;
  End If;

  --����Ŀ��Ƶ������
  If ���_In <> 'C' Then
    Delete �����÷����� Where ��Ŀid = Id_In;
    If ��ĿƵ��_In Is Not Null Then
      v_Strinput := ��ĿƵ��_In || '|';
      n_���     := 0;
    
      While v_Strinput Is Not Null Loop
        v_Strtmp := Substr(v_Strinput, 1, Instr(v_Strinput, '|') - 1);
        v_���   := v_Strtmp;
        n_���   := n_��� + 1;
      
        Insert Into �����÷����� (��Ŀid, ����, Ƶ��) Values (Id_In, n_���, v_���);
        v_Strinput := Replace('|' || v_Strinput, '|' || v_Strtmp || '|');
      End Loop;
    End If;
  End If;
  --ʹ�ÿ���
  If ʹ�ÿ���Ӧ�÷�Χ_In = 1 Then
    If ����id_In Is Null Then
      Open c_������Ŀ For
        Select ID From ������ĿĿ¼ Where ����id Is Null Order By ����;
    Else
      Open c_������Ŀ For
        Select ID From ������ĿĿ¼ Where ����id = ����id_In Order By ����;
    End If;
  Elsif ʹ�ÿ���Ӧ�÷�Χ_In = 2 Then
    If ����id_In Is Null Then
      Open c_������Ŀ For
        Select c.Id
        From ������ĿĿ¼ C, (Select ID From ���Ʒ���Ŀ¼ Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id) D
        Where d.Id = c.����id
        Order By ����;
    Else
      Open c_������Ŀ For
        Select c.Id
        From ������ĿĿ¼ C, (Select ID From ���Ʒ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id) D
        Where d.Id = c.����id
        Order By ����;
    End If;
  Elsif ʹ�ÿ���Ӧ�÷�Χ_In = 3 Then
    Open c_������Ŀ For
      Select ID From ������ĿĿ¼ Where ��� = ���_In Order By ����;
  Else
    Open c_������Ŀ For
      Select ID From ������ĿĿ¼ Where ID = Id_In;
  End If;
  Fetch c_������Ŀ Bulk Collect
    Into t_Id;
  Close c_������Ŀ;

  Forall I In 1 .. t_Id.Count
    Delete �������ÿ��� Where ��Ŀid = t_Id(I) And Instr(',' || ʹ�ÿ���_In || ',', ',' || ����id || ',') = 0;

  If ʹ�ÿ���_In Is Not Null Then
    Forall I In 1 .. t_Id.Count
      Insert Into �������ÿ���
        (��Ŀid, ����id)
        Select t_Id(I), Column_Value
        From Table(f_Num2list(ʹ�ÿ���_In)) A
        Where Not Exists (Select 1 From �������ÿ��� Where ����id = Column_Value And ��Ŀid = t_Id(I));
  End If;
  --��Ѫ�������
  If ���_In = 'K' And ��Ѫ�������_In Is Not Null Then
    v_Strinput := ��Ѫ�������_In || '|';
  
    While v_Strinput Is Not Null Loop
      v_Strtmp := Substr(v_Strinput, 1, Instr(v_Strinput, '|') - 1);
      v_Id     := v_Strtmp;
    
      Insert Into ��Ѫ������� (��Ŀid, ������Ŀid) Values (Id_In, v_Id);
      v_Strinput := Replace('|' || v_Strinput, '|' || v_Strtmp || '|');
    End Loop;
  End If;
  
  if ԭʼid_IN<>0 then
    Zl_�����շ�_Insert(id_In,ԭʼid_IN);
  end if;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������Ŀ_Insert;
/

------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.150.0018' Where ���=&n_System;
Commit;
