----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.34.150升级到 v10.34.150
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------


------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--134222:焦博,2018-11-23,调整Oracle过程Zl_Third_Getregistalter,分段创建返回的XML字符串
Create Or Replace Procedure Zl_Third_Getregistalter
(
  Xml_In  Xmltype,
  Xml_Out Out Xmltype
) Is
  -----------------------------------------------
  --功能：获取当天操作的停换诊安排
  --入参：XML_IN
  --<IN>
  --  <JSKLB>结算卡类别</JSKLB>
  --  <RQ>日期</RQ>
  --</IN>
  --出参:XML_OUT
  --<OUTPUT>
  --  <TZLISTS>          //停诊列表
  --    <ITEM>
  --      <HM>号码</HM>
  --      <YSID>医生ID</YSID>
  --      <YS>医生姓名</YS>
  --      <KSSJ>停诊开始时间</KSSJ>
  --      <JSSJ>停诊结束时间</JSSJ>
  --      <BRLIST>
  --        <INFO>
  --          <YYNO>预约单据号</YYNO>
  --          <BRID>病人ID</BRID>
  --          <YYSJ>预约时间</YYSJ>
  --          <CZSJ>操作时间</CZSJ>
  --          <YYKS>预约科室</YYKS>
  --          <GHLX>号类</GHLX>
  --          <YSXM>医生姓名</YSXM>
  --        </INFO>
  --      </BRLIST>
  --    </ITEM>
  --  </TZLISTS>
  --  <HZLISTS>          //换诊列表
  --    <ITEM>
  --      <BRID>病人ID</BRID>
  --      <YYSJ>预约的操作时间</YYSJ>
  --      <YSJ>原预约时间</YSJ>
  --      <YHM>原号码</YHM>
  --      <YYS>原医生</YYS>
  --      <YZC>原医生的职称</YZC>
  --      <XSJ>现预约时间</XSJ>
  --      <XHM>现号码</XHM>
  --      <XYS>现医生</XYS>
  --      <XZC>现医生的职称</XZC>
  --    </ITEM>
  --  </HZLIST>
  --</OUTPUT>
  -----------------------------------------------------

  d_Date     Date;
  v_Jsklb    Varchar2(100);
  n_卡类别id 医疗卡类别.Id%Type;
  n_Cnt      Number(3);
  v_Temp     Clob;
  v_Brinfo   Varchar2(4000);
  d_启用时间 Date;
  v_Para     Varchar2(2000);
  n_Exists   Number(3);
  n_挂号模式 Number(3);
  x_Templet  Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/JSKLB') Into v_Jsklb From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd')
  Into d_Date
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select b.Id Into n_卡类别id From 医疗卡类别 B Where b.名称 = v_Jsklb And Rownum < 2;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If n_挂号模式 = 1 And Nvl(d_Date, Sysdate) > Nvl(d_启用时间, Sysdate - 30) Then
    --出诊表排班模式
    --获取停诊安排
    For r_停诊 In (Select a.Id As 记录id, b.号码, a.医生id, a.医生姓名, a.停诊开始时间, a.停诊终止时间
                 From 临床出诊记录 A, 临床出诊号源 B, 临床出诊停诊记录 C
                 Where a.Id = c.记录id And a.号源id = b.Id And a.停诊开始时间 Is Not Null And c.审批时间 Between d_Date And
                       d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || r_停诊.号码 || '</HM><YSID>' || r_停诊.医生id || '</YSID><YS>' || r_停诊.医生姓名 ||
                '</YS><KSSJ>' || r_停诊.停诊开始时间 || '</KSSJ><JSSJ>' || r_停诊.停诊终止时间 || '</JSSJ><BRLIST>';
      For r_停诊病人 In (Select a.记录性质, a.No, a.病人id, To_Char(a.发生时间, 'yyyy-mm-dd') As 发生时间,
                            To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.名称, d.号类, c.医生姓名 As 医生姓名
                     From 病人挂号记录 A, 部门表 B, 临床出诊记录 C, 临床出诊号源 D
                     Where a.执行部门id = b.Id And a.出诊记录id = c.Id And c.号源id = d.Id And 记录状态 = 1 And
                           发生时间 Between r_停诊.停诊开始时间 And r_停诊.停诊终止时间 And a.出诊记录id = r_停诊.记录id And Not Exists
                      (Select 1 From 就诊变动记录 Where 挂号单 = a.No)) Loop
        --停诊病人列表，不包含已经换诊和取消了的病人
        If r_停诊病人.记录性质 = 2 Then
          v_Brinfo := '<INFO><YYNO>' || r_停诊病人.No || '</YYNO><BRID>' || r_停诊病人.病人id || '</BRID><YYSJ>' || r_停诊病人.发生时间 ||
                      '</YYSJ><CZSJ>' || r_停诊病人.登记时间 || '</CZSJ>' || '<YYKS>' || r_停诊病人.名称 || '</YYKS><GHLX>' ||
                      r_停诊病人.号类 || '</GHLX><YSXM>' || r_停诊病人.医生姓名 || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        Else
          Begin
            Select 1
            Into n_Exists
            From 病人预交记录
            Where NO = r_停诊病人.No And 记录性质 = 4 And 卡类别id = n_卡类别id;
          Exception
            When Others Then
              n_Exists := 0;
          End;
          If n_Exists = 1 Then
            v_Brinfo := '<INFO><YYNO>' || r_停诊病人.No || '</YYNO><BRID>' || r_停诊病人.病人id || '</BRID><YYSJ>' || r_停诊病人.发生时间 ||
                        '</YYSJ><CZSJ>' || r_停诊病人.登记时间 || '</CZSJ>' || '<YYKS>' || r_停诊病人.名称 || '</YYKS><GHLX>' ||
                        r_停诊病人.号类 || '</GHLX><YSXM>' || r_停诊病人.医生姓名 || '</YSXM></INFO>';
            v_Temp   := v_Temp || v_Brinfo;
          End If;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    --换取换诊列表
    v_Temp := '';
    For r_换诊 In (Select d.记录性质, d.No, a.病人id, To_Char(d.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                        To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.原号码, a.原医生姓名, b.专业技术职务 As 原职务, a.现号码, a.现医生姓名,
                        c.专业技术职务 As 现职务
                 From 就诊变动记录 A, 人员表 B, 人员表 C, 病人挂号记录 D
                 Where a.登记时间 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.原医生id = b.Id And a.现医生id = c.Id And
                       a.挂号单 = d.No) Loop
      --只返回该卡类别挂号的病人         
      If r_换诊.记录性质 = 2 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || r_换诊.病人id || '</BRID><YYSJ>' || r_换诊.登记时间 || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || r_换诊.预约时间 || '</YSJ><YHM>' || r_换诊.原号码 || '</YHM><YYS>' || r_换诊.原医生姓名 ||
                  '</YYS><YZC>' || r_换诊.原职务 || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || r_换诊.预约时间 || '</XSJ><XHM>' || r_换诊.现号码 || '</XHM><XYS>' || r_换诊.现医生姓名 ||
                  '</XYS><XZC>' || r_换诊.现职务 || '</XZC></ITEM>';
      Else
        Begin
          Select 1 Into n_Exists From 病人预交记录 Where NO = r_换诊.No And 记录性质 = 4 And 卡类别id = n_卡类别id;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists = 1 Then
          v_Temp := v_Temp || '<ITEM><BRID>' || r_换诊.病人id || '</BRID><YYSJ>' || r_换诊.登记时间 || '</YYSJ>';
          v_Temp := v_Temp || '<YSJ>' || r_换诊.预约时间 || '</YSJ><YHM>' || r_换诊.原号码 || '</YHM><YYS>' || r_换诊.原医生姓名 ||
                    '</YYS><YZC>' || r_换诊.原职务 || '</YZC>';
          v_Temp := v_Temp || '<XSJ>' || r_换诊.预约时间 || '</XSJ><XHM>' || r_换诊.现号码 || '</XHM><XYS>' || r_换诊.现医生姓名 ||
                    '</XYS><XZC>' || r_换诊.现职务 || '</XZC></ITEM>';
        End If;
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --计划排班模式
    --获取停诊安排
    For Rs In (Select b.号码, b.医生id, b.医生姓名, To_Char(a.开始停止时间, 'yyyy-mm-dd hh24:mi:ss') As 开始停止时间,
                      To_Char(a.结束停止时间, 'yyyy-mm-dd hh24:mi:ss') As 结束停止时间
               From 挂号安排停用状态 A, 挂号安排 B
               Where a.安排id = b.Id And a.制订日期 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || Rs.号码 || '</HM><YSID>' || Rs.医生id || '</YSID><YS>' || Rs.医生姓名 ||
                '</YS><KSSJ>' || Rs.开始停止时间 || '</KSSJ><JSSJ>' || Rs.结束停止时间 || '</JSSJ><BRLIST>';
      ----2015/7/28
      For Rs_Br In (Select a.No, a.病人id, To_Char(a.发生时间, 'yyyy-mm-dd') As 发生时间,
                           To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.名称, c.号类, a.执行人 As 医生姓名
                    From 病人挂号记录 A, 部门表 B, 挂号安排 C
                    Where a.号别 = Rs.号码 And a.执行状态 = 0 And a.执行部门id = b.Id And b.Id = c.科室id And a.号别 = c.号码 And
                          Trunc(发生时间) Between Trunc(To_Date(Rs.开始停止时间, 'yyyy-mm-dd hh24:mi:ss')) And
                          Trunc(To_Date(Rs.结束停止时间, 'yyyy-mm-dd hh24:mi:ss'))) Loop
        --只返回该卡类别挂号的病人
        Select Count(*)
        Into n_Cnt
        From (Select 1
               From 病人预交记录 A
               Where a.No = Rs_Br.No And a.记录性质 = 4 And a.记录状态 = 1 And a.病人id = Rs_Br.病人id And 卡类别id = n_卡类别id
               Union All
               Select 1 From 病人挂号记录 Where NO = Rs_Br.No And 记录状态 = 1 And 交易说明 = v_Jsklb);
        If n_Cnt > 0 Then
          v_Brinfo := '<INFO><YYNO>' || Rs_Br.No || '</YYNO><BRID>' || Rs_Br.病人id || '</BRID><YYSJ>' || Rs_Br.发生时间 ||
                      '</YYSJ><CZSJ>' || Rs_Br.登记时间 || '</CZSJ>' || '<YYKS>' || Rs_Br.名称 || '</YYKS><GHLX>' || Rs_Br.号类 ||
                      '</GHLX><YSXM>' || Rs_Br.医生姓名 || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    --获取换诊记录
    v_Temp := '';
    For Rs In (Select d.No, a.病人id, To_Char(d.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                      To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.原号码, a.原医生姓名, b.专业技术职务 As 原职务, a.现号码, a.现医生姓名,
                      c.专业技术职务 As 现职务
               From 就诊变动记录 A, 人员表 B, 人员表 C, 病人挂号记录 D
               Where a.登记时间 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.原医生id = b.Id And a.现医生id = c.Id And
                     a.挂号单 = d.No) Loop
      --只返回该卡类别挂号的病人         
      Select Count(*)
      Into n_Cnt
      From (Select 1
             From 病人预交记录 A
             Where a.No = Rs.No And a.记录性质 = 4 And a.记录状态 = 1 And a.病人id = Rs.病人id And 卡类别id = n_卡类别id
             Union All
             Select 1 From 病人挂号记录 Where NO = Rs.No And 记录状态 = 1 And 交易说明 = v_Jsklb);
      If n_Cnt > 0 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || Rs.病人id || '</BRID><YYSJ>' || Rs.登记时间 || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || Rs.预约时间 || '</YSJ><YHM>' || Rs.原号码 || '</YHM><YYS>' || Rs.原医生姓名 || '</YYS><YZC>' ||
                  Rs.原职务 || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || Rs.预约时间 || '</XSJ><XHM>' || Rs.现号码 || '</XHM><XYS>' || Rs.现医生姓名 || '</XYS><XZC>' ||
                  Rs.现职务 || '</XZC></ITEM>';
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

--126662:蒋敏,2018-11-23,历史修改把之前的过程覆盖了
Create Or Replace Procedure Zl_诊疗项目_Insert
(
  类别_In             In 诊疗项目目录.类别%Type := Null,
  分类id_In           In 诊疗项目目录.分类id%Type := Null,
  Id_In               In 诊疗项目目录.Id%Type,
  编码_In             In 诊疗项目目录.编码%Type := Null,
  名称_In             In 诊疗项目目录.名称%Type := Null,
  名称拼音_In         In 诊疗项目别名.简码%Type := Null,
  名称五笔_In         In 诊疗项目别名.简码%Type := Null,
  别名_In             诊疗项目目录.名称%Type := Null,
  别名拼音_In         诊疗项目别名.简码%Type := Null,
  别名五笔_In         诊疗项目别名.简码%Type := Null,
  操作类型_In         In 诊疗项目目录.操作类型%Type := Null,
  执行频率_In         In 诊疗项目目录.执行频率%Type := Null,
  单独应用_In         In 诊疗项目目录.单独应用%Type := Null,
  计算方式_In         In 诊疗项目目录.计算方式%Type := Null,
  计算单位_In         In 诊疗项目目录.计算单位%Type := Null,
  适用性别_In         In 诊疗项目目录.适用性别%Type := Null,
  执行安排_In         In 诊疗项目目录.执行安排%Type := Null,
  服务对象_In         In 诊疗项目目录.服务对象%Type := Null,
  组合项目_In         In 诊疗项目目录.组合项目%Type := Null,
  标本部位_In         In 诊疗项目目录.标本部位%Type := Null,
  手术操作id_In       In 疾病诊断对照.疾病id%Type := Null,
  执行科室_In         In 诊疗项目目录.执行科室%Type := Null,
  门诊执行_In         In 诊疗执行科室.执行科室id%Type := Null,
  住院执行_In         In 诊疗执行科室.执行科室id%Type := Null,
  定向执行_In         In Varchar2, --开单科室定向执行的说明串，以'|'分割，每个定向按'开单科室id^执行科室id'形式组织
  参考目录id_In       In 诊疗项目目录.参考目录id%Type := Null,
  应用范围_In         In Number := 0,
  录入限量_In         In 诊疗项目目录.录入限量%Type := Null,
  限量范围_In         In Number := 0,
  执行标记_In         In Number := 0,
  执行分类_In         In 诊疗项目目录.执行分类%Type := 0,
  站点_In             In 诊疗项目目录.站点%Type := Null,
  项目频率_In         In Varchar2 := Null, --该项目的频率设置串：编码|编码......
  计算规则_In         In 诊疗项目目录.计算规则%Type := Null,
  使用科室_In         In Varchar2 := Null, --使用科室的IDs,用逗号分隔
  使用科室应用范围_In In Number := 0, --使用科室应用的范围  0-本项，1-应用于同级，2-分类下所有，3-应用于当前类别
  First_In            In Number := 1, --First：1-需要删除执行科室，再新增，0-不删除执行科室，直接新增
  计算系数_In         In 诊疗项目目录.计算系数%Type := Null,
  输血检验对照_In     In Varchar2 :=Null,
  原始id_IN           In 诊疗项目目录.Id%Type:=0,
  试管编码_In         In 诊疗项目目录.试管编码%Type := Null  
) Is
  Type t_诊疗项目 Is Ref Cursor;
  c_诊疗项目   t_诊疗项目;
  t_Id         t_Numlist;
  v_Id         诊疗项目目录.Id%Type;
  v_Records    Varchar2(4000); --临时记录开单科室定向执行科室的字符串
  v_Currrec    Varchar2(1000); --包含在定向执行科室字符串中的一个定向
  v_Fields     Varchar2(1000);
  v_开单科室id 诊疗执行科室.开单科室id%Type := Null;
  v_执行科室id 诊疗执行科室.执行科室id%Type := Null;
  n_序号       Number;
  v_编号       Varchar2(1000);
  v_Strtmp     Varchar2(1000);
  v_Strinput   Varchar2(1000);
Begin
  If First_In = 1 Then
    Insert Into 诊疗项目目录
      (类别, 分类id, ID, 编码, 名称, 操作类型, 执行频率, 单独应用, 计算方式, 计算单位, 适用性别, 执行安排, 服务对象, 执行科室, 组合项目, 标本部位, 建档时间, 撤档时间, 参考目录id, 录入限量,
       执行标记, 执行分类, 计算规则, 站点, 计算系数,试管编码)
    Values
      (类别_In, 分类id_In, Id_In, 编码_In, 名称_In, 操作类型_In, 执行频率_In, 单独应用_In, 计算方式_In, 计算单位_In, 适用性别_In, 执行安排_In, 服务对象_In,
       执行科室_In, 组合项目_In, Decode(类别_In, 'D', Decode(组合项目_In, 1, '', 标本部位_In), 标本部位_In), Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), 参考目录id_In, 录入限量_In, 执行标记_In, 执行分类_In, 计算规则_In, 站点_In, 计算系数_In,试管编码_In);
    If 手术操作id_In Is Not Null Then
      Insert Into 疾病诊断对照 (疾病id, 诊断id, 手术id) Values (手术操作id_In, Null, Id_In);
    End If;
    If 名称拼音_In Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 名称_In, 1, 名称拼音_In, 1);
    End If;
    If 名称五笔_In Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 名称_In, 1, 名称五笔_In, 2);
    End If;
    If 别名_In Is Not Null And 别名拼音_In Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 别名_In, 9, 别名拼音_In, 1);
    End If;
    If 别名_In Is Not Null And 别名五笔_In Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 别名_In, 9, 别名五笔_In, 2);
    End If;
  End If;
  If 应用范围_In = 1 Then
    If 分类id_In Is Null Then
      Open c_诊疗项目 For
        Select ID From 诊疗项目目录 Where 分类id Is Null Order By 编码;
    Else
      Open c_诊疗项目 For
        Select ID From 诊疗项目目录 Where 分类id = 分类id_In Order By 编码;
    End If;
  Elsif 应用范围_In = 2 Then
    If 分类id_In Is Null Then
      Open c_诊疗项目 For
        Select c.Id
        From 诊疗项目目录 C, (Select ID From 诊疗分类目录 Start With 上级id Is Null Connect By Prior ID = 上级id) D
        Where d.Id = c.分类id
        Order By 编码;
    Else
      Open c_诊疗项目 For
        Select c.Id
        From 诊疗项目目录 C, (Select ID From 诊疗分类目录 Start With ID = 分类id_In Connect By Prior ID = 上级id) D
        Where d.Id = c.分类id
        Order By 编码;
    End If;
  Elsif 应用范围_In = 3 Then
    Open c_诊疗项目 For
      Select ID From 诊疗项目目录 Where 类别 = 类别_In Order By 编码;
  Else
    Open c_诊疗项目 For
      Select ID From 诊疗项目目录 Where ID = Id_In;
  End If;

  Loop
    Fetch c_诊疗项目
      Into v_Id;
    Exit When c_诊疗项目%NotFound;
  
    If First_In = 1 Then
      Delete From 诊疗执行科室 Where 诊疗项目id = v_Id;
      If 执行科室_In = 4 And 门诊执行_In Is Not Null Then
        Insert Into 诊疗执行科室 (诊疗项目id, 病人来源, 开单科室id, 执行科室id) Values (v_Id, 1, Null, 门诊执行_In);
      End If;
      If 执行科室_In = 4 And 住院执行_In Is Not Null Then
        Insert Into 诊疗执行科室 (诊疗项目id, 病人来源, 开单科室id, 执行科室id) Values (v_Id, 2, Null, 住院执行_In);
      End If;
    End If;
    If 执行科室_In <> 4 Or 定向执行_In Is Null Then
      v_Records := Null;
    Else
      v_Records := 定向执行_In || '|';
    End If;
  
    While v_Records Is Not Null Loop
      v_Currrec    := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields     := v_Currrec;
      v_开单科室id := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_执行科室id := To_Number(v_Fields);
      Insert Into 诊疗执行科室
        (诊疗项目id, 病人来源, 开单科室id, 执行科室id)
      Values
        (v_Id, Null, Decode(v_开单科室id, 0, Null, v_开单科室id), v_执行科室id);
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
    If 应用范围_In <> 0 Then
      Update 诊疗项目目录 Set 执行科室 = 执行科室_In Where ID = v_Id;
    End If;
  End Loop;
  Close c_诊疗项目;

  If First_In = 1 Then
    If 类别_In = 'C' Or 类别_In = 'F' Or 类别_In = 'K' Then
      Insert Into 病历单据应用
        (病历文件id, 应用场合, 诊疗项目id)
        Select a.病历文件id, 1, Id_In
        From 病历单据应用 A, 诊疗项目目录 I
        Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 应用场合 = 1 And (服务对象_In = 0 Or 服务对象_In = 1) And Rownum < 2;
      Insert Into 病历单据应用
        (病历文件id, 应用场合, 诊疗项目id)
        Select a.病历文件id, 2, Id_In
        From 病历单据应用 A, 诊疗项目目录 I
        Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 应用场合 = 2 And (服务对象_In = 0 Or 服务对象_In = 2) And Rownum < 2;
    Elsif 类别_In = 'D' Or 类别_In = 'E' Then
      Insert Into 病历单据应用
        (病历文件id, 应用场合, 诊疗项目id)
        Select a.病历文件id, 1, Id_In
        From 病历单据应用 A, 诊疗项目目录 I
        Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 操作类型 = 操作类型_In And 应用场合 = 1 And (服务对象_In = 0 Or 服务对象_In = 1) And
              Rownum < 2;
      Insert Into 病历单据应用
        (病历文件id, 应用场合, 诊疗项目id)
        Select a.病历文件id, 2, Id_In
        From 病历单据应用 A, 诊疗项目目录 I
        Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 操作类型 = 操作类型_In And 应用场合 = 2 And (服务对象_In = 0 Or 服务对象_In = 2) And
              Rownum < 2;
    End If;
  End If;

  If 限量范围_In = 1 Then
    If 分类id_In Is Null Then
      Update 诊疗项目目录 Set 录入限量 = 录入限量_In Where 分类id Is Null;
    Else
      Update 诊疗项目目录 Set 录入限量 = 录入限量_In Where 分类id = 分类id_In;
    End If;
  Elsif 限量范围_In = 2 Then
    If 分类id_In Is Null Then
      Update 诊疗项目目录
      Set 录入限量 = 录入限量_In
      Where 分类id In (Select ID From 诊疗分类目录 Start With 上级id Is Null Connect By Prior ID = 上级id);
    Else
      Update 诊疗项目目录
      Set 录入限量 = 录入限量_In
      Where 分类id In (Select ID From 诊疗分类目录 Start With ID = 分类id_In Connect By Prior ID = 上级id);
    End If;
  Elsif 限量范围_In = 3 Then
    Update 诊疗项目目录 Set 录入限量 = 录入限量_In Where 类别 = 类别_In;
  Elsif 限量范围_In = 4 Then
    Update 诊疗项目目录 Set 录入限量 = 录入限量_In;
  End If;

  --该项目的频率设置
  If 类别_In <> 'C' Then
    Delete 诊疗用法用量 Where 项目id = Id_In;
    If 项目频率_In Is Not Null Then
      v_Strinput := 项目频率_In || '|';
      n_序号     := 0;
    
      While v_Strinput Is Not Null Loop
        v_Strtmp := Substr(v_Strinput, 1, Instr(v_Strinput, '|') - 1);
        v_编号   := v_Strtmp;
        n_序号   := n_序号 + 1;
      
        Insert Into 诊疗用法用量 (项目id, 性质, 频次) Values (Id_In, n_序号, v_编号);
        v_Strinput := Replace('|' || v_Strinput, '|' || v_Strtmp || '|');
      End Loop;
    End If;
  End If;
  --使用科室
  If 使用科室应用范围_In = 1 Then
    If 分类id_In Is Null Then
      Open c_诊疗项目 For
        Select ID From 诊疗项目目录 Where 分类id Is Null Order By 编码;
    Else
      Open c_诊疗项目 For
        Select ID From 诊疗项目目录 Where 分类id = 分类id_In Order By 编码;
    End If;
  Elsif 使用科室应用范围_In = 2 Then
    If 分类id_In Is Null Then
      Open c_诊疗项目 For
        Select c.Id
        From 诊疗项目目录 C, (Select ID From 诊疗分类目录 Start With 上级id Is Null Connect By Prior ID = 上级id) D
        Where d.Id = c.分类id
        Order By 编码;
    Else
      Open c_诊疗项目 For
        Select c.Id
        From 诊疗项目目录 C, (Select ID From 诊疗分类目录 Start With ID = 分类id_In Connect By Prior ID = 上级id) D
        Where d.Id = c.分类id
        Order By 编码;
    End If;
  Elsif 使用科室应用范围_In = 3 Then
    Open c_诊疗项目 For
      Select ID From 诊疗项目目录 Where 类别 = 类别_In Order By 编码;
  Else
    Open c_诊疗项目 For
      Select ID From 诊疗项目目录 Where ID = Id_In;
  End If;
  Fetch c_诊疗项目 Bulk Collect
    Into t_Id;
  Close c_诊疗项目;

  Forall I In 1 .. t_Id.Count
    Delete 诊疗适用科室 Where 项目id = t_Id(I) And Instr(',' || 使用科室_In || ',', ',' || 科室id || ',') = 0;

  If 使用科室_In Is Not Null Then
    Forall I In 1 .. t_Id.Count
      Insert Into 诊疗适用科室
        (项目id, 科室id)
        Select t_Id(I), Column_Value
        From Table(f_Num2list(使用科室_In)) A
        Where Not Exists (Select 1 From 诊疗适用科室 Where 科室id = Column_Value And 项目id = t_Id(I));
  End If;
  --输血检验对照
  If 类别_In = 'K' And 输血检验对照_In Is Not Null Then
    v_Strinput := 输血检验对照_In || '|';
  
    While v_Strinput Is Not Null Loop
      v_Strtmp := Substr(v_Strinput, 1, Instr(v_Strinput, '|') - 1);
      v_Id     := v_Strtmp;
    
      Insert Into 输血检验对照 (项目id, 检验项目id) Values (Id_In, v_Id);
      v_Strinput := Replace('|' || v_Strinput, '|' || v_Strtmp || '|');
    End Loop;
  End If;
  
  if 原始id_IN<>0 then
    Zl_诊疗收费_Insert(id_In,原始id_IN);
  end if;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗项目_Insert;
/

------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.150.0018' Where 编号=&n_System;
Commit;
