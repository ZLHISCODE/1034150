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
--126591:刘兴洪,2018-06-07,增加误差费的处理.

Create Or Replace Procedure Zl_Third_Settlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --功能:三方接口结帐
  --入参:Xml_In:
  --<IN>
  --        <BRID>病人ID</BRID>         //病人ID
  --        <XM>姓名</XM>               //姓名
  --        <SFZH>身份证号</SFZH>       //身份证号
  --        <ZYID>主页ID</ZYID>         //主页ID
  --        <JSLX>2</JSLX>         //结算类型,1-门诊,2-住院.目前固定传2
  --        <JE></JE>         //本次结算总金额
  --       <JSLIST>
  --         <JS>
  --              <JSKLB>支付卡类别</JSKLB >
  --              <JSKH>支付卡号</ JSKH >
  --              <JSFS>支付方式</JSFS> //支付方式:现金;支票,如果是三方卡,可以传空
  --              <JSJE>结算金额</JSJE> //结算金额(正金额：个人补款，负金额：医院退款)<SFCYJ>为1时为冲预交金额
  --              <JYLSH>交易流水号</JYLSH>
  --              <ZY>摘要</ZY>
  --              <SFCYJ>是否冲预交</SFCYJ>  //是否冲预交，0-结算，1-冲预交.允冲预交时,只填JSJE节点
  --              <SFXFK>是否消费卡</SFXFK>  //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点
  --              <EXPENDLIST>  //扩展交易信息
  --                  <EXPEND>
  --                        <JYMC>交易名称</JYMC> //交易名称   退款时,传入冲预交的流水号
  --                        <JYLR>交易内容</JYLR> //交易内容   退款时,传入冲预交的金额
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --出参:Xml_Out
  --  <OUT>
  --       <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --    ――如无下列错误结点则说明正确执行
  --    <ERROR>
  --      <MSG>错误信息</MSG>
  --    </ERROR>
  --  </OUT>
  --------------------------------------------------------------------------------------------------
  n_主页id     病案主页.主页id%Type;
  n_病人id     病案主页.病人id%Type;
  v_姓名       病人信息.姓名%Type;
  v_身份证号   病人信息.身份证号%Type;
  n_结帐总额   病人预交记录.冲预交%Type;
  n_待结帐金额 病人预交记录.冲预交%Type;
  n_结算类型   Number(3);
  v_操作员编码 病人结帐记录.操作员编号%Type;
  v_操作员姓名 病人结帐记录.操作员姓名%Type;
  n_结帐id     病人结帐记录.Id%Type;
  n_冲预交金额 病人预交记录.冲预交%Type;
  d_结帐时间   Date;
  n_预交充值   病人预交记录.金额%Type;
  d_开始日期   Date;
  d_结束日期   Date;
  d_最小日期   Date;
  d_最大日期   Date;

  n_预交id     病人预交记录.Id%Type;
  n_科室id     病案主页.入院科室id%Type;
  n_结算卡序号 卡消费接口目录.编号%Type;
  n_时间类型   Number(3);
  v_Ids        Varchar2(20000);
  v_No         病人结帐记录.No%Type;
  v_预交no     病人预交记录.No%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_结算卡号   病人预交记录.卡号%Type;
  v_结算方式   病人预交记录.结算方式%Type;
  v_Temp       Varchar2(500);
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;

  n_Count    Number(18);
  n_Number   Number(2);
  n_费用id   门诊费用记录.Id%Type;
  n_记录性质 门诊费用记录.记录性质%Type;
  v_费用no   门诊费用记录.No%Type;
  n_序号     门诊费用记录.序号%Type;
  n_记录状态 门诊费用记录.记录状态%Type;
  n_执行状态 门诊费用记录.执行状态%Type;
  n_未结金额 门诊费用记录.实收金额%Type;
  n_结帐金额 门诊费用记录.实收金额%Type;
  n_误差费   门诊费用记录.实收金额%Type;

  v_卡类别     三方交易记录.类别%Type;
  v_消费卡结算 Varchar2(20000);

  Type t_费用结算明细 Is Ref Cursor;
  c_费用结算明细 t_费用结算明细;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/JSLX')),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_主页id, n_病人id, n_结帐总额, n_结算类型, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_结算类型 := Nvl(n_结算类型, 2);
  If n_结算类型 = 1 And Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  --0.相关检查
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许缴费!';
    Raise Err_Item;
  End If;

  --人员id,人员编号,人员姓名
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '系统不能认别有效的操作员,不允许结算!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;
  v_Err_Msg    := Null;

  For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Not (c_交易记录.结算卡类别 Is Null Or Nvl(c_交易记录.是否消费卡, '0') = '1' Or Nvl(c_交易记录.是否冲预交, 0) = 1) Then
    
      Select Decode(Translate(Nvl(c_交易记录.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
      Into n_Count
      From Dual;
    
      If Nvl(n_Count, 0) = 1 Then
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_交易记录.结算卡类别);
      Else
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_交易记录.结算卡类别;
      End If;
    
      If v_卡类别 Is Null Then
        v_Err_Msg := '不支持的结算方式,请检查！';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, 2) = 0 Then
        v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
        Raise Err_Special;
      End If;
    End If;
  End Loop;

  Select Max(入院科室id) Into n_科室id From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
  Select 病人结帐记录_Id.Nextval, Sysdate, Nextno(15) Into n_结帐id, d_结帐时间, v_No From Dual;

  n_时间类型 := Zl_Getsysparameter('结帐费用时间', 1137);

  Select 病人结帐记录_Id.Nextval, Sysdate, Nextno(15) Into n_结帐id, d_结帐时间, v_No From Dual;

  If n_结算类型 = 2 Then
    Open c_费用结算明细 For
      Select Max(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
             Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
             Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
      From 住院费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1
      Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
      Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, 序号;
  Else
  
    Open c_费用结算明细 For
      Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
             Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
             Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
      From 门诊费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 记帐费用 = 1
      Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
      Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
      Union All
      Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
             Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
             Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
      From 住院费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 记帐费用 = 1 And Mod(记录性质, 10) = 5
      Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
      Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, 序号;
  End If;

  n_待结帐金额 := 0;
  Loop
    Fetch c_费用结算明细
      Into n_费用id, n_记录性质, v_费用no, n_序号, n_记录状态, n_执行状态, d_最小日期, d_最大日期, n_未结金额, n_结帐金额;
    Exit When c_费用结算明细%NotFound;
  
    n_待结帐金额 := n_待结帐金额 + Nvl(n_未结金额, 0);
    If d_开始日期 Is Null Then
      d_开始日期 := d_最小日期;
    Elsif d_开始日期 > d_最小日期 Then
      d_开始日期 := d_最小日期;
    End If;
    If d_结束日期 Is Null Then
      d_结束日期 := d_最大日期;
    Elsif d_结束日期 < d_最大日期 Then
      d_结束日期 := d_最大日期;
    End If;
  
    If Nvl(n_结帐金额, 0) = 0 Then
      If n_费用id Is Not Null Then
        If Length(v_Ids || ',' || n_费用id) > 4000 Then
          v_Ids := Substr(v_Ids, 2);
          Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
          v_Ids := '';
        End If;
        v_Ids := v_Ids || ',' || n_费用id;
      Else
        Zl_结帐费用记录_Insert(0, v_费用no, n_记录性质, n_记录状态, n_执行状态, n_序号, n_未结金额, n_结帐id);
      End If;
    Else
      Zl_结帐费用记录_Insert(0, v_费用no, n_记录性质, n_记录状态, n_执行状态, n_序号, n_未结金额, n_结帐id);
    End If;
  
  End Loop;

  If v_Ids Is Not Null Then
    v_Ids := Substr(v_Ids, 2);
    Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
  End If;
  n_待结帐金额 := Round(n_待结帐金额, 6);

  If n_待结帐金额 <> Nvl(n_结帐总额, 0) Then
    v_Err_Msg := '传入的结帐金额与实际结帐金额不符,不允许结算!';
    Raise Err_Item;
  End If;

  Zl_病人结帐记录_Insert(n_结帐id, v_No, n_病人id, d_结帐时间, d_开始日期, d_结束日期, 0, 0, n_主页id, Null, 2, Null, n_结算类型);

  n_结帐金额 := 0;
  n_Count    := 0;
  For r_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_卡类别   := r_结算方式.结算方式;
    n_结帐金额 := n_结帐金额 + Nvl(r_结算方式.结算金额, 0);
  
    If Nvl(r_结算方式.是否冲预交, 0) = 0 Then
      --付款
      If n_Count = 1 Then
        v_Err_Msg := '结帐结算暂不支持多种结算方式!';
        Raise Err_Item;
      End If;
      n_卡类别id := Null;
      If r_结算方式.结算卡类别 Is Not Null Then
        Select Decode(Translate(Nvl(r_结算方式.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
      
        If Nvl(r_结算方式.是否消费卡, 0) = 1 Then
          If Nvl(n_Number, 0) = 1 Then
            Select Max(编号), Max(结算方式), Max(名称)
            Into n_结算卡序号, v_结算方式, v_卡类别
            From 卡消费接口目录
            Where 编号 = n_卡类别id And Nvl(启用, 0) = 1;
          Else
            Select Max(编号), Max(结算方式), Max(名称)
            Into n_结算卡序号, v_结算方式, v_卡类别
            From 卡消费接口目录
            Where 名称 = r_结算方式.结算卡类别 And Nvl(启用, 0) = 1;
          
          End If;
        
          If n_结算卡序号 Is Null Then
            v_Err_Msg := '未找到对应的消费卡信息';
            Raise Err_Item;
          
          End If;
          n_卡类别id := Null;
        
        Else
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID), Max(结算方式), Max(名称)
            Into n_卡类别id, v_结算方式, v_卡类别
            From 医疗卡类别
            Where ID = n_卡类别id And Nvl(是否启用, 0) = 1;
          Else
            Select Max(ID), Max(结算方式), Max(名称)
            Into n_卡类别id, v_结算方式, v_卡类别
            From 医疗卡类别
            Where 名称 = r_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
          End If;
        
          If n_卡类别id Is Null Then
            v_Err_Msg := '未找到对应的医疗卡信息!';
            Raise Err_Item;
          End If;
        End If;
      End If;
    
      If n_卡类别id Is Not Null Then
        --三方卡,生成住院预交款
        v_结算卡号 := r_结算方式.结算卡号;
        If r_结算方式.结算金额 > 0 Then
          --充值部分不应该算作本次结帐 
          Select 病人预交记录_Id.Nextval, Nextno(11) Into n_预交id, v_预交no From Dual;
          Zl_病人预交记录_Insert(n_预交id, v_预交no, Null, n_病人id, n_主页id, n_科室id, r_结算方式.结算金额, v_结算方式, '', '', '', '', '',
                           v_操作员编码, v_操作员姓名, Null, n_结算类型, n_卡类别id, Null, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, Null,
                           d_结帐时间, 0);
          For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                         From Table(Xmlsequence(Extract(r_结算方式.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
            Zl_三方结算交易_Insert(n_卡类别id, 0, r_结算方式.结算卡号, n_预交id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 1);
          End Loop;
	  --结帐处理
	  Zl_结帐预交记录_Insert(n_预交id,v_预交no, 1, r_结算方式.结算金额,  n_结帐id, n_病人id);
        Else
        
          Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, r_结算方式.结算金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间,
                           Null, Null, Null, Null, Null, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明);
        
          For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                         From Table(Xmlsequence(Extract(r_结算方式.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
            Zl_三方结算交易_Insert(n_卡类别id, 0, r_结算方式.结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
          End Loop;
        
        End If;
      
      Else
        If n_结算卡序号 Is Not Null Then
          --消费卡
          v_消费卡结算 := Nvl(v_消费卡结算, '') || '||' || n_结算卡序号 || '|' || r_结算方式.结算卡号 || '|0|' || r_结算方式.结算金额;
        Else
          --其他结算
          Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, r_结算方式.结算金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间,
                           Null, Null, Null, Null, Null, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明);
        End If;
      End If;
      n_Count := 1;
    Else
      --冲预交,目前默认全冲
      n_冲预交金额 := r_结算方式.结算金额;
      For r_预交 In (Select Min(ID) As ID, NO, 结算方式, Sum(Nvl(金额, 0)) - Sum(Nvl(冲预交, 0)) As 金额, 交易流水号
                   From 病人预交记录
                   Where 病人id = n_病人id And Mod(记录性质, 10) = 1 And Nvl(预交类别, 2) = 2 And (主页id = n_主页id Or 主页id Is Null)
                   Group By NO, 结算方式, 交易流水号
                   Having Sum(Nvl(金额, 0)) - Sum(Nvl(冲预交, 0)) <> 0) Loop
        Zl_结帐预交记录_Insert(r_预交.Id, r_预交.No, 1, r_预交.金额, n_结帐id, n_病人id);
        n_冲预交金额 := n_冲预交金额 - Nvl(r_预交.金额, 0);
      End Loop;
      If n_冲预交金额 <> 0 Then
        v_Err_Msg := '传入的预交冲销金额与实际不符,请检查!';
        Raise Err_Item;
      End If;
    End If;
  
    Update 三方交易记录
    Set 业务结算id = n_结帐id
    Where 流水号 = Nvl(r_结算方式.交易流水号, '-') And 类别 = v_卡类别 And 业务类型 = 2;
  End Loop;

  --消费卡处理
  If v_消费卡结算 Is Not Null Then
    v_消费卡结算 := Substr(v_消费卡结算, 3);
  End If;

  n_误差费   := Round(Nvl(n_结帐总额, 0) - Nvl(n_结帐金额, 0), 6);
  v_结算方式 := Null;
  If Abs(Nvl(n_误差费, 0)) > 1 Then
    v_Err_Msg := '计算的误差金额大于了1.00或小于-1.00元,不允许结帐操作,请检查!';
    Raise Err_Item;
  End If;

  n_结帐总额 := n_结帐金额;

  n_结帐金额 := 0;
  If Nvl(n_误差费, 0) <> 0 Then
    Select Nvl(Max(名称), '误差费') Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 9;
    n_结帐金额 := Nvl(n_误差费, 0);
  End If;
  If Nvl(n_误差费, 0) <> 0 Or v_消费卡结算 Is Not Null Then
    Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, n_结帐金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间, Null, Null,
                     Null, Null, Null, Null, Null, Null, Null, v_消费卡结算);
  End If;

  --检查结算信息总额与结算总额是否正确
  Select Sum(冲预交) Into n_结帐金额 From 病人预交记录 Where 结帐id = n_结帐id;
  If Round(n_结帐金额, 6) <> Round(n_结帐总额, 6) Then
  
    v_Err_Msg := '传入的结算合计金额与实际结帐金额合计不符,不允许结算!';
    Raise Err_Item;
  End If;

  Update 病人预交记录 Set 校对标志 = 0 Where 结帐id = n_结帐id And Nvl(校对标志, 0) <> 0;
  v_Temp := '<CZSJ>' || To_Char(d_结帐时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
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
--系统版本号
Update zlSystems Set 版本号='10.34.150.0004' Where 编号=&n_System;
Commit;
