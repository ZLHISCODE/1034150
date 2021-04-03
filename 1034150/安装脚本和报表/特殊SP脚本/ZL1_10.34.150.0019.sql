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
--135887:刘兴洪,2018-12-20,解决结帐作废时三方卡退回了预交的问是
Create Or Replace Procedure Zl_病人结帐记录_Delete
(
  No_In           病人结帐记录.No%Type,
  操作员编号_In   病人结帐记录.操作员编号%Type,
  操作员姓名_In   病人结帐记录.操作员姓名%Type,
  误差金额_In     病人预交记录.冲预交%Type := 0, --医保或预交退现金产生的误差
  结帐作废结算_In Varchar2 := Null, --结算方式|结算金额|结算号码||......
  预交退现金_In   Number := 0, --当预交款退现金时，结算方式及金额通过参数结帐作废结算_In传入
  冲销id_In       病人预交记录.结帐id%Type := Null,
  冲销时间_In     Date := Null,
  缴预交id_In     病人预交记录.Id%Type := Null, --在作废时将相关的金额充值到预交款时填写
  票据号_In       病人结帐记录.实际票号%Type := Null,
  领用id_In       票据领用记录.Id%Type := Null,
  票种_In         票据使用明细.票种%Type := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --该游标用于预交记录相关信息
  Cursor c_Deposit(v_Id 病人预交记录.结帐id%Type) Is
    Select 病人id, 记录性质, 结算方式, 冲预交, 预交类别 From 病人预交记录 Where 结帐id = v_Id;
  r_Depositrow c_Deposit%RowType;

  --该游标用于处理费用相关汇总表
  Cursor c_Money(v_Id 病人预交记录.结帐id%Type) Is
    Select NO, 开单部门id, 病人科室id, 执行部门id, 病人病区id, 病人id, 主页id, 收入项目id, 门诊标志, 结帐金额
    From 住院费用记录
    Where 结帐id = v_Id
    Union All
    Select NO, 开单部门id, 病人科室id, 执行部门id, 0 As 病人病区id, 病人id, 0 As 主页id, 收入项目id, 门诊标志, 结帐金额
    From 门诊费用记录
    Where 结帐id = v_Id;

  r_Moneyrow c_Money%RowType;

  --该游标包含病人的相关信息
  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, b.主页id, b.出院病床, b.当前病区id, b.出院科室id, Nvl(b.费别, a.费别) As 费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 病案主页 B, 医疗付款方式 C
    Where a.病人id = n_病人id And a.病人id = b.病人id(+) And Nvl(a.主页id, 0) = b.主页id(+) And a.医疗付款方式 = c.名称(+);
  r_Pati c_Pati%RowType;

  --过程变量
  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(50);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_实际票号 病人预交记录.实际票号%Type;
  v_误差no   住院费用记录.No%Type;
  v_误差     结算方式.名称%Type;
  n_病人id   病人信息.病人id%Type;

  n_原id   病人结帐记录.Id%Type;
  n_结帐id 病人结帐记录.Id%Type;
  n_打印id 票据打印内容.Id%Type;

  n_来源     Number; --1-门诊;2-住院;3-门诊和住院
  n_返回值   病人余额.预交余额%Type;
  n_组id     财务缴款分组.Id%Type;
  n_预交类别 Number;
  d_Date     Date;
  n_预交id   病人预交记录.Id%Type;
  n_卡结算id 病人结帐记录.Id%Type;
  v_打印id   票据打印内容.Id%Type;

  n_预交合计 病人预交记录.冲预交%Type;
  n_结帐合计 住院费用记录.结帐金额%Type;

Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  Select 名称 Into v_误差 From 结算方式 Where 性质 = 9 And Rownum = 1;

  Begin
    Select ID, 病人id, 实际票号 Into n_原id, n_病人id, v_实际票号 From 病人结帐记录 Where 记录状态 = 1 And NO = No_In;
    --最后一次打印的内容
    Select Max(ID)
    Into n_打印id
    From (Select b.Id
           From 票据使用明细 A, 票据打印内容 B
           Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 3 And b.No = No_In
           Order By a.使用时间 Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      Begin
        v_Err_Msg := '没有发现要作废的结帐单据,可能已经作废！';
        Raise Err_Item;
      End;
  End;

  Open c_Pati(n_病人id);
  Fetch c_Pati
    Into r_Pati; --体检系统调用此过程,团体结帐时没有病人信息

  d_Date := 冲销时间_In;
  If d_Date Is Null Then
    Select Sysdate Into d_Date From Dual;
  End If;

  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 3, No_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, 票种_In, 票据号_In, 1, 6, 领用id_In, v_打印id, d_Date, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  End If;

  n_结帐id := 冲销id_In;
  If Nvl(n_结帐id, 0) = 0 Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  End If;

  --病人结帐记录
  Insert Into 病人结帐记录
    (ID, NO, 实际票号, 记录状态, 中途结帐, 病人id, 操作员编号, 操作员姓名, 开始日期, 结束日期, 收费时间, 备注, 原因, 缴款组id, 结帐类型)
    Select n_结帐id, NO, 实际票号, 2, 中途结帐, 病人id, 操作员编号_In, 操作员姓名_In, 开始日期, 结束日期, d_Date, 备注, 原因, n_组id, 结帐类型
    From 病人结帐记录
    Where ID = n_原id;

  Update 病人结帐记录 Set 记录状态 = 3 Where ID = n_原id;

  --作废收回票据(可能以前没有使用票据,无法收回)
  If n_打印id Is Not Null Then
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
      Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
      From 票据使用明细
      Where 打印id = n_打印id And 票种 In (1, 3) And 性质 = 1;
  End If;

  --病人预交记录(冲预交及缴款)
  If 结帐作废结算_In Is Null Then
    --插入普通的结算信息(不包含三方退款
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 病人id, 主页id, 科室id, Null,
             结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date, 操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 合作单位, 2
      From 病人预交记录
      Where 结帐id = n_原id And Mod(记录性质, 10) <> 1 And
            ((卡类别id Is Not Null And Nvl(冲预交, 0) > 0) Or (Mod(记录性质, 10) <> 1 And 卡类别id Is Null));
  
    --插入预交，但不包含已经退的预交
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, 11 As 记录性质, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
             d_Date, 操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2
      From (Select a.No, Sum(a.冲预交) As 冲预交, Max(a.实际票号) As 实际票号, Max(a.记录状态) As 记录状态, Max(a.病人id) As 病人id,
                    Max(a.主页id) As 主页id, Max(a.科室id) As 科室id, Max(a.结算方式) As 结算方式, Max(a.结算号码) As 结算号码, Max(a.摘要) As 摘要,
                    Max(a.缴款单位) As 缴款单位, Max(a.单位开户行) As 单位开户行, Max(a.单位帐号) As 单位帐号, Max(a.预交类别) As 预交类别,
                    Max(a.卡类别id) As 卡类别id, Max(a.结算卡序号) As 结算卡序号, Max(a.卡号) As 卡号, Max(a.交易流水号) As 交易流水号,
                    Max(a.交易说明) As 交易说明, Max(a.合作单位) As 合作单位
             From (Select a.No, a.冲预交, a.实际票号, a.记录状态, a.病人id, a.主页id, a.科室id, a.结算方式, a.结算号码, a.摘要, a.缴款单位, a.单位开户行,
                           a.单位帐号, a.预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号, a.交易说明, a.合作单位
                    From 病人预交记录 A
                    Where a.结帐id = n_原id And Mod(记录性质, 10) = 1
                    Union All
                    Select b.No, -1 * a.金额 As 冲预交, b.实际票号, b.记录状态, b.病人id, b.主页id, b.科室id, b.结算方式, b.结算号码, b.摘要, b.缴款单位,
                           b.单位开户行, b.单位帐号, b.预交类别, b.卡类别id, b.结算卡序号, b.卡号, b.交易流水号, b.交易说明, b.合作单位

                    From 三方退款信息 A, 病人预交记录 B
                    Where a.结帐id = n_原id And a.记录id = b.Id) A
             Group By a.No) A
      Where Nvl(a.冲预交, 0) <> 0;
  
    --消费卡处理
    For c_消费卡结算 In (Select a.Id, a.结算卡序号, Nvl(b.名称, '消费卡') As 卡名称
                    From 病人预交记录 A, 卡消费接口目录 B
                    Where a.结算卡序号 = b.编号(+) And a.结帐id = n_原id And Nvl(a.结算卡序号, 0) <> 0) Loop
    
      For c_消费卡 In (Select a.Id, a.接口编号, a.消费卡id, a.序号, a.记录状态, a.结算方式, a.结算金额, a.卡号, a.交易流水号, a.交易时间, a.备注, a.结算标志,
                           b.停用日期, b.回收时间
                    From 病人卡结算记录 A, 消费卡目录 B
                    Where a.消费卡id = b.Id(+) And
                          a.Id In (Select 卡结算id From 病人卡结算对照 Where 预交id = c_消费卡结算.Id)) Loop
      
        If Nvl(c_消费卡.停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || c_消费卡.卡号 || '"的' || c_消费卡结算.卡名称 || '已经被他人停用，不能再进行结帐作废,请检查！';
          Raise Err_Item;
        End If;
      
        If c_消费卡.回收时间 < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || c_消费卡.卡号 || '"的' || c_消费卡结算.卡名称 || '已经回收，不能退费,请检查！';
          Raise Err_Item;
        End If;
        Select 病人卡结算记录_Id.Nextval Into n_卡结算id From Dual;
      
        Insert Into 病人卡结算记录
          (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Values
          (n_卡结算id, c_消费卡结算.结算卡序号, c_消费卡.消费卡id, c_消费卡.序号, 2, c_消费卡.结算方式, -1 * c_消费卡.结算金额, c_消费卡.卡号, c_消费卡.交易流水号, d_Date,
           Null, 0);
      
        Insert Into 病人卡结算对照
          (预交id, 卡结算id)
          Select ID, n_卡结算id
          From 病人预交记录
          Where 结帐id = n_结帐id And 结算卡序号 = Nvl(c_消费卡结算.结算卡序号, 0);
      
        Update 消费卡目录 Set 余额 = 余额 + c_消费卡.结算金额 Where ID = c_消费卡.消费卡id;
        If Sql%NotFound Then
          v_Err_Msg := '卡号为' || c_消费卡.卡号 || '的' || c_消费卡结算.卡名称 || '未找到!';
          Raise Err_Item;
        End If;
      End Loop;
    End Loop;
  
  Else
    --1.先处理冲预交部分
    If 预交退现金_In = 0 Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 病人id, 主页id, 科室id,
               Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date, 操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2
        From 病人预交记录
        Where 结帐id = n_原id And 记录性质 In (1, 11) And Nvl(冲预交, 0) <> 0;
    End If;
  
    --2.再处理结帐结算,包括医保和非医保
    v_结算内容 := 结帐作废结算_In || ' ||'; --以空格分开以|结尾,没有结算号码的
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算号码 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, No_In, v_实际票号, 12, 1, n_病人id, r_Pati.主页id, r_Pati.出院科室id, Null, v_结算方式, v_结算号码, '结帐作废退款',
         Null, Null, Null, d_Date, 操作员姓名_In, 操作员编号_In, -1 * n_结算金额, n_结帐id, n_组id, 2);
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;
  --确定结帐的费用记录来源
  Begin
    Select Case
             When Nvl(Max(住院), 0) = 1 And Nvl(Max(门诊), 0) = 1 Then
              3
             When Nvl(Max(住院), 0) = 1 Then
              2
             Else
              1
           End
    Into n_来源
    From (Select 1 As 住院, 0 As 门诊
           From 住院费用记录
           Where 结帐id = n_原id And Rownum = 1
           Union All
           Select 0 As 住院, 1 As 门诊
           From 门诊费用记录
           Where 结帐id = n_原id And Rownum = 1);
  
  Exception
    When Others Then
      n_来源 := 3;
  End;

  If 误差金额_In <> 0 Then
    Update 病人预交记录
    Set 冲预交 = 冲预交 + 误差金额_In
    Where NO = No_In And 记录性质 = 12 And 记录状态 = 1 And 结帐id = n_结帐id And 结算方式 = v_误差;
    If Sql%RowCount = 0 Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, No_In, v_实际票号, 12, 1, n_病人id, r_Pati.主页id, r_Pati.出院科室id, Null, v_误差, Null, '结帐作废退款', Null,
         Null, Null, d_Date, 操作员姓名_In, 操作员编号_In, 误差金额_In, n_结帐id, n_组id, 2);
    End If;
  End If;

  If n_来源 = 2 Or n_来源 = 3 Then
    --作废结帐对应的费用记录:不包含原始结帐产生的误差项目
    Insert Into 住院费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id,
       病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id,
       开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要,
       缴款组id, 医疗小组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 序号, 从属父号, 价格父号, 多病人单,
             记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次,
             加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人,
             执行时间, 操作员姓名, 操作员编号, -1 * 结帐金额, n_结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id, 医疗小组id
      From 住院费用记录
      Where 结帐id = n_原id;
  End If;

  If n_来源 = 1 Or n_来源 = 3 Then
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
       收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
       执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 序号, 从属父号, 价格父号, 记帐单id,
             病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费,
             记帐费用, 收入项目id, 收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号,
             -1 * 结帐金额, n_结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id
      From 门诊费用记录
      Where 结帐id = n_原id;
  End If;
  --相关汇总表处理
  For r_Depositrow In c_Deposit(n_结帐id) Loop
    If r_Depositrow.记录性质 In (1, 11) Then
    
      --病人余额(预交)
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - r_Depositrow.冲预交 --注:新的结帐ID产生的是负数金额
      Where 病人id = r_Depositrow.病人id And 类型 = Nvl(r_Depositrow.预交类别, 2) And 性质 = 1
      Returning 预交余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (r_Depositrow.病人id, 1, Nvl(r_Depositrow.预交类别, 2), -1 * r_Depositrow.冲预交, 0);
        n_返回值 := -1 * r_Depositrow.冲预交;
      End If;
    
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额
        Where 性质 = 1 And 病人id = r_Depositrow.病人id And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    
    Else
      --人员缴款余额,医保不支持作废的结算方式在新的预交结算中已被处理为了退现金,
      --此处用加,表示收回退给病人的现金(结帐时,退款是负,作废时是正)
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + r_Depositrow.冲预交
      Where 收款员 = 操作员姓名_In And 结算方式 = r_Depositrow.结算方式 And 性质 = 1
      Returning 余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Depositrow.结算方式, 1, r_Depositrow.冲预交);
        n_返回值 := -1 * r_Depositrow.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 结算方式 = r_Depositrow.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    End If;
  End Loop;

  For r_Moneyrow In c_Money(n_结帐id) Loop
    --病人余额 ,误差项已结帐,所以不需要更新这两个汇总表
    If Nvl(v_误差no, 'sc') <> Nvl(r_Moneyrow.No, 'sc') Then
      If Nvl(r_Moneyrow.门诊标志, 0) = 1 Or Nvl(r_Moneyrow.门诊标志, 0) = 2 Then
        n_预交类别 := r_Moneyrow.门诊标志;
      Elsif Nvl(r_Moneyrow.主页id, 0) = 0 Or Nvl(r_Moneyrow.门诊标志, 0) = 4 Then
        --体检:门诊病人
        n_预交类别 := 1;
      Else
        n_预交类别 := 2;
      End If;
    
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) - r_Moneyrow.结帐金额 --注:新的结帐ID产生的是负数金额
      Where 病人id = r_Moneyrow.病人id And 类型 = n_预交类别 And 性质 = 1
      Returning 费用余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (r_Moneyrow.病人id, 1, n_预交类别, 0, -1 * r_Moneyrow.结帐金额);
        n_返回值 := -1 * r_Moneyrow.结帐金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额
        Where 病人id = r_Moneyrow.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - r_Moneyrow.结帐金额
      Where 病人id = r_Moneyrow.病人id And Nvl(主页id, 0) = Nvl(r_Moneyrow.主页id, 0) And
            Nvl(病人病区id, 0) = Nvl(r_Moneyrow.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Moneyrow.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(r_Moneyrow.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Moneyrow.执行部门id, 0) And
            收入项目id + 0 = r_Moneyrow.收入项目id And 来源途径 + 0 = r_Moneyrow.门诊标志;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (r_Moneyrow.病人id, Decode(r_Moneyrow.主页id, Null, Null, 0, Null, r_Moneyrow.主页id),
           Decode(r_Moneyrow.病人病区id, Null, Null, 0, Null, r_Moneyrow.病人病区id), r_Moneyrow.病人科室id, r_Moneyrow.开单部门id,
           r_Moneyrow.执行部门id, r_Moneyrow.收入项目id, r_Moneyrow.门诊标志, -1 * r_Moneyrow.结帐金额);
      End If;
    End If;
  End Loop;

  If Nvl(缴预交id_In, 0) <> 0 Then
    --作废时将退款金额充值到预交款帐户,这里标明是本次结帐缴存的
    Update 病人预交记录 Set 结帐id = 冲销id_In Where ID = 缴预交id_In And 结帐id Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '未找到对应的预交款记录！';
      Raise Err_Item;
    End If;
  End If;

  --需要检查退款合计是否相等
  Select Round(Sum(冲预交), 5), Round(Sum(结帐合计), 5)
  Into n_预交合计, n_结帐合计
  From (Select Sum(冲预交) As 冲预交, 0 As 结帐合计
         From 病人预交记录
         Where 结帐id = n_结帐id
         Union All
         Select 0 As 冲预交, Sum(结帐金额) As 结帐合计
         From 住院费用记录
         Where 结帐id = n_结帐id
         Union All
         Select 0 As 冲预交, Sum(结帐金额) As 结帐合计
         From 门诊费用记录
         Where 结帐id = n_结帐id);

  If Nvl(n_预交合计, 0) <> Nvl(n_结帐合计, 0) Then
    v_Err_Msg := '本次结算作废合计(' || Trim(To_Char(Nvl(n_预交合计, 0), '99999999999.99999')) || ')与本次结帐作废费用合计(' ||
                 Trim(To_Char(Nvl(n_结帐合计, 0), '99999999999.99999')) || ')不等，请检查当前结算与费用明细合计是否一致！';
    Raise Err_Item;
  End If;

  Close c_Pati;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结帐记录_Delete;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.150.0019' Where 编号=&n_System;
Commit;
