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
--125588:秦龙,2018-07-02,处理卫材库存实际数量为0的调价
Create Or Replace Procedure Zl_材料收发记录_成本价调价(材料id_In In 药品收发记录.药品id%Type) As
  v_No         药品收发记录.No%Type;
  v_应付id     应付记录.Id%Type; --应付记录的ID 
  v_应付单据号 应付记录.No%Type;
  d_调价时间   Date;
  n_序号       Number(8);
  n_库房id     药品收发记录.库房id%Type;
  n_入出类别id 药品收发记录.入出类别id%Type;
  n_入出系数   药品收发记录.入出系数%Type;
  n_收发id     药品收发记录.Id%Type;
  n_调整额     药品收发记录.零售金额%Type;
  n_原成本价   药品收发记录.成本价%Type;
  n_新成本价   药品收发记录.成本价%Type;
  n_平均成本价 药品库存.平均成本价%Type;
  v_调价id     成本价调价信息.Id%Type;
  v_调价汇总号 成本价调价信息.调价汇总号%Type;
  n_Count      Number(1) := 0;

  Cursor c_Stock Is --当前库存 
    Select 上次供应商id, a.库房id, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, a.上次批号, a.效期, a.上次产地, a.灭菌效期,
           Decode(Sign(Nvl(a.批次, 0)), 1, a.上次采购价, a.平均成本价) As 原成本价
    From 药品库存 A
    Where a.性质 = 1 And Nvl(a.实际数量, 0) <> 0 And a.药品id = 材料id_In
    Order By a.库房id;

  v_Stock c_Stock%RowType;
Begin
  d_调价时间 := Sysdate;
  n_库房id   := 0;

  --判断是否存在无库存调价 
  Begin
    Select ID, 新成本价, 调价汇总号
    Into v_调价id, n_新成本价, v_调价汇总号
    From 成本价调价信息
    Where 执行日期 Is Null And Nvl(库房id, 0) = 0 And 药品id = 材料id_In;
  Exception
    When Others Then
      v_调价id   := 0;
      n_新成本价 := Null;
  End;

  --无库存调价 
  If v_调价id > 0 Then
    --根据当前库存重新产生调价信息 
    For v_Stock In c_Stock Loop
      Zl_材料成本调价_Insert(v_Stock.上次供应商id, v_Stock.库房id, v_Stock.材料id, v_Stock.批次, v_Stock.上次批号, v_Stock.原成本价, n_新成本价,
                       Null, Null, 0, 0, v_调价汇总号);
      n_Count := n_Count + 1;
    End Loop;
  
    If n_Count > 0 Then
      --如果当前有库存记录，则删除无库存调价记录 
      Delete 成本价调价信息 Where ID = v_调价id;
    Else
      Update 成本价调价信息 Set 执行日期 = d_调价时间 Where ID = v_调价id;
    
      Update 材料特性 Set 成本价 = n_新成本价 Where 材料id = 材料id_In And 成本价 <> n_新成本价;
    End If;
  End If;

  --取库存差价调整的入出类别ID 
  Select b.Id, b.系数
  Into n_入出类别id, n_入出系数
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 33 And Rownum < 2;

  For c_成本调整 In (Select a.库房id, a.药品id As 材料id, Nvl(a.批次, 0) 批次, a.上次供应商id, a.实际数量, a.实际金额, a.实际差价, a.上次产地 As 产地,
                        a.上次批号 As 批号, a.灭菌效期, a.效期, a.上次生产日期 As 生产日期, a.批准文号, Nvl(a.平均成本价, 0) As 原成本价, b.新成本价, b.发票号,
                        b.发票日期, b.发票金额, Nvl(a.上次采购价, 0) As 上次采购价, b.Id As 调价id
                 From 药品库存 A, 成本价调价信息 B
                 Where a.药品id = b.药品id And Nvl(a.上次供应商id, 0) = Nvl(b.供药单位id, 0) And a.库房id = b.库房id And
                       Nvl(a.批次, 0) = Nvl(b.批次, 0) And a.性质 = 1 And b.执行日期 Is Null And a.药品id = 材料id_In
                 Order By a.库房id) Loop
    If n_库房id <> c_成本调整.库房id Then
      n_序号   := 1;
      n_库房id := c_成本调整.库房id;
      v_No     := Nextno(71, n_库房id);
    Else
      n_序号 := n_序号 + 1;
    End If;
  
    Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
  
    /*If Nvl(c_成本调整.实际数量, 0) = 0 And Nvl(c_成本调整.实际金额, 0) = 0 And Nvl(c_成本调整.实际差价, 0) = 0 Then
    --数量,金额、差价都为0，则表示数据是填单下可用数量出库产生的单据，此单据还没有审核，因此只需要更新调价信息，其他不更新
    Update 材料特性 Set 成本价 = c_成本调整.新成本价 Where 材料id = c_成本调整.材料id;
    
    Update 成本价调价信息
    Set 收发id = n_收发id, 执行日期 = d_调价时间, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地, 批号 = c_成本调整.批号
    Where ID = c_成本调整.调价id;*/
    If Nvl(c_成本调整.实际数量, 0) = 0 And (Nvl(c_成本调整.实际金额, 0) <> 0 Or Nvl(c_成本调整.实际差价, 0) <> 0) Then
      --数量=0 金额或差价<>0时只更新库存表中对应的平均成本价和特性表中成本价，并产生成本价修正数据但是差价差=0，只记录最新成本价
      --产生调价记录，只记录最新成本价
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期, 审核人,
         审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率)
      Values
        (n_收发id, 1, 18, v_No, n_序号, c_成本调整.库房id, n_入出类别id, c_成本调整.上次供应商id, n_入出系数, c_成本调整.材料id, c_成本调整.批次, c_成本调整.产地,
         c_成本调整.批号, c_成本调整.效期, 0, c_成本调整.实际金额, c_成本调整.实际差价, 0, '卫生材料成本价调价', Zl_Username, d_调价时间, Zl_Username, d_调价时间,
         c_成本调整.生产日期, c_成本调整.批准文号, c_成本调整.新成本价, 1, c_成本调整.原成本价);
      --更新库存      
      Update 药品库存
      Set 平均成本价 = c_成本调整.新成本价, 上次采购价 = c_成本调整.新成本价
      Where 库房id = c_成本调整.库房id And 药品id = c_成本调整.材料id And Nvl(批次, 0) = c_成本调整.批次 And 性质 = 1;
      Update 材料特性 Set 成本价 = c_成本调整.新成本价 Where 材料id = c_成本调整.材料id;
    
      Update 成本价调价信息
      Set 收发id = n_收发id, 执行日期 = d_调价时间, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地, 批号 = c_成本调整.批号
      Where ID = c_成本调整.调价id;
    Else
      --调整相应的库存:原成本金额-实新成本金额 
      n_调整额   := (c_成本调整.实际金额 - c_成本调整.实际差价) - Round(c_成本调整.新成本价 * c_成本调整.实际数量, 2);
      n_原成本价 := c_成本调整.原成本价;
    
      If n_原成本价 <= 0 Then
        n_原成本价 := c_成本调整.上次采购价;
      End If;
    
      --目前：收发记录对应: 
      -- 扣率--> 原成本价 
      -- 单量-->新成本价 
      -- 填写数量-->库存实际数量 
      -- 零售价-->库存实际金额 
      -- 成本价-->库存实际差价 
      -- 差价-->本次调整额 
    
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期, 审核人,
         审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率)
      Values
        (n_收发id, 1, 18, v_No, n_序号, c_成本调整.库房id, n_入出类别id, c_成本调整.上次供应商id, n_入出系数, c_成本调整.材料id, c_成本调整.批次, c_成本调整.产地,
         c_成本调整.批号, c_成本调整.效期, c_成本调整.实际数量, c_成本调整.实际金额, c_成本调整.实际差价, n_调整额, '卫生材料成本价调价', Zl_Username, d_调价时间,
         Zl_Username, d_调价时间, c_成本调整.生产日期, c_成本调整.批准文号, c_成本调整.新成本价, 1, n_原成本价);
    
      --更新库存 
      Update 药品库存
      Set 实际差价 = Nvl(实际差价, 0) + n_调整额
      Where 库房id = c_成本调整.库房id And 药品id = c_成本调整.材料id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 实际差价, 上次批号, 效期, 上次产地, 上次供应商id, 上次生产日期, 批准文号, 灭菌效期)
        Values
          (c_成本调整.库房id, c_成本调整.材料id, c_成本调整.批次, 1, n_调整额, c_成本调整.批号, c_成本调整.效期, c_成本调整.产地, c_成本调整.上次供应商id, c_成本调整.生产日期,
           c_成本调整.批准文号, c_成本调整.灭菌效期);
      End If;
    
      Update 药品库存
      Set 上次采购价 = c_成本调整.新成本价
      Where 药品id = c_成本调整.材料id And 上次采购价 <> c_成本调整.新成本价;
    
      Update 材料特性
      Set 成本价 = c_成本调整.新成本价
      Where 材料id = c_成本调整.材料id And 成本价 <> c_成本调整.新成本价;
    
      --重新计算库存表中的平均成本价 
      Update 药品库存
      Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
      Where 药品id = c_成本调整.材料id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 库房id = c_成本调整.库房id And 性质 = 1 And
            Nvl(实际数量, 0) <> 0;
      If Sql%NotFound Then
        Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = c_成本调整.材料id;
        Update 药品库存
        Set 平均成本价 = n_平均成本价
        Where 药品id = c_成本调整.材料id And 库房id = c_成本调整.库房id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 性质 = 1;
      End If;
    
      --更新成本价调价信息 
      Update 成本价调价信息
      Set 收发id = n_收发id, 执行日期 = d_调价时间, 原成本价 = n_原成本价, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地,
          批号 = c_成本调整.批号
      Where ID = c_成本调整.调价id;
    End If;
  End Loop;

  --产生应付记录 
  For c_应付 In (Select Distinct a.供药单位id, a.药品id, a.发票号, a.发票日期, a.发票金额, b.名称, b.计算单位, b.规格
               From 成本价调价信息 A, 收费项目目录 B
               Where a.药品id = b.Id And Nvl(a.应付款变动, 0) = 1 And Nvl(a.供药单位id, 0) <> 0 And a.药品id = 材料id_In
               Order By a.供药单位id) Loop
  
    v_应付单据号 := Nextno(67);
  
    Select 应付记录_Id.Nextval Into v_应付id From Dual;
  
    Insert Into 应付记录
      (ID, 记录性质, 记录状态, 单位id, NO, 系统标识, 发票号, 发票日期, 发票金额, 品名, 规格, 填制人, 填制日期, 审核人, 审核日期, 摘要)
    Values
      (v_应付id, 1, 1, c_应付.供药单位id, v_应付单据号, 5, c_应付.发票号, c_应付.发票日期, c_应付.发票金额, c_应付.名称, c_应付.规格, Zl_Username, d_调价时间,
       Zl_Username, d_调价时间, '成本价调价自动产生应付款变动记录');
  
    If Nvl(c_应付.供药单位id, 0) <> 0 Then
      Update 应付余额 Set 金额 = Nvl(金额, 0) + Nvl(c_应付.发票金额, 0) Where 单位id = c_应付.供药单位id And 性质 = 1;
      If Sql%NotFound Then
        Insert Into 应付余额 (单位id, 性质, 金额) Values (c_应付.供药单位id, 1, Nvl(c_应付.发票金额, 0));
      End If;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_成本价调价;
/

--125588:秦龙,2018-07-02,处理库存实际数量为0的卫材调价
Create Or Replace Procedure Zl_材料收发记录_Adjust
(
  调价id_In   In Number, --调价记录的ID
  定价_In     In Number := 0, --是否转为定价销售（更新材料特性、收费细目中的变价）
  材料id_In   In Number := 0, --当不为0时表示是成本价调价，不处理售价相关内容
  Billinfo_In In Varchar2 := Null --用于时价卫材按批次调价。格式:"批次1,现价1|批次2,现价2|....."
) As
  n_入出类别id 药品收发记录.入出类别id%Type; --入出类别
  v_调价单据号 药品收发记录.No%Type; --调价单号
  d_生效日期   Date; --调价生效时间
  n_执行调价   Number(1); --调价时刻到了
  n_实价材料   Number(1); --时价药品
  n_收费细目id Number(18); --收费细目ID
  d_审核日期   药品收发记录.审核日期%Type;
  n_零售金额   药品库存.实际金额%Type;
  n_零售价     药品库存.零售价%Type;
  n_序号       Integer(8);
  v_Infotmp    Varchar2(4000);
  v_Fields     Varchar2(4000);
  n_批次       Number(18);
  n_现价       收费价目.现价%Type;
  n_原价       收费价目.原价%Type;
  n_收发id     药品收发记录.Id%Type;
  n_时价分批   Number(1);

  Cursor c_Price --普通调价
  Is
    Select 1 记录状态, 13 单据, v_调价单据号 NO, Rownum 序号, n_入出类别id 入出类别id, m.材料id 药品id, s.批次 批次, Null 批号, s.效期,
           Decode(s.上次产地, Null, q.产地, s.上次产地) 产地, 1 付数, s.实际数量 填写数量, 0 实际数量, a.原价 成本价, 0 成本金额, a.现价 零售价, 0 扣率,
           Nvl(s.零售价, 0) As 库存零售价, s.实际金额 As 库存金额, s.实际差价 As 库存差价, '卫材调价' 摘要, User 填制人, Sysdate 填制日期, s.库房id 库房id,
           1 入出系数, a.Id 价格id, s.上次生产日期, s.灭菌效期, s.批准文号, s.上次供应商id,
           Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, a.原价, Nvl(s.实际金额, 0) / s.实际数量)) As 原售价
    From 药品库存 S, 材料特性 M, 收费价目 A, 收费项目目录 Q
    Where s.药品id = m.材料id And m.材料id = q.Id And m.材料id = a.收费细目id And s.性质 = 1 And a.变动原因 = 0 And a.Id = 调价id_In And
          a.执行日期 <= Sysdate;

  Cursor c_时价按批次调价 --时价卫材按批次调价
  Is
    Select 1 记录状态, 13 单据, v_调价单据号 NO, n_序号 + Rownum 序号, n_入出类别id 入出类别id, s.药品id 药品id, s.批次 批次, Null 批号, s.效期,
           Decode(s.上次产地, Null, b.产地, s.上次产地) 产地, 1 付数, Nvl(s.实际数量, 0) 填写数量, 0 实际数量, a.原价 成本价, 0 成本金额, n_现价 零售价, 0 扣率,
           '卫材调价' 摘要, User 填制人, Sysdate 填制日期, s.库房id 库房id, 1 入出系数, a.Id 价格id, Nvl(b.是否变价, 0) As 时价, s.实际金额 As 库存金额,
           s.实际差价 As 库存差价, Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, a.原价, Nvl(s.实际金额, 0) / s.实际数量)) As 原售价
    From 药品库存 S, 材料特性 M, 收费价目 A, 收费项目目录 B
    Where s.药品id = m.材料id And m.材料id = a.收费细目id And a.收费细目id = b.Id And s.性质 = 1 And a.变动原因 = 0 And a.Id = 调价id_In And
          a.执行日期 <= Sysdate And Nvl(s.批次, 0) = n_批次;
Begin

  If 材料id_In <> 0 Then
    --成本价调价
    Zl_材料收发记录_成本价调价(材料id_In);
    Return;
  End If;

  --取入出类别ID
  Select 类别id Into n_入出类别id From 药品单据性质 Where 单据 = 13;

  --取序列
  Select Nextno(147) Into v_调价单据号 From Dual;
  --取调价记录生效日期
  Select 收费细目id, 执行日期 Into n_收费细目id, d_生效日期 From 收费价目 Where ID = 调价id_In;
  --取该材料是否是时价药品
  Select Nvl(是否变价, 0) Into n_实价材料 From 收费项目目录 Where ID = n_收费细目id;

  If Sysdate >= d_生效日期 Then
    n_执行调价 := 1;
  Else
    n_执行调价 := 0;
  End If;

  If n_执行调价 = 1 Then
    d_审核日期 := Sysdate;
    --普通调价处理
    If Billinfo_In = '' Or Billinfo_In Is Null Then
      --非时价药品调价
      For c_调价 In c_Price Loop
        /*If Nvl(c_调价.填写数量, 0) = 0 And Nvl(c_调价.库存金额, 0) = 0 And Nvl(c_调价.库存差价, 0) = 0 Then
        Null;*/
        If Nvl(c_调价.填写数量, 0) = 0 And (Nvl(c_调价.库存金额, 0) <> 0 Or Nvl(c_调价.库存差价, 0) <> 0) Then
          --数量=0 金额或差价<>0时只更新库存表中对应的零售价,并产生售价修正数据但是金额差=0，只记录最新售价，金额差和差价差不填数据

        
          --产生调价影响记录
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 摘要, 填制人, 填制日期,
             库房id, 入出系数, 价格id, 审核人, 审核日期, 生产日期, 灭菌效期, 批准文号, 供药单位id, 单量, 频次)
          Values
            (药品收发记录_Id.Nextval, c_调价.记录状态, c_调价.单据, c_调价.No, c_调价.序号, c_调价.入出类别id, c_调价.药品id, c_调价.批次, c_调价.批号, c_调价.效期,
             c_调价.产地, c_调价.付数, c_调价.填写数量, c_调价.实际数量, Decode(n_实价材料, 1, c_调价.原售价, c_调价.成本价), c_调价.成本金额, c_调价.零售价, c_调价.扣率,
             c_调价.摘要, c_调价.填制人, c_调价.填制日期, c_调价.库房id, c_调价.入出系数, c_调价.价格id, User, d_审核日期, c_调价.上次生产日期, c_调价.灭菌效期,
             c_调价.批准文号, c_调价.上次供应商id, c_调价.库存金额, c_调价.库存差价);
        
          --更新材料库存 ，只有时价卫材才更新零售价
          Update 药品库存
          Set 零售价 = Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null)
          Where 库房id = c_调价.库房id And 药品id = c_调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(c_调价.批次, 0);
        Else
          If n_实价材料 = 1 Then
            If c_调价.库存零售价 = 0 Then
              n_零售价 := c_调价.库存金额 / c_调价.填写数量;
            Else
              n_零售价 := c_调价.库存零售价;
            End If;
          Else
            n_零售价 := c_调价.成本价;
          End If;
          n_零售金额 := Round((c_调价.零售价 - n_零售价) * c_调价.填写数量, 2);
        
          --产生调价影响记录
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
             填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 生产日期, 灭菌效期, 批准文号, 供药单位id, 单量, 频次)
          Values
            (药品收发记录_Id.Nextval, c_调价.记录状态, c_调价.单据, c_调价.No, c_调价.序号, c_调价.入出类别id, c_调价.药品id, c_调价.批次, c_调价.批号, c_调价.效期,
             c_调价.产地, c_调价.付数, c_调价.填写数量, c_调价.实际数量, Decode(n_实价材料, 1, c_调价.原售价, c_调价.成本价), c_调价.成本金额, c_调价.零售价, c_调价.扣率,
             n_零售金额, n_零售金额, c_调价.摘要, c_调价.填制人, c_调价.填制日期, c_调价.库房id, c_调价.入出系数, c_调价.价格id, User, d_审核日期, c_调价.上次生产日期,
             c_调价.灭菌效期, c_调价.批准文号, c_调价.上次供应商id, c_调价.库存金额, c_调价.库存差价);
        
          --更新材料库存
          Update 药品库存
          Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额,
              零售价 = Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null)
          Where 库房id = c_调价.库房id And 药品id = c_调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(c_调价.批次, 0);
        
          If Sql%RowCount = 0 Then
            Insert Into 药品库存
              (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
            Values
              (c_调价.库房id, c_调价.药品id, c_调价.批次, 1, 0, 0, n_零售金额, n_零售金额, c_调价.效期, c_调价. 灭菌效期, c_调价.上次供应商id, c_调价.成本价,
               c_调价.批号, c_调价.上次生产日期, c_调价.产地, c_调价.批准文号,
               Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null));
          End If;
        End If;
      End Loop;
    Else
      --时价分批调价处理
      n_序号 := 0;
      --时价药品按批次调价
      v_Infotmp := Billinfo_In || '|';
      While v_Infotmp Is Not Null Loop
        --分解单据ID串
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        n_批次    := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        n_现价    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        For v_时价按批次调价 In c_时价按批次调价 Loop
          If v_时价按批次调价.填写数量 <> 0 Then
            n_原价 := Nvl(v_时价按批次调价.库存金额, 0) / v_时价按批次调价.填写数量;
          Else
            n_原价 := v_时价按批次调价.成本价;
          End If;
        
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
          If Nvl(v_时价按批次调价.填写数量, 0) = 0 And Nvl(v_时价按批次调价.库存金额, 0) = 0 And Nvl(v_时价按批次调价.库存差价, 0) = 0 Then
            Null;
          Elsif Nvl(v_时价按批次调价.填写数量, 0) = 0 And (Nvl(v_时价按批次调价.库存金额, 0) <> 0 Or Nvl(v_时价按批次调价.库存差价, 0) <> 0) Then
            --数量=0 金额或差价<>0时只更新库存表中对应的零售价,并产生售价修正数据但是金额差=0，只记录最新售价，金额差和差价差不填数据

          
            --产生调价影响记录
            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 摘要, 填制人, 填制日期,
               库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次)
            Values
              (n_收发id, v_时价按批次调价.记录状态, v_时价按批次调价.单据, v_时价按批次调价.No, v_时价按批次调价.序号, v_时价按批次调价.入出类别id, v_时价按批次调价.药品id,
               v_时价按批次调价.批次, v_时价按批次调价.批号, v_时价按批次调价.效期, v_时价按批次调价.产地, v_时价按批次调价.付数, v_时价按批次调价.填写数量, v_时价按批次调价.实际数量,
               Decode(n_实价材料, 1, v_时价按批次调价.原售价, v_时价按批次调价.成本价), v_时价按批次调价.成本金额, v_时价按批次调价.零售价, v_时价按批次调价.扣率,
               v_时价按批次调价.摘要, v_时价按批次调价.填制人, v_时价按批次调价.填制日期, v_时价按批次调价.库房id, v_时价按批次调价.入出系数, v_时价按批次调价.价格id, User, d_审核日期,
               v_时价按批次调价.库存金额, v_时价按批次调价.库存差价);
            n_序号 := n_序号 + 1;
            --处理库存
            --更新库存零售价,只有时价分批药品才能更新零售价字段
            Update 药品库存
            Set 零售价 = Decode(v_时价按批次调价.时价, 1, Decode(Nvl(v_时价按批次调价.批次, 0), 0, Null, v_时价按批次调价.零售价), Null)
            Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(v_时价按批次调价.批次, 0);
          Else
            n_零售价   := v_时价按批次调价.库存金额 / v_时价按批次调价.填写数量;
            n_零售金额 := Round((n_现价 - n_零售价) * v_时价按批次调价.填写数量, 2);
            --产生调价影响记录
            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要,
               填制人, 填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次)
            Values
              (n_收发id, v_时价按批次调价.记录状态, v_时价按批次调价.单据, v_时价按批次调价.No, v_时价按批次调价.序号, v_时价按批次调价.入出类别id, v_时价按批次调价.药品id,
               v_时价按批次调价.批次, v_时价按批次调价.批号, v_时价按批次调价.效期, v_时价按批次调价.产地, v_时价按批次调价.付数, v_时价按批次调价.填写数量, v_时价按批次调价.实际数量,
               Decode(n_实价材料, 1, v_时价按批次调价.原售价, v_时价按批次调价.成本价), v_时价按批次调价.成本金额, v_时价按批次调价.零售价, v_时价按批次调价.扣率, n_零售金额,
               n_零售金额, v_时价按批次调价.摘要, v_时价按批次调价.填制人, v_时价按批次调价.填制日期, v_时价按批次调价.库房id, v_时价按批次调价.入出系数, v_时价按批次调价.价格id, User,
               d_审核日期, v_时价按批次调价.库存金额, v_时价按批次调价.库存差价);
            n_序号 := n_序号 + 1;
            --处理库存
            If v_时价按批次调价.时价 = 1 And Nvl(v_时价按批次调价.批次, 0) > 0 Then
              n_时价分批 := 1;
            Else
              n_时价分批 := 0;
            End If;
          
            If Nvl(v_时价按批次调价.批次, 0) = 0 Then
              Update 药品库存
              Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额
              Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And (批次 Is Null Or 批次 = 0);
            Else
              Update 药品库存
              Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额,
                  零售价 = Decode(n_时价分批, 1, v_时价按批次调价.零售价, 零售价)
              Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And 批次 = v_时价按批次调价.批次;
            End If;
          
            If Sql%RowCount = 0 Then
              Insert Into 药品库存
                (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 零售价)
              Values
                (v_时价按批次调价.库房id, v_时价按批次调价.药品id, v_时价按批次调价.批次, 1, 0, 0, n_零售金额, n_零售金额,
                 Decode(n_时价分批, 1, v_时价按批次调价.零售价, Null));
            End If;
          End If;
        End Loop;
      End Loop;
    End If;
  
    Update 药品收发记录 Set 审核人 = User, 审核日期 = Sysdate Where 价格id = 调价id_In;
    Update 收费价目 Set 变动原因 = 1 Where ID = 调价id_In;
  
    --更新药品目录、收费细目中的变价
    If 定价_In = 1 Then
      Update 收费项目目录 Set 是否变价 = 0 Where ID = n_收费细目id;
    End If;
    --成本价调价
    Zl_材料收发记录_成本价调价(n_收费细目id);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_Adjust;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.150.0007' Where 编号=&n_System;
Commit;
