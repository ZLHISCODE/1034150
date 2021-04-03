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
--127480:焦博,2018-06-20,修正Oracle过程Zl_挂号安排_Autoupdate
Update 挂号安排计划 Set 实际生效 = To_Date('3000-01-01', 'yyyy-mm-dd') Where 生效时间 > Sysdate And 实际生效 < Sysdate;

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--127480:焦博,2018-06-20,修正Oracle过程Zl_挂号安排_Autoupdate
CREATE OR REPLACE Procedure Zl_挂号安排_Autoupdate Is
  Err_Item Exception;
  v_Date Date;
  -- v_Err_Msg Varchar2(100);
  v_Unitscount Number;
Begin
  --n_更新执行人 ：是否更新病人挂号记录 和门诊费用记录中的执行人
  --               如果计划中更改了 挂号项目 则不允许更新 病人挂号记录和门诊费用记录中的数据
  Select Sysdate Into v_Date From Dual;
  Select Count(0) Into v_Unitscount From 合作单位安排控制 Where Rownum = 1;

  For v_生效 In (Select ID, 安排id, 号码, 生效时间, 失效时间, 周日, 周一, 周二, 周三, 周四, 周五, 周六, 分诊方式, 序号控制, 执行时间 As 上次生效时间, 项目id, 医生姓名, 医生id,
                      序号, 科室id, 是否相同
               From (Select a.Id, a.安排id, a.号码, a.生效时间, a.失效时间, a.周日, a.周一, a.周二, a.周三, a.周四, a.周五, a.周六, a.分诊方式, a.序号控制,
                             b.执行时间, a.项目id, a.医生姓名, a.医生id, Nvl(b.执行计划id, 0) As 执行计划id,
                             Row_Number() Over(Partition By a.安排id Order By a.生效时间 Desc) As 顺序号, b.序号, b.科室id,
                             Case
                               When b.项目id = a.项目id And Nvl(a.医生id, 0) = Nvl(b.医生id, 0) And
                                    Nvl(a.医生姓名, '-') = Nvl(b.医生姓名, '-') Then
                                1
                               Else
                                0
                             End As 是否相同
                      From 挂号安排计划 A, 挂号安排 B
                      Where Sysdate Between a.生效时间 And a.失效时间 And a.安排id = b.Id And
                            a.实际生效 >= To_Date('3000-01-01', 'yyyy-mm-dd') And a.生效时间 + 0 <= Sysdate And 审核人 Is Not Null And
                            b.停用日期 Is Null)
               Where 顺序号 = 1 And ID <> Nvl(执行计划id, 0)) Loop
    Update 挂号安排计划
    Set 实际生效 = v_生效.上次生效时间
    Where 安排id = v_生效.安排id And 失效时间 <= v_生效.失效时间 And 生效时间 < Sysdate And ID <> v_生效.Id And
          实际生效 >= To_Date('3000-01-01', 'yyyy-mm-dd');
  
    Update 挂号安排
    Set 周日 = v_生效.周日, 周一 = v_生效.周一, 周二 = v_生效.周二, 周三 = v_生效.周三, 周四 = v_生效.周四, 周五 = v_生效.周五, 周六 = v_生效.周六,
        分诊方式 = v_生效.分诊方式, 序号控制 = v_生效.序号控制, 开始时间 = Sysdate, 终止时间 = v_生效.失效时间, 项目id = Nvl(v_生效.项目id, 项目id), 执行时间 = v_Date,
        执行计划id = v_生效.Id, 序号 = Decode(v_生效.是否相同, 1, 序号, 9999999), 医生姓名 = v_生效.医生姓名, 医生id = v_生效.医生id
    Where ID = v_生效.安排id;
  
    --重新调整序号
    If Nvl(v_生效.是否相同, 0) <> 1 Then
    
      Update 挂号安排 A
      Set 序号 = -1 * 序号
      Where 项目id = v_生效.项目id And a.科室id = v_生效.科室id And Nvl(a.医生姓名, '-') = Nvl(v_生效.医生姓名, '-') And
            Nvl(a.医生id, 0) = Nvl(v_生效.医生id, 0);
      For v_序号 In (Select a.Id, Rownum As 序号
                   From 挂号安排 A
                   Where a.项目id = v_生效.项目id And a.科室id = v_生效.科室id And Nvl(a.医生姓名, '-') = Nvl(v_生效.医生姓名, '-') And
                         Nvl(a.医生id, 0) = Nvl(v_生效.医生id, 0)
                   Order By a.Id) Loop
        Update 挂号安排 A Set 序号 = v_序号.序号 Where ID = v_序号.Id;
      End Loop;
    End If;
    Delete 挂号安排诊室 Where 号表id = v_生效.安排id;
    Insert Into 挂号安排诊室
      (号表id, 门诊诊室)
      Select v_生效.安排id, 门诊诊室 From 挂号计划诊室 Where 计划id = v_生效.Id;
    Delete 挂号安排限制 Where 安排id = v_生效.安排id;
    Insert Into 挂号安排限制
      (安排id, 限制项目, 限号数, 限约数)
      Select v_生效.安排id, 限制项目, 限号数, 限约数 From 挂号计划限制 Where 计划id = v_生效.Id;
    Delete 挂号安排时段 Where 安排id = v_生效.安排id;
    Insert Into 挂号安排时段
      (安排id, 序号, 开始时间, 结束时间, 限制数量, 是否预约, 星期)
      Select v_生效.安排id, 序号, 开始时间, 结束时间, 限制数量, 是否预约, 星期
      From 挂号计划时段
      Where 计划id = v_生效.Id;
    If Nvl(v_Unitscount, 0) > 0 Then
      Delete 合作单位安排控制 Where 安排id = v_生效.安排id;
      Insert Into 合作单位安排控制
        (安排id, 合作单位, 限制项目, 序号, 数量)
        Select v_生效.安排id, 合作单位, 限制项目, 序号, 数量 From 合作单位计划控制 Where 计划id = v_生效.Id;
    End If;
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号安排_Autoupdate;
/

------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.150.0006' Where 编号=&n_System;
Commit;