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
--131243:余伟节,2018-09-17,身份证校验增加台湾地区码
Create Or Replace Function Zl_Fun_Checkidcard
(
  Idcard_In   In Varchar2,
  Calcdate_In In Date := Null
) Return Varchar2 Is
  -------------------------------------------------------------------------------
  --功能：身份证号码合法性校验,并返回身份证号的出生日期、性别、年龄
  --参数说明:
  -- 入参 IDcard_In:身份证号码
  --    Calcdate_In:计算日期,缺省时按系统时间
  -- 返回值：固定格式XML串
  --<OUTPUT>
  --       <BIRTHDAY></BIRTHDAY>                //出生日期
  --       <SEX></SEX>                  //性别
  --       <AGE></AGE>                //年龄
  --     <MSG></MSG>         //空串-身份证号有效(可从身份证号中获取出生日期和性别)，非空串-返回错误信息
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Count     Number(5);
  n_Sum       Number(5);
  v_校验位    Varchar2(50);
  v_Pattern   Varchar2(500);
  v_Err_Msg   Varchar2(2000);
  v_性别      Varchar2(100);
  v_年龄      Varchar2(100);
  d_Curr_Time Date;
  d_出生日期  Date;
  v_Temp      Varchar2(20);

Begin
  Select Sysdate Into d_Curr_Time From Dual;

  If Idcard_In Is Null Then
    v_Err_Msg := '传入身份证号为空!';
    Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
  Else
    --身份证合法验证
    v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    --地区检验
    If Instr(v_Pattern, Substr(Idcard_In, 1, 2)) = 0 Then
      v_Err_Msg := '身份证前两位地区码不正确!';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    --身份证长度检查
    If Length(Idcard_In) = 15 Then
      --检查身份证号:15位身份证号要求全部为数字
      v_Pattern := '^\d{15}$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中包含非法字符，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --获取性别
      If Mod(To_Number(Substr(Idcard_In, 15, 1)), 2) = 1 Then
        v_性别 := '男';
      Else
        v_性别 := '女';
      End If;
      --出生日期的合法性检查
      v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(Idcard_In, 7, 6), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中的出生日期无效，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
        If Instr(',0229,0230,', ',' || Substr(Idcard_In, 9, 4) || ',') > 0 Then
          v_Temp     := '19' || Substr(Idcard_In, 7, 2) || '0301';
          d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
        Else
          d_出生日期 := To_Date('19' || Substr(Idcard_In, 7, 6), 'yyyy-mm-dd');
        End If;
        If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Elsif Length(Idcard_In) = 18 Then
      -- 18 位身份证号前17 位全部为数字，最后1位可为数字或x
      v_Pattern := '^\d{17}[0-9Xx]$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中包含非法字符!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --获取性别
      If Mod(To_Number(Substr(Idcard_In, 17, 1)), 2) = 1 Then
        v_性别 := '男';
      Else
        v_性别 := '女';
      End If;
      --出生日期的合法性检查
      v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(Idcard_In, 7, 8), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中的出生日期无效，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
        If Instr(',0229,0230,', ',' || Substr(Idcard_In, 11, 4) || ',') > 0 Then
          v_Temp     := Substr(Idcard_In, 7, 4) || '0301';
          d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
        Else
          d_出生日期 := To_Date(Substr(Idcard_In, 7, 8), 'yyyy-mm-dd');
        End If;
        If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
        --计算校验位
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
        v_校验位  := Substr(v_Pattern, n_Count + 1, 1);
        If v_校验位 <> Upper(Substr(Idcard_In, 18, 1)) Then
          v_Err_Msg := '身份证号码不正确，请检查。';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Else
      v_Err_Msg := '身份证长度不对,请检查。';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    v_年龄 := Zl_Age_Calc(0, d_出生日期, Calcdate_In);
  End If;

  Return '<OUTPUT><BIRTHDAY>' || To_Char(d_出生日期, 'YYYY-MM-DD') || '</BIRTHDAY><SEX>' || v_性别 || '</SEX><AGE>' || v_年龄 || '</AGE><MSG></MSG></OUTPUT>';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Checkidcard;
/
--131243:余伟节,2018-09-17,身份证校验增加台湾地区码

Create Or Replace Procedure Zl_Third_Buildpatient
(
  Patiinfo_In  In Xmltype,
  Patiinfo_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------
  --参数说明:
  -- 入参 Patiinfo_In:
  --<IN>
  --  <ZJH></ZJH>                 //证件号，目前仅支持身份证号
  --  <ZJLX></ZJLX>                       //证件类型(目前仅支持身份证,为空时默认为身份证)
  --  <XM></XM>                       //姓名
  --  <SJH></SJH>                      //手机号
  --</IN>

  --出参 Patiinfo_Out：
  --<OUTPUT>
  --       <BRID></BRID>                //病人ID
  --       <MZH></MZH>                  //门诊号
  --     <ERROR></ERROR>         //如果有错误返回该节点
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Pati_Id      病人信息.病人id%Type;
  n_Card_Type_Id 医疗卡类别.Id%Type;
  n_Count        Number(5);
  n_Sum          Number(5);
  v_校验位       Varchar2(50);

  v_姓名         病人信息.姓名%Type;
  v_身份证号     病人信息.身份证号%Type;
  v_手机号       病人信息.家庭电话%Type;
  v_性别         病人信息.性别%Type;
  v_年龄         病人信息.年龄%Type;
  v_操作员       人员表.姓名%Type;
  v_医疗付款方式 病人信息.医疗付款方式%Type;
  n_门诊号       病人信息.门诊号%Type;
  v_证件类型     医疗卡类别.名称%Type;
  v_证件号       病人医疗卡信息.卡号%Type;

  v_Pattern Varchar2(500);
  v_Temp    Varchar2(32767); --临时XML
  v_Err_Msg Varchar2(2000);
  n_存在    Number(2);

  d_出生日期  病人信息.出生日期%Type;
  d_Curr_Time Date;

  Err_Item Exception;
Begin
  Patiinfo_Out := Xmltype('<OUTPUT></OUTPUT>');
  Select Sysdate Into d_Curr_Time From Dual;

  --新建病人：姓名、身份证号、手机号（存在家庭电话中）、出生日期、性别、年龄(后面三项可从身份证中获取)。
  Select Extractvalue(Value(I), 'IN/XM'), Extractvalue(Value(I), 'IN/ZJH'), Extractvalue(Value(I), 'IN/SJH'),
         Extractvalue(Value(I), 'IN/ZJLX')
  Into v_姓名, v_证件号, v_手机号, v_证件类型
  From Table(Xmlsequence(Extract(Patiinfo_In, 'IN'))) I;

  Begin
    If v_证件类型 Is Null Then
      Select 病人id
      Into n_Pati_Id
      From 病人医疗卡信息
      Where 卡号 = v_证件号 And 卡类别id In (Select ID From 医疗卡类别 Where 名称 Like '%身份证%') And Rownum < 2;
    Else
      Select 病人id
      Into n_Pati_Id
      From 病人医疗卡信息
      Where 卡号 = v_证件号 And 卡类别id In (Select ID From 医疗卡类别 Where 名称 = v_证件类型) And Rownum < 2;
    End If;
    n_存在 := 1;
  Exception
    When Others Then
      n_存在 := 0;
  End;

  If Nvl(n_存在, 0) = 1 Then
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    Select 门诊号 Into n_门诊号 From 病人信息 Where 病人id = n_Pati_Id;
    If n_门诊号 Is Null Then
      n_门诊号 := Nextno(3);
      Update 病人信息 Set 门诊号 = n_门诊号 Where 病人id = n_Pati_Id;
    End If;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  Else
    If v_姓名 Is Null Then
      v_Err_Msg := '传入姓名为空!';
      Raise Err_Item;
    End If;
    If v_证件类型 Like '%身份证%' Or v_证件类型 Is Null Then
      v_身份证号 := v_证件号;
    Else
      v_Err_Msg := '目前不支持身份证以外的方式建档！';
      Raise Err_Item;
    End If;
  
    If v_身份证号 Is Null Then
      v_Err_Msg := '传入身份证号为空!';
      Raise Err_Item;
    Else
      --身份证合法验证
      v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    
      --地区检验
      If Instr(v_Pattern, Substr(v_身份证号, 1, 2)) = 0 Then
        v_Err_Msg := '身份证前两位地区码不正确!';
        Raise Err_Item;
      End If;
      --身份证长度检查
      If Length(v_身份证号) = 15 Then
        --检查身份证号:15位身份证号要求全部为数字
        v_Pattern := '^\d{15}$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_身份证号, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中包含非法字符，请检查!';
          Raise Err_Item;
        End If;
        --获取性别
        If Mod(To_Number(Substr(v_身份证号, 15, 1)), 2) = 1 Then
          v_性别 := '男';
        Else
          v_性别 := '女';
        End If;
        --出生日期的合法性检查
      
        v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(v_身份证号, 7, 6), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Raise Err_Item;
        Else
          --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
          If Instr(',0229,0230,', ',' || Substr(v_身份证号, 9, 4) || ',') > 0 Then
            v_Temp     := '19' || Substr(v_身份证号, 7, 2) || '0301';
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_出生日期 := To_Date('19' || Substr(v_身份证号, 7, 6), 'yyyy-mm-dd');
          End If;
          If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '身份证中的出生日期无效，请检查!';
            Raise Err_Item;
          End If;
        End If;
      Elsif Length(v_身份证号) = 18 Then
        -- 18 位身份证号前17 位全部为数字，最后1位可为数字或x
        v_Pattern := '^\d{17}[0-9Xx]$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_身份证号, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中包含非法字符!';
          Raise Err_Item;
        End If;
        --获取性别
        If Mod(To_Number(Substr(v_身份证号, 17, 1)), 2) = 1 Then
          v_性别 := '男';
        Else
          v_性别 := '女';
        End If;
        --出生日期的合法性检查
        v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(v_身份证号, 7, 8), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Raise Err_Item;
        Else
          --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
          If Instr(',0229,0230,', ',' || Substr(v_身份证号, 11, 4) || ',') > 0 Then
            v_Temp     := Substr(v_身份证号, 7, 4) || '0301';
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_出生日期 := To_Date(Substr(v_身份证号, 7, 8), 'yyyy-mm-dd');
          End If;
          If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '身份证中的出生日期无效，请检查!';
            Raise Err_Item;
          End If;
          --计算校验位
          n_Sum     := (To_Number(Substr(v_身份证号, 1, 1)) + To_Number(Substr(v_身份证号, 11, 1))) * 7 +
                       (To_Number(Substr(v_身份证号, 2, 1)) + To_Number(Substr(v_身份证号, 12, 1))) * 9 +
                       (To_Number(Substr(v_身份证号, 3, 1)) + To_Number(Substr(v_身份证号, 13, 1))) * 10 +
                       (To_Number(Substr(v_身份证号, 4, 1)) + To_Number(Substr(v_身份证号, 14, 1))) * 5 +
                       (To_Number(Substr(v_身份证号, 5, 1)) + To_Number(Substr(v_身份证号, 15, 1))) * 8 +
                       (To_Number(Substr(v_身份证号, 6, 1)) + To_Number(Substr(v_身份证号, 16, 1))) * 4 +
                       (To_Number(Substr(v_身份证号, 7, 1)) + To_Number(Substr(v_身份证号, 17, 1))) * 2 +
                       To_Number(Substr(v_身份证号, 8, 1)) * 1 + To_Number(Substr(v_身份证号, 9, 1)) * 6 +
                       To_Number(Substr(v_身份证号, 10, 1)) * 3;
          n_Count   := Mod(n_Sum, 11);
          v_Pattern := '10X98765432';
          v_校验位  := Substr(v_Pattern, n_Count + 1, 1);
          If v_校验位 <> Upper(Substr(v_身份证号, 18, 1)) Then
            v_Err_Msg := '身份证号码不正确，请检查。';
            Raise Err_Item;
          End If;
        End If;
      Else
        v_Err_Msg := '身份证长度不对,请检查。';
        Raise Err_Item;
      End If;
    
      If Nvl(v_年龄, '_') = '_' Then
        v_年龄 := Zl_Age_Calc(0, d_出生日期, d_Curr_Time);
      End If;
    End If;
  
    Select 名称 Into v_医疗付款方式 From 医疗付款方式 Where 缺省标志 = 1;
    n_Pati_Id := Nextno(1);
    n_门诊号  := Nextno(3);
    Insert Into 病人信息
      (病人id, 姓名, 身份证号, 家庭电话, 出生日期, 性别, 年龄, 登记时间, 门诊号, 医疗付款方式, 手机号)
      Select n_Pati_Id, v_姓名, v_身份证号, v_手机号, d_出生日期, v_性别, v_年龄, d_Curr_Time, n_门诊号, v_医疗付款方式, v_手机号



      From Dual;
    --病人信息保存完后，完成医疗卡绑定（二代身份证卡类别的绑定）
    Begin
      If v_证件类型 Is Null Then
        Select ID Into n_Card_Type_Id From 医疗卡类别 Where 名称 Like '%身份证%' And Rownum < 2;
      Else
        Select ID Into n_Card_Type_Id From 医疗卡类别 Where 名称 = v_证件类型 And Rownum < 2;
      End If;
    Exception
      When No_Data_Found Then
        v_Err_Msg := '身份证卡类别不存在！';
        Raise Err_Item;
    End;
    Select b.姓名 Into v_操作员 From 上机人员表 A, 人员表 B Where a.人员id = b.Id And a.用户名 = User;
  
    Zl_医疗卡变动_Insert(11, n_Pati_Id, n_Card_Type_Id, Null, v_身份证号, '创建虚拟卡', Null, v_操作员, d_Curr_Time);
  
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
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
--系统版本号
Update zlSystems Set 版本号='10.34.150.0014' Where 编号=&n_System;
Commit;