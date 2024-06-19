<%
'******************************************************************************
'名　称：clsP_Introduce
'概　要：formで飛んできたP_Introduceテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_Introduce
	Public StaffCode
	Public BranchCode
	Public EmployeeCode
	Public BasePay
	Public OvertimeWorkPayAvg
	Public OtherPay
	Public Bonus
	Public AnnualIncome
	Public SituationHourlyPay
	Public CommutationAllowance
	Public HopeIncomeCode
	Public HopeIncomeMin
	Public AnnualSalarySystemFlag
	Public RaiseTypeCode
	Public BonusFlag
	Public BonusCount
	Public BonusMin
	Public SocietyInsuranceFlag
	Public WelfareAnnuityFlag
	Public EmploymentInsuranceFlag
	Public AccidentInsuranceFlag
	Public SelectJobPoint1
	Public SelectJobPoint2
	Public SelectJobPoint3
	Public ForeignCapitalFlag
	Public CapitalMin
	Public EmployeeNumMin
	Public FounderYear
	Public StartYear
	Public StartMonth
	Public LimitYear
	Public LimitMonth
	Public ActiveFlag
	Public CompetitionFlag
	Public MediaCode1
	Public MediaCode2
	Public MediaCode3
	Public MediaOther
	Public Rank
	Public HopeWeekdayMonFlag
	Public HopeWeekdayTueFlag
	Public HopeWeekdayWedFlag
	Public HopeWeekdayThuFlag
	Public HopeWeekdayFriFlag
	Public HopeWeekdaySatFlag
	Public HopeWeekdaySunFlag
	Public HopeWeekdayOther
	Public HopeHourFlag
	Public HopeTimeFrom
	Public HopeTimeTo
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_Introduce クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_BranchCode_Introduce") <> "" Then IsData = True: BranchCode = Request.Form("CONF_BranchCode_Introduce")
		If Request.Form("CONF_EmployeeCode_Introduce") <> "" Then IsData = True: EmployeeCode = Request.Form("CONF_EmployeeCode_Introduce")
		If Request.Form("CONF_BasePay") <> "" Then IsData = True: BasePay = Request.Form("CONF_BasePay") * 10000
		If Request.Form("CONF_OvertimeWorkPayAvg") <> "" Then IsData = True: OvertimeWorkPayAvg = Request.Form("CONF_OvertimeWorkPayAvg") * 10000
		If Request.Form("CONF_OtherPay") <> "" Then IsData = True: OtherPay = Request.Form("CONF_OtherPay") * 10000
		If Request.Form("CONF_Bonus") <> "" Then IsData = True: Bonus = Request.Form("CONF_Bonus") * 10000
		If Request.Form("CONF_AnnualIncome") <> "" Then IsData = True: AnnualIncome = Request.Form("CONF_AnnualIncome") * 10000
		If Request.Form("CONF_SituationHourlyPay") <> "" Then IsData = True: SituationHourlyPay = Request.Form("CONF_SituationHourlyPay")
		If Request.Form("CONF_CommutationAllowance") <> "" Then IsData = True: CommutationAllowance = Request.Form("CONF_CommutationAllowance")
		If Request.Form("CONF_HopeIncomeCode") <> "" Then IsData = True: HopeIncomeCode = Request.Form("CONF_HopeIncomeCode")
		If Request.Form("CONF_HopeIncomeMin") <> "" Then IsData = True: HopeIncomeMin = Request.Form("CONF_HopeIncomeMin") * 10000
		If Request.Form("CONF_AnnualSalarySystemFlag") <> "" Then IsData = True: AnnualSalarySystemFlag = Request.Form("CONF_AnnualSalarySystemFlag")
		If Request.Form("CONF_RaiseTypeCode") <> "" Then IsData = True: RaiseTypeCode = Request.Form("CONF_RaiseTypeCode")
		If Request.Form("CONF_BonusFlag") <> "" Then IsData = True: BonusFlag = Request.Form("CONF_BonusFlag")
		If Request.Form("CONF_BonusCount") <> "" Then IsData = True: BonusCount = Request.Form("CONF_BonusCount")
		If Request.Form("CONF_BonusMin") <> "" Then IsData = True: BonusMin = Request.Form("CONF_BonusMin") * 10000
		If Request.Form("CONF_SocietyInsuranceFlag2") <> "" Then IsData = True: SocietyInsuranceFlag = Request.Form("CONF_SocietyInsuranceFlag2")
		If Request.Form("CONF_WelfareAnnuityFlag") <> "" Then IsData = True: WelfareAnnuityFlag = Request.Form("CONF_WelfareAnnuityFlag")
		If Request.Form("CONF_EmploymentInsuranceFlag") <> "" Then IsData = True: EmploymentInsuranceFlag = Request.Form("CONF_EmploymentInsuranceFlag")
		If Request.Form("CONF_AccidentInsuranceFlag") <> "" Then IsData = True: AccidentInsuranceFlag = Request.Form("CONF_AccidentInsuranceFlag")
		If Request.Form("CONF_SelectJobPoint1") <> "" Then IsData = True: SelectJobPoint1 = Request.Form("CONF_SelectJobPoint1")
		If Request.Form("CONF_SelectJobPoint2") <> "" Then IsData = True: SelectJobPoint2 = Request.Form("CONF_SelectJobPoint2")
		If Request.Form("CONF_SelectJobPoint3") <> "" Then IsData = True: SelectJobPoint3 = Request.Form("CONF_SelectJobPoint3")
		If Request.Form("CONF_ForeignCapitalFlag") <> "" Then IsData = True: ForeignCapitalFlag = Request.Form("CONF_ForeignCapitalFlag")
		If Request.Form("CONF_CapitalMin") <> "" Then IsData = True: CapitalMin = Request.Form("CONF_CapitalMin")
		If Request.Form("CONF_EmployeeNumMin") <> "" Then IsData = True: EmployeeNumMin = Request.Form("CONF_EmployeeNumMin")
		If Request.Form("CONF_FounderYear") <> "" Then IsData = True: FounderYear = Request.Form("CONF_FounderYear")
		If Request.Form("CONF_StartYear") <> "" Then IsData = True: StartYear = Request.Form("CONF_StartYear")
		If Request.Form("CONF_StartMonth") <> "" Then IsData = True: StartMonth = Request.Form("CONF_StartMonth")
		If Request.Form("CONF_LimitYear") <> "" Then IsData = True: LimitYear = Request.Form("CONF_LimitYear")
		If Request.Form("CONF_LimitMonth") <> "" Then IsData = True: LimitMonth = Request.Form("CONF_LimitMonth")
		If Request.Form("CONF_ActiveFlag") <> "" Then IsData = True: ActiveFlag = Request.Form("CONF_ActiveFlag")
		If Request.Form("CONF_CompetitionFlag") <> "" Then IsData = True: CompetitionFlag = Request.Form("CONF_CompetitionFlag")
		If Request.Form("CONF_MediaCode1") <> "" Then IsData = True: MediaCode1 = Request.Form("CONF_MediaCode1")
		If Request.Form("CONF_MediaCode2") <> "" Then IsData = True: MediaCode2 = Request.Form("CONF_MediaCode2")
		If Request.Form("CONF_MediaCode3") <> "" Then IsData = True: MediaCode3 = Request.Form("CONF_MediaCode3")
		If Request.Form("CONF_MediaOther") <> "" Then IsData = True: MediaOther = Request.Form("CONF_MediaOther")
		If Request.Form("CONF_Rank") <> "" Then IsData = True: Rank = Request.Form("CONF_Rank")
		If Request.Form("CONF_HopeWeekdayMonFlag") <> "" Then IsData = True: HopeWeekdayMonFlag = Request.Form("CONF_HopeWeekdayMonFlag")
		If Request.Form("CONF_HopeWeekdayTueFlag") <> "" Then IsData = True: HopeWeekdayTueFlag = Request.Form("CONF_HopeWeekdayTueFlag")
		If Request.Form("CONF_HopeWeekdayWedFlag") <> "" Then IsData = True: HopeWeekdayWedFlag = Request.Form("CONF_HopeWeekdayWedFlag")
		If Request.Form("CONF_HopeWeekdayThuFlag") <> "" Then IsData = True: HopeWeekdayThuFlag = Request.Form("CONF_HopeWeekdayThuFlag")
		If Request.Form("CONF_HopeWeekdayFriFlag") <> "" Then IsData = True: HopeWeekdayFriFlag = Request.Form("CONF_HopeWeekdayFriFlag")
		If Request.Form("CONF_HopeWeekdaySatFlag") <> "" Then IsData = True: HopeWeekdaySatFlag = Request.Form("CONF_HopeWeekdaySatFlag")
		If Request.Form("CONF_HopeWeekdaySunFlag") <> "" Then IsData = True: HopeWeekdaySunFlag = Request.Form("CONF_HopeWeekdaySunFlag")
		If Request.Form("CONF_HopeWeekdayOther") <> "" Then IsData = True: HopeWeekdayOther = Request.Form("CONF_HopeWeekdayOther")
		If Request.Form("CONF_HopeHourFlag") <> "" Then IsData = True: HopeHourFlag = Request.Form("CONF_HopeHourFlag")
		If Request.Form("CONF_HopeTimeFrom") <> "" Then IsData = True: HopeTimeFrom = Request.Form("CONF_HopeTimeFrom")
		If Request.Form("CONF_HopeTimeTo") <> "" Then IsData = True: HopeTimeTo = Request.Form("CONF_HopeTimeTo")

		If BranchCode <> "" And IsRE(BranchCode, "^[A-Z][A-Z]$", True) = False Then Err = Err & "BranchCode" & vbCrLf
		If EmployeeCode <> "" And IsMainCode(EmployeeCode) = False Then Err = Err & "EmployeeCode" & vbCrLf
		If BasePay <> "" And IsNumber(BasePay, 0, False) = False Then Err = Err & "BasePay" & vbCrLf
		If OvertimeWorkPayAvg <> "" And IsNumber(OvertimeWorkPayAvg, 0, False) = False Then Err = Err & "OvertimeWorkPayAvg" & vbCrLf
		If OtherPay <> "" And IsNumber(OtherPay, 0, False) = False Then Err = Err & "OtherPay" & vbCrLf
		If Bonus <> "" And IsNumber(Bonus, 0, False) = False Then Err = Err & "Bonus" & vbCrLf
		If AnnualIncome <> "" And IsNumber(AnnualIncome, 0, False) = False Then Err = Err & "AnnualIncome" & vbCrLf
		If SituationHourlyPay <> "" And IsNumber(SituationHourlyPay, 0, False) = False Then Err = Err & "SituationHourlyPay" & vbCrLf
		If CommutationAllowance <> "" And IsNumber(CommutationAllowance, 0, False) = False Then Err = Err & "CommutationAllowance" & vbCrLf
		If HopeIncomeCode <> "" And IsNumber(HopeIncomeCode, 3, False) = False Then Err = Err & "HopeIncomeCode" & vbCrLf
		If HopeIncomeMin <> "" And IsNumber(HopeIncomeMin, 0, False) = False Then Err = Err & "HopeIncomeMin" & vbCrLf
		If AnnualSalarySystemFlag <> "" And IsFlag(AnnualSalarySystemFlag) = False Then Err = Err & "AnnualSalarySystemFlag" & vbCrLf
		If RaiseTypeCode <> "" And IsNumber(RaiseTypeCode, 3, False) = False Then Err = Err & "RaiseTypeCode" & vbCrLf
		If BonusFlag <> "" And IsNumber(BonusFlag, 0, False) = False Then Err = Err & "BonusFlag" & vbCrLf
		If BonusCount <> "" And IsNumber(BonusCount, 0, False) = False Then Err = Err & "BonusCount" & vbCrLf
		If BonusMin <> "" And IsNumber(BonusMin, 0, False) = False Then Err = Err & "BonusMin" & vbCrLf
		If SocietyInsuranceFlag <> "" And IsFlag(SocietyInsuranceFlag) = False Then Err = Err & "SocietyInsuranceFlag" & vbCrLf
		If WelfareAnnuityFlag <> "" And IsFlag(WelfareAnnuityFlag) = False Then Err = Err & "WelfareAnnuityFlag" & vbCrLf
		If EmploymentInsuranceFlag <> "" And IsFlag(EmploymentInsuranceFlag) = False Then Err = Err & "EmploymentInsuranceFlag" & vbCrLf
		If AccidentInsuranceFlag <> "" And IsFlag(AccidentInsuranceFlag) = False Then Err = Err & "AccidentInsuranceFlag" & vbCrLf
		If ForeignCapitalFlag <> "" And IsFlag(ForeignCapitalFlag) = False Then Err = Err & "ForeignCapitalFlag" & vbCrLf
		If CapitalMin <> "" And IsNumber(CapitalMin, 0, False) = False Then Err = Err & "CapitalMin" & vbCrLf
		If EmployeeNumMin <> "" And IsNumber(EmployeeNumMin, 0, False) = False Then Err = Err & "EmployeeNumMin" & vbCrLf
		If FounderYear <> "" And IsNumber(FounderYear, 4, False) = False Then Err = Err & "FounderYear" & vbCrLf
		If StartYear <> "" And IsNumber(StartYear, 4, False) = False Then Err = Err & "StartYear" & vbCrLf
		If StartMonth <> "" And IsNumber(StartMonth, 2, False) = False Then Err = Err & "StartMonth" & vbCrLf
		If LimitYear <> "" And IsNumber(LimitYear, 4, False) = False Then Err = Err & "LimitYear" & vbCrLf
		If LimitMonth <> "" And IsNumber(LimitMonth, 2, False) = False Then Err = Err & "LimitMonth" & vbCrLf
		If ActiveFlag <> "" And IsFlag(ActiveFlag) = False Then Err = Err & "ActiveFlag" & vbCrLf
		If CompetitionFlag <> "" And IsFlag(CompetitionFlag) = False Then Err = Err & "CompetitionFlag" & vbCrLf
		If MediaCode1 <> "" And IsNumber(MediaCode1, 3, False) = False Then Err = Err & "MediaCode1" & vbCrLf
		If MediaCode2 <> "" And IsNumber(MediaCode2, 3, False) = False Then Err = Err & "MediaCode2" & vbCrLf
		If MediaCode3 <> "" And IsNumber(MediaCode3, 3, False) = False Then Err = Err & "MediaCode3" & vbCrLf
		If Rank <> "" And IsNumber(Rank, 3, False) = False Then Err = Err & "Rank" & vbCrLf
		If HopeWeekdayMonFlag <> "" And IsFlag(HopeWeekdayMonFlag) = False Then Err = Err & "HopeWeekdayMonFlag" & vbCrLf
		If HopeWeekdayTueFlag <> "" And IsFlag(HopeWeekdayTueFlag) = False Then Err = Err & "HopeWeekdayTueFlag" & vbCrLf
		If HopeWeekdayWedFlag <> "" And IsFlag(HopeWeekdayWedFlag) = False Then Err = Err & "HopeWeekdayWedFlag" & vbCrLf
		If HopeWeekdayThuFlag <> "" And IsFlag(HopeWeekdayThuFlag) = False Then Err = Err & "HopeWeekdayThuFlag" & vbCrLf
		If HopeWeekdayFriFlag <> "" And IsFlag(HopeWeekdayFriFlag) = False Then Err = Err & "HopeWeekdayFriFlag" & vbCrLf
		If HopeWeekdaySatFlag <> "" And IsFlag(HopeWeekdaySatFlag) = False Then Err = Err & "HopeWeekdaySatFlag" & vbCrLf
		If HopeWeekdaySunFlag <> "" And IsFlag(HopeWeekdaySunFlag) = False Then Err = Err & "HopeWeekdaySunFlag" & vbCrLf
		If HopeHourFlag <> "" And IsFlag(HopeHourFlag) = False Then Err = Err & "HopeHourFlag" & vbCrLf
		If HopeTimeFrom <> "" And IsNumber(HopeTimeFrom, 4, False) = False Then Err = Err & "HopeTimeFrom" & vbCrLf
		If HopeTimeTo <> "" And IsNumber(HopeTimeTo, 4, False) = False Then Err = Err & "HopeTimeTo" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_Introduce 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_Introduce" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(BranchCode) & "'" & _
			",'" & ChkSQLStr(EmployeeCode) & "'" & _
			",'" & ChkSQLStr(BasePay) & "'" & _
			",'" & ChkSQLStr(OvertimeWorkPayAvg) & "'" & _
			",'" & ChkSQLStr(OtherPay) & "'" & _
			",'" & ChkSQLStr(Bonus) & "'" & _
			",'" & ChkSQLStr(AnnualIncome) & "'" & _
			",'" & ChkSQLStr(SituationHourlyPay) & "'" & _
			",'" & ChkSQLStr(CommutationAllowance) & "'" & _
			",'" & ChkSQLStr(HopeIncomeCode) & "'" & _
			",'" & ChkSQLStr(HopeIncomeMin) & "'" & _
			",'" & ChkSQLStr(AnnualSalarySystemFlag) & "'" & _
			",'" & ChkSQLStr(RaiseTypeCode) & "'" & _
			",'" & ChkSQLStr(BonusFlag) & "'" & _
			",'" & ChkSQLStr(BonusCount) & "'" & _
			",'" & ChkSQLStr(BonusMin) & "'" & _
			",'" & ChkSQLStr(SocietyInsuranceFlag) & "'" & _
			",'" & ChkSQLStr(WelfareAnnuityFlag) & "'" & _
			",'" & ChkSQLStr(EmploymentInsuranceFlag) & "'" & _
			",'" & ChkSQLStr(AccidentInsuranceFlag) & "'" & _
			",'" & ChkSQLStr(SelectJobPoint1) & "'" & _
			",'" & ChkSQLStr(SelectJobPoint2) & "'" & _
			",'" & ChkSQLStr(SelectJobPoint3) & "'" & _
			",'" & ChkSQLStr(ForeignCapitalFlag) & "'" & _
			",'" & ChkSQLStr(CapitalMin) & "'" & _
			",'" & ChkSQLStr(EmployeeNumMin) & "'" & _
			",'" & ChkSQLStr(FounderYear) & "'" & _
			",'" & ChkSQLStr(StartYear) & "'" & _
			",'" & ChkSQLStr(StartMonth) & "'" & _
			",'" & ChkSQLStr(LimitYear) & "'" & _
			",'" & ChkSQLStr(LimitMonth) & "'" & _
			",'" & ChkSQLStr(ActiveFlag) & "'" & _
			",'" & ChkSQLStr(CompetitionFlag) & "'" & _
			",'" & ChkSQLStr(MediaCode1) & "'" & _
			",'" & ChkSQLStr(MediaCode2) & "'" & _
			",'" & ChkSQLStr(MediaCode3) & "'" & _
			",'" & ChkSQLStr(MediaOther) & "'" & _
			",'" & ChkSQLStr(Rank) & "'" & _
			",'" & ChkSQLStr(HopeWeekdayMonFlag) & "'" & _
			",'" & ChkSQLStr(HopeWeekdayTueFlag) & "'" & _
			",'" & ChkSQLStr(HopeWeekdayWedFlag) & "'" & _
			",'" & ChkSQLStr(HopeWeekdayThuFlag) & "'" & _
			",'" & ChkSQLStr(HopeWeekdayFriFlag) & "'" & _
			",'" & ChkSQLStr(HopeWeekdaySatFlag) & "'" & _
			",'" & ChkSQLStr(HopeWeekdaySunFlag) & "'" & _
			",'" & ChkSQLStr(HopeWeekdayOther) & "'" & _
			",'" & ChkSQLStr(HopeHourFlag) & "'" & _
			",'" & ChkSQLStr(HopeTimeFrom) & "'" & _
			",'" & ChkSQLStr(HopeTimeTo) & "'" & vbCrLf
	End Function
End Class
%>
