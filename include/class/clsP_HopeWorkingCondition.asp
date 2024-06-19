<%
'******************************************************************************
'名　称：clsP_HopeWorkingCondition
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_HopeWorkingCondition
	Public StaffCode
	Public YearlyIncomeMin
	Public YearlyIncomeMax
	Public MonthlyIncomeMin
	Public MonthlyIncomeMax
	Public DailyIncomeMin
	Public DailyIncomeMax
	Public HourlyIncomeMin
	Public HourlyIncomeMax
	Public PercentagePayFlag
	Public IncomeRemark
	Public TrafficFeeFlag
	Public SocietyInsuranceFlag
	Public SanatoriumFlag
	Public EnterprisePensionFlag
	Public WealthShapeFlag
	Public StockOptionFlag
	Public RetirementPayFlag
	Public ResidencePayFlag
	Public FamilyPayFlag
	Public EmployeeDormitoryFlag
	Public CompanyHouseFlag
	Public NewEmployeeTrainingFlag
	Public OverseasTrainingFlag
	Public OtherTrainingFlag
	Public FlexTimeFlag
	Public WorkPeriodFlag
	Public WorkMonthPeriod
	Public WorkStartTime
	Public WorkEndTime
	Public WorkShiftFlag
	Public OverWorkFlag
	Public OverWorkTimeMax
	Public OverWorkTimeOther
	Public MonHolidayFlag
	Public TueHolidayFlag
	Public WedHolidayFlag
	Public ThuHolidayFlag
	Public FriHolidayFlag
	Public SatHolidayFlag
	Public SunHolidayFlag
	Public PublicHolidayFlag
	Public WeeklyHolidayType
	Public HolidayRemark
	Public TransferFlag
	Public HopeWorkStartDay
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_HopeWorkingCondition クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_YearlyIncomeMin") <> "" Then IsData = True: YearlyIncomeMin = Request.Form("CONF_YearlyIncomeMin") * 10000
		If Request.Form("CONF_YearlyIncomeMax") <> "" Then IsData = True: YearlyIncomeMax = Request.Form("CONF_YearlyIncomeMax") * 10000
		If Request.Form("CONF_MonthlyIncomeMin") <> "" Then IsData = True: MonthlyIncomeMin = Request.Form("CONF_MonthlyIncomeMin") * 10000
		If Request.Form("CONF_MonthlyIncomeMax") <> "" Then IsData = True: MonthlyIncomeMax = Request.Form("CONF_MonthlyIncomeMax") * 10000
		If Request.Form("CONF_DailyIncomeMin") <> "" Then IsData = True: DailyIncomeMin = Request.Form("CONF_DailyIncomeMin")
		If Request.Form("CONF_DailyIncomeMax") <> "" Then IsData = True: DailyIncomeMax = Request.Form("CONF_DailyIncomeMax")
		If Request.Form("CONF_HourlyIncomeMin") <> "" Then IsData = True: HourlyIncomeMin = Request.Form("CONF_HourlyIncomeMin")
		If Request.Form("CONF_HourlyIncomeMax") <> "" Then IsData = True: HourlyIncomeMax = Request.Form("CONF_HourlyIncomeMax")
		If Request.Form("CONF_PercentagePayFlag") <> "" Then IsData = True: PercentagePayFlag = Request.Form("CONF_PercentagePayFlag")
		If Request.Form("CONF_IncomeRemark") <> "" Then IsData = True: IncomeRemark = Request.Form("CONF_IncomeRemark")
		If Request.Form("CONF_TrafficFeeFlag") <> "" Then IsData = True: TrafficFeeFlag = Request.Form("CONF_TrafficFeeFlag")
		If Request.Form("CONF_SocietyInsuranceFlag") <> "" Then IsData = True: SocietyInsuranceFlag = Request.Form("CONF_SocietyInsuranceFlag")
		If Request.Form("CONF_SanatoriumFlag") <> "" Then IsData = True: SanatoriumFlag = Request.Form("CONF_SanatoriumFlag")
		If Request.Form("CONF_EnterprisePensionFlag") <> "" Then IsData = True: EnterprisePensionFlag = Request.Form("CONF_EnterprisePensionFlag")
		If Request.Form("CONF_WealthShapeFlag") <> "" Then IsData = True: WealthShapeFlag = Request.Form("CONF_WealthShapeFlag")
		If Request.Form("CONF_StockOptionFlag") <> "" Then IsData = True: StockOptionFlag = Request.Form("CONF_StockOptionFlag")
		If Request.Form("CONF_RetirementPayFlag") <> "" Then IsData = True: RetirementPayFlag = Request.Form("CONF_RetirementPayFlag")
		If Request.Form("CONF_ResidencePayFlag") <> "" Then IsData = True: ResidencePayFlag = Request.Form("CONF_ResidencePayFlag")
		If Request.Form("CONF_FamilyPayFlag") <> "" Then IsData = True: FamilyPayFlag = Request.Form("CONF_FamilyPayFlag")
		If Request.Form("CONF_EmployeeDormitoryFlag") <> "" Then IsData = True: EmployeeDormitoryFlag = Request.Form("CONF_EmployeeDormitoryFlag")
		If Request.Form("CONF_CompanyHouseFlag") <> "" Then IsData = True: CompanyHouseFlag = Request.Form("CONF_CompanyHouseFlag")
		If Request.Form("CONF_NewEmployeeTrainingFlag") <> "" Then IsData = True: NewEmployeeTrainingFlag = Request.Form("CONF_NewEmployeeTrainingFlag")
		If Request.Form("CONF_OverseasTrainingFlag") <> "" Then IsData = True: OverseasTrainingFlag = Request.Form("CONF_OverseasTrainingFlag")
		If Request.Form("CONF_OtherTrainingFlag") <> "" Then IsData = True: OtherTrainingFlag = Request.Form("CONF_OtherTrainingFlag")
		If Request.Form("CONF_FlexTimeFlag") <> "" Then IsData = True: FlexTimeFlag = Request.Form("CONF_FlexTimeFlag")
		If Request.Form("CONF_WorkPeriodTypeFlag") <> "" Then IsData = True: WorkPeriodFlag = Request.Form("CONF_WorkPeriodTypeFlag")
		If Request.Form("CONF_HopeMonthPeriod") <> "" Then IsData = True: WorkMonthPeriod = Request.Form("CONF_HopeMonthPeriod")
		If Request.Form("CONF_WorkStartTime") <> "" Then IsData = True: WorkStartTime = Request.Form("CONF_WorkStartTime")
		If Request.Form("CONF_WorkEndTime") <> "" Then IsData = True: WorkEndTime = Request.Form("CONF_WorkEndTime")
		If Request.Form("CONF_WorkShiftFlag") <> "" Then IsData = True: WorkShiftFlag = Request.Form("CONF_WorkShiftFlag")
		If Request.Form("CONF_OverWorkFlag") <> "" Then IsData = True: OverWorkFlag = Request.Form("CONF_OverWorkFlag")
		If Request.Form("CONF_OverWorkTimeMax") <> "" Then IsData = True: OverWorkTimeMax = Request.Form("CONF_OverWorkTimeMax")
		If Request.Form("CONF_OverWorkTimeOther") <> "" Then IsData = True: OverWorkTimeOther = Request.Form("CONF_OverWorkTimeOther")
		If Request.Form("CONF_MonHolidayFlag") <> "" Then IsData = True: MonHolidayFlag = Request.Form("CONF_MonHolidayFlag")
		If Request.Form("CONF_TueHolidayFlag") <> "" Then IsData = True: TueHolidayFlag = Request.Form("CONF_TueHolidayFlag")
		If Request.Form("CONF_WedHolidayFlag") <> "" Then IsData = True: WedHolidayFlag = Request.Form("CONF_WedHolidayFlag")
		If Request.Form("CONF_ThuHolidayFlag") <> "" Then IsData = True: ThuHolidayFlag = Request.Form("CONF_ThuHolidayFlag")
		If Request.Form("CONF_FriHolidayFlag") <> "" Then IsData = True: FriHolidayFlag = Request.Form("CONF_FriHolidayFlag")
		If Request.Form("CONF_SatHolidayFlag") <> "" Then IsData = True: SatHolidayFlag = Request.Form("CONF_SatHolidayFlag")
		If Request.Form("CONF_SunHolidayFlag") <> "" Then IsData = True: SunHolidayFlag = Request.Form("CONF_SunHolidayFlag")
		If Request.Form("CONF_PublicHolidayFlag") <> "" Then IsData = True: PublicHolidayFlag = Request.Form("CONF_PublicHolidayFlag")
		If Request.Form("CONF_WeeklyHolidayType") <> "" Then IsData = True: WeeklyHolidayType = Request.Form("CONF_WeeklyHolidayType")
		If Request.Form("CONF_HolidayRemark") <> "" Then IsData = True: HolidayRemark = Request.Form("CONF_HolidayRemark")
		If Request.Form("CONF_TransferFlag") <> "" Then IsData = True: TransferFlag = Request.Form("CONF_TransferFlag")
		If Request.Form("CONF_HopeWorkStartDay") <> "" Then IsData = True: HopeWorkStartDay = Request.Form("CONF_HopeWorkStartDay")

		Err = ""
		If YearlyIncomeMin <> "" And IsNumber(YearlyIncomeMin, 0, False) = False Then Err = Err & "YearlyIncomeMin" & vbCrLf
		If YearlyIncomeMax <> "" And IsNumber(YearlyIncomeMax, 0, False) = False Then Err = Err & "YearlyIncomeMax" & vbCrLf
		If MonthlyIncomeMin <> "" And IsNumber(MonthlyIncomeMin, 0, False) = False Then Err = Err & "MonthlyIncomeMin" & vbCrLf
		If MonthlyIncomeMax <> "" And IsNumber(MonthlyIncomeMax, 0, False) = False Then Err = Err & "MonthlyIncomeMax" & vbCrLf
		If DailyIncomeMin <> "" And IsNumber(DailyIncomeMin, 0, False) = False Then Err = Err & "DailyIncomeMin" & vbCrLf
		If DailyIncomeMax <> "" And IsNumber(DailyIncomeMax, 0, False) = False Then Err = Err & "DailyIncomeMax" & vbCrLf
		If HourlyIncomeMin <> "" And IsNumber(HourlyIncomeMin, 0, False) = False Then Err = Err & "HourlyIncomeMin" & vbCrLf
		If HourlyIncomeMax <> "" And IsNumber(HourlyIncomeMax, 0, False) = False Then Err = Err & "HourlyIncomeMax" & vbCrLf
		If PercentagePayFlag <> "" And IsFlag(PercentagePayFlag) = False Then Err = Err & "PercentagePayFlag" & vbCrLf
		If TrafficFeeFlag <> "" And IsFlag(TrafficFeeFlag) = False Then Err = Err & "TrafficFeeFlag" & vbCrLf
		If SocietyInsuranceFlag <> "" And IsFlag(SocietyInsuranceFlag) = False Then Err = Err & "SocietyInsuranceFlag" & vbCrLf
		If SanatoriumFlag <> "" And IsFlag(SanatoriumFlag) = False Then Err = Err & "SanatoriumFlag" & vbCrLf
		If EnterprisePensionFlag <> "" And IsFlag(EnterprisePensionFlag) = False Then Err = Err & "EnterprisePensionFlag" & vbCrLf
		If WealthShapeFlag <> "" And IsFlag(WealthShapeFlag) = False Then Err = Err & "WealthShapeFlag" & vbCrLf
		If StockOptionFlag <> "" And IsFlag(StockOptionFlag) = False Then Err = Err & "StockOptionFlag" & vbCrLf
		If RetirementPayFlag <> "" And IsFlag(RetirementPayFlag) = False Then Err = Err & "RetirementPayFlag" & vbCrLf
		If ResidencePayFlag <> "" And IsFlag(ResidencePayFlag) = False Then Err = Err & "ResidencePayFlag" & vbCrLf
		If FamilyPayFlag <> "" And IsFlag(FamilyPayFlag) = False Then Err = Err & "FamilyPayFlag" & vbCrLf
		If EmployeeDormitoryFlag <> "" And IsFlag(EmployeeDormitoryFlag) = False Then Err = Err & "EmployeeDormitoryFlag" & vbCrLf
		If CompanyHouseFlag <> "" And IsFlag(CompanyHouseFlag) = False Then Err = Err & "CompanyHouseFlag" & vbCrLf
		If NewEmployeeTrainingFlag <> "" And IsFlag(NewEmployeeTrainingFlag) = False Then Err = Err & "NewEmployeeTrainingFlag" & vbCrLf
		If OverseasTrainingFlag <> "" And IsFlag(OverseasTrainingFlag) = False Then Err = Err & "OverseasTrainingFlag" & vbCrLf
		If OtherTrainingFlag <> "" And IsFlag(OtherTrainingFlag) = False Then Err = Err & "OtherTrainingFlag" & vbCrLf
		If FlexTimeFlag <> "" And IsFlag(FlexTimeFlag) = False Then Err = Err & "FlexTimeFlag" & vbCrLf
		If WorkPeriodFlag <> "" And IsFlag(WorkPeriodFlag) = False Then Err = Err & "WorkPeriodFlag" & vbCrLf
		If WorkStartTime <> "" And IsNumber(WorkStartTime, 4, False) = False Then Err = Err & "WorkStartTime" & vbCrLf
		If WorkEndTime <> "" And IsNumber(WorkEndTime, 4, False) = False Then Err = Err & "WorkEndTime" & vbCrLf
		If WorkShiftFlag <> "" And IsFlag(WorkShiftFlag) = False Then Err = Err & "WorkShiftFlag" & vbCrLf
		If OverWorkFlag <> "" And IsFlag(OverWorkFlag) = False Then Err = Err & "OverWorkFlag" & vbCrLf
		If OverWorkTimeMax <> "" And IsNumber(OverWorkTimeMax, 4, False) = False Then Err = Err & "OverWorkTimeMax" & vbCrLf
		If MonHolidayFlag <> "" And IsFlag(MonHolidayFlag) = False Then Err = Err & "MonHolidayFlag" & vbCrLf
		If TueHolidayFlag <> "" And IsFlag(TueHolidayFlag) = False Then Err = Err & "TueHolidayFlag" & vbCrLf
		If WedHolidayFlag <> "" And IsFlag(WedHolidayFlag) = False Then Err = Err & "WedHolidayFlag" & vbCrLf
		If ThuHolidayFlag <> "" And IsFlag(ThuHolidayFlag) = False Then Err = Err & "ThuHolidayFlag" & vbCrLf
		If FriHolidayFlag <> "" And IsFlag(FriHolidayFlag) = False Then Err = Err & "FriHolidayFlag" & vbCrLf
		If SatHolidayFlag <> "" And IsFlag(SatHolidayFlag) = False Then Err = Err & "SatHolidayFlag" & vbCrLf
		If SunHolidayFlag <> "" And IsFlag(SunHolidayFlag) = False Then Err = Err & "SunHolidayFlag" & vbCrLf
		If PublicHolidayFlag <> "" And IsFlag(PublicHolidayFlag) = False Then Err = Err & "PublicHolidayFlag" & vbCrLf
		If WeeklyHolidayType <> "" And IsNumber(WeeklyHolidayType, 3, True) = False Then Err = Err & "WeeklyHolidayType" & vbCrLf
		If TransferFlag <> "" And IsFlag(TransferFlag) = False Then Err = Err & "TransferFlag" & vbCrLf
		If HopeWorkStartDay <> "" And IsDay(HopeWorkStartDay) = False Then Err = Err & "HopeWorkStartDay" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopeWorkingCondition 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_HopeWorkingCondition" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(YearlyIncomeMin) & "'" & _
			",'" & ChkSQLStr(YearlyIncomeMax) & "'" & _
			",'" & ChkSQLStr(MonthlyIncomeMin) & "'" & _
			",'" & ChkSQLStr(MonthlyIncomeMax) & "'" & _
			",'" & ChkSQLStr(DailyIncomeMin) & "'" & _
			",'" & ChkSQLStr(DailyIncomeMax) & "'" & _
			",'" & ChkSQLStr(HourlyIncomeMin) & "'" & _
			",'" & ChkSQLStr(HourlyIncomeMax) & "'" & _
			",'" & ChkSQLStr(PercentagePayFlag) & "'" & _
			",'" & ChkSQLStr(IncomeRemark) & "'" & _
			",'" & ChkSQLStr(TrafficFeeFlag) & "'" & _
			",'" & ChkSQLStr(SocietyInsuranceFlag) & "'" & _
			",'" & ChkSQLStr(SanatoriumFlag) & "'" & _
			",'" & ChkSQLStr(EnterprisePensionFlag) & "'" & _
			",'" & ChkSQLStr(WealthShapeFlag) & "'" & _
			",'" & ChkSQLStr(StockOptionFlag) & "'" & _
			",'" & ChkSQLStr(RetirementPayFlag) & "'" & _
			",'" & ChkSQLStr(ResidencePayFlag) & "'" & _
			",'" & ChkSQLStr(FamilyPayFlag) & "'" & _
			",'" & ChkSQLStr(EmployeeDormitoryFlag) & "'" & _
			",'" & ChkSQLStr(CompanyHouseFlag) & "'" & _
			",'" & ChkSQLStr(NewEmployeeTrainingFlag) & "'" & _
			",'" & ChkSQLStr(OverseasTrainingFlag) & "'" & _
			",'" & ChkSQLStr(OtherTrainingFlag) & "'" & _
			",'" & ChkSQLStr(FlexTimeFlag) & "'" & _
			",'" & ChkSQLStr(WorkPeriodFlag) & "'" & _
			",'" & ChkSQLStr(WorkMonthPeriod) & "'" & _
			",'" & ChkSQLStr(WorkStartTime) & "'" & _
			",'" & ChkSQLStr(WorkEndTime) & "'" & _
			",'" & ChkSQLStr(WorkShiftFlag) & "'" & _
			",'" & ChkSQLStr(OverWorkFlag) & "'" & _
			",'" & ChkSQLStr(OverWorkTimeMax) & "'" & _
			",'" & ChkSQLStr(OverWorkTimeOther) & "'" & _
			",'" & ChkSQLStr(MonHolidayFlag) & "'" & _
			",'" & ChkSQLStr(TueHolidayFlag) & "'" & _
			",'" & ChkSQLStr(WedHolidayFlag) & "'" & _
			",'" & ChkSQLStr(ThuHolidayFlag) & "'" & _
			",'" & ChkSQLStr(FriHolidayFlag) & "'" & _
			",'" & ChkSQLStr(SatHolidayFlag) & "'" & _
			",'" & ChkSQLStr(SunHolidayFlag) & "'" & _
			",'" & ChkSQLStr(PublicHolidayFlag) & "'" & _
			",'" & ChkSQLStr(WeeklyHolidayType) & "'" & _
			",'" & ChkSQLStr(HolidayRemark) & "'" & _
			",'" & ChkSQLStr(TransferFlag) & "'" & _
			",'" & ChkSQLStr(HopeWorkStartDay) & "'" & vbCrLf
	End Function
End Class
%>
