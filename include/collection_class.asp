<%
'******************************************************************************
'概　要：受注テーブル群にデータをInsert, Updateする時に
'　　　：formで飛んできたデータを格納するためのクラス群
'備　考：事前に commonfunc.asp をインクルードしておくこと！
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
%>
<%
'******************************************************************************
'名　称：clsC_Info
'概　要：formで飛んできたC_Infoテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_Info
	Public OrderCode
	Public CompanyCode
	Public PublicFlag
	Public PublicDay
	Public PublicLimitDay
	Public RecruitmentLimitDay
	Public CompetitionFlag
	Public CompetitionRemark
	Public ClientClassFlag
	Public ClientClassRemark
	Public OrderConditionFlag
	Public OrderConditionRemark
	Public OrderType
	Public OrderProgressType
	Public BranchCode
	Public EmployeeCode
	Public CoordinatorCode
	Public JobTypeDetail
	Public BusinessDetail
	Public HopeSchoolHistoryCode
	Public AgeMin
	Public AgeMax
	Public AgeReasonFlag
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
	Public WorkTimeRemark
	Public WeeklyHolidayType
	Public WorkHolidayRemark
	Public UniformFlag
	Public UniformSize
	Public LockerFlag
	Public EmployeeRestaurantFlag
	Public BoardFlag
	Public SmokingFlag
	Public SmokingAreaFlag
	Public DutySystemFlag
	Public DutyType
	Public DutyTimeFlag
	Public WorkingPlaceCompanyName
	Public WorkingPlaceSection
	Public WorkingPlaceTelephoneNumber
	Public WorkingPlaceChargePersonPost
	Public WorkingPlaceChargePersonName
	Public WorkingPlaceArea
	Public WorkingPlacePrefectureCode
	Public WorkingPlaceCity
	Public WorkingPlaceTown
	Public WorkingPlaceAddress
	Public TransferFlag
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Infoクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		CompanyCode = GetForm("CONF_CompanyCode", 1)
		PublicFlag = GetForm("CONF_PublicFlag", 1)
		PublicDay = GetForm("CONF_PublicDay", 1)
		PublicLimitDay = GetForm("CONF_PublicLimitDay", 1)
		RecruitmentLimitDay = GetForm("CONF_RecruitmentLimitDay", 1)
		CompetitionFlag = GetForm("CONF_CompetitionFlag", 1)
		CompetitionRemark = GetForm("CONF_CompetitionRemark", 1)
		ClientClassFlag = GetForm("CONF_ClientClassFlag", 1)
		ClientClassRemark = GetForm("CONF_ClientClassRemark", 1)
		OrderConditionFlag = GetForm("CONF_OrderConditionFlag", 1)
		OrderConditionRemark = GetForm("CONF_OrderConditionRemark", 1)
		AccessCount = GetForm("CONF_AccessCount", 1)
		OrderType = GetForm("CONF_OrderType", 1)
		OrderProgressType = GetForm("CONF_OrderProgressType", 1)
		BranchCode = GetForm("CONF_BranchCode", 1)
		EmployeeCode = GetForm("CONF_EmployeeCode", 1)
		CoordinatorCode = GetForm("CONF_CoordinatorCode", 1)
		JobTypeDetail = GetForm("CONF_JobTypeDetail", 1)
		BusinessDetail = GetForm("CONF_BusinessDetail", 1)
		HopeSchoolHistoryCode = GetForm("CONF_HopeSchoolHistoryCode", 1)
		AgeMin = GetForm("CONF_AgeMin", 1)
		AgeMax = GetForm("CONF_AgeMax", 1)
		AgeReasonFlag = GetForm("CONF_AgeReasonFlag", 1)
		YearlyIncomeMin = GetForm("CONF_YearlyIncomeMin", 1)
		YearlyIncomeMax = GetForm("CONF_YearlyIncomeMax", 1)
		MonthlyIncomeMin = GetForm("CONF_MonthlyIncomeMin", 1)
		MonthlyIncomeMax = GetForm("CONF_MonthlyIncomeMax", 1)
		DailyIncomeMin = GetForm("CONF_DailyIncomeMin", 1)
		DailyIncomeMax = GetForm("CONF_DailyIncomeMax", 1)
		HourlyIncomeMin = GetForm("CONF_HourlyIncomeMin", 1)
		HourlyIncomeMax = GetForm("CONF_HourlyIncomeMax", 1)
		PercentagePayFlag = GetForm("CONF_PercentagePayFlag", 1)
		IncomeRemark = GetForm("CONF_IncomeRemark", 1)
		WorkTimeRemark = GetForm("CONF_WorkTimeRemark", 1)
		WeeklyHolidayType = GetForm("CONF_WeeklyHolidayType", 1)
		WorkHolidayRemark = GetForm("CONF_WorkHolidayRemark", 1)
		UniformFlag = GetForm("CONF_UniformFlag", 1)
		UniformSize = GetForm("CONF_UniformSize", 1)
		LockerFlag = GetForm("CONF_LockerFlag", 1)
		EmployeeRestaurantFlag = GetForm("CONF_EmployeeRestaurantFlag", 1)
		BoardFlag = GetForm("CONF_BoardFlag", 1)
		SmokingFlag = GetForm("CONF_SmokingFlag", 1)
		SmokingAreaFlag = GetForm("CONF_SmokingAreaFlag", 1)
		DutySystemFlag = GetForm("CONF_DutySystemFlag", 1)
		DutyType = GetForm("CONF_DutyType", 1)
		DutyTimeFlag = GetForm("CONF_DutyTimeFlag", 1)
		WorkingPlaceCompanyName = GetForm("CONF_WorkingPlaceCompanyName", 1)
		WorkingPlaceSection = GetForm("CONF_WorkingPlaceSection", 1)
		WorkingPlaceTelephoneNumber = GetForm("CONF_WorkingPlaceTelephone", 1)
		WorkingPlaceChargePersonPost = GetForm("CONF_WorkingPlaceChargePersonPost", 1)
		WorkingPlaceChargePersonName = GetForm("CONF_WorkingPlaceChargePersonName", 1)
		WorkingPlaceArea = GetForm("CONF_WorkingPlaceArea", 1)
		WorkingPlacePrefectureCode = GetForm("CONF_WorkingPlacePrefectureCode", 1)
		WorkingPlaceCity = GetForm("CONF_WorkingPlaceCity", 1)
		WorkingPlaceTown = GetForm("CONF_WorkingPlaceTown", 1)
		WorkingPlaceAddress = GetForm("CONF_WorkingPlaceAddress", 1)
		TransferFlag = GetForm("CONF_TransferFlag", 1)

		IsData = False
		If CompanyCode <> "" Then IsData = True
		If PublicFlag <> "" Then IsData = True
		If PublicDay <> "" Then IsData = True
		If PublicLimitDay <> "" Then IsData = True
		If RecruitmentLimitDay <> "" Then IsData = True
		If CompetitionFlag <> "" Then IsData = True
		If CompetitionRemark <> "" Then IsData = True
		If ClientClassFlag <> "" Then IsData = True
		If ClientClassRemark <> "" Then IsData = True
		If OrderConditionFlag <> "" Then IsData = True
		If OrderConditionRemark <> "" Then IsData = True
		If OrderType <> "" Then IsData = True
		If OrderProgressType <> "" Then IsData = True
		If BranchCode <> "" Then IsData = True
		If EmployeeCode <> "" Then IsData = True
		If CoordinatorCode <> "" Then IsData = True
		If JobTypeDetail <> "" Then IsData = True
		If BusinessDetail <> "" Then IsData = True
		If HopeSchoolHistoryCode <> "" Then IsData = True
		If AgeMin <> "" Then IsData = True
		If AgeMax <> "" Then IsData = True
		If AgeReasonFlag <> "" Then IsData = True
		If YearlyIncomeMin <> "" Then IsData = True
		If YearlyIncomeMax <> "" Then IsData = True
		If MonthlyIncomeMin <> "" Then IsData = True
		If MonthlyIncomeMax <> "" Then IsData = True
		If DailyIncomeMin <> "" Then IsData = True
		If DailyIncomeMax <> "" Then IsData = True
		If HourlyIncomeMin <> "" Then IsData = True
		If HourlyIncomeMax <> "" Then IsData = True
		If PercentagePayFlag <> "" Then IsData = True
		If IncomeRemark <> "" Then IsData = True
		If WorkTimeRemark <> "" Then IsData = True
		If WeeklyHolidayType <> "" Then IsData = True
		If WorkHolidayRemark <> "" Then IsData = True
		If UniformFlag <> "" Then IsData = True
		If UniformSize <> "" Then IsData = True
		If LockerFlag <> "" Then IsData = True
		If EmployeeRestaurantFlag <> "" Then IsData = True
		If BoardFlag <> "" Then IsData = True
		If SmokingFlag <> "" Then IsData = True
		If SmokingAreaFlag <> "" Then IsData = True
		If DutySystemFlag <> "" Then IsData = True
		If DutyType <> "" Then IsData = True
		If DutyTimeFlag <> "" Then IsData = True
		If WorkingPlaceCompanyName <> "" Then IsData = True
		If WorkingPlaceSection <> "" Then IsData = True
		If WorkingPlaceTelephoneNumber <> "" Then IsData = True
		If WorkingPlaceChargePersonPost <> "" Then IsData = True
		If WorkingPlaceChargePersonName <> "" Then IsData = True
		If WorkingPlaceArea <> "" Then IsData = True
		If WorkingPlacePrefectureCode <> "" Then IsData = True
		If WorkingPlaceCity <> "" Then IsData = True
		If WorkingPlaceTown <> "" Then IsData = True
		If WorkingPlaceAddress <> "" Then IsData = True
		If TransferFlag <> "" Then IsData = True

		'値チェック
		Err = ""
		If IsMainCode(CompanyCode) = False Then Err = Err & "CompanyCode" & vbCrLf
		If IsFlag(PublicFlag) = False Then Err = Err & "PublicFlag" & vbCrLf
		If IsDay(PublicDay) = False Then Err = Err & "PublicDay" & vbCrLf
		If IsDay(PublicLimitDay) = False Then Err = Err & "PublicLimitDay" & vbCrLf
		If IsDay(RecruitmentLimitDay) = False Then Err = Err & "RecruitmentLimitDay" & vbCrLf
		If IsFlag(CompetitionFlag) = False Then Err = Err & "CompetitionFlag" & vbCrLf
		If IsRE(ClientClassFlag, "^[123]$", True) = False Then Err = Err & "ClientClassFlag" & vbCrLf
		If IsRE(OrderConditionFlag, "^[12]$", True) = False Then Err = Err & "OrderConditionFlag" & vbCrLf
		If IsRE(OrderType, "^[01234]$", True) = False Then Err = Err & "OrderType" & vbCrLf
		If IsRE(OrderProgressType, "^[12]$", True) = False Then Err = Err & "OrderProgressType" & vbCrLf
		If IsRE(BranchCode, "^[A-Z][A-Z]$", True) = False Then Err = Err & "BranchCode" & vbCrLf
		If IsMainCode(EmployeeCode) = False Then Err = Err & "EmployeeCode" & vbCrLf
		If IsMainCode(CoordinatorCode) = False Then Err = Err & "CoordinatorCode" & vbCrLf
		If IsNumber(HopeSchoolHistoryCode, 3, False) = False Then Err = Err & "HopeSchoolHistoryCode" & vbCrLf
		If IsNumber(AgeMin, 2, False) = False Then Err = Err & "AgeMin" & vbCrLf
		If IsNumber(AgeMax, 2, False) = False Then Err = Err & "AgeMax" & vbCrLf
		If IsRE(AgeReasonFlag, "^[A-Z]$", True) = False Then Err = Err & "AgeReasonFlag" & vbCrLf
		If IsNumber(YearlyIncomeMin, 0, True) = False Then Err = Err & "YearlyIncomeMin" & vbCrLf
		If IsNumber(YearlyIncomeMax, 0, True) = False Then Err = Err & "YearlyIncomeMax" & vbCrLf
		If IsNumber(MonthlyIncomeMin, 0, True) = False Then Err = Err & "MonthlyIncomeMin" & vbCrLf
		If IsNumber(MonthlyIncomeMax, 0, True) = False Then Err = Err & "MonthlyIncomeMax" & vbCrLf
		If IsNumber(DailyIncomeMin, 0, True) = False Then Err = Err & "DailyIncomeMin" & vbCrLf
		If IsNumber(DailyIncomeMax, 0, True) = False Then Err = Err & "DailyIncomeMax" & vbCrLf
		If IsNumber(HourlyIncomeMin, 0, True) = False Then Err = Err & "HourlyIncomeMin" & vbCrLf
		If IsNumber(HourlyIncomeMax, 0, True) = False Then Err = Err & "HourlyIncomeMax" & vbCrLf
		If IsFlag(PercentagePayFlag) = False Then Err = Err & "PercentagePayFlag" & vbCrLf
		If IsNumber(WeeklyHolidayType, 3, False) = False Then Err = Err & "WeeklyHolidayType" & vbCrLf
		If IsFlag(UniformFlag) = False Then Err = Err & "UniformFlag" & vbCrLf
		If IsFlag(LockerFlag) = False Then Err = Err & "LockerFlag" & vbCrLf
		If IsFlag(EmployeeRestaurantFlag) = False Then Err = Err & "EmployeeRestaurantFlag" & vbCrLf
		If IsFlag(BoardFlag) = False Then Err = Err & "BoardFlag" & vbCrLf
		If IsFlag(SmokingFlag) = False Then Err = Err & "SmokingFlag" & vbCrLf
		If IsFlag(SmokingAreaFlag) = False Then Err = Err & "SmokingAreaFlag" & vbCrLf
		If IsFlag(DutySystemFlag) = False Then Err = Err & "DutySystemFlag" & vbCrLf
		If IsRE(DutyType, "^[123]$", True) = False Then Err = Err & "DutyType" & vbCrLf
		If IsRE(DutyTimeFlag, "^[123]$", True) = False Then Err = Err & "DutyTimeFlag" & vbCrLf
		If IsNumber(Replace(WorkingPlaceTelephoneNumber, "-", ""), 0, False) = False Then Err = Err & "WorkingPlaceTelephoneNumber" & vbCrLf
		If IsNumber(Replace(WorkingPlaceChargePersonPost, "-", ""), 0, False) = False Then Err = Err & "WorkingPlaceChargePersonPost" & vbCrLf
		If IsNumber(Replace(WorkingPlaceChargePersonName, "-", ""), 0, False) = False Then Err = Err & "WorkingPlaceChargePersonName" & vbCrLf
		If IsNumber(WorkingPlacePrefectureCode, 3, False) = False Then Err = Err & "WorkingPlacePrefectureCode" & vbCrLf
		If IsFlag(TransferFlag) = False Then Err = Err & "TransferFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Info 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_Info '" & vOrderCode & "'" & _
				",'" & CompanyCode & "'" & _
				",'" & PublicFlag & "'" & _
				",'" & PublicDay & "'" & _
				",'" & PublicLimitDay & "'" & _
				",'" & RecruitmentLimitDay & "'" & _
				",'" & CompetitionFlag & "'" & _
				",'" & CompetitionRemark & "'" & _
				",'" & ClientClassFlag & "'" & _
				",'" & ClientClassRemark & "'" & _
				",'" & OrderConditionFlag & "'" & _
				",'" & OrderConditionRemark & "'" & _
				",'" & OrderType & "'" & _
				",'" & OrderProgressType & "'" & _
				",'" & BranchCode & "'" & _
				",'" & EmployeeCode & "'" & _
				",'" & CoordinatorCode & "'" & _
				",'" & JobTypeDetail & "'" & _
				",'" & BusinessDetail & "'" & _
				",'" & HopeSchoolHistoryCode & "'" & _
				",'" & AgeMin & "'" & _
				",'" & AgeMax & "'" & _
				",'" & AgeReasonFlag & "'" & _
				",'" & YearlyIncomeMin & "'" & _
				",'" & YearlyIncomeMax & "'" & _
				",'" & MonthlyIncomeMin & "'" & _
				",'" & MonthlyIncomeMax & "'" & _
				",'" & DailyIncomeMin & "'" & _
				",'" & DailyIncomeMax & "'" & _
				",'" & HourlyIncomeMin & "'" & _
				",'" & HourlyIncomeMax & "'" & _
				",'" & PercentagePayFlag & "'" & _
				",'" & IncomeRemark & "'" & _
				",'" & WorkTimeRemark & "'" & _
				",'" & WeeklyHolidayType & "'" & _
				",'" & WorkHolidayRemark & "'" & _
				",'" & UniformFlag & "'" & _
				",'" & UniformSize & "'" & _
				",'" & LockerFlag & "'" & _
				",'" & EmployeeRestaurantFlag & "'" & _
				",'" & BoardFlag & "'" & _
				",'" & SmokingFlag & "'" & _
				",'" & SmokingAreaFlag & "'" & _
				",'" & DutySystemFlag & "'" & _
				",'" & DutyType & "'" & _
				",'" & DutyTimeFlag & "'" & _
				",'" & WorkingPlaceCompanyName & "'" & _
				",'" & WorkingPlaceSection & "'" & _
				",'" & WorkingPlaceTelephoneNumber & "'" & _
				",'" & WorkingPlaceChargePersonPost & "'" & _
				",'" & WorkingPlaceChargePersonName & "'" & _
				",'" & WorkingPlaceArea & "'" & _
				",'" & WorkingPlacePrefectureCode & "'" & _
				",'" & WorkingPlaceCity & "'" & _
				",'" & WorkingPlaceTown & "'" & _
				",'" & WorkingPlaceAddress & "'" & _
				",'" & TransferFlag & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_Temp
'概　要：formで飛んできたC_Tempテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_Temp
	Public OrderCode
	Public SelectionPoint
	Public NewGraduateFlag
	Public HourlyPayMin
	Public HourlyPayMax
	Public SalaryPayUnit
	Public OvertimeWorkPayMin
	Public OvertimeWorkPayMax
	Public HolidayHourlyPayMin
	Public HolidayHourlyPayMax
	Public TrafficPayFlag
	Public TrafficPayRemark
	Public WorkStartDay
	Public WorkEndDay
	Public WorkPeriod
	Public WorkUpdateFlag
	Public ManNumber
	Public WomanNumber
	Public NonNumber
	Public HumanNumber
	Public WorkManNumber
	Public WorkManAge
	Public WorkWomanNumber
	Public WorkWomanAge
	Public WorkNumToSection
	Public WorkNumToAll
	Public WorkBuildingType
	Public WorkBuildingRemark
	Public SocietyInsuranceFlag
	Public WelfareAnnuityFlag
	Public EmploymentInsuranceFlag
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Tempクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		SelectionPoint = GetForm("CONF_SelectionPoint", 1)
		NewGraduateFlag = GetForm("CONF_NewGraduateFlag", 1)
		HourlyPayMin = GetForm("CONF_HourlyPayMin", 1)
		HourlyPayMax = GetForm("CONF_HourlyPayMax", 1)
		SalaryPayUnit = GetForm("CONF_SalaryPayUnit", 1)
		OvertimeWorkPayMin = GetForm("CONF_OvertimeWorkPayMin", 1)
		OvertimeWorkPayMax = GetForm("CONF_OvertimeWorkPayMax", 1)
		HolidayHourlyPayMin = GetForm("CONF_HolidayHourlyPayMin", 1)
		HolidayHourlyPayMax = GetForm("CONF_HolidayHourlyPayMax", 1)
		TrafficPayFlag = GetForm("CONF_TrafficPayFlag", 1)
		TrafficPayRemark = GetForm("CONF_TrafficPayRemark", 1)
		WorkStartDay = GetForm("CONF_WorkStartDay", 1)
		WorkEndDay = GetForm("CONF_WorkEndDay", 1)
		WorkPeriod = GetForm("CONF_WorkPeriod", 1)
		WorkUpdateFlag = GetForm("CONF_WorkUpdateFlag", 1)
		ManNumber = GetForm("CONF_ManNumber", 1)
		WomanNumber = GetForm("CONF_WomanNumber", 1)
		NonNumber = GetForm("CONF_NonNumber", 1)
		HumanNumber = GetForm("CONF_HumanNumber", 1)
		WorkManNumber = GetForm("CONF_WorkManNumber", 1)
		WorkManAge = GetForm("CONF_WorkManAge", 1)
		WorkWomanNumber = GetForm("CONF_WorkWomanNumber", 1)
		WorkWomanAge = GetForm("CONF_WorkWomanAge", 1)
		WorkNumToSection = GetForm("CONF_WorkNumToSection", 1)
		WorkNumToAll = GetForm("CONF_WorkNumToAll", 1)
		WorkBuildingType = GetForm("CONF_WorkBuildingType", 1)
		WorkBuildingRemark = GetForm("CONF_WorkBuildingRemark", 1)
		SocietyInsuranceFlag = GetForm("CONF_SocietyInsuranceFlag", 1)
		WelfareAnnuityFlag = GetForm("CONF_WelfareAnnuityFlag", 1)
		EmploymentInsuranceFlag = GetForm("CONF_EmploymentInsuranceFlag", 1)

		IsData = False
		If SelectionPoint <> "" Then IsData = True
		If NewGraduateFlag <> "" Then IsData = True
		If HourlyPayMin <> "" Then IsData = True
		If HourlyPayMax <> "" Then IsData = True
		If SalaryPayUnit <> "" Then IsData = True
		If OvertimeWorkPayMin <> "" Then IsData = True
		If OvertimeWorkPayMax <> "" Then IsData = True
		If HolidayHourlyPayMin <> "" Then IsData = True
		If HolidayHourlyPayMax <> "" Then IsData = True
		If TrafficPayFlag <> "" Then IsData = True
		If TrafficPayRemark <> "" Then IsData = True
		If WorkStartDay <> "" Then IsData = True
		If WorkEndDay <> "" Then IsData = True
		If WorkPeriod <> "" Then IsData = True
		If WorkUpdateFlag <> "" Then IsData = True
		If ManNumber <> "" Then IsData = True
		If WomanNumber <> "" Then IsData = True
		If NonNumber <> "" Then IsData = True
		If HumanNumber <> "" Then IsData = True
		If WorkManNumber <> "" Then IsData = True
		If WorkManAge <> "" Then IsData = True
		If WorkWomanNumber <> "" Then IsData = True
		If WorkWomanAge <> "" Then IsData = True
		If WorkNumToSection <> "" Then IsData = True
		If WorkNumToAll <> "" Then IsData = True
		If WorkBuildingType <> "" Then IsData = True
		If WorkBuildingRemark <> "" Then IsData = True
		If SocietyInsuranceFlag <> "" Then IsData = True
		If WelfareAnnuityFlag <> "" Then IsData = True
		If EmploymentInsuranceFlag <> "" Then IsData = True

		'値チェック
		Err = ""
		If IsFlag(NewGraduateFlag) = False Then Err = Err & "NewGraduateFlag" & vbCrLf
		If IsNumber(HourlyPayMin, 0, True) = False Then Err = Err & "HourlyPayMin" & vbCrLf
		If IsNumber(HourlyPayMax, 0, True) = False Then Err = Err & "HourlyPayMax" & vbCrLf
		If IsRE(SalaryPayUnit, "^[123]$", True) = False Then Err = Err & "SalaryPayUnit" & vbCrLf
		If IsNumber(OvertimeWorkPayMin, 0, True) = False Then Err = Err & "OvertimeWorkPayMin" & vbCrLf
		If IsNumber(OvertimeWorkPayMax, 0, True) = False Then Err = Err & "OvertimeWorkPayMax" & vbCrLf
		If IsNumber(HolidayHourlyPayMin, 0, True) = False Then Err = Err & "HolidayHourlyPayMin" & vbCrLf
		If IsNumber(HolidayHourlyPayMax, 0, True) = False Then Err = Err & "HolidayHourlyPayMax" & vbCrLf
		If IsFlag(TrafficPayFlag) = False Then Err = "TrafficPayFlag" & vbCrLf
		If IsDay(WorkStartDay) = False Then Err = "WorkStartDay" & vbCrLf
		If IsDay(WorkEndDay) = False Then Err = "WorkEndDay" & vbCrLf
		If IsNumber(WorkPeriod, 0, True) = False Then Err = Err & "WorkPeriod" & vbCrLf
		If IsFlag(WorkUpdateFlag) = False Then Err = "WorkUpdateFlag" & vbCrLf
		If IsNumber(ManNumber, 0, False) = False Then Err = Err & "ManNumber" & vbCrLf
		If IsNumber(WomanNumber, 0, False) = False Then Err = Err & "WomanNumber" & vbCrLf
		If IsNumber(NonNumber, 0, False) = False Then Err = Err & "NonNumber" & vbCrLf
		If IsNumber(HumanNumber, 0, False) = False Then Err = Err & "HumanNumber" & vbCrLf
		If IsNumber(WorkManNumber, 0, False) = False Then Err = Err & "WorkManNumber" & vbCrLf
		If IsNumber(WorkManAge, 0, False) = False Then Err = Err & "WorkManAge" & vbCrLf
		If IsNumber(WorkWomanNumber, 0, False) = False Then Err = Err & "WorkWomanNumber" & vbCrLf
		If IsNumber(WorkWomanAge, 0, False) = False Then Err = Err & "WorkWomanAge" & vbCrLf
		If IsNumber(WorkNumToSection, 0, False) = False Then Err = Err & "WorkNumToSection" & vbCrLf
		If IsNumber(WorkNumToAll, 0, False) = False Then Err = Err & "WorkNumToAll" & vbCrLf
		If IsRE(WorkBuildingType, "^[12]$", True) = False Then Err = Err & "WorkBuildingType" & vbCrLf
		If IsFlag(SocietyInsuranceFlag) = False Then Err = "SocietyInsuranceFlag" & vbCrLf
		If IsFlag(WelfareAnnuityFlag) = False Then Err = "WelfareAnnuityFlag" & vbCrLf
		If IsFlag(EmploymentInsuranceFlag) = False Then Err = "EmploymentInsuranceFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Temp 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_Temp '" & vOrderCode & "'" & _
				",'" & SelectionPoint & "'" & _
				",'" & NewGraduateFlag & "'" & _
				",'" & HourlyPayMin & "'" & _
				",'" & HourlyPayMax & "'" & _
				",'" & SalaryPayUnit & "'" & _
				",'" & OvertimeWorkPayMin & "'" & _
				",'" & OvertimeWorkPayMax & "'" & _
				",'" & HolidayHourlyPayMin & "'" & _
				",'" & HolidayHourlyPayMax & "'" & _
				",'" & TrafficPayFlag & "'" & _
				",'" & TrafficPayRemark & "'" & _
				",'" & WorkStartDay & "'" & _
				",'" & WorkEndDay & "'" & _
				",'" & WorkPeriod & "'" & _
				",'" & WorkUpdateFlag & "'" & _
				",'" & ManNumber & "'" & _
				",'" & WomanNumber & "'" & _
				",'" & NonNumber & "'" & _
				",'" & HumanNumber & "'" & _
				",'" & WorkManNumber & "'" & _
				",'" & WorkManAge & "'" & _
				",'" & WorkWomanNumber & "'" & _
				",'" & WorkWomanAge & "'" & _
				",'" & WorkNumToSection & "'" & _
				",'" & WorkNumToAll & "'" & _
				",'" & WorkBuildingType & "'" & _
				",'" & WorkBuildingRemark & "'" & _
				",'" & SocietyInsuranceFlag & "'" & _
				",'" & WelfareAnnuityFlag & "'" & _
				",'" & EmploymentInsuranceFlag & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_Undertake
'概　要：formで飛んできたC_Undertakeテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_Undertake
	Public OrderCode
	Public SelectionPoint
	Public NewGraduateFlag
	Public HourlyPayMin
	Public HourlyPayMax
	Public SalaryPayUnit
	Public OvertimeWorkPayMin
	Public OvertimeWorkPayMax
	Public HolidayHourlyPayMin
	Public HolidayHourlyPayMax
	Public TrafficPayFlag
	Public TrafficPayRemark
	Public WorkStartDay
	Public WorkEndDay
	Public WorkPeriod
	Public WorkUpdateFlag
	Public ManNumber
	Public WomanNumber
	Public NonNumber
	Public HumanNumber
	Public WorkManNumber
	Public WorkManAge
	Public WorkWomanNumber
	Public WorkWomanAge
	Public WorkNumToSection
	Public WorkNumToAll
	Public WorkBuildingType
	Public WorkBuildingRemark
	Public SocietyInsuranceFlag
	Public WelfareAnnuityFlag
	Public EmploymentInsuranceFlag
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Undertakeクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		SelectionPoint = GetForm("CONF_SelectionPoint", 1)
		NewGraduateFlag = GetForm("CONF_NewGraduateFlag", 1)
		HourlyPayMin = GetForm("CONF_HourlyPayMin", 1)
		HourlyPayMax = GetForm("CONF_HourlyPayMax", 1)
		SalaryPayUnit = GetForm("CONF_SalaryPayUnit", 1)
		OvertimeWorkPayMin = GetForm("CONF_OvertimeWorkPayMin", 1)
		OvertimeWorkPayMax = GetForm("CONF_OvertimeWorkPayMax", 1)
		HolidayHourlyPayMin = GetForm("CONF_HolidayHourlyPayMin", 1)
		HolidayHourlyPayMax = GetForm("CONF_HolidayHourlyPayMax", 1)
		TrafficPayFlag = GetForm("CONF_TrafficPayFlag", 1)
		TrafficPayRemark = GetForm("CONF_TrafficPayRemark", 1)
		WorkStartDay = GetForm("CONF_WorkStartDay", 1)
		WorkEndDay = GetForm("CONF_WorkEndDay", 1)
		WorkPeriod = GetForm("CONF_WorkPeriod", 1)
		WorkUpdateFlag = GetForm("CONF_WorkUpdateFlag", 1)
		ManNumber = GetForm("CONF_ManNumber", 1)
		WomanNumber = GetForm("CONF_WomanNumber", 1)
		NonNumber = GetForm("CONF_NonNumber", 1)
		HumanNumber = GetForm("CONF_HumanNumber", 1)
		WorkManNumber = GetForm("CONF_WorkManNumber", 1)
		WorkManAge = GetForm("CONF_WorkManAge", 1)
		WorkWomanNumber = GetForm("CONF_WorkWomanNumber", 1)
		WorkWomanAge = GetForm("CONF_WorkWomanAge", 1)
		WorkNumToSection = GetForm("CONF_WorkNumToSection", 1)
		WorkNumToAll = GetForm("CONF_WorkNumToAll", 1)
		WorkBuildingType = GetForm("CONF_WorkBuildingType", 1)
		WorkBuildingRemark = GetForm("CONF_WorkBuildingRemark", 1)
		SocietyInsuranceFlag = GetForm("CONF_SocietyInsuranceFlag", 1)
		WelfareAnnuityFlag = GetForm("CONF_WelfareAnnuityFlag", 1)
		EmploymentInsuranceFlag = GetForm("CONF_EmploymentInsuranceFlag", 1)

		IsData = False
		If SelectionPoint <> "" Then IsData = True
		If NewGraduateFlag <> "" Then IsData = True
		If HourlyPayMin <> "" Then IsData = True
		If HourlyPayMax <> "" Then IsData = True
		If SalaryPayUnit <> "" Then IsData = True
		If OvertimeWorkPayMin <> "" Then IsData = True
		If OvertimeWorkPayMax <> "" Then IsData = True
		If HolidayHourlyPayMin <> "" Then IsData = True
		If HolidayHourlyPayMax <> "" Then IsData = True
		If TrafficPayFlag <> "" Then IsData = True
		If TrafficPayRemark <> "" Then IsData = True
		If WorkStartDay <> "" Then IsData = True
		If WorkEndDay <> "" Then IsData = True
		If WorkPeriod <> "" Then IsData = True
		If WorkUpdateFlag <> "" Then IsData = True
		If ManNumber <> "" Then IsData = True
		If WomanNumber <> "" Then IsData = True
		If NonNumber <> "" Then IsData = True
		If HumanNumber <> "" Then IsData = True
		If WorkManNumber <> "" Then IsData = True
		If WorkManAge <> "" Then IsData = True
		If WorkWomanNumber <> "" Then IsData = True
		If WorkWomanAge <> "" Then IsData = True
		If WorkNumToSection <> "" Then IsData = True
		If WorkNumToAll <> "" Then IsData = True
		If WorkBuildingType <> "" Then IsData = True
		If WorkBuildingRemark <> "" Then IsData = True
		If SocietyInsuranceFlag <> "" Then IsData = True
		If WelfareAnnuityFlag <> "" Then IsData = True
		If EmploymentInsuranceFlag <> "" Then IsData = True

		'値チェック
		Err = ""
		If IsFlag(NewGraduateFlag) = False Then Err = Err & "NewGraduateFlag" & vbCrLf
		If IsNumber(HourlyPayMin, 0, True) = False Then Err = Err & "HourlyPayMin" & vbCrLf
		If IsNumber(HourlyPayMax, 0, True) = False Then Err = Err & "HourlyPayMax" & vbCrLf
		If IsRE(SalaryPayUnit, "^[123]$", True) = False Then Err = Err & "SalaryPayUnit" & vbCrLf
		If IsNumber(OvertimeWorkPayMin, 0, True) = False Then Err = Err & "OvertimeWorkPayMin" & vbCrLf
		If IsNumber(OvertimeWorkPayMax, 0, True) = False Then Err = Err & "OvertimeWorkPayMax" & vbCrLf
		If IsNumber(HolidayHourlyPayMin, 0, True) = False Then Err = Err & "HolidayHourlyPayMin" & vbCrLf
		If IsNumber(HolidayHourlyPayMax, 0, True) = False Then Err = Err & "HolidayHourlyPayMax" & vbCrLf
		If IsFlag(TrafficPayFlag) = False Then Err = "TrafficPayFlag" & vbCrLf
		If IsDay(WorkStartDay) = False Then Err = "WorkStartDay" & vbCrLf
		If IsDay(WorkEndDay) = False Then Err = "WorkEndDay" & vbCrLf
		If IsNumber(WorkPeriod, 0, True) = False Then Err = Err & "WorkPeriod" & vbCrLf
		If IsFlag(WorkUpdateFlag) = False Then Err = "WorkUpdateFlag" & vbCrLf
		If IsNumber(ManNumber, 0, False) = False Then Err = Err & "ManNumber" & vbCrLf
		If IsNumber(WomanNumber, 0, False) = False Then Err = Err & "WomanNumber" & vbCrLf
		If IsNumber(NonNumber, 0, False) = False Then Err = Err & "NonNumber" & vbCrLf
		If IsNumber(HumanNumber, 0, False) = False Then Err = Err & "HumanNumber" & vbCrLf
		If IsNumber(WorkManNumber, 0, False) = False Then Err = Err & "WorkManNumber" & vbCrLf
		If IsNumber(WorkManAge, 0, False) = False Then Err = Err & "WorkManAge" & vbCrLf
		If IsNumber(WorkWomanNumber, 0, False) = False Then Err = Err & "WorkWomanNumber" & vbCrLf
		If IsNumber(WorkWomanAge, 0, False) = False Then Err = Err & "WorkWomanAge" & vbCrLf
		If IsNumber(WorkNumToSection, 0, False) = False Then Err = Err & "WorkNumToSection" & vbCrLf
		If IsNumber(WorkNumToAll, 0, False) = False Then Err = Err & "WorkNumToAll" & vbCrLf
		If IsRE(WorkBuildingType, "^[12]$", True) = False Then Err = Err & "WorkBuildingType" & vbCrLf
		If IsFlag(SocietyInsuranceFlag) = False Then Err = "SocietyInsuranceFlag" & vbCrLf
		If IsFlag(WelfareAnnuityFlag) = False Then Err = "WelfareAnnuityFlag" & vbCrLf
		If IsFlag(EmploymentInsuranceFlag) = False Then Err = "EmploymentInsuranceFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Undertake 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_Undertake '" & vOrderCode & "'" & _
				",'" & SelectionPoint & "'" & _
				",'" & NewGraduateFlag & "'" & _
				",'" & HourlyPayMin & "'" & _
				",'" & HourlyPayMax & "'" & _
				",'" & SalaryPayUnit & "'" & _
				",'" & OvertimeWorkPayMin & "'" & _
				",'" & OvertimeWorkPayMax & "'" & _
				",'" & HolidayHourlyPayMin & "'" & _
				",'" & HolidayHourlyPayMax & "'" & _
				",'" & TrafficPayFlag & "'" & _
				",'" & TrafficPayRemark & "'" & _
				",'" & WorkStartDay & "'" & _
				",'" & WorkEndDay & "'" & _
				",'" & WorkPeriod & "'" & _
				",'" & WorkUpdateFlag & "'" & _
				",'" & ManNumber & "'" & _
				",'" & WomanNumber & "'" & _
				",'" & NonNumber & "'" & _
				",'" & HumanNumber & "'" & _
				",'" & WorkManNumber & "'" & _
				",'" & WorkManAge & "'" & _
				",'" & WorkWomanNumber & "'" & _
				",'" & WorkWomanAge & "'" & _
				",'" & WorkNumToSection & "'" & _
				",'" & WorkNumToAll & "'" & _
				",'" & WorkBuildingType & "'" & _
				",'" & WorkBuildingRemark & "'" & _
				",'" & SocietyInsuranceFlag & "'" & _
				",'" & WelfareAnnuityFlag & "'" & _
				",'" & EmploymentInsuranceFlag & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_Intro
'概　要：formで飛んできたC_Introテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_Intro
	Public OrderCode
	Public SelectionPoint
	Public NewGraduateFlag
	Public AnnualIncome
	Public IntroduceRate
	Public IntroduceCharge
	Public PenaltyPeriod1
	Public PenaltyRebate1
	Public PenaltyPeriod2
	Public PenaltyRebate2
	Public PenaltyPeriod3
	Public PenaltyRebate3
	Public BonusFlag
	Public BonusSize
	Public RetirementMoneyFlag
	Public EmploymentStartDay
	Public AfterWorkingTypeCode
	Public ManNumber
	Public WomanNumber
	Public NonNumber
	Public HumanNumber
	Public SocietyInsuranceFlag
	Public WelfareAnnuityFlag
	Public EmploymentInsuranceFlag
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Introクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		SelectionPoint = GetForm("CONF_SelectionPoint", 1)
		NewGraduateFlag = GetForm("CONF_NewGraduateFlag", 1)
		AnnualIncome = GetForm("CONF_AnnualIncome", 1)
		IntroduceRate = GetForm("CONF_IntroduceRate", 1)
		IntroduceCharge = GetForm("CONF_IntroduceCharge", 1)
		PenaltyPeriod1 = GetForm("CONF_PenaltyPeriod1", 1)
		PenaltyRebate1 = GetForm("CONF_PenaltyRebate1", 1)
		PenaltyPeriod2 = GetForm("CONF_PenaltyPeriod2", 1)
		PenaltyRebate2 = GetForm("CONF_PenaltyRebate2", 1)
		PenaltyPeriod3 = GetForm("CONF_PenaltyPeriod3", 1)
		PenaltyRebate3 = GetForm("CONF_PenaltyRebate3", 1)
		BonusFlag = GetForm("CONF_BonusFlag", 1)
		BonusSize = GetForm("CONF_BonusSize", 1)
		RetirementMoneyFlag = GetForm("CONF_RetirementMoneyFlag", 1)
		EmploymentStartDay = GetForm("CONF_EmploymentStartDay", 1)
		AfterWorkingTypeCode = GetForm("CONF_AfterWorkingTypeCode", 1)
		ManNumber = GetForm("CONF_ManNumber", 1)
		WomanNumber = GetForm("CONF_WomanNumber", 1)
		NonNumber = GetForm("CONF_NonNumber", 1)
		HumanNumber = GetForm("CONF_HumanNumber", 1)
		SocietyInsuranceFlag = GetForm("CONF_SocietyInsuranceFlag", 1)
		WelfareAnnuityFlag = GetForm("CONF_WelfareAnnuityFlag", 1)
		EmploymentInsuranceFlag = GetForm("CONF_EmploymentInsuranceFlag", 1)

		IsData = False
		If SelectionPoint <> "" Then IsData = True
		If NewGraduateFlag <> "" Then IsData = True
		If AnnualIncome <> "" Then IsData = True
		If IntroduceRate <> "" Then IsData = True
		If IntroduceCharge <> "" Then IsData = True
		If PenaltyPeriod1 <> "" Then IsData = True
		If PenaltyRebate1 <> "" Then IsData = True
		If PenaltyPeriod2 <> "" Then IsData = True
		If PenaltyRebate2 <> "" Then IsData = True
		If PenaltyPeriod3 <> "" Then IsData = True
		If PenaltyRebate3 <> "" Then IsData = True
		If BonusFlag <> "" Then IsData = True
		If BonusSize <> "" Then IsData = True
		If RetirementMoneyFlag <> "" Then IsData = True
		If EmploymentStartDay <> "" Then IsData = True
		If AfterWorkingTypeCode <> "" Then IsData = True
		If ManNumber <> "" Then IsData = True
		If WomanNumber <> "" Then IsData = True
		If NonNumber <> "" Then IsData = True
		If HumanNumber <> "" Then IsData = True
		If SocietyInsuranceFlag <> "" Then IsData = True
		If WelfareAnnuityFlag <> "" Then IsData = True
		If EmploymentInsuranceFlag <> "" Then IsData = True

		'値チェック
		Err = ""
		If IsFlag(NewGraduateFlag) = False Then Err = Err & "NewGraduateFlag" & vbCrLf
		If IsNumber(AnnualIncome, 0, False) = False Then Err = Err & "AnnualIncome" & vbCrLf
		If IsNumber(IntroduceRate, 0, True) = False Then Err = Err & "IntroduceRate" & vbCrLf
		If IsNumber(IntroduceCharge, 0, False) = False Then Err = Err & "IntroduceCharge" & vbCrLf
		If IsNumber(PenaltyRebate1, 0, True) = False Then Err = Err & "PenaltyRebate1" & vbCrLf
		If IsNumber(PenaltyRebate2, 0, True) = False Then Err = Err & "PenaltyRebate2" & vbCrLf
		If IsNumber(PenaltyRebate3, 0, True) = False Then Err = Err & "PenaltyRebate3" & vbCrLf
		If IsFlag(BonusFlag) = False Then Err = Err & "BonusFlag" & vbCrLf
		If IsNumber(BonusSize, 0, True) = False Then Err = Err & "BonusSize" & vbCrLf
		If IsFlag(RetirementMoneyFlag) = False Then Err = Err & "RetirementMoneyFlag" & vbCrLf
		If IsDay(EmploymentStartDay) = False Then Err = Err & "EmploymentStartDay" & vbCrLf
		If IsNumber(AfterWorkingTypeCode, 3, False) = False Then Err = Err & "AfterWorkingTypeCode" & vbCrLf
		If IsNumber(ManNumber, 0, False) = False Then Err = Err & "ManNumber" & vbCrLf
		If IsNumber(WomanNumber, 0, False) = False Then Err = Err & "WomanNumber" & vbCrLf
		If IsNumber(NonNumber, 0, False) = False Then Err = Err & "NonNumber" & vbCrLf
		If IsNumber(HumanNumber, 0, False) = False Then Err = Err & "HumanNumber" & vbCrLf
		If IsFlag(SocietyInsuranceFlag) = False Then Err = Err & "SocietyInsuranceFlag" & vbCrLf
		If IsFlag(WelfareAnnuityFlag) = False Then Err = Err & "WelfareAnnuityFlag" & vbCrLf
		If IsFlag(EmploymentInsuranceFlag) = False Then Err = Err & "EmploymentInsuranceFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Intro 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_Intro '" & vOrderCode & "'" & _
				",'" & SelectionPoint & "'" & _
				",'" & NewGraduateFlag & "'" & _
				",'" & AnnualIncome & "'" & _
				",'" & IntroduceRate & "'" & _
				",'" & IntroduceCharge & "'" & _
				",'" & PenaltyPeriod1 & "'" & _
				",'" & PenaltyRebate1 & "'" & _
				",'" & PenaltyPeriod2 & "'" & _
				",'" & PenaltyRebate2 & "'" & _
				",'" & PenaltyPeriod3 & "'" & _
				",'" & PenaltyRebate3 & "'" & _
				",'" & BonusFlag & "'" & _
				",'" & BonusSize & "'" & _
				",'" & RetirementMoneyFlag & "'" & _
				",'" & EmploymentStartDay & "'" & _
				",'" & AfterWorkingTypeCode & "'" & _
				",'" & ManNumber & "'" & _
				",'" & WomanNumber & "'" & _
				",'" & NonNumber & "'" & _
				",'" & HumanNumber & "'" & _
				",'" & SocietyInsuranceFlag & "'" & _
				",'" & WelfareAnnuityFlag & "'" & _
				",'" & EmploymentInsuranceFlag & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_TTP
'概　要：formで飛んできたC_TTPテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_TTP
	Public OrderCode
	Public SelectionPoint
	Public NewGraduateFlag
	Public HourlyPayMin
	Public HourlyPayMax
	Public SalaryPayUnit
	Public OvertimeWorkPayMin
	Public OvertimeWorkPayMax
	Public HolidayHourlyPayMin
	Public HolidayHourlyPayMax
	Public TrafficPayFlag
	Public TrafficPayRemark
	Public AnnualIncome
	Public IntroduceRate
	Public IntroduceCharge
	Public BonusFlag
	Public BonusSize
	Public RetirementMoneyFlag
	Public WorkStartDay
	Public WorkEndDay
	Public WorkPeriod
	Public WorkPeriodWishFlag
	Public WorkUpdateFlag
	Public EmploymentStartDay
	Public AfterWorkingTypeCode
	Public ManNumber
	Public WomanNumber
	Public NonNumber
	Public HumanNumber
	Public WorkManNumber
	Public WorkManAge
	Public WorkWomanNumber
	Public WorkWomanAge
	Public WorkNumToSection
	Public WorkNumToAll
	Public WorkBuildingType
	Public WorkBuildingRemark
	Public SocietyInsuranceFlag
	Public WelfareAnnuityFlag
	Public EmploymentInsuranceFlag
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_TTPクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		SelectionPoint = GetForm("CONF_SelectionPoint", 1)
		NewGraduateFlag = GetForm("CONF_NewGraduateFlag", 1)
		HourlyPayMin = GetForm("CONF_HourlyPayMin", 1)
		HourlyPayMax = GetForm("CONF_HourlyPayMax", 1)
		SalaryPayUnit = GetForm("CONF_SalaryPayUnit", 1)
		OvertimeWorkPayMin = GetForm("CONF_OvertimeWorkPayMin", 1)
		OvertimeWorkPayMax = GetForm("CONF_OvertimeWorkPayMax", 1)
		HolidayHourlyPayMin = GetForm("CONF_HolidayHourlyPayMin", 1)
		HolidayHourlyPayMax = GetForm("CONF_HolidayHourlyPayMax", 1)
		TrafficPayFlag = GetForm("CONF_TrafficPayFlag", 1)
		TrafficPayRemark = GetForm("CONF_TrafficPayRemark", 1)
		AnnualIncome = GetForm("CONF_AnnualIncome", 1)
		IntroduceRate = GetForm("CONF_IntroduceRate", 1)
		IntroduceCharge = GetForm("CONF_IntroduceCharge", 1)
		BonusFlag = GetForm("CONF_BonusFlag", 1)
		BonusSize = GetForm("CONF_BonusSize", 1)
		RetirementMoneyFlag = GetForm("CONF_RetirementMoneyFlag", 1)
		WorkStartDay = GetForm("CONF_WorkStartDay", 1)
		WorkEndDay = GetForm("CONF_WorkEndDay", 1)
		WorkPeriod = GetForm("CONF_WorkPeriod", 1)
		WorkPeriodWishFlag = GetForm("CONF_WorkPeriodWishFlag", 1)
		WorkUpdateFlag = GetForm("CONF_WorkUpdateFlag", 1)
		EmploymentStartDay = GetForm("CONF_EmploymentStartDay", 1)
		AfterWorkingTypeCode = GetForm("CONF_AfterWorkingTypeCode", 1)
		ManNumber = GetForm("CONF_ManNumber", 1)
		WomanNumber = GetForm("CONF_WomanNumber", 1)
		NonNumber = GetForm("CONF_NonNumber", 1)
		HumanNumber = GetForm("CONF_HumanNumber", 1)
		WorkManNumber = GetForm("CONF_WorkManNumber", 1)
		WorkManAge = GetForm("CONF_WorkManAge", 1)
		WorkWomanNumber = GetForm("CONF_WorkWomanNumber", 1)
		WorkWomanAge = GetForm("CONF_WorkWomanAge", 1)
		WorkNumToSection = GetForm("CONF_WorkNumToSection", 1)
		WorkNumToAll = GetForm("CONF_WorkNumToAll", 1)
		WorkBuildingType = GetForm("CONF_WorkBuildingType", 1)
		WorkBuildingRemark = GetForm("CONF_WorkBuildingRemark", 1)
		SocietyInsuranceFlag = GetForm("CONF_SocietyInsuranceFlag", 1)
		WelfareAnnuityFlag = GetForm("CONF_WelfareAnnuityFlag", 1)
		EmploymentInsuranceFlag = GetForm("CONF_EmploymentInsuranceFlag", 1)

		IsData = False
		If SelectionPoint <> "" Then IsData = True
		If NewGraduateFlag <> "" Then IsData = True
		If HourlyPayMin <> "" Then IsData = True
		If HourlyPayMax <> "" Then IsData = True
		If SalaryPayUnit <> "" Then IsData = True
		If OvertimeWorkPayMin <> "" Then IsData = True
		If OvertimeWorkPayMax <> "" Then IsData = True
		If HolidayHourlyPayMin <> "" Then IsData = True
		If HolidayHourlyPayMax <> "" Then IsData = True
		If TrafficPayFlag <> "" Then IsData = True
		If TrafficPayRemark <> "" Then IsData = True
		If AnnualIncome <> "" Then IsData = True
		If IntroduceRate <> "" Then IsData = True
		If IntroduceCharge <> "" Then IsData = True
		If BonusFlag <> "" Then IsData = True
		If BonusSize <> "" Then IsData = True
		If RetirementMoneyFlag <> "" Then IsData = True
		If WorkStartDay <> "" Then IsData = True
		If WorkEndDay <> "" Then IsData = True
		If WorkPeriod <> "" Then IsData = True
		If WorkPeriodWishFlag <> "" Then IsData = True
		If WorkUpdateFlag <> "" Then IsData = True
		If EmploymentStartDay <> "" Then IsData = True
		If AfterWorkingTypeCode <> "" Then IsData = True
		If ManNumber <> "" Then IsData = True
		If WomanNumber <> "" Then IsData = True
		If NonNumber <> "" Then IsData = True
		If HumanNumber <> "" Then IsData = True
		If WorkManNumber <> "" Then IsData = True
		If WorkManAge <> "" Then IsData = True
		If WorkWomanNumber <> "" Then IsData = True
		If WorkWomanAge <> "" Then IsData = True
		If WorkNumToSection <> "" Then IsData = True
		If WorkNumToAll <> "" Then IsData = True
		If WorkBuildingType <> "" Then IsData = True
		If WorkBuildingRemark <> "" Then IsData = True
		If SocietyInsuranceFlag <> "" Then IsData = True
		If WelfareAnnuityFlag <> "" Then IsData = True
		If EmploymentInsuranceFlag <> "" Then IsData = True

		'値チェック
		Err = ""
		If IsFlag(NewGraduateFlag) = False Then Err = Err & "NewGraduateFlag" & vbCrLf
		If IsNumber(HourlyPayMin, 0, False) = False Then Err = Err & "HourlyPayMin" & vbCrLf
		If IsNumber(HourlyPayMax, 0, False) = False Then Err = Err & "HourlyPayMax" & vbCrLf
		If IsRE(SalaryPayUnit, "^[123]$", True) = False Then Err = Err & "SalaryPayUnit" & vbCrLf
		If IsNumber(OvertimeWorkPayMin, 0, True) = False Then Err = Err & "OvertimeWorkPayMin" & vbCrLf
		If IsNumber(OvertimeWorkPayMax, 0, True) = False Then Err = Err & "OvertimeWorkPayMax" & vbCrLf
		If IsNumber(HolidayHourlyPayMin, 0, True) = False Then Err = Err & "HolidayHourlyPayMin" & vbCrLf
		If IsNumber(HolidayHourlyPayMax, 0, True) = False Then Err = Err & "HolidayHourlyPayMax" & vbCrLf
		If IsFlag(TrafficPayFlag) = False Then Err = Err & "TrafficPayFlag" & vbCrLf
		If IsNumber(AnnualIncome, 0, False) = False Then Err = Err & "AnnualIncome" & vbCrLf
		If IsNumber(IntroduceRate, 0, True) = False Then Err = Err & "IntroduceRate" & vbCrLf
		If IsNumber(IntroduceCharge, 0, False) = False Then Err = Err & "IntroduceCharge" & vbCrLf
		If IsFlag(BonusFlag) = False Then Err = Err & "BonusFlag" & vbCrLf
		If IsNumber(BonusSize, 0, True) = False Then Err = Err & "BonusSize" & vbCrLf
		If IsFlag(RetirementMoneyFlag) = False Then Err = Err & "RetirementMoneyFlag" & vbCrLf
		If IsDay(WorkStartDay) = False Then Err = Err & "WorkStartDay" & vbCrLf
		If IsDay(WorkEndDay) = False Then Err = Err & "WorkEndDay" & vbCrLf
		If IsNumber(WorkPeriod, 0, True) = False Then Err = Err & "WorkPeriod" & vbCrLf
		If IsFlag(WorkPeriodWishFlag) = False Then Err = Err & "WorkPeriodWishFlag" & vbCrLf
		If IsFlag(WorkUpdateFlag) = False Then Err = Err & "WorkUpdateFlag" & vbCrLf
		If IsDay(EmploymentStartDay) = False Then Err = Err & "EmploymentStartDay" & vbCrLf
		If IsNumber(AfterWorkingTypeCode, 3, False) = False Then Err = Err & "AfterWorkingTypeCode" & vbCrLf
		If IsNumber(ManNumber, 0, False) = False Then Err = Err & "ManNumber" & vbCrLf
		If IsNumber(WomanNumber, 0, False) = False Then Err = Err & "WomanNumber" & vbCrLf
		If IsNumber(NonNumber, 0, False) = False Then Err = Err & "NonNumber" & vbCrLf
		If IsNumber(HumanNumber, 0, False) = False Then Err = Err & "HumanNumber" & vbCrLf
		If IsNumber(WorkManNumber, 0, False) = False Then Err = Err & "WorkManNumber" & vbCrLf
		If IsNumber(WorkManAge, 0, False) = False Then Err = Err & "WorkManAge" & vbCrLf
		If IsNumber(WorkWomanNumber, 0, False) = False Then Err = Err & "WorkWomanNumber" & vbCrLf
		If IsNumber(WorkWomanAge, 0, False) = False Then Err = Err & "WorkWomanAge" & vbCrLf
		If IsNumber(WorkNumToSection, 0, False) = False Then Err = Err & "WorkNumToSection" & vbCrLf
		If IsNumber(WorkNumToAll, 0, False) = False Then Err = Err & "WorkNumToAll" & vbCrLf
		If IsRE(WorkBuildingType, "^[12]$", True) = False Then Err = Err & "WorkBuildingType" & vbCrLf
		If IsFlag(SocietyInsuranceFlag) = False Then Err = Err & "SocietyInsuranceFlag" & vbCrLf
		If IsFlag(WelfareAnnuityFlag) = False Then Err = Err & "WelfareAnnuityFlag" & vbCrLf
		If IsFlag(EmploymentInsuranceFlag) = False Then Err = Err & "EmploymentInsuranceFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_TTP 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_TTP '" & vOrderCode & "'" & _
				",'" & SelectionPoint & "'" & _
				",'" & NewGraduateFlag & "'" & _
				",'" & HourlyPayMin & "'" & _
				",'" & HourlyPayMax & "'" & _
				",'" & SalaryPayUnit & "'" & _
				",'" & OvertimeWorkPayMin & "'" & _
				",'" & OvertimeWorkPayMax & "'" & _
				",'" & HolidayHourlyPayMin & "'" & _
				",'" & HolidayHourlyPayMax & "'" & _
				",'" & TrafficPayFlag & "'" & _
				",'" & TrafficPayRemark & "'" & _
				",'" & AnnualIncome & "'" & _
				",'" & IntroduceRate & "'" & _
				",'" & IntroduceCharge & "'" & _
				",'" & BonusFlag & "'" & _
				",'" & BonusSize & "'" & _
				",'" & RetirementMoneyFlag & "'" & _
				",'" & WorkStartDay & "'" & _
				",'" & WorkEndDay & "'" & _
				",'" & WorkPeriod & "'" & _
				",'" & WorkPeriodWishFlag & "'" & _
				",'" & WorkUpdateFlag & "'" & _
				",'" & EmploymentStartDay & "'" & _
				",'" & AfterWorkingTypeCode & "'" & _
				",'" & ManNumber & "'" & _
				",'" & WomanNumber & "'" & _
				",'" & NonNumber & "'" & _
				",'" & HumanNumber & "'" & _
				",'" & WorkManNumber & "'" & _
				",'" & WorkManAge & "'" & _
				",'" & WorkWomanNumber & "'" & _
				",'" & WorkWomanAge & "'" & _
				",'" & WorkNumToSection & "'" & _
				",'" & WorkNumToAll & "'" & _
				",'" & WorkBuildingType & "'" & _
				",'" & WorkBuildingRemark & "'" & _
				",'" & SocietyInsuranceFlag & "'" & _
				",'" & WelfareAnnuityFlag & "'" & _
				",'" & EmploymentInsuranceFlag & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_Navi
'概　要：formで飛んできたC_Naviテーブル用のデータを持つためのクラス
'備　考：
'更　新：2006/03/24 LIS K.Kokubo 作成
'　　　：2008/05/01 LIS K.Kokubo [SecretFlag]シークレットフラグ追加
'******************************************************************************
Class clsC_Navi
	Public OrderCode
	Public WorkStartDay
	Public WorkEndDay
	Public HumanNumber
	Public CollaborationCode
	Public PermitFlag
	Public PermitDay
	Public TrafficPayFlag
	Public SecretFlag
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Naviクラスの初期化関数
	'備　考：
	'更　新：2006/03/24 LIS K.Kokubo 作成
	'　　　：2008/05/01 LIS K.Kokubo [SecretFlag]シークレットフラグ追加
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		WorkStartDay = GetForm("CONF_WorkStartDay", 1)
		WorkEndDay = GetForm("CONF_WorkEndDay", 1)
		HumanNumber = GetForm("CONF_HumanNumber", 1)
		CollaborationCode = GetForm("CONF_CollaborationCode", 1)
		PermitFlag = GetForm("CONF_PermitFlag", 1)
		PermitDay = GetForm("CONF_PermitDay", 1)
		TrafficPayFlag = GetForm("CONF_TrafficPayFlag", 1)
		SecretFlag = GetForm("CONF_SecretFlag", 1)

		IsData = False
		If WorkStartDay <> "" Then IsData = True
		If WorkEndDay <> "" Then IsData = True
		If HumanNumber <> "" Then IsData = True
		If CollaborationCode <> "" Then IsData = True
		If PermitFlag <> "" Then IsData = True
		If PermitDay <> "" Then IsData = True
		If TrafficPayFlag <> "" Then IsData = True
		If SecretFlag <> "" Then IsData = True

		'値チェック
		If IsDay(WorkStartDay) = False Then Err = Err & "WorkStartDay" & vbCrLf
		If IsDay(WorkEndDay) = False Then Err = Err & "WorkEndDay" & vbCrLf
		If IsNumber(HumanNumber, 0, False) = False Then Err = Err & "HumanNumber" & vbCrLf
		If IsMainCode(CollaborationCode) = False Then Err = Err & "CollaborationCode" & vbCrLf
		If IsFlag(PermitFlag) Then Err = Err & "PermitFlag" & vbCrLf
		If IsDay(PermitDay) Then Err = Err & "PermitDay" & vbCrLf
		If IsFlag(TrafficPayFlag) Then Err = Err & "TrafficPayFlag" & vbCrLf
		If IsFlag(SecretFlag) Then Err = Err & "SecretFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Navi 実行SQL取得
	'備　考：
	'更　新：2006/03/24 LIS K.Kokubo 作成
	'　　　：2008/05/01 LIS K.Kokubo [SecretFlag]シークレットフラグ追加
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_Navi '" & vOrderCode & "'" & _
				",'" & WorkStartDay & "'" & _
				",'" & WorkEndDay & "'" & _
				",'" & HumanNumber & "'" & _
				",'" & CollaborationCode & "'" & _
				",'" & PermitFlag & "'" & _
				",'" & PermitDay & "'" & _
				",'" & TrafficPayFlag & "'" & _
				",'" & SecretFlag & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_NearbyStation
'概　要：formで飛んできたC_NearbyStationテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_NearbyStation
	Public OrderCode
	Public StationCode()
	Public ToStation()
	Public ToStationRemark()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_NearbyStationクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1

		MaxIndex = -1
		IsData = False
		OrderCode = GetForm("CONF_OrderCode", 1)

		Err = ""
		Do While True
			If ExistsForm("CONF_StationCode" & idx) =False Then Exit Do

			If GetForm("CONF_StationCode" & idx, 1) <> "" Then
				MaxIndex = MaxIndex + 1
				ReDim Preserve StationCode(MaxIndex)		: StationCode(MaxIndex) = GetForm("CONF_StationCode" & idx, 1)
				ReDim Preserve ToStation(MaxIndex)			: ToStation(MaxIndex) = GetForm("CONF_ToStation" & idx, 1)
				ReDim Preserve ToStationRemark(MaxIndex)	: ToStationRemark(MaxIndex) = GetForm("CONF_ToStationRemark" & idx, 1)

				If IsRE(StationCode(MaxIndex), "^\d\d\d\d\d$", True) = False Then Err = Err & "StationCode" & idx & vbCrLf
				If IsNumber(ToStation(MaxIndex), 0, True) = False Then Err = Err & "ToStation" & idx & vbCrLf

				IsData = True
			End If

			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_NearbyStation 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_C_NearbyStation '" & vOrderCode & "'" & vbCrLf
		If MaxIndex < 0 Then Exit Function
		For idx = 0 To UBound(StationCode)
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_C_NearbyStation '" & vOrderCode & "'" & _
				",''" & _
				",'" & StationCode(idx) & "'" & _
				",'" & ToStation(idx) & "'" & _
				",'" & ToStationRemark(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsC_Contact
'概　要：formで飛んできたC_Contactテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_Contact
	Public OrderCode
	Public CompanyName
	Public SectionName
	Public PersonPost
	Public PersonName
	Public PersonName_F
	Public TelNumber
	Public FaxNumber
	Public Mailaddress
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Contactクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		CompanyName = GetForm("CONF_ContactCompanyName", 1)
		SectionName = GetForm("CONF_ContactSectionName", 1)
		PersonPost = GetForm("CONF_ContactPersonPost", 1)
		PersonName = GetForm("CONF_ContactPersonName", 1)
		PersonName_F = GetForm("CONF_ContactPersonName_F", 1)
		TelNumber = GetForm("CONF_ContactTelNumber", 1)
		FaxNumber = GetForm("CONF_ContactFaxNumber", 1)
		Mailaddress = GetForm("CONF_ContactMailaddress", 1)

		IsData = False
		If CompanyName <> "" Then IsData = True
		If SectionName <> "" Then IsData = True
		If PersonPost <> "" Then IsData = True
		If PersonName <> "" Then IsData = True
		If PersonName_F <> "" Then IsData = True
		If TelNumber <> "" Then IsData = True
		If FaxNumber <> "" Then IsData = True
		If Mailaddress <> "" Then IsData = True

		'値チェック
		If IsNumber(TelNumber, 0, False) = False Then Err = Err & "TelNumber" & vbCrLf
		If IsNumber(FaxNumber, 0, False) = False Then Err = Err & "FaxNumber" & vbCrLf
		If IsMailAddress(Mailaddress) = False Then Err = Err & "MailAddress" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Contact 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_Contact '" & vOrderCode & "'" & _
				",'" & CompanyName & "'" & _
				",'" & SectionName & "'" & _
				",'" & PersonPost & "'" & _
				",'" & PersonName & "'" & _
				",'" & PersonName_F & "'" & _
				",'" & TelNumber & "'" & _
				",'" & FaxNumber & "'" & _
				",'" & Mailaddress & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_Bill
'概　要：formで飛んできたC_Billテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_Bill
	Public OrderCode
	Public BillMin
	Public BillMax
	Public SalaryBillUnit
	Public OvertimeWorkBillMin
	Public OvertimeWorkBillMax
	Public HolidayHourlyBillMin
	Public HolidayHourlyBillMax
	Public TrafficBillFlag
	Public TrafficBillRemark
	Public OtherBill
	Public DayInMonth
	Public Bill
	Public Pay
	Public GrossMargin
	Public GrossMarginPercentage
	Public TightDay
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Billクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		BillMin = GetForm("CONF_BillMin", 1)
		BillMax = GetForm("CONF_BillMax", 1)
		SalaryBillUnit = GetForm("CONF_SalaryBillUnit", 1)
		OvertimeWorkBillMin = GetForm("CONF_OvertimeWorkBillMin", 1)
		OvertimeWorkBillMax = GetForm("CONF_OvertimeWorkBillMax", 1)
		HolidayHourlyBillMin = GetForm("CONF_HolidayHourlyBillMin", 1)
		HolidayHourlyBillMax = GetForm("CONF_HolidayHourlyBillMax", 1)
		TrafficBillFlag = GetForm("CONF_TrafficBillFlag", 1)
		TrafficBillRemark = GetForm("CONF_TrafficBillRemark", 1)
		OtherBill = GetForm("CONF_OtherBill", 1)
		DayInMonth = GetForm("CONF_DayInMonth", 1)
		Bill = GetForm("CONF_Bill", 1)
		Pay = GetForm("CONF_Pay", 1)
		GrossMargin = GetForm("CONF_GrossMargin", 1)
		GrossMarginPercentage = GetForm("CONF_GrossMarginPercentage", 1)
		TightDay = GetForm("CONF_TightDay", 1)

		IsData = False
		If BillMin <> "" Then IsData = True
		If BillMax <> "" Then IsData = True
		If SalaryBillUnit <> "" Then IsData = True
		If OvertimeWorkBillMin <> "" Then IsData = True
		If OvertimeWorkBillMax <> "" Then IsData = True
		If HolidayHourlyBillMin <> "" Then IsData = True
		If HolidayHourlyBillMax <> "" Then IsData = True
		If TrafficBillFlag <> "" Then IsData = True
		If TrafficBillRemark <> "" Then IsData = True
		If OtherBill <> "" Then IsData = True
		If DayInMonth <> "" Then IsData = True
		If Bill <> "" Then IsData = True
		If Pay <> "" Then IsData = True
		If GrossMargin <> "" Then IsData = True
		If GrossMarginPercentage <> "" Then IsData = True
		If TightDay <> "" Then IsData = True

		'値チェック
		Err = ""
		If IsNumber(BillMin, 0, True) = False Then Err = Err & "BillMin" & vbCrLf
		If IsNumber(BillMax, 0, True) = False Then Err = Err & "BillMax" & vbCrLf
		If IsRE(SalaryBillUnit, "^[123]$", True) = False Then Err = Err & "SalaryBillUnit" & vbCrLf
		If IsNumber(OvertimeWorkBillMin, 0, True) = False Then Err = Err & "OvertimeWorkBillMin" & vbCrLf
		If IsNumber(OvertimeWorkBillMax, 0, True) = False Then Err = Err & "OvertimeWorkBillMax" & vbCrLf
		If IsNumber(HolidayHourlyBillMin, 0, True) = False Then Err = Err & "HolidayHourlyBillMin" & vbCrLf
		If IsNumber(HolidayHourlyBillMax, 0, True) = False Then Err = Err & "HolidayHourlyBillMax" & vbCrLf
		If IsFlag(TrafficBillFlag) = False Then Err = Err & "TrafficBillFlag" & vbCrLf
		If IsNumber(DayInMonth, 0, False) = False Then Err = Err & "DayInMonth" & vbCrLf
		If IsNumber(Bill, 0, True) = False Then Err = Err & "Bill" & vbCrLf
		If IsNumber(Pay, 0, True) = False Then Err = Err & "Pay" & vbCrLf
		If IsNumber(GrossMargin, 0, True) = False Then Err = Err & "GrossMargin" & vbCrLf
		If IsNumber(GrossMarginPercentage, 0, True) = False Then Err = Err & "GrossMarginPercentage" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Bill 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_Bill '" & vOrderCode & "'" & _
				",'" & BillMin & "'" & _
				",'" & BillMax & "'" & _
				",'" & SalaryBillUnit & "'" & _
				",'" & OvertimeWorkBillMin & "'" & _
				",'" & OvertimeWorkBillMax & "'" & _
				",'" & HolidayHourlyBillMin & "'" & _
				",'" & HolidayHourlyBillMax & "'" & _
				",'" & TrafficBillFlag & "'" & _
				",'" & TrafficBillRemark & "'" & _
				",'" & OtherBill & "'" & _
				",'" & DayInMonth & "'" & _
				",'" & Bill & "'" & _
				",'" & Pay & "'" & _
				",'" & GrossMargin & "'" & _
				",'" & GrossMarginPercentage & "'" & _
				",'" & TightDay & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_SupplementInfo
'概　要：formで飛んできたC_SupplementInfoテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_SupplementInfo
	Public OrderCode
	Public CompanyCode
	Public CompanySpeciality
	Public CatchCopy
	Public BizName1
	Public BizPercentage1
	Public BizName2
	Public BizPercentage2
	Public BizName3
	Public BizPercentage3
	Public BizName4
	Public BizPercentage4
	Public UITurnFlag
	Public UtilizeLanguageFlag
	Public ManyHolidayFlag
	Public InexperiencedPersonFlag
	Public FlexTimeFlag
	Public NearStationFlag
	Public NoSmokingFlag
	Public NewlyBuiltFlag
	Public LandmarkFlag
	Public RenovationFlag
	Public DesignersFlag
	Public CompanyCafeteriaFlag
	Public ShortOvertimeFlag
	Public MaternityFlag
	Public DressFreeFlag
	Public MammyFlag
	Public FixedTimeFlag
	Public ShortTimeFlag
	Public HandicappedFlag
	Public PRTitle1
	Public PRContents1
	Public PRTitle2
	Public PRContents2
	Public PRTitle3
	Public PRContents3
	Public EntryInfo
	Public Process1
	Public Process2
	Public Process3
	Public Process4
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_SupplementInfoクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		MaxIndex = -1
		OrderCode = GetForm("CONF_OrderCode", 1)
		CompanyCode = GetForm("CONF_CompanyCode", 1)
		CompanySpeciality = GetForm("CONF_CompanySpeciality", 1)
		CatchCopy = GetForm("CONF_CatchCopy", 1)
		BizName1 = GetForm("CONF_BizName1", 1)
		BizPercentage1 = GetForm("CONF_BizPercentage1", 1)
		BizName2 = GetForm("CONF_BizName2", 1)
		BizPercentage2 = GetForm("CONF_BizPercentage2", 1)
		BizName3 = GetForm("CONF_BizName3", 1)
		BizPercentage3 = GetForm("CONF_BizPercentage3", 1)
		BizName4 = GetForm("CONF_BizName4", 1)
		BizPercentage4 = GetForm("CONF_BizPercentage4", 1)
		UITurnFlag = GetForm("CONF_UITurnFlag", 1)
		UtilizeLanguageFlag = GetForm("CONF_UtilizeLanguageFlag", 1)
		ManyHolidayFlag = GetForm("CONF_ManyHolidayFlag", 1)
		InexperiencedPersonFlag = GetForm("CONF_InexperiencedPersonFlag", 1)
		FlexTimeFlag = GetForm("CONF_FlexTimeFlag", 1)
		NearStationFlag = GetForm("CONF_NearStationFlag", 1)
		NoSmokingFlag = GetForm("CONF_NoSmokingFlag", 1)
		NewlyBuiltFlag = GetForm("CONF_NewlyBuiltFlag", 1)
		LandmarkFlag = GetForm("CONF_LandmarkFlag", 1)
		RenovationFlag = GetForm("CONF_RenovationFlag", 1)
		DesignersFlag = GetForm("CONF_DesignersFlag", 1)
		CompanyCafeteriaFlag = GetForm("CONF_CompanyCafeteriaFlag", 1)
		ShortOvertimeFlag = GetForm("CONF_ShortOvertimeFlag", 1)
		MaternityFlag = GetForm("CONF_MaternityFlag", 1)
		DressFreeFlag = GetForm("CONF_DressFreeFlag", 1)
		MammyFlag = GetForm("CONF_MammyFlag", 1)
		FixedTimeFlag = GetForm("CONF_FixedTimeFlag", 1)
		ShortTimeFlag = GetForm("CONF_ShortTimeFlag", 1)
		HandicappedFlag = GetForm("CONF_HandicappedFlag", 1)
		PRTitle1 = GetForm("CONF_PRTitle1", 1)
		PRContents1 = GetForm("CONF_PRContents1", 1)
		PRTitle2 = GetForm("CONF_PRTitle2", 1)
		PRContents2 = GetForm("CONF_PRContents2", 1)
		PRTitle3 = GetForm("CONF_PRTitle3", 1)
		PRContents3 = GetForm("CONF_PRContents3", 1)
		EntryInfo = GetForm("CONF_EntryInfo", 1)
		Process1 = GetForm("CONF_Process1", 1)
		Process2 = GetForm("CONF_Process2", 1)
		Process3 = GetForm("CONF_Process3", 1)
		Process4 = GetForm("CONF_Process4", 1)

		IsData = False
		If CompanyCode <> "" Then IsData = True
		If CompanySpeciality <> "" Then IsData = True
		If CatchCopy <> "" Then IsData = True
		If BizName1 <> "" Then IsData = True
		If BizPercentage1 <> "" Then IsData = True
		If BizName2 <> "" Then IsData = True
		If BizPercentage2 <> "" Then IsData = True
		If BizName3 <> "" Then IsData = True
		If BizPercentage3 <> "" Then IsData = True
		If BizName4 <> "" Then IsData = True
		If BizPercentage4 <> "" Then IsData = True
		If UITurnFlag <> "" Then IsData = True
		If UtilizeLanguageFlag <> "" Then IsData = True
		If ManyHolidayFlag <> "" Then IsData = True
		If InexperiencedPersonFlag <> "" Then IsData = True
		If FlexTimeFlag <> "" Then IsData = True
		If NearStationFlag <> "" Then IsData = True
		If NoSmokingFlag <> "" Then IsData = True
		If NewlyBuiltFlag <> "" Then IsData = True
		If LandmarkFlag <> "" Then IsData = True
		If RenovationFlag <> "" Then IsData = True
		If DesignersFlag <> "" Then IsData = True
		If CompanyCafeteriaFlag <> "" Then IsData = True
		If ShortOvertimeFlag <> "" Then IsData = True
		If MaternityFlag <> "" Then IsData = True
		If DressFreeFlag <> "" Then IsData = True
		If MammyFlag <> "" Then IsData = True
		If FixedTimeFlag <> "" Then IsData = True
		If ShortTimeFlag <> "" Then IsData = True
		If HandicappedFlag <> "" Then IsData = True
		If PRTitle1 <> "" Then IsData = True
		If PRContents1 <> "" Then IsData = True
		If PRTitle2 <> "" Then IsData = True
		If PRContents2 <> "" Then IsData = True
		If PRTitle3 <> "" Then IsData = True
		If PRContents3 <> "" Then IsData = True
		If EntryInfo <> "" Then IsData = True
		If Process1 <> "" Then IsData = True
		If Process2 <> "" Then IsData = True
		If Process3 <> "" Then IsData = True
		If Process4 <> "" Then IsData = True

		'値チェック
		Err = ""
		If IsMainCode(CompanyCode) = False Then Err = Err & "CompanyCode" & vbCrLf
		If IsNumber(BizPercentage1, 0, False) = False Then Err = Err & "BizPercentage1" & vbCrLf
		If IsNumber(BizPercentage2, 0, False) = False Then Err = Err & "BizPercentage2" & vbCrLf
		If IsNumber(BizPercentage3, 0, False) = False Then Err = Err & "BizPercentage3" & vbCrLf
		If IsNumber(BizPercentage4, 9, False) = False Then Err = Err & "BizPercentage4" & vbCrLf
		If IsFlag(UITurnFlag) = False Then Err = Err & "UITurnFlag" & vbCrLf
		If IsFlag(UtilizeLanguageFlag) = False Then Err = Err & "UtilizeLanguageFlag" & vbCrLf
		If IsFlag(ManyHolidayFlag) = False Then Err = Err & "ManyHolidayFlag" & vbCrLf
		If IsFlag(InexperiencedPersonFlag) = False Then Err = Err & "InexperiencedPersonFlag" & vbCrLf
		If IsFlag(FlexTimeFlag) = False Then Err = Err & "FlexTimeFlag" & vbCrLf
		If IsFlag(NearStationFlag) = False Then Err = Err & "NearStationFlag" & vbCrLf
		If IsFlag(NoSmokingFlag) = False Then Err = Err & "NoSmokingFlag" & vbCrLf
		If IsFlag(NewlyBuiltFlag) = False Then Err = Err & "NewlyBuiltFlag" & vbCrLf
		If IsFlag(LandmarkFlag) = False Then Err = Err & "LandmarkFlag" & vbCrLf
		If IsFlag(RenovationFlag) = False Then Err = Err & "RenovationFlag" & vbCrLf
		If IsFlag(DesignersFlag) = False Then Err = Err & "DesignersFlag" & vbCrLf
		If IsFlag(CompanyCafeteriaFlag) = False Then Err = Err & "CompanyCafeteriaFlag" & vbCrLf
		If IsFlag(ShortOvertimeFlag) = False Then Err = Err & "ShortOvertimeFlag" & vbCrLf
		If IsFlag(MaternityFlag) = False Then Err = Err & "MaternityFlag" & vbCrLf
		If IsFlag(DressFreeFlag) = False Then Err = Err & "DressFreeFlag" & vbCrLf
		If IsFlag(MammyFlag) = False Then Err = Err & "MammyFlag" & vbCrLf
		If IsFlag(FixedTimeFlag) = False Then Err = Err & "FixedTimeFlag" & vbCrLf
		If IsFlag(ShortTimeFlag) = False Then Err = Err & "ShortTimeFlag" & vbCrLf
		If IsFlag(HandicappedFlag) = False Then Err = Err & "HandicappedFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_SupplementInfo 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = GetRegSQL & "sp_Reg_C_SupplementInfo '" & vOrderCode & "'" & _
				",'" & CompanyCode & "'" & _
				",'" & CompanySpeciality & "'" & _
				",'" & CatchCopy & "'" & _
				",'" & BizName1 & "'" & _
				",'" & BizPercentage1 & "'" & _
				",'" & BizName2 & "'" & _
				",'" & BizPercentage2 & "'" & _
				",'" & BizName3 & "'" & _
				",'" & BizPercentage3 & "'" & _
				",'" & BizName4 & "'" & _
				",'" & BizPercentage4 & "'" & _
				",'" & UITurnFlag & "'" & _
				",'" & UtilizeLanguageFlag & "'" & _
				",'" & ManyHolidayFlag & "'" & _
				",'" & InexperiencedPersonFlag & "'" & _
				",'" & FlexTimeFlag & "'" & _
				",'" & NearStationFlag & "'" & _
				",'" & NoSmokingFlag & "'" & _
				",'" & NewlyBuiltFlag & "'" & _
				",'" & LandmarkFlag & "'" & _
				",'" & RenovationFlag & "'" & _
				",'" & DesignersFlag & "'" & _
				",'" & CompanyCafeteriaFlag & "'" & _
				",'" & ShortOvertimeFlag & "'" & _
				",'" & MaternityFlag & "'" & _
				",'" & DressFreeFlag & "'" & _
				",'" & MammyFlag & "'" & _
				",'" & FixedTimeFlag & "'" & _
				",'" & ShortTimeFlag & "'" & _
				",'" & HandicappedFlag & "'" & _
				",'" & PRTitle1 & "'" & _
				",'" & PRContents1 & "'" & _
				",'" & PRTitle2 & "'" & _
				",'" & PRContents2 & "'" & _
				",'" & PRTitle3 & "'" & _
				",'" & PRContents3 & "'" & _
				",'" & EntryInfo & "'" & _
				",'" & Process1 & "'" & _
				",'" & Process2 & "'" & _
				",'" & Process3 & "'" & _
				",'" & Process4 & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_WorkingCondition
'概　要：formで飛んできたC_WorkingConditionテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_WorkingCondition
	Public OrderCode
	Public WorkStartTime()
	Public WorkEndTime()
	Public RestStartTime()
	Public RestEndTime()
	Public RestTotalTime()
	Public ContractTime()
	Public TimeUnit()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_WorkingConditionクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg

		MaxIndex = -1
		IsData = False
		OrderCode = GetForm("CONF_OrderCode", 1)

		Do While True
			If ExistsForm("CONF_WorkStartTime" & idx) = False Then Exit Do
			flg = False

			If GetForm("CONF_WorkStartTime" & idx, 1) <> "" Then flg = True
			If GetForm("CONF_WorkEndTime" & idx, 1) <> "" Then flg = True
			If GetForm("CONF_RestStartTime" & idx, 1) <> "" Then flg = True
			If GetForm("CONF_RestEndTime" & idx, 1) <> "" Then flg = True
			If GetForm("CONF_RestTotalTime" & idx, 1) <> "" Then flg = True
			If GetForm("CONF_ContractTime" & idx, 1) <> "" Then flg = True
			If GetForm("CONF_TimeUnit" & idx, 1) <> "" Then flg = True

			If flg = True Then
				IsData = True
				MaxIndex = MaxIndex + 1
				ReDim Preserve WorkStartTime(MaxIndex)	: WorkStartTime(MaxIndex) = GetForm("CONF_WorkStartTime" & idx, 1)
				ReDim Preserve WorkEndTime(MaxIndex)	: WorkEndTime(MaxIndex) = GetForm("CONF_WorkEndTime" & idx, 1)
				ReDim Preserve RestStartTime(MaxIndex)	: RestStartTime(MaxIndex) = GetForm("CONF_RestStartTime" & idx, 1)
				ReDim Preserve RestEndTime(MaxIndex)	: RestEndTime(MaxIndex) = GetForm("CONF_RestEndTime" & idx, 1)
				ReDim Preserve RestTotalTime(MaxIndex)	: RestTotalTime(MaxIndex) = GetForm("CONF_RestTotalTime" & idx, 1)
				ReDim Preserve ContractTime(MaxIndex)	: ContractTime(MaxIndex) = GetForm("CONF_ContractTime" & idx, 1)
				ReDim Preserve TimeUnit(MaxIndex)		: TimeUnit(MaxIndex) = GetForm("CONF_TimeUnit" & idx, 1)

				'値チェック
				Err = ""
				If IsNumber(WorkStartTime(MaxIndex), 4, False) = False Then Err = Err & "WorkStartTime" & MaxIndex & vbCrLf
				If IsNumber(WorkEndTime(MaxIndex), 4, False) = False Then Err = Err & "WorkEndTime" & MaxIndex & vbCrLf
				If IsNumber(RestStartTime(MaxIndex), 4, False) = False Then Err = Err & "RestStartTime" & MaxIndex & vbCrLf
				If IsNumber(RestEndTime(MaxIndex), 4, False) = False Then Err = Err & "RestEndTime" & MaxIndex & vbCrLf
				If IsNumber(RestTotalTime(MaxIndex), 0, False) = False Then Err = Err & "RestTotalTime" & MaxIndex & vbCrLf
				If IsNumber(ContractTime(MaxIndex), 0, True) = False Then Err = Err & "ContractTime" & MaxIndex & vbCrLf
				If IsNumber(TimeUnit(MaxIndex), 0, False) = False Then Err = Err & "TimeUnit" & MaxIndex & vbCrLf
			End If

			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_WorkingCondition 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_C_WorkingCondition '" & vOrderCode & "'" & vbCrLf
		If MaxIndex < 0 Then Exit Function
		For idx = 0 To UBound(WorkStartTime)
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_C_WorkingCondition '" & vOrderCode & "'" & _
				",''" & _
				",'" & WorkStartTime(idx) & "'" & _
				",'" & WorkEndTime(idx) & "'" & _
				",'" & RestStartTime(idx) & "'" & _
				",'" & RestEndTime(idx) & "'" & _
				",'" & RestTotalTime(idx) & "'" & _
				",'" & ContractTime(idx) & "'" & _
				",'" & TimeUnit(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsC_Skill
'概　要：formで飛んできたC_Skillテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_Skill
	Public OrderCode
	Public CategoryCode
	Public Code()
	Public Period()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Skillクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize(ByVal vSkillCategoryCode)
		Dim idx	: idx = 1
		Dim flg

		MaxIndex = -1
		IsData = False
		OrderCode = GetForm("CONF_OrderCode", 1)
		CategoryCode = vSkillCategoryCode

		Do While True
			If ExistsForm("CONF_" & vSkillCategoryCode & idx) = False Then Exit Do
			flg = False

			If GetForm("CONF_" & vSkillCategoryCode & idx, 1) <> "" Then
				MaxIndex = MaxIndex + 1
				IsData = True
				flg = True

				ReDim Preserve Code(MaxIndex)	: Code(MaxIndex) = GetForm("CONF_" & vSkillCategoryCode & idx, 1)
				ReDim Preserve Period(MaxIndex)	: Period(MaxIndex) = GetForm("CONF_" & vSkillCategoryCode & "Period" & idx, 1)

				'値チェック
				If IsNumber(Code(MaxIndex), 3, False) = False Then Err = Err & "Code" & MaxIndex & vbCrLf
				If IsNumber(Period(MaxIndex), 0, True) = False Then Err = Err & "Period" & MaxIndex & vbCrLf
			End If

			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Skill 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_C_Skill '" & vOrderCode & "', '" & CategoryCode & "'" & vbCrLf
		If MaxIndex < 0 Then Exit Function
		For idx = 0 To UBound(Code)
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_C_Skill '" & vOrderCode & "'" & _
				",''" & _
				",'" & CategoryCode & "'" & _
				",'" & Code(idx) & "'" & _
				",'" & Period(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsC_License
'概　要：formで飛んできたC_Licenseテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_License
	Public OrderCode
	Public GroupCode()
	Public CategoryCode()
	Public Code()
	Public GetYear()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Licenseクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim sGroupCode
		Dim sCategoryCode
		Dim sCode
		Dim flg

		MaxIndex = -1
		IsData = False
		OrderCode = GetForm("CONF_OrderCode", 1)

		Do While True
			If ExistsForm("CONF_LicenseCode"  & idx) = False Then Exit Do
			flg = false
			sGroupCode = ""
			sCategoryCode = ""
			sCode = ""

			If GetForm("CONF_LicenseGroupCode" & idx, 1) <> "" _
			And GetForm("CONF_LicenseCategoryCode" & idx, 1) <> "" _
			And GetForm("CONF_LicenseCode" & idx, 1) <> "" Then
				flg = True
				IsData = True
				sGroupCode = GetForm("CONF_LicenseGroupCode" & idx, 1)
				sCategoryCode = GetForm("CONF_LicenseCategoryCode" & idx, 1)
				sCode = GetForm("CONF_LicenseCode" & idx, 1)
			End If
			'第２案
			If Len(GetForm("CONF_LicenseCode" & idx, 1)) = 7 Then
				flg = True
				IsData = True
				sGroupCode = Left(GetForm("CONF_LicenseCode" & idx, 1), 2)
				sCategoryCode = Mid(GetForm("CONF_LicenseCode" & idx, 1), 3, 3)
				sCode = Right(GetForm("CONF_LicenseCode" & idx, 3), 2)
			End If

			If flg = True Then
				MaxIndex = MaxIndex + 1
				ReDim Preserve GroupCode(MaxIndex)		: GroupCode(MaxIndex) = sGroupCode
				ReDim Preserve CategoryCode(MaxIndex)	: CategoryCode(MaxIndex) = sCategoryCode
				ReDim Preserve Code(MaxIndex)			: Code(MaxIndex) = sCode
				ReDim Preserve GetYear(MaxIndex)		: GetYear(MaxIndex) = GetForm("CONF_GetYear" & idx, 1)

				If IsNumber(GroupCode(MaxIndex), 2, False) = False Then Err = Err & "GroupCode" & vbCrLf
				If IsNumber(CategoryCode(MaxIndex), 3, False) = False Then Err = Err & "CategoryCode" & vbCrLf
				If IsNumber(Code(MaxIndex), 2, False) = False Then Err = Err & "Code" & vbCrLf
			End If

			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_License 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_C_License '" & vOrderCode & "'" & vbCrLf
		If MaxIndex < 0 Then Exit Function
		For idx = 0 To UBound(Code)
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_C_License '" & vOrderCode & "'" & _
				",''" & _
				",'" & GroupCode(idx) & "'" & _
				",'" & CategoryCode(idx) & "'" & _
				",'" & Code(idx) & "'" & _
				",'" & GetYear(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsC_Note
'概　要：formで飛んできたC_Noteテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_Note
	Public OrderCode
	Public CategoryCode
	Public Code
	Public Note
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_Noteクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize(vCode)
		Dim flg	: flg = False

		MaxIndex = -1
		IsData = False
		OrderCode = GetForm("CONF_OrderCode", 1)

		CategoryCode = "Note"
		Code = vCode
		If GetForm("CONF_" & vCode, 1) <> "" Then flg = True: Note = GetForm("CONF_" & vCode, 1)

		IsData = flg

		'値チェック
		Err = ""
		If IsRE(Code, "^[A-Z]|[0-9]*$", True) = False Then Err = Err & "Code" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_Note 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_C_Note '" & vOrderCode & "'" & _
			",'" & CategoryCode & "'" & _
			",'" & Code & "'" & _
			",'" & Note & "'" & vbCrLf
	End Function
End Class

'******************************************************************************
'名　称：clsC_JobType
'概　要：formで飛んできたC_JobTypeテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_JobType
	Public OrderCode
	Public JobTypeSetNo
	Public JobTypeCode()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_JobTypeクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1

		MaxIndex = -1
		IsData = False
		OrderCode = GetForm("CONF_OrderCode", 1)
		JobTypeSetNo = GetForm("frmjobtypesetno", 1)

		If JobTypeSetNo <> "" Then IsData = True

		Do While True
			If ExistsForm("CONF_JobTypeCode"  & idx) = False Then Exit Do

			If GetForm("CONF_JobTypeCode" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve JobTypeCode(MaxIndex)	: JobTypeCode(MaxIndex) = GetForm("CONF_JobTypeCode" & idx, 1)

				'値チェック
				Err = ""
				If IsNumber(JobTypeCode(MaxIndex), 3, True) = False Then Err = Err & "JobTypeCode" & vbCrLf
			End If

			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_JobType 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

		Dim idx

		If JobTypeSetNo = "" Then
			'旧ライセンス方式
			If MaxIndex < 0 Then Exit Function

			GetRegSQL = "EXEC sp_Del_C_JobType '" & vOrderCode & "'" & vbCrLf
			For idx = 0 To UBound(JobTypeCode)
				GetRegSQL = GetRegSQL & "EXEC sp_Reg_C_JobType '" & vOrderCode & "'" & _
					",''" & _
					",'" & JobTypeCode(idx) & "'" & vbCrLf
			Next
		Else
			'新ライセンス方式
			GetRegSQL = "EXEC sp_Del_C_JobType '" & vOrderCode & "'" & vbCrLf

			sSQL = "SELECT JobTypeCode FROM vw_NaviLicense_JobType WHERE UserCode = '" & Session("userid") & "' AND JobTypeSetNo = " & JobTypeSetNo & " ORDER BY Seq "
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				GetRegSQL = GetRegSQL & "EXEC sp_Reg_C_JobType '" & vOrderCode & "'" & _
					",''" & _
					",'" & oRS.Collect("JobTypeCode") & "'" & vbCrLf
				oRS.MoveNext
			Loop
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsC_WorkingType
'概　要：formで飛んできたC_WorkingTypeテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_WorkingType
	Public OrderCode
	Public WorkingTypeCode()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_WorkingTypeクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1

		MaxIndex = -1
		IsData = False
		OrderCode = GetForm("CONF_OrderCode", 1)

		Do While True
			If ExistsForm("CONF_WorkingTypeCode" & idx) = False Then Exit Do

			If GetForm("CONF_WorkingTypeCode" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve WorkingTypeCode(MaxIndex)	: WorkingTypeCode(MaxIndex) = GetForm("CONF_WorkingTypeCode" & idx, 1)

				'値チェック
				Err = ""
				If IsNumber(WorkingTypeCode(MaxIndex), 3, True) = False Then Err = Err & "WorkingTypeCode" & vbCrLf
			End If

			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_WorkingType 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		Dim idx

		If MaxIndex < 0 Then Exit Function

		GetRegSQL = "EXEC sp_Del_C_WorkingType '" & vOrderCode & "'" & vbCrLf
		For idx = 0 To UBound(WorkingTypeCode)
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_C_WorkingType '" & vOrderCode & "'" & _
				",''" & _
				",'" & WorkingTypeCode(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsC_NaviTemp
'概　要：formで飛んできたC_NaviTempテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/03/24
'更　新：
'******************************************************************************
Class clsC_NaviTemp
	Public OrderCode
	Public CompanyName
	Public CompanyName_F
	Public EstablishYear
	Public IndustryTypeCode
	Public CapitalAmount
	Public ForeinCapital
	Public ListClass
	Public AllEmployeeNumber
	Public HomepageAddress
	Public Post_U
	Public Post_L
	Public PrefectureCode
	Public City
	Public City_F
	Public Town
	Public Address
	Public TelephoneNumber
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsC_NaviTempクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		OrderCode = GetForm("CONF_OrderCode", 1)
		If GetForm("CONF_TempCompanyName", 1) <> "" Then flg = True: CompanyName = GetForm("CONF_TempCompanyName", 1)
		If GetForm("CONF_TempCompanyName_F", 1) <> "" Then flg = True: CompanyName_F = GetForm("CONF_TempCompanyName_F", 1)
		If GetForm("CONF_TempEstablishYear", 1) <> "" Then flg = True: EstablishYear = GetForm("CONF_TempEstablishYear", 1)
		If GetForm("CONF_TempIndustryTypeCode", 1) <> "" Then flg = True: IndustryTypeCode = GetForm("CONF_TempIndustryTypeCode", 1)
		If GetForm("CONF_TempCapitalAmount", 1) <> "" Then flg = True: CapitalAmount = GetForm("CONF_TempCapitalAmount", 1)
		If GetForm("CONF_TempForeinCapital", 1) <> "" Then flg = True: ForeinCapital = GetForm("CONF_TempForeinCapital", 1)
		If GetForm("CONF_TempListClass", 1) <> "" Then flg = True: ListClass = GetForm("CONF_TempListClass", 1)
		If GetForm("CONF_TempAllEmployeeNumber", 1) <> "" Then flg = True: AllEmployeeNumber = GetForm("CONF_TempAllEmployeeNumber", 1)
		If GetForm("CONF_TempHomepageAddress", 1) <> "" Then flg = True: HomepageAddress = GetForm("CONF_TempHomepageAddress", 1)
		If GetForm("CONF_TempPost_U", 1) <> "" Then flg = True: Post_U = GetForm("CONF_TempPost_U", 1)
		If GetForm("CONF_TempPost_L", 1) <> "" Then flg = True: Post_L = GetForm("CONF_TempPost_L", 1)
		If GetForm("CONF_TempPrefectureCode", 1) <> "" Then flg = True: PrefectureCode = GetForm("CONF_TempPrefectureCode", 1)
		If GetForm("CONF_TempCity", 1) <> "" Then flg = True: City = GetForm("CONF_TempCity", 1)
		If GetForm("CONF_TempCity_F", 1) <> "" Then flg = True: City_F = GetForm("CONF_TempCity_F", 1)
		If GetForm("CONF_TempTown", 1) <> "" Then flg = True: Town = GetForm("CONF_TempTown", 1)
		If GetForm("CONF_TempAddress", 1) <> "" Then flg = True: Address = GetForm("CONF_TempAddress", 1)
		If GetForm("CONF_TempTelephoneNumber", 1) <> "" Then flg = True: TelephoneNumber = GetForm("CONF_TempTelephoneNumber", 1)

		IsData = flg

		'値チェック
		Err = ""
		If EstablishYear <> "" And IsNumber(EstablishYear, 4, False) = False Then Err = Err & "EstablishYear" & vbCrLf
		If IndustryTypeCode <> "" And IsNumber(IndustryTypeCode, 3, False) = False Then Err = Err & "IndustryTypeCode" & vbCrLf
		If AllEmployeeNumber <> "" And IsNumber(AllEmployeeNumber, 0, False) = False Then Err = Err & "AllEmployeeNumber" & vbCrLf
		If Post_U <> "" And IsNumber(Post_U, 3, False) = False Then Err = Err & "Post_U" & vbCrLf
		If Post_L <> "" And IsNumber(Post_L, 4, False) = False Then Err = Err & "Post_L" & vbCrLf
		If PrefectureCode <> "" And IsNumber(PrefectureCode, 3, False) = False Then Err = Err & "PrefectureCode" & vbCrLf
		If TelephoneNumber <> "" And IsNumber(Replace(TelephoneNumber, "-", ""), 0, False) = False Then Err = Err & "TelephoneNumber" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_C_NaviTemp 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_C_NaviTemp '" & vOrderCode & "'" & _
				",'" & WorkStartDay & "'" & _
				",'" & WorkEndDay & "'" & _
				",'" & HumanNumber & "'" & _
				",'" & CollaborationCode & "'" & _
				",'" & PermitFlag & "'" & _
				",'" & PermitDay & "'" & _
				",'" & TrafficPayFlag & "'"
		End If
	End Function
End Class

'******************************************************************************
'名　称：clsReg2
'概　要：company_reg2.aspのformを取得し、求人票作成情報を持つクラス。
'備　考：
'作成者：Lis Kokubo
'作成日：2006/05/17
'更　新：
'******************************************************************************
Class clsReg2
	Public CompanyCode
	Public JobTypeDetail	'C_Info
	Public BusinessDetail	'C_Info
	Public HopeSchoolHistoryCode	'C_Info
	Public AgeMin	'C_Info
	Public AgeMax	'C_Info
	Public AgeReasonFlag	'C_Info
	Public YearlyIncomeMin	'C_Info
	Public YearlyIncomeMax	'C_Info
	Public MonthlyIncomeMin	'C_Info
	Public MonthlyIncomeMax	'C_Info
	Public DailyIncomeMin	'C_Info
	Public DailyIncomeMax	'C_Info
	Public HourlyIncomeMin	'C_Info
	Public HourlyIncomeMax	'C_Info
	Public PercentagePayFlag
	Public IncomeRemark	'C_Info
	Public WorkingPlaceArea	'C_Info
	Public WorkingPlacePrefectureCode	'C_Info
	Public WorkingPlaceCity	'C_Info
	Public WorkingPlaceTown	'C_Info
	Public WorkingPlaceAddress	'C_Info
	Public WorkingPlaceSection	'C_Info
	Public WorkingPlaceTelephoneNumber	'C_Info
	Public TransferFlag	'C_Info
	Public WorkTimeRemark	'C_Info
	Public WeeklyHolidayType	'C_Info
	Public WorkHolidayRemark	'C_Info
	Public JobTypeSetNo	'
	Public JobTypeCode1	'C_JobType
	Public JobTypeCode2	'C_JobType
	Public JobTypeCode3	'C_JobType
	Public WorkingTypeCode1	'C_WorkingType
	Public WorkingTypeCode2	'C_WorkingType
	Public WorkingTypeCode3	'C_WorkingType
	Public WorkStartTime1	'C_WorkingCondition
	Public WorkEndTime1	'C_WorkingCondition
	Public WorkStartTime2	'C_WorkingCondition
	Public WorkEndTime2	'C_WorkingCondition
	Public WorkStartTime3	'C_WorkingCondition
	Public WorkEndTime3	'C_WorkingCondition
	Public PermitFlag	'C_Navi
	Public WorkStartDay	'C_Navi
	Public WorkEndDay	'C_Navi
	Public HumanNumber	'C_Navi
	Public TrafficPayFlag	'C_Navi
	Public SecretFlag	'C_Navi
	Public StationCode1	'C_NearbyStation
	Public ToStationRemark1	'C_NearbyStation
	Public ToStation1	'C_NearbyStation
	Public StationCode2	'C_NearbyStation
	Public ToStationRemark2	'C_NearbyStation
	Public ToStation2	'C_NearbyStation
	Public StationCode3	'C_NearbyStation
	Public ToStationRemark3	'C_NearbyStation
	Public ToStation3	'C_NearbyStation
	Public CompanySpeciality	'C_NearbyStation
	Public CatchCopy	'C_SupplementInfo
	Public PRTitle1	'C_SupplementInfo
	Public PRContents1	'C_SupplementInfo
	Public PRTitle2	'C_SupplementInfo
	Public PRContents2	'C_SupplementInfo
	Public PRTitle3	'C_SupplementInfo
	Public PRContents3	'C_SupplementInfo
	Public BizName1	'C_SupplementInfo
	Public BizPercentage1	'C_SupplementInfo
	Public BizName2	'C_SupplementInfo
	Public BizPercentage2	'C_SupplementInfo
	Public BizName3	'C_SupplementInfo
	Public BizPercentage3	'C_SupplementInfo
	Public BizName4	'C_SupplementInfo
	Public BizPercentage4	'C_SupplementInfo
	Public UITurnFlag
	Public UtilizeLanguageFlag
	Public ManyHolidayFlag
	Public InexperiencedPersonFlag
	Public FlexTimeFlag
	Public NearStationFlag
	Public NoSmokingFlag
	Public NewlyBuiltFlag
	Public LandmarkFlag
	Public RenovationFlag
	Public DesignersFlag
	Public CompanyCafeteriaFlag
	Public ShortOvertimeFlag
	Public MaternityFlag
	Public DressFreeFlag
	Public MammyFlag
	Public FixedTimeFlag
	Public ShortTimeFlag
	Public HandicappedFlag
	Public EntryInfo	'C_SupplementInfo
	Public Process1	'C_SupplementInfo
	Public Process2	'C_SupplementInfo
	Public Process3	'C_SupplementInfo
	Public Process4	'C_SupplementInfo
	Public ContactSectionName	'C_Contact
	Public ContactTelNumber	'C_Contact
	Public ContactPersonName	'C_Contact
	Public ContactPersonName_F	'C_Contact
	Public ContactPersonPost	'C_Contact
	Public ContactMailaddress	'C_Contact
	Public TempCompanyName	'C_NaviTemp
	Public TempCompanyName_F	'C_NaviTemp
	Public TempEstablishYear	'C_NaviTemp
	Public TempIndustryTypeCode	'C_NaviTemp
	Public TempCapitalAmount	'C_NaviTemp
	Public TempForeinCapital	'C_NaviTemp
	Public TempListClass	'C_NaviTemp
	Public TempAllEmployeeNumber	'C_NaviTemp
	Public TempHomepageAddress	'C_NaviTemp
	Public TempPost_U	'C_NaviTemp
	Public TempPost_L	'C_NaviTemp
	Public TempPrefectureCode	'C_NaviTemp
	Public TempCity	'C_NaviTemp
	Public TempCity_F	'C_NaviTemp
	Public TempTown	'C_NaviTemp
	Public TempAddress	'C_NaviTemp
	Public TempTelephoneNumber	'C_NaviTemp
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsReg2クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If GetForm("CONF_JobTypeDetail", 1) <> "" Then JobTypeDetail = GetForm("CONF_JobTypeDetail", 1)
		If GetForm("CONF_BusinessDetail", 1) <> "" Then BusinessDetail = GetForm("CONF_BusinessDetail", 1)
		If GetForm("CONF_HopeSchoolHistoryCode", 1) <> "" Then HopeSchoolHistoryCode = GetForm("CONF_HopeSchoolHistoryCode", 1)
		If GetForm("CONF_AgeMin", 1) <> "" Then AgeMin = GetForm("CONF_AgeMin", 1)
		If GetForm("CONF_AgeMax", 1) <> "" Then AgeMax = GetForm("CONF_AgeMax", 1)
		If GetForm("CONF_AgeReasonFlag", 1) <> "" Then AgeReasonFlag = GetForm("CONF_AgeReasonFlag", 1)
		If GetForm("CONF_YearlyIncomeMin", 1) <> "" Then YearlyIncomeMin = GetForm("CONF_YearlyIncomeMin", 1)
		If GetForm("CONF_YearlyIncomeMax", 1) <> "" Then YearlyIncomeMax = GetForm("CONF_YearlyIncomeMax", 1)
		If GetForm("CONF_MonthlyIncomeMin", 1) <> "" Then MonthlyIncomeMin = GetForm("CONF_MonthlyIncomeMin", 1)
		If GetForm("CONF_MonthlyIncomeMax", 1) <> "" Then MonthlyIncomeMax = GetForm("CONF_MonthlyIncomeMax", 1)
		If GetForm("CONF_DailyIncomeMin", 1) <> "" Then DailyIncomeMin = GetForm("CONF_DailyIncomeMin", 1)
		If GetForm("CONF_DailyIncomeMax", 1) <> "" Then DailyIncomeMax = GetForm("CONF_DailyIncomeMax", 1)
		If GetForm("CONF_HourlyIncomeMin", 1) <> "" Then HourlyIncomeMin = GetForm("CONF_HourlyIncomeMin", 1)
		If GetForm("CONF_HourlyIncomeMax", 1) <> "" Then HourlyIncomeMax = GetForm("CONF_HourlyIncomeMax", 1)
		If GetForm("CONF_PercentagePayFlag", 1) <> "" Then PercentagePayFlag = GetForm("CONF_PercentagePayFlag", 1)
		If GetForm("CONF_IncomeRemark", 1) <> "" Then IncomeRemark = GetForm("CONF_IncomeRemark", 1)
		If GetForm("CONF_WorkingPlaceArea", 1) <> "" Then WorkingPlaceArea = GetForm("CONF_WorkingPlaceArea", 1)
		If GetForm("CONF_WorkingPlacePrefectureCode", 1) <> "" Then WorkingPlacePrefectureCode = GetForm("CONF_WorkingPlacePrefectureCode", 1)
		If GetForm("CONF_WorkingPlaceCity", 1) <> "" Then WorkingPlaceCity = GetForm("CONF_WorkingPlaceCity", 1)
		If GetForm("CONF_WorkingPlaceTown", 1) <> "" Then WorkingPlaceTown = GetForm("CONF_WorkingPlaceTown", 1)
		If GetForm("CONF_WorkingPlaceAddress", 1) <> "" Then WorkingPlaceAddress = GetForm("CONF_WorkingPlaceAddress", 1)
		If GetForm("CONF_WorkingPlaceSection", 1) <> "" Then WorkingPlaceSection = GetForm("CONF_WorkingPlaceSection", 1)
		If GetForm("CONF_WorkingPlaceTelephoneNumber", 1) <> "" Then WorkingPlaceTelephoneNumber = GetForm("CONF_WorkingPlaceTelephoneNumber", 1)
		If GetForm("CONF_TransferFlag", 1) <> "" Then TransferFlag = GetForm("CONF_TransferFlag", 1)
		If GetForm("CONF_WorkTimeRemark", 1) <> "" Then WorkTimeRemark = GetForm("CONF_WorkTimeRemark", 1)
		If GetForm("CONF_WeeklyHolidayType", 1) <> "" Then WeeklyHolidayType = GetForm("CONF_WeeklyHolidayType", 1)
		If GetForm("CONF_WorkHolidayRemark", 1) <> "" Then WorkHolidayRemark = GetForm("CONF_WorkHolidayRemark", 1)
		If GetForm("frmjobtypesetno", 1) <> "" Then JobTypeSetNo = GetForm("frmjobtypesetno", 1)
		If GetForm("CONF_JobTypeCode1", 1) <> "" Then JobTypeCode1 = GetForm("CONF_JobTypeCode1", 1)
		If GetForm("CONF_JobTypeCode2", 1) <> "" Then JobTypeCode2 = GetForm("CONF_JobTypeCode2", 1)
		If GetForm("CONF_JobTypeCode3", 1) <> "" Then JobTypeCode3 = GetForm("CONF_JobTypeCode3", 1)
		If GetForm("CONF_WorkingTypeCode1", 1) <> "" Then WorkingTypeCode1 = GetForm("CONF_WorkingTypeCode1", 1)
		If GetForm("CONF_WorkingTypeCode2", 1) <> "" Then WorkingTypeCode2 = GetForm("CONF_WorkingTypeCode2", 1)
		If GetForm("CONF_WorkingTypeCode3", 1) <> "" Then WorkingTypeCode3 = GetForm("CONF_WorkingTypeCode3", 1)
		If GetForm("CONF_WorkStartTime1", 1) <> "" Then WorkStartTime1 = GetForm("CONF_WorkStartTime1", 1)
		If GetForm("CONF_WorkEndTime1", 1) <> "" Then WorkEndTime1 = GetForm("CONF_WorkEndTime1", 1)
		If GetForm("CONF_WorkStartTime2", 1) <> "" Then WorkStartTime2 = GetForm("CONF_WorkStartTime2", 1)
		If GetForm("CONF_WorkEndTime2", 1) <> "" Then WorkEndTime2 = GetForm("CONF_WorkEndTime2", 1)
		If GetForm("CONF_WorkStartTime3", 1) <> "" Then WorkStartTime3 = GetForm("CONF_WorkStartTime3", 1)
		If GetForm("CONF_WorkEndTime3", 1) <> "" Then WorkEndTime3 = GetForm("CONF_WorkEndTime3", 1)
		If GetForm("CONF_PermitFlag", 1) <> "" Then PermitFlag = GetForm("CONF_PermitFlag", 1)
		If GetForm("CONF_WorkStartDay", 1) <> "" Then WorkStartDay = GetForm("CONF_WorkStartDay", 1)
		If GetForm("CONF_WorkEndDay", 1) <> "" Then WorkEndDay = GetForm("CONF_WorkEndDay", 1)
		If GetForm("CONF_HumanNumber", 1) <> "" Then HumanNumber = GetForm("CONF_HumanNumber", 1)
		If GetForm("CONF_TrafficPayFlag", 1) <> "" Then TrafficPayFlag = GetForm("CONF_TrafficPayFlag", 1)
		If GetForm("frmsecretflag", 1) <> "" Then SecretFlag = GetForm("frmsecretflag", 1)
		If GetForm("CONF_StationCode1", 1) <> "" Then StationCode1 = GetForm("CONF_StationCode1", 1)
		If GetForm("CONF_ToStationRemark1", 1) <> "" Then ToStationRemark1 = GetForm("CONF_ToStationRemark1", 1)
		If GetForm("CONF_ToStation1", 1) <> "" Then ToStation1 = GetForm("CONF_ToStation1", 1)
		If GetForm("CONF_StationCode2", 1) <> "" Then StationCode2 = GetForm("CONF_StationCode2", 1)
		If GetForm("CONF_ToStationRemark2", 1) <> "" Then ToStationRemark2 = GetForm("CONF_ToStationRemark2", 1)
		If GetForm("CONF_ToStation2", 1) <> "" Then ToStation2 = GetForm("CONF_ToStation2", 1)
		If GetForm("CONF_StationCode3", 1) <> "" Then StationCode3 = GetForm("CONF_StationCode3", 1)
		If GetForm("CONF_ToStationRemark3", 1) <> "" Then ToStationRemark3 = GetForm("CONF_ToStationRemark3", 1)
		If GetForm("CONF_ToStation3", 1) <> "" Then ToStation3 = GetForm("CONF_ToStation3", 1)
		If GetForm("CONF_CompanySpeciality", 1) <> "" Then CompanySpeciality = GetForm("CONF_CompanySpeciality", 1)
		If GetForm("CONF_CatchCopy", 1) <> "" Then CatchCopy = GetForm("CONF_CatchCopy", 1)
		If GetForm("CONF_PRTitle1", 1) <> "" Then PRTitle1 = GetForm("CONF_PRTitle1", 1)
		If GetForm("CONF_PRContents1", 1) <> "" Then PRContents1 = GetForm("CONF_PRContents1", 1)
		If GetForm("CONF_PRTitle2", 1) <> "" Then PRTitle2 = GetForm("CONF_PRTitle2", 1)
		If GetForm("CONF_PRContents2", 1) <> "" Then PRContents2 = GetForm("CONF_PRContents2", 1)
		If GetForm("CONF_PRTitle3", 1) <> "" Then PRTitle3 = GetForm("CONF_PRTitle3", 1)
		If GetForm("CONF_PRContents3", 1) <> "" Then PRContents3 = GetForm("CONF_PRContents3", 1)
		If GetForm("CONF_BizName1", 1) <> "" Then BizName1 = GetForm("CONF_BizName1", 1)
		If GetForm("CONF_BizPercentage1", 1) <> "" Then BizPercentage1 = GetForm("CONF_BizPercentage1", 1)
		If GetForm("CONF_BizName2", 1) <> "" Then BizName2 = GetForm("CONF_BizName2", 1)
		If GetForm("CONF_BizPercentage2", 1) <> "" Then BizPercentage2 = GetForm("CONF_BizPercentage2", 1)
		If GetForm("CONF_BizName3", 1) <> "" Then BizName3 = GetForm("CONF_BizName3", 1)
		If GetForm("CONF_BizPercentage3", 1) <> "" Then BizPercentage3 = GetForm("CONF_BizPercentage3", 1)
		If GetForm("CONF_BizName4", 1) <> "" Then BizName4 = GetForm("CONF_BizName4", 1)
		If GetForm("CONF_BizPercentage4", 1) <> "" Then BizPercentage4 = GetForm("CONF_BizPercentage4", 1)
		If GetForm("CONF_UITurnFlag", 1) <> "" Then UITurnFlag = GetForm("CONF_UITurnFlag", 1)
		If GetForm("CONF_UtilizeLanguageFlag", 1) <> "" Then UtilizeLanguageFlag = GetForm("CONF_UtilizeLanguageFlag", 1)
		If GetForm("CONF_ManyHolidayFlag", 1) <> "" Then ManyHolidayFlag = GetForm("CONF_ManyHolidayFlag", 1)
		If GetForm("CONF_InexperiencedPersonFlag", 1) <> "" Then InexperiencedPersonFlag = GetForm("CONF_InexperiencedPersonFlag", 1)
		If GetForm("CONF_FlexTimeFlag", 1) <> "" Then FlexTimeFlag = GetForm("CONF_FlexTimeFlag", 1)
		If GetForm("CONF_NearStationFlag", 1) <> "" Then NearStationFlag = GetForm("CONF_NearStationFlag", 1)
		If GetForm("CONF_NoSmokingFlag", 1) <> "" Then NoSmokingFlag = GetForm("CONF_NoSmokingFlag", 1)
		If GetForm("CONF_NewlyBuiltFlag", 1) <> "" Then NewlyBuiltFlag = GetForm("CONF_NewlyBuiltFlag", 1)
		If GetForm("CONF_LandmarkFlag", 1) <> "" Then LandmarkFlag = GetForm("CONF_LandmarkFlag", 1)
		If GetForm("CONF_RenovationFlag", 1) <> "" Then RenovationFlag = GetForm("CONF_RenovationFlag", 1)
		If GetForm("CONF_DesignersFlag", 1) <> "" Then DesignersFlag = GetForm("CONF_DesignersFlag", 1)
		If GetForm("CONF_CompanyCafeteriaFlag", 1) <> "" Then CompanyCafeteriaFlag = GetForm("CONF_CompanyCafeteriaFlag", 1)
		If GetForm("CONF_ShortOvertimeFlag", 1) <> "" Then ShortOvertimeFlag = GetForm("CONF_ShortOvertimeFlag", 1)
		If GetForm("CONF_MaternityFlag", 1) <> "" Then MaternityFlag = GetForm("CONF_MaternityFlag", 1)
		If GetForm("CONF_DressFreeFlag", 1) <> "" Then DressFreeFlag = GetForm("CONF_DressFreeFlag", 1)
		If GetForm("CONF_MammyFlag", 1) <> "" Then MammyFlag = GetForm("CONF_MammyFlag", 1)
		If GetForm("CONF_FixedTimeFlag", 1) <> "" Then FixedTimeFlag = GetForm("CONF_FixedTimeFlag", 1)
		If GetForm("CONF_ShortTimeFlag", 1) <> "" Then ShortTimeFlag = GetForm("CONF_ShortTimeFlag", 1)
		If GetForm("CONF_HandicappedFlag", 1) <> "" Then HandicappedFlag = GetForm("CONF_HandicappedFlag", 1)
		If GetForm("CONF_EntryInfo", 1) <> "" Then EntryInfo = GetForm("CONF_EntryInfo", 1)
		If GetForm("CONF_Process1", 1) <> "" Then Process1 = GetForm("CONF_Process1", 1)
		If GetForm("CONF_Process2", 1) <> "" Then Process2 = GetForm("CONF_Process2", 1)
		If GetForm("CONF_Process3", 1) <> "" Then Process3 = GetForm("CONF_Process3", 1)
		If GetForm("CONF_Process4", 1) <> "" Then Process4 = GetForm("CONF_Process4", 1)
		If GetForm("CONF_ContactSectionName", 1) <> "" Then ContactSectionName = GetForm("CONF_ContactSectionName", 1)
		If GetForm("CONF_ContactTelNumber", 1) <> "" Then ContactTelNumber = GetForm("CONF_ContactTelNumber", 1)
		If GetForm("CONF_ContactPersonName", 1) <> "" Then ContactPersonName = GetForm("CONF_ContactPersonName", 1)
		If GetForm("CONF_ContactPersonName_F", 1) <> "" Then ContactPersonName_F = GetForm("CONF_ContactPersonName_F", 1)
		If GetForm("CONF_ContactPersonPost", 1) <> "" Then ContactPersonPost = GetForm("CONF_ContactPersonPost", 1)
		If GetForm("CONF_ContactMailaddress", 1) <> "" Then ContactMailaddress = GetForm("CONF_ContactMailaddress", 1)
		If GetForm("CONF_TempCompanyName", 1) <> "" Then TempCompanyName = GetForm("CONF_TempCompanyName", 1)
		If GetForm("CONF_TempCompanyName_F", 1) <> "" Then TempCompanyName_F = GetForm("CONF_TempCompanyName_F", 1)
		If GetForm("CONF_TempEstablishYear", 1) <> "" Then TempEstablishYear = GetForm("CONF_TempEstablishYear", 1)
		If GetForm("CONF_TempIndustryTypeCode", 1) <> "" Then TempIndustryTypeCode = GetForm("CONF_TempIndustryTypeCode", 1)
		If GetForm("CONF_TempCapitalAmount", 1) <> "" Then TempCapitalAmount = GetForm("CONF_TempCapitalAmount", 1)
		If GetForm("CONF_TempForeinCapital", 1) <> "" Then TempForeinCapital = GetForm("CONF_TempForeinCapital", 1)
		If GetForm("CONF_TempListClass", 1) <> "" Then TempListClass = GetForm("CONF_TempListClass", 1)
		If GetForm("CONF_TempAllEmployeeNumber", 1) <> "" Then TempAllEmployeeNumber = GetForm("CONF_TempAllEmployeeNumber", 1)
		If GetForm("CONF_TempHomepageAddress", 1) <> "" Then TempHomepageAddress = GetForm("CONF_TempHomepageAddress", 1)
		If GetForm("CONF_TempPost_U", 1) <> "" Then TempPost_U = GetForm("CONF_TempPost_U", 1)
		If GetForm("CONF_TempPost_L", 1) <> "" Then TempPost_L = GetForm("CONF_TempPost_L", 1)
		If GetForm("CONF_TempPrefectureCode", 1) <> "" Then TempPrefectureCode = GetForm("CONF_TempPrefectureCode", 1)
		If GetForm("CONF_TempCity", 1) <> "" Then TempCity = GetForm("CONF_TempCity", 1)
		If GetForm("CONF_TempCity_F", 1) <> "" Then TempCity_F = GetForm("CONF_TempCity_F", 1)
		If GetForm("CONF_TempTown", 1) <> "" Then TempTown = GetForm("CONF_TempTown", 1)
		If GetForm("CONF_TempAddress", 1) <> "" Then TempAddress = GetForm("CONF_TempAddress", 1)
		If GetForm("CONF_TempTelephoneNumber", 1) <> "" Then TempTelephoneNumber = GetForm("CONF_TempTelephoneNumber", 1)

		If JobTypeDetail <> "" And (JobTypeCode1 & JobTypeCode2 & JobTypeCode3 & JobTypeSetNo) <> "" _
		And (WorkingTypeCode1 & WorkingTypeCode2 & WorkingTypeCode3) <> "" _
		And (WorkStartTime1 & WorkEndTime1 <> "" Or WorkStartTime2 & WorkEndTime2 <> "" Or WorkStartTime3 & WorkEndTime3 <> "") _
		And ContactTelNumber <> "" And ContactPersonName <> "" And ContactPersonName_F <> "" And ContactMailAddress <> "" Then
			IsData = True
		End If

		'値チェック
		Err = ""
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_Reg2 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vOrderCode)
		GetRegSQL = ""
		If IsData = True Then
			GetRegSQL = "sp_Reg_Navi '" & vOrderCode & "'" & _
				",'" & Session("userid") & "'" & _
				",'" & JobTypeDetail & "'" & _
				",'" & BusinessDetail & "'" & _
				",'" & HopeSchoolHistoryCode & "'" & _
				",'" & AgeMin & "'" & _
				",'" & AgeMax & "'" & _
				",'" & AgeReasonFlag & "'" & _
				",'" & YearlyIncomeMin & "'" & _
				",'" & YearlyIncomeMax & "'" & _
				",'" & MonthlyIncomeMin & "'" & _
				",'" & MonthlyIncomeMax & "'" & _
				",'" & DailyIncomeMin & "'" & _
				",'" & DailyIncomeMax & "'" & _
				",'" & HourlyIncomeMin & "'" & _
				",'" & HourlyIncomeMax & "'" & _
				",'" & PercentagePayFlag & "'" & _
				",'" & IncomeRemark & "'" & _
				",'" & WorkingPlaceArea & "'" & _
				",'" & WorkingPlacePrefectureCode & "'" & _
				",'" & WorkingPlaceCity & "'" & _
				",'" & WorkingPlaceTown & "'" & _
				",'" & WorkingPlaceAddress & "'" & _
				",'" & WorkingPlaceSection & "'" & _
				",'" & WorkingPlaceTelephoneNumber & "'" & _
				",'" & TransferFlag & "'" & _
				",'" & WorkTimeRemark & "'" & _
				",'" & WeeklyHolidayType & "'" & _
				",'" & WorkHolidayRemark & "'" & _
				",'" & PermitFlag & "'" & _
				",'" & WorkStartDay & "'" & _
				",'" & WorkEndDay & "'" & _
				",'" & HumanNumber & "'" & _
				",'" & TrafficPayFlag & "'" & _
				",'" & SecretFlag & "'" & _
				",'" & CompanySpeciality & "'" & _
				",'" & CatchCopy & "'" & _
				",'" & PRTitle1 & "'" & _
				",'" & PRContents1 & "'" & _
				",'" & PRTitle2 & "'" & _
				",'" & PRContents2 & "'" & _
				",'" & PRTitle3 & "'" & _
				",'" & PRContents3 & "'" & _
				",'" & BizName1 & "'" & _
				",'" & BizPercentage1 & "'" & _
				",'" & BizName2 & "'" & _
				",'" & BizPercentage2 & "'" & _
				",'" & BizName3 & "'" & _
				",'" & BizPercentage3 & "'" & _
				",'" & BizName4 & "'" & _
				",'" & BizPercentage4 & "'" & _
				",'" & UITurnFlag & "'" & _
				",'" & UtilizeLanguageFlag & "'" & _
				",'" & ManyHolidayFlag & "'" & _
				",'" & InexperiencedPersonFlag & "'" & _
				",'" & FlexTimeFlag & "'" & _
				",'" & NearStationFlag & "'" & _
				",'" & NoSmokingFlag & "'" & _
				",'" & NewlyBuiltFlag & "'" & _
				",'" & LandmarkFlag & "'" & _
				",'" & RenovationFlag & "'" & _
				",'" & DesignersFlag & "'" & _
				",'" & CompanyCafeteriaFlag & "'" & _
				",'" & ShortOvertimeFlag & "'" & _
				",'" & MaternityFlag & "'" & _
				",'" & DressFreeFlag & "'" & _
				",'" & MammyFlag & "'" & _
				",'" & FixedTimeFlag & "'" & _
				",'" & ShortTimeFlag & "'" & _
				",'" & HandicappedFlag & "'" & _
				",'" & EntryInfo & "'" & _
				",'" & Process1 & "'" & _
				",'" & Process2 & "'" & _
				",'" & Process3 & "'" & _
				",'" & Process4 & "'" & _
				",'" & ContactSectionName & "'" & _
				",'" & ContactTelNumber & "'" & _
				",'" & ContactPersonName & "'" & _
				",'" & ContactPersonName_F & "'" & _
				",'" & ContactPersonPost & "'" & _
				",'" & ContactMailaddress & "'" & _
				",'" & TempCompanyName & "'" & _
				",'" & TempCompanyName_F & "'" & _
				",'" & TempEstablishYear & "'" & _
				",'" & TempIndustryTypeCode & "'" & _
				",'" & TempCapitalAmount & "'" & _
				",'" & TempForeinCapital & "'" & _
				",'" & TempListClass	 & "'" & _
				",'" & TempAllEmployeeNumber & "'" & _
				",'" & TempHomepageAddress & "'" & _
				",'" & TempPost_U & "'" & _
				",'" & TempPost_L & "'" & _
				",'" & TempPrefectureCode & "'" & _
				",'" & TempCity & "'" & _
				",'" & TempCity_F & "'" & _
				",'" & TempTown & "'" & _
				",'" & TempAddress & "'" & _
				",'" & TempTelephoneNumber & "'" & vbCrLf
		End If
	End Function
End Class
%>
