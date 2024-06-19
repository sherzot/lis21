<%
'******************************************************************************
'概　要：スタッフテーブル群にデータをInsert, Updateする時に
'　　　：formで飛んできたデータを格納するためのクラス群
'備　考：事前に commonfunc.asp をインクルードしておくこと！
'作成者：Lis Kokubo
'作成日：2006/03/30
'更　新：
'******************************************************************************
'CONF_WorkPeriodTypeFlag
'CONF_HopeMonthPeriod
'CONF_Image
'CONF_StaffCode
%>
<%
'******************************************************************************
'名　称：clsP_UserInfo
'概　要：formで飛んできたP_UserInfoテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_UserInfo
	Public StaffCode
	Public Password
	Public OperateClassComCode
	Public OperateClassWebCode
	Public OperateClassRemark
	Public BranchCode
	Public EmployeeCode
	Public TempFlag
	Public IntroductionFlag
	Public TempToPermFlag
	Public MailMagazineFlag
	Public NewJohoMailFlag
	Public SuspensionFlag
	Public ErasureFlag
	Public OfferFlag
	Public HopeUseFlag
	Public NaviUseFlag
	Public HomeContactFlag
	Public PortableContactFlag
	Public FaxContactFlag
	Public MailContactFlag
	Public ReferRejectFlag
	Public PersonDangerFlag
	Public PriorityJobTypeFlag
	Public PriorityIndustryTypeFlag
	Public PriorityWorkingTypeFlag
	Public PriorityWorkingPlaceFlag
	Public PriorityStationFlag
	Public PriorityWorkingTimeFlag
	Public PrioritySalaryFlag
	Public HopeCommuteTime
	Public LisReserveDay
	Public LisRegistDay
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_UserInfoクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If GetForm("CONF_StaffCode", 1) <> "" Then IsData = True: StaffCode = GetForm("CONF_StaffCode", 1)
		If GetForm("CONF_Password", 1) <> "" Then IsData = True: Password = GetForm("CONF_Password", 1)
		If GetForm("CONF_OperateClassComCode", 1) <> "" Then IsData = True: OperateClassComCode = GetForm("CONF_OperateClassComCode", 1)
		If GetForm("CONF_OperateClassWebCode", 1) <> "" Then IsData = True: OperateClassWebCode = GetForm("CONF_OperateClassWebCode", 1)
		If GetForm("CONF_OperateClassRemark", 1) <> "" Then IsData = True: OperateClassRemark = GetForm("CONF_OperateClassRemark", 1)
		If GetForm("CONF_BranchCode", 1) <> "" Then IsData = True: BranchCode = GetForm("CONF_BranchCode", 1)
		If GetForm("CONF_EmployeeCode", 1) <> "" Then IsData = True: EmployeeCode = GetForm("CONF_EmployeeCode", 1)
		If GetForm("CONF_TempFlag", 1) <> "" Then IsData = True: TempFlag = GetForm("CONF_TempFlag", 1)
		If GetForm("CONF_IntroductionFlag", 1) <> "" Then IsData = True: IntroductionFlag = GetForm("CONF_IntroductionFlag", 1)
		If GetForm("CONF_TempToPermFlag", 1) <> "" Then IsData = True: TempToPermFlag = GetForm("CONF_TempToPermFlag", 1)
		If GetForm("CONF_MailMagazineFlag", 1) <> "" Then IsData = True: MailMagazineFlag = GetForm("CONF_MailMagazineFlag", 1)
		If GetForm("CONF_NewJohoMailFlag", 1) <> "" Then IsData = True: NewJohoMailFlag = GetForm("CONF_NewJohoMailFlag", 1)
		If GetForm("CONF_SuspensionFlag", 1) <> "" Then IsData = True: SuspensionFlag = GetForm("CONF_SuspensionFlag", 1)
		If GetForm("CONF_ErasureFlag", 1) <> "" Then IsData = True: ErasureFlag = GetForm("CONF_ErasureFlag", 1)
		If GetForm("CONF_OfferFlag", 1) <> "" Then IsData = True: OfferFlag = GetForm("CONF_OfferFlag", 1)
		If GetForm("CONF_HopeUseFlag", 1) <> "" Then IsData = True: HopeUseFlag = GetForm("CONF_HopeUseFlag", 1)
		If GetForm("CONF_NaviUseFlag", 1) <> "" Then IsData = True: NaviUseFlag = GetForm("CONF_NaviUseFlag", 1)
		If GetForm("CONF_HomeContactFlag", 1) <> "" Then IsData = True: HomeContactFlag = GetForm("CONF_HomeContactFlag", 1)
		If GetForm("CONF_PortableContactFlag", 1) <> "" Then IsData = True: PortableContactFlag = GetForm("CONF_PortableContactFlag", 1)
		If GetForm("CONF_FaxContactFlag", 1) <> "" Then IsData = True: FaxContactFlag = GetForm("CONF_FaxContactFlag", 1)
		If GetForm("CONF_MailContactFlag", 1) <> "" Then IsData = True: MailContactFlag = GetForm("CONF_MailContactFlag", 1)
		If GetForm("CONF_ReferRejectFlag", 1) <> "" Then IsData = True: ReferRejectFlag = GetForm("CONF_ReferRejectFlag", 1)
		If GetForm("CONF_PersonDangerFlag", 1) <> "" Then IsData = True: PersonDangerFlag = GetForm("CONF_PersonDangerFlag", 1)
		If GetForm("CONF_PriorityJobTypeFlag", 1) <> "" Then IsData = True: PriorityJobTypeFlag = GetForm("CONF_PriorityJobTypeFlag", 1)
		If GetForm("CONF_PriorityIndustryTypeFlag", 1) <> "" Then IsData = True: PriorityIndustryTypeFlag = GetForm("CONF_PriorityIndustryTypeFlag", 1)
		If GetForm("CONF_PriorityWorkingTypeFlag", 1) <> "" Then IsData = True: PriorityWorkingTypeFlag = GetForm("CONF_PriorityWorkingTypeFlag", 1)
		If GetForm("CONF_PriorityWorkingPlaceFlag", 1) <> "" Then IsData = True: PriorityWorkingPlaceFlag = GetForm("CONF_PriorityWorkingPlaceFlag", 1)
		If GetForm("CONF_PriorityStationFlag", 1) <> "" Then IsData = True: PriorityStationFlag = GetForm("CONF_PriorityStationFlag", 1)
		If GetForm("CONF_PriorityWorkingTimeFlag", 1) <> "" Then IsData = True: PriorityWorkingTimeFlag = GetForm("CONF_PriorityWorkingTimeFlag", 1)
		If GetForm("CONF_PrioritySalaryFlag", 1) <> "" Then IsData = True: PrioritySalaryFlag = GetForm("CONF_PrioritySalaryFlag", 1)
		If GetForm("CONF_HopeCommuteTime", 1) <> "" Then IsData = True: HopeCommuteTime = GetForm("CONF_HopeCommuteTime", 1)
		If GetForm("CONF_LisReserveDay", 1) <> "" Then IsData = True: LisReserveDay = GetForm("CONF_LisReserveDay", 1)
		If GetForm("CONF_LisRegistDay", 1) <> "" Then IsData = True: LisRegistDay = GetForm("CONF_LisRegistDay", 1)

		'値チェック
		Err = ""
		If StaffCode <> "" And IsMainCode(StaffCode) = False Then Err = Err & "StaffCode" & vbCrLf
		If OperateClassComCode <> "" And IsNumber(OperateClassComCode, 3, False) = False Then Err = Err & "OperateClassComCode" & vbCrLf
		If OperateClassWebCode <> "" And IsNumber(OperateClassWebCode, 3, False) = False Then Err = Err & "OperateClassWebCode" & vbCrLf
		If BranchCode <> "" And IsRE(BranchCode, "^[A-Z][A-Z]$", True) = False Then Err = Err & "BranchCode" & vbCrLf
		If EmployeeCode <> "" And IsMainCode(EmployeeCode) = False Then Err = Err & "EmployeeCode" & vbCrLf
		If TempFlag <> "" And IsFlag(TempFlag) = False Then Err = Err & "TempFlag" & vbCrLf
		If IntroductionFlag <> "" And IsFlag(IntroductionFlag) = False Then Err = Err & "IntroductionFlag" & vbCrLf
		If TempToPermFlag <> "" And IsFlag(TempToPermFlag) = False Then Err = Err & "TempToPermFlag" & vbCrLf
		If MailMagazineFlag <> "" And IsFlag(MailMagazineFlag) = False Then Err = Err & "MailMagazineFlag" & vbCrLf
		If NewJohoMailFlag <> "" And IsFlag(NewJohoMailFlag) = False Then Err = Err & "NewJohoMailFlag" & vbCrLf
		If SuspensionFlag <> "" And IsFlag(SuspensionFlag) = False Then Err = Err & "SuspensionFlag" & vbCrLf
		If ErasureFlag <> "" And IsFlag(ErasureFlag) = False Then Err = Err & "ErasureFlag" & vbCrLf
		If OfferFlag <> "" And IsFlag(OfferFlag) = False Then Err = Err & "OfferFlag" & vbCrLf
		If HopeUseFlag <> "" And IsFlag(HopeUseFlag) = False Then Err = Err & "HopeUseFlag" & vbCrLf
		If NaviUseFlag <> "" And IsFlag(NaviUseFlag) = False Then Err = Err & "NaviUseFlag" & vbCrLf
		If HomeContactFlag <> "" And IsFlag(HomeContactFlag) = False Then Err = Err & "HomeContactFlag" & vbCrLf
		If PortableContactFlag <> "" And IsFlag(PortableContactFlag) = False Then Err = Err & "PortableContactFlag" & vbCrLf
		If FaxContactFlag <> "" And IsFlag(FaxContactFlag) = False Then Err = Err & "FaxContactFlag" & vbCrLf
		If MailContactFlag <> "" And IsFlag(MailContactFlag) = False Then Err = Err & "MailContactFlag" & vbCrLf
		If ReferRejectFlag <> "" And IsFlag(ReferRejectFlag) = False Then Err = Err & "ReferRejectFlag" & vbCrLf
		If PersonDangerFlag <> "" And IsFlag(PersonDangerFlag) = False Then Err = Err & "PersonDangerFlag" & vbCrLf
		If PriorityJobTypeFlag <> "" And IsFlag(PriorityJobTypeFlag) = False Then Err = Err & "PriorityJobTypeFlag" & vbCrLf
		If PriorityIndustryTypeFlag <> "" And IsFlag(PriorityIndustryTypeFlag) = False Then Err = Err & "PriorityIndustryTypeFlag" & vbCrLf
		If PriorityWorkingTypeFlag <> "" And IsFlag(PriorityWorkingTypeFlag) = False Then Err = Err & "PriorityWorkingTypeFlag" & vbCrLf
		If PriorityWorkingPlaceFlag <> "" And IsFlag(PriorityWorkingPlaceFlag) = False Then Err = Err & "PriorityWorkingPlaceFlag" & vbCrLf
		If PriorityStationFlag <> "" And IsFlag(PriorityStationFlag) = False Then Err = Err & "PriorityStationFlag" & vbCrLf
		If PriorityWorkingTimeFlag <> "" And IsFlag(PriorityWorkingTimeFlag) = False Then Err = Err & "PriorityWorkingTimeFlag" & vbCrLf
		If PrioritySalaryFlag <> "" And IsFlag(PrioritySalaryFlag) = False Then Err = Err & "PrioritySalaryFlag" & vbCrLf
		If HopeCommuteTime <> "" And IsNumber(HopeCommuteTime, 0, False) = False Then Err = Err & "HopeCommuteTime" & vbCrLf
		If LisReserveDay <> "" And IsDay(LisReserveDay) = False Then Err = Err & "LisReserveDay" & vbCrLf
		If LisRegistDay <> "" And IsDay(LisRegistDay) = False Then Err = Err & "LisRegistDay" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_ UserInfo実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "sp_Reg_P_UserInfo '" & vStaffCode & "'" & _
			",'S'" & _
			",'" & Password & "'" & _
			",'" & OperateClassComCode & "'" & _
			",'" & OperateClassWebCode & "'" & _
			",'" & OperateClassRemark & "'" & _
			",'" & BranchCode & "'" & _
			",'" & EmployeeCode & "'" & _
			",'" & TempFlag & "'" & _
			",'" & IntroductionFlag & "'" & _
			",'" & TempToPermFlag & "'" & _
			",'" & MailMagazineFlag & "'" & _
			",'" & NewJohoMailFlag & "'" & _
			",'" & SuspensionFlag & "'" & _
			",'" & ErasureFlag & "'" & _
			",'" & OfferFlag & "'" & _
			",'" & HopeUseFlag & "'" & _
			",'" & NaviUseFlag & "'" & _
			",'" & HomeContactFlag & "'" & _
			",'" & PortableContactFlag & "'" & _
			",'" & FaxContactFlag & "'" & _
			",'" & MailContactFlag & "'" & _
			",'" & ReferRejectFlag & "'" & _
			",'" & PersonDangerFlag & "'" & _
			",'" & PriorityJobTypeFlag & "'" & _
			",'" & PriorityIndustryTypeFlag & "'" & _
			",'" & PriorityWorkingTypeFlag & "'" & _
			",'" & PriorityWorkingPlaceFlag & "'" & _
			",'" & PriorityStationFlag & "'" & _
			",'" & PriorityWorkingTimeFlag & "'" & _
			",'" & PrioritySalaryFlag & "'" & _
			",'" & HopeCommuteTime & "'" & _
			",'" & LisReserveDay & "'" & _
			",'" & LisRegistDay & "'"
	End Function
End Class

'******************************************************************************
'名　称：clsP_Info
'概　要：formで飛んできたP_Infoテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_Info
	Public StaffCode
	Public Name
	Public Name_F
	Public SearchName
	Public SearchName_F
	Public OldName
	Public Birthday
	Public Sex
	Public MarriageFlag
	Public Post_U
	Public Post_L
	Public PrefectureCode
	Public City
	Public City_F
	Public Town
	Public Town_F
	Public Address
	Public Address_F
	Public LivingType
	Public HomeTelephoneNumber
	Public CountryTelephoneNumber
	Public PortableTelephoneNumber
	Public FaxNumber
	Public MailAddress
	Public PortableMailAddress
	Public URL
	Public InfoSourceType
	Public InfoSourceDay
	Public InfoSourceOther
	Public DependentFlag
	Public DependentNumber
	Public SpouseFlag
	Public CurrentCompanyName
	Public CurrentCompanyName_F
	Public SocietyInsuranceIn
	Public SocietyInsuranceLoss
	Public EmployInsuranceIn
	Public EmployInsuranceLoss
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_Info クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If GetForm("CONF_StaffCode", 1) <> "" Then StaffCode = GetForm("CONF_StaffCode", 1)
		If GetForm("CONF_Name_1", 1) <> "" And GetForm("CONF_Name_2", 1) <> "" Then IsData = True: Name = GetForm("CONF_Name_1", 1) & "　" & GetForm("CONF_Name_2", 1)
		If GetForm("CONF_Name_F_1", 1) <> "" And GetForm("CONF_Name_F_2", 1) <> "" Then IsData = True: Name_F = GetForm("CONF_Name_F_1", 1) & "　" & GetForm("CONF_Name_F_2", 1)
		If GetForm("CONF_Name_1", 1) <> "" And GetForm("CONF_Name_2", 1) <> "" Then IsData = True: SearchName = GetForm("CONF_Name_1", 1) & GetForm("CONF_Name_2", 1)
		If GetForm("CONF_Name_F_1", 1) <> "" And GetForm("CONF_Name_F_2", 1) <> "" Then IsData = True: SearchName_F = GetForm("CONF_Name_F_1", 1) & GetForm("CONF_Name_F_2", 1)
		If GetForm("CONF_OldName", 1) <> "" Then IsData = True: OldName = GetForm("CONF_OldName", 1)
		If GetForm("CONF_Birthday", 1) <> "" Then IsData = True: Birthday = GetForm("CONF_Birthday", 1)
		If GetForm("CONF_Sex", 1) <> "" Then IsData = True: Sex = GetForm("CONF_Sex", 1)
		If GetForm("CONF_MarriageFlag", 1) <> "" Then IsData = True: MarriageFlag = GetForm("CONF_MarriageFlag", 1)
		If GetForm("CONF_Post", 1) <> "" Then IsData = True: Post_U = Mid(GetForm("CONF_Post", 1), 1, 3)
		If GetForm("CONF_Post", 1) <> "" Then IsData = True: Post_L = Mid(GetForm("CONF_Post", 1), 4, 4)
		If GetForm("CONF_PrefectureCode", 1) <> "" Then IsData = True: PrefectureCode = GetForm("CONF_PrefectureCode", 1)
		If GetForm("CONF_City", 1) <> "" Then IsData = True: City = GetForm("CONF_City", 1)
		If GetForm("CONF_City_F", 1) <> "" Then IsData = True: City_F = GetForm("CONF_City_F", 1)
		If GetForm("CONF_Town", 1) <> "" Then IsData = True: Town = GetForm("CONF_Town", 1)
		If GetForm("CONF_Town_F", 1) <> "" Then IsData = True: Town_F = GetForm("CONF_Town_F", 1)	'ねぇぞ～
		If GetForm("CONF_Address", 1) <> "" Then IsData = True: Address = GetForm("CONF_Address", 1)
		If GetForm("CONF_Address_F", 1) <> "" Then IsData = True: Address_F = GetForm("CONF_Address_F", 1)	'ねぇぞ～
		If GetForm("CONF_LivingType", 1) <> "" Then IsData = True: LivingType = GetForm("CONF_LivingType", 1)
		If GetForm("CONF_HomeTelephoneNumber", 1) <> "" Then IsData = True: HomeTelephoneNumber = GetForm("CONF_HomeTelephoneNumber", 1)
		If GetForm("CONF_CountryTelephoneNumber", 1) <> "" Then IsData = True: CountryTelephoneNumber = GetForm("CONF_CountryTelephoneNumber", 1)
		If GetForm("CONF_PortableTelephoneNumber", 1) <> "" Then IsData = True: PortableTelephoneNumber = GetForm("CONF_PortableTelephoneNumber", 1)
		If GetForm("CONF_FaxNumber", 1) <> "" Then IsData = True: FaxNumber = GetForm("CONF_FaxNumber", 1)
		If GetForm("CONF_MailAddress", 1) <> "" Then IsData = True: MailAddress = GetForm("CONF_MailAddress", 1)
		If GetForm("CONF_PortableMailAddress", 1) <> "" Then IsData = True: PortableMailAddress = GetForm("CONF_PortableMailAddress", 1)
		If GetForm("CONF_URL", 1) <> "" Then IsData = True: URL = GetForm("CONF_URL", 1)
		If GetForm("CONF_InfoSource", 1) <> "" Then IsData = True: InfoSourceType = GetForm("CONF_InfoSource", 1)	'ねぇぞ～
		If GetForm("CONF_InfoSourceDay", 1) <> "" Then IsData = True: InfoSourceDay = GetForm("CONF_InfoSourceDay", 1)
		If GetForm("CONF_InfoSourceOther", 1) <> "" Then IsData = True: InfoSourceOther = GetForm("CONF_InfoSourceOther", 1)
		If GetForm("CONF_DependentFlag", 1) <> "" Then IsData = True: DependentFlag = GetForm("CONF_DependentFlag", 1)
		If GetForm("CONF_DependentNum", 1) <> "" Then IsData = True: DependentNumber = GetForm("CONF_DependentNum", 1)
		If GetForm("CONF_SpouseFlag", 1) <> "" Then IsData = True: SpouseFlag = GetForm("CONF_SpouseFlag", 1)
		If GetForm("CONF_CurrentCompanyName", 1) <> "" Then IsData = True: CurrentCompanyName = GetForm("CONF_CurrentCompanyName", 1)
		If GetForm("CONF_CurrentCompanyName_F", 1) <> "" Then IsData = True: CurrentCompanyName_F = GetForm("CONF_CurrentCompanyName_F", 1)
		If GetForm("CONF_SocietyInsuranceIn", 1) <> "" Then IsData = True: SocietyInsuranceIn = GetForm("CONF_SocietyInsuranceIn", 1)
		If GetForm("CONF_SocietyInsuranceLoss", 1) <> "" Then IsData = True: SocietyInsuranceLoss = GetForm("CONF_SocietyInsuranceLoss", 1)
		If GetForm("CONF_EmployInsuranceIn", 1) <> "" Then IsData = True: EmployInsuranceIn = GetForm("CONF_EmployInsuranceIn", 1)
		If GetForm("CONF_EmployInsuranceLoss", 1) <> "" Then IsData = True: EmployInsuranceLoss = GetForm("CONF_EmployInsuranceLoss", 1)

		'値チェック
		Err = ""
		If StaffCode <> "" And IsMainCode(StaffCode) = False Then Err = Err & "StaffCode" & vbCrLf
		If Birthday <> "" And IsDay(Birthday) = False Then Err = Err & "Birthday" & vbCrLf
		If Sex <> "" And IsRE(Sex, "^[12]$", True) = False Then Err = Err & "Sex" & vbCrLf
		If MarriageFlag <> "" And IsFlag(MarriageFlag) = False Then Err = Err & "MarriageFlag" & vbCrLf
		If Post_U <> "" And IsNumber(Post_U, 3, False) = False Then Err = Err & "Post_U" & vbCrLf
		If Post_L <> "" And IsNumber(Post_L, 4, False) = False Then Err = Err & "Post_L" & vbCrLf
		If PrefectureCode <> "" And IsNumber(PrefectureCode, 3, False) = False Then Err = Err & "PrefectureCode" & vbCrLf
		If LivingType <> "" And IsRE(LivingType, "^[1234]$", True) = False Then Err = Err & "LivingType" & vbCrLf
		If HomeTelephoneNumber <> "" And IsNumber(HomeTelephoneNumber, 0, False) = False Then Err = Err & "HomeTelephoneNumber" & vbCrLf
		If CountryTelephoneNumber <> "" And IsNumber(CountryTelephoneNumber, 0, False) = False Then Err = Err & "CountryTelephoneNumber" & vbCrLf
		If PortableTelephoneNumber <> "" And IsNumber(PortableTelephoneNumber, 0, False) = False Then Err = Err & "PortableTelephoneNumber" & vbCrLf
		If FaxNumber <> "" And IsNumber(FaxNumber, 0, False) = False Then Err = Err & "FaxNumber" & vbCrLf
		If InfoSourceType <> "" And IsNumber(InfoSourceType, 3, False) = False Then Err = Err & "InfoSourceType" & vbCrLf
		If InfoSourceDay <> "" And IsDay(InfoSourceDay) = False Then Err = Err & "InfoSourceDay" & vbCrLf
		If DependentFlag <> "" And IsFlag(DependentFlag) = False Then Err = Err & "DependentFlag" & vbCrLf
		If DependentNumber <> "" And IsNumber(DependentNumber, 0, False) = False Then Err = Err & "DependentNumber" & vbCrLf
		If SpouseFlag <> "" And IsFlag(SpouseFlag) = False Then Err = Err & "SpouseFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_Info 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "sp_Reg_P_Info" & _
			" '" & vStaffCode & "'" & _
			",'" & Name & "'" & _
			",'" & Name_F & "'" & _
			",'" & SearchName & "'" & _
			",'" & SearchName_F & "'" & _
			",'" & OldName & "'" & _
			",'" & Birthday & "'" & _
			",'" & Sex & "'" & _
			",'" & MarriageFlag & "'" & _
			",'" & Post_U & "'" & _
			",'" & Post_L & "'" & _
			",'" & PrefectureCode & "'" & _
			",'" & City & "'" & _
			",'" & City_F & "'" & _
			",'" & Town & "'" & _
			",'" & Town_F & "'" & _
			",'" & Address & "'" & _
			",'" & Address_F & "'" & _
			",'" & LivingType & "'" & _
			",'" & HomeTelephoneNumber & "'" & _
			",'" & CountryTelephoneNumber & "'" & _
			",'" & PortableTelephoneNumber & "'" & _
			",'" & FaxNumber & "'" & _
			",'" & MailAddress & "'" & _
			",'" & PortableMailAddress & "'" & _
			",'" & URL & "'" & _
			",'" & InfoSourceType & "'" & _
			",'" & InfoSourceDay & "'" & _
			",'" & InfoSourceOther & "'" & _
			",'" & DependentFlag & "'" & _
			",'" & DependentNumber & "'" & _
			",'" & SpouseFlag & "'" & _
			",'" & CurrentCompanyName & "'" & _
			",'" & CurrentCompanyName_F & "'" & _
			",'" & SocietyInsuranceIn & "'" & _
			",'" & SocietyInsuranceLoss & "'" & _
			",'" & EmployInsuranceIn & "'" & _
			",'" & EmployInsuranceLoss & "'" & vbCrLf
	End Function
End Class

'******************************************************************************
'名　称：clsP_NearbyStation
'概　要：formで飛んできたP_NearbyStationテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_NearbyStation
	Public StaffCode
	Public StationCode()
'	Public StationName()
	Public ToStationBusFlag()
	Public ToStationCarFlag()
	Public ToStationBicycleFlag()
	Public ToStationWalkFlag()
	Public OtherTransportation()
	Public ToStationTime()
'	Public RailwayLineCode()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_NearbyStation クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1

		IsData = False
		MaxIndex = -1
		StaffCode = GetForm("CONF_StaffCode", 1)

		Err = ""
		If StaffCode <> "" And IsMainCode(StaffCode) = False Then Err = Err & "StaffCode" & vbCrLf

		Do While True
			If ExistsForm("CONF_StationCode" & idx) = False Then Exit Do

			If GetForm("CONF_StationCode" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve StationCode(MaxIndex) : StationCode(MaxIndex) = GetForm("CONF_StationCode" & idx, 1)
				ReDim Preserve ToStationBusFlag(MaxIndex) : ToStationBusFlag(MaxIndex) = GetForm("CONF_ToStationBusFlag" & idx, 1)
				ReDim Preserve ToStationCarFlag(MaxIndex) : ToStationCarFlag(MaxIndex) = GetForm("CONF_ToStationCarFlag" & idx, 1)
				ReDim Preserve ToStationBicycleFlag(MaxIndex) : ToStationBicycleFlag(MaxIndex) = GetForm("CONF_ToStationBicycleFlag" & idx, 1)
				ReDim Preserve ToStationWalkFlag(MaxIndex) : ToStationWalkFlag(MaxIndex) = GetForm("CONF_ToStationWalkFlag" & idx, 1)
				ReDim Preserve OtherTransportation(MaxIndex) : OtherTransportation(MaxIndex) = GetForm("CONF_OtherTransportation" & idx, 1)
				ReDim Preserve ToStationTime(MaxIndex) : ToStationTime(MaxIndex) = GetForm("CONF_ToStationTime" & idx, 1)

				'値チェック
				If StationCode(MaxIndex) <> "" And IsNumber(StationCode(MaxIndex), 5, False) = False Then Err = Err & "StationCode(" & MaxIndex & ")" & vbCrLf
				If ToStationBusFlag(MaxIndex) <> "" And IsFlag(ToStationBusFlag(MaxIndex)) = False Then Err = Err & "ToStationBusFlag(" & MaxIndex & ")" & vbCrLf
				If ToStationCarFlag(MaxIndex) <> "" And IsFlag(ToStationCarFlag(MaxIndex)) = False Then Err = Err & "ToStationCarFlag(" & MaxIndex & ")" & vbCrLf
				If ToStationBicycleFlag(MaxIndex) <> "" And IsFlag(ToStationBicycleFlag(MaxIndex)) = False Then Err = Err & "ToStationBicycleFlag(" & MaxIndex & ")" & vbCrLf
				If ToStationWalkFlag(MaxIndex) <> "" And IsFlag(ToStationWalkFlag(MaxIndex)) = False Then Err = Err & "ToStationWalkFlag(" & MaxIndex & ")" & vbCrLf
				If ToStationTime(MaxIndex) <> "" And IsNumber(ToStationTime(MaxIndex), 0, True) = False Then Err = Err & "ToStationTime(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_NearbyStation 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_NearbyStation '" & vStaffCode & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_NearbyStation" & _
				" '" & vStaffCode & "'" & _
				",''" & _
				",'" & StationCode(idx) & "'" & _
				",'" & ToStationBusFlag(idx) & "'" & _
				",'" & ToStationCarFlag(idx) & "'" & _
				",'" & ToStationBicycleFlag(idx) & "'" & _
				",'" & ToStationWalkFlag(idx) & "'" & _
				",'" & OtherTransportation(idx) & "'" & _
				",'" & ToStationTime(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_SelectionInfo
'概　要：formで飛んできたP_SelectionInfoテーブル用のデータを持つためのクラス
'備　考：保留
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_SelectionInfo
	Public StaffCode
	Public WorkingTypeCode
	Public OrderCode
	Public SelectionCode
	Public LimitDay
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_SelectionInfo クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1

		StaffCode = GetForm("CONF_StaffCode", 1)
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_SelectionInfo 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		GetRegSQL = "sp_Reg_P_SelectionInfo"
	End Function
End Class

'******************************************************************************
'名　称：clsP_SelfPR
'概　要：formで飛んできたP_SelfPRテーブル用のデータを持つためのクラス
'備　考：navi only
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_SelfPR
	Public StaffCode
	Public SelfPR
	Public WishMotive
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_SelfPR クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If ExistsForm("CONF_SelfPR") = False And ExistsForm("CONF_WishMotive") = False Then Exit Sub

		IsData = True
		If GetForm("CONF_StaffCode", 1) <> "" Then StaffCode = GetForm("CONF_StaffCode", 1)
		If GetForm("CONF_SelfPR", 1) <> "" Then IsData = True: SelfPR = GetForm("CONF_StaffCode", 1)	'ねぇぞ～
		If GetForm("CONF_WishMotive", 1) <> "" Then IsData = True: WishMotive = GetForm("CONF_StaffCode", 1)	'ねぇぞ～
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_SelfPR 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "sp_Reg_P_SelfPR" & _
			" '" & StaffCode & "'" & _
			",'" & SelfPR & "'" & _
			",'" & WishMotive & "'" & vbCrLf
	End Function
End Class

'******************************************************************************
'名　称：clsP_EducateHistory
'概　要：formで飛んできたP_EducateHistoryテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_EducateHistory
	Public StaffCode
	Public EntryDay()
	Public GraduateDay()
	Public EntryTypeCode()
	Public GraduateTypeCode()
	Public SchoolName()
	Public SchoolTypeCode()
	Public Speciality()
	Public CourseType()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_EducateHistory クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sEntryDay
		Dim sGraduateDay
		Dim idx	: idx = 1

		IsData = False
		MaxIndex = -1

		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Err = ""

		Do While True
			If ExistsForm("CONF_EntryDayY" & idx) = False Then Exit Do
			sEntryDay = ""
			sGraduateDay = ""

			sEntryDay = GetForm("CONF_EntryDayY" & idx, 1) & "/"
			If Len(GetForm("CONF_EntryDayM" & idx, 1)) = 1 Then sEntryDay = sEntryDay & "0"
			sEntryDay = sEntryDay & GetForm("CONF_EntryDayM" & idx, 1) & "/01"
			If IsDate(sEntryDay) = False Then sEntryDay = ""

			sGraduateDay = GetForm("CONF_GraduateDayY" & idx, 1) & "/"
			If Len(GetForm("CONF_GraduateDayM" & idx, 1)) = 1 Then sGraduateDay = sGraduateDay & "0"
			sGraduateDay = sGraduateDay & GetForm("CONF_GraduateDayM" & idx, 1) & "/01"
			If IsDate(sGraduateDay) = False Then sGraduateDay = ""

			If IsDate(sEntryDay) = True Then sEntryDay = Replace(sEntryDay, "/", "")
			If IsDate(sGraduateDay) = True Then sGraduateDay = Replace(sGraduateDay, "/", "")

			If sEntryDay <> "" Or sGraduateDay <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve EntryDay(MaxIndex) : EntryDay(MaxIndex) = sEntryDay
				ReDim Preserve GraduateDay(MaxIndex) : GraduateDay(MaxIndex) = sGraduateDay
				ReDim Preserve EntryTypeCode(MaxIndex) : EntryTypeCode(MaxIndex) = GetForm("CONF_EntryTypeCode" & idx, 1)
				ReDim Preserve GraduateTypeCode(MaxIndex) : GraduateTypeCode(MaxIndex) = GetForm("CONF_GraduateTypeCode" & idx, 1)
				ReDim Preserve SchoolName(MaxIndex) : SchoolName(MaxIndex) = GetForm("CONF_SchoolName" & idx, 1)
				ReDim Preserve SchoolTypeCode(MaxIndex) : SchoolTypeCode(MaxIndex) = GetForm("CONF_SchoolTypeCode" & idx, 1)
				ReDim Preserve Speciality(MaxIndex) : Speciality(MaxIndex) = GetForm("CONF_Speciality" & idx, 1)
				ReDim Preserve CourseType(MaxIndex) : CourseType(MaxIndex) = GetForm("CONF_CourseType" & idx, 1)

				'値チェック
				If EntryDay(MaxIndex) <> "" And IsDay(EntryDay(MaxIndex)) = False Then Err = Err & "EntryDay(" & MaxIndex & ")" & vbCrLf
				If GraduateDay(MaxIndex) <> "" And IsDay(GraduateDay(MaxIndex)) = False Then Err = Err & "GraduateDay(" & MaxIndex & ")" & vbCrLf
				If EntryTypeCode(MaxIndex) <> "" And IsNumber(EntryTypeCode(MaxIndex), 3, False) = False Then Err = Err & "EntryTypeCode(" & MaxIndex & ")" & vbCrLf
				If GraduateTypeCode(MaxIndex) <> "" And IsNumber(GraduateTypeCode(MaxIndex), 3, False) = False Then Err = Err & "GraduateTypeCode(" & MaxIndex & ")" & vbCrLf
				If SchoolTypeCode(MaxIndex) <> "" And IsNumber(SchoolTypeCode(MaxIndex), 3, False) = False Then Err = Err & "SchoolTypeCode(" & MaxIndex & ")" & vbCrLf
				If CourseType(MaxIndex) <> "" And IsRE(CourseType(MaxIndex), "^[123]$", True) = False Then Err = Err & "CourseType(" & MaxIndex & ")" & vbCrLf
			End If

			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_EducateHistory 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_EducateHistory '" & vStaffCode & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_EducateHistory" & _
				" '" & vStaffCode & "'" & _
				",''" & _
				",'" & EntryDay(idx) & "'" & _
				",'" & GraduateDay(idx) & "'" & _
				",'" & EntryTypeCode(idx) & "'" & _
				",'" & GraduateTypeCode(idx) & "'" & _
				",'" & SchoolName(idx) & "'" & _
				",'" & SchoolTypeCode(idx) & "'" & _
				",'" & Speciality(idx) & "'" & _
				",'" & CourseType(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_CareerHistory
'概　要：formで飛んできたP_CareerHistoryテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_CareerHistory
	Public StaffCode
	Public IndustryTypeCode()
	Public JobTypeCode()
	Public JobTypeDetail()
	Public WorkingTypeCode()
	Public CompanyName()
	Public CompanyName_F()
	Public EntryDay()
	Public RetireDay()
	Public Period()
	Public BusinessDetail()
	Public RetireReason()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_CareerHistory クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sEntryDay
		Dim sRetireDay
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Err = ""

		Do While True
			If ExistsForm("CONF_EntryDayCY" & idx) = False Then Exit Do
			sEntryDay = ""
			sRetireDay = ""

			sEntryDay = GetForm("CONF_EntryDayCY" & idx, 1) & "/"
			If Len(GetForm("CONF_EntryDayCM" & idx, 1)) = 1 Then sEntryDay = sEntryDay & "0"
			sEntryDay = sEntryDay & GetForm("CONF_EntryDayCM" & idx, 1) & "/01"
			If IsDate(sEntryDay) = False Then sEntryDay = ""

			sRetireDay = GetForm("CONF_RetireDayCY" & idx, 1) & "/"
			If Len(GetForm("CONF_RetireDayCM" & idx, 1)) = 1 Then sRetireDay = sRetireDay & "0"
			sRetireDay = sRetireDay & GetForm("CONF_RetireDayCM" & idx, 1) & "/01"
			If IsDate(sRetireDay) = False Then sRetireDay = ""

			If IsDate(sEntryDay) = True Then sEntryDay = Replace(sEntryDay, "/", "")
			If IsDate(sRetireDay) = True Then sRetireDay = Replace(sRetireDay, "/", "")

			If sEntryDay <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve IndustryTypeCode(MaxIndex) : IndustryTypeCode(MaxIndex) = GetForm("CONF_IndustryTypeCode_C" & idx, 1)
				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = GetForm("CONF_JobTypeCode_C" & idx, 1)
				ReDim Preserve JobTypeDetail(MaxIndex) : JobTypeDetail(MaxIndex) = GetForm("CONF_JobTypeDetail_C" & idx, 1)
				ReDim Preserve WorkingTypeCode(MaxIndex) : WorkingTypeCode(MaxIndex) = GetForm("CONF_WorkingTypeCode_C" & idx, 1)
				ReDim Preserve CompanyName(MaxIndex) : CompanyName(MaxIndex) = GetForm("CONF_CompanyName" & idx, 1)
				ReDim Preserve CompanyName_F(MaxIndex) : CompanyName_F(MaxIndex) = GetForm("CONF_CompanyName_F" & idx, 1)
				ReDim Preserve EntryDay(MaxIndex) : EntryDay(MaxIndex) = sEntryDay
				ReDim Preserve RetireDay(MaxIndex) : RetireDay(MaxIndex) = sRetireDay
				ReDim Preserve Period(MaxIndex) : If IsDate(EntryDay) = True And IsDate(RetireDay) = True Then Period(MaxIndex) = DateDiff("m", EntryDay(MaxIndex), RetireDay(MaxIndex)) / 12	'GetForm("CONF_Period" & idx, 1)
				ReDim Preserve BusinessDetail(MaxIndex) : BusinessDetail(MaxIndex) = GetForm("CONF_BusinessDetail_C" & idx, 1)
				ReDim Preserve RetireReason(MaxIndex) : RetireReason(MaxIndex) = GetForm("CONF_RetireReason" & idx, 1)

				If IndustryTypeCode(MaxIndex) <> "" And IsNumber(IndustryTypeCode(MaxIndex), 3, False) = False Then Err = Err & "IndustryTypeCode(" & MaxIndex & ")" & vbCrLf
				If JobTypeCode(MaxIndex) <> "" And IsNumber(JobTypeCode(MaxIndex), 3, False) = False Then Err = Err & "JobTypeCode(" & MaxIndex & ")" & vbCrLf
				If WorkingTypeCode(MaxIndex) <> "" And IsNumber(WorkingTypeCode(MaxIndex), 3, False) = False Then Err = Err & "WorkingTypeCode(" & MaxIndex & ")" & vbCrLf
				If EntryDay(MaxIndex) <> "" And IsDay(EntryDay(MaxIndex)) = False Then Err = Err & "EntryDay(" & MaxIndex & ")" & vbCrLf
				If RetireDay(MaxIndex) <> "" And IsDay(RetireDay(MaxIndex)) = False Then Err = Err & "RetireDay(" & MaxIndex & ")" & vbCrLf
				If Period(MaxIndex) <> "" And IsRE(Period(MaxIndex), 0, True) = False Then Err = Err & "Period(" & MaxIndex & ")" & vbCrLf

			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_CareerHistory 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_CareerHistory '" & (vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_CareerHistory" & _
				" '" & vStaffCode & "'" & _
				",''" & _
				",'" & IndustryTypeCode(idx) & "'" & _
				",'" & JobTypeCode(idx) & "'" & _
				",'" & JobTypeDetail(idx) & "'" & _
				",'" & WorkingTypeCode(idx) & "'" & _
				",'" & CompanyName(idx) & "'" & _
				",'" & CompanyName_F(idx) & "'" & _
				",'" & EntryDay(idx) & "'" & _
				",'" & RetireDay(idx) & "'" & _
				",'" & Period(idx) & "'" & _
				",'" & BusinessDetail(idx) & "'" & _
				",'" & RetireReason(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_CareerHistoryLis
'概　要：formで飛んできたP_CareerHistoryLisテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_CareerHistoryLis
	Public StaffCode
	Public IndustryTypeCode()
	Public JobTypeCode()
	Public JobTypeDetail()
	Public WorkingTypeCode()
	Public CompanyName()
	Public CompanyName_F()
	Public EntryDay()
	Public RetireDay()
	Public Period()
	Public BusinessDetail()
	Public RetireReason()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_CareerHistoryLis クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sEntryDay
		Dim sRetireDay
		Dim idx	: idx = 1

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Err = ""

		Do While True
			If ExistsForm("CONF_EntryDayCY" & idx) = False Then Exit Do
			sEntryDay = ""
			sRetireDay = ""

			sEntryDay = GetForm("CONF_EntryDayCY" & idx, 1) & "/"
			If Len(GetForm("CONF_EntryDayCM" & idx, 1)) = 1 Then sEntryDay = sEntryDay & "0"
			sEntryDay = sEntryDay & GetForm("CONF_EntryDayCM" & idx, 1) & "/01"
			If IsDate(sEntryDay) = False Then sEntryDay = ""

			sRetireDay = GetForm("CONF_RetireDayCY" & idx, 1) & "/"
			If Len(GetForm("CONF_RetireDayCM" & idx, 1)) = 1 Then sRetireDay = sRetireDay & "0"
			sRetireDay = sRetireDay & GetForm("CONF_RetireDayCM" & idx, 1) & "/01"
			If IsDate(sRetireDay) = False Then sRetireDay = ""

			If IsDate(sEntryDay) = True Then sEntryDay = Replace(sEntryDay, "/", "")
			If IsDate(sRetireDay) = True Then sRetireDay = Replace(sRetireDay, "/", "")

			If sEntryDay <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve IndustryTypeCode(MaxIndex) : IndustryTypeCode(MaxIndex) = GetForm("CONF_IndustryTypeCode_C" & idx, 1)
				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = GetForm("CONF_JobTypeCode_C" & idx, 1)
				ReDim Preserve JobTypeDetail(MaxIndex) : JobTypeDetail(MaxIndex) = GetForm("CONF_JobTypeDetail_C" & idx, 1)
				ReDim Preserve WorkingTypeCode(MaxIndex) : WorkingTypeCode(MaxIndex) = GetForm("CONF_WorkingTypeCode_C" & idx, 1)
				ReDim Preserve CompanyName(MaxIndex) : CompanyName(MaxIndex) = GetForm("CONF_CompanyName" & idx, 1)
				ReDim Preserve CompanyName_F(MaxIndex) : CompanyName_F(MaxIndex) = GetForm("CONF_CompanyName_F" & idx, 1)
				ReDim Preserve EntryDay(MaxIndex) : EntryDay(MaxIndex) = sEntryDay
				ReDim Preserve RetireDay(MaxIndex) : RetireDay(MaxIndex) = sRetireDay
				ReDim Preserve Period(MaxIndex) : Period(MaxIndex) = GetForm("CONF_Period" & idx, 1)
				ReDim Preserve BusinessDetail(MaxIndex) : BusinessDetail(MaxIndex) = GetForm("CONF_BusinessDetail_C" & idx, 1)
				ReDim Preserve RetireReason(MaxIndex) : RetireReason(MaxIndex) = GetForm("CONF_RetireReason" & idx, 1)

				If IndustryTypeCode(MaxIndex) <> "" And IsNumber(IndustryTypeCode(MaxIndex), 3, False) = False Then Err = Err & "IndustryTypeCode(" & MaxIndex & ")" & vbCrLf
				If JobTypeCode(MaxIndex) <> "" And IsNumber(JobTypeCode(MaxIndex), 3, False) = False Then Err = Err & "JobTypeCode(" & MaxIndex & ")" & vbCrLf
				If WorkingTypeCode(MaxIndex) <> "" And IsNumber(WorkingTypeCode(MaxIndex), 3, False) = False Then Err = Err & "WorkingTypeCode(" & MaxIndex & ")" & vbCrLf
				If EntryDay(MaxIndex) <> "" And IsDay(sEntryDay) = False Then Err = Err & "EntryDay(" & MaxIndex & ")" & vbCrLf
				If RetireDay(MaxIndex) <> "" And IsDay(sRetireDay) = False Then Err = Err & "RetireDay(" & MaxIndex & ")" & vbCrLf
				If Period(MaxIndex) <> "" And IsRE(Period(MaxIndex), 0, True) = False Then Err = Err & "Period(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_CareerHistoryLis 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_CareerHistoryLis '" & vStaffCode & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_CareerHistoryLis" & _
				" '" & vStaffCode & "'" & _
				",''" & _
				",'" & IndustryTypeCode(idx) & "'" & _
				",'" & JobTypeCode(idx) & "'" & _
				",'" & JobTypeDetail(idx) & "'" & _
				",'" & WorkingTypeCode(idx) & "'" & _
				",'" & CompanyName(idx) & "'" & _
				",'" & CompanyName_F(idx) & "'" & _
				",'" & EntryDay(idx) & "'" & _
				",'" & RetireDay(idx) & "'" & _
				",'" & Period(idx) & "'" & _
				",'" & BusinessDetail(idx) & "'" & _
				",'" & RetireReason(idx) & "'" & vbCrLf
		Next
	End Function
	
End Class

'******************************************************************************
'名　称：clsP_CareerHistoryIT
'概　要：formで飛んできたP_CareerHistoryITテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_CareerHistoryIT
	Public StaffCode
	Public StartDay()
	Public EndDay()
	Public Number()
	Public PMFlag()
	Public PLFlag()
	Public SEFlag()
	Public PGFlag()
	Public TSFlag()
	Public SystemAnalysisFlag()
	Public DesignFlag()
	Public DevelopmentFlag()
	Public TestFlag()
	Public MaintenanceFlag()
	Public DevelopmentRemark()
	Public DevelopmentDetail()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_CareerHistoryIT クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sStartDay
		Dim sEndDay
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Err = ""

		Do While True
			If ExistsForm("CONF_DevelopmentDetail" & idx) = False Then Exit Do
			sStartDay = ""
			sEndDay = ""

			sStartDay = GetForm("CONF_StartDayITY" & idx, 1) & "/"
			If Len(GetForm("CONF_StartDayITM" & idx, 1)) = 1 Then sStartDay = sStartDay & "0"
			sStartDay = sStartDay & GetForm("CONF_StartDayITM" & idx, 1) & "/01"
			If IsDate(sStartDay) = False Then sStartDay = ""

			sEndDay = GetForm("CONF_EndDayITY" & idx, 1) & "/"
			If Len(GetForm("CONF_EndDayITM" & idx, 1)) = 1 Then sEndDay = sEndDay & "0"
			sEndDay = sEndDay & GetForm("CONF_EndDayITM" & idx, 1) & "/01"
			If IsDate(sEndDay) = False Then sEndDay = ""

			If IsDate(sStartDay) = True Then sStartDay = Replace(sStartDay, "/", "")
			If IsDate(sEndDay) = True Then sEndDay = Replace(sEndDay, "/", "")

			If GetForm("CONF_DevelopmentDetail" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve StartDay(MaxIndex) : StartDay(MaxIndex) = sStartDay
				ReDim Preserve EndDay(MaxIndex) : EndDay(MaxIndex) = sEndDay
				ReDim Preserve Number(MaxIndex) : Number(MaxIndex) = GetForm("CONF_Number_IT" & idx, 1)
				ReDim Preserve PMFlag(MaxIndex) : PMFlag(MaxIndex) = GetForm("CONF_PMFlag" & idx, 1)
				ReDim Preserve PLFlag(MaxIndex) : PLFlag(MaxIndex) = GetForm("CONF_PLFlag" & idx, 1)
				ReDim Preserve SEFlag(MaxIndex) : SEFlag(MaxIndex) = GetForm("CONF_SEFlag" & idx, 1)
				ReDim Preserve PGFlag(MaxIndex) : PGFlag(MaxIndex) = GetForm("CONF_PGFlag" & idx, 1)
				ReDim Preserve TSFlag(MaxIndex) : TSFlag(MaxIndex) = GetForm("CONF_TSFlag" & idx, 1)
				ReDim Preserve SystemAnalysisFlag(MaxIndex) : SystemAnalysisFlag(MaxIndex) = GetForm("CONF_SystemAnalysisFlag" & idx, 1)
				ReDim Preserve DesignFlag(MaxIndex) : DesignFlag(MaxIndex) = GetForm("CONF_DesignFlag" & idx, 1)
				ReDim Preserve DevelopmentFlag(MaxIndex) : DevelopmentFlag(MaxIndex) = GetForm("CONF_DevelopmentFlag" & idx, 1)
				ReDim Preserve TestFlag(MaxIndex) : TestFlag(MaxIndex) = GetForm("CONF_TestFlag" & idx, 1)
				ReDim Preserve MaintenanceFlag(MaxIndex) : MaintenanceFlag(MaxIndex) = GetForm("CONF_MaintenanceFlag" & idx, 1)
				ReDim Preserve DevelopmentRemark(MaxIndex) : DevelopmentRemark(MaxIndex) = GetForm("CONF_DevelopmentRemark" & idx, 1)
				ReDim Preserve DevelopmentDetail(MaxIndex) : DevelopmentDetail(MaxIndex) = GetForm("CONF_DevelopmentDetail" & idx, 1)

				If StartDay(MaxIndex) <> "" And IsDay(StartDay(MaxIndex)) = False Then Err = Err & "StartDay(" & MaxIndex & ")" & vbCrLf
				If EndDay(MaxIndex) <> "" And IsDay(EndDay(MaxIndex)) = False Then Err = Err & "EndDay(" & MaxIndex & ")" & vbCrLf
				If Number(MaxIndex) <> "" And IsNumber(Number(MaxIndex), 0, False) = False Then Err = Err & "Number(" & MaxIndex & ")" & vbCrLf
				If PMFlag(MaxIndex) <> "" And IsFlag(PMFlag(MaxIndex)) = False Then Err = Err & "PMFlag(" & MaxIndex & ")" & vbCrLf
				If PLFlag(MaxIndex) <> "" And IsFlag(PLFlag(MaxIndex)) = False Then Err = Err & "PLFlag(" & MaxIndex & ")" & vbCrLf
				If SEFlag(MaxIndex) <> "" And IsFlag(SEFlag(MaxIndex)) = False Then Err = Err & "SEFlag(" & MaxIndex & ")" & vbCrLf
				If PGFlag(MaxIndex) <> "" And IsFlag(PGFlag(MaxIndex)) = False Then Err = Err & "PGFlag(" & MaxIndex & ")" & vbCrLf
				If TSFlag(MaxIndex) <> "" And IsFlag(TSFlag(MaxIndex)) = False Then Err = Err & "TSFlag(" & MaxIndex & ")" & vbCrLf
				If SystemAnalysisFlag(MaxIndex) <> "" And IsFlag(SystemAnalysisFlag(MaxIndex)) = False Then Err = Err & "SystemAnalysisFlag(" & MaxIndex & ")" & vbCrLf
				If DesignFlag(MaxIndex) <> "" And IsFlag(DesignFlag(MaxIndex)) = False Then Err = Err & "DesignFlag(" & MaxIndex & ")" & vbCrLf
				If DevelopmentFlag(MaxIndex) <> "" And IsFlag(DevelopmentFlag(MaxIndex)) = False Then Err = Err & "DevelopmentFlag(" & MaxIndex & ")" & vbCrLf
				If TestFlag(MaxIndex) <> "" And IsFlag(TestFlag(MaxIndex)) = False Then Err = Err & "TestFlag(" & MaxIndex & ")" & vbCrLf
				If MaintenanceFlag(MaxIndex) <> "" And IsFlag(MaintenanceFlag(MaxIndex)) = False Then Err = Err & "MaintenanceFlag(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_CareerHistoryIT 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_CareerHistoryIT '" & vStaffCode & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_CareerHistoryIT" & _
				" '" & vStaffCode & "'" & _
				",''" & _
				",'" & StartDay(idx) & "'" & _
				",'" & EndDay(idx) & "'" & _
				",'" & Number(idx) & "'" & _
				",'" & PMFlag(idx) & "'" & _
				",'" & PLFlag(idx) & "'" & _
				",'" & SEFlag(idx) & "'" & _
				",'" & PGFlag(idx) & "'" & _
				",'" & TSFlag(idx) & "'" & _
				",'" & SystemAnalysisFlag(idx) & "'" & _
				",'" & DesignFlag(idx) & "'" & _
				",'" & DevelopmentFlag(idx) & "'" & _
				",'" & TestFlag(idx) & "'" & _
				",'" & MaintenanceFlag(idx) & "'" & _
				",'" & DevelopmentRemark(idx) & "'" & _
				",'" & DevelopmentDetail(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_DevelopmentTool
'概　要：formで飛んできたP_DevelopmentToolテーブル用のデータを持つためのクラス
'　　　：CategoryCode毎にこのクラスを作成して使用する。
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_DevelopmentTool
	Public StaffCode
	Public CareerHistoryID()
	Public CategoryCode
	Public Code()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_DevelopmentTool クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize(vCategoryCode)
		Dim idx		: idx = 1
		Dim idx1
		Dim idx2
		Dim flag	: flag = False

		IsData = False
		MaxIndex = -1
		StaffCode = GetForm("CONF_StaffCode", 1)

		Err = ""

		'CONF_ の名前が Lang, App, DB と略されてしまっている事への処置
		Select Case vCategoryCode
			Case "Lang": CategoryCode = "DevelopmentLanguage"
			Case "App": CategoryCode = "Application"
			Case "DB": CategoryCode = "Database"
			Case Else: CategoryCode = vCategoryCode
		End Select

		Do While True
			If ExistsForm("CONF_DevelopmentTool_" & vCategoryCode & idx) = False Then Exit Do

			If GetForm("CONF_DevelopmentTool_" & vCategoryCode & idx, 1) <> "" Then
				MaxIndex = MaxIndex + 1

				ReDim Preserve CareerHistoryID(MaxIndex): CareerHistoryID(MaxIndex) = idx
				ReDim Preserve Code(MaxIndex): Code(MaxIndex) = Split(GetForm("CONF_DevelopmentTool_" & vCategoryCode & idx, 1), ",")

				If CareerHistoryID(MaxIndex) <> "" And IsNumber(CareerHistoryID(MaxIndex), 0, False) = False Then Err = Err & "CareerHistoryID(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop

		For idx1 = 0 To MaxIndex
			For idx2 = LBound(Code(idx1)) To UBound(Code(idx1))
				Code(idx1)(idx2) = Trim(Code(idx1)(idx2))
				If Code(idx1)(idx2) <> "" And IsNumber(Code(idx1)(idx2), 3, False) = True Then
					IsData = True
				Else
					Err = Err & "Code(" & idx1 & ")(" & idx2 & ")" & vbCrLf
				End If
			Next
		Next
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_DevelopmentTool 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx
		Dim idxCode

		GetRegSQL = "EXEC sp_Del_P_DevelopmentTool '" & vStaffCode & "', '" & CategoryCode & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			For idxCode = LBound(Code(idx)) To UBound(Code(idx))
				GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_DevelopmentTool" & _
					" '" & vStaffCode & "'" & _
					",'" & CareerHistoryID(idx) & "'" & _
					",''" & _
					",'" & CategoryCode & "'" & _
					",'" & Code(idx)(idxCode) & "'" & vbCrLf
			Next
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_License
'概　要：formで飛んできたP_Licenseテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_License
	Public StaffCode
	Public GroupCode()
	Public CategoryCode()
	Public Code()
	Public GetDay()
	Public Remark()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_License クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sGetDay
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Err = ""

		Do While True
			If ExistsForm("CONF_LicenseCode" & idx) = False Then Exit Do
			sGetDay = ""

			sGetDay = GetForm("CONF_GetDayY" & idx, 1) & "/"
			If Len(GetForm("CONF_GetDayM" & idx, 1)) = 1 Then sGetDay = sGetDay & "0"
			sGetDay = sGetDay & GetForm("CONF_GetDayM" & idx, 1) & "/01"
			If IsDate(sGetDay) = False Then sGetDay = ""
			sGetDay = Replace(sGetDay, "/", "")

			If GetForm("CONF_LicenseCode" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve GroupCode(MaxIndex) : GroupCode(MaxIndex) = Mid(GetForm("CONF_LicenseCode" & idx, 1), 1, 2)
				ReDim Preserve CategoryCode(MaxIndex) : CategoryCode(MaxIndex) = Mid(GetForm("CONF_LicenseCode" & idx, 1), 3, 3)
				ReDim Preserve Code(MaxIndex) : Code(MaxIndex) = Mid(GetForm("CONF_LicenseCode" & idx, 1), 6, 2)
				ReDim Preserve GetDay(MaxIndex) : GetDay(MaxIndex) = sGetDay
				ReDim Preserve Remark(MaxIndex) : Remark(MaxIndex) = GetForm("CONF_LicenseRemark" & idx, 1)

				If GroupCode(MaxIndex) <> "" And IsNumber(GroupCode(MaxIndex), 2, False) = False Then Err = Err & "GroupCode(" & MaxIndex & ")" & vbCrLf
				If CategoryCode(MaxIndex) <> "" And IsNumber(CategoryCode(MaxIndex), 3, False) = False Then Err = Err & "CategoryCode(" & MaxIndex & ")" & vbCrLf
				If Code(MaxIndex) <> "" And IsNumber(Code(MaxIndex), 2, False) = False Then Err = Err & "Code(" & MaxIndex & ")" & vbCrLf
				If GetDay(MaxIndex) <> "" And IsDay(GetDay(MaxIndex)) = False Then Err = Err & "GetDay(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_License 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_License '" & vStaffCode & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_License" & _
				" '" & vStaffCode & "'" & _
				",''" & _
				",'" & GroupCode(idx) & "'" & _
				",'" & CategoryCode(idx) & "'" & _
				",'" & Code(idx) & "'" & _
				",'" & GetDay(idx) & "'" & _
				",'" & Remark(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_Skill
'概　要：formで飛んできたP_Skillテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_Skill
	Public StaffCode
	Public CategoryCode
	Public Code()
	Public StartDay()
	Public Period()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_Skill クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize(vCategoryCode)
		Dim sStartDay
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Err = ""

		Select Case vCategoryCode
			Case "OS":		CategoryCode = "OS"
			Case "App":		CategoryCode = "Application"
			Case "Lang":	CategoryCode = "DevelopmentLanguage"
			Case "DB":		CategoryCode = "Database"
			Case Else:		CategoryCode = vCategoryCode
		End Select

		Do While True
			If ExistsForm("CONF_" & vCategoryCode & idx) = False Then Exit Do
			sStartDay = GetForm("CONF_StartDay" & vCategoryCode & "Y" & idx, 1) & "/"
			If Len(GetForm("CONF_StartDay" & vCategoryCode & "M" & idx, 1)) = 1 Then sStartDay = sStartDay & "0"
			sStartDay = sStartDay & GetForm("CONF_StartDay" & vCategoryCode & "M" & idx, 1) & "/01"
			If IsDate(sStartDay) = False Then sStartDay = ""
			sStartDay = Replace(sStartDay, "/", "")

			If GetForm("CONF_" & vCategoryCode & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve Code(MaxIndex) : Code(MaxIndex) = GetForm("CONF_" & vCategoryCode & idx, 1)
				ReDim Preserve StartDay(MaxIndex) : StartDay(MaxIndex) = sStartDay
				ReDim Preserve Period(MaxIndex) : Period(MaxIndex) = GetForm("CONF_Period_" & vCategoryCode & idx, 1)
				If Code(MaxIndex) <> "" And IsNumber(Code(MaxIndex), 3, False) = False Then Err = Err & "Code(" & MaxIndex & ")" & vbCrLf
				If StartDay(MaxIndex) <> "" And IsDay(StartDay(MaxIndex)) = False Then Err = Err & "StartDay(" & MaxIndex & ")" & vbCrLf
				If Period(MaxIndex) <> "" And IsNumber(Period(MaxIndex), 0, False) = False Then Err = Err & "Period(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_Skill 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_Skill '" & vStaffCode & "', '" & CategoryCode & "'" & vbCrLf
		'If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_Skill" & _
				" '" & vStaffCode & "'"  & _
				",''"  & _
				",'" & CategoryCode & "'" & _
				",'" & Code(idx) & "'" & _
				",'" & StartDay(idx) & "'" & _
				",'" & Period(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_HopeJobType
'概　要：formで飛んできたP_HopeJobTypeテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_HopeJobType
	Public StaffCode
	Public JobTypeCode()
	Public JobTypeDetail()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_HopeJobType クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Err = ""

		Do While True
			If ExistsForm("CONF_JobTypeCode" & idx) = False Then Exit Do

			If GetForm("CONF_JobTypeCode" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = GetForm("CONF_JobTypeCode" & idx, 1)
				ReDim Preserve JobTypeDetail(MaxIndex) : JobTypeDetail(MaxIndex) = GetForm("CONF_JobTypeDetail" & idx, 1)

				If JobTypeCode(MaxIndex) <> "" And IsNumber(JobTypeCode(MaxIndex), 3, False) = False Then Err = Err & "JobTypeCode(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopeJobType 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeJobType '" & vStaffCode & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeJobType" & _
				" '" & vStaffCode & "'"  & _
				",''" & _
				",'" & JobTypeCode(idx) & "'" & _
				",'" & JobTypeDetail(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_HopeIndustryType
'概　要：formで飛んできたP_HopeIndustryTypeテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_HopeIndustryType
	Public StaffCode
	Public IndustryTypeCode
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_HopIndustryType クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)
		If GetForm("CONF_IndustryTypeCode", 1) <> "" Then IsData = True: IndustryTypeCode = GetForm("CONF_IndustryTypeCode", 1)

		Err = ""

		If IndustryTypeCode <> "" And IsNumber(IndustryTypeCode, 3, False) = False Then Err = Err & "IndustryTypeCode" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopIndustryType 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		GetRegSQL = "EXEC sp_Del_P_HopeIndustryType '" & (vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeIndustryType" & _
			" '" & vStaffCode & "'" & _
			",''" & _
			",'" & IndustryTypeCode & "'" & vbCrLf
	End Function
End Class

'******************************************************************************
'名　称：clsP_HopeWorkingPlace
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_HopeWorkingPlace
	Public StaffCode
	Public PrefectureCode()
	Public City()
	Public Area()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_HopeWorkingPlace クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Do While True
			If ExistsForm("CONF_HopePrefecture" & idx) = False Then Exit Do

			If GetForm("CONF_HopePrefecture" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve PrefectureCode(MaxIndex): PrefectureCode(MaxIndex) = GetForm("CONF_HopePrefecture" & idx, 1)
				ReDim Preserve City(MaxIndex): City(MaxIndex) = GetForm("CONF_HopeCity" & idx, 1)
				ReDim Preserve Area(MaxIndex): Area(MaxIndex) = GetForm("CONF_HopeArea" & idx, 1)

				If PrefectureCode(MaxIndex) <> "" And IsNumber(PrefectureCode(MaxIndex), 3, False) = False Then Err = Err & "PrefectureCode(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopeWorkingPlace 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeWorkingPlace '" & (vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeWorkingPlace" & _
				" '" & vStaffCode & "'"  & _
				",''" & _
				",'" & PrefectureCode(idx) & "'" & _
				",'" & City(idx) & "'" & _
				",'" & Area(idx) & "'" & vbCrLf
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_HopeWorkingType
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_HopeWorkingType
	Public StaffCode
	Public WorkingTypeCode()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_HopeWorkingType クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Do While True
			If ExistsForm("CONF_HopeWorkingType" & idx) = False Then Exit Do

			If GetForm("CONF_HopeWorkingType" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve WorkingTypeCode(MaxIndex) : WorkingTypeCode(MaxIndex) = GetForm("CONF_HopeWorkingType" & idx, 1)

				If WorkingTypeCode(MaxIndex) <> "" And IsNumber(WorkingTypeCode(MaxIndex), 3, False) = False Then Err = Err & "WorkingTypeCode(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopeWorkingType 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeWorkingType '" & (vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeWorkingType" & _
				" '" & vStaffCode & "'" & _
				",''" & _
				",'" & WorkingTypeCode(idx) & "'"
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_HopeCommuting
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：希望駅等
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_HopeCommuting
	Public StaffCode
	Public StationCode()
	Public MinuteToStation()
	Public CommuteTime()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_HopeCommuting クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Do While True
			If ExistsForm("CONF_StationCodeHope" & idx) = False Then Exit Do

			If GetForm("CONF_StationCodeHope" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve StationCode(MaxIndex):		StationCode(MaxIndex) = GetForm("CONF_StationCodeHope" & idx, "1")
				ReDim Preserve MinuteToStation(MaxIndex):	MinuteToStation(MaxIndex) = GetForm("CONF_MinuteToStation" & idx, "1")
				ReDim Preserve CommuteTime(MaxIndex):		CommuteTime(MaxIndex) = GetForm("CONF_HopeCommuteTime" & idx, "1")

				If StationCode(MaxIndex) <> "" And IsNumber(StationCode(MaxIndex), 5, False) = False Then Err = Err & "StationCode(" & MaxIndex & ")" & vbCrLf
				If MinuteToStation(MaxIndex) <> "" And IsNumber(MinuteToStation(MaxIndex), 0, False) = False Then Err = Err & "MinuteToStation(" & MaxIndex & ")" & vbCrLf
				If CommuteTime(MaxIndex) <> "" And IsNumber(CommuteTime(MaxIndex), 0, False) = False Then Err = Err & "CommuteTime(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopeCommuting 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeCommuting '" & (vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeCommuting" & _
				" '" & vStaffCode & "'" & _
				",''" & _
				",'" & StationCode(idx) & "'" & _
				",'" & MinuteToStation(idx) & "'" & _
				",'" & CommuteTime(idx) & "'" & vbCrLf
		Next
	End Function
End Class

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

		If GetForm("CONF_StaffCode", 1) <> "" Then StaffCode = GetForm("CONF_StaffCode", 1)
		If GetForm("CONF_YearlyIncomeMin", 1) <> "" Then IsData = True: YearlyIncomeMin = GetForm("CONF_YearlyIncomeMin", 1) * 10000
		If GetForm("CONF_YearlyIncomeMax", 1) <> "" Then IsData = True: YearlyIncomeMax = GetForm("CONF_YearlyIncomeMax", 1) * 10000
		If GetForm("CONF_MonthlyIncomeMin", 1) <> "" Then IsData = True: MonthlyIncomeMin = GetForm("CONF_MonthlyIncomeMin", 1) * 10000
		If GetForm("CONF_MonthlyIncomeMax", 1) <> "" Then IsData = True: MonthlyIncomeMax = GetForm("CONF_MonthlyIncomeMax", 1) * 10000
		If GetForm("CONF_DailyIncomeMin", 1) <> "" Then IsData = True: DailyIncomeMin = GetForm("CONF_DailyIncomeMin", 1)
		If GetForm("CONF_DailyIncomeMax", 1) <> "" Then IsData = True: DailyIncomeMax = GetForm("CONF_DailyIncomeMax", 1)
		If GetForm("CONF_HourlyIncomeMin", 1) <> "" Then IsData = True: HourlyIncomeMin = GetForm("CONF_HourlyIncomeMin", 1)
		If GetForm("CONF_HourlyIncomeMax", 1) <> "" Then IsData = True: HourlyIncomeMax = GetForm("CONF_HourlyIncomeMax", 1)
		If GetForm("CONF_PercentagePayFlag", 1) <> "" Then IsData = True: PercentagePayFlag = GetForm("CONF_PercentagePayFlag", 1)
		If GetForm("CONF_IncomeRemark", 1) <> "" Then IsData = True: IncomeRemark = GetForm("CONF_IncomeRemark", 1)
		If GetForm("CONF_TrafficFeeFlag", 1) <> "" Then IsData = True: TrafficFeeFlag = GetForm("CONF_TrafficFeeFlag", 1)
		If GetForm("CONF_SocietyInsuranceFlag", 1) <> "" Then IsData = True: SocietyInsuranceFlag = GetForm("CONF_SocietyInsuranceFlag", 1)
		If GetForm("CONF_SanatoriumFlag", 1) <> "" Then IsData = True: SanatoriumFlag = GetForm("CONF_SanatoriumFlag", 1)
		If GetForm("CONF_EnterprisePensionFlag", 1) <> "" Then IsData = True: EnterprisePensionFlag = GetForm("CONF_EnterprisePensionFlag", 1)
		If GetForm("CONF_WealthShapeFlag", 1) <> "" Then IsData = True: WealthShapeFlag = GetForm("CONF_WealthShapeFlag", 1)
		If GetForm("CONF_StockOptionFlag", 1) <> "" Then IsData = True: StockOptionFlag = GetForm("CONF_StockOptionFlag", 1)
		If GetForm("CONF_RetirementPayFlag", 1) <> "" Then IsData = True: RetirementPayFlag = GetForm("CONF_RetirementPayFlag", 1)
		If GetForm("CONF_ResidencePayFlag", 1) <> "" Then IsData = True: ResidencePayFlag = GetForm("CONF_ResidencePayFlag", 1)
		If GetForm("CONF_FamilyPayFlag", 1) <> "" Then IsData = True: FamilyPayFlag = GetForm("CONF_FamilyPayFlag", 1)
		If GetForm("CONF_EmployeeDormitoryFlag", 1) <> "" Then IsData = True: EmployeeDormitoryFlag = GetForm("CONF_EmployeeDormitoryFlag", 1)
		If GetForm("CONF_CompanyHouseFlag", 1) <> "" Then IsData = True: CompanyHouseFlag = GetForm("CONF_CompanyHouseFlag", 1)
		If GetForm("CONF_NewEmployeeTrainingFlag", 1) <> "" Then IsData = True: NewEmployeeTrainingFlag = GetForm("CONF_NewEmployeeTrainingFlag", 1)
		If GetForm("CONF_OverseasTrainingFlag", 1) <> "" Then IsData = True: OverseasTrainingFlag = GetForm("CONF_OverseasTrainingFlag", 1)
		If GetForm("CONF_OtherTrainingFlag", 1) <> "" Then IsData = True: OtherTrainingFlag = GetForm("CONF_OtherTrainingFlag", 1)
		If GetForm("CONF_FlexTimeFlag", 1) <> "" Then IsData = True: FlexTimeFlag = GetForm("CONF_FlexTimeFlag", 1)
		If GetForm("CONF_WorkPeriodTypeFlag", 1) <> "" Then IsData = True: WorkPeriodFlag = GetForm("CONF_WorkPeriodTypeFlag", 1)
		If GetForm("CONF_HopeMonthPeriod", 1) <> "" Then IsData = True: WorkMonthPeriod = GetForm("CONF_HopeMonthPeriod", 1)
		If GetForm("CONF_WorkStartTime", 1) <> "" Then IsData = True: WorkStartTime = GetForm("CONF_WorkStartTime", 1)
		If GetForm("CONF_WorkEndTime", 1) <> "" Then IsData = True: WorkEndTime = GetForm("CONF_WorkEndTime", 1)
		If GetForm("CONF_WorkShiftFlag", 1) <> "" Then IsData = True: WorkShiftFlag = GetForm("CONF_WorkShiftFlag", 1)
		If GetForm("CONF_OverWorkFlag", 1) <> "" Then IsData = True: OverWorkFlag = GetForm("CONF_OverWorkFlag", 1)
		If GetForm("CONF_OverWorkTimeMax", 1) <> "" Then IsData = True: OverWorkTimeMax = GetForm("CONF_OverWorkTimeMax", 1)
		If GetForm("CONF_OverWorkTimeOther", 1) <> "" Then IsData = True: OverWorkTimeOther = GetForm("CONF_OverWorkTimeOther", 1)	'ねぇぞ～
		If GetForm("CONF_MonHolidayFlag", 1) <> "" Then IsData = True: MonHolidayFlag = GetForm("CONF_MonHolidayFlag", 1)
		If GetForm("CONF_TueHolidayFlag", 1) <> "" Then IsData = True: TueHolidayFlag = GetForm("CONF_TueHolidayFlag", 1)
		If GetForm("CONF_WedHolidayFlag", 1) <> "" Then IsData = True: WedHolidayFlag = GetForm("CONF_WedHolidayFlag", 1)
		If GetForm("CONF_ThuHolidayFlag", 1) <> "" Then IsData = True: ThuHolidayFlag = GetForm("CONF_ThuHolidayFlag", 1)
		If GetForm("CONF_FriHolidayFlag", 1) <> "" Then IsData = True: FriHolidayFlag = GetForm("CONF_FriHolidayFlag", 1)
		If GetForm("CONF_SatHolidayFlag", 1) <> "" Then IsData = True: SatHolidayFlag = GetForm("CONF_SatHolidayFlag", 1)
		If GetForm("CONF_SunHolidayFlag", 1) <> "" Then IsData = True: SunHolidayFlag = GetForm("CONF_SunHolidayFlag", 1)
		If GetForm("CONF_PublicHolidayFlag", 1) <> "" Then IsData = True: PublicHolidayFlag = GetForm("CONF_PublicHolidayFlag", 1)
		If GetForm("CONF_WeeklyHolidayType", 1) <> "" Then IsData = True: WeeklyHolidayType = GetForm("CONF_WeeklyHolidayType", 1)
		If GetForm("CONF_HolidayRemark", 1) <> "" Then IsData = True: HolidayRemark = GetForm("CONF_HolidayRemark", 1)
		If GetForm("CONF_TransferFlag", 1) <> "" Then IsData = True: TransferFlag = GetForm("CONF_TransferFlag", 1)
		If GetForm("CONF_HopeWorkStartDay", 1) <> "" Then IsData = True: HopeWorkStartDay = GetForm("CONF_HopeWorkStartDay", 1)

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
			" '" & vStaffCode & "'" & _
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
			",'" & TrafficFeeFlag & "'" & _
			",'" & SocietyInsuranceFlag & "'" & _
			",'" & SanatoriumFlag & "'" & _
			",'" & EnterprisePensionFlag & "'" & _
			",'" & WealthShapeFlag & "'" & _
			",'" & StockOptionFlag & "'" & _
			",'" & RetirementPayFlag & "'" & _
			",'" & ResidencePayFlag & "'" & _
			",'" & FamilyPayFlag & "'" & _
			",'" & EmployeeDormitoryFlag & "'" & _
			",'" & CompanyHouseFlag & "'" & _
			",'" & NewEmployeeTrainingFlag & "'" & _
			",'" & OverseasTrainingFlag & "'" & _
			",'" & OtherTrainingFlag & "'" & _
			",'" & FlexTimeFlag & "'" & _
			",'" & WorkPeriodFlag & "'" & _
			",'" & WorkMonthPeriod & "'" & _
			",'" & WorkStartTime & "'" & _
			",'" & WorkEndTime & "'" & _
			",'" & WorkShiftFlag & "'" & _
			",'" & OverWorkFlag & "'" & _
			",'" & OverWorkTimeMax & "'" & _
			",'" & OverWorkTimeOther & "'" & _
			",'" & MonHolidayFlag & "'" & _
			",'" & TueHolidayFlag & "'" & _
			",'" & WedHolidayFlag & "'" & _
			",'" & ThuHolidayFlag & "'" & _
			",'" & FriHolidayFlag & "'" & _
			",'" & SatHolidayFlag & "'" & _
			",'" & SunHolidayFlag & "'" & _
			",'" & PublicHolidayFlag & "'" & _
			",'" & WeeklyHolidayType & "'" & _
			",'" & HolidayRemark & "'" & _
			",'" & TransferFlag & "'" & _
			",'" & HopeWorkStartDay & "'" & vbCrLf
	End Function
End Class

'******************************************************************************
'名　称：clsP_SkillTest
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_SkillTest
	Public StaffCode
	Public ExecuteDay1
	Public Kana_M1
	Public RomanChar_M1
	Public TenKeyTime1
	Public TenKeyCorrect1
	Public TenKeyStroke1
	Public ExecuteDay2
	Public Kana_M2
	Public RomanChar_M2
	Public TenKeyTime2
	Public TenKeyCorrect2
	Public TenKeyStroke2
	Public ExecuteDay3
	Public Kana_M3
	Public RomanChar_M3
	Public TenKeyTime3
	Public TenKeyCorrect3
	Public TenKeyStroke3
	Public ExecuteDay4
	Public Kana_M4
	Public RomanChar_M4
	Public TenKeyTime4
	Public TenKeyCorrect4
	Public TenKeyStroke4
	Public Behavior
	Public Durability
	Public Leader
	Public Challenge
	Public Sympathy
	Public Stability
	Public Originality
	Public Innovation
	Public Thinking
	Public Flexibility
	Public Sensitivity
	Public Carefulness
	Public DutySynthesis
	Public DutyRank
	Public GeneralSynthesis
	Public GeneralRank
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_SkillTest クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1
		ExecuteDay1 = ""
		ExecuteDay2 = ""
		ExecuteDay3 = ""
		ExecuteDay4 = ""

		ExecuteDay1 = GetForm("CONF_ExecuteDayY1", 1) & "/"
		If Len(GetForm("CONF_ExecuteDayM1", 1)) = 1 Then ExecuteDay1 = ExecuteDay1 & "0"
		ExecuteDay1 = ExecuteDay1 & GetForm("CONF_ExecuteDayM1", 1) & "/01"
		If IsDate(ExecuteDay1) = True Then
			ExecuteDay1 = Replace(ExecuteDay1, "/", "")
		Else
			ExecuteDay1 = ""
		End If

		ExecuteDay2 = GetForm("CONF_ExecuteDayY2", 1) & "/"
		If Len(GetForm("CONF_ExecuteDayM2", 1)) = 1 Then ExecuteDay2 = ExecuteDay2 & "0"
		ExecuteDay2 = ExecuteDay2 & GetForm("CONF_ExecuteDayM2", 1) & "/01"
		If IsDate(ExecuteDay2) = True Then
			ExecuteDay2 = Replace(ExecuteDay2, "/", "")
		Else
			ExecuteDay2 = ""
		End If

		ExecuteDay3 = GetForm("CONF_ExecuteDayY3", 1) & "/"
		If Len(GetForm("CONF_ExecuteDayM3", 1)) = 1 Then ExecuteDay3 = ExecuteDay3 & "0"
		ExecuteDay3 = ExecuteDay3 & GetForm("CONF_ExecuteDayM3", 1) & "/01"
		If IsDate(ExecuteDay3) = True Then
			ExecuteDay3 = Replace(ExecuteDay3, "/", "")
		Else
			ExecuteDay3 = ""
		End If

		ExecuteDay4 = GetForm("CONF_ExecuteDayY4", 1) & "/"
		If Len(GetForm("CONF_ExecuteDayM4", 1)) = 1 Then ExecuteDay4 = ExecuteDay4 & "0"
		ExecuteDay4 = ExecuteDay4 & GetForm("CONF_ExecuteDayM4", 1) & "/01"
		If IsDate(ExecuteDay4) = True Then
			ExecuteDay4 = Replace(ExecuteDay4, "/", "")
		Else
			ExecuteDay4 = ""
		End If

		If GetForm("CONF_StaffCode", 1) <> "" Then StaffCode = GetForm("CONF_StaffCode", 1)
		If GetForm("CONF_Kana_M1", 1) <> "" Then IsData = True: Kana_M1 = GetForm("CONF_Kana_M1", 1)
		If GetForm("CONF_RomanChar_M1", 1) <> "" Then IsData = True: RomanChar_M1 = GetForm("CONF_RomanChar_M1", 1)
		If GetForm("CONF_TenKeyTime1", 1) <> "" Then IsData = True: TenKeyTime1 = GetForm("CONF_TenKeyTime1", 1)
		If GetForm("CONF_TenKeyCorrect1", 1) <> "" Then IsData = True: TenKeyCorrect1 = GetForm("CONF_TenKeyCorrect1", 1)
		If GetForm("CONF_TenKeyStroke1", 1) <> "" Then IsData = True: TenKeyStroke1 = GetForm("CONF_TenKeyStroke1", 1)
		If GetForm("CONF_Kana_M2", 1) <> "" Then IsData = True: Kana_M2 = GetForm("CONF_Kana_M2", 1)
		If GetForm("CONF_RomanChar_M2", 1) <> "" Then IsData = True: RomanChar_M2 = GetForm("CONF_RomanChar_M2", 1)
		If GetForm("CONF_TenKeyTime2", 1) <> "" Then IsData = True: TenKeyTime2 = GetForm("CONF_TenKeyTime2", 1)
		If GetForm("CONF_TenKeyCorrect2", 1) <> "" Then IsData = True: TenKeyCorrect2 = GetForm("CONF_TenKeyCorrect2", 1)
		If GetForm("CONF_TenKeyStroke2", 1) <> "" Then IsData = True: TenKeyStroke2 = GetForm("CONF_TenKeyStroke2", 1)
		If GetForm("CONF_Kana_M3", 1) <> "" Then IsData = True: Kana_M3 = GetForm("CONF_Kana_M3", 1)
		If GetForm("CONF_RomanChar_M3", 1) <> "" Then IsData = True: RomanChar_M3 = GetForm("CONF_RomanChar_M3", 1)
		If GetForm("CONF_TenKeyTime3", 1) <> "" Then IsData = True: TenKeyTime3 = GetForm("CONF_TenKeyTime3", 1)
		If GetForm("CONF_TenKeyCorrect3", 1) <> "" Then IsData = True: TenKeyCorrect3 = GetForm("CONF_TenKeyCorrect3", 1)
		If GetForm("CONF_TenKeyStroke3", 1) <> "" Then IsData = True: TenKeyStroke3 = GetForm("CONF_TenKeyStroke3", 1)
		If GetForm("CONF_Kana_M4", 1) <> "" Then IsData = True: Kana_M4 = GetForm("CONF_Kana_M4", 1)
		If GetForm("CONF_RomanChar_M4", 1) <> "" Then IsData = True: RomanChar_M4 = GetForm("CONF_RomanChar_M4", 1)
		If GetForm("CONF_TenKeyTime4", 1) <> "" Then IsData = True: TenKeyTime4 = GetForm("CONF_TenKeyTime4", 1)
		If GetForm("CONF_TenKeyCorrect4", 1) <> "" Then IsData = True: TenKeyCorrect4 = GetForm("CONF_TenKeyCorrect4", 1)
		If GetForm("CONF_TenKeyStroke4", 1) <> "" Then IsData = True: TenKeyStroke4 = GetForm("CONF_TenKeyStroke4", 1)
		If GetForm("CONF_Behavior", 1) <> "" Then IsData = True: Behavior = GetForm("CONF_Behavior", 1)
		If GetForm("CONF_Durability", 1) <> "" Then IsData = True: Durability = GetForm("CONF_Durability", 1)
		If GetForm("CONF_Leader", 1) <> "" Then IsData = True: Leader = GetForm("CONF_Leader", 1)
		If GetForm("CONF_Challenge", 1) <> "" Then IsData = True: Challenge = GetForm("CONF_Challenge", 1)
		If GetForm("CONF_Sympathy", 1) <> "" Then IsData = True: Sympathy = GetForm("CONF_Sympathy", 1)
		If GetForm("CONF_Stability", 1) <> "" Then IsData = True: Stability = GetForm("CONF_Stability", 1)
		If GetForm("CONF_Originality", 1) <> "" Then IsData = True: Originality = GetForm("CONF_Originality", 1)
		If GetForm("CONF_Innovation", 1) <> "" Then IsData = True: Innovation = GetForm("CONF_Innovation", 1)
		If GetForm("CONF_Thinking", 1) <> "" Then IsData = True: Thinking = GetForm("CONF_Thinking", 1)
		If GetForm("CONF_Flexibility", 1) <> "" Then IsData = True: Flexibility = GetForm("CONF_Flexibility", 1)
		If GetForm("CONF_Sensitivity", 1) <> "" Then IsData = True: Sensitivity = GetForm("CONF_Sensitivity", 1)
		If GetForm("CONF_Carefulness", 1) <> "" Then IsData = True: Carefulness = GetForm("CONF_Carefulness", 1)
		If GetForm("CONF_DutySynthesis", 1) <> "" Then IsData = True: DutySynthesis = GetForm("CONF_DutySynthesis", 1)
		If GetForm("CONF_DutyRank", 1) <> "" Then IsData = True: DutyRank = GetForm("CONF_DutyRank", 1)
		If GetForm("CONF_GeneralSynthesis", 1) <> "" Then IsData = True: GeneralSynthesis = GetForm("CONF_GeneralSynthesis", 1)
		If GetForm("CONF_GeneralRank", 1) <> "" Then IsData = True: GeneralRank = GetForm("CONF_GeneralRank", 1)

		Err = ""
		If ExecuteDay1 <> "" And IsDay(ExecuteDay1) = False Then Err = Err & "ExecuteDay1" & vbCrLf
		If ExecuteDay2 <> "" And IsDay(ExecuteDay2) = False Then Err = Err & "ExecuteDay2" & vbCrLf
		If ExecuteDay3 <> "" And IsDay(ExecuteDay3) = False Then Err = Err & "ExecuteDay3" & vbCrLf
		If ExecuteDay4 <> "" And IsDay(ExecuteDay4) = False Then Err = Err & "ExecuteDay4" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_SkillTest 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_SkillTest" & _
			" '" & vStaffCode & "'" & _
			",'" & ExecuteDay1 & "'" & _
			",'" & Kana_M1 & "'" & _
			",'" & RomanChar_M1 & "'" & _
			",'" & TenKeyTime1 & "'" & _
			",'" & TenKeyCorrect1 & "'" & _
			",'" & TenKeyStroke1 & "'" & _
			",'" & ExecuteDay2 & "'" & _
			",'" & Kana_M2 & "'" & _
			",'" & RomanChar_M2 & "'" & _
			",'" & TenKeyTime2 & "'" & _
			",'" & TenKeyCorrect2 & "'" & _
			",'" & TenKeyStroke2 & "'" & _
			",'" & ExecuteDay3 & "'" & _
			",'" & Kana_M3 & "'" & _
			",'" & RomanChar_M3 & "'" & _
			",'" & TenKeyTime3 & "'" & _
			",'" & TenKeyCorrect3 & "'" & _
			",'" & TenKeyStroke3 & "'" & _
			",'" & ExecuteDay4 & "'" & _
			",'" & Kana_M4 & "'" & _
			",'" & RomanChar_M4 & "'" & _
			",'" & TenKeyTime4 & "'" & _
			",'" & TenKeyCorrect4 & "'" & _
			",'" & TenKeyStroke4 & "'" & _
			",'" & Behavior & "'" & _
			",'" & Durability & "'" & _
			",'" & Leader & "'" & _
			",'" & Challenge & "'" & _
			",'" & Sympathy & "'" & _
			",'" & Stability & "'" & _
			",'" & Originality & "'" & _
			",'" & Innovation & "'" & _
			",'" & Thinking & "'" & _
			",'" & Flexibility & "'" & _
			",'" & Sensitivity & "'" & _
			",'" & Carefulness & "'" & _
			",'" & DutySynthesis & "'" & _
			",'" & DutyRank & "'" & _
			",'" & GeneralSynthesis & "'" & _
			",'" & GeneralRank & "'" & vbCrLf
	End Function
End Class

'******************************************************************************
'名　称：clsP_Talent
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_Talent
	Public StaffCode
	Public CompanyCode
	Public LicenseNumber
	Public EmploymentDivisionFlag
	Public RecommendationLetter
	Public WorkDivisionFlag
	Public PL_StateFlag
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_Talent クラスの初期化関数
	'備　考：社内システムでは用なし
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If GetForm("CONF_StaffCode", 1) <> "" Then StaffCode = GetForm("CONF_StaffCode", 1)
		If GetForm("CONF_CompanyCode", 1) <> "" Then IsData = True: CompanyCode = GetForm("CONF_CompanyCode", 1)	'ねぇぞ～
		If GetForm("CONF_LicenseNumber", 1) <> "" Then IsData = True: LicenseNumber = GetForm("CONF_LicenseNumber", 1)	'ねぇぞ～
		If GetForm("CONF_EmploymentDivisionFlag", 1) <> "" Then IsData = True: EmploymentDivisionFlag = GetForm("CONF_EmploymentDivisionFlag", 1)	'ねぇぞ～
		If GetForm("CONF_RecommendationLetter", 1) <> "" Then IsData = True: RecommendationLetter = GetForm("CONF_RecommendationLetter", 1)	'ねぇぞ～
		If GetForm("CONF_WorkDivisionFlag", 1) <> "" Then IsData = True: WorkDivisionFlag = GetForm("CONF_WorkDivisionFlag", 1)	'ねぇぞ～
		If GetForm("CONF_PL_StateFlag", 1) <> "" Then IsData = True: PL_StateFlag = GetForm("CONF_PL_StateFlag", 1)	'ねぇぞ～

		Err = ""
		If CompanyCode <> "" And IsMainCode(CompanyCode) = False Then Err = Err & "CompanyCode" & vbCrLf
		If LicenseNumber <> "" And IsNumber(LicenseNumber, 0, False) = False Then Err = Err & "LicenseNumber" & vbCrLf
		If EmploymentDivisionFlag <> "" And IsFlag(EmploymentDivisionFlag) = False Then Err = Err & "EmploymentDivisionFlag" & vbCrLf
		If WorkDivisionFlag <> "" And IsFlag(WorkDivisionFlag) = False Then Err = Err & "WorkDivisionFlag" & vbCrLf
		If PL_StateFlag <> "" And IsFlag(PL_StateFlag) = False Then Err = Err & "PL_StateFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_Talent 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_Talent" & _
			" '" & vStaffCode & "'" & _
			",'" & CompanyCode & "'" & _
			",'" & LicenseNumber & "'" & _
			",'" & EmploymentDivisionFlag & "'" & _
			",'" & RecommendationLetter & "'" & _
			",'" & WorkDivisionFlag & "'" & _
			",'" & PL_StateFlag & "'" & vbCrLf
	End Function
End Class

'******************************************************************************
'名　称：clsP_Note
'概　要：formで飛んできたP_Noteテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_Note
	Public StaffCode
	Public CategoryCode
	Public Code
	Public Note
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_Note クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize(vCode)
		IsData = False
		MaxIndex = -1

		StaffCode = GetForm("CONF_StaffCode", 1)
		CategoryCode = "Note"
		Code = vCode
		Note = GetForm("CONF_Note_" & vCode, 1)
		If Note <> "" Then IsData = True
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_Note 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		GetRegSQL = "sp_Del_P_Note '" & vStaffCode & "', '" & CategoryCode & "', '" & Code & "'" & vbCrLf
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_Note" & _
			" '" & vStaffCode & "'" & _
			",'" & CategoryCode & "'" & _
			",'" & Code & "'" & _
			",'" & Note & "'"
	End Function
End Class

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

		If GetForm("CONF_StaffCode", 1) <> "" Then StaffCode = GetForm("CONF_StaffCode", 1)
		If GetForm("CONF_BranchCode_Introduce", 1) <> "" Then IsData = True: BranchCode = GetForm("CONF_BranchCode_Introduce", 1)
		If GetForm("CONF_EmployeeCode_Introduce", 1) <> "" Then IsData = True: EmployeeCode = GetForm("CONF_EmployeeCode_Introduce", 1)
		If GetForm("CONF_BasePay", 1) <> "" Then IsData = True: BasePay = GetForm("CONF_BasePay", 1) * 10000
		If GetForm("CONF_OvertimeWorkPayAvg", 1) <> "" Then IsData = True: OvertimeWorkPayAvg = GetForm("CONF_OvertimeWorkPayAvg", 1) * 10000
		If GetForm("CONF_OtherPay", 1) <> "" Then IsData = True: OtherPay = GetForm("CONF_OtherPay", 1) * 10000
		If GetForm("CONF_Bonus", 1) <> "" Then IsData = True: Bonus = GetForm("CONF_Bonus", 1) * 10000
		If GetForm("CONF_AnnualIncome", 1) <> "" Then IsData = True: AnnualIncome = GetForm("CONF_AnnualIncome", 1) * 10000
		If GetForm("CONF_SituationHourlyPay", 1) <> "" Then IsData = True: SituationHourlyPay = GetForm("CONF_SituationHourlyPay", 1)
		If GetForm("CONF_CommutationAllowance", 1) <> "" Then IsData = True: CommutationAllowance = GetForm("CONF_CommutationAllowance", 1)
		If GetForm("CONF_HopeIncomeCode", 1) <> "" Then IsData = True: HopeIncomeCode = GetForm("CONF_HopeIncomeCode", 1)
		If GetForm("CONF_HopeIncomeMin", 1) <> "" Then IsData = True: HopeIncomeMin = GetForm("CONF_HopeIncomeMin", 1) * 10000
		If GetForm("CONF_AnnualSalarySystemFlag", 1) <> "" Then IsData = True: AnnualSalarySystemFlag = GetForm("CONF_AnnualSalarySystemFlag", 1)
		If GetForm("CONF_RaiseTypeCode", 1) <> "" Then IsData = True: RaiseTypeCode = GetForm("CONF_RaiseTypeCode", 1)
		If GetForm("CONF_BonusFlag", 1) <> "" Then IsData = True: BonusFlag = GetForm("CONF_BonusFlag", 1)
		If GetForm("CONF_BonusCount", 1) <> "" Then IsData = True: BonusCount = GetForm("CONF_BonusCount", 1)
		If GetForm("CONF_BonusMin", 1) <> "" Then IsData = True: BonusMin = GetForm("CONF_BonusMin", 1) * 10000
		If GetForm("CONF_SocietyInsuranceFlag2", 1) <> "" Then IsData = True: SocietyInsuranceFlag = GetForm("CONF_SocietyInsuranceFlag2", 1)
		If GetForm("CONF_WelfareAnnuityFlag", 1) <> "" Then IsData = True: WelfareAnnuityFlag = GetForm("CONF_WelfareAnnuityFlag", 1)
		If GetForm("CONF_EmploymentInsuranceFlag", 1) <> "" Then IsData = True: EmploymentInsuranceFlag = GetForm("CONF_EmploymentInsuranceFlag", 1)
		If GetForm("CONF_AccidentInsuranceFlag", 1) <> "" Then IsData = True: AccidentInsuranceFlag = GetForm("CONF_AccidentInsuranceFlag", 1)
		If GetForm("CONF_SelectJobPoint1", 1) <> "" Then IsData = True: SelectJobPoint1 = GetForm("CONF_SelectJobPoint1", 1)
		If GetForm("CONF_SelectJobPoint2", 1) <> "" Then IsData = True: SelectJobPoint2 = GetForm("CONF_SelectJobPoint2", 1)
		If GetForm("CONF_SelectJobPoint3", 1) <> "" Then IsData = True: SelectJobPoint3 = GetForm("CONF_SelectJobPoint3", 1)
		If GetForm("CONF_ForeignCapitalFlag", 1) <> "" Then IsData = True: ForeignCapitalFlag = GetForm("CONF_ForeignCapitalFlag", 1)
		If GetForm("CONF_CapitalMin", 1) <> "" Then IsData = True: CapitalMin = GetForm("CONF_CapitalMin", 1)
		If GetForm("CONF_EmployeeNumMin", 1) <> "" Then IsData = True: EmployeeNumMin = GetForm("CONF_EmployeeNumMin", 1)
		If GetForm("CONF_FounderYear", 1) <> "" Then IsData = True: FounderYear = GetForm("CONF_FounderYear", 1)
		If GetForm("CONF_StartYear", 1) <> "" Then IsData = True: StartYear = GetForm("CONF_StartYear", 1)
		If GetForm("CONF_StartMonth", 1) <> "" Then IsData = True: StartMonth = GetForm("CONF_StartMonth", 1)
		If GetForm("CONF_LimitYear", 1) <> "" Then IsData = True: LimitYear = GetForm("CONF_LimitYear", 1)
		If GetForm("CONF_LimitMonth", 1) <> "" Then IsData = True: LimitMonth = GetForm("CONF_LimitMonth", 1)
		If GetForm("CONF_ActiveFlag", 1) <> "" Then IsData = True: ActiveFlag = GetForm("CONF_ActiveFlag", 1)
		If GetForm("CONF_CompetitionFlag", 1) <> "" Then IsData = True: CompetitionFlag = GetForm("CONF_CompetitionFlag", 1)
		If GetForm("CONF_MediaCode1", 1) <> "" Then IsData = True: MediaCode1 = GetForm("CONF_MediaCode1", 1)
		If GetForm("CONF_MediaCode2", 1) <> "" Then IsData = True: MediaCode2 = GetForm("CONF_MediaCode2", 1)
		If GetForm("CONF_MediaCode3", 1) <> "" Then IsData = True: MediaCode3 = GetForm("CONF_MediaCode3", 1)
		If GetForm("CONF_MediaOther", 1) <> "" Then IsData = True: MediaOther = GetForm("CONF_MediaOther", 1)
		If GetForm("CONF_Rank", 1) <> "" Then IsData = True: Rank = GetForm("CONF_Rank", 1)
		If GetForm("CONF_HopeWeekdayMonFlag", 1) <> "" Then IsData = True: HopeWeekdayMonFlag = GetForm("CONF_HopeWeekdayMonFlag", 1)
		If GetForm("CONF_HopeWeekdayTueFlag", 1) <> "" Then IsData = True: HopeWeekdayTueFlag = GetForm("CONF_HopeWeekdayTueFlag", 1)
		If GetForm("CONF_HopeWeekdayWedFlag", 1) <> "" Then IsData = True: HopeWeekdayWedFlag = GetForm("CONF_HopeWeekdayWedFlag", 1)
		If GetForm("CONF_HopeWeekdayThuFlag", 1) <> "" Then IsData = True: HopeWeekdayThuFlag = GetForm("CONF_HopeWeekdayThuFlag", 1)
		If GetForm("CONF_HopeWeekdayFriFlag", 1) <> "" Then IsData = True: HopeWeekdayFriFlag = GetForm("CONF_HopeWeekdayFriFlag", 1)
		If GetForm("CONF_HopeWeekdaySatFlag", 1) <> "" Then IsData = True: HopeWeekdaySatFlag = GetForm("CONF_HopeWeekdaySatFlag", 1)
		If GetForm("CONF_HopeWeekdaySunFlag", 1) <> "" Then IsData = True: HopeWeekdaySunFlag = GetForm("CONF_HopeWeekdaySunFlag", 1)
		If GetForm("CONF_HopeWeekdayOther", 1) <> "" Then IsData = True: HopeWeekdayOther = GetForm("CONF_HopeWeekdayOther", 1)
		If GetForm("CONF_HopeHourFlag", 1) <> "" Then IsData = True: HopeHourFlag = GetForm("CONF_HopeHourFlag", 1)
		If GetForm("CONF_HopeTimeFrom", 1) <> "" Then IsData = True: HopeTimeFrom = GetForm("CONF_HopeTimeFrom", 1)
		If GetForm("CONF_HopeTimeTo", 1) <> "" Then IsData = True: HopeTimeTo = GetForm("CONF_HopeTimeTo", 1)

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
			" '" & vStaffCode & "'" & _
			",'" & BranchCode & "'" & _
			",'" & EmployeeCode & "'" & _
			",'" & BasePay & "'" & _
			",'" & OvertimeWorkPayAvg & "'" & _
			",'" & OtherPay & "'" & _
			",'" & Bonus & "'" & _
			",'" & AnnualIncome & "'" & _
			",'" & SituationHourlyPay & "'" & _
			",'" & CommutationAllowance & "'" & _
			",'" & HopeIncomeCode & "'" & _
			",'" & HopeIncomeMin & "'" & _
			",'" & AnnualSalarySystemFlag & "'" & _
			",'" & RaiseTypeCode & "'" & _
			",'" & BonusFlag & "'" & _
			",'" & BonusCount & "'" & _
			",'" & BonusMin & "'" & _
			",'" & SocietyInsuranceFlag & "'" & _
			",'" & WelfareAnnuityFlag & "'" & _
			",'" & EmploymentInsuranceFlag & "'" & _
			",'" & AccidentInsuranceFlag & "'" & _
			",'" & SelectJobPoint1 & "'" & _
			",'" & SelectJobPoint2 & "'" & _
			",'" & SelectJobPoint3 & "'" & _
			",'" & ForeignCapitalFlag & "'" & _
			",'" & CapitalMin & "'" & _
			",'" & EmployeeNumMin & "'" & _
			",'" & FounderYear & "'" & _
			",'" & StartYear & "'" & _
			",'" & StartMonth & "'" & _
			",'" & LimitYear & "'" & _
			",'" & LimitMonth & "'" & _
			",'" & ActiveFlag & "'" & _
			",'" & CompetitionFlag & "'" & _
			",'" & MediaCode1 & "'" & _
			",'" & MediaCode2 & "'" & _
			",'" & MediaCode3 & "'" & _
			",'" & MediaOther & "'" & _
			",'" & Rank & "'" & _
			",'" & HopeWeekdayMonFlag & "'" & _
			",'" & HopeWeekdayTueFlag & "'" & _
			",'" & HopeWeekdayWedFlag & "'" & _
			",'" & HopeWeekdayThuFlag & "'" & _
			",'" & HopeWeekdayFriFlag & "'" & _
			",'" & HopeWeekdaySatFlag & "'" & _
			",'" & HopeWeekdaySunFlag & "'" & _
			",'" & HopeWeekdayOther & "'" & _
			",'" & HopeHourFlag & "'" & _
			",'" & HopeTimeFrom & "'" & _
			",'" & HopeTimeTo & "'"
	End Function
End Class

'******************************************************************************
'名　称：clsP_IntroductionComment
'概　要：formで飛んできたP_IntroductionCommentテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_IntroductionComment
	Public StaffCode
	Public Comment(16)
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_IntroductionComment クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sidx
		Dim idx

		IsData = False
		MaxIndex = UBound(Comment)
		StaffCode = GetForm("CONF_StaffCode", 1)

		For idx = 1 To UBound(Comment)
			If idx <= 9 Then
				sidx = "00" & idx
			Else
				sidx = "0" & idx
			End If

			Comment(idx) = GetForm("CONF_IntroductionComment" & sidx, "1")	'ねぇぞ～
			If Comment(idx) <> "" Then IsData = True
		Next
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_IntroductionComment 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx
		Dim sidx

		If IsData = False Then Exit Function

		GetRegSQL = ""
		For idx = 1 To MaxIndex
			If Comment(idx) <> "" Then
				If idx <= 9 Then
					sidx = "00" & idx
				Else
					sidx = "0" & idx
				End If

				GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_IntroductionComment" & _
					" '" & vStaffCode & "', 'IntroductionComment', '" & sidx & "', '" & Comment(idx) & "'" & vbCrLf
			End If
		Next
	End Function
End Class

'******************************************************************************
'名　称：clsP_BankAccount
'概　要：formで飛んできたP_BankAccountテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_BankAccount
	Public StaffCode
	Public BankName
	Public BankNo
	Public BankBranchName
	Public BankBranchNo
	Public AccountNo
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_BankAccount クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If GetForm("CONF_StaffCode", 1) <> "" Then StaffCode = GetForm("CONF_StaffCode", 1)
		If GetForm("CONF_BankName", 1) <> "" Then IsData = True: BankName = GetForm("CONF_BankName", 1)
		If GetForm("CONF_BankNo", 1) <> "" Then IsData = True: BankNo = GetForm("CONF_BankNo", 1)
		If GetForm("CONF_BankBranchName", 1) <> "" Then IsData = True: BankBranchName = GetForm("CONF_BankBranchName", 1)
		If GetForm("CONF_BankBranchNo", 1) <> "" Then IsData = True: BankBranchNo = GetForm("CONF_BankBranchNo", 1)
		If GetForm("CONF_AccountNo", 1) <> "" Then IsData = True: AccountNo = GetForm("CONF_AccountNo", 1)
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_BankAccount 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_BankAccount" & _
			" '" & vStaffCode & "'" & _
			",'" & BankName & "'" & _
			",'" & BankNo & "'" & _
			",'" & BankBranchName & "'" & _
			",'" & BankBranchNo & "'" & _
			",'" & AccountNo & "'" & vbCrLf
	End Function
End Class

'******************************************************************************
'名　称：clsP_CompetitionSelection
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_CompetitionSelection
	Public StaffCode
	Public IndustryTypeCode()
	Public JobTypeCode()
	Public CompanyName()
	Public SelectionTypeCode()
	Public MediaCode()
	Public OtherMedia()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_CompetitionSelection クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If GetForm("StaffCode", 1) <> "" Then StaffCode = GetForm("StaffCode", 1)

		Do While True
			If ExistsForm("CONF_SelectionTypeCode" & idx) = False Then Exit Do

			If GetForm("CONF_SelectionTypeCode" & idx, 1) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve IndustryTypeCode(MaxIndex) : IndustryTypeCode(MaxIndex) = GetForm("CONF_IndustryTypeCode_S" & idx, 1)
				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = GetForm("CONF_JobTypeCode_S" & idx, 1)
				ReDim Preserve CompanyName(MaxIndex) : CompanyName(MaxIndex) = GetForm("CONF_CompanyName_S" & idx, 1)
				ReDim Preserve SelectionTypeCode(MaxIndex) : SelectionTypeCode(MaxIndex) = GetForm("CONF_SelectionTypeCode" & idx, 1)
				ReDim Preserve MediaCode(MaxIndex) : MediaCode(MaxIndex) = GetForm("CONF_MediaCode" & idx, 1)
				ReDim Preserve OtherMedia(MaxIndex) : OtherMedia(MaxIndex) = GetForm("CONF_OtherMedia" & idx, 1)
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_CompetitionSelection 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_CompetitionSelection '" & (vStaffCode) & "'"
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_CompetitionSelection" & _
			" '" & vStaffCode & "'" & _
			",''" & _
			",'" & IndustryTypeCode(idx) & "'" & _
			",'" & JobTypeCode(idx) & "'" & _
			",'" & CompanyName(idx) & "'" & _
			",'" & SelectionTypeCode(idx) & "'" & _
			",'" & MediaCode(idx) & "'" & _
			",'" & OtherMedia(idx) & "'" & vbCrLf
		Next
	End Function
End Class
%>
