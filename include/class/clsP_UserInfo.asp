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
	Public NoticeMailFlag
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
	End Sub

	Public Function ChkData()
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
		If NoticeMailFlag <> "" And IsNumber(NoticeMailFlag, 1, False) = False Then Err = Err & "NoticeMailFlag" & vbCrLf
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
		GetRegSQL = "up_Reg_P_UserInfo '" & vStaffCode & "'" & _
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
			",'" & NoticeMailFlag & "'" & _
			",'" & LisReserveDay & "'" & _
			",'" & LisRegistDay & "'"
	End Function
End Class
%>
