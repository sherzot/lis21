<%
'******************************************************************************
'名　称：clsN_UserInfo
'概　要：formで飛んできたP_UserInfoテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/10/06
'更　新：
'******************************************************************************
Class clsN_UserInfo
	Public StaffCode
	Public PenName
	Public Password
	Public GraduateYear
	Public OperateClassComCode
	Public OperateClassWebCode
	Public OperateClassRemark
	Public BranchCode
	Public EmployeeCode
	Public MailMagazineFlag
	Public NewJohoMailFlag
	Public SuspensionFlag
	Public ErasureFlag
	Public NaviUseFlag
	Public HomeContactFlag
	Public PortableContactFlag
	Public MailContactFlag
	Public PersonDangerFlag
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

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_PenName") <> "" Then PenName = Request.Form("CONF_PenName")
		If Request.Form("CONF_Password") <> "" Then Password = Request.Form("CONF_Password")
		If Request.Form("CONF_GraduateYear") <> "" Then GraduateYear = Request.Form("CONF_GraduateYear")
		If Request.Form("CONF_OperateClassComCode") <> "" Then OperateClassComCode = Request.Form("CONF_OperateClassComCode")
		If Request.Form("CONF_OperateClassWebCode") <> "" Then OperateClassWebCode = Request.Form("CONF_OperateClassWebCode")
		If Request.Form("CONF_OperateClassRemark") <> "" Then OperateClassRemark = Request.Form("CONF_OperateClassRemark")
		If Request.Form("CONF_BranchCode") <> "" Then BranchCode = Request.Form("CONF_BranchCode")
		If Request.Form("CONF_EmployeeCode") <> "" Then EmployeeCode = Request.Form("CONF_EmployeeCode")
		If Request.Form("CONF_MailMagazineFlag") <> "" Then MailMagazineFlag = Request.Form("CONF_MailMagazineFlag")
		If Request.Form("CONF_NewJohoMailFlag") <> "" Then NewJohoMailFlag = Request.Form("CONF_NewJohoMailFlag")
		If Request.Form("CONF_SuspensionFlag") <> "" Then SuspensionFlag = Request.Form("CONF_SuspensionFlag")
		If Request.Form("CONF_ErasureFlag") <> "" Then ErasureFlag = Request.Form("CONF_ErasureFlag")
		If Request.Form("CONF_NaviUseFlag") <> "" Then NaviUseFlag = Request.Form("CONF_NaviUseFlag")
		If Request.Form("CONF_HomeContactFlag") <> "" Then HomeContactFlag = Request.Form("CONF_HomeContactFlag")
		If Request.Form("CONF_PortableContactFlag") <> "" Then PortableContactFlag = Request.Form("CONF_PortableContactFlag")
		If Request.Form("CONF_MailContactFlag") <> "" Then MailContactFlag = Request.Form("CONF_MailContactFlag")
		If Request.Form("CONF_PersonDangerFlag") <> "" Then PersonDangerFlag = Request.Form("CONF_PersonDangerFlag")
		If Request.Form("CONF_HopeCommuteTime") <> "" Then HopeCommuteTime = Request.Form("CONF_HopeCommuteTime")
		If Request.Form("CONF_NoticeMailFlag") <> "" Then NoticeMailFlag = Request.Form("CONF_NoticeMailFlag")
		If Request.Form("CONF_LisReserveDay") <> "" Then LisReserveDay = Request.Form("CONF_LisReserveDay")
		If Request.Form("CONF_LisRegistDay") <> "" Then LisRegistDay = Request.Form("CONF_LisRegistDay")

		If StaffCode & Password & GraduateYear <> "" Then IsData = True

		'値チェック
		Err = ""
		If StaffCode <> "" And IsMainCode(StaffCode) = False Then Err = Err & "StaffCode" & vbCrLf
		If GraduateYear <> "" And IsDay(GraduateYear & "01") = False Then Err = Err & "GraduateYear" & vbCrLf
		If OperateClassComCode <> "" And IsNumber(OperateClassComCode, 3, False) = False Then Err = Err & "OperateClassComCode" & vbCrLf
		If OperateClassWebCode <> "" And IsNumber(OperateClassWebCode, 3, False) = False Then Err = Err & "OperateClassWebCode" & vbCrLf
		If BranchCode <> "" And IsRE(BranchCode, "^[A-Z][A-Z]$", True) = False Then Err = Err & "BranchCode" & vbCrLf
		If EmployeeCode <> "" And IsMainCode(EmployeeCode) = False Then Err = Err & "EmployeeCode" & vbCrLf
		If MailMagazineFlag <> "" And IsFlag(MailMagazineFlag) = False Then Err = Err & "MailMagazineFlag" & vbCrLf
		If NewJohoMailFlag <> "" And IsFlag(NewJohoMailFlag) = False Then Err = Err & "NewJohoMailFlag" & vbCrLf
		If SuspensionFlag <> "" And IsFlag(SuspensionFlag) = False Then Err = Err & "SuspensionFlag" & vbCrLf
		If ErasureFlag <> "" And IsFlag(ErasureFlag) = False Then Err = Err & "ErasureFlag" & vbCrLf
		If NaviUseFlag <> "" And IsFlag(NaviUseFlag) = False Then Err = Err & "NaviUseFlag" & vbCrLf
		If HomeContactFlag <> "" And IsFlag(HomeContactFlag) = False Then Err = Err & "HomeContactFlag" & vbCrLf
		If PortableContactFlag <> "" And IsFlag(PortableContactFlag) = False Then Err = Err & "PortableContactFlag" & vbCrLf
		If MailContactFlag <> "" And IsFlag(MailContactFlag) = False Then Err = Err & "MailContactFlag" & vbCrLf
		If PersonDangerFlag <> "" And IsFlag(PersonDangerFlag) = False Then Err = Err & "PersonDangerFlag" & vbCrLf
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
		If IsData = False Then Exit Function

		GetRegSQL = "up_Reg_N_UserInfo" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(PenName) & "'" & _
			",'" & ChkSQLStr(Password) & "'" & _
			",'" & ChkSQLStr(GraduateYear) & "'" & _
			",'" & ChkSQLStr(OperateClassComCode) & "'" & _
			",'" & ChkSQLStr(OperateClassWebCode) & "'" & _
			",'" & ChkSQLStr(OperateClassRemark) & "'" & _
			",'" & ChkSQLStr(BranchCode) & "'" & _
			",'" & ChkSQLStr(EmployeeCode) & "'" & _
			",'" & ChkSQLStr(MailMagazineFlag) & "'" & _
			",'" & ChkSQLStr(NewJohoMailFlag) & "'" & _
			",'" & ChkSQLStr(SuspensionFlag) & "'" & _
			",'" & ChkSQLStr(ErasureFlag) & "'" & _
			",'" & ChkSQLStr(NaviUseFlag) & "'" & _
			",'" & ChkSQLStr(HomeContactFlag) & "'" & _
			",'" & ChkSQLStr(PortableContactFlag) & "'" & _
			",'" & ChkSQLStr(MailContactFlag) & "'" & _
			",'" & ChkSQLStr(PersonDangerFlag) & "'" & _
			",'" & ChkSQLStr(HopeCommuteTime) & "'" & _
			",'" & ChkSQLStr(NoticeMailFlag) & "'" & _
			",'" & ChkSQLStr(LisReserveDay) & "'" & _
			",'" & ChkSQLStr(LisRegistDay) & "'" & vbCrLf
	End Function
End Class
%>
