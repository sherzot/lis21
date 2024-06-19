<%
'******************************************************************************
'���@�́FclsP_Talent
'�T�@�v�Fform�Ŕ��ł���P_�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
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
	'���@�́FInitialize
	'�T�@�v�FclsP_Talent �N���X�̏������֐�
	'���@�l�F�Г��V�X�e���ł͗p�Ȃ�
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_CompanyCode") <> "" Then IsData = True: CompanyCode = Request.Form("CONF_CompanyCode")
		If Request.Form("CONF_LicenseNumber") <> "" Then IsData = True: LicenseNumber = Request.Form("CONF_LicenseNumber")
		If Request.Form("CONF_EmploymentDivisionFlag") <> "" Then IsData = True: EmploymentDivisionFlag = Request.Form("CONF_EmploymentDivisionFlag")
		If Request.Form("CONF_RecommendationLetter") <> "" Then IsData = True: RecommendationLetter = Request.Form("CONF_RecommendationLetter")
		If Request.Form("CONF_WorkDivisionFlag") <> "" Then IsData = True: WorkDivisionFlag = Request.Form("CONF_WorkDivisionFlag")
		If Request.Form("CONF_PL_StateFlag") <> "" Then IsData = True: PL_StateFlag = Request.Form("CONF_PL_StateFlag")

		Err = ""
		If CompanyCode <> "" And IsMainCode(CompanyCode) = False Then Err = Err & "CompanyCode" & vbCrLf
		If LicenseNumber <> "" And IsNumber(LicenseNumber, 0, False) = False Then Err = Err & "LicenseNumber" & vbCrLf
		If EmploymentDivisionFlag <> "" And IsFlag(EmploymentDivisionFlag) = False Then Err = Err & "EmploymentDivisionFlag" & vbCrLf
		If WorkDivisionFlag <> "" And IsFlag(WorkDivisionFlag) = False Then Err = Err & "WorkDivisionFlag" & vbCrLf
		If PL_StateFlag <> "" And IsFlag(PL_StateFlag) = False Then Err = Err & "PL_StateFlag" & vbCrLf
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_Talent ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_Talent" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(CompanyCode) & "'" & _
			",'" & ChkSQLStr(LicenseNumber) & "'" & _
			",'" & ChkSQLStr(EmploymentDivisionFlag) & "'" & _
			",'" & ChkSQLStr(RecommendationLetter) & "'" & _
			",'" & ChkSQLStr(WorkDivisionFlag) & "'" & _
			",'" & ChkSQLStr(PL_StateFlag) & "'" & vbCrLf
	End Function
End Class
%>
