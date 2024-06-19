<%
'******************************************************************************
'���@�́FclsN_School
'�T�@�v�Fform�Ŕ��ł���P_UserInfo�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/10/06
'�X�@�V�F
'******************************************************************************
Class clsN_School
	Public StaffCode
	Public SchoolCode
	Public OtherSchoolName
	Public Department
	Public Studyroom
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsN_School�N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("conf_staffcode") <> "" Then StaffCode = Request.Form("conf_staffcode")
		If Request.Form("conf_schoolcode") <> "" Then SchoolCode = Request.Form("conf_schoolcode")
		If Request.Form("conf_otherschoolname") <> "" Then OtherSchoolName = Request.Form("conf_otherschoolname")
		If Request.Form("conf_department") <> "" Then Department = Request.Form("conf_department")
		If Request.Form("conf_studyroom") <> "" Then Studyroom = Request.Form("conf_studyroom")

		If SchoolCode & OtherSchoolName <> "" Then IsData = True

		'�l�`�F�b�N
		Err = ""
		If StaffCode <> "" And IsMainCode(StaffCode) = False Then Err = Err & "StaffCode" & vbCrLf
		If SchoolCode <> "" And IsRE(SchoolCode, "^\d\d\d\d\d$", False) = False Then Err = Err & "SchoolCode" & vbCrLf
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_ UserInfo���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "up_Reg_N_School" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(SchoolCode) & "'" & _
			",'" & ChkSQLStr(OtherSchoolName) & "'" & _
			",'" & ChkSQLStr(Department) & "'" & _
			",'" & ChkSQLStr(Studyroom) & "'" & vbCrLf
	End Function
End Class
%>
