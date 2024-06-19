<%
'******************************************************************************
'���@�́FclsP_HopeIndustryType
'�T�@�v�Fform�Ŕ��ł���P_HopeIndustryType�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_HopeIndustryType
	Public StaffCode
	Public IndustryTypeCode
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_HopIndustryType �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")
		If Request.Form("CONF_IndustryTypeCode") <> "" Then IsData = True: IndustryTypeCode = Request.Form("CONF_IndustryTypeCode")

		Err = ""

		If IndustryTypeCode <> "" And IsNumber(IndustryTypeCode, 3, False) = False Then Err = Err & "IndustryTypeCode" & vbCrLf
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_HopIndustryType ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		GetRegSQL = "EXEC sp_Del_P_HopeIndustryType '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeIndustryType" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",''" & _
			",'" & ChkSQLStr(IndustryTypeCode) & "'" & vbCrLf
	End Function
End Class
%>
