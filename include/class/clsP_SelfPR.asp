<%
'******************************************************************************
'���@�́FclsP_SelfPR
'�T�@�v�Fform�Ŕ��ł���P_SelfPR�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�Fnavi only
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_SelfPR
	Public StaffCode
	Public SelfPR
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_SelfPR �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If ExistsForm("CONF_SelfPR") = False Then Exit Sub

		IsData = True
		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_SelfPR") <> "" Then SelfPR = Request.Form("CONF_SelfPR")
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_SelfPR ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "sp_Reg_P_SelfPR" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(SelfPR) & "'" & vbCrLf
	End Function
End Class
%>
