<%
'******************************************************************************
'���@�́FclsP_BankAccount
'�T�@�v�Fform�Ŕ��ł���P_BankAccount�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
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
	'���@�́FInitialize
	'�T�@�v�FclsP_BankAccount �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_BankName") <> "" Then IsData = True: BankName = Request.Form("CONF_BankName")
		If Request.Form("CONF_BankNo") <> "" Then IsData = True: BankNo = Request.Form("CONF_BankNo")
		If Request.Form("CONF_BankBranchName") <> "" Then IsData = True: BankBranchName = Request.Form("CONF_BankBranchName")
		If Request.Form("CONF_BankBranchNo") <> "" Then IsData = True: BankBranchNo = Request.Form("CONF_BankBranchNo")
		If Request.Form("CONF_AccountNo") <> "" Then IsData = True: AccountNo = Request.Form("CONF_AccountNo")
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_BankAccount ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_BankAccount" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(BankName) & "'" & _
			",'" & ChkSQLStr(BankNo) & "'" & _
			",'" & ChkSQLStr(BankBranchName) & "'" & _
			",'" & ChkSQLStr(BankBranchNo) & "'" & _
			",'" & ChkSQLStr(AccountNo) & "'" & vbCrLf
	End Function
End Class
%>
