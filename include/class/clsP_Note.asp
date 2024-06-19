<%
'******************************************************************************
'���@�́FclsP_Note
'�T�@�v�Fform�Ŕ��ł���P_Note�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
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
	'���@�́FInitialize
	'�T�@�v�FclsP_Note �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize(vCode)
		IsData = False
		MaxIndex = -1

		StaffCode = Request.Form("CONF_StaffCode")
		CategoryCode = "Note"
		Code = vCode
		Note = Request.Form("CONF_Note_" Code, 1)
		If Note <> "" Then IsData = True
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_Note ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		GetRegSQL = "EXEC sp_Del_P_Note '" & ChkSQLStr(vStaffCode) & "', '" & ChkSQLStr(CategoryCode) & "', '" & ChkSQLStr(Code) & "'" & vbCrLf
		If IsData = False Then Exit Function
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_Note" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(CategoryCode) & "'" & _
			",'" & ChkSQLStr(Code) & "'" & _
			",'" & ChkSQLStr(Note) & "'" & vbCrLf
	End Function
End Class
%>
