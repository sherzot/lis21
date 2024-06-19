<%
'******************************************************************************
'���@�́FclsN_Image
'�T�@�v�Fform�Ŕ��ł���N_Image�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/10/06
'�X�@�V�F
'******************************************************************************
Class clsN_Image
	Public StaffCode
	Public CategoryCode
	Public Code
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsN_Image �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/10/06
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize(ByVal vCategoryCode)
		Dim sTemp

		CategoryCode = vCategoryCode
		If Request.Form("conf_" & CategoryCode) <> "" Then sTemp = Replace(Request.Form("conf_" & CategoryCode), " ", "")

		IsData = False
		If sTemp <> "" Then
			Code = Split(sTemp, ",")
			MaxIndex = UBound(Code)
			IsData = True
		Else
			MaxIndex = -1
		End If

		Err = ""
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fup_Reg_N_Image ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC up_Del_N_Image '" & ChkSQLStr(vStaffCode) & "', '" & ChkSQLStr(CategoryCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC up_Reg_N_Image" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",'" & ChkSQLStr(CategoryCode) & "'" & _
				",'" & ChkSQLStr(Code(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
