<%
'******************************************************************************
'���@�́FclsP_HopeWorkingType
'�T�@�v�Fform�Ŕ��ł���P_�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_HopeWorkingType
	Public StaffCode
	Public WorkingTypeCode()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_HopeWorkingType �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")

		Do While True
			If ExistsForm("CONF_HopeWorkingType" & idx) = False Then Exit Do

			If Request.Form("CONF_HopeWorkingType" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve WorkingTypeCode(MaxIndex) : WorkingTypeCode(MaxIndex) = Request.Form("CONF_HopeWorkingType" & idx)

				If WorkingTypeCode(MaxIndex) <> "" And IsNumber(WorkingTypeCode(MaxIndex), 3, False) = False Then Err = Err & "WorkingTypeCode(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_HopeWorkingType ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeWorkingType '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeWorkingType" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(WorkingTypeCode(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
