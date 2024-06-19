<%
'******************************************************************************
'���@�́FclsP_License
'�T�@�v�Fform�Ŕ��ł���P_License�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_License
	Public StaffCode
	Public GroupCode()
	Public CategoryCode()
	Public Code()
	Public GetDay()
	Public Remark()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_License �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		Dim sGetDay
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")

		Err = ""

		Do While True
			If ExistsForm("CONF_LicenseCode" & idx) = False Then Exit Do
			sGetDay = ""

			sGetDay = Request.Form("CONF_GetDayY" & idx) & "/"
			If Len(Request.Form("CONF_GetDayM" & idx)) = 1 Then sGetDay = sGetDay & "0"
			sGetDay = sGetDay & Request.Form("CONF_GetDayM" & idx) & "/01"
			If IsDate(sGetDay) = False Then sGetDay = ""
			sGetDay = Replace(sGetDay, "/", "")

			If Request.Form("CONF_LicenseCode" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve GroupCode(MaxIndex) : GroupCode(MaxIndex) = Mid(Request.Form("CONF_LicenseCode" & idx), 1, 2)
				ReDim Preserve CategoryCode(MaxIndex) : CategoryCode(MaxIndex) = Mid(Request.Form("CONF_LicenseCode" & idx), 3, 3)
				ReDim Preserve Code(MaxIndex) : Code(MaxIndex) = Mid(Request.Form("CONF_LicenseCode" & idx), 6, 2)
				ReDim Preserve GetDay(MaxIndex) : GetDay(MaxIndex) = sGetDay
				ReDim Preserve Remark(MaxIndex) : Remark(MaxIndex) = Request.Form("CONF_LicenseRemark" & idx)

				If GroupCode(MaxIndex) <> "" And IsNumber(GroupCode(MaxIndex), 2, False) = False Then Err = Err & "GroupCode(" & MaxIndex & ")" & vbCrLf
				If CategoryCode(MaxIndex) <> "" And IsNumber(CategoryCode(MaxIndex), 3, False) = False Then Err = Err & "CategoryCode(" & MaxIndex & ")" & vbCrLf
				If Code(MaxIndex) <> "" And IsNumber(Code(MaxIndex), 2, False) = False Then Err = Err & "Code(" & MaxIndex & ")" & vbCrLf
				If GetDay(MaxIndex) <> "" And IsDay(GetDay(MaxIndex)) = False Then Err = Err & "GetDay(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_License ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_License '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_License" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(GroupCode(idx)) & "'" & _
				",'" & ChkSQLStr(CategoryCode(idx)) & "'" & _
				",'" & ChkSQLStr(Code(idx)) & "'" & _
				",'" & ChkSQLStr(GetDay(idx)) & "'" & _
				",'" & ChkSQLStr(Remark(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
