<%
'******************************************************************************
'���@�́FclsP_Skill
'�T�@�v�Fform�Ŕ��ł���P_Skill�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_Skill
	Public StaffCode
	Public CategoryCode
	Public Code()
	Public StartDay()
	Public Period()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_Skill �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize(vCategoryCode)
		Dim sStartDay
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")

		Err = ""

		Select Case vCategoryCode
			Case "OS":		CategoryCode = "OS"
			Case "App":		CategoryCode = "Application"
			Case "Lang":	CategoryCode = "DevelopmentLanguage"
			Case "DB":		CategoryCode = "Database"
			Case Else:		CategoryCode = vCategoryCode
		End Select

		Do While True
			If ExistsForm("CONF_" & vCategoryCode & idx) = False Then Exit Do
			sStartDay = Request.Form("CONF_StartDay" & vCategoryCode & "Y" & idx) & "/"
			If Len(Request.Form("CONF_StartDay" & vCategoryCode & "M" & idx)) = 1 Then sStartDay = sStartDay & "0"
			sStartDay = sStartDay & Request.Form("CONF_StartDay" & vCategoryCode & "M" & idx) & "/01"
			If IsDate(sStartDay) = False Then sStartDay = ""
			sStartDay = Replace(sStartDay, "/", "")

			If Request.Form("CONF_" & vCategoryCode & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve Code(MaxIndex) : Code(MaxIndex) = Request.Form("CONF_" & vCategoryCode & idx)
				ReDim Preserve StartDay(MaxIndex) : StartDay(MaxIndex) = sStartDay
				ReDim Preserve Period(MaxIndex) : Period(MaxIndex) = Request.Form("CONF_Period_" & vCategoryCode & idx)
				If Code(MaxIndex) <> "" And IsNumber(Code(MaxIndex), 3, False) = False Then Err = Err & "Code(" & MaxIndex & ")" & vbCrLf
				If StartDay(MaxIndex) <> "" And IsDay(StartDay(MaxIndex)) = False Then Err = Err & "StartDay(" & MaxIndex & ")" & vbCrLf
				If Period(MaxIndex) <> "" And IsNumber(Period(MaxIndex), 0, False) = False Then Err = Err & "Period(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_Skill ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_Skill '" & ChkSQLStr(vStaffCode) & "', '" & ChkSQLStr(CategoryCode) & "'" & vbCrLf
		'If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_Skill" & _
				" '" & ChkSQLStr(vStaffCode )& "'"  & _
				",''"  & _
				",'" & ChkSQLStr(CategoryCode) & "'" & _
				",'" & ChkSQLStr(Code(idx)) & "'" & _
				",'" & ChkSQLStr(StartDay(idx)) & "'" & _
				",'" & ChkSQLStr(Period(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
