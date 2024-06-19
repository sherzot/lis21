<%
'******************************************************************************
'���@�́FclsP_EducateHistory
'�T�@�v�Fform�Ŕ��ł���P_EducateHistory�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_EducateHistory
	Public StaffCode
	Public EntryDay()
	Public GraduateDay()
	Public EntryTypeCode()
	Public GraduateTypeCode()
	Public SchoolName()
	Public SchoolTypeCode()
	Public Speciality()
	Public CourseType()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_EducateHistory �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		Dim sEntryDay
		Dim sGraduateDay
		Dim idx	: idx = 1

		IsData = False
		MaxIndex = -1

		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")

		Err = ""

		Do While True
			If ExistsForm("CONF_EntryDayY" & idx) = False Then Exit Do

			sEntryDay = ""
			sGraduateDay = ""

			sEntryDay = Request.Form("CONF_EntryDayY" & idx) & "/"
			If Len(Request.Form("CONF_EntryDayM" & idx)) = 1 Then sEntryDay = sEntryDay & "0"
			sEntryDay = sEntryDay & Request.Form("CONF_EntryDayM" & idx) & "/01"
			If IsDate(sEntryDay) = False Then sEntryDay = ""

			sGraduateDay = Request.Form("CONF_GraduateDayY" & idx) & "/"
			If Len(Request.Form("CONF_GraduateDayM" & idx)) = 1 Then sGraduateDay = sGraduateDay & "0"
			sGraduateDay = sGraduateDay & Request.Form("CONF_GraduateDayM" & idx) & "/01"
			If IsDate(sGraduateDay) = False Then sGraduateDay = ""

			If IsDate(sEntryDay) = True Then sEntryDay = Replace(sEntryDay, "/", "")
			If IsDate(sGraduateDay) = True Then sGraduateDay = Replace(sGraduateDay, "/", "")

			If sEntryDay <> "" Or sGraduateDay <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve EntryDay(MaxIndex) : EntryDay(MaxIndex) = sEntryDay
				ReDim Preserve GraduateDay(MaxIndex) : GraduateDay(MaxIndex) = sGraduateDay
				ReDim Preserve EntryTypeCode(MaxIndex) : EntryTypeCode(MaxIndex) = Request.Form("CONF_EntryTypeCode" & idx)
				ReDim Preserve GraduateTypeCode(MaxIndex) : GraduateTypeCode(MaxIndex) = Request.Form("CONF_GraduateTypeCode" & idx)
				ReDim Preserve SchoolName(MaxIndex) : SchoolName(MaxIndex) = Request.Form("CONF_SchoolName" & idx)
				ReDim Preserve SchoolTypeCode(MaxIndex) : SchoolTypeCode(MaxIndex) = Request.Form("CONF_SchoolTypeCode" & idx)
				ReDim Preserve Speciality(MaxIndex) : Speciality(MaxIndex) = Request.Form("CONF_Speciality" & idx)
				ReDim Preserve CourseType(MaxIndex) : CourseType(MaxIndex) = Request.Form("CONF_CourseType" & idx)

				'�l�`�F�b�N
				If EntryDay(MaxIndex) <> "" And IsDay(EntryDay(MaxIndex)) = False Then Err = Err & "EntryDay(" & MaxIndex & ")" & vbCrLf
				If GraduateDay(MaxIndex) <> "" And IsDay(GraduateDay(MaxIndex)) = False Then Err = Err & "GraduateDay(" & MaxIndex & ")" & vbCrLf
				If EntryTypeCode(MaxIndex) <> "" And IsNumber(EntryTypeCode(MaxIndex), 3, False) = False Then Err = Err & "EntryTypeCode(" & MaxIndex & ")" & vbCrLf
				If GraduateTypeCode(MaxIndex) <> "" And IsNumber(GraduateTypeCode(MaxIndex), 3, False) = False Then Err = Err & "GraduateTypeCode(" & MaxIndex & ")" & vbCrLf
				If SchoolTypeCode(MaxIndex) <> "" And IsNumber(SchoolTypeCode(MaxIndex), 3, False) = False Then Err = Err & "SchoolTypeCode(" & MaxIndex & ")" & vbCrLf
				If CourseType(MaxIndex) <> "" And IsRE(CourseType(MaxIndex), "^[123]$", True) = False Then Err = Err & "CourseType(" & MaxIndex & ")" & vbCrLf
			End If

			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_EducateHistory ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_EducateHistory '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_EducateHistory" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(EntryDay(idx)) & "'" & _
				",'" & ChkSQLStr(GraduateDay(idx)) & "'" & _
				",'" & ChkSQLStr(EntryTypeCode(idx)) & "'" & _
				",'" & ChkSQLStr(GraduateTypeCode(idx)) & "'" & _
				",'" & ChkSQLStr(SchoolName(idx)) & "'" & _
				",'" & ChkSQLStr(SchoolTypeCode(idx)) & "'" & _
				",'" & ChkSQLStr(Speciality(idx)) & "'" & _
				",'" & ChkSQLStr(CourseType(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
