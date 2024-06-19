<%
'******************************************************************************
'���@�́FclsP_CareerHistoryLis
'�T�@�v�Fform�Ŕ��ł���P_CareerHistoryLis�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_CareerHistoryLis
	Public StaffCode
	Public IndustryTypeCode()
	Public JobTypeCode()
	Public JobTypeDetail()
	Public WorkingTypeCode()
	Public CompanyName()
	Public CompanyName_F()
	Public EntryDay()
	Public RetireDay()
	Public Period()
	Public BusinessDetail()
	Public RetireReason()
	Public Capital()
	Public NumberEmployees()
	Public Summary()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_CareerHistoryLis �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		Dim sEntryDay
		Dim sRetireDay
		Dim idx	: idx = 1

		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")

		Err = ""

		Do While True
			If ExistsForm("CONF_EntryDayCY" & idx) = False Then Exit Do
			sEntryDay = ""
			sRetireDay = ""

			sEntryDay = Request.Form("CONF_EntryDayCY" & idx) & "/"
			If Len(Request.Form("CONF_EntryDayCM" & idx)) = 1 Then sEntryDay = sEntryDay & "0"
			sEntryDay = sEntryDay & Request.Form("CONF_EntryDayCM" & idx) & "/01"
			If IsDate(sEntryDay) = False Then sEntryDay = ""

			sRetireDay = Request.Form("CONF_RetireDayCY" & idx) & "/"
			If Len(Request.Form("CONF_RetireDayCM" & idx)) = 1 Then sRetireDay = sRetireDay & "0"
			sRetireDay = sRetireDay & Request.Form("CONF_RetireDayCM" & idx) & "/01"
			If IsDate(sRetireDay) = False Then sRetireDay = ""

			If IsDate(sEntryDay) = True Then sEntryDay = Replace(sEntryDay, "/", "")
			If IsDate(sRetireDay) = True Then sRetireDay = Replace(sRetireDay, "/", "")

			If sEntryDay <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve IndustryTypeCode(MaxIndex) : IndustryTypeCode(MaxIndex) = Request.Form("CONF_IndustryTypeCode_C" & idx)
				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = Request.Form("CONF_JobTypeCode_C" & idx)
				ReDim Preserve JobTypeDetail(MaxIndex) : JobTypeDetail(MaxIndex) = Request.Form("CONF_JobTypeDetail_C" & idx)
				ReDim Preserve WorkingTypeCode(MaxIndex) : WorkingTypeCode(MaxIndex) = Request.Form("CONF_WorkingTypeCode_C" & idx)
				ReDim Preserve CompanyName(MaxIndex) : CompanyName(MaxIndex) = Request.Form("CONF_CompanyName" & idx)
				ReDim Preserve CompanyName_F(MaxIndex) : CompanyName_F(MaxIndex) = Request.Form("CONF_CompanyName_F" & idx)
				ReDim Preserve EntryDay(MaxIndex) : EntryDay(MaxIndex) = sEntryDay
				ReDim Preserve RetireDay(MaxIndex) : RetireDay(MaxIndex) = sRetireDay
				ReDim Preserve Period(MaxIndex) : Period(MaxIndex) = Request.Form("CONF_Period" & idx)
				ReDim Preserve BusinessDetail(MaxIndex) : BusinessDetail(MaxIndex) = Request.Form("CONF_BusinessDetail_C" & idx)
				ReDim Preserve RetireReason(MaxIndex) : RetireReason(MaxIndex) = Request.Form("CONF_RetireReason" & idx)
				ReDim Preserve Capital(MaxIndex) : Capital(MaxIndex) = Request.Form("CONF_Capital" & idx)
				ReDim Preserve NumberEmployees(MaxIndex) : NumberEmployees(MaxIndex) = Request.Form("CONF_NumberEmployees" & idx)
				ReDim Preserve Summary(MaxIndex) : Summary(MaxIndex) = Request.Form("CONF_Summary" & idx)

				If IndustryTypeCode(MaxIndex) <> "" And IsNumber(IndustryTypeCode(MaxIndex), 3, False) = False Then Err = Err & "IndustryTypeCode(" & MaxIndex & ")" & vbCrLf
				If JobTypeCode(MaxIndex) <> "" And IsNumber(JobTypeCode(MaxIndex), 3, False) = False Then Err = Err & "JobTypeCode(" & MaxIndex & ")" & vbCrLf
				If WorkingTypeCode(MaxIndex) <> "" And IsNumber(WorkingTypeCode(MaxIndex), 3, False) = False Then Err = Err & "WorkingTypeCode(" & MaxIndex & ")" & vbCrLf
				If EntryDay(MaxIndex) <> "" And IsDay(sEntryDay) = False Then Err = Err & "EntryDay(" & MaxIndex & ")" & vbCrLf
				If RetireDay(MaxIndex) <> "" And IsDay(sRetireDay) = False Then Err = Err & "RetireDay(" & MaxIndex & ")" & vbCrLf
				If Period(MaxIndex) <> "" And IsRE(Period(MaxIndex), 0, True) = False Then Err = Err & "Period(" & MaxIndex & ")" & vbCrLf
				If Capital(MaxIndex) <> "" And IsNumber(Capital(MaxIndex), 0, False) = False Then Err = Err & "Capital(" & MaxIndex & ")" & vbCrLf
				If NumberEmployees(MaxIndex) <> "" And IsNumber(NumberEmployees(MaxIndex), 0, False) = False Then Err = Err & "NumberEmployees(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_CareerHistoryLis ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_CareerHistoryLis '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC up_Reg_P_CareerHistoryLis" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(IndustryTypeCode(idx)) & "'" & _
				",'" & ChkSQLStr(JobTypeCode(idx)) & "'" & _
				",'" & ChkSQLStr(JobTypeDetail(idx)) & "'" & _
				",'" & ChkSQLStr(WorkingTypeCode(idx)) & "'" & _
				",'" & ChkSQLStr(CompanyName(idx)) & "'" & _
				",'" & ChkSQLStr(CompanyName_F(idx)) & "'" & _
				",'" & ChkSQLStr(EntryDay(idx)) & "'" & _
				",'" & ChkSQLStr(RetireDay(idx)) & "'" & _
				",'" & ChkSQLStr(Period(idx)) & "'" & _
				",'" & ChkSQLStr(BusinessDetail(idx)) & "'" & _
				",'" & ChkSQLStr(RetireReason(idx)) & "'" & _
				",'" & ChkSQLStr(Capital(idx)) & "'" & _
				",'" & ChkSQLStr(NumberEmployees(idx)) & "'" & _
				",'" & ChkSQLStr(Summary(idx)) & "'" & vbCrLf
		Next
	End Function
	
End Class
%>
