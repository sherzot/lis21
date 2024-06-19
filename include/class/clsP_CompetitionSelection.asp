<%
'******************************************************************************
'���@�́FclsP_CompetitionSelection
'�T�@�v�Fform�Ŕ��ł���P_�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_CompetitionSelection
	Public StaffCode
	Public IndustryTypeCode()
	Public JobTypeCode()
	Public CompanyName()
	Public SelectionTypeCode()
	Public MediaCode()
	Public OtherMedia()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_CompetitionSelection �N���X�̏������֐�
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
			If ExistsForm("CONF_SelectionTypeCode" & idx) = False Then Exit Do

			If Request.Form("CONF_SelectionTypeCode" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve IndustryTypeCode(MaxIndex) : IndustryTypeCode(MaxIndex) = Request.Form("CONF_IndustryTypeCode_S" & idx)
				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = Request.Form("CONF_JobTypeCode_S" & idx)
				ReDim Preserve CompanyName(MaxIndex) : CompanyName(MaxIndex) = Request.Form("CONF_CompanyName_S" & idx)
				ReDim Preserve SelectionTypeCode(MaxIndex) : SelectionTypeCode(MaxIndex) = Request.Form("CONF_SelectionTypeCode" & idx)
				ReDim Preserve MediaCode(MaxIndex) : MediaCode(MaxIndex) = Request.Form("CONF_MediaCode" & idx)
				ReDim Preserve OtherMedia(MaxIndex) : OtherMedia(MaxIndex) = Request.Form("CONF_OtherMedia" & idx)
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_CompetitionSelection ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_CompetitionSelection '" & ChkSQLStr(vStaffCode) & "'"
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_CompetitionSelection" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",''" & _
			",'" & ChkSQLStr(IndustryTypeCode(idx)) & "'" & _
			",'" & ChkSQLStr(JobTypeCode(idx)) & "'" & _
			",'" & ChkSQLStr(CompanyName(idx)) & "'" & _
			",'" & ChkSQLStr(SelectionTypeCode(idx)) & "'" & _
			",'" & ChkSQLStr(MediaCode(idx)) & "'" & _
			",'" & ChkSQLStr(OtherMedia(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
