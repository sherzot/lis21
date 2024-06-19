<%
'******************************************************************************
'���@�́FclsP_HopeJobType
'�T�@�v�Fform�Ŕ��ł���P_HopeJobType�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_HopeJobType
	Public StaffCode
	Public JobTypeCode()
	Public JobTypeDetail()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_HopeJobType �N���X�̏������֐�
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

		Err = ""

		Do While True
			If ExistsForm("CONF_JobTypeCode" & idx) = False Then Exit Do

			If Request.Form("CONF_JobTypeCode" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = Request.Form("CONF_JobTypeCode" & idx)
				ReDim Preserve JobTypeDetail(MaxIndex) : JobTypeDetail(MaxIndex) = Request.Form("CONF_JobTypeDetail" & idx)

				If JobTypeCode(MaxIndex) <> "" And IsNumber(JobTypeCode(MaxIndex), 3, False) = False Then Err = Err & "JobTypeCode(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_HopeJobType ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeJobType '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeJobType" & _
				" '" & ChkSQLStr(vStaffCode )& "'"  & _
				",''" & _
				",'" & ChkSQLStr(JobTypeCode(idx)) & "'" & _
				",'" & ChkSQLStr(JobTypeDetail(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
