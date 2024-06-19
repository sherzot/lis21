<%
'******************************************************************************
'���@�́FclsP_HopeCommuting
'�T�@�v�Fform�Ŕ��ł���P_�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F��]�w��
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_HopeCommuting
	Public StaffCode
	Public StationCode()
	Public MinuteToStation()
	Public CommuteTime()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_HopeCommuting �N���X�̏������֐�
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
			If ExistsForm("CONF_StationCodeHope" & idx) = False Then Exit Do

			If Request.Form("CONF_StationCodeHope" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve StationCode(MaxIndex):		StationCode(MaxIndex) = Request.Form("CONF_StationCodeHope" & idx,")
				ReDim Preserve MinuteToStation(MaxIndex):	MinuteToStation(MaxIndex) = Request.Form("CONF_MinuteToStation" & idx,")
				ReDim Preserve CommuteTime(MaxIndex):		CommuteTime(MaxIndex) = Request.Form("CONF_HopeCommuteTime" & idx,")

				If StationCode(MaxIndex) <> "" And IsNumber(StationCode(MaxIndex), 5, False) = False Then Err = Err & "StationCode(" & MaxIndex & ")" & vbCrLf
				If MinuteToStation(MaxIndex) <> "" And IsNumber(MinuteToStation(MaxIndex), 0, False) = False Then Err = Err & "MinuteToStation(" & MaxIndex & ")" & vbCrLf
				If CommuteTime(MaxIndex) <> "" And IsNumber(CommuteTime(MaxIndex), 0, False) = False Then Err = Err & "CommuteTime(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_HopeCommuting ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeCommuting '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeCommuting" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(StationCode(idx)) & "'" & _
				",'" & ChkSQLStr(MinuteToStation(idx)) & "'" & _
				",'" & ChkSQLStr(CommuteTime(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
