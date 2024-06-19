<%
'******************************************************************************
'���@�́FclsP_HopeWorkingPlace
'�T�@�v�Fform�Ŕ��ł���P_�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_HopeWorkingPlace
	Public StaffCode
	Public PrefectureCode()
	Public City()
	Public Area()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_HopeWorkingPlace �N���X�̏������֐�
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
			If ExistsForm("CONF_HopePrefecture" & idx) = False Then Exit Do

			If Request.Form("CONF_HopePrefecture" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve PrefectureCode(MaxIndex): PrefectureCode(MaxIndex) = Request.Form("CONF_HopePrefecture" & idx)
				ReDim Preserve City(MaxIndex): City(MaxIndex) = Request.Form("CONF_HopeCity" & idx)
				ReDim Preserve Area(MaxIndex): Area(MaxIndex) = Request.Form("CONF_HopeArea" & idx)

				If PrefectureCode(MaxIndex) <> "" And IsNumber(PrefectureCode(MaxIndex), 3, False) = False Then Err = Err & "PrefectureCode(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_HopeWorkingPlace ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeWorkingPlace '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeWorkingPlace" & _
				" '" & ChkSQLStr(vStaffCode )& "'"  & _
				",''" & _
				",'" & ChkSQLStr(PrefectureCode(idx)) & "'" & _
				",'" & ChkSQLStr(City(idx)) & "'" & _
				",'" & ChkSQLStr(Area(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
