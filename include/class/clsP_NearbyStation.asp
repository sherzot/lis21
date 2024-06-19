<%
'******************************************************************************
'���@�́FclsP_NearbyStation
'�T�@�v�Fform�Ŕ��ł���P_NearbyStation�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_NearbyStation
	Public StaffCode
	Public StationCode()
'	Public StationName()
	Public ToStationBusFlag()
	Public ToStationCarFlag()
	Public ToStationBicycleFlag()
	Public ToStationWalkFlag()
	Public OtherTransportation()
	Public ToStationTime()
'	Public RailwayLineCode()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_NearbyStation �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1

		IsData = False
		MaxIndex = -1
		StaffCode = Request.Form("CONF_StaffCode")

		Err = ""
		If StaffCode <> "" And IsMainCode(StaffCode) = False Then Err = Err & "StaffCode" & vbCrLf

		Do While True
			If ExistsForm("CONF_StationCode" & idx) = False Then Exit Do

			If Request.Form("CONF_StationCode" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve StationCode(MaxIndex) : StationCode(MaxIndex) = Request.Form("CONF_StationCode" & idx)
				ReDim Preserve ToStationBusFlag(MaxIndex) : ToStationBusFlag(MaxIndex) = Request.Form("CONF_ToStationBusFlag" & idx)
				ReDim Preserve ToStationCarFlag(MaxIndex) : ToStationCarFlag(MaxIndex) = Request.Form("CONF_ToStationCarFlag" & idx)
				ReDim Preserve ToStationBicycleFlag(MaxIndex) : ToStationBicycleFlag(MaxIndex) = Request.Form("CONF_ToStationBicycleFlag" & idx)
				ReDim Preserve ToStationWalkFlag(MaxIndex) : ToStationWalkFlag(MaxIndex) = Request.Form("CONF_ToStationWalkFlag" & idx)
				ReDim Preserve OtherTransportation(MaxIndex) : OtherTransportation(MaxIndex) = Request.Form("CONF_OtherTransportation" & idx)
				ReDim Preserve ToStationTime(MaxIndex) : ToStationTime(MaxIndex) = Request.Form("CONF_ToStationTime" & idx)

				'�l�`�F�b�N
				If StationCode(MaxIndex) <> "" And IsNumber(StationCode(MaxIndex), 5, False) = False Then Err = Err & "StationCode(" & MaxIndex & ")" & vbCrLf
				If ToStationBusFlag(MaxIndex) <> "" And IsFlag(ToStationBusFlag(MaxIndex)) = False Then Err = Err & "ToStationBusFlag(" & MaxIndex & ")" & vbCrLf
				If ToStationCarFlag(MaxIndex) <> "" And IsFlag(ToStationCarFlag(MaxIndex)) = False Then Err = Err & "ToStationCarFlag(" & MaxIndex & ")" & vbCrLf
				If ToStationBicycleFlag(MaxIndex) <> "" And IsFlag(ToStationBicycleFlag(MaxIndex)) = False Then Err = Err & "ToStationBicycleFlag(" & MaxIndex & ")" & vbCrLf
				If ToStationWalkFlag(MaxIndex) <> "" And IsFlag(ToStationWalkFlag(MaxIndex)) = False Then Err = Err & "ToStationWalkFlag(" & MaxIndex & ")" & vbCrLf
				If ToStationTime(MaxIndex) <> "" And IsNumber(ToStationTime(MaxIndex), 0, True) = False Then Err = Err & "ToStationTime(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_NearbyStation ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_NearbyStation '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_NearbyStation" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(StationCode(idx)) & "'" & _
				",'" & ChkSQLStr(ToStationBusFlag(idx)) & "'" & _
				",'" & ChkSQLStr(ToStationCarFlag(idx)) & "'" & _
				",'" & ChkSQLStr(ToStationBicycleFlag(idx)) & "'" & _
				",'" & ChkSQLStr(ToStationWalkFlag(idx)) & "'" & _
				",'" & ChkSQLStr(OtherTransportation(idx)) & "'" & _
				",'" & ChkSQLStr(ToStationTime(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
