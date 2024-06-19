<%
'******************************************************************************
'名　称：clsP_HopeCommuting
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：希望駅等
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
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
	'名　称：Initialize
	'概　要：clsP_HopeCommuting クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
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
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopeCommuting 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
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
