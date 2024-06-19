<%
'******************************************************************************
'名　称：clsP_HopeWorkingPlace
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
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
	'名　称：Initialize
	'概　要：clsP_HopeWorkingPlace クラスの初期化関数
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
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopeWorkingPlace 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
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
