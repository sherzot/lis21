<%
'******************************************************************************
'名　称：clsP_HopeIndustryType
'概　要：formで飛んできたP_HopeIndustryTypeテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_HopeIndustryType
	Public StaffCode
	Public IndustryTypeCode
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_HopIndustryType クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")
		If Request.Form("CONF_IndustryTypeCode") <> "" Then IsData = True: IndustryTypeCode = Request.Form("CONF_IndustryTypeCode")

		Err = ""

		If IndustryTypeCode <> "" And IsNumber(IndustryTypeCode, 3, False) = False Then Err = Err & "IndustryTypeCode" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopIndustryType 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		GetRegSQL = "EXEC sp_Del_P_HopeIndustryType '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeIndustryType" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",''" & _
			",'" & ChkSQLStr(IndustryTypeCode) & "'" & vbCrLf
	End Function
End Class
%>
