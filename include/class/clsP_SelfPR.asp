<%
'******************************************************************************
'名　称：clsP_SelfPR
'概　要：formで飛んできたP_SelfPRテーブル用のデータを持つためのクラス
'備　考：navi only
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_SelfPR
	Public StaffCode
	Public SelfPR
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_SelfPR クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If ExistsForm("CONF_SelfPR") = False Then Exit Sub

		IsData = True
		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_SelfPR") <> "" Then SelfPR = Request.Form("CONF_SelfPR")
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_SelfPR 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "sp_Reg_P_SelfPR" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(SelfPR) & "'" & vbCrLf
	End Function
End Class
%>
