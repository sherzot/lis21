<%
'******************************************************************************
'名　称：clsP_BankAccount
'概　要：formで飛んできたP_BankAccountテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_BankAccount
	Public StaffCode
	Public BankName
	Public BankNo
	Public BankBranchName
	Public BankBranchNo
	Public AccountNo
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_BankAccount クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_BankName") <> "" Then IsData = True: BankName = Request.Form("CONF_BankName")
		If Request.Form("CONF_BankNo") <> "" Then IsData = True: BankNo = Request.Form("CONF_BankNo")
		If Request.Form("CONF_BankBranchName") <> "" Then IsData = True: BankBranchName = Request.Form("CONF_BankBranchName")
		If Request.Form("CONF_BankBranchNo") <> "" Then IsData = True: BankBranchNo = Request.Form("CONF_BankBranchNo")
		If Request.Form("CONF_AccountNo") <> "" Then IsData = True: AccountNo = Request.Form("CONF_AccountNo")
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_BankAccount 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_BankAccount" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(BankName) & "'" & _
			",'" & ChkSQLStr(BankNo) & "'" & _
			",'" & ChkSQLStr(BankBranchName) & "'" & _
			",'" & ChkSQLStr(BankBranchNo) & "'" & _
			",'" & ChkSQLStr(AccountNo) & "'" & vbCrLf
	End Function
End Class
%>
