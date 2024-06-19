<%
'******************************************************************************
'名　称：clsP_Note
'概　要：formで飛んできたP_Noteテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_Note
	Public StaffCode
	Public CategoryCode
	Public Code
	Public Note
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_Note クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize(vCode)
		IsData = False
		MaxIndex = -1

		StaffCode = Request.Form("CONF_StaffCode")
		CategoryCode = "Note"
		Code = vCode
		Note = Request.Form("CONF_Note_" Code, 1)
		If Note <> "" Then IsData = True
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_Note 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		GetRegSQL = "EXEC sp_Del_P_Note '" & ChkSQLStr(vStaffCode) & "', '" & ChkSQLStr(CategoryCode) & "', '" & ChkSQLStr(Code) & "'" & vbCrLf
		If IsData = False Then Exit Function
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_Note" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(CategoryCode) & "'" & _
			",'" & ChkSQLStr(Code) & "'" & _
			",'" & ChkSQLStr(Note) & "'" & vbCrLf
	End Function
End Class
%>
