<%
'******************************************************************************
'名　称：clsN_School
'概　要：formで飛んできたP_UserInfoテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/10/06
'更　新：
'******************************************************************************
Class clsN_School
	Public StaffCode
	Public SchoolCode
	Public OtherSchoolName
	Public Department
	Public Studyroom
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsN_Schoolクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("conf_staffcode") <> "" Then StaffCode = Request.Form("conf_staffcode")
		If Request.Form("conf_schoolcode") <> "" Then SchoolCode = Request.Form("conf_schoolcode")
		If Request.Form("conf_otherschoolname") <> "" Then OtherSchoolName = Request.Form("conf_otherschoolname")
		If Request.Form("conf_department") <> "" Then Department = Request.Form("conf_department")
		If Request.Form("conf_studyroom") <> "" Then Studyroom = Request.Form("conf_studyroom")

		If SchoolCode & OtherSchoolName <> "" Then IsData = True

		'値チェック
		Err = ""
		If StaffCode <> "" And IsMainCode(StaffCode) = False Then Err = Err & "StaffCode" & vbCrLf
		If SchoolCode <> "" And IsRE(SchoolCode, "^\d\d\d\d\d$", False) = False Then Err = Err & "SchoolCode" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_ UserInfo実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "up_Reg_N_School" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(SchoolCode) & "'" & _
			",'" & ChkSQLStr(OtherSchoolName) & "'" & _
			",'" & ChkSQLStr(Department) & "'" & _
			",'" & ChkSQLStr(Studyroom) & "'" & vbCrLf
	End Function
End Class
%>
