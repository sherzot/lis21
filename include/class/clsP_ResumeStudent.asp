<%
'******************************************************************************
'名　称：clsP_ResumeStudent
'概　要：formで飛んできたP_ResumeStudentテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/10/23
'更　新：
'******************************************************************************
Class clsP_ResumeStudent
	Public StaffCode
	Public Good
	Public Health
	Public Activity
	Public Specialty
	Public IsData
	Public MaxIndex
	Public Err
	Public ErrStyle

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_ResumeStudent クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/10/23
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_ResumeGood") <> "" Then IsData = True: Good = Request.Form("CONF_ResumeGood")
		If Request.Form("CONF_ResumeHealth") <> "" Then IsData = True: Health = Request.Form("CONF_ResumeHealth")
		If Request.Form("CONF_ResumeActivity") <> "" Then IsData = True: Activity = Request.Form("CONF_ResumeActivity")
		If Request.Form("CONF_ResumeSpecialty") <> "" Then IsData = True: Specialty = Request.Form("CONF_ResumeSpecialty")

		'値チェック
		Err = ""
		Set ErrStyle = Server.CreateObject("Scripting.Dictionary")
		ErrStyle.CompareMode = 1

		'得意分野・科目
		If Good <> "" And ChkLen(Good, 500) = False Then
			Call DicAdd(ErrStyle, "CONF_ResumeGood", "style=""background-color:#ffff00;""")
			Err = Err & "得意分野・科目は半角１文字、全角２文字と数えて５００文字までです。<br>"
		End If
		'健康状態
		If Health <> "" And ChkLen(Health, 500) = False Then
			Call DicAdd(ErrStyle, "CONF_ResumeHealth", "style=""background-color:#ffff00;""")
			Err = Err & "健康状態は半角１文字、全角２文字と数えて５００文字までです。<br>"
		End If
		'クラブ活動・文化活動
		If Activity <> "" And ChkLen(Activity, 500) = False Then
			Call DicAdd(ErrStyle, "CONF_ResumeActivity", "style=""background-color:#ffff00;""")
			Err = Err & "クラブ活動・文化活動は半角１文字、全角２文字と数えて５００文字までです。<br>"
		End If
		'趣味・特技
		If Specialty <> "" And ChkLen(Specialty, 500) = False Then
			Call DicAdd(ErrStyle, "CONF_ResumeSpecialty", "style=""background-color:#ffff00;""")
			Err = Err & "趣味・特技は半角１文字、全角２文字と数えて５００文字までです。<br>"
		End If

		If ErrStyle.Count > 0 Then IsData = False
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_ResumeStudent 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/10/23
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If Good & Health & Activity & Specialty = "" Then
			GetRegSQL = "sp_Del_P_ResumeStudent '" & ChkSQLStr(vStaffCode) & "'"
			Exit Function
		End If

		If IsData = False Then Exit Function
		GetRegSQL = "up_Reg_P_ResumeStudent" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(Good) & "'" & _
			",'" & ChkSQLStr(Health) & "'" & _
			",'" & ChkSQLStr(Activity) & "'" & _
			",'" & ChkSQLStr(Specialty) & "'" & vbCrLf
	End Function
End Class
%>
