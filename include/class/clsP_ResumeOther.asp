<%
'******************************************************************************
'名　称：clsP_ResumeOther
'概　要：formで飛んできたP_ResumeOtherテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_ResumeOther
	Public StaffCode
	Public PrintFlag()
	Public Subject()
	Public WishMotive()
	Public CommuteTime()
	Public HopeColumn()
	Public IsData
	Public MaxIndex
	Public Err
	Public ErrStyle

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_ResumeOther クラスの初期化関数
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
			If ExistsForm("CONF_Subject" & idx) = False Then Exit Do

			If Request.Form("CONF_Subject" & idx) & Request.Form("CONF_WishMotive" & idx) & Request.Form("CONF_CommuteTime" & idx) & Request.Form("CONF_HopeColumn" & idx) <> "" Then
				MaxIndex = MaxIndex + 1

				ReDim Preserve Subject(MaxIndex) : Subject(MaxIndex) = Request.Form("CONF_Subject" & idx)
				ReDim Preserve WishMotive(MaxIndex) : WishMotive(MaxIndex) = Request.Form("CONF_WishMotive" & idx)
				ReDim Preserve CommuteTime(MaxIndex) : CommuteTime(MaxIndex) = Request.Form("CONF_CommuteTime" & idx)
				ReDim Preserve HopeColumn(MaxIndex) : HopeColumn(MaxIndex) = Request.Form("CONF_HopeColumn" & idx)

				If CommuteTime(MaxIndex) <> "" And IsNumber(CommuteTime(MaxIndex), 0, False) = False Then CommuteTime(MaxIndex) = "": Err = Err & "BranchCode" & vbCrLf
			End If
			idx = idx + 1
		Loop

		'値チェック
		Set ErrStyle = Server.CreateObject("scripting.dictionary")
		ErrStyle.CompareMode = 1

		For idx = 1 To MaxIndex
			'タイトル
			If Subject(idx) <> "" And ChkLen(Subject(idx), 200) = False Then
				Call DicAdd(ErrStyle, "CONF_Subject" & idx, "style=""background-color:#ffff00;""")
				Err = Err & "タイトルは半角１文字、全角２文字と数えて２００文字までです。<br>"
			End If

			'志望動機
			If WishMotive(idx) <> "" And ChkLen(WishMotive(idx), 2000) = False Then
				Call DicAdd(ErrStyle, "CONF_WishMotive" & idx, "style=""background-color:#ffff00;""")
				Err = Err & "志望動機は半角１文字、全角２文字と数えて２０００文字までです。<br>"
			End If

			'希望通勤時間
			If CommuteTime(idx) <> "" And IsNumber(CommuteTime(idx), 0, True) = False Then
				Call DicAdd(ErrStyle, "CONF_CommuteTime" & idx, "style=""background-color:#ffff00;""")
				Err = Err & "希望通勤時間は半角数字で入力して下さい。<br>"
			End If

			'本人希望
			If HopeColumn(idx) <> "" And ChkLen(HopeColumn(idx), 2000) = False Then
				Call DicAdd(ErrStyle, "CONF_HopeColumn" & idx, "style=""background-color:#ffff00;""")
				Err = Err & "本人希望は半角１文字、全角２文字と数えて２０００文字までです。<br>"
			End If
		Next

		'If Err = "" Then IsData = True
		IsData = True
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：clsP_ResumeOther 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_ResumeOther '" & ChkSQLStr(vStaffCode) & "'"
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_ResumeOther" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(Subject(idx)) & "'" & _
				",'" & ChkSQLStr(WishMotive(idx)) & "'" & _
				",'" & ChkSQLStr(CommuteTime(idx)) & "'" & _
				",'" & ChkSQLStr(HopeColumn(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
