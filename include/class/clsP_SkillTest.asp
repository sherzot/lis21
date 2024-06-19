<%
'******************************************************************************
'名　称：clsP_SkillTest
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_SkillTest
	Public StaffCode
	Public ExecuteDay1
	Public Kana_M1
	Public RomanChar_M1
	Public TenKeyTime1
	Public TenKeyCorrect1
	Public TenKeyStroke1
	Public ExecuteDay2
	Public Kana_M2
	Public RomanChar_M2
	Public TenKeyTime2
	Public TenKeyCorrect2
	Public TenKeyStroke2
	Public ExecuteDay3
	Public Kana_M3
	Public RomanChar_M3
	Public TenKeyTime3
	Public TenKeyCorrect3
	Public TenKeyStroke3
	Public ExecuteDay4
	Public Kana_M4
	Public RomanChar_M4
	Public TenKeyTime4
	Public TenKeyCorrect4
	Public TenKeyStroke4
	Public Behavior
	Public Durability
	Public Leader
	Public Challenge
	Public Sympathy
	Public Stability
	Public Originality
	Public Innovation
	Public Thinking
	Public Flexibility
	Public Sensitivity
	Public Carefulness
	Public DutySynthesis
	Public DutyRank
	Public GeneralSynthesis
	Public GeneralRank
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_SkillTest クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1
		ExecuteDay1 = ""
		ExecuteDay2 = ""
		ExecuteDay3 = ""
		ExecuteDay4 = ""

		ExecuteDay1 = Request.Form("CONF_ExecuteDayY1") & "/"
		If Len(Request.Form("CONF_ExecuteDayM1")) = 1 Then ExecuteDay1 = ExecuteDay1 & "0"
		ExecuteDay1 = ExecuteDay1 & Request.Form("CONF_ExecuteDayM1") & "/01"
		If IsDate(ExecuteDay1) = True Then
			ExecuteDay1 = Replace(ExecuteDay1, "/", "")
		Else
			ExecuteDay1 = ""
		End If

		ExecuteDay2 = Request.Form("CONF_ExecuteDayY2") & "/"
		If Len(Request.Form("CONF_ExecuteDayM2")) = 1 Then ExecuteDay2 = ExecuteDay2 & "0"
		ExecuteDay2 = ExecuteDay2 & Request.Form("CONF_ExecuteDayM2") & "/01"
		If IsDate(ExecuteDay2) = True Then
			ExecuteDay2 = Replace(ExecuteDay2, "/", "")
		Else
			ExecuteDay2 = ""
		End If

		ExecuteDay3 = Request.Form("CONF_ExecuteDayY3") & "/"
		If Len(Request.Form("CONF_ExecuteDayM3")) = 1 Then ExecuteDay3 = ExecuteDay3 & "0"
		ExecuteDay3 = ExecuteDay3 & Request.Form("CONF_ExecuteDayM3") & "/01"
		If IsDate(ExecuteDay3) = True Then
			ExecuteDay3 = Replace(ExecuteDay3, "/", "")
		Else
			ExecuteDay3 = ""
		End If

		ExecuteDay4 = Request.Form("CONF_ExecuteDayY4") & "/"
		If Len(Request.Form("CONF_ExecuteDayM4")) = 1 Then ExecuteDay4 = ExecuteDay4 & "0"
		ExecuteDay4 = ExecuteDay4 & Request.Form("CONF_ExecuteDayM4") & "/01"
		If IsDate(ExecuteDay4) = True Then
			ExecuteDay4 = Replace(ExecuteDay4, "/", "")
		Else
			ExecuteDay4 = ""
		End If

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_Kana_M1") <> "" Then IsData = True: Kana_M1 = Request.Form("CONF_Kana_M1")
		If Request.Form("CONF_RomanChar_M1") <> "" Then IsData = True: RomanChar_M1 = Request.Form("CONF_RomanChar_M1")
		If Request.Form("CONF_TenKeyTime1") <> "" Then IsData = True: TenKeyTime1 = Request.Form("CONF_TenKeyTime1")
		If Request.Form("CONF_TenKeyCorrect1") <> "" Then IsData = True: TenKeyCorrect1 = Request.Form("CONF_TenKeyCorrect1")
		If Request.Form("CONF_TenKeyStroke1") <> "" Then IsData = True: TenKeyStroke1 = Request.Form("CONF_TenKeyStroke1")
		If Request.Form("CONF_Kana_M2") <> "" Then IsData = True: Kana_M2 = Request.Form("CONF_Kana_M2")
		If Request.Form("CONF_RomanChar_M2") <> "" Then IsData = True: RomanChar_M2 = Request.Form("CONF_RomanChar_M2")
		If Request.Form("CONF_TenKeyTime2") <> "" Then IsData = True: TenKeyTime2 = Request.Form("CONF_TenKeyTime2")
		If Request.Form("CONF_TenKeyCorrect2") <> "" Then IsData = True: TenKeyCorrect2 = Request.Form("CONF_TenKeyCorrect2")
		If Request.Form("CONF_TenKeyStroke2") <> "" Then IsData = True: TenKeyStroke2 = Request.Form("CONF_TenKeyStroke2")
		If Request.Form("CONF_Kana_M3") <> "" Then IsData = True: Kana_M3 = Request.Form("CONF_Kana_M3")
		If Request.Form("CONF_RomanChar_M3") <> "" Then IsData = True: RomanChar_M3 = Request.Form("CONF_RomanChar_M3")
		If Request.Form("CONF_TenKeyTime3") <> "" Then IsData = True: TenKeyTime3 = Request.Form("CONF_TenKeyTime3")
		If Request.Form("CONF_TenKeyCorrect3") <> "" Then IsData = True: TenKeyCorrect3 = Request.Form("CONF_TenKeyCorrect3")
		If Request.Form("CONF_TenKeyStroke3") <> "" Then IsData = True: TenKeyStroke3 = Request.Form("CONF_TenKeyStroke3")
		If Request.Form("CONF_Kana_M4") <> "" Then IsData = True: Kana_M4 = Request.Form("CONF_Kana_M4")
		If Request.Form("CONF_RomanChar_M4") <> "" Then IsData = True: RomanChar_M4 = Request.Form("CONF_RomanChar_M4")
		If Request.Form("CONF_TenKeyTime4") <> "" Then IsData = True: TenKeyTime4 = Request.Form("CONF_TenKeyTime4")
		If Request.Form("CONF_TenKeyCorrect4") <> "" Then IsData = True: TenKeyCorrect4 = Request.Form("CONF_TenKeyCorrect4")
		If Request.Form("CONF_TenKeyStroke4") <> "" Then IsData = True: TenKeyStroke4 = Request.Form("CONF_TenKeyStroke4")
		If Request.Form("CONF_Behavior") <> "" Then IsData = True: Behavior = Request.Form("CONF_Behavior")
		If Request.Form("CONF_Durability") <> "" Then IsData = True: Durability = Request.Form("CONF_Durability")
		If Request.Form("CONF_Leader") <> "" Then IsData = True: Leader = Request.Form("CONF_Leader")
		If Request.Form("CONF_Challenge") <> "" Then IsData = True: Challenge = Request.Form("CONF_Challenge")
		If Request.Form("CONF_Sympathy") <> "" Then IsData = True: Sympathy = Request.Form("CONF_Sympathy")
		If Request.Form("CONF_Stability") <> "" Then IsData = True: Stability = Request.Form("CONF_Stability")
		If Request.Form("CONF_Originality") <> "" Then IsData = True: Originality = Request.Form("CONF_Originality")
		If Request.Form("CONF_Innovation") <> "" Then IsData = True: Innovation = Request.Form("CONF_Innovation")
		If Request.Form("CONF_Thinking") <> "" Then IsData = True: Thinking = Request.Form("CONF_Thinking")
		If Request.Form("CONF_Flexibility") <> "" Then IsData = True: Flexibility = Request.Form("CONF_Flexibility")
		If Request.Form("CONF_Sensitivity") <> "" Then IsData = True: Sensitivity = Request.Form("CONF_Sensitivity")
		If Request.Form("CONF_Carefulness") <> "" Then IsData = True: Carefulness = Request.Form("CONF_Carefulness")
		If Request.Form("CONF_DutySynthesis") <> "" Then IsData = True: DutySynthesis = Request.Form("CONF_DutySynthesis")
		If Request.Form("CONF_DutyRank") <> "" Then IsData = True: DutyRank = Request.Form("CONF_DutyRank")
		If Request.Form("CONF_GeneralSynthesis") <> "" Then IsData = True: GeneralSynthesis = Request.Form("CONF_GeneralSynthesis")
		If Request.Form("CONF_GeneralRank") <> "" Then IsData = True: GeneralRank = Request.Form("CONF_GeneralRank")

		Err = ""
		If ExecuteDay1 <> "" And IsDay(ExecuteDay1) = False Then Err = Err & "ExecuteDay1" & vbCrLf
		If ExecuteDay2 <> "" And IsDay(ExecuteDay2) = False Then Err = Err & "ExecuteDay2" & vbCrLf
		If ExecuteDay3 <> "" And IsDay(ExecuteDay3) = False Then Err = Err & "ExecuteDay3" & vbCrLf
		If ExecuteDay4 <> "" And IsDay(ExecuteDay4) = False Then Err = Err & "ExecuteDay4" & vbCrLf
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_SkillTest 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function
		GetRegSQL = "sp_Reg_P_SkillTest" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(ExecuteDay1) & "'" & _
			",'" & ChkSQLStr(Kana_M1) & "'" & _
			",'" & ChkSQLStr(RomanChar_M1) & "'" & _
			",'" & ChkSQLStr(TenKeyTime1) & "'" & _
			",'" & ChkSQLStr(TenKeyCorrect1) & "'" & _
			",'" & ChkSQLStr(TenKeyStroke1) & "'" & _
			",'" & ChkSQLStr(ExecuteDay2) & "'" & _
			",'" & ChkSQLStr(Kana_M2) & "'" & _
			",'" & ChkSQLStr(RomanChar_M2) & "'" & _
			",'" & ChkSQLStr(TenKeyTime2) & "'" & _
			",'" & ChkSQLStr(TenKeyCorrect2) & "'" & _
			",'" & ChkSQLStr(TenKeyStroke2) & "'" & _
			",'" & ChkSQLStr(ExecuteDay3) & "'" & _
			",'" & ChkSQLStr(Kana_M3) & "'" & _
			",'" & ChkSQLStr(RomanChar_M3) & "'" & _
			",'" & ChkSQLStr(TenKeyTime3) & "'" & _
			",'" & ChkSQLStr(TenKeyCorrect3) & "'" & _
			",'" & ChkSQLStr(TenKeyStroke3) & "'" & _
			",'" & ChkSQLStr(ExecuteDay4) & "'" & _
			",'" & ChkSQLStr(Kana_M4) & "'" & _
			",'" & ChkSQLStr(RomanChar_M4) & "'" & _
			",'" & ChkSQLStr(TenKeyTime4) & "'" & _
			",'" & ChkSQLStr(TenKeyCorrect4) & "'" & _
			",'" & ChkSQLStr(TenKeyStroke4) & "'" & _
			",'" & ChkSQLStr(Behavior) & "'" & _
			",'" & ChkSQLStr(Durability) & "'" & _
			",'" & ChkSQLStr(Leader) & "'" & _
			",'" & ChkSQLStr(Challenge) & "'" & _
			",'" & ChkSQLStr(Sympathy) & "'" & _
			",'" & ChkSQLStr(Stability) & "'" & _
			",'" & ChkSQLStr(Originality) & "'" & _
			",'" & ChkSQLStr(Innovation) & "'" & _
			",'" & ChkSQLStr(Thinking) & "'" & _
			",'" & ChkSQLStr(Flexibility) & "'" & _
			",'" & ChkSQLStr(Sensitivity) & "'" & _
			",'" & ChkSQLStr(Carefulness) & "'" & _
			",'" & ChkSQLStr(DutySynthesis) & "'" & _
			",'" & ChkSQLStr(DutyRank) & "'" & _
			",'" & ChkSQLStr(GeneralSynthesis) & "'" & _
			",'" & ChkSQLStr(GeneralRank) & "'" & vbCrLf
	End Function
End Class
%>
