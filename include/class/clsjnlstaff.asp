<%
'******************************************************************************
'概　要：ＬＩＳジャーナルの求職者配信予約内容クラス(構造体)
'関　数：■Private
'　　　：■Public
'備　考：
'更　新：2009/01/14 LIS K.Kokubo 作成
'　　　：2009/02/17 LIS K.Kokubo 追加 [HopeJobType][HopeYearlyIncome][EducateHistory]
'******************************************************************************
Class clsJNLStaff
	Public BranchCode
	Public ContentsType
	Public vol
	Public StaffCode
	Public Age
	Public Sex
	Public Address
	Public WorkStartDay
	Public CareerHistory
	Public CounselingView
	Public Skill
	Public RecentConditions
	Public Hope
	Public HopeJobType
	Public HopeYearlyIncome
	Public EducateHistory
	Public Certify

	'******************************************************************************
	'概　要：メンバ変数初期化
	'引　数：
	'備　考：
	'使　用：社内/mailservice/lisjournal/reserve.asp
	'更　新：2009/01/14 LIS K.Kokubo 作成
	'　　　：2009/02/17 LIS K.Kokubo 追加 [HopeJobType][HopeYearlyIncome][EducateHistory]
	'******************************************************************************
	Public Function Clear()
		BranchCode = Empty
		ContentsType = Empty
		vol = Empty
		StaffCode = Empty
		Age = Empty
		Sex = Empty
		Address = Empty
		WorkStartDay = Empty
		CareerHistory = Empty
		CounselingView = Empty
		Skill = Empty
		RecentConditions = Empty
		Hope = Empty
		HopeJobType = Empty
		HopeYearlyIncome = Empty
		EducateHistory = Empty
		Certify = Empty
	End Function

	'******************************************************************************
	'概　要：ＬＩＳジャーナルの勤務開始予定日変換
	'引　数：
	'備　考：
	'使　用：社内/mailservice/lisjournal/reserve.asp
	'更　新：2007/09/10 LIS K.Kokubo 作成
	'　　　：2008/04/24 LIS K.Kokubo 大阪支社だけ「入社日ご相談」文言
	'　　　：2009/01/14 LIS K.Kokubo クラス用に変更
	'******************************************************************************
	Public Function ChgWorkStartDay()
		On Error Resume Next
		Dim dWorkStartDay

		ChgWorkStartDay = ""

		If IsDate(WorkStartDay) = True Then
			dWorkStartDay = CDate(WorkStartDay)

			If DateDiff("d", dWorkStartDay, Date) < 0 Then
				ChgWorkStartDay = Year(dWorkStartDay) & "年" & Month(dWorkStartDay) & "月" & Day(dWorkStartDay) & "日"
			Else
				ChgWorkStartDay = "即日可能"
			End If
		ElseIf BranchCode = "OR" Then
			ChgWorkStartDay = "入社日ご相談"
		Else
			ChgWorkStartDay = "条件次第"
		End If
	End Function
End Class
%>
