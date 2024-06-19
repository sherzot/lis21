<%
'******************************************************************************
'概　要：検索条件を保持するクラス
'関　数：■Public
'　　　：
'　　　：■Private
'　　　：
'備　考：■■■ 詳細検索用パラメータ （アドホックなＳＱＬ生成）
'更　新：2009/09/09 LIS K.Kokubo 作成
'******************************************************************************
Class clsSearchOrderAdvice
	Public StaffCode
	'検索条件
	Public HopeJobTypeCode		'希望職種CSV
	Public HopePrefectureCode	'希望勤務地等道府県CSV
	Public HopeWorkingTypeCode	'希望勤務形態CSV
	Public HopeIndustryTypeCode	'希望業種CSV
	Public YearlyIncomeMin		'年収下限
	Public YearlyIncomeMax		'年収上限
	Public MonthlyIncomeMin		'月給下限
	Public MonthlyIncomeMax		'月給上限
	Public DailyIncomeMin		'日給下限
	Public DailyIncomeMax		'日給上限
	Public HourlyIncomeMin		'時給下限
	Public HourlyIncomeMax		'時給上限
	Public WorkStartTime		'勤務開始時刻 HHMM
	Public WorkEndTime			'勤務終了時刻 HHMM
	Public WeeklyHolidayTypeCode'週休種類
	Public LicenseCount			'資格数
	Public LicenseGroupCode		'資格大分類配列
	Public LicenseCategoryCode	'資格中分類配列
	Public LicenseCode			'資格小分類配列
	Public OSCode				'OSCSV
	Public APCode				'アプリケーションCSV
	Public DLCode				'開発言語CSV
	Public DBCode				'データベースCSV
	'
	Public HopeJobTypeFlag
	Public HopeWorkingPlaceFlag
	Public HopeWorkingTypeFlag
	Public HopeIndustryTypeFlag
	Public SalaryFlag
	Public WorkTimeFlag
	Public HolidayFlag
	Public LicenseFlag
	Public OSFlag
	Public APFlag
	Public DLFlag
	Public DBFlag
	'
	Public HopeJobTypeName
	Public HopePrefectureName
	Public HopeWorkingTypeName
	Public HopeIndustryTypeName
	Public WeeklyHolidayTypeName
	Public LicenseName
	Public OSName
	Public APName
	Public DLName
	Public DBName

	'******************************************************************************
	'概　要：コンストラクタ
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Private Sub Class_Initialize()
		LicenseCount = 0

		'パラメータから検索条件を取得
		Call ReadParam()
	End Sub

	'******************************************************************************
	'概　要：GETデータの読み込み
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Sub ReadParam()
		Dim idx

		If GetForm("hjtf",2) <> "" Then HopeJobTypeFlag = GetForm("hjtf",2)
		If GetForm("hwpf",2) <> "" Then HopeWorkingPlaceFlag = GetForm("hwpf",2)
		If GetForm("hwtf",2) <> "" Then HopeWorkingTypeFlag = GetForm("hwtf",2)
		If GetForm("hitf",2) <> "" Then HopeIndustryTypeFlag = GetForm("hitf",2)
		If GetForm("sf",2) <> "" Then SalaryFlag = GetForm("sf",2)
		If GetForm("wtf",2) <> "" Then WorkTimeFlag = GetForm("wtf",2)
		If GetForm("hf",2) <> "" Then HolidayFlag = GetForm("hf",2)
		If GetForm("lf",2) <> "" Then LicenseFlag = GetForm("lf",2)
		If GetForm("osf",2) <> "" Then OSFlag = GetForm("osf",2)
		If GetForm("apf",2) <> "" Then APFlag = GetForm("apf",2)
		If GetForm("dlf",2) <> "" Then DLFlag = GetForm("dlf",2)
		If GetForm("dbf",2) <> "" Then DBFlag = GetForm("dbf",2)

		'データ整合性チェック
		Call ChkData()

		'コード対応名称取得
		Call SetData()
	End Sub

	'******************************************************************************
	'概　要：コードに対応した名称を取得する
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Sub SetData()
		Call SetHopeJobType()
		Call SetHopeWorkingPlace()
		Call SetHopeWorkingType()
		Call SetHopeIndustryType()
		Call SetStaffData()
		Call SetLicense()
		Call SetSkill()
	End Sub

	'******************************************************************************
	'概　要：データの整合性をチェック
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Private Sub ChkData()
	End Sub

	'******************************************************************************
	'概　要：お仕事詳細検索ページへ渡すGETパラメータを生成して取得。
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function GetSearchParam()
		Dim idx
		Dim sParam

		sParam = ""

		If HopeJobTypeFlag <> "" Then sParam = sParam & "&hjtf=" & HopeJobTypeFlag
		If HopeWorkingPlaceFlag <> "" Then sParam = sParam & "&hwpf=" & HopeWorkingPlaceFlag
		If HopeWorkingTypeFlag <> "" Then sParam = sParam & "&hwtf=" & HopeWorkingTypeFlag
		If HopeIndustryTypeFlag <> "" Then sParam = sParam & "&hitf=" & HopeIndustryTypeFlag
		If SalaryFlag <> "" Then sParam = sParam & "&sf=" & SalaryFlag
		If WorkTimeFlag <> "" Then sParam = sParam & "&wtf=" & WorkTimeFlag
		If HolidayFlag <> "" Then sParam = sParam & "&hf=" & HolidayFlag
		If LicenseFlag <> "" Then sParam = sParam & "&lf=" & LicenseFlag
		If OSFlag <> "" Then sParam = sParam & "&osf=" & OSFlag
		If APFlag <> "" Then sParam = sParam & "&apf=" & APFlag
		If DLFlag <> "" Then sParam = sParam & "&dlf=" & DLFlag
		If DBFlag <> "" Then sParam = sParam & "&dbf=" & DBFlag

		If sParam <> "" Then
			'頭の&を？に変換
			sParam = "?" & Mid(GetSearchParam, 2)

			'ＩＥの仕様はパラメータの上限が２０４８バイト
			sParam = Left(sParam, 2048)
		End If

		GetSearchParam = sParam
	End Function

	'******************************************************************************
	'概　要：求人票詳細検索ＳＱＬを取得
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Function sqlSearchOrderAdvice()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sCount

		Dim sSQL
		Dim sSQL2
		Dim tmp1
		Dim tmp2
		Dim tmp3
		Dim iPrmNo
		Dim iPrmNo2

		Dim aValue
		Dim idx

		sDeclare = ""
		sParams = ""
		sJoin = ""
		sCount = ""

		'データ整合性チェック
		Call ChkData()

		'<希望職種>
		tmp1 = ""
		iPrmNo = 1
		If HopeJobTypeCode <> "" Then
			aValue = Split(HopeJobTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHopeJobTypeCode" & iPrmNo & " VARCHAR(7)"
					sParams = sParams & ",@vHopeJobTypeCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vHopeJobTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If HopeJobTypeFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_JobType AS A WHERE A.JobTypeCode IN (" & tmp1 & ")) AS CJT ON VWOC.OrderCode = CJT.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT 'HopeJobType', COUNT(*) FROM BASE AS A WHERE EXISTS(SELECT * FROM C_JobType AS B WHERE A.OrderCode = B.OrderCode AND B.JobTypeCode IN (" & tmp1 & "))" & vbCrLf
			End If
		End If
		'</希望職種>

		'<希望勤務地>
		tmp1 = ""
		iPrmNo = 1
		If HopePrefectureCode <> "" Then
			aValue = Split(HopePrefectureCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHopePrefectureCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vHopePrefectureCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vHopePrefectureCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If HopeWorkingPlaceFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_WorkingPlace AS A WHERE A.PrefectureCode IN (" & tmp1 & ")) AS CWP ON VWOC.OrderCode = CWP.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT 'HopeWorkingPlace', COUNT(*) FROM BASE AS A WHERE EXISTS(SELECT * FROM C_WorkingPlace AS B WHERE A.OrderCode = B.OrderCode AND B.PrefectureCode IN (" & tmp1 & "))" & vbCrLf
			End If
		End If
		'</希望職種>

		'<希望勤務形態>
		tmp1 = ""
		iPrmNo = 1
		If HopeWorkingTypeCode <> "" Then
			aValue = Split(HopeWorkingTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHopeWorkingTypeCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vHopeWorkingTypeCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vHopeWorkingTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If HopeWorkingTypeFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_WorkingType AS A WHERE A.WorkingTypeCode IN (" & tmp1 & ")) AS CWT ON VWOC.OrderCode = CWT.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT 'HopeWorkingType', COUNT(*) FROM BASE AS A WHERE EXISTS(SELECT * FROM C_WorkingType AS B WHERE A.OrderCode = B.OrderCode AND B.WorkingTypeCode IN (" & tmp1 & "))" & vbCrLf
			End If
		End If
		'</希望勤務形態>

		'<希望業種>
		tmp1 = ""
		iPrmNo = 1
		If HopeIndustryTypeCode <> "" Then
			aValue = Split(HopeIndustryTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHopeIndustryTypeCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vHopeIndustryTypeCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vHopeIndustryTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If HopeIndustryTypeFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_Info AS A WHERE EXISTS(SELECT * FROM CompanyInfo AS B WHERE A.CompanyCode = B.CompanyCode AND B.IndustryType IN (" & tmp1 & "))) AS CIT ON VWOC.OrderCode = CIT.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT 'HopeIndustryType', COUNT(*) FROM BASE AS A WHERE EXISTS(SELECT * FROM C_Info AS B WHERE A.OrderCode = B.OrderCode AND EXISTS(SELECT * FROM CompanyInfo AS C WHERE B.CompanyCode = C.CompanyCode AND C.IndustryType IN (" & tmp1 & ")))" & vbCrLf
			End If
		End If
		'</希望業種>

		'<給与>
		tmp1 = ""
		If YearlyIncomeMin & YearlyIncomeMax & MonthlyIncomeMin & MonthlyIncomeMax & DailyIncomeMin & DailyIncomeMax & HourlyIncomeMin & HourlyIncomeMax <> "" Then
			'<年収>
			If YearlyIncomeMin & YearlyIncomeMax <> "" Then
				If YearlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vYearlyIncomeMin INT"
					sParams = sParams & ",@vYearlyIncomeMin = " & YearlyIncomeMin
				End If

				If YearlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vYearlyIncomeMax INT"
					sParams = sParams & ",@vYearlyIncomeMax = " & YearlyIncomeMax
				End If

				If tmp1 <> "" Then tmp1 = tmp1 & "OR "
				If YearlyIncomeMin <> "" And YearlyIncomeMax <> "" Then
					'年収下限,上限両方の入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND (A.YearlyIncomeMin BETWEEN @vYearlyIncomeMin AND @vYearlyIncomeMax)) OR (COALESCE(A.YearlyIncomeMax, 0) > 0 AND (A.YearlyIncomeMax BETWEEN @vYearlyIncomeMin AND @vYearlyIncomeMax))) "
				ElseIf YearlyIncomeMin <> "" Then
					'年収下限のみ入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND A.YearlyIncomeMin >= @vYearlyIncomeMin) OR (COALESCE(A.YearlyIncomeMax, 0) > 0 AND A.YearlyIncomeMax >= @vYearlyIncomeMin)) "
				ElseIf YearlyIncomeMax <> "" Then
					'年収上限のみ入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND A.YearlyIncomeMin <= @vYearlyIncomeMax) OR (COALESCE(A.YearlyIncomeMin, 0) = 0 AND COALESCE(A.YearlyIncomeMax, 0) > 0 AND A.YearlyIncomeMax <= @vYearlyIncomeMax)) "
				End If
			End If
			'</年収>

			'<月給>
			If MonthlyIncomeMin & MonthlyIncomeMax <> "" Then
				If MonthlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vMonthlyIncomeMin INT"
					sParams = sParams & ",@vMonthlyIncomeMin = " & MonthlyIncomeMin
				End If

				If MonthlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vMonthlyIncomeMax INT"
					sParams = sParams & ",@vMonthlyIncomeMax = " & MonthlyIncomeMax
				End If

				If tmp1 <> "" Then tmp1 = tmp1 & "OR "
				If MonthlyIncomeMin <> "" And MonthlyIncomeMax <> "" Then
					'月給下限,上限両方の入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND (A.MonthlyIncomeMin BETWEEN @vMonthlyIncomeMin AND @vMonthlyIncomeMax)) OR (COALESCE(A.MonthlyIncomeMax, 0) > 0 AND (A.MonthlyIncomeMax BETWEEN @vMonthlyIncomeMin AND @vMonthlyIncomeMax))) "
				ElseIf MonthlyIncomeMin <> "" Then
					'月給下限のみ入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND A.MonthlyIncomeMin >= @vMonthlyIncomeMin) OR (COALESCE(A.MonthlyIncomeMax, 0) > 0 AND A.MonthlyIncomeMax >= @vMonthlyIncomeMin)) "
				ElseIf MonthlyIncomeMax <> "" Then
					'月給上限のみ入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND A.MonthlyIncomeMin <= @vMonthlyIncomeMax) OR (COALESCE(A.MonthlyIncomeMin, 0) = 0 AND COALESCE(A.MonthlyIncomeMax, 0) > 0 AND A.MonthlyIncomeMax <= @vMonthlyIncomeMax)) "
				End If
			End If
			'</月給>

			'<日給>
			If DailyIncomeMin & DailyIncomeMax <> "" Then
				If DailyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDailyIncomeMin INT"
					sParams = sParams & ",@vDailyIncomeMin = " & DailyIncomeMin
				End If

				If DailyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDailyIncomeMax INT"
					sParams = sParams & ",@vDailyIncomeMax = " & DailyIncomeMax
				End If

				If tmp1 <> "" Then tmp1 = tmp1 & "OR "
				If DailyIncomeMin <> "" And DailyIncomeMax <> "" Then
					'日給下限,上限両方の入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND (A.DailyIncomeMin BETWEEN @vDailyIncomeMin AND @vDailyIncomeMax)) OR (COALESCE(A.DailyIncomeMax, 0) > 0 AND (A.DailyIncomeMax BETWEEN @vDailyIncomeMin AND @vDailyIncomeMax))) "
				ElseIf DailyIncomeMin <> "" Then
					'日給下限のみ入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND A.DailyIncomeMin >= @vDailyIncomeMin) OR (COALESCE(A.DailyIncomeMax, 0) > 0 AND A.DailyIncomeMax >= @vDailyIncomeMin)) "
				ElseIf DailyIncomeMax <> "" Then
					'日給上限のみ入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND A.DailyIncomeMin <= @vDailyIncomeMax) OR (COALESCE(A.DailyIncomeMin, 0) = 0 AND COALESCE(A.DailyIncomeMax, 0) > 0 AND A.DailyIncomeMax <= @vDailyIncomeMax)) "
				End If
			End If
			'</日給>

			'<時給>
			If HourlyIncomeMin & HourlyIncomeMax <> "" Then
				If HourlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHourlyIncomeMin INT"
					sParams = sParams & ",@vHourlyIncomeMin = " & HourlyIncomeMin
				End If

				If HourlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHourlyIncomeMax INT"
					sParams = sParams & ",@vHourlyIncomeMax = " & HourlyIncomeMax
				End If

				If tmp1 <> "" Then tmp1 = tmp1 & "OR "
				If HourlyIncomeMin <> "" And HourlyIncomeMax <> "" Then
					'時給下限,上限両方の入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND (A.HourlyIncomeMin BETWEEN @vHourlyIncomeMin AND @vHourlyIncomeMax)) OR (COALESCE(A.HourlyIncomeMax, 0) > 0 AND (A.HourlyIncomeMax BETWEEN @vHourlyIncomeMin AND @vHourlyIncomeMax))) "
				ElseIf HourlyIncomeMin <> "" Then
					'時給下限のみ入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND A.HourlyIncomeMin >= @vHourlyIncomeMin) OR (COALESCE(A.HourlyIncomeMax, 0) > 0 AND A.HourlyIncomeMax >= @vHourlyIncomeMin)) "
				ElseIf HourlyIncomeMax <> "" Then
					'時給上限のみ入力がある場合
					tmp1 = tmp1 & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND A.HourlyIncomeMin <= @vHourlyIncomeMax) OR (COALESCE(A.HourlyIncomeMin, 0) = 0 AND COALESCE(A.HourlyIncomeMax, 0) > 0 AND A.HourlyIncomeMax <= @vHourlyIncomeMax)) "
				End If
			End If
			'</時給>

			If SalaryFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_Info AS A WHERE " & Trim(tmp1) & ") AS CSLY ON VWOC.OrderCode = CSLY.OrderCode " & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT 'Salary', COUNT(*) FROM BASE AS BASE WHERE EXISTS(SELECT * FROM C_Info AS A WHERE BASE.OrderCode = A.OrderCode AND (" & Trim(tmp1) & "))" & vbCrLf
			End If
		End If
		'<給与>

		'<勤務時間>
		tmp1 = ""
		tmp2 = ""
		If WorkStartTime & WorkEndTime <> "" Then
			If WorkStartTime <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkStartTime VARCHAR(4) "
				sParams = sParams & ",@vWorkStartTime = N'" & WorkStartTime & "'"

				If tmp1 <> "" Then tmp1 = tmp1 & "AND "
				tmp1 = tmp1 & "A.WorkStartTime >= @vWorkStartTime "
			End If

			If WorkEndTime <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkEndTime VARCHAR(4) "
				sParams = sParams & ",@vWorkEndTime = N'" & WorkEndTime & "'"

				If tmp1 <> "" Then tmp1 = tmp1 & "AND "
				tmp1 = tmp1 & "A.WorkEndTime <= @vWorkEndTime + @vWorkEndTime "
			End If

			If WorkStartTime <> "" And WorkEndTime <> "" Then
				If WorkStartTime < WorkEndTime Then
					'勤務開始時間 < 勤務終了時間の場合、夜間の業務時間を除くようにする
					tmp2 = "AND A.WorkStartTime < A.WorkEndTime "
				End If
			End If

			If WorkTimeFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_WorkingCondition AS A WHERE " & tmp1 & RTrim(tmp2) & ") AS CTM ON VWOC.OrderCode = CTM.OrderCode " & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT 'WorkTime', COUNT(*) FROM BASE AS BASE WHERE EXISTS(SELECT * FROM C_WorkingCondition AS A WHERE BASE.OrderCode = A.OrderCode AND " & tmp1 & RTrim(tmp2) & ")" & vbCrLf
			End If
		End If
		'</勤務時間>

		'<週休種類>
		tmp1 = ""
		If WeeklyHolidayTypeCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vWeeklyHolidayTypeCode VARCHAR(3) "
			sParams = sParams & ",@vWeeklyHolidayTypeCode = N'" & WeeklyHolidayTypeCode & "'"

			tmp1 = tmp1 & "A.WeeklyHolidayType = @vWeeklyHolidayTypeCode "

			If HolidayFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_Info AS A WHERE " & tmp1 & ") AS CWHT ON VWOC.OrderCode = CWHT.OrderCode " & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT 'Holiday', COUNT(*) FROM BASE AS BASE WHERE EXISTS(SELECT * FROM C_Info AS A WHERE BASE.OrderCode = A.OrderCode AND " & tmp1 & ")" & vbCrLf
			End If
		End If
		'</週休種類>

		'<資格>
		tmp1 = ""
		tmp2 = ""
		iPrmNo = 1
		If LicenseCount > 0 Then
			For idx = 0 To LicenseCount - 1
				tmp1 = ""
				If LicenseGroupCode(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLicenseGroupCode" & iPrmNo & " VARCHAR(2)"
					sParams = sParams & ",@vLicenseGroupCode" & iPrmNo & " = N'" & LicenseGroupCode(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & "AND "
					tmp1 = tmp1 & "A.GroupCode = @vLicenseGroupCode" & iPrmNo & " "
				End If

				If LicenseGroupCode(idx) <> "" And LicenseCategoryCode(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLicenseCategoryCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vLicenseCategoryCode" & iPrmNo & " = N'" & LicenseCategoryCode(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & "AND "
					tmp1 = tmp1 & "A.CategoryCode = @vLicenseCategoryCode" & iPrmNo & " "
				End If

				If LicenseGroupCode(idx) <> "" And LicenseCategoryCode(idx) <> "" And LicenseCode(idx) Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLicenseCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vLicenseCode" & iPrmNo & " = N'" & LicenseCode(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & "AND "
					tmp1 = tmp1 & "A.Code = @vLicenseCode" & iPrmNo & " "
				End If

				If LicenseGroupCode(idx) <> "" Then iPrmNo = iPrmNo + 1

				If tmp2 <> "" Then tmp2 = tmp2 & "OR "
				tmp2 = tmp2 & "(" & RTrim(tmp1) & ")"
			Next

			If LicenseFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_License AS A WHERE (" & tmp2 & ")) AS CL ON VWOC.OrderCode = CL.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT 'License', COUNT(*) FROM BASE AS BASE WHERE EXISTS(SELECT * FROM C_License AS A WHERE BASE.OrderCode = A.OrderCode AND (" & tmp2 & "))" & vbCrLf
			End If
		End If
		'</資格>

		iPrmNo2 = 1
		'<ＯＳ>
		tmp1 = ""
		iPrmNo = 1
		If OSCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategory" & iPrmNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategory" & iPrmNo2 & " = N'OS'"

			aValue = Split(OSCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vOSCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vOSCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vOSCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If OSFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WHERE A.CategoryCode = @vSkillCategory" & iPrmNo2 & " AND A.Code IN (" & tmp1 & ")) AS COS ON VWOC.OrderCode = COS.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT @vSkillCategory" & iPrmNo2 & ", COUNT(*) FROM BASE AS BASE WHERE EXISTS(SELECT * FROM C_Skill AS A WHERE BASE.OrderCode = A.OrderCode AND A.CategoryCode = @vSkillCategory" & iPrmNo2 & " AND A.Code IN (" & tmp1 & "))" & vbCrLf
			End If

			iPrmNo2 = iPrmNo2 + 1
		End If
		'</ＯＳ>

		'<アプリケーション>
		tmp1 = ""
		iPrmNo = 1
		If APCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategory" & iPrmNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategory" & iPrmNo2 & " = N'Application'"

			aValue = Split(APCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vAPCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vAPCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vAPCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If APFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WHERE A.CategoryCode = @vSkillCategory" & iPrmNo2 & " AND A.Code IN (" & tmp1 & ")) AS CAP ON VWOC.OrderCode = CAP.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT @vSkillCategory" & iPrmNo2 & ", COUNT(*) FROM BASE AS BASE WHERE EXISTS(SELECT * FROM C_Skill AS A WHERE BASE.OrderCode = A.OrderCode AND A.CategoryCode = @vSkillCategory" & iPrmNo2 & " AND A.Code IN (" & tmp1 & "))" & vbCrLf
			End If

			iPrmNo2 = iPrmNo2 + 1
		End If
		'</アプリケーション>

		'<開発言語>
		tmp1 = ""
		iPrmNo = 1
		If DLCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategory" & iPrmNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategory" & iPrmNo2 & " = N'DevelopmentLanguage'"

			aValue = Split(DLCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDLCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vDLCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vDLCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If DLFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WHERE A.CategoryCode = @vSkillCategory" & iPrmNo2 & " AND A.Code IN (" & tmp1 & ")) AS CDL ON VWOC.OrderCode = CDL.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT @vSkillCategory" & iPrmNo2 & ", COUNT(*) FROM BASE AS BASE WHERE EXISTS(SELECT * FROM C_Skill AS A WHERE BASE.OrderCode = A.OrderCode AND A.CategoryCode = @vSkillCategory" & iPrmNo2 & " AND A.Code IN (" & tmp1 & "))" & vbCrLf
			End If

			iPrmNo2 = iPrmNo2 + 1
		End If
		'</開発言語>

		'<データベース>
		tmp1 = ""
		iPrmNo = 1
		If DBCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategory" & iPrmNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategory" & iPrmNo2 & " = N'Database'"

			aValue = Split(DBCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDBCode" & iPrmNo & " VARCHAR(3)"
					sParams = sParams & ",@vDBCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vDBCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If DBFlag = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WHERE A.CategoryCode = @vSkillCategory" & iPrmNo2 & " AND A.Code IN (" & tmp1 & ")) AS CDB ON VWOC.OrderCode = CDB.OrderCode" & vbCrLf
			Else
				sCount = sCount & "UNION "
				sCount = sCount & "SELECT @vSkillCategory" & iPrmNo2 & ", COUNT(*) FROM BASE AS BASE WHERE EXISTS(SELECT * FROM C_Skill AS A WHERE BASE.OrderCode = A.OrderCode AND A.CategoryCode = @vSkillCategory" & iPrmNo2 & " AND A.Code IN (" & tmp1 & "))" & vbCrLf
			End If

			iPrmNo2 = iPrmNo2 + 1
		End If
		'</データベース>

		sSQL = ""
		sSQL = sSQL & "WITH BASE(OrderCode) AS (" & vbCrLf
		sSQL = sSQL & "SELECT VWOC.OrderCode" & vbCrLf
		sSQL = sSQL & "FROM vw_OrderCode AS VWOC" & vbCrLf
		sSQL = sSQL & sJoin
		sSQL = sSQL & ")" & vbCrLf
		sSQL = sSQL & "SELECT 'Result' AS CountType,(SELECT COUNT(*) FROM BASE) AS OrderCnt" & vbCrLf
		sSQL = sSQL & sCount

		sSQL2 = ""
		sSQL2 = sSQL2 & "/*ナビ・求人票詳細検索*/" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlSearchOrderAdvice = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'概　要：
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function SetHopeJobType()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr
		Dim tmp1
		Dim tmp2

		tmp1 = ""
		tmp2 = ""

		sSQL = "EXEC up_LstP_HopeJobType '" & StaffCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
		Do While GetRSState(oRS) = True
			If tmp1 <> "" Then tmp1 = tmp1 & ","
			tmp1 = tmp1 & oRS.Collect("JobTypeCode")

			If tmp2 <> "" Then tmp2 = tmp2 & ","
			tmp2 = tmp2 & oRS.Collect("JobTypeDetail")

			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		HopeJobTypeCode = tmp1
		HopeJobTypeName = tmp2
	End Function

	'******************************************************************************
	'概　要：
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function SetHopeWorkingPlace()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr
		Dim tmp1
		Dim tmp2

		tmp1 = ""
		tmp2 = ""

		sSQL = "EXEC up_LstP_HopeWorkingPlace '" & StaffCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
		Do While GetRSState(oRS) = True
			If tmp1 <> "" Then tmp1 = tmp1 & ","
			tmp1 = tmp1 & oRS.Collect("PrefectureCode")

			If tmp2 <> "" Then tmp2 = tmp2 & ","
			tmp2 = tmp2 & oRS.Collect("PrefectureName")

			oRS.MoveNext
		Loop

		HopePrefectureCode = tmp1
		HopePrefectureName = tmp2
	End Function

	'******************************************************************************
	'概　要：
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function SetHopeWorkingType()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr
		Dim tmp1
		Dim tmp2

		tmp1 = ""
		tmp2 = ""

		sSQL = "EXEC up_LstP_HopeWorkingType '" & StaffCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
		Do While GetRSState(oRS) = True
			If tmp1 <> "" Then tmp1 = tmp1 & ","
			tmp1 = tmp1 & oRS.Collect("WorkingTypeCode")

			If tmp2 <> "" Then tmp2 = tmp2 & ","
			tmp2 = tmp2 & oRS.Collect("WorkingTypeName")

			oRS.MoveNext
		Loop

		HopeWorkingTypeCode = tmp1
		HopeWorkingTypeName = tmp2
	End Function

	'******************************************************************************
	'概　要：
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function SetHopeIndustryType()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr
		Dim tmp1
		Dim tmp2

		tmp1 = ""
		tmp2 = ""

		sSQL = "EXEC up_LstP_HopeIndustryType '" & StaffCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
		Do While GetRSState(oRS) = True
			If tmp1 <> "" Then tmp1 = tmp1 & ","
			tmp1 = tmp1 & oRS.Collect("IndustryTypeCode")

			If tmp2 <> "" Then tmp2 = tmp2 & ","
			tmp2 = tmp2 & oRS.Collect("IndustryTypeName")

			oRS.MoveNext
		Loop

		HopeIndustryTypeCode = tmp1
		HopeIndustryTypeName = tmp2
	End Function

	'******************************************************************************
	'概　要：
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function SetStaffData()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr

		sSQL = "EXEC up_DtlStaff '" & StaffCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
		If GetRSState(oRS) = True Then
			YearlyIncomeMin = ChkStr(oRS.Collect("YearlyIncomeMin"))
			YearlyIncomeMax = ChkStr(oRS.Collect("YearlyIncomeMax"))
			MonthlyIncomeMin = ChkStr(oRS.Collect("MonthlyIncomeMin"))
			MonthlyIncomeMax = ChkStr(oRS.Collect("MonthlyIncomeMax"))
			DailyIncomeMin = ChkStr(oRS.Collect("DailyIncomeMin"))
			DailyIncomeMax = ChkStr(oRS.Collect("DailyIncomeMax"))
			HourlyIncomeMin = ChkStr(oRS.Collect("HourlyIncomeMin"))
			HourlyIncomeMax = ChkStr(oRS.Collect("HourlyIncomeMax"))
			WorkStartTime = ChkStr(oRS.Collect("WorkStartTime"))
			WorkEndTime = ChkStr(oRS.Collect("WorkEndTime"))
			WeeklyHolidayTypeCode = ChkStr(oRS.Collect("WeeklyHolidayTypeCode"))
			WeeklyHolidayTypeName = ChkStr(oRS.Collect("WeeklyHolidayType"))
		End If
		Call RSClose(oRS)
	End Function

	'******************************************************************************
	'概　要：
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function SetLicense()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr
		Dim tmp1
		Dim idx

		tmp1 = ""

		sSQL = "EXEC up_LstP_License '" & StaffCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
		If GetRSState(oRS) = True Then
			LicenseCount = oRS.RecordCount

			ReDim LicenseGroupCode(LicenseCount - 1)
			ReDim LicenseCategoryCode(LicenseCount - 1)
			ReDim LicenseCode(LicenseCount - 1)

			idx = 0
			Do While GetRSState(oRS) = True
				LicenseGroupCode(idx) = oRS.Collect("LicenseGroupCode")
				LicenseCategoryCode(idx) = oRS.Collect("LicenseCategoryCode")
				LicenseCode(idx) = oRS.Collect("LicenseCode")

				If tmp1 <> "" Then tmp1 = tmp1 & ","
				tmp1 = tmp1 & oRS.Collect("LicenseNameDsp")

				idx = idx + 1
				oRS.MoveNext
			Loop
		End If

		LicenseName = tmp1
	End Function

	'******************************************************************************
	'概　要：
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function SetSkill()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr
		Dim tmp1
		Dim tmp2

		sSQL = "EXEC sp_GetDataSkill '" & StaffCode & "','';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
		If GetRSState(oRS) = True Then
			Set oRS.ActiveConnection = Nothing

			'<ＯＳ>
			tmp1 = ""
			tmp2 = ""
			oRS.Filter = "CategoryCode = 'OS'"
			If GetRSState(oRS) = True Then
				Do While GetRSState(oRS) = True
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & oRS.Collect("Code")

					If tmp2 <> "" Then tmp2 = tmp2 & ","
					tmp2 = tmp2 & oRS.Collect("SkillName")

					oRS.MoveNext
				Loop
			End If
			OSCode = tmp1
			OSName = tmp2
			oRS.Filter = 0
			'</ＯＳ>

			'<アプリケーション>
			tmp1 = ""
			tmp2 = ""
			oRS.Filter = "CategoryCode = 'Application'"
			If GetRSState(oRS) = True Then
				Do While GetRSState(oRS) = True
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & oRS.Collect("Code")

					If tmp2 <> "" Then tmp2 = tmp2 & ","
					tmp2 = tmp2 & oRS.Collect("SkillName")

					oRS.MoveNext
				Loop
			End If
			APCode = tmp1
			APName = tmp2
			oRS.Filter = 0
			'</アプリケーション>

			'<開発言語>
			tmp1 = ""
			tmp2 = ""
			oRS.Filter = "CategoryCode = 'DevelopmentLanguage'"
			If GetRSState(oRS) = True Then
				Do While GetRSState(oRS) = True
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & oRS.Collect("Code")

					If tmp2 <> "" Then tmp2 = tmp2 & ","
					tmp2 = tmp2 & oRS.Collect("SkillName")

					oRS.MoveNext
				Loop
			End If
			oRS.Filter = 0
			DLCode = tmp1
			DLName = tmp2
			'</開発言語>

			'<データベース>
			tmp1 = ""
			tmp2 = ""
			oRS.Filter = "CategoryCode = 'Database'"
			If GetRSState(oRS) = True Then
				Do While GetRSState(oRS) = True
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & oRS.Collect("Code")

					If tmp2 <> "" Then tmp2 = tmp2 & ","
					tmp2 = tmp2 & oRS.Collect("SkillName")

					oRS.MoveNext
				Loop
			End If
			DBCode = tmp1
			DBName = tmp2
			oRS.Filter = 0
			'</データベース>
		End If
		Call RSClose(oRS)
	End Function

	'******************************************************************************
	'概　要：アドバイスツールのデータをお仕事検索パラメータに変換する
	'引　数：
	'備　考：
	'履　歴：2009/09/15 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function ChgParamSearchDetail()
		Dim oSOC

		Dim tmp
		Dim aValue
		Dim idx
		Dim idx2

		Set oSOC = New clsSearchOrderCondition

		oSOC.SearchDetailFlag = "1"

		'<希望職種>
		If HopeJobTypeFlag = "1" Then
			idx2 = 1
			aValue = Split(HopeJobTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If idx2 = 1 Then
						oSOC.JobTypeBigCode1 = Left(aValue(idx),2)
						oSOC.JobTypeCode1 = aValue(idx)
					ElseIf idx2 = 2 Then
						oSOC.JobTypeBigCode2 = Left(aValue(idx),2)
						oSOC.JobTypeCode2 = aValue(idx)
					ElseIf idx2 = 3 Then
						oSOC.JobTypeBigCode3 = Left(aValue(idx),2)
						oSOC.JobTypeCode3 = aValue(idx)
					End If

					idx2 = idx2 + 1
				End If
			Next
		End If
		'</希望職種>

		'<希望勤務地>
		If HopeWorkingPlaceFlag = "1" Then
			idx2 = 1
			aValue = Split(HopePrefectureCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If idx2 = 1 Then
						oSOC.PrefectureCode = aValue(idx)
					ElseIf idx2 = 2 Then
						oSOC.PrefectureCode = aValue(idx)
					End If

					idx2 = idx2 + 1
				End If
			Next
		End If
		'</希望勤務地>

		'<希望勤務形態>
		If HopeWorkingTypeFlag = "1" Then
			idx2 = 1
			aValue = Split(HopeWorkingTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If idx2 = 1 Then
						oSOC.WorkingTypeCode1 = aValue(idx)
					ElseIf idx2 = 2 Then
						oSOC.WorkingTypeCode2 = aValue(idx)
					ElseIf idx2 = 3 Then
						oSOC.WorkingTypeCode3 = aValue(idx)
					End If

					idx2 = idx2 + 1
				End If
			Next
		End If
		'</希望勤務形態>

		'<希望業種>
		If HopeIndustryTypeFlag = "1" Then
			oSOC.IndustryTypeCode = HopeIndustryTypeCode
		End If
		'</希望業種>

		'<給与>
		If SalaryFlag = "1" Then
			oSOC.YearlyIncomeMin = YearlyIncomeMin
			oSOC.YearlyIncomeMax = YearlyIncomeMax
			oSOC.MonthlyIncomeMin = MonthlyIncomeMin
			oSOC.MonthlyIncomeMax = MonthlyIncomeMax
			oSOC.DailyIncomeMin = DailyIncomeMin
			oSOC.DailyIncomeMax = DailyIncomeMax
			oSOC.HourlyIncomeMin = HourlyIncomeMin
			oSOC.HourlyIncomeMax = HourlyIncomeMax
		End If
		'</給与>

		'<勤務時間>
		If WorkTimeFlag = "1" Then
			oSOC.WorkStartHour = Left(WorkStartTime,2)
			oSOC.WorkStartMinute = Right(WorkStartTime,2)
			oSOC.WorkEndHour = Left(WorkEndTime,2)
			oSOC.WorkEndMinute = Right(WorkEndTime,2)
		End If
		'</勤務時間>

		'<週休種類>
		If HolidayFlag = "1" Then
			oSOC.WeeklyHolidayType = WeeklyHolidayTypeCode
		End If
		'</週休種類>

		'<資格>
		If LicenseFlag = "1" Then
			oSOC.LicenseCount = LicenseCount
			oSOC.LicenseGroupCode = LicenseGroupCode
			oSOC.LicenseCategoryCode = LicenseCategoryCode
			oSOC.LicenseCode = LicenseCode
		End If
		'</資格>

		'<スキル>
		If OSFlag = "1" Then
			oSOC.OSCode = OSCode
		End If
		If APFlag = "1" Then
			oSOC.ApplicationCode = APCode
		End If
		If DLFlag = "1" Then
			oSOC.DevelopmentLanguageCode = DLCode
		End If
		If DBFlag = "1" Then
			oSOC.DatabaseCode = DBCode
		End If
		'</スキル>

		ChgParamSearchDetail = oSOC.GetSearchParam()
	End Function
End Class
%>
