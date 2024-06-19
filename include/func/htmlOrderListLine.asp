<%
'*******************************************************************************
'概　要：求人票一覧の１行表示バージョン
'引　数：rDB	：DBコネクション
'　　　：rRS	：求人票一覧のレコードセット
'　　　：vMaxRow：求人表示件数
'出　力：
'戻り値：String
'備　考：
'履　歴：2010/11/05 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlOrderListLine(ByRef rDB,ByRef rRS,ByVal vMaxRow)
	Dim oRS,sSQL,flgQE,sSQLErr
	Dim dbOrderCode,dbJobTypeDetail
	Dim dbWorkingPlacePrefectureName,dbWorkingPlaceCity,dbWorkingTypeCode,dbWorkingTypeName
	Dim dbYearlyIncomeMin,dbYearlyIncomeMax,dbMonthlyIncomeMin,dbMonthlyIncomeMax,dbDailyIncomeMin,dbDailyIncomeMax,dbHourlyIncomeMin,dbHourlyIncomeMax
	Dim sHTML,sWP,sWT,sSalary
	Dim idx,idx2

	If GetRSState(rRS) = False Then Exit Function

	idx = 0
	Do While GetRSState(rRS) = True And idx < vMaxRow
		dbOrderCode = rRS.Collect("OrderCode")

		'<勤務地>
		sWP = ""
		sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
		flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
		If GetRSState(oRS) = True Then
			Set oRS.ActiveConnection = Nothing

			idx2 = 0
			Do While GetRSState(oRS)
				dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
				dbWorkingPlaceCity = ChkStr(oRS.Collect("WorkingPlaceCity"))

				If sWP <> "" Then sWP = sWP & ","

				sWP = sWP & dbWorkingPlacePrefectureName & dbWorkingPlaceCity

				idx2 = idx2 + 1
				oRS.MoveNext

				If idx2 > 3 Then sWP = sWP & ",…"
			Loop
		End If
		Call RSClose(oRS)
		'</勤務地>

		'<勤務形態>
		sWT = ""
		sSQL = "EXEC up_LstC_WorkingType '" & dbOrderCode & "';"
		flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
		If GetRSState(oRS) = True Then
			Set oRS.ActiveConnection = Nothing

			Do While GetRSState(oRS)
				dbWorkingTypeCode = ChkStr(oRS.Collect("WorkingTypeCode"))
				dbWorkingTypeName = ChkStr(oRS.Collect("WorkingTypeName"))
				Select Case dbWorkingTypeCode
					Case "001": sWT = sWT & "<img src=""/img/haken.gif"" alt=""派遣"" style=""margin-right:1px;border-width:0px;"">"
					Case "002": sWT = sWT & "<img src=""/img/seishain.gif"" alt=""正社員"" style=""margin-right:1px;border-width:0px;"">"
					Case "003": sWT = sWT & "<img src=""/img/keiyaku.gif"" alt=""契約社員"" style=""margin-right:1px;border-width:0px;"">"
					Case "004": sWT = sWT & "<img src=""/img/syoha.gif"" alt=""紹介予定派遣"" style=""margin-right:1px;border-width:0px;"">"
					Case "005": sWT = sWT & "<img src=""/img/arbeit.gif"" alt=""アルバイト・パート"" style=""margin-right:1px;border-width:0px;"">"
					Case "006": sWT = sWT & "<img src=""/img/soho.gif"" alt=""SOHO"" style=""margin-right:1px;border-width:0px;"">"
					Case "007": sWT = sWT & "<img src=""/img/fc.gif"" alt=""FC"" style=""margin-right:1px;border-width:0px;"">"
				End Select

				'If sWT <> "" Then sWT = sWT & ","
				'sWT = sWT & dbWorkingTypeName

				oRS.MoveNext
			Loop
		End If
		Call RSClose(oRS)
		'<勤務形態>

		'<求人票詳細>
		sSQL = "EXEC up_DtlOrder '" & dbOrderCode & "','';"
		flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
		If GetRSState(oRS) = True Then
			Set oRS.ActiveConnection = Nothing

			dbJobTypeDetail = oRS.Collect("JobTypeDetail")
			dbYearlyIncomeMin = ChkStr(oRS.Collect("YearlyIncomeMin"))
			dbYearlyIncomeMax = ChkStr(oRS.Collect("YearlyIncomeMax"))
			dbMonthlyIncomeMin = ChkStr(oRS.Collect("MonthlyIncomeMin"))
			dbMonthlyIncomeMax = ChkStr(oRS.Collect("MonthlyIncomeMax"))
			dbDailyIncomeMin = ChkStr(oRS.Collect("DailyIncomeMin"))
			dbDailyIncomeMax = ChkStr(oRS.Collect("DailyIncomeMax"))
			dbHourlyIncomeMin = ChkStr(oRS.Collect("HourlyIncomeMin"))
			dbHourlyIncomeMax = ChkStr(oRS.Collect("HourlyIncomeMax"))

			'<給与>
			If dbYearlyIncomeMin <> "" Then
				sSalary = "【年収】" & GetJapaneseYen(dbYearlyIncomeMin) & "～"
				If dbYearlyIncomeMax <> "" Then sSalary = sSalary & GetJapaneseYen(dbYearlyIncomeMax)
			ElseIf dbMonthlyIncomeMin <> "" Then
				sSalary = "【月給】" & GetJapaneseYen(dbMonthlyIncomeMin) & "～"
				If dbMonthlyIncomeMax <> "" Then sSalary = sSalary & GetJapaneseYen(dbMonthlyIncomeMax)
			ElseIf dbDailyIncomeMin <> "" Then
				sSalary = "【日給】" & GetJapaneseYen(dbDailyIncomeMin) & "～"
				If dbDailyIncomeMax <> "" Then sSalary = sSalary & GetJapaneseYen(dbDailyIncomeMax)
			ElseIf dbHourlyIncomeMin <> "" Then
				sSalary = "【時給】" & GetJapaneseYen(dbHourlyIncomeMin) & "～"
				If dbHourlyIncomeMax <> "" Then sSalary = sSalary & GetJapaneseYen(dbHourlyIncomeMax)
			End If
			'</給与>
		End If
		Call RSClose(oRS)
		'</求人票詳細>

		sHTML = sHTML & "<div style=""margin-bottom:3px;""><a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?ordercode=" & dbOrderCode & """>" & sWP & "/" & sSalary & "/" & dbJobTypeDetail & "/" & sWT & "</a></div>"

		idx = idx + 1
		rRS.MoveNext
	Loop
	Call RSClose(rRS)

	htmlOrderListLine = sHTML
End Function
%>
