<%
'*******************************************************************************
'概　要：求人票一覧表示パターン01
'引　数：rDB	：DBコネクション
'　　　：vXML	：求人票一覧のXML
'　　　：vStr	：SEO対策用の追加文言
'出　力：
'戻り値：String
'備　考：
'履　歴：2011/02/17 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlOrderList01(ByRef rDB,ByVal vXML,ByVal vStr)
	Dim sSQL,oRS,sSQLErr,flgQE
	Dim dbOrderCode,dbJobTypeDetail,dbBusinessDetail,dbYearlyIncomeMin,dbYearlyIncomeMax,dbMonthlyIncomeMin,dbMonthlyIncomeMax,dbDailyIncomeMin,dbDailyIncomeMax,dbHourlyIncomeMin,dbHourlyIncomeMax,dbCompanyName,dbWorkingType,dbWorkingPlace,dbNearbyStation,dbPicSrc
	Dim aTmp,sCompanyName,sWorkingPlace,sImg,sSalary

	Dim sHTML

	sSQL = "EXEC up_DtlOrder_Recommend_XML '" & vXML & "';"
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sSQLErr)
	Do While GetRSState(oRS) = True
		dbOrderCode = oRS.Collect("OrderCode")
		dbJobTypeDetail = oRS.Collect("JobTypeDetail")
		dbBusinessDetail = ChkStr(oRS.Collect("BusinessDetail"))
		dbYearlyIncomeMin = oRS.Collect("YearlyIncomeMin")
		dbYearlyIncomeMax = oRS.Collect("YearlyIncomeMax")
		dbMonthlyIncomeMin = oRS.Collect("MonthlyIncomeMin")
		dbMonthlyIncomeMax = oRS.Collect("MonthlyIncomeMax")
		dbDailyIncomeMin = oRS.Collect("DailyIncomeMin")
		dbDailyIncomeMax = oRS.Collect("DailyIncomeMax")
		dbHourlyIncomeMin = oRS.Collect("HourlyIncomeMin")
		dbHourlyIncomeMax = oRS.Collect("HourlyIncomeMax")
		dbCompanyName = oRS.Collect("CompanyName")
		dbWorkingType = oRS.Collect("WorkingType")
		dbWorkingPlace = oRS.Collect("WorkingPlace")
		dbNearbyStation = oRS.Collect("NearbyStation")
		dbPicSrc = ChkStr(oRS.Collect("PicSrc"))

		'<画像src調整>
		If dbPicSrc <> "" Then
			sImg = "<a href=""/order/order_detail.asp?ordercode=" & dbOrderCode & """ target=""_blank""><img src=""" & dbPicSrc & """ alt=""" & dbJobTypeDetail & """ style=""border:1px solid #cccccc; max-width:280px; max-height:210px;""></a>"
		End If
		'</画像src調整>

		'<勤務地文言調整>
		sWorkingPlace = Left(dbWorkingPlace,30)
		If Len(dbWorkingPlace) > 30 Then sWorkingPlace = sWorkingPlace & "..."
		'</勤務地文言調整>

		'<会社名調整>
		aTmp = Split(dbCompanyName,vbTab)
		sCompanyName = aTmp(0)
		'</会社名調整>

		'<給与文言調整>
		sSalary = ""
		If dbYearlyIncomeMin + dbYearlyIncomeMax > 0 Then
			sSalary = sSalary & "[年収]&nbsp;"
			If dbYearlyIncomeMin = dbYearlyIncomeMax Then
				sSalary = sSalary & GetJapaneseYen(dbYearlyIncomeMin) & "<br>"
			Else
				If dbYearlyIncomeMin > 0 Then sSalary = sSalary & GetJapaneseYen(dbYearlyIncomeMin) & "&nbsp;"
				sSalary = sSalary & "～"
				If dbYearlyIncomeMax > 0 Then sSalary = sSalary & "&nbsp;" & GetJapaneseYen(dbYearlyIncomeMax) & "　"
			End If
		End If
		If dbMonthlyIncomeMin + dbMonthlyIncomeMax > 0 Then
			sSalary = sSalary & "[月給]&nbsp;"
			If dbMonthlyIncomeMin = dbMonthlyIncomeMax Then
				sSalary = sSalary & GetJapaneseYen(dbMonthlyIncomeMin)
			Else
				If dbMonthlyIncomeMin > 0 Then sSalary = sSalary & GetJapaneseYen(dbMonthlyIncomeMin) & "&nbsp;"
				sSalary = sSalary & "～"
				If dbMonthlyIncomeMax > 0 Then sSalary = sSalary & "&nbsp;" & GetJapaneseYen(dbMonthlyIncomeMax) & "　"
			End If
		End If
		If dbDailyIncomeMin + dbDailyIncomeMax > 0 Then
			sSalary = sSalary & "[日給]&nbsp;"
			If dbDailyIncomeMin = dbDailyIncomeMax Then
				sSalary = sSalary & GetJapaneseYen(dbDailyIncomeMin) & "<br>"
			Else
				If dbDailyIncomeMin > 0 Then sSalary = sSalary & GetJapaneseYen(dbDailyIncomeMin) & "&nbsp;"
				sSalary = sSalary & "～"
				If dbDailyIncomeMax > 0 Then sSalary = sSalary & "&nbsp;" & GetJapaneseYen(dbDailyIncomeMax) & "　"
			End If
		End If
		If dbHourlyIncomeMin + dbHourlyIncomeMax > 0 Then
			sSalary = sSalary & "[時給]&nbsp;"
			If dbHourlyIncomeMin = dbHourlyIncomeMax Then
				sSalary = sSalary & GetJapaneseYen(dbHourlyIncomeMin) & "<br>"
			Else
				If dbHourlyIncomeMin > 0 Then sSalary = sSalary & GetJapaneseYen(dbHourlyIncomeMin) & "&nbsp;"
				sSalary = sSalary & "～"
				If dbHourlyIncomeMax > 0 Then sSalary = sSalary & "&nbsp;" & GetJapaneseYen(dbHourlyIncomeMax) & "　"
			End If
		End If
		'</給与文言調整>


		sHTML = sHTML & "<div class=""description2"">"

		sHTML = sHTML & "<p class=""m0"" style=""font-weight:bold;"">" & sCompanyName & "</p>"
		sHTML = sHTML & "<p><a href=""/order/order_detail.asp?ordercode=" & dbOrderCode & """ target=""_blank"">" & dbJobTypeDetail & "</a></p>"

		sHTML = sHTML & "<div style=""float:left;width:280px; height:240px; text-align:center; vertical-align:middle;"" class=""center"">"
		sHTML = sHTML & sImg & "<br>"
		sHTML = sHTML & "<p style=""font-size:8pt;text-align:center;"">" & vStr & "</p>"
		sHTML = sHTML & "</div>"

		sHTML = sHTML & "<div style=""float:right;width:430px;"" class=""inpSmart"">"
		sHTML = sHTML & "<p>" & Replace(dbBusinessDetail,vbCrLf,"<br>") & "</p>"
		sHTML = sHTML & "<div class=""line1""></div>"
		sHTML = sHTML & "<div style=""float:left;width:25%;"">【勤務形態】</div><div style=""float:right;width:70%;"">" & dbWorkingType & "</div><div style=""clear:both;""></div>"
		sHTML = sHTML & "<div class=""line1""></div>"
		sHTML = sHTML & "<div style=""float:left;width:25%;"">【就業場所】</div><div style=""float:right;width:70%;"">" & sWorkingPlace & "</div><div style=""clear:both;""></div>"
		sHTML = sHTML & "<div class=""line1""></div>"
		If dbNearbyStation <> "" Then
			sHTML = sHTML & "<div style=""float:left;width:25%;"">【最寄駅】</div><div style=""float:right;width:70%;"">" & dbNearbyStation & "</div><div style=""clear:both;""></div>"
			sHTML = sHTML & "<div class=""line1""></div>"
		End If
		sHTML = sHTML & "<div style=""float:left;width:25%;"">【給与】</div><div style=""float:right;width:70%;"">" & sSalary & "</div><div style=""clear:both;""></div>"

		sHTML = sHTML & "</div><div style=""clear:both;""></div>"

		sHTML = sHTML & "</div>"

		oRS.MoveNext
	Loop

	htmlOrderList01 = sHTML
End Function
%>
