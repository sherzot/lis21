<%
'*******************************************************************************
'T@vFl[κΜPs\¦o[W
'ψ@FrDB	FDBRlNV
'@@@FvXML	Fl[κΜXML
'@@@FvMaxRowFl\¦
'o@ΝF
'ίθlFString
'υ@lF
'@πF2011/02/17 LIS K.Kokubo μ¬
'*******************************************************************************
Function htmlOrderListBlock01(ByRef rDB,ByVal vXML,ByVal vMaxCols)
	Dim sSQL,oRS,sSQLErr,flgQE
	Dim dbOrderCode,dbJobTypeDetail,dbYearlyIncomeMin,dbYearlyIncomeMax,dbMonthlyIncomeMin,dbMonthlyIncomeMax,dbDailyIncomeMin,dbDailyIncomeMax,dbHourlyIncomeMin,dbHourlyIncomeMax,dbCompanyName,dbWorkingType,dbWorkingPlace,dbNearbyStation,dbPicSrc
	Dim aTmp,sCompanyName
	Dim sHTML
	Dim idx

	vMaxCols = 3 'Ζθ ¦Έ3ΕΕθB
	idx = 1

	sSQL = "EXEC up_DtlOrder_Recommend_XML '" & vXML & "';"
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sSQLErr)
	Do While GetRSState(oRS) = True
		dbOrderCode = oRS.Collect("OrderCode")
		dbJobTypeDetail = oRS.Collect("JobTypeDetail")
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

		If Len(dbWorkingPlace) > 15 Then dbWorkingPlace = Left(dbWorkingPlace,15) & "..."
		aTmp = Split(dbCompanyName,vbTab)
		sCompanyName = aTmp(0)

		sHTML = sHTML & "<div style=""float:left; width:200px;"">"

		If dbPicSrc <> "" Then sHTML = sHTML & "<div style=""text-align:center;margin-bottom:2px;""><a href=""/order/order_detail.asp?ordercode=" & dbOrderCode & """ target=""_blank""><img src=""" & dbPicSrc & """ alt=""" & dbJobTypeDetail & """ width=""180"" height=""135"" style=""border:1px solid #cccccc;""></a></div>"
		sHTML = sHTML & "<p style=""margin-bottom:2px;font-size:10px;text-align:center;"">ξρR[hF" & dbOrderCode & "</p>"

		sHTML = sHTML & "<div class=""description2"" style=""margin:3px;"">"

		If dbCompanyName <> "" Then sHTML = sHTML & "yοΠz<br>" & sCompanyName & "<div class=""line1""></div>"
		If dbJobTypeDetail <> "" Then sHTML = sHTML & "yEνz<br><a href=""/order/order_detail.asp?ordercode=" & dbOrderCode & """ target=""_blank"">" & dbJobTypeDetail & "</a><div class=""line1""></div>"
		If dbWorkingType <> "" Then sHTML = sHTML & "yΞ±`Τz<br>" & dbWorkingType  & "<div class=""line1""></div>"
		If dbWorkingPlace <> "" Then sHTML = sHTML & "yAΖκz<br>" & dbWorkingPlace & "<div class=""line1""></div>"
		If dbNearbyStation <> "" Then sHTML = sHTML & "yΕρwz<br>" & dbNearbyStation & "<div class=""line1""></div>"
		sHTML = sHTML & "y^z<br>"
		If dbYearlyIncomeMin + dbYearlyIncomeMax > 0 Then
			sHTML = sHTML & "[Nϋ]&nbsp;"
			If dbYearlyIncomeMin = dbYearlyIncomeMax Then
				sHTML = sHTML & GetJapaneseYen(dbYearlyIncomeMin) & "<br>"
			Else
				If dbYearlyIncomeMin > 0 Then sHTML = sHTML & GetJapaneseYen(dbYearlyIncomeMin) & "&nbsp;"
				sHTML = sHTML & "`"
				If dbYearlyIncomeMax > 0 Then sHTML = sHTML & "&nbsp;" & GetJapaneseYen(dbYearlyIncomeMax) & "<br>"
			End If
		End If
		If dbMonthlyIncomeMin + dbMonthlyIncomeMax > 0 Then
			sHTML = sHTML & "[]&nbsp;"
			If dbMonthlyIncomeMin = dbMonthlyIncomeMax Then
				sHTML = sHTML & GetJapaneseYen(dbMonthlyIncomeMin)
			Else
				If dbMonthlyIncomeMin > 0 Then sHTML = sHTML & GetJapaneseYen(dbMonthlyIncomeMin) & "&nbsp;"
				sHTML = sHTML & "`"
				If dbMonthlyIncomeMax > 0 Then sHTML = sHTML & "&nbsp;" & GetJapaneseYen(dbMonthlyIncomeMax) & "<br>"
			End If
		End If
		If dbDailyIncomeMin + dbDailyIncomeMax > 0 Then
			sHTML = sHTML & "[ϊ]&nbsp;"
			If dbDailyIncomeMin = dbDailyIncomeMax Then
				sHTML = sHTML & GetJapaneseYen(dbDailyIncomeMin) & "<br>"
			Else
				If dbDailyIncomeMin > 0 Then sHTML = sHTML & GetJapaneseYen(dbDailyIncomeMin) & "&nbsp;"
				sHTML = sHTML & "`"
				If dbDailyIncomeMax > 0 Then sHTML = sHTML & "&nbsp;" & GetJapaneseYen(dbDailyIncomeMax) & "<br>"
			End If
		End If
		If dbHourlyIncomeMin + dbHourlyIncomeMax > 0 Then
			sHTML = sHTML & "[]&nbsp;"
			If dbHourlyIncomeMin = dbHourlyIncomeMax Then
				sHTML = sHTML & GetJapaneseYen(dbHourlyIncomeMin) & "<br>"
			Else
				If dbHourlyIncomeMin > 0 Then sHTML = sHTML & GetJapaneseYen(dbHourlyIncomeMin) & "&nbsp;"
				sHTML = sHTML & "`"
				If dbHourlyIncomeMax > 0 Then sHTML = sHTML & "&nbsp;" & GetJapaneseYen(dbHourlyIncomeMax) & "<br>"
			End If
		End If

		sHTML = sHTML & "</div></div>"

		oRS.MoveNext

		If GetRSState(oRS) = False Or idx Mod vMaxCols = 0 Then sHTML = sHTML & "<br clear=""all"">"

		idx = idx + 1
	Loop

	htmlOrderListBlock01 = sHTML
End Function
%>
