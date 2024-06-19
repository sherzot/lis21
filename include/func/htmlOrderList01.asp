<%
'*******************************************************************************
'�T�@�v�F���l�[�ꗗ�\���p�^�[��01
'���@���FrDB	�FDB�R�l�N�V����
'�@�@�@�FvXML	�F���l�[�ꗗ��XML
'�@�@�@�FvStr	�FSEO�΍��p�̒ǉ�����
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2011/02/17 LIS K.Kokubo �쐬
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

		'<�摜src����>
		If dbPicSrc <> "" Then
			sImg = "<a href=""/order/order_detail.asp?ordercode=" & dbOrderCode & """ target=""_blank""><img src=""" & dbPicSrc & """ alt=""" & dbJobTypeDetail & """ style=""border:1px solid #cccccc; max-width:280px; max-height:210px;""></a>"
		End If
		'</�摜src����>

		'<�Ζ��n��������>
		sWorkingPlace = Left(dbWorkingPlace,30)
		If Len(dbWorkingPlace) > 30 Then sWorkingPlace = sWorkingPlace & "..."
		'</�Ζ��n��������>

		'<��Ж�����>
		aTmp = Split(dbCompanyName,vbTab)
		sCompanyName = aTmp(0)
		'</��Ж�����>

		'<���^��������>
		sSalary = ""
		If dbYearlyIncomeMin + dbYearlyIncomeMax > 0 Then
			sSalary = sSalary & "[�N��]&nbsp;"
			If dbYearlyIncomeMin = dbYearlyIncomeMax Then
				sSalary = sSalary & GetJapaneseYen(dbYearlyIncomeMin) & "<br>"
			Else
				If dbYearlyIncomeMin > 0 Then sSalary = sSalary & GetJapaneseYen(dbYearlyIncomeMin) & "&nbsp;"
				sSalary = sSalary & "�`"
				If dbYearlyIncomeMax > 0 Then sSalary = sSalary & "&nbsp;" & GetJapaneseYen(dbYearlyIncomeMax) & "�@"
			End If
		End If
		If dbMonthlyIncomeMin + dbMonthlyIncomeMax > 0 Then
			sSalary = sSalary & "[����]&nbsp;"
			If dbMonthlyIncomeMin = dbMonthlyIncomeMax Then
				sSalary = sSalary & GetJapaneseYen(dbMonthlyIncomeMin)
			Else
				If dbMonthlyIncomeMin > 0 Then sSalary = sSalary & GetJapaneseYen(dbMonthlyIncomeMin) & "&nbsp;"
				sSalary = sSalary & "�`"
				If dbMonthlyIncomeMax > 0 Then sSalary = sSalary & "&nbsp;" & GetJapaneseYen(dbMonthlyIncomeMax) & "�@"
			End If
		End If
		If dbDailyIncomeMin + dbDailyIncomeMax > 0 Then
			sSalary = sSalary & "[����]&nbsp;"
			If dbDailyIncomeMin = dbDailyIncomeMax Then
				sSalary = sSalary & GetJapaneseYen(dbDailyIncomeMin) & "<br>"
			Else
				If dbDailyIncomeMin > 0 Then sSalary = sSalary & GetJapaneseYen(dbDailyIncomeMin) & "&nbsp;"
				sSalary = sSalary & "�`"
				If dbDailyIncomeMax > 0 Then sSalary = sSalary & "&nbsp;" & GetJapaneseYen(dbDailyIncomeMax) & "�@"
			End If
		End If
		If dbHourlyIncomeMin + dbHourlyIncomeMax > 0 Then
			sSalary = sSalary & "[����]&nbsp;"
			If dbHourlyIncomeMin = dbHourlyIncomeMax Then
				sSalary = sSalary & GetJapaneseYen(dbHourlyIncomeMin) & "<br>"
			Else
				If dbHourlyIncomeMin > 0 Then sSalary = sSalary & GetJapaneseYen(dbHourlyIncomeMin) & "&nbsp;"
				sSalary = sSalary & "�`"
				If dbHourlyIncomeMax > 0 Then sSalary = sSalary & "&nbsp;" & GetJapaneseYen(dbHourlyIncomeMax) & "�@"
			End If
		End If
		'</���^��������>


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
		sHTML = sHTML & "<div style=""float:left;width:25%;"">�y�Ζ��`�ԁz</div><div style=""float:right;width:70%;"">" & dbWorkingType & "</div><div style=""clear:both;""></div>"
		sHTML = sHTML & "<div class=""line1""></div>"
		sHTML = sHTML & "<div style=""float:left;width:25%;"">�y�A�Əꏊ�z</div><div style=""float:right;width:70%;"">" & sWorkingPlace & "</div><div style=""clear:both;""></div>"
		sHTML = sHTML & "<div class=""line1""></div>"
		If dbNearbyStation <> "" Then
			sHTML = sHTML & "<div style=""float:left;width:25%;"">�y�Ŋ�w�z</div><div style=""float:right;width:70%;"">" & dbNearbyStation & "</div><div style=""clear:both;""></div>"
			sHTML = sHTML & "<div class=""line1""></div>"
		End If
		sHTML = sHTML & "<div style=""float:left;width:25%;"">�y���^�z</div><div style=""float:right;width:70%;"">" & sSalary & "</div><div style=""clear:both;""></div>"

		sHTML = sHTML & "</div><div style=""clear:both;""></div>"

		sHTML = sHTML & "</div>"

		oRS.MoveNext
	Loop

	htmlOrderList01 = sHTML
End Function
%>
