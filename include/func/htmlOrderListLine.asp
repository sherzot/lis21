<%
'*******************************************************************************
'�T�@�v�F���l�[�ꗗ�̂P�s�\���o�[�W����
'���@���FrDB	�FDB�R�l�N�V����
'�@�@�@�FrRS	�F���l�[�ꗗ�̃��R�[�h�Z�b�g
'�@�@�@�FvMaxRow�F���l�\������
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2010/11/05 LIS K.Kokubo �쐬
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

		'<�Ζ��n>
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

				If idx2 > 3 Then sWP = sWP & ",�c"
			Loop
		End If
		Call RSClose(oRS)
		'</�Ζ��n>

		'<�Ζ��`��>
		sWT = ""
		sSQL = "EXEC up_LstC_WorkingType '" & dbOrderCode & "';"
		flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
		If GetRSState(oRS) = True Then
			Set oRS.ActiveConnection = Nothing

			Do While GetRSState(oRS)
				dbWorkingTypeCode = ChkStr(oRS.Collect("WorkingTypeCode"))
				dbWorkingTypeName = ChkStr(oRS.Collect("WorkingTypeName"))
				Select Case dbWorkingTypeCode
					Case "001": sWT = sWT & "<img src=""/img/haken.gif"" alt=""�h��"" style=""margin-right:1px;border-width:0px;"">"
					Case "002": sWT = sWT & "<img src=""/img/seishain.gif"" alt=""���Ј�"" style=""margin-right:1px;border-width:0px;"">"
					Case "003": sWT = sWT & "<img src=""/img/keiyaku.gif"" alt=""�_��Ј�"" style=""margin-right:1px;border-width:0px;"">"
					Case "004": sWT = sWT & "<img src=""/img/syoha.gif"" alt=""�Љ�\��h��"" style=""margin-right:1px;border-width:0px;"">"
					Case "005": sWT = sWT & "<img src=""/img/arbeit.gif"" alt=""�A���o�C�g�E�p�[�g"" style=""margin-right:1px;border-width:0px;"">"
					Case "006": sWT = sWT & "<img src=""/img/soho.gif"" alt=""SOHO"" style=""margin-right:1px;border-width:0px;"">"
					Case "007": sWT = sWT & "<img src=""/img/fc.gif"" alt=""FC"" style=""margin-right:1px;border-width:0px;"">"
				End Select

				'If sWT <> "" Then sWT = sWT & ","
				'sWT = sWT & dbWorkingTypeName

				oRS.MoveNext
			Loop
		End If
		Call RSClose(oRS)
		'<�Ζ��`��>

		'<���l�[�ڍ�>
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

			'<���^>
			If dbYearlyIncomeMin <> "" Then
				sSalary = "�y�N���z" & GetJapaneseYen(dbYearlyIncomeMin) & "�`"
				If dbYearlyIncomeMax <> "" Then sSalary = sSalary & GetJapaneseYen(dbYearlyIncomeMax)
			ElseIf dbMonthlyIncomeMin <> "" Then
				sSalary = "�y�����z" & GetJapaneseYen(dbMonthlyIncomeMin) & "�`"
				If dbMonthlyIncomeMax <> "" Then sSalary = sSalary & GetJapaneseYen(dbMonthlyIncomeMax)
			ElseIf dbDailyIncomeMin <> "" Then
				sSalary = "�y�����z" & GetJapaneseYen(dbDailyIncomeMin) & "�`"
				If dbDailyIncomeMax <> "" Then sSalary = sSalary & GetJapaneseYen(dbDailyIncomeMax)
			ElseIf dbHourlyIncomeMin <> "" Then
				sSalary = "�y�����z" & GetJapaneseYen(dbHourlyIncomeMin) & "�`"
				If dbHourlyIncomeMax <> "" Then sSalary = sSalary & GetJapaneseYen(dbHourlyIncomeMax)
			End If
			'</���^>
		End If
		Call RSClose(oRS)
		'</���l�[�ڍ�>

		sHTML = sHTML & "<div style=""margin-bottom:3px;""><a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?ordercode=" & dbOrderCode & """>" & sWP & "/" & sSalary & "/" & dbJobTypeDetail & "/" & sWT & "</a></div>"

		idx = idx + 1
		rRS.MoveNext
	Loop
	Call RSClose(rRS)

	htmlOrderListLine = sHTML
End Function
%>
