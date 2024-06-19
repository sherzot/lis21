<%
'*******************************************************************************
'概　要：ナビのＴＯＰページに表示するログイン済み求職者の求人検索結果一覧（自動検索）
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2010/10/29 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlNaviTopOrderList2(ByRef rDB,ByVal vUserID)
	Dim oRS,oRS2,sSQL,flgQE,sSQLErr
	Dim dbOrderCode,dbJobTypeDetail
	Dim dbWorkingPlacePrefectureName,dbWorkingPlaceCity,dbWorkingTypeName
	Dim dbYearlyIncomeMin,dbYearlyIncomeMax,dbMonthlyIncomeMin,dbMonthlyIncomeMax,dbDailyIncomeMin,dbDailyIncomeMax,dbHourlyIncomeMin,dbHourlyIncomeMax
	Dim sHTML,sWP,sWT,sSalary
	Dim idx,idx2

	sSQL = "EXEC up_SearchOrderAuto '" & vUserID & "','';"
	flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing

		sHTML = htmlOrderListLine(rDB,oRS,5)
	Else
		sHTML = "<p>あなたの希望条件にマッチする求人が見つかりませんでした。</p>"
	End If

	htmlNaviTopOrderList2 = sHTML
End Function
%>
