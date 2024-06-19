<%
'*******************************************************************************
'概　要：ナビのＴＯＰページに表示するログイン済み求職者の求人検索結果一覧（保存してある条件から）
'引　数：
'出　力：
'戻り値：String
'備　考：保存してある検索条件の番号 seq=1
'履　歴：2010/11/06 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlNaviTopOrderList1(ByRef rDB,ByVal vUserID)
	Dim oRS,oRS2,sSQL,flgQE,sSQLErr
	Dim dbSearchName,dbSearchParam
	Dim oSOC
	Dim iCnt

	sSQL = "EXEC up_LstP_SearchOrderCondition '" & vUserID & "';"
	flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing

		dbSearchName = oRS.Collect("SearchName")
		dbSearchParam = oRS.Collect("SearchParam")

		Set oSOC = New clsSearchOrderCondition
		oSOC.SetData_Param(dbSearchParam)
		sSQL = oSOC.GetSQLOrderSearchDetail()
		flgQE = QUERYEXE(rDB,oRS2,sSQL,sSQLErr)
		If GetRSState(oRS2) = True Then
			Set oRS2.ActiveConnection = Nothing

			iCnt = oRS2.RecordCount

			sHTML = htmlOrderListLine(rDB,oRS2,5)
			If iCnt > 5 Then
				sHTML = sHTML & "<p style=""text-align:right;""><a href=""" & HTTP_CURRENTURL & "order/order_list.asp?" & Replace(dbSearchParam,"&","&amp;") & """>" & _
					"&gt;&gt;もっと検索結果を見る" & _
					"</a></p>"
			End If
		Else
			sHTML = "<p>保存した求人の検索条件にマッチするお求人が見つかりませんでした。</p>"
		End If
	Else
		sHTML = "<p>求人の検索条件が保存されていません。" & _
			"求人の検索条件を保存するには、求人の検索結果一覧で「この検索条件を保存する」をクリックします。" & _
			"求人の検索は<a href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">コチラ</a>からどうぞ。</p>"
	End If
	Call RSClose(oRS)

	htmlNaviTopOrderList1 = sHTML
End Function
%>
