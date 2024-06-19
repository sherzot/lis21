<%
'*******************************************************************************
'概　要：採用改善サポートシステムの部門・店舗別費用対効果TABLEを取得
'引　数：rDB	：接続中DBコネクション
'　　　：vUserID：ログイン中ユーザID
'　　　：vYM1	：集計期間下限年月
'　　　：vYM2	：集計期間上限年月
'戻り値：String
'備　考：
'履　歴：2010/03/10 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlCostPerformance_Branch(ByRef rDB, ByRef rRS, ByVal vUserID, ByVal rCP, ByVal vPageSize, ByVal vPage, ByVal vSort, ByVal vBranchName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr
	'DB
	Dim dbCompanyCode
	Dim dbBranchSeq
	Dim dbBranchName
	Dim dbMedName
	Dim dbCost
	Dim dbAdoptNumPlan
	Dim dbAdoptNumResult
	Dim dbUnitCost
	Dim dbAdoptNumPlanPeriod

	Dim iCost
	Dim iUnitCost
	Dim iAdoptNumPlan
	Dim iAdoptNum

	Dim idx
	Dim sHref
	Dim sFlag
	Dim sHTML
	Dim sFilter

	Dim aBN

	sHTML = ""

	If GetRSState(rRS) = True Then
		rRS.PageSize = vPageSize
		If IsNumber(vPage,0,False) = False Then vPage = 1
		If rRS.PageCount < CInt(vPage) Then vPage = rRS.PageCount

		If vSort = "" Then
			rRS.Sort = "SortNum,UnitCost"
		Else
			rRS.Sort = "SortNum DESC,UnitCost DESC"
		End If

		iCost = 0
		iAdoptNumPlan = 0
		iAdoptNum = 0
		Do While GetRSState(rRS) = True
			dbCost = rRS.Collect("Cost")
			dbAdoptNumPlan = rRS.Collect("AdoptNumPlan")
			dbAdoptNumResult = rRS.Collect("AdoptNumResult")
			dbAdoptNumPlanPeriod = rRS.Collect("AdoptNumPlanPeriod")

			iCost = iCost + dbCost
			iAdoptNumPlan = iAdoptNumPlan + dbAdoptNumPlanPeriod
			iAdoptNum = iAdoptNum + dbAdoptNumResult

			rRS.MoveNext
		Loop
		rRS.MoveFirst
		iUnitCost = 0
		If iAdoptNum > 0 Then iUnitCost = iCost / iAdoptNum

		If vBranchName <> "" Then
			sFilter = ""
			aBN = Split(Replace(vBranchName,"　"," ")," ")
			For idx = 0 To UBound(aBN)
				If sFilter <> "" Then sFilter = sFilter & "OR "
				sFilter = sFilter & "BranchName LIKE '%" & aBN(idx) & "%' "
			Next
			rRS.Filter = Trim(sFilter)
		End If
	End If

	If GetRSState(rRS) = True Then
		sHTML = sHTML & "<table class=""pattern6"" border=""0"" style=""width:100%;margin-bottom:15px;"">"
		sHTML = sHTML & "<colgroup>"
		sHTML = sHTML & "<col style=""width:5%;"">"
		sHTML = sHTML & "<col style=""width:19%;"">"
		sHTML = sHTML & "<col style=""width:16%;"">"
		sHTML = sHTML & "<col style=""width:16%;"">"
		sHTML = sHTML & "<col style=""width:16%;"">"
		sHTML = sHTML & "<col style=""width:25%;"">"
		sHTML = sHTML & "</colgroup>"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th></th>"
		sHTML = sHTML & "<th style=""text-align:center;"">部門・店舗</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">採用ｺｽﾄ総額</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">採用人数<br>(実績/計画)</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">1名採用ｺｽﾄ</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">媒体</th>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</thead>"
		sHTML = sHTML & "<tbody>"

		rRS.AbsolutePage = vPage
		idx = 0
		Do While GetRSState(rRS) = True And idx < vPageSize
			dbCompanyCode = rRS.Collect("CompanyCode")
			dbBranchSeq = rRS.Collect("BranchSeq")
			dbBranchName = rRS.Collect("BranchName")
			dbCost = rRS.Collect("Cost")
			dbAdoptNumPlan = rRS.Collect("AdoptNumPlan")
			dbAdoptNumResult = rRS.Collect("AdoptNumResult")
			dbUnitCost = rRS.Collect("UnitCost")
			dbMedName = rRS.Collect("MedName")
			dbAdoptNumPlanPeriod = rRS.Collect("AdoptNumPlanPeriod")

			sFlag = ""
			If dbUnitCost > iUnitCost Or dbUnitCost = 0 Then sFlag = "▲"

			sHref = rCP.GetSearchParam()
			If sHref <> "" Then
				sHref = HTTPS_CURRENTURL & "company/costperformance/branch.asp" & sHref & "&dcc=" & dbCompanyCode & "&branchseq=" & dbBranchSeq
			Else
				sHref = HTTPS_CURRENTURL & "company/costperformance/branch.asp?dcc=" & dbCompanyCode & "&branchseq=" & dbBranchSeq
			End If

			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>" & sFlag & "</td>"
			sHTML = sHTML & "<td><a href=""" & Replace(sHref,"&","&amp;") & """>" & dbBranchName & "</a></td>"
			sHTML = sHTML & "<td style=""text-align:right;"">" & FormatCurrency(Round(dbCost)) & "</td>"
			sHTML = sHTML & "<td style=""text-align:right;"">"
			If dbAdoptNumResult > 0 Then
				sHTML = sHTML & dbAdoptNumResult
			Else
				sHTML = sHTML & "-&nbsp;"
			End If
			sHTML = sHTML & "/" & RoundUp(dbAdoptNumPlanPeriod,0) & "名"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td style=""text-align:right;"">"
			If dbUnitCost > 0 Then
				sHTML = sHTML & FormatCurrency(Round(dbUnitCost)) & "/名"
			Else
				sHTML = sHTML & "-"
			End If
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>" & dbMedName & "</td>"
			sHTML = sHTML & "</tr>"

			idx = idx + 1
			rRS.MoveNext
		Loop

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th colspan=""2"" style=""border-top:2px solid #cecfff;"">全部門・店舗の合計・平均</th>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">" & FormatCurrency(Round(iCost)) & "</td>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">" & iAdoptNum & "/" & RoundUp(iAdoptNumPlan,0) & "名</td>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">"
		If iUnitCost > 0 Then
			sHTML = sHTML & FormatCurrency(Round(iUnitCost)) & "/名"
		Else
			sHTML = sHTML & "-"
		End If
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<th style=""border-top:2px solid #cecfff;""></th>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"

		rRS.MoveFirst
	End If

	htmlCostPerformance_Branch = sHTML
End Function
%>
