<%
'*******************************************************************************
'概　要：採用改善サポートシステムの就業形態別費用対効果TABLEを取得
'引　数：rDB	：接続中DBコネクション
'　　　：vUserID：ログイン中ユーザID
'　　　：vYM1	：集計期間下限年月
'　　　：vYM2	：集計期間上限年月
'戻り値：String
'備　考：
'履　歴：2010/03/10 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlCostPerformance_WorkingType(ByRef rDB, ByVal vUserID, ByVal vYM1, ByVal vYM2)
	Dim sSQL
	Dim oRS
	Dim sSQLErr
	Dim flgQE
	'DB
	Dim dbWorkingTypeName
	Dim dbCost
	Dim dbAdoptNumPlan
	Dim dbAdoptNum
	Dim dbUnitCost
	Dim dbAdoptNumPlanPeriod

	Dim iCost
	Dim iAdoptNumPlan
	Dim iAdoptNum
	Dim iUnitCost

	Dim sHTML

	iCost = 0
	iAdoptNumPlan = 0
	iAdoptNum = 0
	iUnitCost = 0

	sSQL = "EXEC up_LstCMPCostPerformance_CompanyAll '" & vUserID & "','" & vYM1 & "','" & vYM2 & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
	If GetRSState(oRS) = True = True Then
		Set oRS.ActiveConnection = Nothing

		sHTML = sHTML & "<table class=""pattern6"" border=""0"" style=""width:100%;"">"
		sHTML = sHTML & "<colgroup>"
		sHTML = sHTML & "<col style=""width:25%;"">"
		sHTML = sHTML & "<col style=""width:25%;"">"
		sHTML = sHTML & "<col style=""width:25%;"">"
		sHTML = sHTML & "<col style=""width:25%;"">"
		sHTML = sHTML & "</colgroup>"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th></th>"
		sHTML = sHTML & "<th style=""text-align:center;"">採用ｺｽﾄ総額</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">採用人数(実績/計画)</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">1名採用ｺｽﾄ</th>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</thead>"
		sHTML = sHTML & "<tbody>"

		Do While GetRSState(oRS) = True
			dbWorkingTypeName = oRS.Collect("WorkingTypeName")
			dbCost = oRS.Collect("Cost")
			dbAdoptNumPlan = oRS.Collect("AdoptNumPlan")
			dbAdoptNum = oRS.Collect("AdoptNum")
			dbUnitCost = oRS.Collect("UnitCost")
			dbAdoptNumPlanPeriod = oRS.Collect("AdoptNumPlanPeriod")

			iCost = iCost + dbCost
			iAdoptNumPlan = iAdoptNumPlan + dbAdoptNumPlanPeriod
			iAdoptNum = iAdoptNum + dbAdoptNum

			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>" & dbWorkingTypeName & "</td>"
			sHTML = sHTML & "<td style=""text-align:right;"">" & FormatCurrency(dbCost) & "</td>"
			sHTML = sHTML & "<td style=""text-align:right;"">"
			If dbAdoptNum > 0 Then
				sHTML = sHTML & dbAdoptNum
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
			sHTML = sHTML & "</tr>"

			oRS.MoveNext
		Loop

		If iAdoptNum > 0 Then iUnitCost = iCost / iAdoptNum

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border-top:2px solid #cecfff;"">合計・平均</th>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">" & FormatCurrency(iCost) & "</td>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">" & iAdoptNum & "/" & RoundUp(iAdoptNumPlan,0) & "名</td>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">"
		If iUnitCost > 0 Then
			sHTML = sHTML & FormatCurrency(Round(iUnitCost)) & "/名"
		Else
			sHTML = sHTML & "-"
		End If
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
	End If
	Call RSClose(oRS)

	htmlCostPerformance_WorkingType = sHTML
End Function
%>
