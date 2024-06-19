<%
'******************************************************************************
'概　要：求人票の特徴
'引　数：rDB		：
'　　　：rRS		：
'　　　：vHTMLType	：
'戻り値：
'備　考：
'履　歴：2011/09/21 LIS K.Kokubo 作成
'******************************************************************************
Function htmlOrderSpecialityImg(ByRef rDB, ByRef rRS, ByVal vHTMLType)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName

	Dim sHTML
	Dim sSlash
	Dim sWorkingCode

	If GetRSState(rRS) = False Then Exit Function

	If LCase(vHTMLType) = "xhtml" Then sSlash = " /"

	dbOrderCode = rRS.Collect("OrderCode")

	sHTML = ""
	'アクセス数が100を超えていれば「HOT」表示（リス安藤）
	If rRS.Collect("AccessCount") > 100 Then sHTML = sHTML & "<img src=""/img/c_HOT_green.gif"" alt=""人気"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'UPDATEと今日から10日引いた日で「新着」表示(リス安藤)
	If rRS.Collect("Updateday") > NOW()-10 Then sHTML = sHTML & "<img src=""/img/c_NEW_green.gif"" alt=""新着"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'未経験者ＯＫの場合、わかばマーク表示(リス安藤)
	If rRS.Collect("InexperiencedPersonFlag") = "1" Then sHTML = sHTML & "<img src=""/img/no_experience.gif"" alt=""未経験者／第二新卒歓迎"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'Ｕターン・Ｉターン
	If rRS.Collect("UITurnFlag") = "1" Then sHTML = sHTML & "<img src=""/img/ui_turn.gif"" alt=""Ｕターン・Ｉターン"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'語学を活かす仕事
	If rRS.Collect("UtilizeLanguageFlag") = "1" Then sHTML = sHTML & "<img src=""/img/linguistic_job.gif"" alt=""語学を活かす仕事"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'年間休日120日以上
	If rRS.Collect("ManyHolidayFlag") = "1" Then sHTML = sHTML & "<img src=""/img/year_holidaycnt.gif"" alt=""年間休日120日以上"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2006/01/10 M.Hayashi ADD フレックスタイム制度あり
	If rRS.Collect("FlexTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_flextime.gif"" alt=""フレックスタイム制度あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NearStationFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_nearstation.gif"" alt=""駅近(徒歩5分以内)"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NoSmokingFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_nosmoking.gif"" alt=""禁煙・分煙"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NewlyBuiltFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_newlybuilt.gif"" alt=""新築ビル・オフィス(5年以内)"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("LandmarkFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_landmark.gif"" alt=""高層(15階以上)ビル"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("RenovationFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_renovation.gif"" alt=""リノベーションビル・オフィス(5年以内)"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DesignersFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_designers.gif"" alt=""デザイナーズビル・オフィス"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("CompanyCafeteriaFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_companycafeteria.gif"" alt=""社員食堂"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortOvertimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_shortovertime.gif"" alt=""残業10h/月以内"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MaternityFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_maternity.gif"" alt=""産休・育休実績あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DressFreeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_dressfree.gif"" alt=""服装自由"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MammyFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mammy.gif"" alt=""子育てママ歓迎"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("FixedTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_fixedtime.gif"" alt=""18時までに退社"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_shorttime.gif"" alt=""1日6時間以内労働"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("HandicappedFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_handicapped.gif"" alt=""障害者歓迎"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentAllFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_rentallflag.gif"" alt=""住宅費用全額補助あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentPartFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_rentpartflag.gif"" alt=""住宅費用一部補助あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mealsflag.gif"" alt=""食事・賄い付き案件"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsAssistanceFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mealsassistanceflag.gif"" alt=""食事補助制度あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("TrainingCostFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_trainingcostflag.gif"" alt=""研修費助成制度あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("EntrepreneurCostFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_entrepreneurcostflag.gif"" alt=""起業機材補助制度あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MoneyFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_moneyflag.gif"" alt=""無利子・低利子補助制度あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("LandShopFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_landshopflag.gif"" alt=""土地・店舗等提供制度あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("FindJobFestiveFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_findjobfestiveflag.gif"" alt=""就職お祝い金制度あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("AppointmentFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_appointmentflag.gif"" alt=""正社員登用制度あり"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_societyinsuranceflag.gif"" alt=""社保完備"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/05/08 LIS K.Kokubo ADD シークレット求人
	If rRS.Collect("SecretFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order/secret.gif"" alt=""スカウトを受けた人だけが閲覧できる求人情報"" width=""50"" height=""15""" & sSlash & ">&nbsp;"

	'直接Yahoo!の検索からお仕事情報詳細ページへ来る人へアイコン表示
	If InStr(Request.ServerVariables("HTTP_REFERER"),"search.yahoo.co.jp/") <> 0 Then
		sSQL = "sp_GetDataWorkingType '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			sWorkingcode = oRS.Collect("WorkingTypecode")

			sHTML = sHTML & "<img src=""/img/order_detail_icon/icon_w" & sWorkingcode & ".gif"" alt=""派遣社員"" width=""50"" height=""15""" & sSlash & ">&nbsp;"

			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		'<勤務地>
		sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbWorkingPlacePrefectureCode = ChkStr(oRS.Collect("WorkingPlacePrefectureCode"))
			dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
			If InStr(sHTML, "/icon_p" & dbWorkingPlacePrefectureCode & ".gif") = 0 Then
				'同じ都道府県アイコンは出さない！
				sHTML = sHTML & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15""" & sSlash & ">&nbsp;"
			End If
		End If
		Call RSClose(oRS)
		'</勤務地>
	End If

	htmlOrderSpecialityImg = sHTML
End Function
%>
