<%
'**********************************************************************************************************************
'概　要：メール詳細ページ /staff/mailhistory_person_entity.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　メール詳細ページ出力用　■■■
'　　　：DspNoticeMailLink
'　　　：DspMailReturnBtn	：返信ボタン出力
'　　　：DspMailDetail		：メール詳細を出力
'　　　：DspNoMailDetail	：メールが無い場合の文言出力
'　　　：■■■　メール詳細ページ更新系　■■■
'　　　：RegNoticeScoutMailUnRead_OpenFlag	：未開封通知メールログの開封フラグを立てる処理
'**********************************************************************************************************************

'******************************************************************************
'概　要：スケジュール通知サービスリンクを出力
'引　数：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vAnswerNG	：返信可否 ["0"]返信可 ["1"]返信不可
'備　考：
'使用元：/staff/maildetail_person.asp
'更　新：2007/03/02 LIS K.Kokubo 
'******************************************************************************
Function DspNoticeMailLink(ByVal vMode, ByVal vAnswerNG)
	If vMode <> "1" And vAnswerNG = "0" Then
		Response.Write "<div style=""padding:0px;"">"
		Response.Write "<input type=""button"" value=""スケジュール通知サービスに登録"" style=""width:190px; color:#aa3300;"" onclick=""window.open('" & HTTPS_NAVI_CURRENTURL & "staff/notification_mail_service.asp?popup=1','notification_mail_service','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=630');return false;"">"
		Response.Write "<span style=""font-size:10px;"">...面接などの日程が決まったら、スケジュール通知サービスに登録して忘却防止！</span>"
		Response.Write "</div>"
		Response.Write "<hr size=""1"">"
	End If
End Function

'******************************************************************************
'概　要：返信ボタンを出力
'引　数：rRS		：メール詳細のレコードセット(up_GetDetailMail)
'　　　：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vAnswerNG	：返信可否 ["0"]返信可 ["1"]返信不可
'備　考：
'使用元：ナビ/staff/maildetail_person_entity.asp
'履　歴：2007/03/02 LIS K.Kokubo 作成
'　　　：2009/03/27 LIS K.Kokubo up_ChkScoutMailにReceiveFlagを追加した対応
'******************************************************************************
Function DspMailReturnBtn(ByRef rRS, ByVal vMode, ByVal vAnswerNG)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbID
	Dim dbScoutMailFlag
	Dim dbReceiveFlag

	If GetRSState(rRS) = False Then Exit Function

	dbID = rRS.Collect("ID")
	dbScoutMailFlag = "0"
	dbReceiveFlag = "0"

	sSQL = "EXEC up_ChkScoutMail '" & dbID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbScoutMailFlag = ChkStr(oRS.Collect("ScoutMailFlag"))
		dbReceiveFlag = ChkStr(oRS.Collect("ReceiveFlag"))
	End If
	Call RSClose(oRS)

	If vMode <> "1" Then
		If vAnswerNG = "0" Then
			Response.Write "<div align=""center"" style=""padding:5px 0px;"">"
			If dbScoutMailFlag = "0" Or dbReceiveFlag = "1" Then
				Response.Write "<input type=""button"" value=""返　信"" onclick=""SendAnswer();""><br><span style=""font-size:10px;"">★企業へメールを返信します。</span>"
			Else
				Response.Write "<fieldset style=""width:175px; float:left;"">"
				Response.Write "<legend>メール本文を入力して返信</legend>"
				Response.Write "<input type=""button"" value=""返　信"" onclick=""SendAnswer();""><br>"
				Response.Write "<p class=""m0"" style=""font-size:10px;"">★メールを返信します。(応募,質問など)</p>"
				Response.Write "</fieldset>" & vbCrLf

				Response.Write "<fieldset style=""width:375px; float:right;"">"
				Response.Write "<legend>カンタン返信(メールの作成は不要)</legend>"
				Response.Write "<div style=""float:left; width:50%;"">"
				Response.Write "<form action=""mail_detail_person.asp?id=" & dbID & """ method=""post"" onsubmit=""return confirm('応募を保留（検討）しますか？');"">"
				Response.Write "<input name=""frmmailtype"" type=""hidden"" value=""1"">"
				Response.Write "<input type=""submit"" value=""保　留""><br>"
				Response.Write "<p class=""m0"" style=""font-size:10px;"">★応募を保留(検討)します。[<span style=""color:#0045f9; cursor:pointer;"" onclick=""window.open('/infomation/horyubutton.asp','autologin','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=220');""><u>？</u></span>]</p>"
				Response.Write "</form>"
				Response.Write "</div>" & vbCrLf

				Response.Write "<div style=""float:left; width:50%;"">"
				Response.Write "<form action=""mail_detail_person.asp?id=" & dbID & """ method=""post"" onsubmit=""return confirm('スカウトを辞退しますか？');"">"
				Response.Write "<input name=""frmmailtype"" type=""hidden"" value=""2"">"
				Response.Write "<input type=""submit"" value=""辞　退""><br>"
				Response.Write "<p class=""m0"" style=""font-size:10px;"">★スカウトを辞退します。</p>"
				Response.Write "</form>"
				Response.Write "</div>" & vbCrLf
				Response.Write "<div style=""clear:both;""></div>"
				Response.Write "</fieldset>"
			End If
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "</div>"
		Else
			Response.Write "<b>この求人票は掲載を終了しており、連絡を取ることができません。</b><br><br>"
		End If
	End If
End Function

'******************************************************************************
'概　要：メール詳細を出力
'引　数：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vAnswerNG	：返信可否 ["0"]返信可 ["1"]返信不可
'備　考：
'使用元：ナビ/staff/maildetail_person_entity.asp
'更　新：2007/03/02 LIS K.kokubo 作成
'　　　：2008/05/08 LIS K.Kokubo ロジック変更
'　　　：2008/08/20 Lis 林 特徴フラグの追加とフレックス移動
'******************************************************************************
Function DspMailDetail(ByRef rRS, ByVal vMode, ByVal vAnswerNG)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode	'情報コード
	Dim dbCompanyCode	'企業コード
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName

	Dim sJobTypeDetail
	Dim sAccessCount
	Dim sUpdateDay
	Dim sInexperiencedPersonFlag
	Dim sUITurnFlag
	Dim sUtilizeLanguageFlag
	Dim sManyHolidayFlag
	Dim sFlexTimeFlag
	'**TOP 08/08/19 Lis林 ADD
	Dim sNearStationFlag,sNoSmokingFlag,sNewlyBuiltFlag,sLandmarkFlag
	Dim sRenovationFlag,sDesignersFlag,sCompanyCafeteriaFlag,sShortOvertimeFlag,sMaternityFlag
	Dim sDressFreeFlag,sMammyFlag,sFixedTimeFlag,sShortTimeFlag,sHandicappedFlag
	'**BTM 08/08/19 Lis林 ADD
	Dim sOrderType
	Dim sCompanyKbn
	Dim sImgOrderState
	Dim sWorkingPlacePrefectureCode
	Dim sWorkingcode
	Dim sWorkingname
	Dim sWorkingPlacePrefectureName

	Dim dbSecretFlag
	Dim sYearlyIncomeMin
	Dim sYearlyIncomeMax
	Dim sMonthlyIncomeMin
	Dim sMonthlyIncomeMax
	Dim sDailyIncomeMin
	Dim sDailyIncomeMax
	Dim sDailyIncome
	Dim sHourlyIncomeMin
	Dim sHourlyIncomeMax
	Dim sYearlyIncome
	Dim sMonthlyIncome
	Dim sHourlyIncome
	Dim sImgMain
	Dim sImgSub
	Dim sCompanyPictureFlag
	Dim flgImg
	Dim idx

	'具体的職種名の取得
	dbOrderCode = rRS.Collect("OrderCode")
	sSQL = "select OrderType,JobTypeDetail,AccessCount,UpdateDay,Companycode from C_info where ordercode = '" & dbOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sJobTypeDetail = oRS.Collect("JobTypeDetail")
		sAccessCount = oRS.Collect("AccessCount")
		sUpdateDay = oRS.Collect("UpdateDay")
		dbCompanyCode = oRS.Collect("Companycode")
		sOrderType = oRS.Collect("OrderType")
	End if

	'**TOP 08/08/19 Lis林 REP
	'sSQL = "select InexperiencedPersonFlag,UITurnFlag,UtilizeLanguageFlag,ManyHolidayFlag from C_SupplementInfo where ordercode = '" & dbOrderCode & "'"
	sSQL = "select InexperiencedPersonFlag,UITurnFlag,UtilizeLanguageFlag,ManyHolidayFlag"
	sSQL = sSQL & ",FlexTimeFlag,NearStationFlag,NoSmokingFlag,NewlyBuiltFlag,LandmarkFlag"
	sSQL = sSQL & ",RenovationFlag,DesignersFlag,CompanyCafeteriaFlag,ShortOvertimeFlag,MaternityFlag"
	sSQL = sSQL & ",DressFreeFlag,MammyFlag,FixedTimeFlag,ShortTimeFlag,HandicappedFlag"
	sSQL = sSQL & " from C_SupplementInfo where ordercode = '" & dbOrderCode & "'"
	'**BTM 08/08/19 Lis林 REP
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sInexperiencedPersonFlag = oRS.Collect("InexperiencedPersonFlag")
		sUITurnFlag = oRS.Collect("UITurnFlag")
		sUtilizeLanguageFlag = oRS.Collect("UtilizeLanguageFlag")
		sManyHolidayFlag = oRS.Collect("ManyHolidayFlag")
		'**TOP 08/08/19 Lis林 REP
		sFlexTimeFlag = oRS.Collect("FlexTimeFlag")
		sNearStationFlag = oRS.Collect("NearStationFlag")
		sNoSmokingFlag = oRS.Collect("NoSmokingFlag")
		sNewlyBuiltFlag = oRS.Collect("NewlyBuiltFlag")
		sLandmarkFlag = oRS.Collect("LandmarkFlag")
		sRenovationFlag = oRS.Collect("RenovationFlag")
		sDesignersFlag = oRS.Collect("DesignersFlag")
		sCompanyCafeteriaFlag = oRS.Collect("CompanyCafeteriaFlag")
		sShortOvertimeFlag = oRS.Collect("ShortOvertimeFlag")
		sMaternityFlag = oRS.Collect("MaternityFlag")
		sDressFreeFlag = oRS.Collect("DressFreeFlag")
		sMammyFlag = oRS.Collect("MammyFlag")
		sFixedTimeFlag = oRS.Collect("FixedTimeFlag")
		sShortTimeFlag = oRS.Collect("ShortTimeFlag")
		sHandicappedFlag = oRS.Collect("HandicappedFlag")
		'**BTM 08/08/19 Lis林 REP
	End if

	'**TOP 08/08/19 Lis林 REP
	'sSQL = "select FlexTime,CompanyKbn from Companyinfo where companycode = '" & dbCompanyCode & "'"
	sSQL = "select CompanyKbn from Companyinfo where companycode = '" & dbCompanyCode & "'"
	'**BTM 08/08/19 Lis林 REP
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		'sFlexTimeFlag = oRS.Collect("FlexTime")		'08/08/19 Lis林 DEL
		sCompanyKbn = oRS.Collect("CompanyKbn")
	End if


	sSQL = "up_DtlOrder '" & rRS.Collect("OrderCode") & "', ''"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbSecretFlag = oRS.Collect("SecretFlag")

		'******************************************************************************
		'給与 start　10月1日一覧変更用に表示追加_新名
		'------------------------------------------------------------------------------

		'年収
		If GetRSState(oRS) = True Then
			sYearlyIncomeMin = ChkStr(oRS.Collect("YearlyIncomeMin"))
			sYearlyIncomeMax = ChkStr(oRS.Collect("YearlyIncomeMax"))
			If sYearlyIncomeMin = "0" Then sYearlyIncomeMin = ""
			If sYearlyIncomeMax = "0" Then sYearlyIncomeMax = ""
			If sYearlyIncomeMin <> "" Then sYearlyIncomeMin = GetJapaneseYen(sYearlyIncomeMin)
			If sYearlyIncomeMax <> "" Then sYearlyIncomeMax = GetJapaneseYen(sYearlyIncomeMax)
			If sYearlyIncomeMin & sYearlyIncomeMax <> "" Then
				If sYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMin
				sYearlyIncome = sYearlyIncome & "&nbsp;～&nbsp;"
				If sYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMax
			End If
			'月給
			sMonthlyIncomeMin = ChkStr(oRS.Collect("MonthlyIncomeMin"))
			sMonthlyIncomeMax = ChkStr(oRS.Collect("MonthlyIncomeMax"))
			If sMonthlyIncomeMin = "0" Then sMonthlyIncomeMin = ""
			If sMonthlyIncomeMax = "0" Then sMonthlyIncomeMax = ""
			If sMonthlyIncomeMin <> "" Then sMonthlyIncomeMin = GetJapaneseYen(sMonthlyIncomeMin)
			If sMonthlyIncomeMax <> "" Then sMonthlyIncomeMax = GetJapaneseYen(sMonthlyIncomeMax)
			If sMonthlyIncomeMin & sMonthlyIncomeMax <> "" Then
				If sMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMin
				sMonthlyIncome = sMonthlyIncome & "&nbsp;～&nbsp;"
				If sMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMax
			End If
			'日給
			sDailyIncomeMin = ChkStr(oRS.Collect("DailyIncomeMin"))
			sDailyIncomeMax = ChkStr(oRS.Collect("DailyIncomeMax"))
			If sDailyIncomeMin = "0" Then sDailyIncomeMin = ""
			If sDailyIncomeMax = "0" Then sDailyIncomeMax = ""
			If sDailyIncomeMin <> "" Then sDailyIncomeMin = GetJapaneseYen(sDailyIncomeMin)
			If sDailyIncomeMax <> "" Then sDailyIncomeMax = GetJapaneseYen(sDailyIncomeMax)
			If sDailyIncomeMin & sDailyIncomeMax <> "" Then
				If sDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMin
				sDailyIncome = sDailyIncome & "&nbsp;～&nbsp;"
				If sDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMax
			End If
			'時給
			sHourlyIncomeMin = ChkStr(oRS.Collect("HourlyIncomeMin"))
			sHourlyIncomeMax = ChkStr(oRS.Collect("HourlyIncomeMax"))
			If sHourlyIncomeMin = "0" Then sHourlyIncomeMin = ""
			If sHourlyIncomeMax = "0" Then sHourlyIncomeMax = ""
			If sHourlyIncomeMin <> "" Then sHourlyIncomeMin = GetJapaneseYen(sHourlyIncomeMin)
			If sHourlyIncomeMax <> "" Then sHourlyIncomeMax = GetJapaneseYen(sHourlyIncomeMax)
			If sHourlyIncomeMin & sHourlyIncomeMax <> "" Then
				If sHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMin
				sHourlyIncome = sHourlyIncome & "&nbsp;～&nbsp;"
				If sHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMax
			End If
		End If

		'------------------------------------------------------------------------------
		'給与 end
		'******************************************************************************

		'**************************************************************************
		'画像 start
		'--------------------------------------------------------------------------
		sSQL = "select ordercode,CASE WHEN ISNULL(CONVERT(VARBINARY, OP.Picture), 0x00) > 0x00 THEN '1' ELSE '0' END AS CompanyPictureFlag from c_info as CI LEFT JOIN OptionPicture AS OP ON CI.CompanyCode = OP.CompanyCode AND OptionNo = 1 Where CI.ordercode='" & dbOrderCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sCompanyPictureFlag = ChkStr(oRS.Collect("CompanyPictureFlag"))
		End If

		flgImg = False
		sImgMain = ""
		sImgSub = ""

		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("OptionNo1")) <> "" Or (sOrderType = "0" And sCompanyPictureFlag = "1") Then
				sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo1") & """ alt="""" border=""0"" width=""100"" height=""75"" style=""float:left; margin-right:5px;"">"
				flgImg = True
			End If
			If sImgSub <> "" Then sImgSub = sImgSub & "<div style=""clear:both;""></div>"
		Else
			If sCompanyPictureFlag = "1" And sOrderType = "0" Then
				sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"" alt="""" border=""0"" width=""100"" height=""75"">"
				flgImg = True
			End If
		End If

		Call RSClose(oRS)
		'--------------------------------------------------------------------------
		'画像 end
		'**************************************************************************

		'**************************************************************************
		'状態img start
		'--------------------------------------------------------------------------
		sImgOrderState = ""
		'アクセス数が100を超えていれば「HOT」表示（リス安藤）
		If sAccessCount > 100 Then
			sImgOrderState = sImgOrderState & "<img src=""/img/c_HOT_green.gif"" alt=""人気"">&nbsp;"
		End If

		'UPDATEと今日から10日引いた日で「新着」表示(リス安藤)
		If sUpdateDay > NOW()-10 Then
			sImgOrderState = sImgOrderState & "<img src=""/img/c_NEW_green.gif"" alt=""新着"">&nbsp;"
		End If

		'未経験者ＯＫの場合、わかばマーク表示(リス安藤)
		If sInexperiencedPersonFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/no_experience.gif"" alt=""未経験者／第二新卒歓迎"">&nbsp;"
		End If

		'Ｕターン・Ｉターン
		If sUITurnFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/ui_turn.gif"" alt=""Ｕターン・Ｉターン"">&nbsp;"
		End If

		'語学を活かす仕事
		If sUtilizeLanguageFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/linguistic_job.gif"" alt=""語学を活かす仕事"">&nbsp;"
		End If

		'年間休日120日以上
		If sManyHolidayFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/year_holidaycnt.gif"" alt=""年間休日120日以上"">&nbsp;"
		End If

		'フレックスタイム制度あり ------2006/01/10 Hayashi ADD
		'**TOP 08/08/19 Lis林 REP
		'If sFlexTimeFlag = "ON" And sOrderType = "0" And sCompanyKbn = "1" Then
		'	sImgOrderState = sImgOrderState & "<img src=""/img/flextime.gif"" alt=""フレックスタイム制度あり"">&nbsp;"
		'End If
		If sFlexTimeFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_flextime.gif"" alt=""フレックスタイム制度あり"">&nbsp;"
		End If
		if sNearStationFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_nearstation.gif"" alt=""駅近(徒歩5分以内)"">&nbsp;"
		end if
		if sNoSmokingFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_nosmoking.gif"" alt=""禁煙・分煙"">&nbsp;"
		end if
		if sNewlyBuiltFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_newlybuilt.gif"" alt=""新築ビル・オフィス(5年以内)"">&nbsp;"
		end if
		if sLandmarkFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_landmark.gif"" alt=""高層(15階以上)ビル"">&nbsp;"
		end if
		if sRenovationFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_renovation.gif"" alt=""リノベーションビル・オフィス(5年以内)"">&nbsp;"
		end if
		if sDesignersFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_designers.gif"" alt=""デザイナーズビル・オフィス"">&nbsp;"
		end if
		if sCompanyCafeteriaFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_companycafeteria.gif"" alt=""社員食堂"">&nbsp;"
		end if
		if sShortOvertimeFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_shortovertime.gif"" alt=""残業10h/月以内"">&nbsp;"
		end if
		if sMaternityFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_maternity.gif"" alt=""産休・育休実績あり"">&nbsp;"
		end if
		if sDressFreeFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_dressfree.gif"" alt=""服装自由"">&nbsp;"
		end if
		if sMammyFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_mammy.gif"" alt=""子育てママ歓迎"">&nbsp;"
		end if
		if sFixedTimeFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_fixedtime.gif"" alt=""18時までに退社"">&nbsp;"
		end if
		if sShortTimeFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_shorttime.gif"" alt=""1日6時間以内労働"">&nbsp;"
		end if
		if sHandicappedFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_handicapped.gif"" alt=""障害者歓迎"">&nbsp;"
		end if
		'**BTM 08/08/19 Lis林 REP

		'シークレット求人 ------2008/05/08 LIS K.Kokubo ADD
		If dbSecretFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/order/secret.gif"" alt=""スカウトを受けた人だけが閲覧できる求人情報"" width=""50"" height=""15"">&nbsp;"
		End If

		'<希望勤務形態アイコン　10月1日一覧変更用に表示追加_新名>
		sSQL = "sp_GetDataWorkingType '" & dbOrderCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			sWorkingcode = oRS.Collect("WorkingTypecode")
			sWorkingname = GetDetail("WorkingType", sWorkingcode)

			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_w" & sWorkingcode & ".gif"" alt=""" & sWorkingname & """ width=""50"" height=""15"">&nbsp;"

			oRS.MoveNext
		Loop
		Call RSClose(oRS)
		'</希望勤務形態アイコン　10月1日一覧変更用に表示追加_新名>

		'<勤務地アイコン>
		idx = 0
		sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True And idx < 3
			dbWorkingPlacePrefectureCode = ChkStr(oRS.Collect("WorkingPlacePrefectureCode"))
			dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15"">&nbsp;"

			oRS.MoveNext
			idx = idx + 1
		Loop
		Call RSClose(oRS)
		'</勤務地アイコン>
		'--------------------------------------------------------------------------
		'状態img end
		'**************************************************************************
	End If

	Response.Write "<table class=""pattern1"" border=""0"" style=""width:600px; table-layout:fixed;"">" & vbCrLf
	Response.Write "<thead>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write "<th style=""font-size:16px; text-align:left; padding:4px 0px 2px 10px;""><span style=""color:#66cc33;"">■ </span>" & rRS.Collect("Subject") & "</th>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "</thead>" & vbCrLf
	Response.Write "<tbody>" & vbCrLf
	Response.Write "<tr>" & vbCrLf

	If vMode = "1" Then
		'送信画面の場合は宛先
	Else
		'開封済みにする
		sSQL = "sp_Reg_MailOpenDay '" & rRS.Collect("ID") & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		'受信画面の場合は差出人
	End If

	Response.Write "<td>"

	If sCompanyPictureFlag <> "" Then
		If sAnswerNG = "0" And Trim(dbOrderCode) <> "" Then
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "order/order_detail.asp?ordercode=" & rRS.Collect("OrderCode") & """ title=""" & sJobTypeDetail & """>" & sImgMain & "</a>"
		ElseIf sAnswerNG = "1" Then
			Response.Write sImgMain
		End If
	End If

	Response.Write rRS.Collect("CompanyName")
	If sAnswerNG = "0" And dbOrderCode <> "" Then
		Response.Write "<br>[<a href=""" & HTTPS_CURRENTURL & "order/order_detail.asp?OrderCode=" & dbOrderCode & """>" & sJobTypeDetail & "（" & dbOrderCode & "）</a>]"
	ElseIf sAnswerNG = "1" Then
		Response.Write "<br>" & rRS.Collect("OrderCode")
	End If

	If sImgOrderState <> "" Then
		Response.Write "<div style=""margin-top:5px;width:480px;"">" & sImgOrderState & "</div>"
	End if

	If sYearlyIncome <> "" Then Response.Write "年収【" & sYearlyIncome & "】"
	If sMonthlyIncome <> "" Then Response.Write "月給【" & sMonthlyIncome & "】"
	If sDailyIncome <> "" Then Response.Write "日給【" & sDailyIncome & "】"
	If sHourlyIncome <> "" Then Response.Write "時給【" & sHourlyIncome & "】"

	Response.Write "</td>"
	Response.Write "</tr>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write "<td>" & vbCrLf
	Response.Write "<div readonly style=""border:solid 1px #cccccc; overflow:visible; width:576px; background-color:#fffff5; padding:5px;"">" & vbCrLf
	Response.Write Replace(Replace(Replace(rRS.Collect("Body"),vbCrLf,"<br>"), vbCr, "<br>"), vbLf, "<br>") & vbCrLf
	Response.Write "</div>" & vbCrLf
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "</tbody>" & vbCrLf
	Response.Write "</table>" & vbCrLf
	Response.Write "<br>" & vbCrLf
End Function

'******************************************************************************
'概　要：メールが無い場合の文言出力
'引　数：
'備　考：
'使用元：/staff/maildetail_person.asp
'更　新：2007/03/02 LIS K.Kokubo 
'******************************************************************************
Function DspNoMailDetail()
	Response.Write "<b>指定されたメールは存在しないか、削除されています。</b><br><br>"
End Function

'******************************************************************************
'概　要：未開封通知メールログの開封フラグを立てる処理
'引　数：rDB	：接続中ＤＢコネクション
'　　　：rRS	：メール詳細のレコードセット(up_GetDetailMail)
'備　考：
'使用元：/staff/maildetail_person.asp
'更　新：2007/03/02 LIS K.Kokubo 
'******************************************************************************
Function RegNoticeScoutMailUnRead_OpenFlag(ByRef rDB, ByRef rRS, ByVal vMode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbMailID

	If GetRSState(rRS) = False Then Exit Function
	If vMode = "1" Then Exit Function

	dbMailID = rRS.Collect("ID")

	sSQL = "/*ナビ：未開封通知メールログの開封フラグを立てる*/" & vbCrLf
	sSQL = sSQL & "EXEC up_UpdLOG_Notice_ScoutMailUnRead_OpenFlag '" & dbMailID & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Call RSClose(oRS)
End Function
%>
