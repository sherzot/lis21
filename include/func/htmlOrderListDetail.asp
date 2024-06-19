<%
'******************************************************************************
'概　要：求人票一覧ページの各求人票項目を表示
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_SearchOrder or 求人票詳細検索SQL で生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder		：利用中ユーザの自社求人票か否か ["1"]自社求人票 ["0"]自社求人票でない
'　　　：vHTMLType		：[xhtml]XHTML形式 [html]HTML形式
'使用元：/rss/job.asp
'備　考：
'履　歴：2011/09/21 LIS K.Kokubo 作成
'******************************************************************************
Function htmlOrderListDetail(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder, ByVal vHTMLType)
	Const PICSIZEW = 240
	Const PICSIZEH = 180
	Const PICSIZESUBW = 72
	Const PICSIZESUBH = 56

	Dim sHTML
	Dim sSlash

	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'情報コード
	Dim dbCompanyCode		'企業コード
	Dim sOrderType			'受注種類
	Dim sPlanType			'ライセンスプラン種類
	Dim iImageLimit			'写真掲載数制限数
	Dim sTitleJobName		'職種
	Dim sTitleCompanyName	'会社名
	Dim sImgMail			'送信済みメール画像
	Dim sImgOrderState		'状態画像 HOT,新着,未経験OK,UIターン,語学,休日120日,フレックス
	Dim sCatchCopy			'キャッチコピー
	Dim flgImg				'画像の有無フラグ(画像の有無でレイアウトが変化) [True]有 [False]無
	Dim sImgMain			'大きい画像
	Dim sImgSub				'小さい画像
	Dim sImg1,sImg2,sImg3,sImg4	'画像URL
	Dim sBusinessDetail		'担当業務
	Dim sWorkingType		'勤務形態
	Dim sWorkingPlace		'勤務地 都道府県+市区郡
	Dim sProgress			'求人票審査状況
	Dim sPublicDay			'掲載日
	Dim sPublicListDsp		'掲載非掲載 リストボックス表示スタイル [style="display:none;"]
	Dim sPublicFlag1		'掲載 selected
	Dim sPublicFlag0		'非掲載 selected
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sRegistDay			'登録日
	Dim sPublishLimitStr	'求人票掲載終了日
	Dim sStationName		'駅名
	Dim sYearlyIncomeMin	'年収下限
	Dim sYearlyIncomeMax	'年収上限
	Dim sMonthlyIncomeMin	'月給下限
	Dim sMonthlyIncomeMax	'月給上限
	Dim sDailyIncomeMin		'月給下限
	Dim sDailyIncomeMax		'月給上限
	Dim sHourlyIncomeMin	'時給下限
	Dim sHourlyIncomeMax	'時給上限
	Dim dbTopInterviewFlag	'トップインタビューフラグ
	Dim dbWValueURL			'ＷバリューのＵＲＬ

	Dim sYearlyIncome		'年収表示用
	Dim sDailyIncome		'月給表示用
	Dim sMonthlyIncome		'日給表示用
	Dim sHourlyIncome		'時給表示用
	'希望勤務形態・希望勤務地アイコン　10月1日一覧変更用に表示追加_新名
	Dim sWorkingCode
	Dim sWorkingName
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName
	Dim dbWorkingPlaceCity
	Dim sBiz
	Dim sBizName1
	Dim sBizName2
	Dim sBizName3
	Dim sBizName4
	Dim sBizPercentage1
	Dim sBizPercentage2
	Dim sBizPercentage3
	Dim sBizPercentage4
	Dim flgBusiness
	Dim idx

	If GetRSState(rRS) = False Then Exit Function

	If LCase(vHTMLType) = "xhtml" Then sSlash = " /"

	dbOrderCode = rRS.Collect("OrderCode")

	If G_USEFLAG = "0" And vMyOrder = "1" And G_OLDAPPLICATIONCODE <> "" Then
		sSQL = "EXEC up_DtlOrder '" & rRS.Collect("OrderCode") & "', '" & G_OLDAPPLICATIONCODE & "';"
	Else
		sSQL = "EXEC up_DtlOrder '" & rRS.Collect("OrderCode") & "', '';"
	End If

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	dbCompanyCode = oRS.Collect("CompanyCode")
	sOrderType = ChkStr(oRS.Collect("OrderType"))
	sPlanType = ChkStr(oRS.Collect("PlanTypeName"))
	iImageLimit = oRS.Collect("ImageLimit")

	'**************************************************************************
	'職種／会社名 start
	'--------------------------------------------------------------------------
	sTitleCompanyName = ""
	'STEP1：具体的職種名取得
	If oRS.Collect("JobTypeDetail") <> "" Then
		If Len(oRS.Collect("JobTypeDetail")) >= 50 Then
			sTitleJobName = Left(oRS.Collect("JobTypeDetail"), 50)
		Else
			sTitleJobName = oRS.Collect("JobTypeDetail")
		End If
	End If

	'STEP2：具体的職種名があれば／を追加
	'If sTitleCompanyName <> "" Then sTitleCompanyName = sTitleCompanyName & "／"
	'STEP3：企業名取得
	If oRS.Collect("CompanySpeciality") <>"" THEN 
			sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanySpeciality")
	Else
		If oRS.Collect("Companykbn") ="4" Then
			sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanyName")
		ElseIf oRS.Collect("OrderType") > "0" then
				sTitleCompanyName = sTitleCompanyName & "リス株式会社"
			Else
				sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanyName")
		End If
	End If
	'--------------------------------------------------------------------------
	'職種／会社名 end
	'**************************************************************************

	'******************************************************************************
	'給与 start　10月1日一覧変更用に表示追加_新名
	'------------------------------------------------------------------------------
	'年収
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

	'------------------------------------------------------------------------------
	'給与 end
	'******************************************************************************

	'******************************************************************************
	'最寄駅 start　10月1日一覧変更用に表示追加_新名
	'2008/10/22 LIS K.Kokubo 勤務地複数化により表示量が増える恐れがあるために非表示に。
	'------------------------------------------------------------------------------
	'sStationName = ""
	'sSQL = "sp_GetDataNearbyStation '" & dbOrderCode & "'"
	'flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	'If GetRSState(oRS2) = True Then
	'	sStationName ="【" & sStationName & GetStrNearbyStation(oRS2.Collect("StationName"), "", "") & "】"
	'End If
	'------------------------------------------------------------------------------
	'最寄駅 end
	'******************************************************************************

	'**************************************************************************
	'メール送信済み確認 start
	'--------------------------------------------------------------------------
	If vUserType = "staff" Then
		sSQL = "up_DtlMailHistory_Order '" & vUserID & "', '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			sImgMail = "<img src=""/img/s_contact.gif"" alt=""メール送信済み""" & sSlash & ">"
		End If
		Call RSClose(oRS2)
	End If

	'「求人票をメール送信」のリンクにぶつからないように職種名を削る(2007/08/01 T.Sotome追加)
	If LenByte(sTitleCompanyName) > 72 Then
		sTitleCompanyName = LeftByte(sTitleCompanyName, 70) & "..."
	End If
	'「ウォッチリストへ保存」のリンクにぶつからないように職種名を削る(2007/06/26 T.Sotome追加)
	If sImgMail = "" Then
		If LenByte(sTitleJobName) > 46 Then
			sTitleJobName = LeftByte(sTitleJobName, 44) & "..."
		End If
	Else
		If LenByte(sTitleJobName) > 36 Then
			sTitleJobName = LeftByte(sTitleJobName, 34) & "..."
		End If
	End If

	'--------------------------------------------------------------------------
	'メール送信済み確認 end
	'**************************************************************************

	'**************************************************************************
	'状態img start
	'--------------------------------------------------------------------------
	sImgOrderState = htmlOrderSpecialityImg(rDB, oRS, vHTMLType)
	'--------------------------------------------------------------------------
	'状態img end
	'**************************************************************************

	'**************************************************************************
	'キャッチコピー start
	'--------------------------------------------------------------------------
	sCatchCopy = ""
	sCatchCopy = oRS.Collect("CatchCopy")
	'--------------------------------------------------------------------------
	'キャッチコピー end
	'**************************************************************************

	'**************************************************************************
	'画像 start
	'--------------------------------------------------------------------------
	flgImg = False
	If sOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS2,sSQL,sError)
		If GetRSState(oRS2) = True Then
			If ChkStr(oRS2.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo1")
			End If
			If ChkStr(oRS2.Collect("PicNo2")) <> "" Then
				sImg2 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo2")
			End If
			If ChkStr(oRS2.Collect("PicNo3")) <> "" Then
				sImg3 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo3")
			End If
			If ChkStr(oRS2.Collect("PicNo4")) <> "" Then
				sImg4 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo4")
			End If
		End If
		Call RSClose(oRS2)
	ElseIf iImageLimit > 0 Then
		sCompanyPictureFlag = ChkStr(oRS.Collect("CompanyPictureFlag"))

		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & oRS.Collect("OrderCode") & "', 'orderpicture'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			If ChkStr(oRS2.Collect("OptionNo1")) <> "" Or (sOrderType = "0" And sCompanyPictureFlag = "1") Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo1")
			End If

			If sPlanType = "platinum" Or sPlanType = "old" Or iImageLimit > 1 Then
				If ChkStr(oRS2.Collect("OptionNo2")) <> "" Then
					sImg2 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo2")
				End If
				If ChkStr(oRS2.Collect("OptionNo3")) <> "" Then
					sImg3 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo3")
				End If
				If ChkStr(oRS2.Collect("OptionNo4")) <> "" Then
					sImg4 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo4")
				End If
			End If
		Else
			If sCompanyPictureFlag = "1" And sOrderType = "0" Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If

		Call RSClose(oRS2)
	End If

	If sImg1 & sImg2 & sImg3 & sImg4 <> "" Then flgImg = True

	If sImg1 <> "" Then
		sImgMain = "<img src=""" & sImg1 & """ alt="""" border=""0"" width=""" & PICSIZEW & """ height=""" & PICSIZEH & """" & sSlash & ">"
	End If

	If sImg2 <> "" Then
		sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
			"<img src=""" & sImg2 & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""" & sSlash & "><br" & sSlash & ">"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	If sImg3 <> "" Then
		sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
			"<img src=""" & sImg3 & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""" & sSlash & "><br" & sSlash & ">"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	If sImg4 <> "" Then
		sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
			"<img src=""" & sImg4 & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""" & sSlash & "><br" & sSlash & ">"
		sImgSub = sImgSub & "</div>"
	End If

	If sImgSub <> "" Then sImgSub = "<div style=""padding-top:1px;"">" & sImgSub & "<div style=""clear:both;""></div></div>"
	'--------------------------------------------------------------------------
	'画像 end
	'**************************************************************************

	'**************************************************************************
	'担当業務 start
	'--------------------------------------------------------------------------
	If flgImg = True Then
		'画像が有る場合は文章を短めにカット
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),100) & "&nbsp;"
		If Len(sBusinessDetail) > 100 Then sBusinessDetail = sBusinessDetail & "..."
	Else
		'画像が無い場合は文章を長めにカット
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),155) & "&nbsp;"
		If Len(sBusinessDetail) > 155 Then sBusinessDetail = sBusinessDetail & "..."
	End If
	'--------------------------------------------------------------------------
	'担当業務 end
	'**************************************************************************

	'**************************************************************************
	'勤務形態 start
	'--------------------------------------------------------------------------
	sWorkingType = ""
	sSQL = "sp_GetDataWorkingType '" & oRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		sWorkingType = sWorkingType & oRS2.Collect("WorkingTypeName")
		If (oRS.Collect("OrderType") ="0" And oRS.Collect("Companykbn") = "2") Or oRS.Collect("OrderType") ="1" Or oRS.Collect("OrderType") ="2" Or oRS.Collect("OrderType") ="3" Then
			Select Case oRS2.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "【<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>派遣とは</a>】" 
				Case "002","003": sWorkingType = sWorkingType & "【<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>人材紹介とは</a>】" 
				Case "004": sWorkingType = sWorkingType & "【<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>紹介予定派遣とは</a>】" 
			End Select
		End If
		sWorkingType = sWorkingType & "<br" & sSlash & ">"
		oRS2.MoveNext
	Loop
	Call RSClose(oRS2)
	'--------------------------------------------------------------------------
	'勤務形態 end
	'**************************************************************************

	'**************************************************************************
	'勤務地 start
	'--------------------------------------------------------------------------
	idx = 0
	sWorkingPlace = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True And idx < 3
		dbWorkingPlacePrefectureCode = ChkStr(oRS2.Collect("WorkingPlacePrefectureCode"))
		dbWorkingPlacePrefectureName = ChkStr(oRS2.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS2.Collect("WorkingPlaceCity"))
		'<勤務地アイコン>
		If InStr(sImgOrderState, "/icon_p" & dbWorkingPlacePrefectureCode & ".gif") = 0 Then
			'同じ都道府県アイコンは出さない！
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15""" & sSlash & ">&nbsp;"
		End If
		'</勤務地アイコン>

		'<勤務地>
		If sWorkingPlace <> "" Then sWorkingPlace = sWorkingPlace & "/"
		sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		'</勤務地>

		oRS2.MoveNext
		idx = idx + 1
	Loop
	If oRS2.RecordCount > 3 Then sWorkingPlace = sWorkingPlace & "&nbsp;...他"
	Call RSClose(oRS2)
	'--------------------------------------------------------------------------
	'勤務地 end
	'**************************************************************************

	'**************************************************************************
	'掲載状態リストボックス start
	'--------------------------------------------------------------------------
	sPublicFlag1 = ""
	sPublicFlag0 = ""
	If oRS.Collect("PublicFlag") = "1" Then
		sPublicFlag1 = " selected"
	Else
		sPublicFlag0 = " selected"
	End If
	'--------------------------------------------------------------------------
	'掲載状態リストボックス start
	'**************************************************************************

	'**************************************************************************
	'審査の進捗 start
	'--------------------------------------------------------------------------
	sProgress = ""
	sPublicListDsp = ""
	sPublicDay = ""

	'審査状況
	If oRS.Collect("PermitFlag") = "0" Then
		'リス未審査
		sProgress = "リス審査中"
		sPublicListDsp = "style=""display:none;"""
	ElseIf oRS.Collect("PermitFlag") = "1" Then
		'リス許可済
		If oRS.Collect("PublicFlag") = "0" Then
			sProgress = "リス許可済(非掲載)"
		Else
			sProgress = "掲載中"
		End If
	Else
		'以外
		If oRS.Collect("PublicFlag") = "1" And oRS.Collect("PermitFlag") = "1" Then
			sProgress = "掲載"
		Else
			sProgress = "非掲載"
		End If
		sPublicListDsp = "style=""display:none;"""
	End If

	'掲載日
	sPublicDay = GetDateStr(oRS.Collect("PublicDay"), "/")
	If oRS.Collect("PermitFlag") = "1" And oRS.Collect("PublicDay") > Date Then
		sPublicDay = "<span style=""color:#ff0000;"">未(" & sPublicDay & ")</span>"
		sPublicListDsp = "style=""display:none;"""
	End If
	'--------------------------------------------------------------------------
	'審査の進捗 end
	'**************************************************************************

	'**************************************************************************
	'登録日 start
	'--------------------------------------------------------------------------
	sRegistDay = GetDateStr(oRS.Collect("RegistDay"), "/")
	'--------------------------------------------------------------------------
	'登録日 end
	'**************************************************************************

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	'企業ログイン時以外のときに掲載期限を表示
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(oRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(oRS.Collect("PublicLimitDay"))
	End If

	If sPublishLimitStr = "" Then
		If oRS.Collect("NowPublicFlag") = "0" Then
			'ライセンス切れのときは"掲載終了"と表示
			sPublishLimitStr = "掲載終了"
		Else
			sPublishLimitStr = "常時募集中"
		End If
	End If

	sPublishLimitStr = sPublishLimitStr & "&nbsp;"
	'------------------------------------------------------------------------------
	'求人票掲載期限 end
	'******************************************************************************

	'******************************************************************************
	'仕事の割合 start　10月1日一覧変更用に表示追加_新名
	'------------------------------------------------------------------------------
	If sPlanType = "platinum" Or sPlanType = "gold" Or sPlanType = "old" Then
		sBiz = ""
		sBizName1 = ""
		sBizName2 = ""
		sBizName3 = ""
		sBizName4 = ""
		sBizPercentage1 = ""
		sBizPercentage2 = ""
		sBizPercentage3 = ""
		sBizPercentage4 = ""

		sBizName1 = ChkStr(oRS.Collect("BizName1"))
		sBizName2 = ChkStr(oRS.Collect("BizName2"))
		sBizName3 = ChkStr(oRS.Collect("BizName3"))
		sBizName4 = ChkStr(oRS.Collect("BizName4"))
		sBizPercentage1 = ChkStr(oRS.Collect("BizPercentage1"))
		sBizPercentage2 = ChkStr(oRS.Collect("BizPercentage2"))
		sBizPercentage3 = ChkStr(oRS.Collect("BizPercentage3"))
		sBizPercentage4 = ChkStr(oRS.Collect("BizPercentage4"))
		If sBizPercentage1 = "" Then sBizPercentage1 = "0"
		If sBizPercentage2 = "" Then sBizPercentage2 = "0"
		If sBizPercentage3 = "" Then sBizPercentage3 = "0"
		If sBizPercentage4 = "" Then sBizPercentage4 = "0"

		If Len(sBizName1) >= 17 Then sBizName1 = Left(sBizName1,17) & "..."
		If Len(sBizName2) >= 17 Then sBizName2 = Left(sBizName2,17) & "..."
		If Len(sBizName3) >= 17 Then sBizName3 = Left(sBizName3,17) & "..."
		If Len(sBizName4) >= 17 Then sBizName4 = Left(sBizName4,17) & "..."

		If sBizName1 & sBizName2 & sBizName3 & sBizName4 <> "" Then
			If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""" & sSlash & "></td></tr>"
			If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""" & sSlash & "></td></tr>"
			If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""" & sSlash & "></td></tr>"
			If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""" & sSlash & "></td></tr>"
			sBiz = "<table>" & sBiz & "</table>"
			flgBusiness = True
		End If
	End If
	'------------------------------------------------------------------------------
	'仕事の割合 end
	'******************************************************************************

	'******************************************************************************
	'トップインタビュー start
	'------------------------------------------------------------------------------
	dbTopInterviewFlag = oRS.Collect("TopInterviewFlag")
	'------------------------------------------------------------------------------
	'トップインタビュー end
	'******************************************************************************

	'******************************************************************************
	'ＷバリューＵＲＬ start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(oRS.Collect("WValueURL"))
	'------------------------------------------------------------------------------
	'ＷバリューＵＲＬ end
	'******************************************************************************

	sHTML = sHTML & "<input type=""hidden"" name=""CONF_OrderCodes"" value=""" & oRS.Collect("OrderCode") & """>"
	sHTML = sHTML & "<table border=""0"" class=""old"">"
	sHTML = sHTML & "<tbody>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td class=""old11"" style=""padding-left:0px; width:600px;"" valign=""middle"">"

	If vUserType = "" Or vUserType = "staff" Then
		'非ログイン時、スタッフログイン時

		'・求人票ＵＲＬをメール送信
		'・ウォッチリストへ保存
		sHTML = sHTML & "<div style=""float:left;width:420px;"">"
		sHTML = sHTML & "<img src=""/img/list_companyicon.gif"" alt="""" align=""left""" & sSlash & ">" & sTitleCompanyName
		sHTML = sHTML & "<h3 style=""margin-left:5px;"">■<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & sTitleJobName & "</a>" & sImgMail & "</h3>"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div align=""right"" style=""float:right;font-size:11px;width:113px;"">"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "order/sendmail_jobofferaddress.asp?OrderCode=" & oRS.Collect("OrderCode") & """ onclick=""window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;""><img src=""/img/order/ordermail.gif"" style=""margin-bottom:6px;"" border=""0"" alt=""求人情報をメール送信"" align=""top""" & sSlash & "></a>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "order/sendmail_jobofferaddress.asp?OrderCode=" & oRS.Collect("OrderCode") & """ onclick=""window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;""><img src=""/img/order/orderwachlist.gif"" border=""0"" alt=""ウォッチリストに追加"" align=""top""" & sSlash & "></a>"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	ElseIf vUserType = "company" Then
		'企業ログイン時
		sHTML = sHTML & "<p class=""m0""><img src=""/img/list_companyicon.gif"" alt="""" align=""left""" & sSlash & ">" & sTitleCompanyName & "</p>"
		sHTML = sHTML & "<h3 style=""margin-left:5px;"">■<a href=""../order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & sTitleJobName & "</a>" & sImgMail & "</h3>"
	End If

	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td class=""old12"">"
	'**TOP 08/08/19 Lis林 REP
	'sHTML = sHTML & "<div style=""float:left;"">" & sImgOrderState & "</div>"
	'sHTML = sHTML & "<div align=""right"" style=""font-size:10px;line-height:14px;"">掲載期限：" & sPublishLimitStr & "</div>"
	'sHTML = sHTML & "<div style=""clear:both;""></div>"
	sHTML = sHTML & "<table style='width:600px;'><tr><td style='width:500px;padding-left:5px;'>" & sImgOrderState & "</td>"
	sHTML = sHTML & "<td style='width:100px;vertical-align:top;font-size:10px;text-align:right;'>掲載期限："
	sHTML = sHTML & sPublishLimitStr & "</td></tr></table>"
	'**BTM 08/08/19 Lis林 REP
	sHTML = sHTML & "<table border=""0"" class=""old2"">"

	If sCatchCopy <> "" Then
		sHTML = sHTML & "<caption>" & sCatchCopy & "</caption>"
	End If

	sHTML = sHTML & "<tbody>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td rowspan=""2"" valign=""top"">"

	If flgImg = True Then
		'画像が有る場合のレイアウト
		sHTML = sHTML & "<div class=""old21"" style=""margin:0px 12px;"">"
		sHTML = sHTML & "<b>【担当業務の説明】</b><br" & sSlash & ">" & sBusinessDetail
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div class=""old21"" style=""width:240px; float:left; margin:0px 5px;"">"
		sHTML = sHTML & "<a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """ title=""" & sTitleCompanyName & """>" & sImgMain & "</a>"
		sHTML = sHTML & sImgSub
		sHTML = sHTML & "</div>"
	Else
		'画像が無い場合のレイアウト
		sHTML = sHTML & "<div class=""old21"" style=""width:239px; float:left; margin:0px 5px;"">"
		sHTML = sHTML & "<b>【担当業務の説明】</b><br" & sSlash & ">" & sBusinessDetail
		sHTML = sHTML & "</div><br" & sSlash & ">"
	End If

	sHTML = sHTML & "<table style=""width:330px; margin-left:3px;"">"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
	sHTML = sHTML & "勤務形態"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td style=""background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;"">"
	sHTML = sHTML & sWorkingType
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
	sHTML = sHTML & "勤務地"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td style=""background-color:#eeeeee; padding-left:10px; border-bottom:solid 3px #ffffff;"">"
	sHTML = sHTML & sWorkingPlace & "" & sStationName
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	If sYearlyIncome & sMonthlyIncome & sDailyIncome & sHourlyIncome <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
		sHTML = sHTML & "給与"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;"">"

		If sYearlyIncome <> "" Then
			sHTML = sHTML & "<p>年収&nbsp;" & sYearlyIncome & "</p>"
		End If

		If sMonthlyIncome <> "" Then
			sHTML = sHTML & "<p>月給&nbsp;" & sMonthlyIncome & "</p>"
		End If

		If sDailyIncome <> "" Then
			sHTML = sHTML & "<p>日給&nbsp;" & sDailyIncome & "</p>"
		End If

		If sHourlyIncome <> "" Then
			sHTML = sHTML & "<p>時給&nbsp;" & sHourlyIncome & "</p>"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	If sBizName1 <> "" Then

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; border-bottom:solid 3px #ffffff; text-align:center;"">"
		sHTML = sHTML & "仕事の割合"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""background-color:#eeeeee; border-bottom:solid 3px #ffffff; padding-left:0px; line-height:14px;"">"
		sHTML = sHTML & "<table>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""padding:5px 0px 5px 7px;"">"
		sHTML = sHTML & "<script type=""text/javascript"" language=""javascript"">"
		sHTML = sHTML & "viewWorkAvg(" & sBizPercentage1 & ", " & sBizPercentage2 & ", " & sBizPercentage3 & ", " & sBizPercentage4 & ")"
		sHTML = sHTML & "</script>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"

		If sBizName1 <> "" Then sHTML = sHTML & "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ff9999;"">■</span>" & sBizPercentage1 & "%　" & sBizName1 & "</p>"
		If sBizName2 <> "" Then sHTML = sHTML & "<p style=""font-size:10px; line-height:12px;""><span style=""color:#9999ff;"">■</span>" & sBizPercentage2 & "%　" & sBizName2 & "</p>"
		If sBizName3 <> "" Then sHTML = sHTML & "<p style=""font-size:10px; line-height:12px;""><span style=""color:#99ff99;"">■</span>" & sBizPercentage3 & "%　" & sBizName3 & "</p>"
		If sBizName4 <> "" Then sHTML = sHTML & "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ffff99;"">■</span>" & sBizPercentage4 & "%　" & sBizName4 & "</p>"

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "</table>"
	sHTML = sHTML & "<div align=""right"" style=""margin:3px 5px;"">"

	If dbWValueURL <> "" Then
		sHTML = sHTML & "<a href=""" & dbWValueURL & """ target=""_blank""><img src=""/img/order/btn_wvalue.gif"" border=""0"" alt=""求人情報:" & sTitleCompanyName & "の自社採用ページ""" & sSlash & "></a>"
	End If

	If dbTopInterviewFlag = "1" Then
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "order/order_interview.asp?ordercode=" & dbOrderCode & """><img src=""/img/order/interview_icon.gif"" border=""0"" alt=""求人情報:トップインタビュー""" & sSlash & "></a>"
	End If

	sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """><img src=""/img/detail_button2.gif"" border=""0"" alt=""求人情報:詳細""" & sSlash & "></a>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" And G_USEFLAG = "1" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td class=""old13"">"
		sHTML = sHTML & "<table class=""old3"">"
		sHTML = sHTML & "<tbody>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td class=""old31"">情報コード(" & oRS.Collect("OrderCode") & ")</td>"
		sHTML = sHTML & "<td class=""old32"">状態</td>"
		sHTML = sHTML & "<td class=""old33"">"
		sHTML = sHTML & sProgress & "&nbsp;"
		sHTML = sHTML & "<select name=""CONF_PublicFlags"" " & sPublicListDsp & ">"
		If oRS.Collect("PublicFlag") = "1" Then
			sHTML = sHTML & "<option value=""1"" selected>掲載</option>"
			sHTML = sHTML & "<option value=""0"">非掲載</option>"
		Else
			sHTML = sHTML & "<option value=""1"">掲載</option>"
			sHTML = sHTML & "<option value=""0"" selected>非掲載</option>"
		End If
		sHTML = sHTML & "</select>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td class=""old34"">掲載日<br" & sSlash & ">登録日</td>"
		sHTML = sHTML & "<td class=""old35"">" & sPublicDay & "<br" & sSlash & ">" & sRegistDay & "</td>"
		'sHTML = sHTML & "<td class=""old36""><input type=""checkbox"" name=""CONF_DeleteFlags"" value=""" & oRS.Collect("OrderCode") & """>削除</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td class=""old14""></td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"

	htmlOrderListDetail = sHTML
End Function
%>
