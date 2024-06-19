<%
'**********************************************************************************************************************
'概　要：求人票一覧ページ /order/order_list_entity.asp
'　　　：求人票詳細ページ /order/order_detail_entity.asp
'　　　：企業情報ページ /order/company_order.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　求人票一覧ページ出力用　■■■
'　　　：DspOrderListDetail			：求人票一覧ページの各求人票単位を表示
'　　　：DspOrderListDetail2		：求人票一覧横並びバージョン1
'　　　：DspOrderListDetail3		：求人票一覧横並びバージョン2
'　　　：DspPageControl				：求人票一覧ページのページコントロール
'　　　：
'　　　：■■■　企業情報ページ出力用　■■■
'　　　：DspCompanyInfo				：企業情報の基本情報を出力
'　　　：DspCompanyPR				：企業情報のＰＲ情報を出力
'　　　：
'　　　：■■■　求人票詳細ページ出力用　■■■
'　　　：DspLisOrderCompanyInfo		：求人票詳細ページのリスの紹介先・派遣先企業情報を出力
'　　　：DspTempOrderCompanyInfo	：求人票詳細ページの派遣企業の派遣先企業情報を出力
'　　　：DspOrderControlButton		：求人票詳細ページのコントロールボタン（ログイン済みユーザ用）
'　　　：JSOrderControlButton		：求人票詳細ページのコントロールボタンで利用するjavascriptの出力
'　　　：FrmOrderControlButton		：求人票詳細ページのコントロールボタンで利用するFORMデータの出力
'　　　：DspOrderCompanyName		：求人票詳細ページの企業名を出力
'　　　：DspOrderShowTypeSwitch		：求人票詳細ページの会社情報・職種情報切り替えボタンと参照回数を出力
'　　　：DspOrderCatchCopy			：求人票詳細ページのキャッチコピー部分（大きい画像など）を出力
'　　　：DspOrderFreePR				：求人票詳細ページのフリーＰＲを出力
'　　　：DspOrderPictureNow			：求人票詳細ページの小さい画像を出力
'　　　：DspBusiness				：求人票詳細ページの業務内容を出力
'　　　：DspCondition				：求人票詳細ページの勤務条件を出力
'　　　：DspNeedCondition			：求人票詳細ページの必要条件を出力
'　　　：DspHowToEntry				：求人票詳細ページの応募情報を出力
'　　　：DspContact					：求人票詳細ページの担当者連絡先を出力
'　　　：DspConsultantComment		：リスの案件担当者、コンサル所見を出力
'　　　：DspNewMail					：求人票詳細ページの最新の送信済みメールを出力
'　　　：GetWorkingType				：求人票詳細ページの勤務形態部分
'　　　：GetJobType					：求人票詳細ページの職種部分
'　　　：GetWorkingTime				：求人票詳細ページの勤務形態部分
'　　　：GetNearbyStation			：求人票詳細ページの最寄駅部分
'　　　：GetNearbyRailway			：求人票詳細ページの最寄沿線部分
'　　　：GetSkill					：求人票詳細ページのスキル部分
'　　　：GetLicense					：求人票詳細ページの資格部分
'　　　：GetOrderNote				：求人票詳細ページの資格部分
'　　　：GetOrderTitle				：求人票詳細ページのタイトルとディスクリプションを取得
'　　　：GetSkillList				：求人票詳細ページのスキルの各項目表示
'　　　：
'　　　：■■■　レコメンド　■■■
'　　　：DspRecommendOrderList		：レコメンドお仕事情報一覧出力
'　　　：GetRecommendValues			：レコメンドの求人票一覧の、求人票一つ一つの各項目（職種、企業名など）を取得
'　　　：
'　　　：■■■　求人票詳細ページチェック用　■■■
'　　　：ChkMyOrder					：自社求人票か否かをチェックする ["0"]自社求人票以外 ["1"]自社求人票
'　　　：
'　　　：■■■　掲載状態変更・求人票削除　■■■
'　　　：UpdMyOrderPublicFlag		：自社求人票の掲載状態を変更する
'　　　：DelMyOrder					：自社求人票を削除する
'　　　：
'　　　：■■■　共通利用　■■■
'　　　：GetImgOrderSpeciality		：求人票の特徴
'　　　：
'　　　：■■■　＠履歴書としごとナビで表示が異なる部分用　■■■
'　　　：DspTopRegButton			：しごとナビの求人票詳細ページの上部に置く、ログイン誘導ボタン。
'　　　：DspTopRegButtonResume		：＠履歴書の求人票詳細ページの上部に置く、ログイン誘導ボタン。
'　　　：DspBottomRegButton			：しごとナビの求人票詳細ページの下部に置く、ログイン誘導ボタン。
'　　　：DspBottomRegButtonResume	：＠履歴書の求人票詳細ページの下部に置く、ログイン誘導ボタン。
'　　　：
'　　　：■■■　求人票詳細アクセス時の制御　■■■
'　　　：MailMagazineAccess			：新着求人メールからのアクセス時のログ書き込み
'　　　：MailMagazineDelivery		：求人メルマガからのアクセス時のログ書き込み
'　　　：AccessHistoryOrder			：足跡ログの書き込み
'　　　：AccessCountUp				：アクセス回数のカウントアップ
'**********************************************************************************************************************

'******************************************************************************
'概　要：求人票一覧ページの各求人票項目を表示
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_SearchOrder or 求人票詳細検索SQL で生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder		：利用中ユーザの自社求人票か否か ["1"]自社求人票 ["0"]自社求人票でない
'使用元：order/order_list_entity.asp
'備　考：
'更　新：2006/05/13 LIS K.Kokubo 作成
'　　　：2007/11/22 LIS K.Kokubo up_SearchOrderを必要最小限のものだけを取ってくるようにしたことによる変更。sp_GetDetailOrderからデータを取得。
'******************************************************************************
Function DspOrderListDetail(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Const PICSIZEW = 240
	Const PICSIZEH = 180
	Const PICSIZESUBW = 72
	Const PICSIZESUBH = 56

	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim sOrderCode			'情報コード
	Dim sOrderType			'受注種類
	Dim sTitleJobName		'職種
	Dim sTitleCompanyName	'会社名
	Dim sImgMail			'送信済みメール画像
	Dim sImgOrderState		'状態画像 HOT,新着,未経験OK,UIターン,語学,休日120日,フレックス
	Dim sCatchCopy			'キャッチコピー
	Dim flgImg				'画像の有無フラグ(画像の有無でレイアウトが変化) [True]有 [False]無
	Dim sImgMain			'大きい画像
	Dim sImgSub				'小さい画像
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
	Dim sMonthlyIncomeMin	'月収下限
	Dim sMonthlyIncomeMax	'月収上限
	Dim sDailyIncomeMin		'月給下限
	Dim sDailyIncomeMax		'月給上限
	Dim sHourlyIncomeMin	'時給下限
	Dim sHourlyIncomeMax	'時給上限
	Dim sYearlyIncome		'年収表示用
	Dim sDailyIncome		'月収表示用
	Dim sMonthlyIncome		'日給表示用
	Dim sHourlyIncome		'時給表示用
	'希望勤務形態・希望勤務地アイコン　10月1日一覧変更用に表示追加_新名
	Dim sWorkingCode
	Dim sWorkingName
	Dim sWorkingPlacePrefectureName
	Dim sBiz
	Dim sBizName1
	Dim sBizName2
	Dim sBizName3
	Dim sBizName4
	Dim sBizPercentage1
	Dim sBizPercentage2
	Dim sBizPercentage3
	Dim sBizPercentage4
	Dim flgAddWatchList
	Dim flgBusiness

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = rRS.Collect("OrderCode")

	DspOrderListDetail = False

	sSQL = "sp_GetDetailOrder '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	sOrderType = ChkStr(oRS.Collect("OrderType"))

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
	'月収
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
	'------------------------------------------------------------------------------
	sStationName = ""
	sSQL = "sp_GetDataNearbyStation '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		sStationName ="【" & sStationName & GetStrNearbyStation(oRS2.Collect("StationName"), "", "") & "】"
	End If
	'------------------------------------------------------------------------------
	'最寄駅 end
	'******************************************************************************

	'**************************************************************************
	'メール送信済み確認 start
	'--------------------------------------------------------------------------
	If vUserType = "staff" Then
		sSQL = "sp_GetDataMailHistory '" & vUserID & "', '', '" & sOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			sImgMail = "<img src=""/img/s_contact.gif"" alt=""メール送信済み"">"
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
	sImgOrderState = "&nbsp;"
	'アクセス数が100を超えていれば「HOT」表示（リス安藤）
	If oRS.Collect("AccessCount") > 100 Then
		sImgOrderState = sImgOrderState & "<img src=""/img/c_HOT_green.gif"" alt=""人気"">&nbsp;"
	End If

	'UPDATEと今日から10日引いた日で「新着」表示(リス安藤)
	If oRS.Collect("UpdateDay") > NOW()-10 Then
		sImgOrderState = sImgOrderState & "<img src=""/img/c_NEW_green.gif"" alt=""新着"">&nbsp;"
	End If

	'未経験者ＯＫの場合、わかばマーク表示(リス安藤)
	If oRS.Collect("InexperiencedPersonFlag") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/no_experience.gif"" alt=""未経験者／第二新卒歓迎"">&nbsp;"
	End If

	'Ｕターン・Ｉターン
	If oRS.Collect("UITurnFlag") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/ui_turn.gif"" alt=""Ｕターン・Ｉターン"">&nbsp;"
	End If

	'語学を活かす仕事
	If oRS.Collect("UtilizeLanguageFlag") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/linguistic_job.gif"" alt=""語学を活かす仕事"">&nbsp;"
	End If

	'年間休日120日以上
	If oRS.Collect("ManyHolidayFlag") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/year_holidaycnt.gif"" alt=""年間休日120日以上"">&nbsp;"
	End If

	'フレックスタイム制度あり ------2006/01/10 Hayashi ADD
	If oRS.Collect("FlexTimeFlag") = "1" And oRS.Collect("OrderType") = "0" And oRS.Collect("CompanyKbn") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/flextime.gif"" alt=""フレックスタイム制度あり"">&nbsp;"
	End If

	sSQL = "sp_GetDataWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		sWorkingCode = oRS2.Collect("WorkingTypeCode")
		sWorkingName = oRS2.Collect("WorkingTypeName")

		sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_w" & sWorkingCode & ".gif"" alt=""" & sWorkingName & """ width=""50"" height=""15"">&nbsp;"

		oRS2.MoveNext
	Loop
	sWorkingPlacePrefectureName = oRS.Collect("WorkingPlacePrefectureName")
	If oRS.Collect("Workingplaceprefecturecode") >= "048" Then
		sWorkingPlacePrefectureName = "海外"
	End If

	sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & oRS.Collect("WorkingPlacePrefectureCode") & ".gif"" alt=""" & sWorkingplaceprefecturename & """ width=""50"" height=""15"">&nbsp;"

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
	sImgMain = ""
	sImgSub = ""
	sCompanyPictureFlag = ChkStr(oRS.Collect("CompanyPictureFlag"))

	sSQL = "up_GetListOrderPictureNow '" & oRS.Collect("CompanyCode") & "', '" & oRS.Collect("OrderCode") & "', 'orderpicture'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		If ChkStr(oRS2.Collect("OptionNo1")) <> "" Or (sOrderType = "0" And sCompanyPictureFlag = "1") Then
			sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=" & oRS2.Collect("OptionNo1") & """ alt="""" border=""0"" width=""" & PICSIZEW & """ height=""" & PICSIZEH & """>"
			flgImg = True
		End If

		If ChkStr(oRS2.Collect("OptionNo2")) <> "" Then
			sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
				"<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=" & oRS2.Collect("OptionNo2") & """ alt=""" & oRS2.Collect("Caption2") & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""><br>"
			sImgSub = sImgSub & "</div>"
			flgImg = True
		End If

		If ChkStr(oRS2.Collect("OptionNo3")) <> "" Then
			sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
				"<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=" & oRS2.Collect("OptionNo3") & """ alt=""" & oRS2.Collect("Caption3") & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""><br>"
			sImgSub = sImgSub & "</div>"
			flgImg = True
		End If

		If ChkStr(oRS2.Collect("OptionNo4")) <> "" Then
			sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
				"<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=" & oRS2.Collect("OptionNo4") & """ alt=""" & oRS2.Collect("Caption4") & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""><br>"
			sImgSub = sImgSub & "</div>"
			flgImg = True
		End If
		If sImgSub <> "" Then sImgSub = sImgSub & "<div style=""clear:both;""></div>"
	Else
		If sCompanyPictureFlag = "1" And sOrderType = "0" Then
			sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=1"" alt="""" border=""0"" width=""" & PICSIZEW & """ height=""" & PICSIZEH & """>"
			flgImg = True
		End If
	End If

	Call RSClose(oRS2)
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
		If (oRS.Collect("OrderType") ="0" And oRS.Collect("Companykbn") = "2") Or oRS.Collect("OrderType") ="2" Then
			sWorkingType = sWorkingType & "【<a href=""javascript:void(0)"" onclick=""window.open('/staff/s_shokai.htm','count','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=300,height=200')"">人材紹介</a>】"
		End If
		sWorkingType = sWorkingType & "<br>"
		oRS2.MoveNext
	Loop
	Call RSClose(oRS2)
	'--------------------------------------------------------------------------
	'勤務形態 end
	'**************************************************************************

	'**************************************************************************
	'勤務地 start
	'--------------------------------------------------------------------------
	sWorkingPlace = oRS.Collect("WorkingPlacePrefectureName") & oRS.Collect("WorkingPlaceCity") & "&nbsp;"
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
	'企業コード start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "up_GetDataWatchList '" & vUserID & "', '', '', '" & sOrderCode & "', ''"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	If GetRSState(oRS2) = False Then
		flgAddWatchList = True
	End If
	Call RSClose(oRS2)
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	'企業ログイン時以外のときに掲載期限を表示
	sPublishLimitStr = GetDateStr(oRS.Collect("riyotodate"), "/")

	If sPublishLimitStr = "" Then
		sPublishLimitStr = "常時募集中" 
	End If

	sPublishLimitStr = sPublishLimitStr & "&nbsp;"
	'------------------------------------------------------------------------------
	'求人票掲載期限 end
	'******************************************************************************

	'******************************************************************************
	'仕事の割合 start　10月1日一覧変更用に表示追加_新名
	'------------------------------------------------------------------------------
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
		If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""></td></tr>"
		If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""></td></tr>"
		If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""></td></tr>"
		If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""></td></tr>"
		sBiz = "<table>" & sBiz & "</table>"
		flgBusiness = True
	End If
	'------------------------------------------------------------------------------
	'仕事の割合 end
	'******************************************************************************


%>
<input type="hidden" name="CONF_OrderCodes" value="<%= oRS.Collect("OrderCode") %>">
<table border="0" class="old">
	<tbody>
	<tr>
		<td class="old11" style="padding-left:0px; width:600px;" valign="middle">
<%
	If vUserType = "" Or vUserType = "staff" Then
		'非ログイン時、スタッフログイン時

		If G_USERID <> "" And G_FLGRESUME = False or G_FLGRESUME = False Then
			'しごとナビの求人票一覧の場合は以下を表示
			'・求人票ＵＲＬをメール送信
			'・ウォッチリストへ保存
%>
			<div style="float:left;width:420px;">
			<img src="/img/list_companyicon.gif" alt="" align="left"><%= sTitleCompanyName %>
			<h3 style="margin-left:5px;">■<a href="<%= HTTP_NAVI_CURRENTURL %>order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>"><%= sTitleJobName %></a><%= sImgMail %></h3>
			</div>
			<div align="right" style="float:right;font-size:11px;width:113px;">
			<a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = oRS.Collect("OrderCode") %>" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;"><img src="/img/order/ordermail.gif" style="margin-bottom:6px;" border="0" alt="求人情報をメール送信" align="top"></a>
			<a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = oRS.Collect("OrderCode") %>" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;"><img src="/img/order/orderwachlist.gif" border="0" alt="ウォッチリストに追加" align="top"></a>
			</div>
			<div style="clear:both;"></div>
<%
		Else
			'＠履歴書の求人票一覧の場合は以下を表示しない！
			'・ＵＲＬをメール送信
			'・ウォッチリストへ保存
%>
			<p class="m0"><img src="/img/list_companyicon.gif" alt="" align="left"><%= sTitleCompanyName %></p>
			<h3 style="margin-left:5px;">■<a href="../order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>"><%= sTitleJobName %></a><%= sImgMail %></h3>
<%
		End If

	ElseIf vUserType = "company" Then
		'企業ログイン時
%>
			<p class="m0"><img src="/img/list_companyicon.gif" alt="" align="left"><%= sTitleCompanyName %></p>
			<h3 style="margin-left:5px;">■<a href="../order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>"><%= sTitleJobName %></a><%= sImgMail %></h3>
<%
	End If
%>
		</td>
	</tr>
	<tr>
		<td class="old12">
			<div style="float:left;"><%= sImgOrderState %></div>
			<div align="right" style="font-size:10px;line-height:14px;">掲載期限：<%= sPublishLimitStr %></div>
			<div style="clear:both;"></div>
			<table border="0" class="old2">
<%
	If sCatchCopy <> "" Then
%>
				<caption><%= sCatchCopy %></caption>
<%
	End If
%>
				<tbody>
				<tr>
					<td rowspan="2" valign="top">
<%
	If flgImg = True Then
		'画像が有る場合のレイアウト
%>
					<div class="old21" valign="top" style="margin:0px 12px;">
					<b>【担当業務の説明】</b><br><%= sBusinessDetail %>
					</div>
					<div class="old21" valign="top" style="width:240px; float:left; margin:0px 5px;">
						<a href="../order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>" title="<%= sTitleCompanyName %>"><%= sImgMain %></a>
						<%= sImgSub %>
					</div>
<%
	Else
		'画像が無い場合のレイアウト
%>
					<div class="old21" valign="top" style="width:239px; float:left; margin:0px 5px;">
					<b>【担当業務の説明】</b><br><%= sBusinessDetail %>
					</div><br>
<%
	End If
%>
					<table style="width:330px; margin-left:3px;">
						<tr>
							<td style="font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;">
							勤務形態
							</td>
							<td style="background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;">
							<%= sWorkingType %>
							</td>
						</tr>
						<tr>
							<td style="font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;">
							勤務地
							</td>
							<td style="background-color:#eeeeee; padding-left:10px; border-bottom:solid 3px #ffffff;">
							<%= sWorkingPlace %><%= sStationName %>
							</td>
						</tr>
						<tr>
							<td style="font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;">
							給与
							</td>
							<td style="background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;">
<%
			If sYearlyIncome <> "" Then
%>
							<p>年収 <%= sYearlyIncome %></p>
<%
			End If

			If sMonthlyIncome <> "" Then
%>
							<p>月収 <%= sMonthlyIncome %></p>
<%
			End If

			If sDailyIncome <> "" Then
%>
							<p>日給 <%= sDailyIncome %></p>
<%
			End If

			If sHourlyIncome <> "" Then
%>
							<p>時給 <%= sHourlyIncome %></p>
<%
			End If
%>
							</td>
						</tr>
<%
	If sBizName1 <> "" Then
%>

						<tr>
							<td style="font-weight:bold; background-color:#E1FBCD; width:70px; border-bottom:solid 3px #ffffff; text-align:center;">
							仕事の割合
							</td>
							<td style="background-color:#eeeeee; border-bottom:solid 3px #ffffff; padding-left:0px; line-height:14px;">
								<table>
									<tr>
										<td style="padding:5px 0px 5px 7px;">
										<script type="text/javascript" language="javascript">
											viewWorkAvg(<%= sBizPercentage1 %>, <%= sBizPercentage2 %>, <%= sBizPercentage3 %>, <%= sBizPercentage4 %>)
										</script>
										</td>
										<td>
<%
		If sBizName1 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ff9999;"">■</span>" & sBizPercentage1 & "%　" & sBizName1 & "</p>"
		If sBizName2 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#9999ff;"">■</span>" & sBizPercentage2 & "%　" & sBizName2 & "</p>"
		If sBizName3 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#99ff99;"">■</span>" & sBizPercentage3 & "%　" & sBizName3 & "</p>"
		If sBizName4 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ffff99;"">■</span>" & sBizPercentage4 & "%　" & sBizName4 & "</p>"
%>
										</td>
									</tr>
								</table>
							</td>
						</tr>
<%
	End If
%>
					</table>
						<div align="right" style="margin:3px 5px;">
							<a href="../order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>">
								<img src="/img/detail_button2.gif" border="0" alt="">
							</a>
						</div>
					</td>
				</tr>
			</table>
			<div style="clear:both;"></div>
		</td>
	</tr>
<%
	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" Then
%>
	<tr>
		<td class="old13">
			<table class="old3">
				<tbody>
				<tr>
					<td class="old31">情報コード(<% = oRS.Collect("OrderCode") %>)</td>
					<td class="old32">状態</td>
					<td class="old33">
						<%= sProgress %>
						<select name="CONF_PublicFlags" <%= sPublicListDsp %>>
						<option value="1"<% If oRS.Collect("PublicFlag") = "1" Then Response.Write(" selected") %>>掲載</option>
						<option value="0"<% If oRS.Collect("PublicFlag") = "0" Then Response.Write(" selected") %>>非掲載</option>
						</select>
					</td>
					<td class="old34">掲載日<br>登録日</td>
					<td class="old35"><%= sPublicDay %><br><%= sRegistDay %></td>
					<td class="old36"><input type="checkbox" name="CONF_DeleteFlags" value="<%= oRS.Collect("OrderCode") %>">削除</td>
				</tr>
				</tbody>
			</table>
		</td>
	</tr>
<%
	End If
%>
	<tr>
		<td class="old14"></td>
	</tr>
</table>
<%
	DspOrderListDetail = True
End Function

'******************************************************************************
'概　要：求人票一覧の横並びバージョン
'引　数：rDB		：DB接続オブジェクト
'　　　：rRS		：求人票一覧のレコードセット
'　　　：vCols		：現在の列数
'　　　：vMaxCols	：列最大数
'戻り値：
'作成日：2007/05/23
'作成者：Lis Kokubo
'備　考：
'更　新：
'******************************************************************************
Function DspOrderListDetail2(ByRef rDB, ByRef rRS, ByVal vCols, ByVal vMaxCols)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim sOrderCode			'情報コード
	Dim sOrderType			'受注区分
	Dim sCompanyKbn			'会社区分
	Dim sCompanyName		'企業名
	Dim sCompanyNameF		'企業名カナ
	Dim sCompanySpeciality	'企業名（特徴）
	Dim sJobTypeDetail		'具体的職種名(altやtitleで出力する)
	Dim sViewJobTypeDetail	'求職者に見える具体的職種名(長い文字列はカットされる)
	Dim sBusinessDetail		'担当業務
	Dim sYearlyIncome		'年収
	Dim sYearlyIncomeMin	'年収下限
	Dim sYearlyIncomeMax	'年収上限
	Dim sMonthlyIncome		'月収
	Dim sMonthlyIncomeMin	'月収下限
	Dim sMonthlyIncomeMax	'月収上限
	Dim sDailyIncome		'日給
	Dim sDailyIncomeMin		'日給下限
	Dim sDailyIncomeMax		'日給上限
	Dim sHourlyIncome		'時給
	Dim sHourlyIncomeMin	'時給下限
	Dim sHourlyIncomeMax	'時給上限
	Dim sWorkingTypeIcon	'勤務形態アイコン並び
	Dim sStation			'最寄駅
	Dim sImg				'画像URL

	Dim sURL				'求人票詳細のURL
	Dim sAlign				'枠寄せ [vCols = 1]left [vCols = vMaxCols]right [それ以外]center

	If GetRSState(rRS) = False Then Exit Function

	sURL = HTTP_CURRENTURL & "order/order_detail.asp"

	If vCols = 1 Then
		sAlign = "left"
	ElseIf vCols = vMaxCols Then
		sAlign = "right"
	Else
		sAlign = "center"
	End If

	sSQL = "sp_GetDetailOrder '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'情報コード
	sOrderCode = ChkStr(oRS.Collect("OrderCode"))
	'受注区分
	sOrderType = ChkStr(oRS.Collect("OrderType"))
	'企業区分
	sCompanyKbn = ChkStr(oRS.Collect("CompanyKbn"))
	'企業名, 企業名カナ
	sCompanyName = ChkStr(oRS.Collect("CompanyName"))
	sCompanyNameF = ChkStr(oRS.Collect("CompanyName_F"))
	sCompanySpeciality = ChkStr(oRS.Collect("CompanySpeciality"))
	Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'具体的職種名
	sJobTypeDetail = ChkStr(oRS.Collect("JobTypeDetail"))
	sViewJobTypeDetail = sJobTypeDetail
	If Len(sViewJobTypeDetail) > 14 Then sViewJobTypeDetail = Left(sViewJobTypeDetail, 14) & ".."
	'担当業務
	sBusinessDetail = ChkStr(oRS.Collect("BusinessDetail"))

	'******************************************************************************
	'給与 start
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
	'月収
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
	'勤務形態アイコン start
	'------------------------------------------------------------------------------
	sWorkingTypeIcon = ""
	sSQL = "sp_GetListWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		Select Case ChkStr(oRS2.Collect("WorkingTypeCode"))
			Case "001": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/haken.gif"" alt=""派遣"">&nbsp;"
			Case "002": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/seishain.gif"" alt=""正社員"">&nbsp;"
			Case "003": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/keiyaku.gif"" alt=""契約社員"">&nbsp;"
			Case "004": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/syoha.gif"" alt=""紹介予定派遣"">&nbsp;"
			Case "005": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/arbeit.gif"" alt=""アルバイト・パート"">&nbsp;"
			Case "006": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/soho.gif"" alt=""SOHO"">&nbsp;"
			Case "007": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/fc.gif"" alt=""FC"">&nbsp;"
		End Select
		oRS2.MoveNext
	Loop
	Call RSClose(oRS2)
	'------------------------------------------------------------------------------
	'勤務形態アイコン end
	'******************************************************************************

	'******************************************************************************
	'画像 start
	'------------------------------------------------------------------------------
	sImg = ""
	sSQL = "up_GetListOrderPictureNow '" & sCompanyCode & "', '" & sOrderCode & "', 'orderpicture'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo1")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo1")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo2")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo2")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo3")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo3")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo4")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo4")
	End If

	If sImg = "" And sOrderType = "0" Then
		sSQL = "sp_GetDataPicture '" & sCompanyCode & "', '1'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"
		End If
	End If

	If sImg = "" Then sImg = "/img/nopicture180.gif"
	'------------------------------------------------------------------------------
	'画像 end
	'******************************************************************************

	'******************************************************************************
	'最寄駅 start
	'------------------------------------------------------------------------------
	sStation = ""
	sSQL = "sp_GetDataNearbyStation '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		sStation = sStation & GetStrNearbyStation(oRS2.Collect("StationName"), oRS2.Collect("ToStationTime"), oRS2.Collect("ToStationRemark"))
		oRS2.MoveNext
		If GetRSState(oRS2) = True Then sStation = sStation & "<br>"
	Loop
	'------------------------------------------------------------------------------
	'最寄駅 end
	'******************************************************************************
%>
<div align="<%= sAlign %>" style="float:left; width:200px;">
	<table class="pattern1" border="0" style="width:195px;">
		<thead>
		<tr>
			<th colspan="2" valign="top" style="width:183px;">
				<div style="float:left; width:64px;"><img src="<%= sImg %>" alt="<%= sJobTypeDetail %>" width="64" height="48"></div>
				<div style="float:left; width:114px; margin-left:5px;"><a href="<%= sURL %>?ordercode=<%= sOrderCode %>"><%= sViewJobTypeDetail %></a></div>
				<br clear="all">
			</th>
		</tr>
		</thead>
		<tbody>
<!--
		<tr>
			<td colspan="2" align="center">
				<a href="<%= sURL %>?ordercode=<%= sOrderCode %>" title="<%= sJobTypeDetail %>">
					<img src="<%= sImg %>" alt="<%= sJobTypeDetail %>" border="1" width="180" height="135" style="border-color:#999999;">
				</a>
			</td>
		</tr>
-->
		<tr>
			<th style="width:63px;">会社名</th>
			<td style="width:109px;"><%= sCompanyName %></td>
		</tr>
		<tr>
			<th>勤務形態</th>
			<td><%= sWorkingTypeIcon %></td>
		</tr>
<!--
		<tr>
			<th>担当業務</th>
			<td><%= sBusinessDetail %></td>
		</tr>
-->
		<tr>
			<th>最寄駅</th>
			<td><%= sStation %></td>
		</tr>
<%
			If sYearlyIncome <> "" Then
%>
		<tr>
			<th>年収</th>
			<td><%= sYearlyIncome %></td>
		</tr>
<%
			End If

			If sMonthlyIncome <> "" Then
%>
		<tr>
			<th>月収</th>
			<td><%= sMonthlyIncome %></td>
		</tr>
<%
			End If

			If sDailyIncome <> "" Then
%>
		<tr>
			<th>日給</th>
			<td><%= sDailyIncome %></td>
		</tr>
<%
			End If

			If sHourlyIncome <> "" Then
%>
		<tr>
			<th>時給</th>
			<td><%= sHourlyIncome %></td>
		</tr>
<%
			End If
%>
		</tbody>
	</table>
</div>
<%
End Function

'******************************************************************************
'概　要：求人票一覧横並びバージョン2
'引　数：rDB		：DB接続オブジェクト
'　　　：rRS		：お仕事検索結果を保持するのレコードセット
'　　　：vPageSize	：１ページあたりの求人票件数
'　　　：vPage		：現在表示中のページ
'　　　：vRCMD		：レコメンド種類 ["1"]こんなお仕事情報も見てます ["2"]近い条件のお仕事情報 ["3"]資格
'戻り値：
'作成日：2007/05/31
'作成者：Lis Kokubo
'備　考：
'更　新：
'******************************************************************************
Function DspOrderListDetail3(ByRef rDB, ByRef rRS, ByVal vPageSize, ByVal vPage, ByVal vRCMD)
	Const MAXCOLS = 3

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTitle
	Dim iRecordCnt	'レコード件数
	Dim idx			'ループカウントアップ変数
	Dim iCols		'列数
	Dim aPadding(2)	'各列のパディング
	Dim aJobTypeDetail()
	Dim aCompanyName()
	Dim aImg()
	Dim aWorkingTypeIcon()
	Dim aWorkingPlace()
	Dim aStation()
	Dim aYearlyIncome()
	Dim aMonthlyIncome()
	Dim aDailyIncome()
	Dim aHourlyIncome()

	If GetRSState(rRS) = False Then Exit Function
	If IsNumeric(vPageSize) = False Then Exit Function

	If IsNumeric(vPage) = False Then vPage = 1
	rRS.PageSize = vPageSize
	rRS.AbsolutePage = vPage

	If GetRSState(rRS) = False Then Exit Function

	iRecordCnt = 0
	idx = 0
	Do While GetRSState(rRS) = True And idx < vPageSize
		ReDim Preserve aJobTypeDetail(idx)
		ReDim Preserve aCompanyName(idx)
		ReDim Preserve aImg(idx)
		ReDim Preserve aWorkingTypeIcon(idx)
		ReDim Preserve aWorkingPlace(idx)
		ReDim Preserve aStation(idx)
		ReDim Preserve aYearlyIncome(idx)
		ReDim Preserve aMonthlyIncome(idx)
		ReDim Preserve aDailyIncome(idx)
		ReDim Preserve aHourlyIncome(idx)

		Call GetRecommendValues(rDB, rRS, vRCMD, aJobTypeDetail(idx), aCompanyName(idx), aImg(idx), aWorkingTypeIcon(idx), aWorkingPlace(idx), aStation(idx), aYearlyIncome(idx), aMonthlyIncome(idx), aDailyIncome(idx), aHourlyIncome(idx))
		idx = idx + 1
		iRecordCnt = iRecordCnt + 1
		rRS.MoveNext
	Loop

	'各列のパディング
	aPadding(0) = "padding:0px 4px 0px 0px;"
	aPadding(1) = "padding:0px 2px 0px 2px;"
	aPadding(2) = "padding:0px 0px 0px 4px;"

	idx = 0
	Do While idx < iRecordCnt
		For iCols = 0 To MAXCOLS - 1
			If idx >= iRecordCnt Then Exit For

			Response.Write "<div style=""float:left; width:200px;""><div style=""line-height:16px; " & aPadding(iCols) & """>"

			Response.Write aImg(idx)
			If aJobTypeDetail(idx) <> "" Then Response.Write "【職種】" & aJobTypeDetail(idx) & "<br>" & vbCrLf
			'If aCompanyName(idx) <> "" Then Response.Write "【企業】" & aCompanyName(idx) & "<br>" & vbCrLf
			If aWorkingTypeIcon(idx) <> "" Then Response.Write "【形態】" & aWorkingTypeIcon(idx)  & "<br>"& vbCrLf
			If aWorkingPlace(idx) <> "" Then Response.Write "【場所】" & aWorkingPlace(idx) & "<br>" & vbCrLf
			If aStation(idx) <> "" Then Response.Write "【最寄】" & Replace(aStation(idx), "<br>", "、") & "<br>" & vbCrLf
			Response.Write "【給与】"
			If aYearlyIncome(idx) <> "" Then Response.Write "[年収]" & aYearlyIncome(idx)
			If aMonthlyIncome(idx) <> "" Then Response.Write "[月収]" & aMonthlyIncome(idx)
			If aDailyIncome(idx) <> "" Then Response.Write "[日給]" & aDailyIncome(idx)
			If aHourlyIncome(idx) <> "" Then Response.Write "[時給]" & aHourlyIncome(idx)

			idx = idx + 1
			Response.Write "</div></div>"
		Next

		Response.Write "<div style=""padding-bottom:15px; clear:both;""></div>"
	Loop
End Function

'******************************************************************************
'概　要：求人票一覧ページのページコントロール
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_SearchOrder or 求人票詳細検索SQL で生成されたレコードセットオブジェクト
'　　　：vPageSize		：１ページあたりの表示件数
'　　　：vPage			：表示中ページ
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_list_entity.asp
'　　　：しごとナビ/order/company_order.asp
'******************************************************************************
Function DspPageControl(ByRef rDB, ByRef rRS, ByVal vPageSize, ByVal vPage)
	Dim iMaxPage
	Dim iLine
	Dim S_Page
	Dim E_Page
	Dim Sort
	Dim idx

	If GetRSState(rRS) = False Then Exit Function

	If vPage <> "" Then vPage = CInt(vPage)

	'ページあたりの表示件数
	rRS.PageSize = vPageSize

	iMaxPage = rRS.PageCount
	If vPage > iMaxPage Then vPage = iMaxPage
	rRS.AbsolutePage = vPage

	'画面上に表示する開始・終了ページ番号を設定
	'表示開始ページ番号を指定
	S_Page = vPage - 5
	If S_Page < 1 Then
		S_Page = 1
	End If

	'表示終了ページ番号を指定
	E_Page = vPage + 4
	If E_Page < 10 Then E_Page = 10
	If E_Page > iMaxPage Then
		E_Page = iMaxPage
	End If
	If S_Page > iMaxPage - 9 And iMaxPage - 9 > 0 Then S_Page = iMaxPage - 9
%>
<table style="width:600px; margin:10px 0px;">
	<tbody>
	<tr>
		<td style="width:88px; padding:5px; border-width:1px 0px 1px 1px; text-align:center;">
<%
	If vPage > 1 Then Response.Write "<a href='javascript:ChgPage(" & vPage - 1 & ");'>前のページ</a>"
%>
		</td>
		<td style="width:489px; padding:5px; border-width:1px 0px 1px 0px; text-align:center;">
<%
	If S_Page <> 1 Then Response.Write "…"
	For idx = S_Page To E_Page	'ページ番号を表示
		Response.write "　"
		If idx = vPage Then		'指定ページの表示
			Response.Write "[" & idx & "]"
		Else
			Response.Write "<a href='javascript:ChgPage(" & idx & ");'>" & idx & "</a>"
		End If
	Next
	If E_Page < iMaxPage Then Response.Write "　…"
%>
		</td>
		<td style="width:89px; padding:5px; border-width:1px 1px 1px 0px; text-align:center;">
<%
	If vPage < iMaxPage Then Response.Write "<a href='javascript:ChgPage(" & vPage + 1 & ");'>次のページ</a>"
%>
		</td>
	</tr>
	</tbody>
</table>
<%
End Function

'******************************************************************************
'概　要：企業情報の基本情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vOrderCode		：
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'******************************************************************************
Function DspCompanyInfo(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Dim sCompanyCode		'企業コード
	Dim sCompanyName		'企業名称
	Dim sCompanyNameF		'企業名称カナ
	Dim sOrderType			'求人種類 ["0"]しごとナビ一般 ["1"]派遣 ["2"]紹介 ["3"]
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sCompanyKbn			'企業区分
	Dim sCompanySpeciality	'企業特徴
	Dim sEstablishYear		'設立年度
	Dim sCapitalAmount		'資本額
	Dim sListClass			'株式公開
	Dim sEmployeeNum		'社員数
	Dim sIndustryType		'業種
	Dim sAddress			'本社住所
	Dim sHomePage			'ホームページ
	Dim sClass				'使用するスタイルシートのクラス　画像の有無で変化
	Dim sLineClass			'
	Dim flgLine				'線引きフラグ
	Dim sAddTitle			'派遣企業の情報の場合は「派遣」を項目名に付ける

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'会社名 start
	'------------------------------------------------------------------------------
	sCompanyName = rRS.Collect("CompanyName")
	sCompanyNameF = rRS.Collect("CompanyName_F")
	sOrderType = rRS.Collect("OrderType")
	sCompanyPictureFlag = rRS.Collect("CompanyPictureFlag")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	sCompanySpeciality = rRS.Collect("CompanySpeciality")

	If sOrderType = "0" And sCompanyKbn = "4" Then sAddTitle = "派遣企業の"

	Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'------------------------------------------------------------------------------
	'会社名 end
	'******************************************************************************

	'******************************************************************************
	'設立年度 start
	'------------------------------------------------------------------------------
	sEstablishYear = ""
	sEstablishYear = rRS.Collect("EstablishYear")
	If sEstablishYear <> "" Then sEstablishYear = sEstablishYear & "年"
	'------------------------------------------------------------------------------
	'設立年度 end
	'******************************************************************************

	'******************************************************************************
	'資本額 start
	'------------------------------------------------------------------------------
	sCapitalAmount = ""
	sCapitalAmount = rRS.Collect("CapitalAmount")
	If IsNumeric(sCapitalAmount) = True Then sCapitalAmount = GetJapaneseYen(sCapitalAmount)
	'------------------------------------------------------------------------------
	'資本額 end
	'******************************************************************************

	'******************************************************************************
	'株式公開 start
	'------------------------------------------------------------------------------
	sListClass = ""
	sListClass = rRS.Collect("ListClass")
	'------------------------------------------------------------------------------
	'株式公開 end
	'******************************************************************************

	'******************************************************************************
	'社員数 start
	'------------------------------------------------------------------------------
	sEmployeeNum = ""
	If ChkStr(rRS.Collect("ManEmployeeNum")) <> "" Or ChkStr(rRS.Collect("WomanEmployeeNum")) <> "" Then
		If rRS.Collect("ManEmployeeNum") <> "" Then
			sEmployeeNum = sEmployeeNum & "男性" & rRS.Collect("ManEmployeeNum") & "人"
		End If
		If ChkStr(rRS.Collect("WomanEmployeeNum")) <> "" Then
			If sEmployeeNum <> "" Then sEmployeeNum = sEmployeeNum & "　"
			sEmployeeNum = sEmployeeNum & "女性" & rRS.Collect("WomanEmployeeNum") & "人"
		End If
		sEmployeeNum = "(" & sEmployeeNum & ")"
	End If
	If ChkStr(rRS.Collect("AllEmployeeNum")) <> "" Then
		sEmployeeNum = rRS.Collect("AllEmployeeNum") & "人" & sEmployeeNum
	End If
	'------------------------------------------------------------------------------
	'社員数 end
	'******************************************************************************

	'******************************************************************************
	'業種 start
	'------------------------------------------------------------------------------
	sIndustryType = ""
	sIndustryType = rRS.Collect("IndustryTypeName")
	'------------------------------------------------------------------------------
	'株式公開 end
	'******************************************************************************

	'******************************************************************************
	'本社住所 start
	'------------------------------------------------------------------------------
	sAddress = ""
	If rRS.Collect("Post_U") & rRS.Collect("Post_L") <> "" Then
		sAddress = "〒" & rRS.Collect("Post_U") & "-" & rRS.Collect("Post_L") & "<br>"
	End If
	sAddress = sAddress & rRS.Collect("Address")
	'------------------------------------------------------------------------------
	'本社住所 end
	'******************************************************************************

	'******************************************************************************
	'ホームページ start
	'------------------------------------------------------------------------------
	sHomePage = ""
	If rRS.Collect("HomepageAddress") <> "" And sOrderType = "0" Then
		sHomePage = rRS.Collect("HomePageAddress")
	End If
	'------------------------------------------------------------------------------
	'ホームページ end
	'******************************************************************************

	If sCompanyPictureFlag = "1" Then
		sClass = "value1"
		sLineClass = "odline2"
	Else
		sClass = "value2"
		sLineClass = "odline1"
	End If

	flgLine = False
%>
<div class="companyblock">
	<h3><%= sAddTitle %>企業情報</h3>
<%
	If sCompanyPictureFlag = "1" Then
%>
	<div style="width:302px; float:right;"><img id="imgcompany" src="<%= HTTPS_NAVI_CURRENTURL %>company/imgdsp.asp?companycode=<%= sCompanyCode %>&amp;optionno=1" alt="イメージ写真" width="300" height="225" style="border:1px solid #999999;"></div>
	<div style="float:left; width:295px;">
<%
	End If

	If sCompanyCode <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>企業コード</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sCompanyCode %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sEstablishYear <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>設立年度</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sEstablishYear %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sCapitalAmount <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>資本額</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sCapitalAmount %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sListClass <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>株式公開</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sListClass %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sEmployeeNum <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>社員数</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sEmployeeNum %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sIndustryType <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>業種</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sIndustryType %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sAddress <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>本社住所</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sAddress %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sHomePage <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>ホームページ</h4></div>
		<div class="<%= sClass %>"><p class="m0"><a href="<%= sHomePage %>" target="_blank">この企業のホームページ</a></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sCompanyPictureFlag = "1" Then
%>
	</div>
	<div style="clear:both;"></div>
<%
	End If
%>
</div>
<%
End Function

'******************************************************************************
'概　要：企業情報のＰＲ情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vOrderCode		：
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'******************************************************************************
Function DspCompanyPR(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Const WELFARECOL = "3"	'福利厚生の１行あたりの列数

	Dim sOrderType			'受注種類
	Dim sCompanyKbn			'企業区分
	Dim sBusiness			'事業内容
	Dim sPR					'企業紹介
	Dim sWelfare			'福利厚生
	Dim iWelfare			'福利厚生カウント
	Dim idx
	Dim flgPR
	Dim flgLine				'線引きフラグ
	Dim sClass
	Dim sAddTitle			'派遣企業の情報の場合は「派遣企業の」を項目名に付ける

	If GetRSState(rRS) = False Then Exit Function

	sOrderType = rRS.Collect("OrderType")
	sCompanyKbn = rRS.Collect("CompanyKbn")

	If sOrderType = "0" And sCompanyKbn = "4" Then sAddTitle = "派遣企業の"

	'******************************************************************************
	'事業内容 start
	'------------------------------------------------------------------------------
	sBusiness = ""
	sBusiness = Replace(ChkStr(rRS.Collect("BusinessContents")), vbCrLf, "<br>")
	sBusiness = Replace(sBusiness, vbCr, "<br>")
	sBusiness = Replace(sBusiness, vbLf, "<br>")
	'------------------------------------------------------------------------------
	'事業内容 end
	'******************************************************************************

	'******************************************************************************
	'会社紹介 start
	'------------------------------------------------------------------------------
	sPR = ""
	sPR = Replace(ChkStr(rRS.Collect("CompanyPR")), vbCrLf, "<br>")
	sPR = Replace(sPR, vbCr, "<br>")
	sPR = Replace(sPR, vbLf, "<br>")
	'------------------------------------------------------------------------------
	'会社紹介 end
	'******************************************************************************

	'******************************************************************************
	'福利厚生 start
	'------------------------------------------------------------------------------
	sWelfare = ""
	iWelfare = 0

	If ChkStr(rRS.Collect("SocietyInsuranceFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">社会保険完備</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("SanatoriumFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">保養所</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("EnterprisePensionFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">企業年金</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("WealthShapeFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">財形貯蓄</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("StockOptionFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">持株制度(ストックオプション)</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("RetirementPayFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">退職金制度</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("ResidencePayFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">住宅手当</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("FamilyPayFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">家族手当</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("EmployeeDormitoryFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">社員寮</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("CompanyHouseFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">社宅</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("NewEmployeeTrainingFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">新入社員研修</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("OverseasTrainingFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">海外研修</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("OtherTrainingFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">各種研修</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("FlexTimeFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">フレックスタイム</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	'中途半端で終わった場合の調整
	If iWelfare Mod WELFARECOL <> 0 Then
		For idx = 1 To (WELFARECOL - iWelfare Mod WELFARECOL)
			sWelfare = sWelfare & "<td class=""welfare""></td>"
		Next
		sWelfare = sWelfare & "</tr>"
	End If

	If sWelfare <> "" Then
		sWelfare = "<table class=""welfare"">" & sWelfare & "</table>"
	End If
	'------------------------------------------------------------------------------
	'福利厚生 end
	'******************************************************************************

	flgPR = False
	If sBusiness & sPR & sWelfare <> "" Then flgPR = True

	flgLine = False
	sClass = "value2"

	If flgPR = True Then
%>
<div class="companyblock">
	<h3><%= sAddTitle %>ＰＲ情報</h3>
<%
		If sBusiness <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
	<div class="category"><h4>事業内容</h4></div>
	<div class="<%= sClass %>"><p class="m0"><%= sBusiness %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sPR <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
	<div class="category"><h4>会社ＰＲ</h4></div>
	<div class="<%= sClass %>"><p class="m0"><%= sPR %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWelfare <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
	<div class="category"><h4>福利厚生</h4></div>
	<div class="<%= sClass %>"><p class="m0"><%= sWelfare %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<br>
<%
	End If
End Function

'******************************************************************************
'概　要：求人票詳細ページのリスの紹介先・派遣先企業情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspLisOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'受注区分
	Dim sListClass			'株式公開
	Dim sIndustryType		'業種
	Dim sPR					'事業内容
	Dim sImgTitle			'タイトルイメージ
	Dim sIntrDisp			'派遣 or 紹介文言
	Dim flgDsp
	Dim flgLine				'線引きフラグ

	DspLisOrderCompanyInfo = False

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		If sOrderType = "2" Then
			sImgTitle = "/img/order/lisorderpr2.gif"
			sIntrDisp = "紹介先"
		Else
			sImgTitle = "/img/order/lisorderpr1.gif"
			sIntrDisp = "派遣先"
		End If
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************

		'******************************************************************************
		'株式公開 start
		'------------------------------------------------------------------------------
		sListClass = ""
		sListClass = rRS.Collect("ListClass")
		'------------------------------------------------------------------------------
		'株式公開 end
		'******************************************************************************

		'******************************************************************************
		'業種 start
		'------------------------------------------------------------------------------
		sIndustryType = ""
		sIndustryType = ChkStr(rRS.Collect("IndustryTypeName"))
		'------------------------------------------------------------------------------
		'株式公開 end
		'******************************************************************************

		'******************************************************************************
		'会社紹介 start
		'------------------------------------------------------------------------------
		sPR = ""
		sPR = Replace(ChkStr(rRS.Collect("BusinessContents")), vbCrLf, "<br>")
		sPR = Replace(sPR, vbCr, "<br>")
		sPR = Replace(sPR, vbLf, "<br>")
		'------------------------------------------------------------------------------
		'会社紹介 end
		'******************************************************************************
	End If

	flgLine = False

	If sListClass & sIndustryType & sPR <> "" Then
		DspLisOrderCompanyInfo = True
%>
<h3><%= sIntrDisp %>企業情報</h3>
<%
		If sListClass <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>株式公開</h4></div>
<div class="value1"><p class="m0"><%= sListClass %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>業種</h4></div>
<div class="value1"><p class="m0"><%= sIndustryType %></p></div>
<div style="clear:both;"></div>
<%
		End If


		If sPR <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			

%>
<div class="category1"><h4>事業内容</h4></div>
<div class="value1"><p class="m0"><%= sPR %></p></div>
<div style="clear:both;"></div>
<% End If %>
				<p class="m0" style="font-size:10px;margin:0 0 20px 20px;">
				※人材<%= left(sIntrDisp,2) %>でご案内するお仕事のため、詳しい会社情報は下のボタンやお電話などで直接お問合せください。
		</p>
<%
	End If
End Function

'******************************************************************************
'概　要：派遣企業の派遣先企業情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder		：自社求人票フラグ
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'******************************************************************************
Function DspTempOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Dim sCompanyCode		'企業コード
	Dim sCompanyName		'会社名
	Dim sCompanyName_F		'会社名カナ
	Dim sAddress			'住所
	Dim sTel				'電話番号
	Dim sIndustryType		'業種
	Dim sCapitalAmount		'資本額
	Dim sListClass			'株式公開
	Dim sEmployeeNum		'社員数
	Dim flgLine				'線引きフラグ
	Dim flgData				'出力データの有無フラグ

	flgData = False
	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'会社名 start
	'------------------------------------------------------------------------------
	sCompanyName = ChkStr(rRS.Collect("TempCompanyName"))
	sCompanyName_F = ChkStr(rRS.Collect("TempCompanyName_F"))
	
	If sCompanyName_F <> "" Then sCompanyName = sCompanyName & "(" & sCompanyName_F & ")"
	'------------------------------------------------------------------------------
	'会社名 end
	'******************************************************************************

	'******************************************************************************
	'住所 start
	'------------------------------------------------------------------------------
	sAddress = ""
	If rRS.Collect("TempPost_U") & rRS.Collect("TempPost_L") <> "" Then
		sAddress = "〒" & rRS.Collect("TempPost_U") & "-" & rRS.Collect("TempPost_L") & "<br>"
	End If
	sAddress = sAddress & rRS.Collect("TempPrefectureName") & rRS.Collect("TempCity") & rRS.Collect("TempTown") & rRS.Collect("TempAddress")
	'------------------------------------------------------------------------------
	'住所 end
	'******************************************************************************

	'******************************************************************************
	'電話番号 start
	'------------------------------------------------------------------------------
	sTel = ChkStr(rRS.Collect("TempTelephoneNumber"))
	'------------------------------------------------------------------------------
	'電話番号 end
	'******************************************************************************

	'******************************************************************************
	'業種 start
	'------------------------------------------------------------------------------
	sIndustryType = ChkStr(rRS.Collect("TempIndustryTypeName"))
	If sIndustryType <> "" Then flgData = True
	'------------------------------------------------------------------------------
	'業種 end
	'******************************************************************************

	'******************************************************************************
	'資本額 start
	'------------------------------------------------------------------------------
	sCapitalAmount = ChkStr(rRS.Collect("TempCapitalAmount"))
	sCapitalAmount = GetJapaneseYen(sCapitalAmount)
	If sCapitalAmount <> "" Then flgData = True
	'------------------------------------------------------------------------------
	'資本額 end
	'******************************************************************************

	'******************************************************************************
	'株式公開 start
	'------------------------------------------------------------------------------
	sListClass = ChkStr(rRS.Collect("TempListClass"))
	If sListClass <> "" Then flgData = True
	'------------------------------------------------------------------------------
	'株式公開 end
	'******************************************************************************

	'******************************************************************************
	'社員数 start
	'------------------------------------------------------------------------------
	sEmployeeNum = ChkStr(rRS.Collect("TempAllEmployeeNumber"))
	If sEmployeeNum <> "" Then sEmployeeNum = sEmployeeNum & "人"
	If sEmployeeNum <> "" Then flgData = True
	'------------------------------------------------------------------------------
	'社員数 end
	'******************************************************************************

	flgLine = False

	If flgData = True Then
%>
<h3>派遣先企業情報</h3>
<%
		If vMyOrder = "1" Then
%>
<p class="m0" style="margin:0px 0px 10px 20px;">※企業名、住所、電話番号は非公開情報です。</p>
<%
			If sCompanyName <> "" Then
				If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
				flgLine = True
%>
<div class="category1"><h4>企業名</h4></div>
<div class="value1"><p class="m0"><%= sCompanyName %></p></div>
<div style="clear:both;"></div>
<%
			End If

			If sAddress <> "" Then
				If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
				flgLine = True
%>
<div class="category1"><h4>住所</h4></div>
<div class="value1"><p class="m0"><%= sAddress %></p></div>
<div style="clear:both;"></div>
<%
			End If

			If sTel <> "" Then
				If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
				flgLine = True
%>
<div class="category1"><h4>電話番号</h4></div>
<div class="value1"><p class="m0"><%= sTel %></p></div>
<div style="clear:both;"></div>
<%
			End If
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>業種</h4></div>
<div class="value1"><p class="m0"><%= sIndustryType %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>資本額</h4></div>
<div class="value1"><p class="m0"><%= sCapitalAmount %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>株式公開</h4></div>
<div class="value1"><p class="m0"><%= sListClass %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>社員数</h4></div>
<div class="value1"><p class="m0"><%= sEmployeeNum %></p></div>
<div style="clear:both;"></div>
<%
		End If

		Response.Write "<br>"
	End If
End Function

'******************************************************************************
'概　要：求人票コントロールボタン
'引　数：rDB				：接続中のDBConnection
'　　　：rRS				：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType			：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID			：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder			：自社求人票か否か ["1"]自社求人票 ["0"]自社求人票でない
'　　　：vJobTypeLimitFlag	：職種数が制限を越えていないか ["1"]OK ["0"]NO
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderControlButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder, ByVal vJobTypeLimitFlag)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError
	Dim sOrderCode
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'受注種類
	Dim sPermitFlag			'掲載許可フラグ
	Dim sPublicFlag			'掲載フラグ
	Dim sRiyoFlag			'掲載開始日
	Dim sHakouFlag			'利用開始日（ライセンス発効日）
	Dim flgAddWatchList
	Dim iMailTemplateCnt	'メールテンプレートの件数
	Dim sAncMT				'メールテンプレートへのリンク

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	sPermitFlag = rRS.Collect("PermitFlag")
	sPublicFlag = rRS.Collect("PublicFlag")
	sRiyoFlag = rRS.Collect("RiyoFlag")
	sHakouFlag = rRS.Collect("HakouFlag")
	iMailTemplateCnt = rRS.Collect("MailTemplateCnt")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "up_GetDataWatchList '" & vUserID & "', '', '', '" & sOrderCode & "', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = False Then
		flgAddWatchList = True
	End If
	Call RSClose(oRS2)
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	If vMyOrder = "1" Then
		'******************************************************************************
		'自社求人票の場合 start
		'------------------------------------------------------------------------------
		If sHakouFlag = "1" Then
%>
<h2 class="csubtitle">自社求人票の操作</h2>
<div class="subcontent">
<%
			If sPermitFlag = "1" And sRiyoFlag = "0" Then
%>
	<p class="cctrltitle">求職者検索・スカウトメール</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<p style="color:#ff0000;">この求人票はまだ掲載されておりません（掲載開始日前です）。そのため、求職者の検索は利用できません。</p>
		</div>
	</div>
<%
			ElseIf sPermitFlag = "0" Then
%>
	<p class="cctrltitle">求職者検索・スカウトメール</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<p style="color:#ff0000;">この求人票はまだ掲載されておりません（審査中です）。そのため、求職者の検索は利用できません。</p>
		</div>
	</div>
<%
			ElseIf sPermitFlag = "1" And sPublicFlag = "1" And sRiyoFlag = "1" Then
%>
	<p class="cctrltitle">求職者検索・スカウトメール</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<input type="button" value="求職者を自動検索" style="width:150px; color:#aa3300;" onclick="Go_Edit('10');">
			<span style="font-size:10px; color:#666666;">・・・この求人票の、職種・勤務地・雇用形態を満たす求職者を検索します。</span>
		</div>
		<div style="padding:0px 0px 5px 15px;">
			<input type="button" value="求職者を詳細検索" style="width:150px; color:#aa3300;" onclick="Go_Edit('11');">
			<span style="font-size:10px; color:#666666;">・・・この求人票から、詳細な検索条件を指定して求職者を検索します。</span><br>
		</div>
	</div>
<%
			Else
%>
	<p class="cctrltitle">求職者検索・スカウトメール</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<p style="color:#ff0000;">掲載されていない求人票からのスカウトはできません。</p>
		</div>
	</div>
<%
			End If

			If vJobTypeLimitFlag = True Then
				'職種数が制限を越えていなければ「求人票コピー作成」ボタンの表示
%>
	<p class="cctrltitle">求人票コピー作成</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<input type="button" value="求人票をコピー" style="width:100px; color:#3333ff;" onclick="Go_Edit('4');">
			<span style="font-size:10px; color:#666666;">・・・この求人票をもとに、新しい求人票を作成します。</span><br>
		</div>
	</div>
<%
			End If
%>
	<p class="cctrltitle">求人情報を編集する</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<div style="float:left; width:290px;">
				<input type="button" value="自社情報更新" style="width:100px;" onclick="Go_Edit('1');">
				<span style="font-size:10px; color:#666666;">・・・自社情報を更新します。</span>
			</div>
			<div style="float:right; width:290px;">
				<input type="button" value="画像登録" style="width:100px;" onclick="Go_Edit('5');">
				<span style="font-size:10px; color:#666666;">・・・画像を掲載します。</span><br>
			</div>
			<div style="clear:both;"></div>
		</div>
		<div style="padding:0px 0px 5px 15px;">
			<div style="float:left; width:290px; margin:0px;">
				<input type="button" value="募集情報編集" style="width:100px;" onclick="Go_Edit('2');">
				<span style="font-size:10px; color:#666666;">・・・ＰＲ・募集要項を編集します。</span>
			</div>
			<div style="float:right; width:290px;">
				<input type="button" value="スキル条件編集" style="width:100px;" onclick="Go_Edit('3');">
				<span style="font-size:10px; color:#666666;">・・・必要スキル・資格を編集します。</span><br>
			</div>
			<div style="clear:both;"></div>
		</div>
	</div>

	<p class="cctrltitle">メールテンプレート</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
<%
			If iMailTemplateCnt >= 5 Then
				'メールテンプレート数が上限に達している場合は新規作成できない
%>
			<p style="color:#ff0000; font-size:10px;">メールテンプレート数が上限に達しているので、これ以上作成できません。</p>
<%
			Else
				'メールテンプレート数が上限に達していない場合は新規作成できる
%>
			<input type="button" value="新規作成" style="width:100px;" onclick="location.href = '<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/regist.asp?ordercode=<%= sOrderCode %>';">
			<span style="font-size:10px; color:#666666;">・・・この求人のメールテンプレートを新規に作成します。</span><br>
<%
			End If
%>
			<p style="color:#ff0000; font-size:10px;">※メールテンプレートは求人票毎に作成します。</p>
<%
			sSQL = "up_GetListMailTemplate '" & G_USERID & "', '" & sOrderCode & "'"
			flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
			If GetRSState(oRS2) = True Then Response.Write "<hr size=""1"">"
			Do While GetRSState(oRS2) = True
				sAncMT = "?ordercode=" & oRS2.Collect("OrderCode") & "&amp;seq=" & oRS2.Collect("SEQ")
				sAncMT = "<a href=""" & HTTPS_NAVI_CURRENTURL & "mailtemplate/regist.asp" & sAncMT & """>" & oRS2.Collect("Subject") & "</a>"
%>
			<div style="width:585px;">
				<div style="float:left; width:85px;"><%= GetDetail("MailTemplateType", oRS2.Collect("MailTemplateTypeCode")) %></div>
				<div style="float:left; width:500px;"><%= sAncMT %></div>
				<div style="clear:both;"></div>
			</div>
<%
				oRS2.MoveNext
			Loop
%>
		</div>
	</div>
</div>
<%
		End If
		'------------------------------------------------------------------------------
		'自社求人票の場合 end
		'******************************************************************************
	ElseIf vUserType = "staff" Then
		'******************************************************************************
		'ログイン求職者の場合 start
		'------------------------------------------------------------------------------
		If rRS.Collect("PublicFlag") = "1" Then
%>
<div class="subcontent" style="margin-bottom:15px;">
	<div style="padding:5px 0px;">
		<p class="sctrltitle">応募・質問・ウォッチリスト</p>
		<div style="padding:0px 0px 5px 15px;">
			<div style="float:left; width:195px;">
				<p class="m0" style="margin-right:20px; font-size:10px; color:#666666; text-align:center;">▼この募集へ応募メールの作成</p>
				<input type="button" value="応募メールを送信する" style="width:180px;" onclick="contactCompany('');">
			</div>
			<div align="center" style="float:left; width:195px;">
				<p class="m0" style="font-size:10px; color:#666666; text-align:center;">▼この募集へ質問メールの作成</p>
				<input type="button" value="質問メールを送信する" onclick="contactCompany('1');">
			</div>
			<div style="float:left; width:195px;">
				<p class="m0" style="margin-left:20px; font-size:10px; color:#666666; text-align:center;">▼<a href="watchlist_info.htm" onclick="window.open(this.href, 'mywindow6', 'width=300, height=150, menubar=no, toolbar=no, scrollbars=yes'); return false;" style="color:#0045F9;">ウォッチリスト</A>へ追加</p>
<%
			If flgAddWatchList = True Then
%>
				<div align="right"><input type="button" value="この求人票を追加する" style="width:180px;" onclick="document.forms.frmMain.action='../staff/watchlist_register.asp';document.forms.frmMain.submit();"></div>
<%
			Else
%>
				<p class="m0" style="margin-left:20px; text-align:center; font-weight:bold;">既に登録済みです</p>
<%
			End If
%>
			</div>
			<div style="clear:both;"></div>
		</div>
	</div>
</div>
<%
		Else
%>
	<div align="center"><b>この求人票は掲載が終了しています。メール送信はできません。</b></div>
<%
		End If
		'------------------------------------------------------------------------------
		'ログイン求職者の場合 end
		'******************************************************************************
	End If
End Function

'******************************************************************************
'概　要：求人票詳細ページのコントロールボタンで利用するjavascriptの出力
'　　　：自社求人票 or ログイン中の求職者の場合は、編集ボタン or メール送信ボタンを処理する
'　　　：javascriptを出力
'引　数：rDB				：接続中のDBConnection
'　　　：rRS				：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType			：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID			：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder			：自社求人票か否か ["1"]自社求人票 ["0"]自社求人票でない
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function JSOrderControlButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sOrderCode

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'情報コード
		sOrderCode = rRS.Collect("OrderCode")
	End If

	If vMyOrder = "1" Then
		'******************************************************************************
		'自社求人票の場合 start
		'------------------------------------------------------------------------------
%>
<script type="text/javascript">
<!--
function Go_Edit(pNo){
	switch (pNo){
		case '1':
			//「会社情報更新」へ
			location.href = '<%= HTTPS_NAVI_CURRENTURL & vUserType %>/company_reg1.asp';
			return true;
		case '2':
			//「募集情報編集」へ
			document.forms.frmMain.mode.value="edit"
			document.forms.frmMain.action="<%= HTTPS_NAVI_CURRENTURL & vUserType %>/company_reg2.asp";
			break;
		case '3':
			//「スキル」へ
			document.forms.frmMain.mode.value="edit"
			document.forms.frmMain.action="<%= HTTPS_NAVI_CURRENTURL & vUserType %>/company_reg3.asp";
			break;
		case '4':
			//「コピーして求人票の作成」へ
			document.forms.frmMain.mode.value="copy"
			document.forms.frmMain.action="<%= HTTPS_NAVI_CURRENTURL & vUserType %>/company_reg2.asp";
			break;
		case '5':
			//「求人票写真登録」へ
			location.href = '<%= HTTP_NAVI_CURRENTURL %>company/order_img_listnow.asp?ordercode=<%= sOrderCode %>';
			return true;
		case '10':
			//自動検索
			document.forms.frmMain.action="<%= HTTP_NAVI_CURRENTURL %>staff/person_list.asp";
			break;
		case '11':
			//詳細検索
			document.forms.frmMain.action="<%= HTTP_NAVI_CURRENTURL %>staff/person_search_detail.asp";
			break;
		default:
			return false;
	}
	document.forms.frmMain.submit();
}
//-->
</script>
<%
		'------------------------------------------------------------------------------
		'自社求人票の場合 end
		'******************************************************************************
	ElseIf vUserType = "staff" Then
		'******************************************************************************
		'ログイン求職者の場合 start
		'------------------------------------------------------------------------------
		If rRS.Collect("PublicFlag") = "1" Then
%>
<script type="text/javascript">
function contactCompany(vflg) {
	var sQ = '';
	if(vflg){
		if(vflg.length > 0)sQ = 'q=1&';
	}
	MailWin = window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompany.asp?' + sQ + 'ordercode=<%= sOrderCode %>','mail','width=480,height=580,resizable=1,scrollbars=no');
}
</script>
<%
		End If
		'------------------------------------------------------------------------------
		'ログイン求職者の場合 end
		'******************************************************************************
	End If
End Function

'******************************************************************************
'概　要：求人票詳細ページのコントロールボタンで使用するFORMデータを出力
'引　数：rDB				：接続中のDBConnection
'　　　：rRS				：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType			：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID			：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder			：自社求人票か否か ["1"]自社求人票 ["0"]自社求人票でない
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function FrmOrderControlButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sOrderCode
	Dim sCompanyCode		'企業コード
	Dim sOrderType

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sOrderCode = rRS.Collect("OrderCode")
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************
	End If
%>
	<form id="frmMain" action="./" method="post">
	<input type="hidden" name="CONF_OrderCode" value="<%= sOrderCode %>">
	<input type="hidden" name="CONF_CompanyCode" value="<%= sCompanyCode %>">
	<input type="hidden" name="CONF_OrderType" value="<%= sOrderType %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="CONF_SearchMode" value="">
	</form>
<%
End Function

'******************************************************************************
'概　要：求人票の企業名称を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderCompanyName(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType
	Dim sCompanyCode		'企業コード
	Dim sCompanyName		'企業名称
	Dim sCompanyNameF		'企業名称カナ
	Dim sCompanyKbn			'企業区分
	Dim sCompanySpeciality	'企業特徴
	Dim sPublishLimitStr	'掲載期限表示用文字列
	Dim sCautionStr			'掲載期限表示注意文言文字列
	Dim flgNowPublic		'現在掲載中の求人票判定 '[True]掲載中 [False]非掲載

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'会社名 start
	'------------------------------------------------------------------------------
	sCompanyName = rRS.Collect("CompanyName")
	sCompanyNameF = rRS.Collect("CompanyName_F")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	sCompanySpeciality = rRS.Collect("CompanySpeciality")
	sOrderType = rRS.Collect("OrderType")

	Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'------------------------------------------------------------------------------
	'会社名 end
	'******************************************************************************

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p style=""line-height:11px;text-align:right;font-size:11px;"">※期限前に掲載終了する場合があります。</p>"

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならriyotodateを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(rRS.Collect("riyotodate"), "/")
	Else
		sPublishLimitStr = rRS.Collect("PublicLimitDay")
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = 0 Then
			'ライセンス切れのときは"掲載終了"と表示
			sPublishLimitStr = "掲載終了"
			sCautionStr = ""
		Else
			sPublishLimitStr = "常時募集中"
		End If
	End If
	'------------------------------------------------------------------------------
	'求人票掲載期限 end
	'******************************************************************************
%>
<div style="width:600px; margin-bottom:10px;">
<%
	If sOrderType = "2" Then
		'リス紹介案件の場合は「転職相談案件」イメージを表示
%>
	<img src="/img/order/counselable_order.gif" width="150" height="25" alt="転職支援を受けて応募する求人です">
<%
	End If

	If vUserType = "" Or vUserType = "staff" Then
		'非ログイン時、スタッフログイン時

		If G_USERID <> "" And G_FLGRESUME = False And flgNowPublic = True Then
			'しごとナビにログイン中の場合は、企業名＋掲載期限＋求人票ＵＲＬメール送信
%>
	<div class="m0" style="width:420px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="float:right; padding:0px;"><img src="../ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode"></div>
	<div style="text-align:right; font-size:11px; padding-top:6px;"><a href="../order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=470');return false;"><img src="/img/staff/mail/mailhei.gif" border="0" align="bottom" alt="求人票をメール送信"> 求人票をメール送信</a></div>
	<p style="text-align:right;padding:4px 0px 0px 0px;">掲載期限：<%= sPublishLimitStr %></p>
	<div style="clear:both;"></div>
	<%= sCautionStr %>
	<div style="clear:both;"></div>
<%
		ElseIf G_FLGRESUME = False And flgNowPublic = True Then
			'しごとナビに非ログインの場合は、企業名＋掲載期限＋求人票ＵＲＬメール送信
%>
	<div class="m0" style="width:420px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="float:right; padding:0px;"><img src="../ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode"></div>
	<div style="text-align:right; font-size:11px; padding-top:6px;"><a href="../order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=380');return false;"><img src="/img/staff/mail/mailhei.gif" border="0" align="bottom" alt="求人票をメール送信"> 求人票をメール送信</a></div>
	<p style="text-align:right;padding:4px 0px 0px 0px;">掲載期限：<%= sPublishLimitStr %></p>
	<div style="clear:both;"></div>
	<%= sCautionStr %>
	<div style="clear:both;"></div>
<%
		Else
			'＠履歴書の求人票の場合は、企業名＋掲載期限のみ
%>
	<div class="m0" style="width:420px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="float:right; padding:0px;"><img src="../ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode"></div>
	<p style="text-align:right;padding-top:21px;">掲載期限：<%= sPublishLimitStr %></p>
	<div style="clear:both;"></div>
	<%= sCautionStr %>
	<div style="clear:both;"></div>
<%
		End If
	Else
%>
	<div class="m0" style="width:420px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="float:right; padding:0px;"><img src="../ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode"></div>
	<p style="text-align:right;padding-top:21px;">掲載期限：<%= sPublishLimitStr %></p>
	<div style="clear:both;"></div>
	<%= sCautionStr %>
	<div style="clear:both;"></div>
<%
	End If
%>
</div>
<%
End Function

'******************************************************************************
'概　要：求人票詳細ページの会社情報・職種情報切り替えボタンと参照回数を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vType			：表示中情報の種類 ["0"]職種情報 ["1"]会社情報
'　　　：vAccessCount	：表示中求人票のアクセス回数
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderShowTypeSwitch(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vType, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sOrderType
	Dim sJobTypeDetail
	Dim sUpdateDay

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sOrderType = rRS.Collect("OrderType")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'具体的職種名
	sJobTypeDetail = rRS.Collect("JobTypeDetail")
	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	If sJobTypeDetail <> "" Then sJobTypeDetail = sJobTypeDetail & "のお仕事情報詳細"
%>
<div style="width:600px; margin-bottom:5px;">
	<div style="float:left; width:350px; margin:0px;">
<%
	If vType = "0" Then
		'仕事情報を表示中の場合
%>
		<div style="float:left; width:93px; margin:0px;"><img src="/img/order/tab_orderdetail_on.gif" alt="<%= sJobTypeDetail %>" border="0" width="93" height="22"></div>
<%
		If sOrderType = "0" Then
			'一般の求人広告の場合は会社情報へのリンクを表示
%>
		<div style="float:left; width:93px; margin:0px;"><a href="./company_order.asp?poc=<%= sOrderCode %>" title="会社情報"><img src="/img/order/tab_companyinfo_off.gif" alt="会社情報" border="0" width="93" height="22"></a></div>
<%
		End If
	ElseIf vType = "1" Then
		'会社情報を表示中の場合
%>
		<div style="float:left; width:93px; margin:0px;"><a href="./order_detail.asp?ordercode=<%= sOrderCode %>"><img src="/img/order/tab_orderdetail_off.gif" alt="<%= sJobTypeDetail %>" border="0" width="93" height="22"></a></div>
<%
		If sOrderType = "0" Then
			'一般の求人広告の場合は会社情報へのリンクを表示
%>
		<div style="float:left; width:93px; margin:0px;"><img src="/img/order/tab_companyinfo_on.gif" alt="会社情報" border="0" width="93" height="22"></div>
<%
		End If
	End If
%>
		<div class="clear:both; margin:0px;"></div>
	</div>
	<div align="right" style="float:right; width:250px;">
		<p class="m0">月間参照回数：<%= vAccessCount %>回　更新日：<%= sUpdateDay %></p>
	</div>
	<div style="clear:both;"><img src="/img/order/tab_border.gif" alt="" width="600" height="5"></div>
</div>
<%
End Function

'******************************************************************************
'概　要：求人票のキャッチコピー部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'　　　：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderCatchCopy(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType
	Dim sCompanyCode
	Dim sOrderCode

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1
	Dim sClass

	If GetRSState(rRS) = False Then Exit Function

	sOrderType = rRS.Collect("OrderType")
	sOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")

	'******************************************************************************
	'大きい画像 start
	'------------------------------------------------------------------------------
	sOptionNo = ""
	sImg1 = ""
	sSQL = "up_GetListOrderPictureNow '" & sCompanyCode & "', '" & sOrderCode & "', 'orderpicture'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
			sOptionNo = oRS.Collect("OptionNo1")
			sImg1 = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & sOptionNo
		End If
	End If

	If sImg1 = "" And sOrderType = "0" Then
		sSQL = "sp_GetDataPicture '" & sCompanyCode & "', '1'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sImg1 = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"
		End If
	End If
	'------------------------------------------------------------------------------
	'大きい画像 end
	'******************************************************************************

	If sImg1 <> "" Then
%>
<div id="catchcopy" style="width:600px;">
	<div style="float:right; width:300px;"><img class="big" src="<%= sImg1 %>" alt="" border="1" width="300" height="225" style="border:1px solid #999999;"></div>
	<h2><%= rRS.Collect("JobTypeDetail") %></h2>
	<div style="margin:10px 0px;"><%= GetImgOrderSpeciality(rDB, rRS) %></div>
	<p class="m0"><%= rRS.Collect("CatchCopy") %></p>
	<br clear="all">
</div>
<%
	Else
%>
<div id="catchcopy" style="width:600px;">
	<h2 style="width:600px;"><%= rRS.Collect("JobTypeDetail") %></h2>
	<div style="margin:10px 0px;"><%= GetImgOrderSpeciality(rDB, rRS) %></div>
	<p class="m0" style="width:600px;"><%= rRS.Collect("CatchCopy") %></p>
	<div style="clear:both;"></div>
</div>
<%
	End If
End Function

'******************************************************************************
'概　要：求人票詳細ページのフリーＰＲを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'　　　：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderFreePR(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sPRTitle1			'ＰＲタイトル1
	Dim sPRTitle2			'ＰＲタイトル2
	Dim sPRTitle3			'ＰＲタイトル3
	Dim sPRContents1		'ＰＲ文1
	Dim sPRContents2		'ＰＲ文2
	Dim sPRContents3		'ＰＲ文3
	Dim flgPR				'ＰＲ有無フラグ [True]有 [False]無

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'PR start
	'------------------------------------------------------------------------------
	flgPR = False
	sPRTitle1 = ChkStr(rRS.Collect("PRTitle1"))
	sPRTitle2 = ChkStr(rRS.Collect("PRTitle2"))
	sPRTitle3 = ChkStr(rRS.Collect("PRTitle3"))
	sPRContents1 = Replace(ChkStr(rRS.Collect("PRContents1")), vbCrLf, "<br>")
	sPRContents1 = Replace(sPRContents1, vbCr, "<br>")
	sPRContents1 = Replace(sPRContents1, vbLf, "<br>")
	sPRContents2 = Replace(ChkStr(rRS.Collect("PRContents2")), vbCrLf, "<br>")
	sPRContents2 = Replace(sPRContents2, vbCr, "<br>")
	sPRContents2 = Replace(sPRContents2, vbLf, "<br>")
	sPRContents3 = Replace(ChkStr(rRS.Collect("PRContents3")), vbCrLf, "<br>")
	sPRContents3 = Replace(sPRContents3, vbCr, "<br>")
	sPRContents3 = Replace(sPRContents3, vbLf, "<br>")

	If sPRTitle1 & sPRTitle2 & sPRTitle3 & sPRContents1 & sPRContents2 & sPRContents3 <> "" Then flgPR = True
	'------------------------------------------------------------------------------
	'PR end
	'******************************************************************************

	If flgPR = True Then
%>
	<h3>ＰＲ</h3>
	<div class="freeprblock">
<%
		If sPRTitle1 <> "" Or sPRContents1 <> "" Then
%>
		<h4><%= sPRTitle1 %></h4>
		<div style="clear:both;"></div>
		<p class="m0"><%= sPRContents1 %></p>
<%
		End If

		If sPRTitle2 <> "" Or sPRContents2 <> "" Then
%>
		<h4><%= sPRTitle2 %></h4>
		<div style="clear:both;"></div>
		<p class="m0"><%= sPRContents2 %></p>
<%
		End If

		If sPRTitle3 <> "" Or sPRContents3 <> "" Then
%>
		<h4><%= sPRTitle3 %></h4>
		<div style="clear:both;"></div>
		<p class="m0"><%= sPRContents3 %></p>
<%
		End If
%>
	</div>
<%
	End If
End Function

'******************************************************************************
'概　要：求人企業画像一覧表示ＨＴＭＬ表示
'作成者：Lis Kokubo
'作成日：2006/12/27
'引　数：vCompanyCode	：企業コード
'　　　：vOrderCode		：情報コード
'　　　：vCategoryCode	：カテゴリコード
'使用先：
'備　考：
'******************************************************************************
Function DspOrderPictureNow(ByVal vCompanyCode, ByVal vOrderCode, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sURL
	Dim flgPicture

	flgPicture = False
	sSQL = "up_GetListOrderPictureNow '" & vCompanyCode & "', '" & vOrderCode & "', '" & vCategoryCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
		If Len(oRS.Collect("OptionNo2")) > 0 Or Len(oRS.Collect("OptionNo3")) > 0 Or Len(oRS.Collect("OptionNo4")) > 0 Then
%>
<div align="center" style="padding:5px 15px; background-color:#e1fbcd; margin-bottom:40px;">
<div style="width:570px;">
<%
			sURL = ""
			If Len(oRS.Collect("OptionNo2")) > 0 Then
				sURL = "/company/imgdsp.asp?companycode=" & vCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
%>
<div align="right" style="float:left; width:190px;">
	<div style="width:182px; background-color:#ffffff;"><img src="<%= sURL %>" alt="<%= oRS.Collect("Caption2") %>" width="180" height="135" border="1" style="border:1px solid #999999;"></div>
	<p class="m0" align="left" style="width:182px; font-size:10px;"><%= oRS.Collect("Caption2") %></p>
</div>
<%
			End If

			sURL = ""
			If Len(oRS.Collect("OptionNo3")) > 0 Then
				sURL = "/company/imgdsp.asp?companycode=" & vCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
%>
<div align="right" style="float:left; width:190px;">
	<div style="width:182px; background-color:#ffffff;"><img src="<%= sURL %>" alt="<%= oRS.Collect("Caption3") %>" width="180" height="135" border="1" style="border:1px solid #999999;"></div>
	<p class="m0" align="left" style="width:182px; font-size:10px;"><%= oRS.Collect("Caption3") %></p>
</div>
<%
			End If

			sURL = ""
			If Len(oRS.Collect("OptionNo4")) > 0 Then
				sURL = "/company/imgdsp.asp?companycode=" & vCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
%>
<div align="right" style="float:left; width:190px;">
	<div style="width:182px; background-color:#ffffff;"><img src="<%= sURL %>" alt="<%= oRS.Collect("Caption4") %>" width="180" height="135" border="1" style="border:1px solid #999999;"></div>
	<p class="m0" align="left" style="width:182px; font-size:10px;"><%= oRS.Collect("Caption4") %></p>
</div>
<%
			End If

			Response.Write "<br clear=""all"">"
%>
</div>
</div>
<%
		End If
	End If
End Function

'******************************************************************************
'概　要：求人票の業務内容を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspBusiness(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode			'情報コード
	Dim sCompanyCode		'企業コード
	Dim sBizName1			'仕事割合文言1
	Dim sBizName2			'仕事割合文言2
	Dim sBizName3			'仕事割合文言3
	Dim sBizName4			'仕事割合文言4
	Dim sBizPercentage1		'仕事割合1
	Dim sBizPercentage2		'仕事割合2
	Dim sBizPercentage3		'仕事割合3
	Dim sBizPercentage4		'仕事割合4
	Dim sBiz				'仕事割合HTML
	Dim sBusinessDetail		'担当業務
	Dim sClearSolid
	Dim flgBusiness
	Dim flgLine				'線引きフラグ

	If GetRSState(rRS) = False Then Exit Function

	flgBusiness = False
	If GetRSState(rRS) = True Then

		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sOrderCode = rRS.Collect("OrderCode")
		sCompanyCode = rRS.Collect("CompanyCode")
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************

		'******************************************************************************
		'仕事の割合 start
		'------------------------------------------------------------------------------
		sBiz = ""
		sBizName1 = ""
		sBizName2 = ""
		sBizName3 = ""
		sBizName4 = ""
		sBizPercentage1 = ""
		sBizPercentage2 = ""
		sBizPercentage3 = ""
		sBizPercentage4 = ""

		sBizName1 = ChkStr(rRS.Collect("BizName1"))
		sBizName2 = ChkStr(rRS.Collect("BizName2"))
		sBizName3 = ChkStr(rRS.Collect("BizName3"))
		sBizName4 = ChkStr(rRS.Collect("BizName4"))
		sBizPercentage1 = ChkStr(rRS.Collect("BizPercentage1"))
		sBizPercentage2 = ChkStr(rRS.Collect("BizPercentage2"))
		sBizPercentage3 = ChkStr(rRS.Collect("BizPercentage3"))
		sBizPercentage4 = ChkStr(rRS.Collect("BizPercentage4"))

		If sBizName1 & sBizName2 & sBizName3 & sBizName4 <> "" Then
			If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""></td></tr>"
			If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""></td></tr>"
			If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""></td></tr>"
			If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""></td></tr>"
			sBiz = "<table>" & sBiz & "</table>"
			flgBusiness = True
		End If
		'------------------------------------------------------------------------------
		'仕事の割合 end
		'******************************************************************************

		'******************************************************************************
		'担当業務 start
		'------------------------------------------------------------------------------
		sBusinessDetail = Replace(ChkStr(rRS.Collect("BusinessDetail")), vbCrLf, "<br>")
		sBusinessDetail = Replace(sBusinessDetail, vbCr, "<br>")
		sBusinessDetail = Replace(sBusinessDetail, vbLf, "<br>")
		If sBusinessDetail <> "" Then flgBusiness = True
		'------------------------------------------------------------------------------
		'担当業務 end
		'******************************************************************************
	End If

	flgLine = False
	If flgBusiness = True Then
%>
<h3>業務内容</h3>
<%
		If sBusinessDetail <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>担当業務</h4></div>
<div class="value1"><p class="m0"><%= sBusinessDetail %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sBiz <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>仕事の割合</h4></div>
<%'<div class="value1"><%= sBiz % ></div>%>
<div class="value1">
	<table border="0">
		<tbody>
		<tr>
			<td>
<script type="text/javascript" language="javascript">
	viewWorkAvg(<%= sBizPercentage1 %>, <%= sBizPercentage2 %>, <%= sBizPercentage3 %>, <%= sBizPercentage4 %>)
</script>
			</td>
			<td style="padding-left:5px; vertical-align:middle;">
				<table border="0">
					<tbody>
<%
			If sBizName1 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#ff9999; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage1 & "%</td><td>" & sBizName1 & "</td></tr>"
			If sBizName2 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#9999ff; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage2 & "%</td><td>" & sBizName2 & "</td></tr>"
			If sBizName3 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#99ff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage3 & "%</td><td>" & sBizName3 & "</td></tr>"
			If sBizName4 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#ffff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage4 & "%</td><td>" & sBizName4 & "</td></tr>"
%>
					</tbody>
				</table>
			</td>
		</tr>
		</tbody>
	</table>
</div>
<div style="clear:both;"></div>
<%
		End If
%>
<br>
<%
	End If
End Function

'******************************************************************************
'概　要：求人票の勤務条件を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode			'情報コード
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'求人票種類
	Dim sCompanyKbn			'企業区分
	Dim sJobTypeDetail		'職種詳細
	Dim sSalary				'給与
	Dim sYearlyIncome		'年収
	Dim sYearlyIncomeMin	'年収
	Dim sYearlyIncomeMax	'年収
	Dim sMonthlyIncome		'月収
	Dim sMonthlyIncomeMin	'月収
	Dim sMonthlyIncomeMax	'月収
	Dim sDailyIncome		'日給
	Dim sDailyIncomeMin		'日給
	Dim sDailyIncomeMax		'日給
	Dim sHourlyIncome		'時給
	Dim sHourlyIncomeMin	'時給
	Dim sHourlyIncomeMax	'時給
	Dim sPercentagePay		'歩合制
	Dim sSalaryRemark		'給与備考
	Dim sTrafficFee			'交通費
	Dim sTrafficFeeType		'
	Dim sTrafficFeeMonth	'交通費／１ヶ月
	Dim sTime				'時間
	Dim sWorkRange			'就業期間
	Dim sWorkStartDay		'就業開始日
	Dim sWorkEndDay			'就業終了日
	Dim sWorkUpdate			'就業期間の更新有無
	Dim sWorkingTime		'就業時間
	Dim sWorkTimeRemark		'就業時間備考
	Dim sHoliday			'休日
	Dim sWeeklyHolidayType	'週休
	Dim sHolidayRemark		'休日備考
	Dim sWorkingPlace		'就業場所
	Dim sWPSection			'勤務先部署
	Dim sWPTel				'勤務先電話番号
	Dim sWPAddress			'勤務先住所
	Dim sMAP				'地図情報
	Dim sNearbyStation		'最寄駅
	Dim sNearbyRailway		'最寄沿線
	Dim sTransfer
	Dim sClearSolid
	Dim flgSalary
	Dim flgTime
	Dim flgHoliday
	Dim flgWorkingPlace
	Dim flgLine
	Dim flgLine2

	DspCondition = False

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sOrderCode = rRS.Collect("OrderCode")
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		sCompanyKbn = rRS.Collect("CompanyKbn")
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************

		'******************************************************************************
		'職種詳細 start
		'------------------------------------------------------------------------------
		sJobTypeDetail = rRS.Collect("JobTypeDetail")
		'------------------------------------------------------------------------------
		'職種詳細 end
		'******************************************************************************

		'******************************************************************************
		'給与 start
		'------------------------------------------------------------------------------
		sYearlyIncomeMin = ChkStr(rRS.Collect("YearlyIncomeMin"))
		sYearlyIncomeMax = ChkStr(rRS.Collect("YearlyIncomeMax"))
		If sYearlyIncomeMin = "0" Then sYearlyIncomeMin = ""
		If sYearlyIncomeMax = "0" Then sYearlyIncomeMax = ""
		If sYearlyIncomeMin <> "" Then sYearlyIncomeMin = GetJapaneseYen(sYearlyIncomeMin)
		If sYearlyIncomeMax <> "" Then sYearlyIncomeMax = GetJapaneseYen(sYearlyIncomeMax)
		If sYearlyIncomeMin & sYearlyIncomeMax <> "" Then
			If sYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMin
			sYearlyIncome = sYearlyIncome & "&nbsp;～&nbsp;"
			If sYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMax
		End If

		sMonthlyIncomeMin = ChkStr(rRS.Collect("MonthlyIncomeMin"))
		sMonthlyIncomeMax = ChkStr(rRS.Collect("MonthlyIncomeMax"))
		If sMonthlyIncomeMin = "0" Then sMonthlyIncomeMin = ""
		If sMonthlyIncomeMax = "0" Then sMonthlyIncomeMax = ""
		If sMonthlyIncomeMin <> "" Then sMonthlyIncomeMin = GetJapaneseYen(sMonthlyIncomeMin)
		If sMonthlyIncomeMax <> "" Then sMonthlyIncomeMax = GetJapaneseYen(sMonthlyIncomeMax)
		If sMonthlyIncomeMin & sMonthlyIncomeMax <> "" Then
			If sMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMin
			sMonthlyIncome = sMonthlyIncome & "&nbsp;～&nbsp;"
			If sMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMax
		End If

		sDailyIncomeMin = ChkStr(rRS.Collect("DailyIncomeMin"))
		sDailyIncomeMax = ChkStr(rRS.Collect("DailyIncomeMax"))
		If sDailyIncomeMin = "0" Then sDailyIncomeMin = ""
		If sDailyIncomeMax = "0" Then sDailyIncomeMax = ""
		If sDailyIncomeMin <> "" Then sDailyIncomeMin = GetJapaneseYen(sDailyIncomeMin)
		If sDailyIncomeMax <> "" Then sDailyIncomeMax = GetJapaneseYen(sDailyIncomeMax)
		If sDailyIncomeMin & sDailyIncomeMax <> "" Then
			If sDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMin
			sDailyIncome = sDailyIncome & "&nbsp;～&nbsp;"
			If sDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMax
		End If

		sHourlyIncomeMin = ChkStr(rRS.Collect("HourlyIncomeMin"))
		sHourlyIncomeMax = ChkStr(rRS.Collect("HourlyIncomeMax"))
		If sHourlyIncomeMin = "0" Then sHourlyIncomeMin = ""
		If sHourlyIncomeMax = "0" Then sHourlyIncomeMax = ""
		If sHourlyIncomeMin <> "" Then sHourlyIncomeMin = GetJapaneseYen(sHourlyIncomeMin)
		If sHourlyIncomeMax <> "" Then sHourlyIncomeMax = GetJapaneseYen(sHourlyIncomeMax)
		If sHourlyIncomeMin & sHourlyIncomeMax <> "" Then
			If sHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMin
			sHourlyIncome = sHourlyIncome & "&nbsp;～&nbsp;"
			If sHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMax
		End If

'		sYearlyIncome = GetMoneyRange(ChkStr(rRS.Collect("YearlyIncomeMin")), ChkStr(rRS.Collect("YearlyIncomeMax")), 1)
'		sMonthlyIncome = GetMoneyRange(ChkStr(rRS.Collect("MonthlyIncomeMin")), ChkStr(rRS.Collect("MonthlyIncomeMax")), 1)
'		sDailyIncome = GetMoneyRange(ChkStr(rRS.Collect("DailyIncomeMin")), ChkStr(rRS.Collect("DailyIncomeMax")), 1)
'		sHourlyIncome = GetMoneyRange(ChkStr(rRS.Collect("HourlyIncomeMin")), ChkStr(rRS.Collect("HourlyIncomeMax")), 1)
		sPercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
		sSalaryRemark = Replace(ChkStr(rRS.Collect("IncomeRemark")), vbCrLf, "<br>")
		sSalaryRemark = Replace(sSalaryRemark, vbCr, "<br>")
		sSalaryRemark = Replace(sSalaryRemark, vbLf, "<br>")
		sTrafficFee = ""
		sTrafficFeeType = ChkStr(rRS.Collect("TrafficFeeType"))
		sTrafficFeeMonth = ChkStr(rRS.Collect("MonthTrafficFee"))
		flgSalary = False

		'給与
		sSalary = ""
		If sYearlyIncome <> "" Then
			sSalary = sSalary & sYearlyIncome
			flgSalary = True
		End If
		If sMonthlyIncome <> "" Then
			If sSalary <> "" Then sSalary = sSalary & "<br>"
			sSalary = sSalary & sMonthlyIncome
			flgSalary = True
		End If
		If sDailyIncome <> "" Then
			If sSalary <> "" Then sSalary = sSalary & "<br>"
			sSalary = sSalary & sDailyIncome
			flgSalary = True
		End If
		If sHourlyIncome <> "" Then
			If sSalary <> "" Then sSalary = sSalary & "<br>"
			sSalary = sSalary & sHourlyIncome
			flgSalary = True
		End If

		'歩合制
		If sPercentagePay <> "" Then
			If sPercentagePay = "1" Then
				sPercentagePay = "あり"
			ElseIf sPercentagePay = "0" Then
				sPercentagePay = "なし"
			End If
			flgSalary = True
		End If

		'交通費
		If ChkStr(rRS.Collect("NaviTrafficPayFlag")) = "1" Then 
			sTrafficFee = "交通費支給あり" & sTrafficFeeType
			If IsNumber(sTrafficFeeMonth, 0, False) = True Then
				sTrafficFee = sTrafficFee & "(" & FormatCanma(sTrafficFeeMonth) & "円／月)"
			End If
			flgSalary = True
		End If

		If flgSalary = True Then DspCondition = True
		'------------------------------------------------------------------------------
		'給与 end
		'******************************************************************************

		'******************************************************************************
		'時間 start
		'------------------------------------------------------------------------------
		sWorkRange = ""
		sWorkStartDay = ChkStr(rRS.Collect("WorkStartDay"))
		sWorkEndDay = ChkStr(rRS.Collect("WorkEndDay"))
		sWorkingTime = GetWorkingTime(rDB, rRS)
		sWorkTimeRemark = ChkStr(rRS.Collect("WorkTimeRemark"))
		flgTime = False

		'就業期間
		If sWorkStartDay & sWorkEndDay <> "" Then
			If sWorkStartDay <> "" Then sWorkRange = sWorkRange & GetDateStr(sWorkStartDay, "/")
			If sWorkRange <> "" Then sWorkRange = sWorkRange & "～"
			If sWorkEndDay <> "" Then sWorkRange = sWorkRange & GetDateStr(sWorkEndDay, "/")
		End If
		If sOrderType = "1" Then
			If rRS.Collect("WorkUpdateFlag") = "1" Then
				sWorkUpdate = "有"
			Else
				sWorkUpdate = "無"
			End If
			sWorkRange = sWorkRange & "(更新" & sWorkUpdate & ")"
		End If

		If sWorkRange & sWorkingTime & sWorkTimeRemark <> "" Then
			flgTime = True
			DspCondition = True
		End If
		'------------------------------------------------------------------------------
		'時間 end
		'******************************************************************************

		'******************************************************************************
		'休日 start
		'------------------------------------------------------------------------------
		sWeeklyHolidayType = ChkStr(rRS.Collect("WeeklyHolidayTypeName"))
		sHolidayRemark = ChkStr(rRS.Collect("HolidayRemark"))
		flgHoliday = False

		If sWeeklyHolidayType & sHolidayRemark <> "" Then
			flgHoliday = True
			DspCondition = True
		End If
		'------------------------------------------------------------------------------
		'休日 end
		'******************************************************************************

		'******************************************************************************
		'勤務先 start
		'------------------------------------------------------------------------------
		sWorkingPlace = ""
		sWPSection = ""
		sWPTel = ""
		sWPAddress = ""
		sMAP = ""
		sNearbyStation = GetNearbyStation(rDB, rRS)
		sNearbyRailway = GetNearbyRailway(rDB, rRS)
		flgWorkingPlace = False

		If sOrderType = "0" Then
			sWPSection = ChkStr(rRS.Collect("WorkingPlaceSection"))
			sWPTel = ChkStr(rRS.Collect("WorkingPlaceTelephoneNumber"))
			sWPAddress = ChkStr(rRS.Collect("WorkingPlaceAddressAll"))
		Else
			sWPAddress = ChkStr(rRS.Collect("WorkingPlacePrefectureName")) & ChkStr(rRS.Collect("WorkingPlaceCity"))
		End If
		If ChkStr(rRS.Collect("ExistsMap")) = "1" Then sMAP = "<div style=""margin:5px 0px;""><input type=""button"" value=""地図確認"" onclick=""open('/map/showmap.asp?mapOrderCode=" & sOrderCode & "', 'map', 'width=700,height=650');""></div>"

		'転勤
		If (sOrderType = "0" Or sOrderType = "2") And sCompanyKbn <> "4" Then
			'ﾘｽの派遣求人票 または 派遣会社の求人票の場合は表示しない

			sTransfer = ChkStr(rRS.Collect("Transfer"))
			If sTransfer <> "" Then
				If sTransfer = "1" Then
					sTransfer = "あり"
				Else
					sTransfer = "なし"
				End If

				sWorkingPlace = sWorkingPlace & "<tr><td><img src=""/img/order/transfer.gif"" alt=""転勤""></td><td style=""padding-left:5px;""><p class=""m0"">" & sTransfer & "</p></td></tr>"
			End If
		End If

		flgWorkingPlace = True
		DspCondition = True
		'------------------------------------------------------------------------------
		'勤務先 end
		'******************************************************************************
	End If

	flgLine = False
%>
<h3>勤務条件</h3>
<%
	If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>勤務形態</h4></div>
<div class="value1"><p class="m0"><%= GetWorkingType(rDB, rRS) %></p></div>
<div style="clear:both;"></div>
<%
	If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>職種</h4></div>
<div class="value1">
	<p class="m0"><strong><%= sJobTypeDetail %></strong></p>
	<p class="m0"><%= GetJobType(rDB, rRS) %></p>
</div>
<div style="clear:both;"></div>
<%
	If flgSalary = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>給与</h4></div>
<div class="value1">
<%
		If sYearlyIncome <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>年収</h5>
	<div class="value2"><p class="m0"><%= sYearlyIncome %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sMonthlyIncome <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>月収</h5>
	<div class="value2"><p class="m0"><%= sMonthlyIncome %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sDailyIncome <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>日給</h5>
	<div class="value2"><p class="m0"><%= sDailyIncome %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sHourlyIncome <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>時給</h5>
	<div class="value2"><p class="m0"><%= sHourlyIncome %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSalaryRemark <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>給与備考</h5>
	<div class="value2"><p class="m0"><%= sSalaryRemark %></p></div>
	<div style="clear:both; margin:0px;"></div>
<%
		End If

		If sTrafficFee <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>交通費</h5>
	<div class="value2"><p class="m0"><%= sTrafficFee %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
	<p class="m0" style="font-size:10px;">
		※最低額は条件に関係なく得られる額です。(年収の最低額は条件に関係なく得られる月給の合計です。)
	</p>
</div>
<div style="clear:both;"></div>
<%
	End If

	If flgTime = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>時間</h4></div>
<div class="value1">
<%
		If sWorkRange <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>就業期間</h5>
	<div class="value2"><p class="m0"><%= sWorkRange %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWorkingTime <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>就業時間</h5>
	<div class="value2"><p class="m0"><%= sWorkingTime %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWorkTimeRemark <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>就業時間備考</h5>
	<div class="value2"><p class="m0"><%= sWorkTimeRemark %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If

	If flgHoliday = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>休日</h4></div>
<div class="value1">
<%
		If sWeeklyHolidayType <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>休日</h5>
	<div class="value2"><p class="m0"><%= sWeeklyHolidayType %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sHolidayRemark <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>休日備考</h5>
	<div class="value2"><p class="m0"><%= sHolidayRemark %></p></div>
	<div style="clear:both;"></div>
<%
			sClearSolid = ""
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If

	If flgWorkingPlace = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>勤務先</h4></div>
<div class="value1">
<%
		If sWPSection <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>部署名</h5>
	<div class="value2"><p class="m0"><%= sWPSection %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWPTel <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>電話番号</h5>
	<div class="value2"><p class="m0"><%= sWPTel %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWPAddress <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>勤務地</h5>
	<div class="value2"><p class="m0"><%= sWPAddress %></p><%= sMAP %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sNearbyStation <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>最寄駅</h5>
	<div class="value2"><%= sNearbyStation %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sNearbyRailway <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>沿線</h5>
	<div class="value2"><%= sNearbyRailway %></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If

	If DspCondition = True Then Response.Write "<br>"
End Function














'******************************************************************************
'概　要：求人票の必要条件を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspNeedCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode			'情報コード
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'求人票種類
	Dim sCompanyKbn			'企業区分
	Dim sAge				'年齢制限
	Dim sAgeMin				'年齢下限
	Dim sAgeMax				'年齢上限
	Dim sAgeReasonFlag		'年齢理由フラグ
	Dim sAgeReason			'年齢理由
	Dim sAgeReasonDetail	'年齢制限理由詳細
	Dim sFEHistory			'学歴
	Dim sSkillOS			'ＯＳ
	Dim sSkillApp			'アプリケーション
	Dim sSkillDL			'開発言語
	Dim sSkillDB			'ＤＢ
	Dim sSkillOther			'その他スキル
	Dim sLicense			'資格
	Dim sLicenseOther		'その他資格
	Dim sOtherNote			'その他特記事項
	Dim sClearSolid			'border消去用
	Dim flgLicense			'ライセンスの項目の有無 [True]有 [False]無
	Dim flgSkill			'スキルの項目の有無 [True]有 [False]無
	Dim flgLine				'線引きフラグ
	Dim flgLine2			'線引きフラグ２

	DspNeedCondition = False

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sOrderCode = rRS.Collect("OrderCode")
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		sCompanyKbn = rRS.Collect("CompanyKbn")
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************

		'******************************************************************************
		'年齢 start
		'------------------------------------------------------------------------------
		sAge = ""
		sAgeMin = ChkStr(rRS.Collect("AgeMin"))
		sAgeMax = ChkStr(rRS.Collect("AgeMax"))
		sAgeReasonFlag = ChkStr(rRS.Collect("AgeReasonFlag"))
		sAgeReason = ChkStr(rRS.Collect("AgeReason"))
		sAgeReasonDetail = Replace(ChkStr(rRS.Collect("AgeReasonDetail")), vbCrLf, "<br>")

		If sAgeReasonFlag = "0" Or sAgeReasonFlag = "" Or (sAgeMin & sAgeMax = "") Then
			sAge = "年齢不問<br>"
			sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[？]制限について</a>"
		ElseIf sOrderType = "1" Or (sOrderType = "0" And sCompanyKbn = "4") Then
			sAge = "派遣案件のため、年齢掲載していません。<br>"
			sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[？]制限について</a>"
		Else
			If sAgeMin <> "" Then sAgeMin = sAgeMin & "歳"
			If sAgeMax <> "" Then sAgeMax = sAgeMax & "歳"
			sAge = sAgeMin & "～" & sAgeMax
			If sAgeReason <> "" Then sAge = sAge & "&nbsp;(" & sAgeReason & ")<br>"
			If sAgeReasonDetail <> "" Then sAge = sAge & sAgeReasonDetail & "<br>"
			sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[？]制限について</a><br>"
		End If

		If sAge <> "" Then DspNeedCondition = True
		'------------------------------------------------------------------------------
		'年齢 end
		'******************************************************************************

		'******************************************************************************
		'学歴 start
		'------------------------------------------------------------------------------
		sFEHistory = ChkStr(rRS.Collect("HopeSchoolHistory"))
		If sFEHistory <> "" Then sFEHistory = sFEHistory & "卒以上"
		If sFEHistory <> "" Then DspNeedCondition = True
		'------------------------------------------------------------------------------
		'学歴 end
		'******************************************************************************

		'******************************************************************************
		'資格 start
		'------------------------------------------------------------------------------
		sLicense = GetLicense(rDB, rRS)
		sLicenseOther = GetOrderNote(rDB, rRS, "OtherLicense")
		flgLicense = False
		If sLicense & sLicenseOther <> "" Then
			flgLicense = True
			DspNeedCondition = True
		End If
		'------------------------------------------------------------------------------
		'資格 end
		'******************************************************************************

		'******************************************************************************
		'スキル start
		'------------------------------------------------------------------------------
		sSkillOS = GetSkill(rDB, rRS, "OS")
		sSkillApp = GetSkill(rDB, rRS, "Application")
		sSkillDL = GetSkill(rDB, rRS, "DevelopmentLanguage")
		sSkillDB = GetSkill(rDB, rRS, "Database")
		sSkillOther = GetSkill(rDB, rRS, "OtherSkill")
		flgSkill = False
		If sSkillOS & sSkillApp & sSkillDL & sSkillDB & sSkillOther <> "" Then
			flgSkill = True
			DspNeedCondition = True
		End If
		'------------------------------------------------------------------------------
		'スキル end
		'******************************************************************************

		'******************************************************************************
		'その他特記事項 start
		'------------------------------------------------------------------------------
		sOtherNote = ""
		If sOrderType = "0" Then
			sOtherNote = GetOrderNote(rDB, rRS, "OtherNote")
			DspNeedCondition = True
		End If
		'------------------------------------------------------------------------------
		'その他特記事項 end
		'******************************************************************************
	End If

	flgLine = False
%>
<h3>必要条件</h3>
<%
	If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>年齢</h4></div>
<div class="value1"><p class="m0"><%= sAge %></p></div>
<div style="clear:both;"></div>
<%
	If sFEHistory <> "" Then
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>希望学歴</h4></div>
<div class="value1"><p class="m0"><%= sFEHistory %></p></div>
<div style="clear:both;"></div>
<%
	End If

	'******************************************************************************
	'資格出力 start
	'------------------------------------------------------------------------------
	sClearSolid = " style=""border-top-width:0px;"""
	If flgLicense = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>資格</h4></div>
<div class="value1">
<%
		If sLicense <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>資格</h5>
	<div class="value2"><%= sLicense %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sLicenseOther <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>その他資格</h5>
	<div class="value2"><p class="m0"><%= sLicenseOther %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If
	'------------------------------------------------------------------------------
	'資格出力 end
	'******************************************************************************

	'******************************************************************************
	'スキル出力 start
	'------------------------------------------------------------------------------
	sClearSolid = " style=""border-top-width:0px;"""
	If flgSkill = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>スキル</h4></div>
<div class="value1">
<%
		If sSkillOS <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>ＯＳ</h5>
	<div class="value2"><%= sSkillOS %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSkillApp <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>ｱﾌﾟﾘｹｰｼｮﾝ</h5>
	<div class="value2"><%= sSkillApp %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSkillDL <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>開発言語</h5>
	<div class="value2"><%= sSkillDL %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSkillDB <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>データベース</h5>
	<div class="value2"><%= sSkillDB %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSkillOther <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>その他スキル</h5>
	<div class="value2"><%= sSkillOther %></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If
	'------------------------------------------------------------------------------
	'スキル出力 end
	'******************************************************************************

	'******************************************************************************
	'その他特記事項 start
	'------------------------------------------------------------------------------
	If sOtherNote <> "" Then
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>特記事項</h4></div>
<div class="value1"><p class="m0"><%= sOtherNote %></p></div>
<div style="clear:both;"></div>
<%
		sClearSolid = ""
	End If
	'------------------------------------------------------------------------------
	'その他特記事項 end
	'******************************************************************************

	If DspNeedCondition = True Then Response.Write "<br>"
End Function

'******************************************************************************
'概　要：求人票の応募情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'******************************************************************************
Function DspHowToEntry(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sOrderCode			'情報コード
	Dim sCompanyCode		'企業コード
	Dim sEntryInfo			'応募方法
	Dim sProcess1			'STEP1
	Dim sProcess2			'STEP2
	Dim sProcess3			'STEP3
	Dim sProcess4			'STEP4
	Dim sCSectionName		'リス担当部署
	Dim sCPersonName		'リス担当者名
	Dim sCTel				'リス連絡先
	Dim sLis				'リス担当者
	Dim flgEntryInfo		'応募情報が有るか無いか [True]ある [False]ない
	Dim flgProcess			'選考手順が有るか無いか [True]ある [False]ない
	Dim sClearSolid
	Dim flgLine				'線引きフラグ

	DspHowToEntry = False

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sOrderType = ChkStr(rRS.Collect("OrderType"))
		sOrderCode = ChkStr(rRS.Collect("OrderCode"))
		sCompanyCode = rRS.Collect("CompanyCode")
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************

		'******************************************************************************
		'応募方法 start
		'------------------------------------------------------------------------------
		flgEntryInfo = False

		sEntryInfo = Replace(ChkStr(rRS.Collect("EntryInfo")), vbCrLf, "<br>")
		sEntryInfo = Replace(sEntryInfo, vbCr, "<br>")
		sEntryInfo = Replace(sEntryInfo, vbLf, "<br>")

		If sEntryInfo <> "" Then
			flgEntryInfo = True
			DspHowToEntry = True
		End If
		'------------------------------------------------------------------------------
		'応募方法 end
		'******************************************************************************

		'******************************************************************************
		'選考手順 start
		'------------------------------------------------------------------------------
		flgProcess = False

		sProcess1 = ChkStr(rRS.Collect("Process1"))
		sProcess2 = ChkStr(rRS.Collect("Process2"))
		sProcess3 = ChkStr(rRS.Collect("Process3"))
		sProcess4 = ChkStr(rRS.Collect("Process4"))

		If sProcess1 & sProcess2 & sProcess3 & sProcess4 <> "" Then
			flgProcess = True
			DspHowToEntry = True
		End If
		'------------------------------------------------------------------------------
		'選考手順 end
		'******************************************************************************

		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
		sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
		sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
		sLis = sCPersonName & "［リス株式会社" & sCSectionName & "］　" & sCTel & "<br>(この案件はリス株式会社が取りまとめています。)"
		DspHowToEntry = True
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************

	End If


	flgLine = False
%>
<h3>応募情報</h3>
<%
	If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>情報コード</h4></div>
<div class="value1"><p class="m0"><%= sOrderCode %></p></div>
<div style="clear:both;"></div>
<%
	If flgEntryInfo = True Then
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>応募方法</h4></div>
<div class="value1"><p class="m0"><%= sEntryInfo %></p></div>
<div style="clear:both;"></div>
<%
	End If

	If flgProcess = True Then
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>選考手順</h4></div>
<div class="value1">
<%
		If sProcess1 <> "" Then
%>
	<h5>ステップ１</h5>
	<div class="value2"><p class="m0"><%= sProcess1 %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sProcess2 <> "" Then
%>
	<h5>ステップ２</h5>
	<div class="value2"><p class="m0"><%= sProcess2 %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sProcess3 <> "" Then
%>
	<h5>ステップ３</h5>
	<div class="value2"><p class="m0"><%= sProcess3 %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sProcess4 <> "" Then
%>
	<h5>ステップ４</h5>
	<div class="value2"><p class="m0"><%= sProcess4 %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If

	If DspHowToEntry = True Then Response.Write "<br>"
End Function

'******************************************************************************
'概　要：求人票の担当者連絡先を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'******************************************************************************
Function DspContact(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sCompanyCode		'企業コード
	Dim sCompanyName		'企業名称
	Dim sCompanyNameF		'企業名称カナ
	Dim sCompanyKbn			'企業区分
	Dim sCompanySpeciality	'企業特徴
	Dim sCSectionName		'仕事の連絡先担当部署
	Dim sCPersonPost		'仕事の連絡先担当者役職
	Dim sCPersonName		'仕事の連絡先担当者名
	Dim sCPersonNameF		'仕事の連絡先担当者カナ
	Dim sCTel				'仕事の連絡先電話番号
	Dim sCMail				'仕事の連絡先メールアドレス
	Dim sPerson
	Dim sContact
	Dim sOrderType
	Dim flgLine				'線引きフラグ

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		If sOrderType <> "0" Then Exit Function
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************

		'******************************************************************************
		'会社名 start
		'------------------------------------------------------------------------------
		sCompanyName = rRS.Collect("CompanyName")
		sCompanyNameF = rRS.Collect("CompanyName_F")
		sCompanyKbn = rRS.Collect("CompanyKbn")
		sCompanySpeciality = rRS.Collect("CompanySpeciality")

		Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
		'------------------------------------------------------------------------------
		'会社名 end
		'******************************************************************************

		'******************************************************************************
		'仕事の連絡先 start
		'------------------------------------------------------------------------------
		If sOrderType = "0" Then
			sCSectionName = ChkStr(rRS.Collect("ContactSectionName"))
			sCPersonPost = ChkStr(rRS.Collect("ContactPersonPost"))
			sCPersonName = ChkStr(rRS.Collect("ContactPersonName"))
			sCPersonNameF = ChkStr(rRS.Collect("ContactPersonName_F"))
			sCTel = ChkStr(rRS.Collect("ContactTelNumber"))
			sCMail = ChkStr(rRS.Collect("ContactMailAddress"))

			If sCompanyKbn = "2" Or sCompanyKbn = "4" Then
				'人材会社の求人票の場合は「名前」＋「人材会社名」
				sPerson = sCPersonName & "(" & sCompanyName & ")"
			Else
				'一般企業の求人票の場合は「名前」＋「カナ」
				sPerson = sCPersonName
				If sCPersonNameF <> "" Then sPerson = sPerson & "(" & sCPersonNameF & ")"
			End If
'		Else
'			'リス受注票の場合は「リス担当者名」＋「リス担当者カナ」
'			sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
'			sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
'			sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
'			sPerson = sCPersonName
'			If sPerson <> "" Then sPerson = sPerson & "(人材会社：リス株式会社)"
		End If

		sContact = ""
		If sCTel <> "" Then sContact = sContact & sCTel & "	<SPAN style='font-size:10px;'>　※電話等でのお問い合わせの際、「しごとナビを見た」と言うとスムーズです。</SPAN>"
		If sContact <> "" Then sContact = sContact & "<br>"
		If sCMail <> "" Then sContact = sContact & sCMail
		'------------------------------------------------------------------------------
		'仕事の連絡先
		'******************************************************************************
	End If

	flgLine = False
%>
<h3 class="sp">担当者連絡先</h3>

<%
	If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>担当者</h4></div>
<div class="value1"><p class="m0"><%= sPerson %></p></div>
<div style="clear:both;"></div>
<%
	If sCSectionName <> "" Then
		If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>担当部署</h4></div>
<div class="value1"><p class="m0"><%= sCSectionName %></p></div>
<div style="clear:both;"></div>
<%
	End If

	If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>連絡先</h4></div>

<div class="value1"><p class="m0"><%= sContact %></p></div>
<div style="clear:both;"></div>
<br>
<%
End Function

'******************************************************************************
'概　要：リスの案件担当者、コンサル所見を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'******************************************************************************
Function DspConsultantComment(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sCompanyCode			'企業コード
	Dim sOrderType				'受注種類
	Dim sEmployeeCode			'コンサルタント社員番号
	Dim sEmployeeName			'コンサルタント名
	Dim sBranchName				'コンサルタントの拠点名
	Dim sTel					'コンサルタントの拠点の電話番号
	Dim sConsultantLink			'コンサル紹介ページへのリンク
	Dim sImg					'コンサルタントの写真
	Dim sComment				'コンサルタントコメント
	Dim sConsultantPublicFlag	'コンサルタントの紹介ページ掲載フラグ
	Dim sPictureFlag			'コンサルタント写真フラグ
	Dim sTitle					'タイトル　所見があれば"この求人情報を担当しているコンサルタントの所見"　なければ"担当者連絡先"
	Dim sClearSolid
	Dim flgLine

	If GetRSState(rRS) = False Then Exit Function

	flgLine = False

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'コンサルタント start
	'------------------------------------------------------------------------------
	'リス受注票の場合は「リス担当者名」＋「リス担当者カナ」
	sEmployeeCode = ChkStr(rRS.Collect("EmployeeCode"))
	sEmployeeName = ChkStr(rRS.Collect("EmployeeName"))
	sBranchName = ChkStr(rRS.Collect("LisDepartment"))
	sTel = ChkStr(rRS.Collect("LisTelephoneNumber"))

	sImg = "<img src=""/consultant/consultantimage.asp?ec=" & sEmployeeCode & """ alt=""この求人情報を担当しているコンサルタント"" border=""1"" width=""180"" height=""180"" style=""border-color:#666666;"">"
	sComment = Replace(ChkStr(rRS.Collect("ConsultantComment")), vbCrLf, "<br>")
	sComment = Replace(sComment, vbCr, "<br>")
	sComment = Replace(sComment, vbLf, "<br>")
	sConsultantPublicFlag = ChkStr(rRS.Collect("ConsultantPublicFlag"))
	sPictureFlag = ChkStr(rRS.Collect("ConsultantPictureFlag"))

	sConsultantLink = sEmployeeName
	If sConsultantPublicFlag = "1" Then
		sConsultantLink = "<a href=""" & HTTP_NAVI_CURRENTURL & "consultant/consultantdetail.asp?ec=" & sEmployeeCode & """>" & sEmployeeName & "</a>"
	End If
	sConsultantLink = sConsultantLink & "(人材会社：リス株式会社)"
	'------------------------------------------------------------------------------
	'コンサルタント end
	'******************************************************************************

	sTitle = "担当者連絡先"
	If sComment <> "" Then sTitle = "この求人情報を担当しているコンサルタントの所見"
%>
<h3 class="sp"><%= sTitle %></h3>
<div class="category1"><h4>コンサルタント</h4></div>
<div class="value1"><p class="m0"><%= sConsultantLink %></p></div>
<div style="clear:both;"></div>
<%
	Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
%>
<div class="category1"><h4>担当部署</h4></div>
<div class="value1"><p class="m0"><%= sBranchName %></p></div>
<div style="clear:both;"></div>
<%
	Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
%>
<div class="category1"><h4>連絡先</h4></div>
<div class="value1"><p class="m0"><%= sTel %><SPAN style='font-size:10px;'>　※お問い合わせの際、上記「情報コード」と「しごとナビを見た」と言うとスムーズです。</SPAN></p>	</div>
<div style="clear:both;"></div>
<%
	If sComment <> "" Then
		Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
%>
<div class="category1"><h4>所見</h4></div>
<div class="value1"><p class="m0"><%= sComment %></p></div>
<div style="clear:both;"></div>
<br>
<%
	End If
End Function

'******************************************************************************
'概　要：最新メールを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/company_order.asp
'******************************************************************************
Function DspNewMail(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sDateTime
	Dim sSubject
	Dim sDetail
	Dim flgLine

	DspNewMail = False

	If GetRSState(rRS) = False Then Exit Function

	flgLine = False

	If vUserType = "staff" THen
		sSQL = "sp_GetDataMailHistory '" & vUserID & "', '" & rRS.Collect("CompanyCode") & "', '" & rRS.Collect("OrderCode") & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sDateTime = GetDateStr(oRS.Collect("SendDay"), "/") & "　" & GetTimeStr(oRS.Collect("SendDay"), ":")
			sSubject = ChkStr(oRS.Collect("Subject"))
			sDetail = Replace(ChkStr(oRS.Collect("Body")), vbCrLf, "<br>")
			sDetail = Replace(sDetail, vbCr, "<br>")
			sDetail = Replace(sDetail, vbLf, "<br>")
%>
<h3 class="sp">最新の送信済みメール</h3>
<%
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>送信日時</h4></div>
<div class="value1"><p class="m0"><%= sDateTime %></p></div>
<div style="clear:both;"></div>
<%
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>サブジェクト</h4></div>
<div class="value1"><p class="m0"><%= sSubject %></p></div>
<div style="clear:both;"></div>
<%
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>内容</h4></div>
<div class="value1"><p class="m0"><%= sDetail %></p></div>
<div style="clear:both;"></div>
<br>
<%
		End If
	End If

	Call RSClose(oRS)

	DspNewMail = True
End Function

'******************************************************************************
'概　要：求人票詳細ページの勤務形態部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：staff/company_detail.asp
'******************************************************************************
Function GetWorkingType(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sWorkingType

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = rRS.Collect("OrderCode")
	sWorkingType = ""
	sSQL = "sp_GetDataWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	Do While GetRSState(oRS) = True
		sWorkingType = sWorkingType & oRS.Fields("WorkingTypeName").Value

		'リス紹介or紹介会社'従来版If (rRS.Fields("OrderType") ="" and rRS.Fields("Companykbn") = "2") or (rRS.Fields("OrderType") ="2") Then
		If (rRS.Collect("OrderType") ="0" And rRS.Collect("Companykbn") = "2") Or (rRS.Collect("OrderType") ="2") Then
			sWorkingType = sWorkingType & "【<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=300,height=200"")'>人材紹介</a>】" 
		End If

		oRS.MoveNext
		If GetRSState(oRS) = True Then sWorkingType = sWorkingType & "<br>"
	Loop
	Call RSClose(oRS)

	GetWorkingType = sWorkingType
End Function

'******************************************************************************
'概　要：求人票詳細ページの職種部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：staff/company_detail.asp
'******************************************************************************
Function GetJobType(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sJobType

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = rRS.Collect("OrderCode")
	sJobType = ""

	sSQL = "sp_GetDataJobType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	Do While GetRSState(oRS) = True
		sJobType = sJobType & oRS.Collect("JobTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sJobType = sJobType & "<br>"
	Loop
	Call RSClose(oRS)

	GetJobType = sJobType
End Function

'******************************************************************************
'概　要：求人票詳細ページの勤務形態部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：staff/company_detail.asp
'******************************************************************************
Function GetWorkingTime(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWST
	Dim sWET

	Dim sWorkingTime

	If GetRSState(rRS) = False Then Exit Function

	sWorkingTime = ""
	sSQL = "sp_GetDataWorkingTime '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sWST = ChkStr(oRS.Collect("DspWorkStartTime"))
		sWET = ChkStr(oRS.Collect("DspWorkEndTime"))
		If sWST & sWET <> "" Then
			sWorkingTime = sWorkingTime & sWST & "～" & sWET
		End If
		oRS.MoveNext
		If GetRSState(oRS) = True And sWST & sWET <> "" Then sWorkingTime = sWorkingTime & "<br>"
	Loop
	Call RSClose(oRS)

	GetWorkingTime = sWorkingTime
End Function

'******************************************************************************
'概　要：求人票詳細ページの最寄駅部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：
'******************************************************************************
Function GetNearbyStation(ByRef rDB, ByRef rRS)
	Const STATIONCOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim sStation
	Dim sToStation
	Dim iStation

	If GetRSState(rRS) = False Then Exit Function

	iStation = 0
	sStation = ""
	sSQL = "sp_GetDataNearbyStation '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iStation = iStation + 1

		sToStation = ""
		If ChkStr(oRS.Collect("ToStationTime")) <> "" Then sToStation = oRS.Collect("ToStationTime") & "分"
		If ChkStr(oRS.Collect("ToStationRemark")) <> "" Then sToStation = oRS.Collect("ToStationRemark") & sToStation
		If sToStation <> "" Then sToStation = "(" & sToStation & ")"

		sStation = sStation & "<p style=""width:50%; float:left;"">" & oRS.Collect("StationName") & "駅" & sToStation & "</p>"
		If iStation Mod STATIONCOL = 0 Then sStation = sStation & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'中途半端で終わった場合の調整
	If sStation <> "" And iStation Mod STATIONCOL <> 0 Then sStation = sStation & "<br clear=""all"">"

	GetNearbyStation = sStation
End Function

'******************************************************************************
'概　要：求人票詳細ページの最寄沿線部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：
'******************************************************************************
Function GetNearbyRailway(ByRef rDB, ByRef rRS)
	Const RAILWAYCOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim sRailway
	Dim iRailway

	If GetRSState(rRS) = False Then Exit Function

	iRailway = 0
	sRailway = ""
	sSQL = "sp_GetDataNearbyRailwayLine '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iRailway = iRailway + 1

		sRailway = sRailway & "<p style=""width:50%; float:left;"">" & oRS.Collect("RailwayLineName2") & "</p>"
		If iRailway Mod RAILWAYCOL = 0 Then sRailway = sRailway & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'中途半端で終わった場合の調整
	If sRailway <> "" And iRailway Mod RAILWAYCOL <> 0 Then
		sRailway = sRailway & "<br clear=""all"">"
	End If

	GetNearbyRailway = sRailway
End Function

'******************************************************************************
'概　要：求人票詳細ページのスキル部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：
'******************************************************************************
Function GetSkill(ByRef rDB, ByRef rRS, ByVal vCategoryCode)
	Const SKILLCOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim sSkill
	Dim iSkill

	If GetRSState(rRS) = False Then Exit Function

	iSkill = 0
	sSkill = ""
	sSQL = "sp_GetDataSkill '" & rRS.Collect("OrderCode") & "', '" & vCategoryCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iSkill = iSkill + 1

		sSkill = sSkill & "<p style=""width:50%; float:left;"">" & oRS.Collect("SkillName")
		If ChkStr(oRS.Collect("Period")) <> "" Then
			sSkill = sSkill & "<br>　<span style=""color:#339933;"">■</span>" & oRS.Collect("Period") & "年以上は尚可"
		End If
		sSkill = sSkill & "</p>"
		If iSkill Mod SKILLCOL = 0 Then sSkill = sSkill & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'中途半端で終わった場合の調整
	If sSkill <> "" And iSkill Mod SKILLCOL <> 0 Then sSkill = sSkill & "<br clear=""all"">"

	GetSkill = sSkill
End Function

'******************************************************************************
'概　要：求人票詳細ページの資格部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'******************************************************************************
Function GetLicense(ByRef rDB, ByRef rRS)
	Const LICENSECOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim iLicense
	Dim sLicense

	If GetRSState(rRS) = False Then Exit Function

	iLicense = 0
	sLicense = ""

	sSQL = "sp_GetDataLicense '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iLicense = iLicense + 1

		sLicense = sLicense & "<p style=""width:50%; float:left;"">" & oRS.Collect("LicenseName") & "</p>"
		If iLicense Mod LICENSECOL = 0 Then sLicense = sLicense & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'中途半端で終わった場合の調整
	If sLicense <> "" And iLicense Mod LICENSECOL <> 0 Then sLicense = sLicense & "<br clear=""all"">"

	GetLicense = sLicense
End Function

'******************************************************************************
'概　要：求人票詳細ページのその他情報取得
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailOrderで生成されたレコードセットオブジェクト
'　　　：vCode			：C_Noteテーブルの Code フィールド値
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'******************************************************************************
Function GetOrderNote(ByRef rDB, ByRef rRS, ByVal vCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sNote

	If GetRSState(rRS) = False Then Exit Function

	sSQL = "sp_GetDataNote '" & rRS.Collect("OrderCode") & "', '"  & vCode &"'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sNote = oRS.Collect("Note")
	End If
	Call RSClose(oRS)

	GetOrderNote = sNote
End Function

'******************************************************************************
'概　要：求人票詳細のタイトルとディスクリプションを取得
'作成者：Lis Kokubo
'作成日：2007/02/12
'戻り値：rTitle			：タイトル（具体的職種名）
'　　　：rDescription	：説明文（担当業務）
'使用元：しごとナビ/order/order_detail.asp
'備　考：
'******************************************************************************
Function GetOrderTitle(ByRef rDB, ByVal vOrderCode, ByRef rTitle, ByRef rDescription)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWorkingType

	sSQL = "up_GetOrderTitle '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		rTitle = ChkStr(oRS.Collect("JobTypeDetail"))
		rDescription = ChkStr(oRS.Collect("BusinessDetail"))
	End If
	Call RSClose(oRS)

	sSQL = "sp_GetListWorkingType '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sWorkingType = ""
	Do While GetRSState(oRS) = True
		If sWorkingType <> "" Then sWorkingType = sWorkingType & ","
		sWorkingType = sWorkingType & oRS.Collect("WorkingTypeName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	If rTitle <> "" Then rTitle = rTitle & "&nbsp;"
	rTitle = rTitle & sWorkingType

	GetOrderTitle = flgQE
End Function

'******************************************************************************
'概　要：スキルの各項目表示
'作成者：Lis Kokubo
'作成日：2007/02/14
'戻り値：
'　　　：
'使用元：しごとナビ/order/order_detail.asp
'備　考：
'******************************************************************************
Function GetSkillList(ByVal vTitleImg, ByVal vTitleAlt, ByVal vSkill)
	GetSkillList = ""
	If Len(vSkill) = 0 Then Exit Function
	GetSkillList = "<tr><td valign=""top""><img src=""" & vTitleImg & """ alt=""" & vTitleAlt & """ width=""50"" height=""12""></td><td style=""padding-left:5px;"">" & vSkill & "</td></tr>"
End Function

'******************************************************************************
'概　要：レコメンドお仕事情報一覧出力
'引　数：rDB		：DB接続オブジェクト
'　　　：vUserType	：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID	：利用中ユーザのユーザID [Session("userid")]
'　　　：vOrderCode	：閲覧中求人票の情報コード
'　　　：vRCMD		：レコメンド種類 ["1"]こんなお仕事情報も見てます ["2"]近い条件のお仕事情報
'　　　：vMyOrder	：自社求人票か否か ["1"]自社求人票
'戻り値：
'作成日：2007/05/31
'作成者：Lis Kokubo
'備　考：
'更　新：
'******************************************************************************
Function DspRecommendOrderList(ByRef rDB, ByVal vUserType, ByVal vUserID, ByVal vOrderCode, ByVal vRCMD, ByVal vMyOrder)
	Const MAXCOLS = 3

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTitle
	Dim idx			'ループカウントアップ変数
	Dim iCols		'列数
	Dim aPadding(2)	'各列のパディング
	Dim aJobTypeDetail()
	Dim aCompanyName()
	Dim aImg()
	Dim aWorkingTypeIcon()
	Dim aWorkingPlace()
	Dim aStation()
	Dim aYearlyIncome()
	Dim aMonthlyIncome()
	Dim aDailyIncome()
	Dim aHourlyIncome()

	If vMyOrder = "1" Then Exit Function

	Select Case vRCMD
		Case "1"
			sSQL = "up_SearchRelationAccessOrder '" & CONF_OrderCode & "'"
			sTitle = "この求人情報を見た人はこんな求人情報も見ています"
		Case "2"
			sSQL = "up_SearchHighRelationOrder '" & CONF_OrderCode & "'"
			sTitle = "この求人情報の条件に近い求人情報"
		Case Else
			Exit Function
	End Select

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = False Then Exit Function
%>
<h2 class="ssubtitle"><%= sTitle %></h2>
<div class="subcontent" style="margin-bottom:15px;">
<%
	Call DspOrderListDetail3(rDB, oRS, 3, 1, vRCMD)
%>
</div>
<%
End Function

'******************************************************************************
'概　要：レコメンドの求人票一覧の、求人票一つ一つの各項目（職種、企業名など）を取得
'引　数：rDB		：DB接続オブジェクト
'　　　：rRS		：求人票一覧のレコードセット
'　　　：vRCMD		：レコメンド種類 ["1"]こんなお仕事情報も見てます ["2"]近い条件のお仕事情報
'　　　：[OUTPUT]rJobTypeDetail		：具体的職種名
'　　　：[OUTPUT]rCompanyName		：企業名
'　　　：[OUTPUT]rImg				：企業イメージ
'　　　：[OUTPUT]rWorkingTypeIcon	：勤務形態アイコン
'　　　：[OUTPUT]rWorkingPlace		：勤務地
'　　　：[OUTPUT]rStation			：最寄駅
'　　　：[OUTPUT]rYearlyIncome		：年収
'　　　：[OUTPUT]rMonthlyIncome		：月収
'　　　：[OUTPUT]rDailyIncome		：日給
'　　　：[OUTPUT]rHourlyIncome		：時給
'戻り値：
'作成日：2007/05/31
'作成者：Lis Kokubo
'備　考：
'更　新：
'******************************************************************************
Function GetRecommendValues(ByRef rDB, ByRef rRS, ByVal vRCMD, ByRef rJobTypeDetail, ByRef rCompanyName, ByRef rImg, ByRef rWorkingTypeIcon, ByRef rWorkingPlace, ByRef rStation, ByRef rYearlyIncome, ByRef rMonthlyIncome, ByRef rDailyIncome, ByRef rHourlyIncome)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim sOrderCode			'情報コード
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'受注区分
	Dim sCompanyKbn			'会社区分
	Dim sCompanyName		'企業名
	Dim sCompanyNameF		'企業名カナ
	Dim sCompanySpeciality	'企業名（特徴）
	Dim sJobTypeDetail		'具体的職種名(altやtitleで出力する)
	Dim sViewJobTypeDetail	'求職者に見える具体的職種名(長い文字列はカットされる)
	Dim sBusinessDetail		'担当業務
	Dim sYearlyIncome		'年収
	Dim sYearlyIncomeMin	'年収下限
	Dim sYearlyIncomeMax	'年収上限
	Dim sMonthlyIncome		'月収
	Dim sMonthlyIncomeMin	'月収下限
	Dim sMonthlyIncomeMax	'月収上限
	Dim sDailyIncome		'日給
	Dim sDailyIncomeMin		'日給下限
	Dim sDailyIncomeMax		'日給上限
	Dim sHourlyIncome		'時給
	Dim sHourlyIncomeMin	'時給下限
	Dim sHourlyIncomeMax	'時給上限
	Dim sWorkingTypeIcon	'勤務形態アイコン並び
	Dim sWorkingPlace		'勤務地
	Dim sStation			'最寄駅
	Dim sImg				'画像URL

	Dim sURL				'求人票詳細のURL
	Dim sAlign				'枠寄せ [vCols = 1]left [vCols = vMaxCols]right [それ以外]center

	If GetRSState(rRS) = False Then Exit Function

	sURL = HTTP_CURRENTURL & "order/order_detail.asp"

	sSQL = "sp_GetDetailOrder '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'情報コード
	sOrderCode = ChkStr(oRS.Collect("OrderCode"))
	'企業コード
	sCompanyCode = ChkStr(oRS.Collect("CompanyCode"))
	'受注区分
	sOrderType = ChkStr(oRS.Collect("OrderType"))
	'企業区分
	sCompanyKbn = ChkStr(oRS.Collect("CompanyKbn"))
	'企業名, 企業名カナ
	sCompanyName = ChkStr(oRS.Collect("CompanyName"))
	sCompanyNameF = ChkStr(oRS.Collect("CompanyName_F"))
	sCompanySpeciality = ChkStr(oRS.Collect("CompanySpeciality"))
	Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'具体的職種名
	sJobTypeDetail = ChkStr(oRS.Collect("JobTypeDetail"))
	sViewJobTypeDetail = sJobTypeDetail
	If Len(sViewJobTypeDetail) > 14 Then sViewJobTypeDetail = Left(sViewJobTypeDetail, 14) & ".."
	'担当業務
	sBusinessDetail = ChkStr(oRS.Collect("BusinessDetail"))

	'******************************************************************************
	'給与 start
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
	'月収
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
	'勤務形態アイコン start
	'------------------------------------------------------------------------------
	sWorkingTypeIcon = ""
	sSQL = "sp_GetListWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		Select Case ChkStr(oRS2.Collect("WorkingTypeCode"))
			Case "001": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/haken.gif"" alt=""派遣"" style=""margin-right:1px;"">"
			Case "002": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/seishain.gif"" alt=""正社員"" style=""margin-right:1px;"">"
			Case "003": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/keiyaku.gif"" alt=""契約社員"" style=""margin-right:1px;"">"
			Case "004": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/syoha.gif"" alt=""紹介予定派遣"" style=""margin-right:1px;"">"
			Case "005": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/arbeit.gif"" alt=""アルバイト・パート"" style=""margin-right:1px;"">"
			Case "006": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/soho.gif"" alt=""SOHO"" style=""margin-right:1px;"">"
			Case "007": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/fc.gif"" alt=""FC"" style=""margin-right:1px;"">"
		End Select
		oRS2.MoveNext
	Loop
	Call RSClose(oRS2)
	'------------------------------------------------------------------------------
	'勤務形態アイコン end
	'******************************************************************************

	'******************************************************************************
	'画像 start
	'------------------------------------------------------------------------------
	sImg = ""
	sSQL = "up_GetListOrderPictureNow '" & sCompanyCode & "', '" & sOrderCode & "', 'orderpicture'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo1")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo1")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo2")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo2")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo3")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo3")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo4")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo4")
	End If

	If sImg = "" And sOrderType = "0" Then
		sSQL = "sp_GetDataPicture '" & sCompanyCode & "', '1'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"
		End If
	End If

	If sImg = "" Then sImg = "/img/nopicture180.gif"
	'sImg = "<img src=""" & sImg & """ alt=""" & sCompanyName & """ width=""156"" height=""117"">"
	sImg = "<img src=""" & sImg & """ alt=""" & sCompanyName & """ width=""88"" height=""66"" border=""0"" align=""left"" style=""margin:0px; padding:0px;"">"
	'------------------------------------------------------------------------------
	'画像 end
	'******************************************************************************

	'******************************************************************************
	'勤務地 start
	'------------------------------------------------------------------------------
	sWorkingPlace = ""
	If sOrderType = "0" Then
		sWorkingPlace = ChkStr(oRS.Collect("WorkingPlaceAddressAll"))
	Else
		sWorkingPlace = ChkStr(oRS.Collect("WorkingPlacePrefectureName")) & ChkStr(oRS.Collect("WorkingPlaceCity"))
	End If
	'------------------------------------------------------------------------------
	'最寄駅 end
	'******************************************************************************

	'******************************************************************************
	'最寄駅 start
	'------------------------------------------------------------------------------
	sStation = ""
	sSQL = "sp_GetDataNearbyStation '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		sStation = sStation & GetStrNearbyStation(oRS2.Collect("StationName"), oRS2.Collect("ToStationTime"), oRS2.Collect("ToStationRemark"))
		oRS2.MoveNext
		If GetRSState(oRS2) = True Then sStation = sStation & "<br>"
	Loop
	'------------------------------------------------------------------------------
	'最寄駅 end
	'******************************************************************************

	rJobTypeDetail = "<a href=""" & sURL & "?ordercode=" & sOrderCode & "&amp;rcmd=" & vRCMD & """>" & sViewJobTypeDetail & "</a>"
	rCompanyName = sCompanyName
	rImg = "<a href=""" & sURL & "?ordercode=" & sOrderCode & "&amp;rcmd=" & vRCMD & """>" & sImg & "</a>"
	rWorkingTypeIcon = sWorkingTypeIcon
	rWorkingPlace = sWorkingPlace
	rStation = sStation
	rYearlyIncome = sYearlyIncome
	rMonthlyIncome = sMonthlyIncome
	rDailyIncome = sDailyIncome
	rHourlyIncome = sHourlyIncome
End Function

'******************************************************************************
'概　要：自社求人票の掲載状態を変更する
'引　数：rDB			：接続中のDBConnection
'　　　：vOrderCodes	：更新対象の情報コード群（カンマ区切り）
'　　　：vPublicFlags	：更新対象の公開フラグ群（カンマ区切り）
'作成者：Lis Kokubo
'作成日：2007/04/02
'備　考：
'使用元：しごとナビ/order/order_list_entity.asp
'******************************************************************************
Function UpdMyOrderPublicFlag(ByRef rDB, ByVal vOrderCodes, ByVal vPublicFlags)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim aOrderCode
	Dim aPublicFlag
	Dim idx

	flgQE = True
	aOrderCode = Split(Replace(vOrderCodes, " ", ""), ",")
	aPublicFlag = Split(Replace(vPublicFlags, " ", ""), ",")

	sSQL = ""
	For idx = LBound(aOrderCode) To UBOund(aOrderCode)
		If aPublicFlag(idx) <> "" Then
			sSQL = sSQL & "EXEC sp_Reg_PublicFlag" & _
				" '" & CONF_CompanyCode & "'" & _
				",'" & aOrderCode(idx) & "'" & _
				",'" & aPublicFlag(idx) & "'" & vbCrLf
		End If
	Next
	If sSQL <> "" Then flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	UpdMyOrderPublicFlag = flgQE
End Function

'******************************************************************************
'概　要：自社求人票を削除する
'引　数：rDB			：接続中のDBConnection
'　　　：vOrderCodes	：更新対象の情報コード群（カンマ区切り）
'作成者：Lis Kokubo
'作成日：2007/04/02
'備　考：
'使用元：しごとナビ/order/order_list_entity.asp
'******************************************************************************
Function DelMyOrder(ByRef rDB, vOrderCodes)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim aOrderCode
	Dim idx

	aOrderCode = Split(Replace(vOrderCodes, " ", ""), ",")
	For idx = LBound(aOrderCode) To UBound(aOrderCode)
		If aOrderCode(idx) <> "" Then
			sSQL = sSQL & "EXEC sp_Reg_RegistCommit" & _
				" '" & Replace(aOrderCode(idx), " ", "") & "'" & vbCrLf & _
				",'0'"
		End If
	Next
	If sSQL <> "" Then flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
End Function

'******************************************************************************
'概　要：求人票の特徴
'引　数：rDB
'　　　：rRS
'作成者：Lis Kokubo
'作成日：2007/02/14
'戻り値：
'　　　：
'使用元：しごとナビ/order/order_detail.asp
'備　考：
'******************************************************************************
Function GetImgOrderSpeciality(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWorkingCode
	Dim sOrderType
	Dim sCompanyKbn

	If GetRSState(rRS) = False Then Exit Function

	sOrderType = rRS.Collect("OrderType")
	sCompanyKbn = rRS.Collect("CompanyKbn")

	GetImgOrderSpeciality = ""
	'アクセス数が100を超えていれば「HOT」表示（リス安藤）
	If rRS.Collect("AccessCount") > 100 Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/c_HOT_green.gif"" alt=""人気"" width=""50"" height=""15"">&nbsp;"
	End If

	'UPDATEと今日から10日引いた日で「新着」表示(リス安藤)
	If rRS.Collect("Updateday") > NOW()-10 Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/c_NEW_green.gif"" alt=""新着"" width=""50"" height=""15"">&nbsp;"
	End If

	'未経験者ＯＫの場合、わかばマーク表示(リス安藤)
	If rRS.Collect("InexperiencedPersonFlag") = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/no_experience.gif"" alt=""未経験者／第二新卒歓迎"" width=""50"" height=""15"">&nbsp;"
	End If

	'Ｕターン・Ｉターン
	If rRS.Collect("UITurnFlag") = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/ui_turn.gif"" alt=""Ｕターン・Ｉターン"" width=""50"" height=""15"">&nbsp;"
	End If

	'語学を活かす仕事
	If rRS.Collect("UtilizeLanguageFlag") = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/linguistic_job.gif"" alt=""語学を活かす仕事"" width=""50"" height=""15"">&nbsp;"
	End If

	'年間休日120日以上
	If rRS.Collect("ManyHolidayFlag") = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/year_holidaycnt.gif"" alt=""年間休日120日以上"" width=""50"" height=""15"">&nbsp;"
	End If

	'フレックスタイム制度あり ------2006/01/10 Hayashi ADD
	If rRS.Collect("FlexTimeFlag") = "1" And sOrderType = "0" And sCompanyKbn = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/flextime.gif"" alt=""フレックスタイム制度あり"" width=""50"" height=""15"">&nbsp;"
	End If

'直接Yahoo!の検索からお仕事情報詳細ページへ来る人へアイコン表示
if G_FLGRESUME = False Then
	if InStr(Request.ServerVariables("HTTP_REFERER"),"search.yahoo.co.jp/") <> 0 Then

	sSQL = "sp_GetDataWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sWorkingcode = oRS.Collect("WorkingTypecode")

		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/order_detail_icon/icon_w" & sWorkingcode & ".gif"" alt=""派遣社員"" width=""50"" height=""15"">&nbsp;"

		oRS.MoveNext
	Loop

	GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/order_detail_icon/icon_p" & rRS.Collect("Workingplaceprefecturecode") & ".gif"" alt=""北海道"" width=""50"" height=""15"">&nbsp;"
	End if
End if
'/直接Yahoo!の検索からお仕事情報詳細ページへ来る人へアイコン表示

	If GetImgOrderSpeciality <> "" Then GetImgOrderSpeciality = "<div>" & GetImgOrderSpeciality & "</div>"

End Function

'******************************************************************************
'概　要：しごとナビの求人票詳細ページの上部に置く、ログイン誘導ボタン。
'引　数：vOrderCode	：ログイン後の飛び先情報コード
'作成者：Lis Kokubo
'作成日：2007/02/20
'戻り値：×
'使用元：しごとナビ/order/order_detail.asp
'備　考：
'******************************************************************************
Sub DspTopRegButton(ByVal vOrderCode)
%>
<div align="right" style="width:600px; margin-bottom:5px;">
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button3.gif" alt="ログインして応募" border="0"></a></div>
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="履歴書登録して応募" border="0"></a></div>
	<div style="clear:both;"></div>
<!--
	<div align="center">
	<form action="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>" method="post">
	<input type="submit" value="会員の方はこちらから応募できます">
	</form>
	</div>
-->
</div>
<%
End Sub

'******************************************************************************
'概　要：＠履歴書の求人票詳細ページの上部に置く、ログイン誘導ボタン。
'引　数：vOrderCode	：ログイン後の飛び先情報コード
'作成者：Lis Kokubo
'作成日：2007/02/20
'戻り値：×
'使用元：しごとナビ/resume/order/order_detail.asp
'備　考：
'******************************************************************************
Sub DspTopRegButtonResume(ByVal vOrderCode)
%>
<div align="right" style="width:600px; margin-bottom:5px;">
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button3.gif" alt="ログインして応募" border="0"></a></div>
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="履歴書登録して応募" border="0"></a></div>
	<div style="clear:both;"></div>
</div>
<%
End Sub

'******************************************************************************
'概　要：しごとナビの求人票詳細ページの下部に置く、ログイン誘導ボタン。
'引　数：vOrderCode	：ログイン後の飛び先情報コード
'作成者：Lis Kokubo
'作成日：2007/02/20
'戻り値：×
'使用元：しごとナビ/order/order_detail.asp
'備　考：
'******************************************************************************
Sub DspBottomRegButton(ByVal vOrderCode)
%>
<div align="center">
	<hr size="1">
	<p style="color:#ff0000;">
▼▼会員登録すれば応募や質問が可能になります！▼▼<BR>
応募のための履歴書も自動作成されます。</p>
	<hr size="1">
	<div align="center" style="float:left; width:300px;color:#C51035;">＜まだIDをお持ちでない方＞<br><a href="<%= HTTPS_NAVI_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="履歴書登録して応募" border="0"></a></div>
	<div align="center" style="float:right; width:300px;color:#C51035;">＜すでにIDをお持ちの方＞<br><a href="<%= HTTPS_NAVI_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button3.gif" alt="ログインして応募" border="0"></a></div>
	<div style="clear:both;"></div>
	<br>
</div>
<%
End Sub

'******************************************************************************
'概　要：＠履歴書の求人票詳細ページの下部に置く、ログイン誘導ボタン。
'引　数：vOrderCode	：ログイン後の飛び先情報コード
'作成者：Lis Kokubo
'作成日：2007/02/20
'戻り値：×
'使用元：しごとナビ/resume/order/order_detail.asp
'備　考：
'******************************************************************************
Sub DspBottomRegButtonResume(ByVal vOrderCode)
%>
<div align="center">
	<hr size="1">
	<p style="color:#ff0000;">▼会員登録すれば応募や質問が可能になります！▼</p>
	<hr size="1">
	<div align="center" style="float:left; width:300px;color:#C51035;">＜まだIDをお持ちでない方＞<br><a href="<%= HTTPS_NAVI_CURRENTURL %>resume/staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="履歴書登録して応募" border="0"></a></div>
	<div align="center" style="float:right; width:300px;color:#C51035;">＜すでにIDをお持ちの方＞<br><a href="<%= HTTPS_NAVI_CURRENTURL %>resume/login/login.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/resume/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button3.gif" alt="ログインして応募" border="0"></a></div>
	<div style="clear:both;"></div>
	<br>
</div>
<%
End Sub

'******************************************************************************
'概　要：新着求人情報メールからアクセスがあった場合のログ書き込み
'引　数：rDB		
'　　　：rRS		
'　　　：vMU		：メルマガユーザＩＤ
'　　　：vOrderCode	：閲覧中求人票
'作成者：Lis Kokubo
'作成日：2007/05/08
'戻り値：
'　　　：
'使用元：しごとナビ/order/order_detail_entity.asp
'備　考：
'******************************************************************************
Function MailMagazineAccess(ByRef rDB, ByVal vMU, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	If IsNumber(vMU, 0, False) = True Then
		sSQL = "up_Reg_LOG_MailMagazineAccess '" & vMU & "', '" & vOrderCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'概　要：求人メルマガからアクセスがあった場合のログ書き込み
'引　数：rDB		
'　　　：rRS		
'　　　：vMU		：メルマガユーザＩＤ
'　　　：vOrderCode	：閲覧中求人票
'作成者：Lis Kokubo
'作成日：2007/05/08
'戻り値：
'　　　：
'使用元：しごとナビ/order/order_detail_entity.asp
'備　考：
'******************************************************************************
Function MailMagazineDelivery(ByRef rDB, ByVal vMI, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	If IsNumber(vMI, 0, False) = True Then
		sSQL = "up_Reg_LOG_MailMagazineDelivery '" & vMI & "', '" & vOrderCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'概　要：足跡ログの書き込み
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_SearchOrder or 求人票詳細検索SQL で生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vOrderCode		：閲覧中求人票
'作成者：Lis Kokubo
'作成日：2007/05/08
'備　考：
'使用元：order/order_detail_entity.asp
'******************************************************************************
Function AccessHistoryOrder(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	If vUserType = "staff" Then
		sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & vUserID & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Call RSClose(oRS)
	ElseIf IsRE(Request.Cookies("id_memory"), "^S\d\d\d\d\d\d\d$", True) = True Then
		sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & Request.Cookies("id_memory") & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'概　要：アクセス回数のカウントアップ
'引　数：rDB		：接続中のDBConnection
'　　　：vOrderCode	：閲覧中求人票の情報コード
'作成者：Lis Kokubo
'作成日：2007/05/08
'備　考：
'使用元：order/order_detail_entity.asp
'******************************************************************************
Function AccessCountUp(ByRef rDB, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	AccessCountUp = 0

	sSQL = "sp_Reg_AccessCountUp '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS2) = True Then
		AccessCountUp = oRS.Collect("AccessCount")
	End If
	Call RSClose(oRS)
End Function

'*******************************************************************************
'概　要：全角半角が混じった文字列のバイト数を正確に返す(Webからの引用)
'引　数：string		:対象文字列
'戻り値：Interger	:対象文字列のバイト数
'作成日：2007/05/23 Lis Sotome
'更　新：
'*******************************************************************************
Function LenByte(ByRef string)

    Dim c, i, k

    c = 0

    For i = 0 To Len(string) - 1
        k = Mid(string, i + 1, 1)
        If (Asc(k) And &HFF00) = 0 Then
            c = c + 1
        Else
            c = c + 2
        End If
    Next

    LenByte = c

End Function

'*******************************************************************************
'概　要：文字列の左端から指定されたバイト数分の文字列を抽出する(全角半角の混じった文字列対応)
'　　　：※指定されたバイト数で収まらない全角文字は削られます
'　　　：ex:sStr="aaあ", vByte=3 ・・・戻り値:"aa"
'引　数：sStr		:対象文字列
'      ：vByte		:抽出する文字列のバイト数
'戻り値：String		:抽出後の文字列
'作成日：2007/05/23 Lis Sotome
'更　新：
'*******************************************************************************
Function LeftByte(ByRef sStr, ByRef vByte)

    Dim cnt, i, k
	Dim sBuf	'文字列用バッファ

    cnt = 0

    For i = 0 To Len(sStr) - 1
        k = Mid(sStr, i + 1, 1)
        If (Asc(k) And &HFF00) = 0 Then
            cnt = cnt + 1
        Else
            cnt = cnt + 2
        End If

		If cnt > vByte Then	'目的の文字数を超えた(半角、全角と続いた)とき
			LeftByte = sBuf
			Exit Function	'処理終了
		Elseif cnt = vByte Then	'目的の文字数の(半角、半角または全角、全角と続いた)とき
			sBuf = sBuf & k
			LeftByte = sBuf
			Exit Function	'処理終了
		Elseif cnt < vByte Then
			sBuf = sBuf & k
		End If
	Next

	LeftByte = sBuf

End Function
%>
