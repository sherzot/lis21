<%
'**********************************************************************************************************************
'概　要：求人票一覧ページ /order/order_list_entity.asp
'　　　：求人票詳細ページ /order/order_detail_entity.asp
'　　　：企業情報ページ /order/company_order.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：main_pics
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　求人票一覧ページ出力用　■■■
'　　　：DspOrderListDetail			：求人票一覧ページの各求人票単位を表示
'　　　：DspOrderListDetail2		：求人票一覧横並びバージョン1
'　　　：DspOrderListDetail3		：求人票一覧横並びバージョン2
'　　　：DspOrderListDetail4		：求人票一覧横並びバージョン3 (2重応募対応入り)
'　　　：DspPageControl				：求人票一覧ページのページコントロール
'　　　：
'　　　：■■■　企業情報ページ出力用　■■■
'　　　：DspCompanyInfo				：企業情報の基本情報を出力
'       :DspCompanyInfoNEO          ：企業情報の基本情報を出力(NEO用)
'　　　：DspCompanyPR				：企業情報のＰＲ情報を出力
'　　　：
'　　　：■■■　求人票詳細ページ出力用　■■■
'　　　：DspLisOrderCompanyInfo		：求人票詳細ページのリスの紹介先・派遣先企業情報を出力
'　　　：DspTempOrderCompanyInfo	：求人票詳細ページの派遣企業の派遣先企業情報を出力
'　　　：DspOrderControlButton		：求人票詳細ページのコントロールボタン（ログイン済みユーザ用）
'　　　：JSOrderControlButton		：求人票詳細ページのコントロールボタンで利用するjavascriptの出力
'　　　：FrmOrderControlButton		：求人票詳細ページのコントロールボタンで利用するFORMデータの出力
'　　　：DspOrderCompanyName		：求人票詳細ページの企業名を出力
'　　　：DspOrderShowTypeSwitch		：求人票詳細ページの会社情報・職種情報・インタビュー切り替えボタンと参照回数を出力
'　　　：DspOrderCatchCopy			：求人票詳細ページのキャッチコピー部分（大きい画像など）を出力
'　　　：DspOrderCatchCopy_OldPlan		：求人票詳細ページの過去求人バナーを出力
'　　　：DspOrderFreePR				：求人票詳細ページのフリーＰＲを出力
'　　　：DspOrderPictureNow			：求人票詳細ページの小さい画像を出力
'　　　：DspOrderBackGround			：求人票詳細ページの採用の背景を出力
'　　　：DspBusiness				：求人票詳細ページの業務内容を出力
'　　　：DspCondition				：求人票詳細ページの勤務条件を出力
'　　　：DspNeedCondition			：求人票詳細ページの必要条件を出力
'　　　：DspHowToEntry				：求人票詳細ページの応募情報を出力
'　　　：DspContact					：求人票詳細ページの担当者連絡先を出力
'　　　：DspElderInterview			：求人票詳細ページの先輩インタビューを出力
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
'　　　：DspBottomRegButton_OldPlan		：しごとナビの求人票詳細ページの下部に置く、ログイン誘導ボタン。
'　　　：DspBottomRegButtonResume	：＠履歴書の求人票詳細ページの下部に置く、ログイン誘導ボタン。
'　　　：
'　　　：■■■　求人票詳細アクセス時の制御　■■■
'　　　：MailMagazineAccess			：新着求人メールからのアクセス時のログ書き込み
'　　　：MailMagazineDelivery		：求人メルマガからのアクセス時のログ書き込み
'　　　：AccessHistoryOrder			：足跡ログの書き込み
'　　　：AccessCountUp				：アクセス回数のカウントアップ
'　　　：PVCountUp					：求人票の日別ＰＶのカウントアップ
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
'履　歴：2006/05/13 LIS K.Kokubo 作成
'　　　：2007/11/22 LIS K.Kokubo up_SearchOrderを必要最小限のものだけを取ってくるようにしたことによる変更。up_DtlOrderからデータを取得。
'　　　：2008/03/04 LIS K.Kokubo 掲載終了日を[RiyoToDate]→[DspPublicLimitDay]に変更
'　　　：2008/03/11 LIS K.Kokubo トップインタビューへのリンクを出力
'　　　：2008/08/01 LIS K.Kokubo Ｗバリューのリンクを出力
'　　　：2008/08/19 LIS 林 特徴フラグの追加とフレックス移動
'　　　：2008/10/20 LIS K.Kokubo 勤務地複数化による修正
'　　　：2010/01/28 LIS K.Kokubo 給与の記載が無い場合は給与の項目を表示しない（FC・SOHO対応）
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
	'HTTPクラス変更用変数
	Dim HtmlClassName
	Dim HtmlWorkingType
    Dim HimlOiwai

	
	
	
	
	
	Dim flgAddWatchList
	Dim iMailTemplateCnt	'メールテンプレートの件数
	Dim sAncMT				'メールテンプレートへのリンク
	Dim sOrderCode
	
	
	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "EXEC up_ChkWatchListExists_Staff '" & vUserID & "', '" & sOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
	End If
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************
	
	
	
	
	
	
	
'	Dim qsOrderCode				'オーダーコード(受注表番号)
'	Dim iDetail				'求人票詳細からのフラグ
'	
'	qsOrderCode = GetForm("ordercode", 2)
'	iDetail = GetForm("Detail", 2)
'
'	'******************************************************************************
'	'企業コード start
'	'------------------------------------------------------------------------------
'	flgAddWatchList = False
'	
'	sSQL = "EXEC up_ChkWatchListExists_Staff '" & G_USERID & "', '" & qsOrderCode & "';"
'	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
'	If GetRSState(oRS) = True Then
'		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
'	End If
'	Call RSClose(oRS)
'	'------------------------------------------------------------------------------
'	'企業コード end
'	'******************************************************************************
	
	
	
	
	
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	DspOrderListDetail = False

	If G_USEFLAG = "0" And vMyOrder = "1" And G_OLDAPPLICATIONCODE <> "" Then
		sSQL = "EXEC up_DtlOrder_NEO '" & rRS.Collect("OrderCode") & "', '" & G_OLDAPPLICATIONCODE & "';"
	Else
		sSQL = "EXEC up_DtlOrder_NEO '" & rRS.Collect("OrderCode") & "', '';"
	End If

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	dbCompanyCode = oRS.Collect("CompanyCode")
	sOrderType = ChkStr(oRS.Collect("OrderType"))
	sPlanType = ChkStr(oRS.Collect("PlanTypeName"))
	iImageLimit = oRS.Collect("ImageLimit")
    'お祝い金設定
    HimlOiwai = oRS.Collect("CongratulationPrice")

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
	sImgOrderState = GetImgOrderSpeciality(rDB, oRS)
	'--------------------------------------------------------------------------
	'状態img end
	'**************************************************************************

	'**************************************************************************
	'キャッチコピー start
	'--------------------------------------------------------------------------
	sCatchCopy = ""
	sCatchCopy = chkstr(oRS.Collect("CatchCopy"))
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
		sImgMain = "background:url("& sImg1 &") no-repeat;"
	End If

	If sImg2 <> "" Then
		sImgSub = sImgSub & "<div class=""sub_img"">" & _
			"<img src=""" & sImg2 & """><br>"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	If sImg3 <> "" Then
		sImgSub = sImgSub & "<div class=""sub_img"" style=""margin-top: 5px;"">" & _
			"<img src=""" & sImg3 & """><br>"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	'If sImg4 <> "" Then
	'	sImgSub = sImgSub & "<div class=""sub_img"">" & _
	'		"<img src=""" & sImg4 & """><br>"
	'	sImgSub = sImgSub & "</div>"
	'End If

	If sImgSub <> "" Then sImgSub =  sImgSub 
	'--------------------------------------------------------------------------
	'画像 end
	'**************************************************************************

	'**************************************************************************
	'担当業務 start
	'--------------------------------------------------------------------------
	If flgImg = True Then
		'画像が有る場合は文章を短めにカット
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),300) & "&nbsp;"
		If Len(sBusinessDetail) > 300 Then sBusinessDetail = sBusinessDetail & "..."
	Else
		'画像が無い場合は文章を長めにカット
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),465) & "&nbsp;"
		If Len(sBusinessDetail) > 465 Then sBusinessDetail = sBusinessDetail & "..."
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
	Dim Counter
	Counter = 1
	Do While GetRSState(oRS2) = True
		if Counter = 1 Then
			HtmlWorkingType = oRS2.Collect("WorkingTypeCode")
			Counter = Counter + 1
		End If
		sWorkingType = sWorkingType & oRS2.Collect("WorkingTypeName")
		If (oRS.Collect("OrderType") ="0" And oRS.Collect("Companykbn") = "2") Or oRS.Collect("OrderType") ="1" Or oRS.Collect("OrderType") ="2" Or oRS.Collect("OrderType") ="3" Then
			Select Case oRS2.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""haken_tr"">派遣とは</a>】</span>" 
				Case "002","003": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""shokai_tr"">人材紹介とは</a>】</span>" 
				Case "004": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>紹介予定派遣とは</a>】</span>" 
			End Select
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
	idx = 0
	sWorkingPlace = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True And idx < 20
		dbWorkingPlacePrefectureCode = ChkStr(oRS2.Collect("WorkingPlacePrefectureCode"))
		dbWorkingPlacePrefectureName = ChkStr(oRS2.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS2.Collect("WorkingPlaceCity"))
		'<勤務地アイコン>
		If InStr(sImgOrderState, "/icon_p" & dbWorkingPlacePrefectureCode & ".gif") = 0 Then
			'同じ都道府県アイコンは出さない！
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15"">&nbsp;"
		End If
		'</勤務地アイコン>

		'<勤務地>
		If sWorkingPlace <> "" Then sWorkingPlace = sWorkingPlace & "/"
		sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		'</勤務地>

		oRS2.MoveNext
		idx = idx + 1
	Loop
	If oRS2.RecordCount > 20 Then sWorkingPlace = sWorkingPlace & "&nbsp;...他"
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

    '<無期限延長機能対応>
    '2016/04/01 池田改修
    If sPublishLimitStr = "9999/12/31" Then
        '無期限の場合は、掲載期限に月末を指定。更新日に月初を指定。
        sPublishLimitStr = DateSerial(Year(Date()), Month(Date()) + 1, 0)
        'sUpdateDay       = DateSerial(Year(Date()), Month(Date()), 1)
    End If
    '</無期限延長機能対応>


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
			If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""></td></tr>"
			If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""></td></tr>"
			If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""></td></tr>"
			If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""></td></tr>"
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

	Response.Write "<input type=""hidden"" name=""CONF_OrderCodes"" value=""" & oRS.Collect("OrderCode") & """>"
	
    'クラス変更
    If sOrderType = "2" Then
    '紹介の時 
        HtmlClassName = "neo_shokai"
    Elseif sOrderType = "1" Then
    '派遣の時
        HtmlClassName = "neo_haken"
    Elseif sOrderType = "3" Then
    '紹介予定派遣の時
        HtmlClassName = "neo_ttp"
    Elseif sOrderType = "0" Then
    '広告のとき
        if HtmlWorkingType = "005" Then
            HtmlClassName = "neo_beit"
		Elseif HtmlWorkingType = "006" Then
			 HtmlClassName = "neo_soho"
        Else
            HtmlClassName = "neo_shain"
        End if
    End if


	'広告一覧

	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" And G_USEFLAG = "1" Then

		%>
<div class="my_order">
<div>
<span>情報コード</span>(<%= oRS.Collect("OrderCode") %>)
</div>
<div>
<span>状態 </span><%= sProgress %>
</div>
<div>
<select name="CONF_PublicFlags" <%= sPublicListDsp %>>
<%		If oRS.Collect("PublicFlag") = "1" Then		%>
			<option value="1" selected>掲載</option>
			<option value="0">非掲載</option>
<%		Else	%>
            <option value="1">掲載</option>
            <option value="0" selected>非掲載</option>
<%		End If	%>
</select>
</div>
<div>
<span>掲載日</span>	<%= sPublicDay %>
</div>
<div>
<span>登録日</span> <%= sRegistDay %>
</div>
<div>
<input type="checkbox" name="CONF_DeleteFlags" value="<%= oRS.Collect("OrderCode") %>">削除
</div>
</div>
<br clear="both">
<%	End If	%>

 <% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then  %>   
    
	<table border="0" class="old <%= HtmlClassName %>">
 <% else %>   
    <table border"0" class="old motto_old <%= HtmlClassName %>">
 <% end if %>
    
	<tbody>
	<tr>
	<td class="old11" valign="middle" colspan="2">
    
    <%

	If vUserType = "" Or vUserType = "staff" Then
		'非ログイン時、スタッフログイン時

		'・求人票ＵＲＬをメール送信
		'・ウォッチリストへ保存
		%>
		<div class="order_titele">
			<%= sTitleCompanyName %>
            <h3><a href="<%= HTTPS_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>">
			<% if sCatchCopy = "" then
			 response.write "詳細はこちら"
			  else
			  response.write sCatchCopy
			  end if %>
            </a><%= sImgMail %></h3>
		</div><!--/order_titele-->
		      
        <%
        	ElseIf vUserType = "company" Then
		'企業ログイン時
		%>
		<div class="order_titele">
			<%= sTitleCompanyName %>
            <h3><a href="<%= HTTPS_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>">
			<% if sCatchCopy = "" then
			 response.write "詳細はこちら"
			  else
			  response.write sCatchCopy
			  end if %>
            </a><%= sImgMail %></h3>
		</div><!--/order_titele-->
<%	End If %>
	<div class="support_type oiwai_<%= HimlOiwai %>">
    	
    </div>
	</td><!--/old11-->
	</tr>
	
    <tr>
	<td class="old12" colspan="2">
    <div class="order_state"><%= sImgOrderState %></div>
    <div class="publish_limit">掲載期限：<%= sPublishLimitStr %>　<span class="jSpan">情報コード</span>(<%= oRS.Collect("OrderCode") %>)</div> 
    <div class="arrow_img"></div>
    </td>
    </tr>
   
   <% If flgImg = True Then '画像が有る場合のレイアウト %>
    <tr>
    <td class="old27 td_point">
    <div>
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_typ.png">
    <p><%= sTitleJobName %></p>
    </div><!--職種-->
    
    <div>
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_emp.png">
    <p><%= sWorkingType %></p>
     </div>
     
     <div>
     <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_sal.png">
      <p id="salary">
    <% If sYearlyIncome <> "" Then %>
	<span>年収</span>
	<span><%= sYearlyIncome %></span><br>
    <% End If %>
    
    <% If sMonthlyIncome <> "" Then %>
	<span>月給</span>
	<span><%= sMonthlyIncome %></span><br>
    <% End If %>
    
    <% If sHourlyIncome <> "" Then %>
	<span>時給</span>
	<span><%= sHourlyIncome %></span>
    <% End If %>
     </p>
     </div>
    
    <div id="kinmuchi">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_loc.png">
    <p><%= sWorkingPlace %><%= sStationName%></p>
    </div>
    

    
    </td><!--/old27-->
    	<% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then  %> 
    		<td class="td_img">
    	<% else %>
    		<td class="td_img" style="background:#f5f5fa;">
    	<% end if %>
    
    <div id="pics_out">
        <div id="pics_main_out">
        	<a href="<%= HTTP_NAVI_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>" title="<%= sTitleCompanyName %>" style="width: 316px; height: 237px; display: block;<%= sImgMain %> background-size: contain; background-position: 50%;"></a>
        </div>
	<%= sImgSub %>
    </div>
	
    </td><!--/td_img-->
    </tr>
    <tr>
    	<td colspan="2" class="b_detail">
        	<b>【担当業務の説明】</b><br>
    		<%= sBusinessDetail %>
        </td>
    </tr>
    <tr>
    <td class="old28 td_point" colspan="2">
	<a href="<%= HTTPS_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>" class="neo_reg" target="_self">詳細をみる</a>
    
    
	
    
    <% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then 
    

			If flgAddWatchList = True Then
				Response.Write "<div class=""order_button""><span class=""m0 kentoZumi"">この求人票は既にお気に入りリストに登録済みです</span></div>"
			Else
			
				If vUserType = "staff" Then
					%>
                
                    <div class="order_button">
                        <!--<a href="#" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;">		
                            ウォッチリストに追加</a>-->
                        <form id="<%= sOrderCode %>frmSendMailJobOfferAddress">
                        	<%
								If sOrderType <> "0"  Then
							%>
                            	<input type="button" value="求人への問合せ" class="qmail" onclick="contactCompany('1',this);" name="<%= sOrderCode %>">
                            <%
								end if
							%>
                        
                            <input type="button" value="お気に入りリスト" onclick="kento(this)" class="kento" name="<%= sOrderCode %>">
                            <input type="hidden" name="CONF_OrderCode" value="<%= sOrderCode %>">
                        </form>
                    </div><!--/order_button-->
                <%
					
					else
					%>
                    	<div class="order_button">
							<form>
                            <%
							If sOrderType <> "0"  Then
							%>
                            	<input type="button" value="求人への問合せ" class="qmail" onclick="forRegQ(this)" name="<%= sOrderCode %>">
                            <%
							end if
							%>
                            
                            	<input type="button" value="お気に入りリスト" onclick="forReg(this)" class="kento" name="<%= sOrderCode %>">
                            </form>
                        </div>
					<%
					
				end if
			
				
			End If
		
		

	end if %>
    
    </td><!--/old28-->
    </tr>
		
        	

    

<% Else '画像が無い場合のレイアウト %>    
    
    <tr>
    <td class="old21 td_point">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_typ.png">
    <p><%= sTitleJobName %></p>
    </td><!--/old21-->
        <td class="old22 td_point">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_emp.png">
    <p><%= sWorkingType %></p>
    </td><!--/old22-->
    </tr> 
 
     <tr>
    <td class="old24 td_point">
   <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_sal.png">
    <p id="salary">
    <% If sYearlyIncome <> "" Then %>
	<span>年収</span>
	<span><%= sYearlyIncome %></span><br>
    <% End If %>
    
    <% If sMonthlyIncome <> "" Then %>
	<span>月給</span>
	<span><%= sMonthlyIncome %></span><br>
    <% End If %>
    
    <% If sHourlyIncome <> "" Then %>
	<span>時給</span>
	<span><%= sHourlyIncome %></span>
    <% End If %>
    </p>
    </td><!--/old24-->
    <td class="old23 td_point" id="kinmuchi2">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_loc.png">
    <p><%= sWorkingPlace %><%= sStationName%></p>
    </td><!--/old23-->
    </tr> 
    
      
    <tr>
    <td class="old25" colspan="2">
    <b>【担当業務の説明】</b><br>
    <%= sBusinessDetail %>
    </td><!--/old23-->
    </tr>
    
    <tr>
    <td class="old26" colspan="2">
	<a href="<%= HTTPS_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>" class="neo_reg" target="_self">詳細をみる</a>

	
	
   <% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then 
    

			If flgAddWatchList = True Then
				Response.Write "<div class=""order_button""><span class=""m0 kentoZumi"">この求人票は既にお気に入りリストに登録済みです</span></div>"
			Else
				If vUserType = "staff" Then
					%>
                
                    <div class="order_button">
                        <!--<a href="#" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;">		
                            ウォッチリストに追加</a>-->

                        <form id="<%= sOrderCode %>frmSendMailJobOfferAddress" method="post" action="../staff/watchlist_register.asp" onSubmit="return Submit();">
                        
                        	<%
								If sOrderType <> "0"  Then
							%>
                            	<input type="button" value="求人への問合せ" class="qmail" onclick="contactCompany('1',this);" name="<%= sOrderCode %>">
                            <%
								end if
							%>
                        
                            <input type="button" value="お気に入りリスト" onclick="kento(this)" class="kento" name="<%= sOrderCode %>">
                            <input type="hidden" name="CONF_OrderCode" value="<%= sOrderCode %>">
                        </form>
                    </div><!--/order_button-->
                <%
					
					else
					%>
						<div class="order_button">
							<form>
                            
                            <%
							If sOrderType <> "0"  Then
							%>
                            	<input type="button" value="求人への問合せ" class="qmail" onclick="forRegQ(this)" name="<%= sOrderCode %>">
                            <%
							end if
							%>
                            
                            	<input type="button" value="お気に入りリスト" onclick="forReg(this)" class="kento" name="<%= sOrderCode %>">
                            </form>
                        </div>
					<%
					
				end if
			End If
			

	end if %>

    </td><!--/old26-->
    </tr>
      
    
<% End If %>  

    </tbody>
    </table>




  
<%




	DspOrderListDetail = True
End Function

'******************************************************************************
'応募確認ページ
'******************************************************************************
Function DspOrderListDetail4(ByRef rDB, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Const PICSIZEW = 240
	Const PICSIZEH = 180
	Const PICSIZESUBW = 72
	Const PICSIZESUBH = 56

	Dim sSQL
	Dim oRS
	Dim oRS2
    Dim oRS3                '応募済画像表示に使用
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
	'HTTPクラス変更用変数
	Dim HtmlClassName
	Dim HtmlWorkingType
    Dim HimlOiwai

    
	'If GetRSState(rRS) = False Then Exit Function
	dbOrderCode = vMyOrder
    'response.Write vMyOrder
	DspOrderListDetail4 = False

	If G_USEFLAG = "0" And vMyOrder = "1" And G_OLDAPPLICATIONCODE <> "" Then
		sSQL = "EXEC up_DtlOrder_NEO '" & vMyOrder & "', '" & G_OLDAPPLICATIONCODE & "';"
	Else
		sSQL = "EXEC up_DtlOrder_NEO '" & vMyOrder & "', '';"
	End If

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
'デバックコード
'Response.write "sSQL:" & sSQL & "<br>"

	If GetRSState(oRS) = False Then Exit Function	'求人票が物理削除されている場合、応募データがあるのに、求人データを表示できない潜在的バグ回避のためコード追加 2014/07/25 池田

	dbCompanyCode = ChkStr(oRS.Collect("CompanyCode"))
	sOrderType    = ChkStr(oRS.Collect("OrderType"))
	sPlanType     = ChkStr(oRS.Collect("PlanTypeName"))
	iImageLimit   = oRS.Collect("ImageLimit")
	'お祝い金設定
	HimlOiwai = oRS.Collect("CongratulationPrice")

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
	sImgOrderState = GetImgOrderSpeciality(rDB, oRS)
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
		sImgMain = "<img src=""" & sImg1 & """>"
	End If

	If sImg2 <> "" Then
		sImgSub = sImgSub & "<div class=""sub_img"">" & _
			"<img src=""" & sImg2 & """><br>"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	If sImg3 <> "" Then
		sImgSub = sImgSub & "<div class=""sub_img"" style=""margin-top: 5px;"">" & _
			"<img src=""" & sImg3 & """><br>"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	'If sImg4 <> "" Then
	'	sImgSub = sImgSub & "<div class=""sub_img"">" & _
	'		"<img src=""" & sImg4 & """><br>"
	'	sImgSub = sImgSub & "</div>"
	'End If

	If sImgSub <> "" Then sImgSub =  sImgSub 
	'--------------------------------------------------------------------------
	'画像 end
	'**************************************************************************

	'**************************************************************************
	'担当業務 start
	'--------------------------------------------------------------------------
	If flgImg = True Then
		'画像が有る場合は文章を短めにカット
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),300) & "&nbsp;"
		If Len(sBusinessDetail) > 300 Then sBusinessDetail = sBusinessDetail & "..."
	Else
		'画像が無い場合は文章を長めにカット
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),465) & "&nbsp;"
		If Len(sBusinessDetail) > 465 Then sBusinessDetail = sBusinessDetail & "..."
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
	Dim Counter
	Counter = 1
	Do While GetRSState(oRS2) = True
		if Counter = 1 Then
			HtmlWorkingType = oRS2.Collect("WorkingTypeCode")
			Counter = Counter + 1
		End If
		sWorkingType = sWorkingType & oRS2.Collect("WorkingTypeName")
		If (oRS.Collect("OrderType") ="0" And oRS.Collect("Companykbn") = "2") Or oRS.Collect("OrderType") ="1" Or oRS.Collect("OrderType") ="2" Or oRS.Collect("OrderType") ="3" Then
			Select Case oRS2.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""haken_tr"">派遣とは</a>】</span>" 
				Case "002","003": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""shokai_tr"">人材紹介とは</a>】</span>" 
				Case "004": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>紹介予定派遣とは</a>】</span>" 
			End Select
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
	idx = 0
	sWorkingPlace = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True And idx < 20
		dbWorkingPlacePrefectureCode = ChkStr(oRS2.Collect("WorkingPlacePrefectureCode"))
		dbWorkingPlacePrefectureName = ChkStr(oRS2.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS2.Collect("WorkingPlaceCity"))
		'<勤務地アイコン>
		If InStr(sImgOrderState, "/icon_p" & dbWorkingPlacePrefectureCode & ".gif") = 0 Then
			'同じ都道府県アイコンは出さない！
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15"">&nbsp;"
		End If
		'</勤務地アイコン>

		'<勤務地>
		If sWorkingPlace <> "" Then sWorkingPlace = sWorkingPlace & "/"
		sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		'</勤務地>

		oRS2.MoveNext
		idx = idx + 1
	Loop
	If oRS2.RecordCount > 20 Then sWorkingPlace = sWorkingPlace & "&nbsp;...他"
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
			If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""></td></tr>"
			If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""></td></tr>"
			If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""></td></tr>"
			If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""></td></tr>"
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

	Response.Write "<input type=""hidden"" name=""CONF_OrderCodes"" value=""" & oRS.Collect("OrderCode") & """>"
	
    'クラス変更
    If sOrderType = "2" Then
    '紹介の時 
        HtmlClassName = "neo_shokai"
    Elseif sOrderType = "1" Then
    '派遣の時
        HtmlClassName = "neo_haken"
    Elseif sOrderType = "3" Then
    '紹介予定派遣の時
        HtmlClassName = "neo_ttp"
    Elseif sOrderType = "0" Then
    '広告のとき
        if HtmlWorkingType = "005" Then
            HtmlClassName = "neo_beit"
		Elseif HtmlWorkingType = "006" Then
			HtmlClassName = "neo_soho"
        Else
            HtmlClassName = "neo_shain"
        End if
    End if


    '******************************************************************************
    '2014/05/22 ikeda 追加
    '「応募済」と画像表示する
    '対象条件：
    '   既に応募済みの求人
    '   同じ企業の他求人に応募している場合
    '------------------------------------------------------------------------------
    Dim iExists_C_Flag  '同企業への応募データが存在する場合
    Dim iExists_O_Flag  '同求人への　　　　”
                        '0: なし 1:存在する
	sSQL = "EXEC up_ChkAdoptionExists '" & dbCompanyCode & "','" & vMyOrder & "','" & vUserID & "';"
	flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)

	iExists_C_Flag = oRS3.Collect("Exists_C_Flag")
	iExists_O_Flag = oRS3.Collect("Exists_O_Flag")

    Call RSClose(oRS3)

    '------------------------------------------------------------------------------
    '2014/05/22 ikeda ここまで
    '******************************************************************************


	'広告一覧

	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" And G_USEFLAG = "1" Then

		%>
<div class="my_order">
<div>
<span>情報コード</span>(<%= oRS.Collect("OrderCode") %>)
</div>
<div>
<span>状態 </span><%= sProgress %>
</div>
<div>
<select name="CONF_PublicFlags" <%= sPublicListDsp %>>
<%		If oRS.Collect("PublicFlag") = "1" Then		%>
			<option value="1" selected>掲載</option>
			<option value="0">非掲載</option>
<%		Else	%>
            <option value="1">掲載</option>
            <option value="0" selected>非掲載</option>
<%		End If	%>
</select>
</div>
<div>
<span>掲載日</span>	<%= sPublicDay %>
</div>
<div>
<span>登録日</span> <%= sRegistDay %>
</div>
<div>
<input type="checkbox" name="CONF_DeleteFlags" value="<%= oRS.Collect("OrderCode") %>">削除
</div>
</div>
<br clear="both">
<%	End If	%>

 <% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then  %>   
    
	<table border="0" class="old delSmart <%= HtmlClassName %>">
 <% else %>   
    <table border"0" class="old motto_old delSmart <%= HtmlClassName %>">
 <% end if %>
    
	<tbody>
	<tr>
	<td class="old11" valign="middle" colspan="2">

		<div class="order_titele">
			<%= sTitleCompanyName %>
            <h3><%= sCatchCopy %><%= sImgMail %></h3>
		</div><!--/order_titele-->

	<div class="support_type oiwai_<%= HimlOiwai %>">
    	
    </div>
	</td><!--/old11-->
	</tr>
	
    <tr>
	<td class="old12" colspan="2">
    <div class="order_state"><%= sImgOrderState %></div>
    <div class="publish_limit">掲載期限：<%= sPublishLimitStr %></div> 
    <div class="arrow_img"></div>

    <% '応募済み画像表示(同企業の他求人)
       If iExists_O_Flag = 1 Then %>
        <div class="adoption_img"></div>
    <% End If %>

    </td>
    </tr>
   

    <tr>
    <td class="old21 td_point">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_typ.png">
    <p><%= sTitleJobName %></p>
    </td><!--/old21-->
        <td class="old22 td_point">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_emp.png">
    <p><%= sWorkingType %></p>
    </td><!--/old22-->
    </tr> 
 
     <tr>
    <td class="old24 td_point">
   <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_sal.png">
    <p id="salary">
    <% If sYearlyIncome <> "" Then %>
	<span>年収</span>
	<span><%= sYearlyIncome %></span><br>
    <% End If %>
    
    <% If sMonthlyIncome <> "" Then %>
	<span>月給</span>
	<span><%= sMonthlyIncome %></span><br>
    <% End If %>
    
    <% If sHourlyIncome <> "" Then %>
	<span>時給</span>
	<span><%= sHourlyIncome %></span>
    <% End If %>
    </p>
    </td><!--/old24-->
    <td class="old23 td_point" id="kinmuchi2">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_loc.png">
    <p><%= sWorkingPlace %><%= sStationName%></p>
    </td><!--/old23-->
    </tr> 
    
      
    <tr>
    <td class="old25" colspan="2">
    <b>【担当業務の説明】</b><br>
    <%= sBusinessDetail %>
    </td><!--/old23-->
    </tr>


    </tbody>
    </table>




  
<%




	DspOrderListDetail4 = True
End Function



'******************************************************************************
'概　要：条件を指定して検索し直す
'引　数：rDB		：DB接続オブジェクト
'　　　：rRS		：求人票一覧のレコードセット
'　　　：vOrderCode	：現在の列数
'戻り値：
'備　考：
'履　歴：LIS K.NIINA
'　　　：2008/10/20 LIS K.Kokubo 勤務地複数化による修正
'******************************************************************************
Function Retrieval(byref rDB, byref rRS, ByVal vOrderCode)
	Dim oRS
	Dim sSQL
	Dim sError
	Dim sWT
	Dim sAC2
	Dim sJT2

	Dim dbWorkingPlacePrefectureCode

	'<勤務形態>
	sSQL = "EXEC sp_GetDataWorkingType '" & vOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = true Then
		sWT = oRS.Collect("WorkingTypeCode")
	End If
	Call RSClose(oRS)
	'</勤務形態>

	'<勤務地>
	sAC2 = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & vOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbWorkingPlacePrefectureCode = oRS.Collect("WorkingPlacePrefectureCode")

		If sAC2 <> "" Then sAC2 = sAC2 & ","
		sAC2 = sAC2 & dbWorkingPlacePrefectureCode
		oRS.MoveNext
	End If
	Call RSClose(oRS)
	'</勤務地>

	'<職種>
	sSQL = "sp_GetDataJobType '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = true Then
		sJT2 = oRS.Collect("JobTypeCode")
	End If
	Call RSClose(oRS)
	'</職種>

	Retrieval = "<div align=""right""><a href=""/order/order_list.asp?wt=" & sWT & "&amp;ac2=" & sAC2 & "&amp;jt2=" & sJT2 & """><img src=""/img/order_detail_icon/serchimage.gif"" border=""0"" style=""vertical-align:bottom;"">条件を指定して検索し直す⇒</a></div>"
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
'履　歴：
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
	Dim sMonthlyIncome		'月給
	Dim sMonthlyIncomeMin	'月給下限
	Dim sMonthlyIncomeMax	'月給上限
	Dim sDailyIncome		'日給
	Dim sDailyIncomeMin		'日給下限
	Dim sDailyIncomeMax		'日給上限
	Dim sHourlyIncome		'時給
	Dim sHourlyIncomeMin	'時給下限
	Dim sHourlyIncomeMax	'時給上限
	Dim sWorkingTypeIcon	'勤務形態アイコン並び
	Dim sImg				'画像URL

	Dim sURL				'求人票詳細のURL
	Dim sAlign				'枠寄せ [vCols = 1]left [vCols = vMaxCols]right [それ以外]center

	If GetRSState(rRS) = False Then Exit Function

	sURL = HTTPS_CURRENTURL & "order/order_detail.asp"

	If vCols = 1 Then
		sAlign = "left"
	ElseIf vCols = vMaxCols Then
		sAlign = "right"
	Else
		sAlign = "center"
	End If

	sSQL = "up_DtlOrder '" & rRS.Collect("OrderCode") & "', ''"
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
			<th>月給</th>
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
'履　歴：
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

			Response.Write "<div class=""sec_div"" style=""float:left; width:235px; margin:0 2px;""><div class=""thr_div"" style=""line-height:16px; " & aPadding(iCols) & """>"

			Response.Write "<div class=""jovimg"">" & aImg(idx) & "</div>"
			Response.Write "<div class=""jovtext"">"
			If aJobTypeDetail(idx) <> "" Then Response.Write "【職種】" & aJobTypeDetail(idx) & "<br>" & vbCrLf
			'If aCompanyName(idx) <> "" Then Response.Write "【企業】" & aCompanyName(idx) & "<br>" & vbCrLf
			If aWorkingTypeIcon(idx) <> "" Then Response.Write "【形態】" & aWorkingTypeIcon(idx)  & "<br>"& vbCrLf
			If aWorkingPlace(idx) <> "" Then Response.Write "【場所】" & aWorkingPlace(idx) & "<br>" & vbCrLf
			If aStation(idx) <> "" Then Response.Write "【最寄】" & Replace(aStation(idx), "<br>", "、") & "<br>" & vbCrLf
			Response.Write "【給与】"
			If aYearlyIncome(idx) <> "" Then Response.Write "[年収]" & aYearlyIncome(idx)
			If aMonthlyIncome(idx) <> "" Then Response.Write "[月給]" & aMonthlyIncome(idx)
			If aDailyIncome(idx) <> "" Then Response.Write "[日給]" & aDailyIncome(idx)
			If aHourlyIncome(idx) <> "" Then Response.Write "[時給]" & aHourlyIncome(idx)

			idx = idx + 1
			Response.Write "</div></div></div>"
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
<table style="margin:10px 0px;" class="smartPager">
	<tbody>
	<tr>
		<td style="width:88px; padding:5px; border-width:1px 0px 1px 1px; text-align:center;">
<%
	If vPage > 1 Then Response.Write "<a href='javascript:ChgPage(" & vPage - 1 & ");'>前のページ</a>"
%>
		</td>
		<td style="padding:5px; border-width:1px 0px 1px 0px; text-align:center;">
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
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vOrderCode		：
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：しごとナビ/order/company_order.asp
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'　　　：2008/06/25 LIS K.Kokubo 最寄駅追加
'******************************************************************************
Function DspCompanyInfo(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbPlanTypeName		'ライセンスプランタイプ
	Dim dbImageLimit		'最大画像掲載数
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
	Dim dbStationName1		'最寄駅１
	Dim dbToStation1		'最寄駅１から会社までの所要時間
	Dim dbToStationRemark1	'最寄駅１までの交通手段
	Dim dbStationName2		'最寄駅２
	Dim dbToStation2		'最寄駅２から会社までの所要時間
	Dim dbToStationRemark2	'最寄駅２までの交通手段

	Dim sNearbyStation		'最寄駅
	Dim sClass				'使用するスタイルシートのクラス　画像の有無で変化
	Dim sLineClass			'
	Dim flgLine				'線引きフラグ
	Dim sAddTitle			'派遣企業の情報の場合は「派遣」を項目名に付ける

	If GetRSState(rRS) = False Then Exit Function

	dbPlanTypeName = rRS.Collect("PlanTypeName")
	dbImageLimit = rRS.Collect("ImageLimit")

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

	'******************************************************************************
	'最寄駅 start
	'------------------------------------------------------------------------------
	sSQL = "/* ナビ：企業情報ページの最寄駅取得 */"
	sSQL = sSQL & "EXEC sp_GetDetailCompany '" & sCompanyCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sNearbyStation = ""
		dbStationName1 = ChkStr(oRS.Collect("StationName1"))
		dbStationName2 = ChkStr(oRS.Collect("StationName2"))
		If dbStationName1 & dbStationName2 <> "" And sOrderType = "0" Then
			dbToStation1 = ChkStr(oRS.Collect("WorkOrBus1"))
			dbToStationRemark1 = ChkStr(oRS.Collect("CompanySyudan1_1"))
			dbToStation2 = ChkStr(oRS.Collect("WorkOrBus2"))
			dbToStationRemark2 = ChkStr(oRS.Collect("CompanySyudan2_1"))

			If dbStationName1 <> "" Then
				If sNearbyStation <> "" Then sNearbyStation = sNearbyStation & "<br>"

				sNearbyStation = sNearbyStation & dbStationName1 & "駅"
				If dbToStation1 <> "" Then sNearbyStation = sNearbyStation & "(" & dbToStationRemark1 & dbToStation1 & "分)"
			End If

			If dbStationName2 <> "" Then
				If sNearbyStation <> "" Then sNearbyStation = sNearbyStation & "<br>"

				sNearbyStation = sNearbyStation & dbStationName2 & "駅"
				If dbToStation2 <> "" Then sNearbyStation = sNearbyStation & "(" & dbToStationRemark2 & dbToStation2 & "分)"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'最寄駅 end
	'******************************************************************************

	If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
		sClass = "value1"
		sLineClass = "odline2"
	Else
		sClass = "value2"
		sLineClass = "odline1"
	End If

	flgLine = False
	Response.Write "<div class=""companyblock"">"
	Response.Write "<h3>" & sAddTitle & "企業情報</h3>"
	'If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
'	Response.Write "<div id=""imgcompany""><img id=""imgcompany"" src=""" & HTTPS_NAVI_CURRENTURL & "company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"" alt=""イメージ写真""></div>"
'	End If
		
		%>
        <table id="company_code">
            <tbody>
            	<tr>
                	<th>企業コード</th>
                    <td><%= sCompanyCode %></td>
            	</tr>
                <tr>
                	<th>設立年度</th>
                    <td><%= sEstablishYear %></td>
                </tr>
                <tr>
                	<th>資本額</th>
                    <td><%= sCapitalAmount %></td>
                </tr>
                <tr>
                	<th>株式公開</th>
                    <td><%= sListClass %></td>
                </tr>
                <tr>
                	<th>社員数</th>
                    <td><%= sEmployeeNum %></td>
                </tr>
                <tr>
                	<th>業種</th>
                    <td><%= sIndustryType %></td>
                </tr>
                <tr>
                	<th>本社住所</th>
                    <td><%= sAddress %></td>
                </tr>
                <tr>
                	<th>本社最寄駅</th>
                    <td><%= sNearbyStation %></td>
                </tr>

                <!-- 2014/04/22 池田 企業HPから応募してしまう可能性があるので、コメントアウト
                <tr>
                	<th>コーポレートサイト</th>
                    <td>
                    <%
                        if replace(sHomePage," ","") <> "" then
                        %>
                            <a href="<%= sHomePage %>" target="_blank" rel="nofollow">この企業のホームページ</a>
                        <%
                        else
                            
                        end if
                    %>
                    </td>
                </tr>
                -->
                 
            </tbody>
        </table>
		<%


	Response.Write "</div>"

End Function

'******************************************************************************
'概　要：企業情報の基本情報を出力(NEO専用)
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vOrderCode		：
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：しごとナビ/order/company_order.asp
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'　　　：2008/06/25 LIS K.Kokubo 最寄駅追加
'******************************************************************************
Function DspCompanyInfoNEO(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbPlanTypeName		'ライセンスプランタイプ
	Dim dbImageLimit		'最大画像掲載数
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
	Dim dbStationName1		'最寄駅１
	Dim dbToStation1		'最寄駅１から会社までの所要時間
	Dim dbToStationRemark1	'最寄駅１までの交通手段
	Dim dbStationName2		'最寄駅２
	Dim dbToStation2		'最寄駅２から会社までの所要時間
	Dim dbToStationRemark2	'最寄駅２までの交通手段

	Dim sNearbyStation		'最寄駅
	Dim sClass				'使用するスタイルシートのクラス　画像の有無で変化
	Dim sLineClass			'
	Dim flgLine				'線引きフラグ
	Dim sAddTitle			'派遣企業の情報の場合は「派遣」を項目名に付ける

	If GetRSState(rRS) = False Then Exit Function

	dbPlanTypeName = rRS.Collect("PlanTypeName")
	dbImageLimit = rRS.Collect("ImageLimit")

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

	'******************************************************************************
	'最寄駅 start
	'------------------------------------------------------------------------------
	sSQL = "/* ナビ：企業情報ページの最寄駅取得 */"
	sSQL = sSQL & "EXEC sp_GetDetailCompany '" & sCompanyCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sNearbyStation = ""
		dbStationName1 = ChkStr(oRS.Collect("StationName1"))
		dbStationName2 = ChkStr(oRS.Collect("StationName2"))
		If dbStationName1 & dbStationName2 <> "" And sOrderType = "0" Then
			dbToStation1 = ChkStr(oRS.Collect("WorkOrBus1"))
			dbToStationRemark1 = ChkStr(oRS.Collect("CompanySyudan1_1"))
			dbToStation2 = ChkStr(oRS.Collect("WorkOrBus2"))
			dbToStationRemark2 = ChkStr(oRS.Collect("CompanySyudan2_1"))

			If dbStationName1 <> "" Then
				If sNearbyStation <> "" Then sNearbyStation = sNearbyStation & "<br>"

				sNearbyStation = sNearbyStation & dbStationName1 & "駅"
				If dbToStation1 <> "" Then sNearbyStation = sNearbyStation & "(" & dbToStationRemark1 & dbToStation1 & "分)"
			End If

			If dbStationName2 <> "" Then
				If sNearbyStation <> "" Then sNearbyStation = sNearbyStation & "<br>"

				sNearbyStation = sNearbyStation & dbStationName2 & "駅"
				If dbToStation2 <> "" Then sNearbyStation = sNearbyStation & "(" & dbToStationRemark2 & dbToStation2 & "分)"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'最寄駅 end
	'******************************************************************************

	If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
		sClass = "value1"
		sLineClass = "odline2"
	Else
		sClass = "value2"
		sLineClass = "odline1"
	End If

	flgLine = False
	Response.Write "<div class=""companyblock"">"
	Response.Write "<h3>" & sAddTitle & "企業情報</h3>"
	'If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
'	Response.Write "<div id=""imgcompany""><img id=""imgcompany"" src=""" & HTTPS_NAVI_CURRENTURL & "company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"" alt=""イメージ写真""></div>"
'	End If
		
		%>
        <table id="company_code">
            <tbody>
            	<tr>
                	<th>企業名</th>
                    <td><%= sCompanyName %></td>
            	</tr>
                <tr>
                	<th>設立年度</th>
                    <td><%= sEstablishYear %></td>
                </tr>
                <tr>
                	<th>資本額</th>
                    <td><%= sCapitalAmount %></td>
                </tr>
                <tr>
                	<th>株式公開</th>
                    <td><%= sListClass %></td>
                </tr>
                <tr>
                	<th>社員数</th>
                    <td><%= sEmployeeNum %></td>
                </tr>
                <tr>
                	<th>業種</th>
                    <td><%= sIndustryType %></td>
                </tr>
                <tr>
                	<th>本社住所</th>
                    <td><%= sAddress %></td>
                </tr>
                <tr>
                	<th>本社最寄駅</th>
                    <td><%= sNearbyStation %></td>
                </tr>

                <!-- 2014/04/22 池田 企業HPから応募してしまう可能性があるので、コメントアウト
                <tr>
                	<th>コーポレートサイト</th>
                    <td>
                    <%
                        if replace(sHomePage," ","") <> "" then
                        %>
                            <a href="<%= sHomePage %>" target="_blank" rel="nofollow">この企業のホームページ</a>
                        <%
                        else
                            
                        end if
                    %>
                    </td>
                </tr>
                -->
                 
            </tbody>
        </table>
		<%


	Response.Write "</div>"

End Function

'******************************************************************************
'概　要：企業情報のＰＲ情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vOrderCode		：
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使用元：しごとナビ/order/company_order.asp
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'　　　：2009/01/06 LIS K.Kokubo 福利厚生備考追加
'******************************************************************************
Function DspCompanyPR(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Const WELFARECOL = "3"	'福利厚生の１行あたりの列数

	Dim sOrderType			'受注種類
	Dim sCompanyKbn			'企業区分
	Dim sBusiness			'事業内容
	Dim sPR					'企業紹介
	Dim sAtmosphere			'会社の雰囲気
	Dim sWelfare			'福利厚生
	Dim iWelfare			'福利厚生カウント
	Dim sWelfareProgramRemark'福利厚生備考
	Dim idx
	Dim flgPR
	Dim flgLine				'線引きフラグ
	Dim sClass
	Dim sAddTitle			'派遣企業の情報の場合は「派遣企業の」を項目名に付ける

	If GetRSState(rRS) = False Then Exit Function

	sOrderType = rRS.Collect("OrderType")
	sCompanyKbn = rRS.Collect("CompanyKbn")

	If sOrderType = "0" And sCompanyKbn = "4" Then sAddTitle = "派遣企業の"

	'<事業内容>
	sBusiness = ""
	sBusiness = Replace(ChkStr(rRS.Collect("BusinessContents")), vbCrLf, "<br>")
	sBusiness = Replace(sBusiness, vbCr, "<br>")
	sBusiness = Replace(sBusiness, vbLf, "<br>")
	'</事業内容>

	'<会社紹介>
	sPR = ""
	sPR = Replace(ChkStr(rRS.Collect("CompanyPR")), vbCrLf, "<br>")
	sPR = Replace(sPR, vbCr, "<br>")
	sPR = Replace(sPR, vbLf, "<br>")
	'</会社紹介>

	'<会社の雰囲気>
	sAtmosphere = ""
	sAtmosphere = Replace(ChkStr(rRS.Collect("Atmosphere")), vbCrLf, "<br>")
	sAtmosphere = Replace(sAtmosphere, vbCr, "<br>")
	sAtmosphere = Replace(sAtmosphere, vbLf, "<br>")
	'</会社の雰囲気>

	'<福利厚生>
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

	'**TOP 08/08/19 Lis林 DEL
	'If ChkStr(rRS.Collect("FlexTimeFlag")) = "1" Then
	'	iWelfare = iWelfare + 1
	'	If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
	'	sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">フレックスタイム</p></td>"
	'	If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	'End If
	'**BTM 08/08/19 Lis林 DEL

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

	sWelfareProgramRemark = Replace(ChkStr(rRS.Collect("WelfareProgramRemark")),VbCrLf,"<br>")
	'<福利厚生>
	
	%>
    <div class="companyblock">
        <h3><%= sAddTitle %>PR情報</h3>
        <table id="company_code">
            <tbody>
                <tr>
                    <th>事業内容</th>
                    <td><%= sBusiness %></td>
                </tr>
                <tr>
                    <th>会社PR</th>
                    <td><%= sPR %></td>
            	</tr>
                <tr>
                    <th>会社の雰囲気</th>
                    <td><%= sAtmosphere %></td>
            	</tr>
                <tr>
                    <th>福利厚生</th>
                    <td><%= sWelfare & sWelfareProgramRemark %></td>
            	</tr>
            </tbody>
        </table>
	</div>
	<%

End Function

'******************************************************************************
'概　要：求人票詳細ページのリスの紹介先・派遣先企業情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
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
	Dim sCapitalAmount		'資本額		'**TOP 08/08/21 Lis林 ADD
	Dim sEmployeeNum		'社員数
	Dim sAccountingPeriod1	'決算期1
	Dim sSalesAmount1		'売上高1
	Dim sOrdinaryProfit1	'経常利益1
	Dim sAccountingPeriod2	'決算期2
	Dim sSalesAmount2		'売上高2
	Dim sOrdinaryProfit2	'経常利益2
	Dim sAccountingPeriod3	'決算期3
	Dim sSalesAmount3		'売上高3
	Dim sOrdinaryProfit3	'経常利益3
	Dim sImportantNotice	'特記事項
	Dim sflgAct							'**BTM 08/08/21 Lis林 ADD
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
			
		Else
			sImgTitle = "/img/order/lisorderpr1.gif"
			
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
		'業種 end
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
		'**TOP 08/08/21 Lis林 ADD
		'******************************************************************************
		'資本額 start
		'------------------------------------------------------------------------------
		sCapitalAmount = ""
		sCapitalAmount = ChkStr(rRS.Collect("CapitalAmount"))
		if IsNumeric(sCapitalAmount) = True then
			sCapitalAmount = GetJapaneseYen(sCapitalAmount)
		elseif sCapitalAmount <> "" then
			if InStr(sCapitalAmount,"円") > 0 then		'"円"が入っていたらそのまま
			else
				sCapitalAmount = sCapitalAmount & "円"
			end if
		end if
		'------------------------------------------------------------------------------
		'資本額 end
		'******************************************************************************

		'******************************************************************************
		'社員数 start
		'------------------------------------------------------------------------------
		sEmployeeNum = ""
		sEmployeeNum = ChkStr(rRS.Collect("AllEmployeeNum"))
		If sEmployeeNum <> "" Then sEmployeeNum = sEmployeeNum & "人"
		'------------------------------------------------------------------------------
		'社員数 end
		'******************************************************************************
		
		'******************************************************************************
		'決算期・売上高・経常利益 start
		'------------------------------------------------------------------------------
		sAccountingPeriod1 = ""
		sSalesAmount1 = ""
		sOrdinaryProfit1 = ""
		sAccountingPeriod2 = ""
		sSalesAmount2 = ""
		sOrdinaryProfit2 = ""
		sAccountingPeriod3 = ""
		sSalesAmount3 = ""
		sOrdinaryProfit3 = ""
		sImportantNotice = ""
		sAccountingPeriod1 = ChkStr(rRS.Collect("AccountingPeriod1"))
		sSalesAmount1 = ChkStr(rRS.Collect("SalesAmount1"))
		'if sSalesAmount1 <> "" and InStr(sSalesAmount1,"円") <= 0 then sSalesAmount1 = sSalesAmount1 & "円"
		sOrdinaryProfit1 = ChkStr(rRS.Collect("OrdinaryProfit1"))
		'if sOrdinaryProfit1 <> "" and InStr(sOrdinaryProfit1,"円") <= 0 then sOrdinaryProfit1 = sOrdinaryProfit1 & "円"
		sAccountingPeriod2 = ChkStr(rRS.Collect("AccountingPeriod2"))
		sSalesAmount2 = ChkStr(rRS.Collect("SalesAmount2"))
		'if sSalesAmount2 <> "" and InStr(sSalesAmount2,"円") <= 0 then sSalesAmount2 = sSalesAmount2 & "円"
		sOrdinaryProfit2 = ChkStr(rRS.Collect("OrdinaryProfit2"))
		'if sOrdinaryProfit2 <> "" and InStr(sOrdinaryProfit2,"円") <= 0 then sOrdinaryProfit2 = sOrdinaryProfit2 & "円"
		sAccountingPeriod3 = ChkStr(rRS.Collect("AccountingPeriod3"))
		sSalesAmount3 = ChkStr(rRS.Collect("SalesAmount3"))
		'if sSalesAmount3 <> "" and InStr(sSalesAmount3,"円") <= 0 then sSalesAmount3 = sSalesAmount3 & "円"
		sOrdinaryProfit3 = ChkStr(rRS.Collect("OrdinaryProfit3"))
		'if sOrdinaryProfit3 <> "" and InStr(sOrdinaryProfit3,"円") <= 0 then sOrdinaryProfit3 = sOrdinaryProfit3 & "円"
		sImportantNotice = ChkStr(rRS.Collect("ImportantNotice"))
		'------------------------------------------------------------------------------
		'決算期・売上高・経常利益 end
		'******************************************************************************
		'**BTM 08/08/21 Lis林 ADD
	End If

	flgLine = False

	'**TOP 08/08/21 Lis林 REP
	'If sListClass & sIndustryType & sPR <> "" Then
	If sListClass & sIndustryType & sPR & sCapitalAmount & sEmployeeNum <> "" or _
		(InStr(sImportantNotice,"非公開") <= 0 and _
		((sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"非公開") <= 0) or _
		 (sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"非公開") <= 0) or _
		 (sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"非公開") <= 0))) Then
	'**BTM 08/08/21 Lis林 REP
		DspLisOrderCompanyInfo = True
%>
	<img src="/img/order/tab_detail_sk.png" class="tab_img">
    <table class="detail_table">
	<tbody>
    <tr>
    <th class="fast_th"><%= sIntrDisp %></th>
    <td>
    </td>
    </tr>
    <% If sListClass <> "" Then %>
    <tr>
    <th class="dborder_bottom">株式公開</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sListClass %></p>
    </td>
    <% End If %>
    
	<% If sIndustryType <> "" Then %>
    <tr>
    <th class="dborder_bottom">業種</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sIndustryType %></p>
    </td>
	</tr>     
   	<% End If %>
    
   	<% If sPR <> "" Then %>
    <tr>
    <th class="dborder_bottom">事業内容</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sPR %></p>
    </td>
	</tr>     
   	<% End If %>
    
   	<% If sCapitalAmount <> "" Then %>
    <tr>
    <th class="dborder_bottom">資本金</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sCapitalAmount %></p>
    </td>
	</tr>     
   	<% End If %>    

   	<% If sEmployeeNum <> "" Then %>
    <tr>
    <th class="dborder_bottom">社員数</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sEmployeeNum %></p>
    </td>
	</tr>     
   	<% End If %> 
    
   	<% sflgAct = ""
		If InStr(sImportantNotice,"非公開") <= 0 and _
		((sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"非公開") <= 0) or _
		 (sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"非公開") <= 0) or _
		 (sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"非公開") <= 0)) then %>
    <tr>
    <th class="dborder_bottom">売上実績</th>
    <td class="dborder_bottom">
    <p class="m0"><% 			'売上高１・経常利益１・決算期１
			if sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"非公開") <= 0 then
				if sSalesAmount1 <> "" and InStr(sSalesAmount1,"非公開") <= 0 then
					response.write "売上高：" & sSalesAmount1 & "　"
				end if
				if sOrdinaryProfit1 <> "" and InStr(sOrdinaryProfit1,"非公開") <= 0 then
					response.write "経常利益：" & sOrdinaryProfit1
				end if
				if sAccountingPeriod1 <> "" and InStr(sAccountingPeriod1,"非公開") <= 0 then
					response.write "（決算期：" & sAccountingPeriod1 & "）<br>"
				end if
				sflgAct = "1"
			end if
			'売上高２・経常利益２・決算期２
			if sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"非公開") <= 0 then
				if sSalesAmount2 <> "" and InStr(sSalesAmount2,"非公開") <= 0 then
					response.write "売上高：" & sSalesAmount2 & "　"
				end if
				if sOrdinaryProfit2 <> "" and InStr(sOrdinaryProfit2,"非公開") <= 0 then
					response.write "経常利益：" & sOrdinaryProfit2
				end if
				if sAccountingPeriod2 <> "" and InStr(sAccountingPeriod2,"非公開") <= 0 then
					response.write "（決算期：" & sAccountingPeriod2 & "）<br>"
				end if
				sflgAct = "1"
			end if
			'売上高３・経常利益３・決算期３
			if sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"非公開") <= 0 then
				if sSalesAmount3 <> "" and InStr(sSalesAmount3,"非公開") <= 0 then
					response.write "売上高：" & sSalesAmount3 & "　"
				end if
				if sOrdinaryProfit3 <> "" and InStr(sOrdinaryProfit3,"非公開") <= 0 then
					response.write "経常利益：" & sOrdinaryProfit3
				end if
				if sAccountingPeriod3 <> "" and InStr(sAccountingPeriod3,"非公開") <= 0 then
					response.write "（決算期：" & sAccountingPeriod3 & "）<br>"
				end if
				sflgAct = "1"
			end if
			'特記事項
			If sflgAct = "1" and sImportantNotice <> "" and InStr(sImportantNotice,"非公開") <= 0 then
				response.write "（"
				if InStr(sImportantNotice,"※") <= 0 then response.write "※"
				response.write  sImportantNotice & "）<br>"
			End If
			 %></p>
    </td>
	</tr>     
   	<% End If %>     
       
    </tbody>
    </table>
<p class="m0">※人材<% If sOrderType = "2" Then response.write "紹介" Else response.write "派遣" End If %>でご案内するお仕事のため、詳しい会社情報は下のボタンやお電話などで直接お問合せください。</p>
<div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div> 

<%

	End If
End Function

'******************************************************************************
'概　要：派遣企業の派遣先企業情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder		：自社求人票フラグ
'備　考：
'使用元：しごとナビ/order/company_order.asp
'履　歴：2007/02/11 LIS K.Kokubo 作成
'******************************************************************************
Function DspTempOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Dim dbOrderCode			'情報コード
	Dim dbTempCompanyName
	Dim dbTempCompanyName_F
	Dim dbTempEstablishYear
	Dim dbTempIndustryTypeName
	Dim dbTempCapitalAmount
	Dim dbTempForeinCapital
	Dim dbTempListClass
	Dim dbTempAllEmployeeNumber
	Dim dbTempHomepageAddress
	Dim dbTempPost_U
	Dim dbTempPost_L
	Dim dbTempPrefectureCode
	Dim dbTempCity
	Dim dbTempCity_F
	Dim dbTempTown
	Dim dbTempAddress
	Dim dbTempTelephoneNumber

	Dim sClearSolid
	Dim flgLine				'線引きフラグ
	Dim flgData
	Dim sCapital
	Dim sTempAllEmployeeNumber

	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

	flgData = False

	'<派遣先企業情報取得>
	dbOrderCode = ChkStr(rRS.Collect("OrderCode"))
	'dbTempCompanyName = ChkStr(rRS.Collect("TempCompanyName"))
	'dbTempCompanyName_F = ChkStr(rRS.Collect("TempCompanyName_F"))
	dbTempEstablishYear = ChkStr(rRS.Collect("TempEstablishYear"))
	dbTempIndustryTypeName = ChkStr(rRS.Collect("TempIndustryTypeName"))
	dbTempCapitalAmount = ChkStr(rRS.Collect("TempCapitalAmount"))
	dbTempForeinCapital = ChkStr(rRS.Collect("TempForeinCapital"))
	dbTempListClass = ChkStr(rRS.Collect("TempListClass"))
	dbTempAllEmployeeNumber = ChkStr(rRS.Collect("TempAllEmployeeNumber"))
	'dbTempHomepageAddress = ChkStr(rRS.Collect("TempHomepageAddress"))
	'dbTempPost_U = ChkStr(rRS.Collect("TempPost_U"))
	'dbTempPost_L = ChkStr(rRS.Collect("TempPost_L"))
	'dbTempPrefectureCode = ChkStr(rRS.Collect("TempPrefectureCode"))
	'dbTempCity = ChkStr(rRS.Collect("TempCity"))
	'dbTempCity_F = ChkStr(rRS.Collect("TempCity_F"))
	'dbTempTown = ChkStr(rRS.Collect("TempTown"))
	'dbTempAddress = ChkStr(rRS.Collect("TempAddress"))
	'dbTempTelephoneNumber = ChkStr(rRS.Collect("TempTelephoneNumber"))
	'</派遣先企業情報取得>

	'<設立年度>
	If dbTempEstablishYear <> "" Then
		dbTempEstablishYear = dbTempEstablishYear & "年"
		flgData = True
	End If
	'</設立年度>

	'<業種>
	If dbTempIndustryTypeName <> "" Then
		flgData = True
	End If
	'</業種>

	'<資本>
	sCapital = ""
	If dbTempCapitalAmount & dbTempForeinCapital <> "" Then
		If dbTempCapitalAmount <> "" Then
			sCapital = sCapital & GetJapaneseYen(dbTempCapitalAmount)
		End If

		If dbTempForeinCapital <> "" Then
			sCapital = sCapital & "&nbsp;（外資：" & dbTempForeinCapital & "）"
		End If

		flgData = True
	End If
	'</資本>

	'<株式>
	If dbTempListClass <> "" Then
		flgData = True
	End If
	'</株式>

	'<社員数>
	If dbTempAllEmployeeNumber <> "" Then
		sTempAllEmployeeNumber = dbTempAllEmployeeNumber & "人"
		flgData = True
	End If
	'</社員数>

	flgLine = False

	If flgData = True Then
		sHTML = sHTML & "<h3>派遣先企業情報</h3>" & vbCrLf

		If dbTempEstablishYear <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>設立年度</h4></div>"
			sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempEstablishYear & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If dbTempIndustryTypeName <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>業種</h4></div>"
			sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempIndustryTypeName & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sCapital <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>資本金</h4></div>"
			sHTML = sHTML & "<div class=""value1"">"
			sHTML = sHTML & "<p class=""m0"">" & sCapital & "</p>"
			sHTML = sHTML & "</div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If dbTempListClass <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>株式</h4></div>"
			sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempListClass & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sTempAllEmployeeNumber <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>社員数</h4></div>"
			sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sTempAllEmployeeNumber & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		sHTML = sHTML & "<br>" & vbCrLf
	End If

	Response.Write sHTML
End Function

'******************************************************************************
'概　要：求人票コントロールボタン
'引　数：rDB				：接続中のDBConnection
'　　　：rRS				：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType			：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID			：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder			：自社求人票か否か ["1"]自社求人票 ["0"]自社求人票でない
'　　　：vJobTypeLimitFlag	：職種数が制限を越えていないか ["1"]OK ["0"]NO
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'履　歴：2007/02/11 LIS K.Kokubo 作成
'　　　：2009/03/11 LIS K.Kokubo 変更 インタビュー編集ボタンの表示方法変更(ナビ無料化対応)
'******************************************************************************
Function DspOrderControlButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder, ByVal vJobTypeLimitFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbPlanTypeName		'ライセンスプラン種類
	Dim sOrderCode
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'受注種類
	Dim sPermitFlag			'掲載許可フラグ
	Dim sPublicFlag			'掲載フラグ
	Dim sRiyoFlag			'掲載開始日
	Dim sHakouFlag			'利用開始日（ライセンス発効日）
	Dim dbSearchName		'保存検索条件名
	Dim dbSearchParam		'保存検索条件パラメータ

	Dim flgAddWatchList
	Dim iMailTemplateCnt	'メールテンプレートの件数
	Dim sAncMT				'メールテンプレートへのリンク

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	dbPlanTypeName = rRS.Collect("PlanTypeName")
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
	sSQL = "EXEC up_ChkWatchListExists_Staff '" & vUserID & "', '" & sOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
	End If
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	
End Function

'******************************************************************************
'概　要：求人票コントロールボタン2
'引　数：rDB				：接続中のDBConnection
'　　　：rRS				：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType			：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID			：利用中ユーザのユーザID [Session("userid")]
'　　　：vMyOrder			：自社求人票か否か ["1"]自社求人票 ["0"]自社求人票でない
'　　　：vJobTypeLimitFlag	：職種数が制限を越えていないか ["1"]OK ["0"]NO
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'履　歴：応募ボタンのみ
'******************************************************************************
Function DspOrderControlButton2(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder, ByVal vJobTypeLimitFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbPlanTypeName		'ライセンスプラン種類
	Dim sOrderCode
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'受注種類
	Dim sPermitFlag			'掲載許可フラグ
	Dim sPublicFlag			'掲載フラグ
	Dim sRiyoFlag			'掲載開始日
	Dim sHakouFlag			'利用開始日（ライセンス発効日）
	Dim dbSearchName		'保存検索条件名
	Dim dbSearchParam		'保存検索条件パラメータ

	Dim flgAddWatchList
	Dim iMailTemplateCnt	'メールテンプレートの件数
	Dim sAncMT				'メールテンプレートへのリンク

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	dbPlanTypeName = rRS.Collect("PlanTypeName")
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
	sSQL = "EXEC up_ChkWatchListExists_Staff '" & vUserID & "', '" & sOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
	End If
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************
	
	Dim qsOrderCode				'オーダーコード(受注表番号)
	Dim iDetail				'求人票詳細からのフラグ
	
	qsOrderCode = GetForm("ordercode", 2)
	iDetail = GetForm("Detail", 2)
	
	
	
	
	
	
	
	
	If vUserType = "staff" Then
	
		'******************************************************************************
		'ログイン求職者の場合 start
		'------------------------------------------------------------------------------
'		Response.Write "<div id=""login_watch"">"
'		If rRS.Collect("NowPublicFlag") = "1" Then
'
'            if sOrderType = "0" Then
'			    Response.Write "<input type=""button"" value=""この求人に応募メールを送信する"" onclick=""contactCompanyAdv('');"" class=""obo"">"
'            Else
'                Response.Write "<input type=""button"" value=""この求人に応募メールを送信する"" onclick=""contactCompanyLis('');"" class=""obo"">"
'			    Response.Write "<input type=""button"" value=""この求人の質問メールを送信する"" onclick=""contactCompany('1');"" class=""qmail"">"
'            End If
'
'
'			If flgAddWatchList = True Then
'				Response.Write "<span class=""m0 kentoZumi"">この求人票は既にお気に入りリストに登録済みです</span>"
'			Else
'				Response.Write "<input type=""button"" value=""お気に入りリスト"" onclick=""ListAdd()"" class=""kento"">"
'			End If
'			Response.Write "<br clear=""both"">"
'
'		Else
'			Response.Write "<div class=""description"" align=""center""><b>この求人票は掲載が終了しています。メール送信はできません。</b></div>"
'		End If
'		Response.Write "</div>"
		'------------------------------------------------------------------------------
		'ログイン求職者の場合 end
		'******************************************************************************

		'******************************************************************************
		'ログイン求職者の場合 start
		'------------------------------------------------------------------------------
		Response.Write "<div id=""login_watch"">"
		Response.Write "<ul>"
		If rRS.Collect("NowPublicFlag") = "1" Then

            if sOrderType = "0" Then
			    Response.Write "<li><input type=""button"" value=""この求人に応募メールを送信する"" onclick=""contactCompanyAdv('');"" class=""obo""></li>"
            Else
                Response.Write "<li><input type=""button"" value=""この求人に応募メールを送信する"" onclick=""contactCompanyLis('');"" class=""obo""></li>"
			    Response.Write "<li><input type=""button"" value=""この求人の質問メールを送信する"" onclick=""contactCompany('1');"" class=""qmail""></li>"
            End If


			If flgAddWatchList = True Then
				Response.Write "<li><span class=""m0 kentoZumi"">この求人票は既にお気に入りリストに登録済みです</span></li>"
			Else
				response.write "<li>"
				'2017/04/04 木村：お気に入り追加がIEのみ機能しないのを修正　※onSubmit="return Submit();"は入力チェック用JSなので外しました
				'response.write "<form id=""frmSendMailJobOfferAddress"" method=""post"" action=""../staff/watchlist_register.asp"" onSubmit=""return Submit();"">"
				'Response.Write "<input type=""button"" value=""お気に入りリスト"" onclick=""document.forms.frmSendMailJobOfferAddress.submit();return false;"" class=""kento"">"
				'response.write "<input type=""hidden"" name=""CONF_OrderCode"" value='"& qsOrderCode &"'>"

				response.write "<form id=""frmSendMailJobOfferAddress"" method=""post"" action=""../staff/watchlist_register.asp"" onSubmit=""document.forms.frmSendMailJobOfferAddress.submit();return false;"">"
				Response.Write "<input type=""submit"" value=""お気に入りリスト"" class=""kento"">"
				response.write "<input type=""hidden"" name=""CONF_OrderCode"" value='"& qsOrderCode &"'>"

				response.write "</form></li>"
			End If
			

		Else
			Response.Write "<li id=""finKokoku""><div class=""description"" align=""center""><b>この求人票は掲載が終了しています。メール送信はできません。</b></div></li>"
		End If
		Response.Write "<br clear=""both"">"
		Response.Write "</ul></div>"
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
'　　　：rRS				：up_DtlOrderで生成されたレコードセットオブジェクト
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

	If vUserType = "staff" Then
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
	MailWin = window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompany.asp?' + sQ + 'ordercode=<%= sOrderCode %>','_blank');
}
function contactCompanyAdv(vflg) {
    var sQ = '';
    if (vflg) {
        if (vflg.length > 0) sQ = 'q=1&';
    }

//  別ウィンドウ表示から同ウィンドウ内ページ遷移へ修正　池田 2014/04/07    
//  MailWin = window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompanyAdv.asp?' + sQ + 'ordercode=<%= sOrderCode %>', '_blank');
    document.location = '<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompanyAdv.asp?' + sQ + 'ordercode=<%= sOrderCode %>';
    
}
function contactCompanyLis(vflg) {
    var sQ = '';
    if (vflg) {
        if (vflg.length > 0) sQ = 'q=1&';
    }
    MailWin = window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompanyLis.asp?' + sQ + 'ordercode=<%= sOrderCode %>', '_blank');
}
function ListAdd() {
    MailWin2 = window.open('<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?ordercode=<%= sOrderCode %>', 'list', 'menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=470');
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
'　　　：rRS				：up_DtlOrderで生成されたレコードセットオブジェクト
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
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'　　　：2008/03/04 LIS K.Kokubo 掲載終了日を[RiyoToDate]→[DspPublicLimitDay]に変更
'　　　：2009/03/18 LIS K.Kokubo vReplaceFlag追加
'******************************************************************************
Function DspOrderCompanyName(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vReplaceFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType
	Dim dbSecretFlag		'シークレットフラグ
	Dim sCompanyCode		'企業コード
	Dim sCompanyName		'企業名称
	Dim sCompanyNameF		'企業名称カナ
	Dim sCompanyKbn			'企業区分
	Dim sCompanySpeciality	'企業特徴
	Dim sPublishLimitStr	'掲載期限表示用文字列
	Dim sCautionStr			'掲載期限表示注意文言文字列
	Dim dbTempOrderFlag		'派遣案件フラグ
	Dim dbTTPOrderFlag		'紹介予定派遣案件フラグ
	Dim flgNowPublic		'現在掲載中の求人票判定 '[True]掲載中 [False]非掲載

	Dim sUpdateDay
	Dim vAccessCount
	
	If GetRSState(rRS) = False Then Exit Function

	dbSecretFlag = rRS.Collect("SecretFlag")

	'******************************************************************************
	'会社名 start
	'------------------------------------------------------------------------------
	sCompanyName = rRS.Collect("CompanyName")
	sCompanyNameF = rRS.Collect("CompanyName_F")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	sCompanySpeciality = rRS.Collect("CompanySpeciality")
	sOrderType = rRS.Collect("OrderType")
	dbTempOrderFlag = rRS.Collect("TempOrderFlag")
	dbTTPOrderFlag = rRS.Collect("TTPOrderFlag")

	If vReplaceFlag = "1" Then
		Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	End If
	'------------------------------------------------------------------------------
	'会社名 end
	'******************************************************************************

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">※期限前に掲載終了する場合があります。</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならDspPublicLimitDayを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
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
<!--
<div style="margin-bottom:10px;">
<%
	'リス紹介案件,人材会社紹介案件の場合は「転職相談案件」イメージを表示
	If sOrderType = "2" Or (sCompanyKbn = "2" And dbTempOrderFlag = "0" And dbTTPOrderFlag = "0") Then
		Response.Write "<div style=""text-align:left;padding-bottom:10px;font-size:8pt; color:#666666;"">"
		Response.Write "<img src=""/img/order/counselable_order.gif"" width=""150"" height=""25"" alt=""転職支援を受けて応募する求人です"">"
		Response.Write "&nbsp;&nbsp;［人材会社の転職支援を受けて応募できる求人］&nbsp;&nbsp;&nbsp;"
	else
		Response.Write "<div style=""font-size:8pt;color:#666666;text-align:right"">"
	end if


	'Twitterフォローボタン
	if Request.ServerVariables("HTTPS") <> "on" then
		Response.Write "<div style=""float:right;width:160px"">"
		
		Select Case rRS.Collect("BranchCode")
			Case "OY"
				Response.Write "<a href=""http://"
				Response.Write "twitter.com/navi_okayama"" class=""twitter-follow-button"" data-show-count=""false"" data-lang=""ja"""
				Response.Write ">Follow @navi_okayama</a><script type=""text/javascript"" src="""
				Response.Write "http://platform.twitter.com/widgets.js"" charset=""utf-8""></script>"
			Case Else
				Response.Write "<a href=""http://"
				Response.Write "twitter.com/shigoto_navi"" class=""twitter-follow-button"" data-show-count=""false"" data-lang=""ja"""
				Response.Write ">Follow @shigoto_navi</a><script type=""text/javascript"" src="""
				Response.Write "http://platform.twitter.com/widgets.js"" charset=""utf-8""></script>"
		End Select
		Response.Write "</div>"
	end if
	'FaceBookボタン
	Response.write "<iframe src="""
	if Request.ServerVariables("HTTPS") = "on" then
		Response.Write "https://"
	else
		Response.Write "http://"
	end if
	Response.Write "www.facebook.com/plugins/like.php?href=http%3a%2f%2fwww.shigotonavi.co.jp%2forder%2forder_detail.asp%3fOrderCode%3d" & rRS.Collect("OrderCode")
	Response.Write "&amp;layout=button_count&amp;show_faces=true&amp;width=30&amp;action=like&amp;colorscheme=light&amp;height=21"
	Response.Write """ scrolling=""no"" frameborder=""0"" style=""float:right;border:none; overflow:hidden; width:80px; height:21px;"" allowTransparency=""true""></iframe>"

    'Google+1ボタン
    Response.Write "<div style=""float:right"">"
    Response.Write "<g:plusone size=""medium"" count=""true"" href=""http://www.shigotonavi.co.jp/""></g:plusone>"
    Response.Write "</div>"


	Response.Write "</div>"

	'シークレット求人の場合は「シークレット求人」イメージを表示
	'If dbSecretFlag = "1" Then Response.Write "<img src=""/img/order/secret_order.gif"" width=""150"" height=""25"" alt=""この求人からスカウトを受けた人だけが閲覧できる求人です"">"
	If dbSecretFlag = "1" Then Response.Write "<p class=""m0"" style=""color:#ff9933; font-weight:bold;"">■スカウトを受けた人だけが閲覧できる求人情報です。</p>"

	If vUserType = "" Or vUserType = "staff" Then
		'非ログイン時、スタッフログイン時

		If G_USERID <> "" And flgNowPublic = True Then
			'しごとナビにログイン中の場合は、企業名＋掲載期限＋求人票ＵＲＬメール送信
%>
	<div style="width:400px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="width:200px; float:left;">
		<div style="float:right; padding:0px;">

<%
	if Request.ServerVariables("HTTPS") = "on" then
		Response.Write "<img src=""https://www.google.com/chart?chs=82x82&cht=qr&chl="
	else
		Response.Write "<img src=""http://chart.apis.google.com/chart?chs=82x82&cht=qr&chl="
	end if
	Response.Write Server.URLEncode("http://m.shigotonavi.jp/detail/JobDetail.asp?OrderCode=" & rRS.Collect("OrderCode") & "&an=QROD")
	Response.Write """ border=""0"" align=""right"" alt=""QRCode"">"
%>
		</div>
		<div style="text-align:right; font-size:11px; padding-top:6px;">
			<a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&amp;detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=470');return false;"><img src="<%= HTTP_NAVI_CURRENTURL %>img/staff/mail/mailhei.gif" border="0" align="bottom" alt="求人票をメール送信"> 求人票をメール送信</a>
		</div>
		<p class="m0" style="text-align:right;padding:4px 0px 0px 0px;">掲載期限：<%= sPublishLimitStr %></p>

		<%= sCautionStr %>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
		ElseIf flgNowPublic = True Then
			'しごとナビに非ログインの場合は、企業名＋掲載期限＋求人票ＵＲＬメール送信
%>
	<div style="width:400px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="width:200px; float:left;">

<%
	if Request.ServerVariables("HTTPS") = "on" then
		Response.Write "<img src=""https://www.google.com/chart?chs=82x82&cht=qr&chl="
	else
		Response.Write "<img src=""http://chart.apis.google.com/chart?chs=82x82&cht=qr&chl="
	end if
	Response.Write Server.URLEncode("http://m.shigotonavi.jp/detail/JobDetail.asp?OrderCode=" & rRS.Collect("OrderCode") & "&an=QROD")
	Response.Write """ border=""0"" align=""right"" alt=""QRCode"">"
%>
		<div style="text-align:right; font-size:11px; padding-top:6px;"><a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&amp;detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=380');return false;"><img src="<%= HTTP_NAVI_CURRENTURL %>img/staff/mail/mailhei.gif" border="0" align="bottom" alt="求人票をメール送信"> 求人票をメール送信</a></div>
		<p class="m0" style="text-align:right;padding:4px 0px 0px 0px;">掲載期限：<%= sPublishLimitStr %></p>
		<%= sCautionStr %>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
		Else
%>
	<div style="width:400px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="width:200px; float:left;">
		<p class="m0" style="text-align:right; padding-top:21px;">掲載期限：<%= sPublishLimitStr %></p>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
		End If
	Else
%>
	<div style="width:400px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="width:200px; float:left;">

<%
	if Request.ServerVariables("HTTPS") = "on" then
		Response.Write "<img src=""https://www.google.com/chart?chs=82x82&cht=qr&chl="
	else
		Response.Write "<img src=""http://chart.apis.google.com/chart?chs=82x82&cht=qr&chl="
	end if
	Response.Write Server.URLEncode("http://m.shigotonavi.jp/detail/JobDetail.asp?OrderCode=" & rRS.Collect("OrderCode") & "&an=QROD")
	Response.Write """ border=""0"" align=""right"" alt=""QRCode"">"
%>		<p class="m0" style="text-align:right; width:156px; padding-top:21px;">掲載期限：<%= sPublishLimitStr %></p>

		<%= sCautionStr %>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
	End If
%>
</div>-->


<%
End Function

'******************************************************************************
'概　要：求人票詳細ページの会社情報・職種情報・インタビュー切り替えボタンと参照回数を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vType			：表示中情報の種類 ["0"]職種情報 ["1"]会社情報 ["2"]インタビュー
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
	Dim dbTopInterviewFlag
	Dim dbPlanType

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sOrderType = rRS.Collect("OrderType")
	dbPlanType = ChkStr(rRS.Collect("PlanTypeName"))
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'具体的職種名
	sJobTypeDetail = rRS.Collect("JobTypeDetail")
	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")
	'トップインタビュー
	dbTopInterviewFlag = rRS.Collect("TopInterviewFlag")

	If sJobTypeDetail <> "" Then sJobTypeDetail = sJobTypeDetail & "のお仕事情報詳細"

	Response.Write "<div id=""tab_switch"">"
	Response.Write "<div class=""left"">"
	If vType = "0" Then
		'仕事情報を表示中の場合
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_orderdetail_on.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></div>"
		If sOrderType = "0" Then
			'一般の求人広告の場合は会社情報へのリンクを表示
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></a></div>"
		End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_interview_off.gif"" alt=""インタビュー"" border=""0"" width=""93"" height=""22""></a></div>"
		End If
	ElseIf vType = "1" Then
		'会社情報を表示中の場合
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		If sOrderType = "0" Then
			'一般の求人広告の場合は会社情報を表示
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_companyinfo_on.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></div>"
		End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_interview_off.gif"" alt=""インタビュー"" border=""0"" width=""93"" height=""22""></a></div>"
		End If

	ElseIf vType = "2" Then
		'インタビューを表示中の場合
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_interview_on.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></div>"
	End If

	Response.Write "</div>"


	Response.Write "<br clear=""both""></div>" & vbCrLf
End Function

'******************************************************************************
'概　要：求人票のキャッチコピー部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'　　　：2010/05/06 LIS K.Kokubo 社内案件用写真
'******************************************************************************
Function DspOrderCatchCopy(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim vAccessCount
	Dim sPublishLimitStr
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'大きい画像 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'大きい画像 end
	'******************************************************************************

	'<社内案件用写真>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</社内案件用写真>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)

	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics""><div>"
		Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"
		Response.Write "</div></div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div>"

		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write "</div>"

		%>
       		<div id="lissapo">
			<div><span>転職サポート案件</span><br>
			人材会社の転職支援を受けて応募できる求人
			</div>
			<p>掲載期限：<%= sPublishLimitStr %><br>
			月間閲覧回数：<%= vAccessCount %>回<br>
			更新日：<%= sUpdateDay %></p>
			<span>期限前に掲載終了する場合があります。</span>
			</div>
           <br clear="all">
           
           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>
            <% End If %>
            
		<%	

		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
	Else
		Response.Write "<div id=""catchcopy2"">"
		Response.Write "<div id=""in_catch"">"		
		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"


		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write"</div>"
		
			%>
       		<div id="lissapo">
			<div><span>転職サポート案件</span><br>
			人材会社の転職支援を受けて応募できる求人
			</div>
			<p>掲載期限：<%= sPublishLimitStr %><br>
			月間閲覧回数：<%= vAccessCount %>回<br>
			更新日：<%= sUpdateDay %></p>
			<span>期限前に掲載終了する場合があります。</span>
			</div>
           <br clear="all">
           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>
            <% End If %>

		<%
		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
			
		  If G_USERTYPE = "" Then  %>
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>

		<% End If 

	End If
End Function


'******************************************************************************
'概　要：求人票のキャッチコピー部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：2

'******************************************************************************
Function DspOrderCatchCopy2(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode
    Dim dbCompanyName

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

        if dbOrderType = "0" Then
            dbCompanyName = rRS.Collect("CompanyName")
        else
            dbCompanyName = ""
        End if
        %>
        <div id="c_name"><%= dbCompanyName %></div>
        <%
	'******************************************************************************
	'大きい画像 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'大きい画像 end
	'******************************************************************************

	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">※期限前に掲載終了する場合があります。</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならDspPublicLimitDayを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
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

	'<社内案件用写真>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</社内案件用写真>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics""><div>"
		Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"
		Response.Write "</div></div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div>"

		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write "</div>"

		%>
        	<div id="lissapo">
			<div><span>直接ご応募案件</span><br>
			このページから企業へ直接ご応募できる求人 
			</div>
			<p>掲載期限：<%= sPublishLimitStr %><br>
			月間閲覧回数：<%= vAccessCount %>回<br>
			更新日：<%= sUpdateDay %></p>
			<span>期限前に掲載終了する場合があります。</span>
			</div>
           <br clear="all">
           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>
            <% End If %> 
		<%	

		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
	Else
		Response.Write "<div id=""catchcopy2"">"
		Response.Write "<div id=""in_catch"">"		
		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"


		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write"</div>"
		
			%>

        	<div id="lissapo">
			<div><span>直接ご応募案件</span><br>
			このページから企業へ直接ご応募できる求人 
			</div>
			<p>掲載期限：<%= sPublishLimitStr %><br>
			月間閲覧回数：<%= vAccessCount %>回<br>
			更新日：<%= sUpdateDay %></p>
			<span>期限前に掲載終了する場合があります。</span>
			</div>
           
 

		<%
		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
			
		  If G_USERTYPE = "" Then  %>
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>

		<% End If 

	End If
End Function

'******************************************************************************
'概　要：求人票のキャッチコピー部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：リスサポート案件用

'******************************************************************************
Function DspOrderCatchCopy3(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic
	Dim dbCompanyName 'リス自体の求人に使う会社名

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

        if dbCompanyCode = "C0001533" Then
            dbCompanyName = rRS.Collect("CompanyName")
            %>
            <div id="c_name"><%= dbCompanyName %></div>
            <%
        else
            dbCompanyName = ""
        End if
	'******************************************************************************
	'大きい画像 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'大きい画像 end
	'******************************************************************************

	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">※期限前に掲載終了する場合があります。</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならDspPublicLimitDayを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
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

	'<社内案件用写真>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</社内案件用写真>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics""><div>"
		Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"
		Response.Write "</div></div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div>"

		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write "</div>"

		%>
			<div id="lissapo">
			<div><span>転職サポート案件</span><br>
			人材会社の転職支援を受けて応募できる求人
			</div>
			<p>掲載期限：<%= sPublishLimitStr %><br>
			月間閲覧回数：<%= vAccessCount %>回<br>
			更新日：<%= sUpdateDay %></p>
			<span>期限前に掲載終了する場合があります。</span>
			</div>
           <br clear="all">
           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>
            <% End If %>
		

		<br clear="all">
		</div>
       
    <%
	Else
		Response.Write "<div id=""catchcopy2"">"
		Response.Write "<div id=""in_catch"">"		
		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"


		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write"</div>"
		
			%>

        	<div id="lissapo">
			<div><span>転職サポート案件</span><br>
			人材会社の転職支援を受けて応募できる求人
			</div>
			<p>掲載期限：<%= sPublishLimitStr %><br>
			月間閲覧回数：<%= vAccessCount %>回<br>
			更新日：<%= sUpdateDay %></p>
			<span>期限前に掲載終了する場合があります。</span>
			</div>

		<%
		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
			
		  If G_USERTYPE = "" Then  %>
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>
			
		<% End If 

	End If
End Function

'******************************************************************************
'概　要：求人票のキャッチコピー部分を出力(過去求人)
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：リスサポート案件用

'******************************************************************************
Function DspOrderCatchCopy_OldPlan(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount,ByVal YearlyIncomeMin,ByVal MonthlyIncomeMin,ByVal DailyIncomeMin,ByVal HourlyIncomeMin)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic

	Dim JobTypeBigCode
	Dim JobTypeCode
	Dim WorkingTypeCode1
	Dim WorkingTypeCode2
	Dim WorkingTypeCode3
	Dim PrefectureCode
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'大きい画像 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'大きい画像 end
	'******************************************************************************

	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">※期限前に掲載終了する場合があります。</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならDspPublicLimitDayを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
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

	'<社内案件用写真>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</社内案件用写真>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics""><div>"
		Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"
		Response.Write "</div></div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div>"

		If sImgSpeciality <> "" Then

			
			%>			

<div class="right">
	<img src="/img/order/oubo_end.png" class="spSmart">
</div>
 <br clear="both">  

		<div class="center" style="margin-top:25px;">
                    <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>">
           		<img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_reg_button03.png" alt="会員登録をして条件に近い求人へ応募" border="0">
            </a>
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_list.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>">
           		<img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_login_button03.png" alt="ログインして条件に近い求人へ応募" border="0">
            </a>
      
            
		</div>
  
<% 
			
		End If

		Response.Write "</div>"


		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
	Else
		Response.Write "<div id=""catchcopy3"" class=""left"">"
	
		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"



		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
			
%>			

<div class="right">
	<img src="/img/order/oubo_end.png" class="spSmart">
</div>
 <br clear="both">   
<%    		
		If G_USERTYPE = "" Then
			'雇用形態、勤務地、職種検索
			sSQL = "select CJT.jobtypecode, BJT.Bigclasscode from C_JobType AS CJT INNER JOIN B_JobType AS BJT ON CJT.JobTypeCode = BJT.AllConnectCode where CJT.id = '1' and CJT.OrderCode = '" & dbOrderCode & "';"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("Bigclasscode")) <> "" Then
					JobTypeBigCode = oRS.Collect("Bigclasscode")
				End If
				If ChkStr(oRS.Collect("jobtypecode")) <> "" Then
					JobTypeCode = oRS.Collect("jobtypecode")
				End If
			End If
			Call RSClose(oRS)

			sSQL = "select prefecturecode from c_workingplace where ordercode = '" & dbOrderCode & "';"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			PrefectureCode  = ""
			Do While GetRSState(oRS) = True
				If ChkStr(oRS.Collect("prefecturecode")) <> "" Then
					PrefectureCode = PrefectureCode & oRS.Collect("prefecturecode") & ","
				End If
				oRS.MoveNext
			Loop
			Call RSClose(oRS)
			PrefectureCode = Left(PrefectureCode, Len(PrefectureCode) -1)

			sSQL = "select CWT1.workingtypecode as workingtypecode1,CWT2.workingtypecode as workingtypecode2,CWT3.workingtypecode as workingtypecode3 from c_workingtype AS CWT1 "
			sSQL = sSQL & " left join c_workingtype AS CWT2 on CWT1.ordercode = '" & dbOrderCode & "' and CWT2.id = 2"
			sSQL = sSQL & " left join c_workingtype AS CWT3 on CWT2.ordercode = '" & dbOrderCode & "' and CWT3.id = 3"
			sSQL = sSQL & " where CWT3.ordercode = '" & dbOrderCode & "' and CWT1.id = 1;"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("workingtypecode1")) <> "" Then
					WorkingTypeCode1 = oRS.Collect("workingtypecode1")
				End If
				If ChkStr(oRS.Collect("workingtypecode2")) <> "" Then
					WorkingTypeCode2 = oRS.Collect("workingtypecode2")
				End If
				If ChkStr(oRS.Collect("workingtypecode3")) <> "" Then
					WorkingTypeCode3 = oRS.Collect("workingtypecode3")
				End If
			End If
			Call RSClose(oRS)
%>

			
		<div class="center">
		<a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>">
		<img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_reg_button03.png" alt="会員登録をして条件に近い求人へ応募" border="0">
		</a>
		<a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_list.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>">
		<img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_login_button03.png" alt="ログインして条件に近い求人へ応募" border="0">
		</a>
		</div>
<% End If 

	End If
	

End Function

'******************************************************************************
'概　要：求人票詳細ページのフリーＰＲを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'******************************************************************************
Function DspOrderFreePR(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sPRTitle1			'ＰＲタイトル1
	Dim sPRTitle2			'ＰＲタイトル2
	Dim sPRTitle3			'ＰＲタイトル3
	Dim sPRContents1		'ＰＲ文1
	Dim sPRContents2		'ＰＲ文2
	Dim sPRContents3		'ＰＲ文3
	Dim flgPR				'ＰＲ有無フラグ [True]有 [False]無

	Dim dbOrderCode

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

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
    	<img src="/img/order/tab_detail_pr.png" class="tab_img">
		<table class="detail_table">
        <tbody>
        <tr>
        <th class="fast_th"></th>
        <td>
        <%
		Response.Write "<div>"
		If sPRTitle1 <> "" Or sPRContents1 <> "" Then
			Response.Write "<h4>" & sPRTitle1 & "</h4>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<p class=""m0"">" & sPRContents1 & "</p>"
		End If

		If sPRTitle2 <> "" Or sPRContents2 <> "" Then
			Response.Write "<h4>" & sPRTitle2 & "</h4>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<p class=""m0"">" & sPRContents2 & "</p>"
		End If

		If sPRTitle3 <> "" Or sPRContents3 <> "" Then
			Response.Write "<h4>" & sPRTitle3 & "</h4>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<p class=""m0"">" & sPRContents3 & "</p>"
		End If
		Response.Write "</div>"
		%>
        </td>
        </tr>
        </tbody>
        </table>
        
        <div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div>  
        
     <%   
	End If
End Function

'******************************************************************************
'概　要：求人企業画像一覧表示ＨＴＭＬ表示
'引　数：rDB			：接続中ＤＢオブジェクト
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vCategoryCode	：カテゴリコード
'使　用：ナビ/order/order_detail.asp
'備　考：
'履　歴：2006/12/27 LIS K.Kokubo 作成
'　　　：2008/01/28 LIS K.Kokubo ライセンス変更による対応
'　　　：2010/05/06 LIS K.Kokubo 社内案件用写真
'******************************************************************************
Function DspOrderPictureNow(ByRef rDB, ByRef rRS, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbCompanyCode
	Dim dbOrderType
	Dim dbImageLimit

	Dim sURL
	Dim sImg1,sImg2,sImg3,sCap1,sCap2,sCap3

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbCompanyCode = rRS.Collect("CompanyCode")
	dbOrderType = rRS.Collect("OrderType")
	dbImageLimit = rRS.Collect("ImageLimit")

	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo2")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo2")
				sCap1 = ChkStr(oRS.Collect("Caption2"))
			End If
			If ChkStr(oRS.Collect("PicNo3")) <> "" Then
				sImg2 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo3")
				sCap2 = ChkStr(oRS.Collect("Caption3"))
			End If
			If ChkStr(oRS.Collect("PicNo4")) <> "" Then
				sImg3 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo4")
				sCap3 = ChkStr(oRS.Collect("Caption4"))
			End If
		End If
		Call RSClose(oRS)
	ElseIf dbImageLimit > 1 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', '" & vCategoryCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("OptionNo2")) <> "" Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
				sCap1 = ChkStr(oRS.Collect("Caption2"))
			End If
			If ChkStr(oRS.Collect("OptionNo3")) <> "" Then
				sImg2 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
				sCap2 = ChkStr(oRS.Collect("Caption3"))
			End If
			If ChkStr(oRS.Collect("OptionNo4")) <> "" Then
				sImg3 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
				sCap3 = ChkStr(oRS.Collect("Caption4"))
			End If
		End If
	End If

	If sImg1 & sImg2 & sImg3 <> "" Then
		Response.Write "<div id=""sub_pics"">"
		Response.Write "<div class=""auto"">"

		If sImg1 <> "" Then
			Response.Write "<div class=""sub_waku"">"
			Response.Write "<div class=""sub_pics sub_pics1""><div><img src=""" & sImg1 & """ alt=""" & sCap1 & """></div></div>"
			Response.Write "<p class=""m0"" align=""left"" style=""width:213px; font-size:10px;"">" & sCap1 & "</p>"
			Response.Write "</div>"
		End If

		If sImg2 <> "" Then
			Response.Write "<div class=""sub_waku"">"
			Response.Write "<div class=""sub_pics sub_pics2""><div><img src=""" & sImg2 & """ alt=""" & sCap2 & """></div></div>"
			Response.Write "<p class=""m0"" align=""left"" style=""width:213px; font-size:10px;"">" & sCap2 & "</p>"
			Response.Write "</div>"
		End If

		If sImg3 <> "" Then
			Response.Write "<div class=""sub_waku"">"
			Response.Write "<div class=""sub_pics sub_pics3""><div><img src=""" & sImg3 & """ alt=""" & sCap3 & """></div></div>"
			Response.Write "<p class=""m0"" align=""left"" style=""width:213px; font-size:10px;"">" & sCap3 & "</p>"
			Response.Write "</div>"
		End If


		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
		Response.Write "</div>"
	End If
End Function

'******************************************************************************
'概　要：求人票の採用の背景を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：
'履　歴：2007/05/13 LIS K.Kokubo 作成
'******************************************************************************
Function DspOrderBackGround(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderBackGround	'採用の背景

	DspOrderBackGround = False

	If GetRSState(rRS) = False Then Exit Function

	'採用の背景取得
	dbOrderBackGround = Replace(ChkStr(rRS.Collect("OrderBackGround")), vbCrLf, "<br>")

	'採用の背景出力
	If dbOrderBackGround <> "" Then
	%>
    <img src="/img/order/tab_detail_bb.png" class="tab_img">
	<table class="detail_table">
    <tbody>
    <tr>
    <th class="fast_th"></th>
    <td><p class="m0"><%= dbOrderBackGround %></p></td>
    </tr>
    </tbody>
    </table>
    
	<div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div>  
		
        <%
        DspOrderBackGround = True
	End If


End Function

'******************************************************************************
'概　要：求人票の業務内容を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'******************************************************************************
Function DspBusiness(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode			'情報コード
	Dim sCompanyCode		'企業コード
	Dim sPlanType			'求人票ライセンスプラン種類
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
	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")
	sPlanType = rRS.Collect("PlanTypeName")
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

	flgLine = False
	If flgBusiness = True Then
			If sBusinessDetail <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
	%>
    <img src="/img/order/tab_detail_job.png" class="tab_img">
   	<table class="detail_table">
    <tbody>
    <tr>
    <th rowspan="2" class="fast_th"></th>
    <td><h4>担当業務</h4>
    <p class=""m0""><%= sBusinessDetail %></p></td>
    </tr>
    <tr>
    <td><%
	 If (sPlanType = "platinum" Or sPlanType = "gold" Or sPlanType = "old") And sBiz <> "" Then


			Response.Write "<h4>仕事の割合</h4>"
			Response.Write "<div class=""value1"">"
			Response.Write "<table border=""0"">"
			Response.Write "<tbody>"
			Response.Write "<tr>"
			Response.Write "<td>"
			Response.Write "<script type=""text/javascript"" language=""javascript"">"
			Response.Write "viewWorkAvg(" & sBizPercentage1 & ", " & sBizPercentage2 & ", " & sBizPercentage3 & ", " & sBizPercentage4 & ")"
			Response.Write "</script>"
			Response.Write "</td>"
			Response.Write "<td style=""padding-left:5px; vertical-align:middle;"">"
			Response.Write "<table border=""0"">"
			Response.Write "<tbody>"
			If sBizName1 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#ff9999; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage1 & "%</td><td>" & sBizName1 & "</td></tr>"
			If sBizName2 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#9999ff; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage2 & "%</td><td>" & sBizName2 & "</td></tr>"
			If sBizName3 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#99ff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage3 & "%</td><td>" & sBizName3 & "</td></tr>"
			If sBizName4 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#ffff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage4 & "%</td><td>" & sBizName4 & "</td></tr>"
			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "</div>"
		End If
		 %>
         </td>
    </tr>
    </tbody>
    </table>
    <div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div>  
    
    <%
  		End If


	End If
End Function

'******************************************************************************
'概　要：求人票の勤務条件を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/include/func_order.asp
'備　考：
'履　歴：2007/02/11 作成
'　　　：2008/10/22 LIS K.Kokubo 勤務地複数化対応
'　　　：2009/04/16 LIS K.Kokubo メール課金ライセンスの場合は勤務地の表示を一般の求人広告でも市区郡までしか表示させない
'　　　：2009/04/22 LIS K.Kokubo 紹介後の勤務形態(TTP用)対応
'　　　：2009/11/02 LIS K.Kokubo ＳＯＨＯ,ＦＣの勤務地表示対応
'******************************************************************************
Function DspCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	'<変数宣言>
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim oRS3
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'情報コード
	Dim dbOrderType			'求人票種類
	Dim dbCompanyKbn		'企業区分
	Dim dbJobTypeDetail		'職種詳細
	Dim dbYearlyIncomeMin	'年収下限
	Dim dbYearlyIncomeMax	'年収上限
	Dim dbMonthlyIncomeMin	'月給下限
	Dim dbMonthlyIncomeMax	'月給上限
	Dim dbDailyIncomeMin	'日給下限
	Dim dbDailyIncomeMax	'日給上限
	Dim dbHourlyIncomeMin	'時給下限
	Dim dbHourlyIncomeMax	'時給上限
	Dim dbPercentagePay		'歩合制
	Dim dbSalaryRemark		'給与備考
	Dim dbTrafficFeeType	'
	Dim dbTrafficFeeMonth	'交通費／１ヶ月
	Dim dbAfterWorkingTypeCode'紹介後の勤務形態
	Dim dbWorkStartDay		'就業開始日
	Dim dbWorkEndDay		'就業終了日
	Dim dbWorkTimeRemark	'就業時間備考
	Dim dbWeeklyHolidayType	'週休
	Dim dbHolidayRemark		'休日備考
	Dim dbHumanNumber		'募集人数
	Dim dbWorkingPlaceSeq	'勤務地番号
	Dim dbWorkingPlacePrefectureName'勤務地都道府県名
	Dim dbWorkingPlaceCity	'勤務地市区郡
	Dim dbWorkingPlaceAddressAll'勤務地住所全体
	Dim dbWorkingPlaceSection'勤務地部署
	Dim dbWorkingPlaceTelephoneNumber'勤務地TEL
	Dim dbMapFlag			'地図有無フラグ
	Dim dbTransfer			'転勤
	Dim dbPlanTypeName		'ナビライセンスの種類
	Dim dbTTPOrderFlag		'紹介予定派遣案件フラグ

	Dim sHTML
	Dim sWorkingType		'勤務形態
	Dim sJobType			'職種
	Dim sSalary				'給与
	Dim sYearlyIncome		'年収
	Dim sMonthlyIncome		'月給
	Dim sDailyIncome		'日給
	Dim sHourlyIncome		'時給
	Dim sTrafficFee			'交通費
	Dim sAfterWorkingType	'紹介後の勤務形態
	Dim sWorkRange			'就業期間
	Dim sWorkUpdate			'就業期間の更新有無
	Dim sWorkingTime		'就業時間
	Dim sMAP				'地図情報
	Dim sWorkingPlace		'就業場所
	Dim sNearbyStation		'最寄駅
	Dim sNearbyRailway		'沿線
	Dim sNearbyStationBlock	'最寄駅,沿線ブロック
	Dim iMaxRow
	Dim sDisplay
	Dim flgDspWorkingType
	Dim flgDspJobType
	Dim flgDspSalary
	Dim flgDspTime
	Dim flgDspHoliday
	Dim flgDspHumanNumber
	Dim flgDspWorkingPlace
	Dim flgLine
	Dim flgSOHOFC
	'</変数宣言>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbPlanTypeName = rRS.Collect("PlanTypeName")
	dbTTPOrderFlag = rRS.Collect("TTPOrderFlag")
	
		'<勤務形態>
	flgDspWorkingType = False
	dbAfterWorkingTypeCode = ChkStr(rRS.Collect("AfterWorkingTypeCode"))
	dbWorkStartDay = ChkStr(rRS.Collect("WorkStartDay"))
	dbWorkEndDay = ChkStr(rRS.Collect("WorkEndDay"))

	'勤務形態
	sWorkingType = GetWorkingType(rDB, rRS)
	flgSOHOFC = False
	If IsRE(sWorkingType,"((SOHO)|(FC))",True) = True Then flgSOHOFC = True

	'紹介後の勤務形態
	sAfterWorkingType = ""
	If dbAfterWorkingTypeCode <> "" Then
		sAfterWorkingType = "※紹介後の勤務形態&nbsp;･･･&nbsp;" & GetDetail("WorkingType", dbAfterWorkingTypeCode)
	End If

	'就業期間
	sWorkRange = ""
	If dbWorkStartDay & dbWorkEndDay <> "" Then
		If dbWorkStartDay <> "" Then sWorkRange = sWorkRange & GetDateStr(dbWorkStartDay, "/")
		If sWorkRange <> "" Then sWorkRange = sWorkRange & "～"
		If dbWorkEndDay <> "" Then sWorkRange = sWorkRange & GetDateStr(dbWorkEndDay, "/")
	End If

	If dbOrderType = "1" Then
		If rRS.Collect("WorkUpdateFlag") = "1" Then
			sWorkUpdate = "有"
		Else
			sWorkUpdate = "無"
		End If
		sWorkRange = sWorkRange & "(更新" & sWorkUpdate & ")"
	End If

	If sWorkingType & sAfterWorkingType & sWorkRange <> "" Then flgDspWorkingType = True
	'</勤務形態>

	'<職種>
	flgDspJobType = False
	sJobType = GetJobType(rDB, rRS)
	dbJobTypeDetail = rRS.Collect("JobTypeDetail")
	If sJobType & dbJobTypeDetail <> "" Then flgDspJobType = True
	'</職種>

	'<給与>
	flgDspSalary = False
	dbYearlyIncomeMin = ChkStr(rRS.Collect("YearlyIncomeMin"))
	dbYearlyIncomeMax = ChkStr(rRS.Collect("YearlyIncomeMax"))
	If dbYearlyIncomeMin = "0" Then dbYearlyIncomeMin = ""
	If dbYearlyIncomeMax = "0" Then dbYearlyIncomeMax = ""
	If dbYearlyIncomeMin <> "" Then dbYearlyIncomeMin = GetJapaneseYen(dbYearlyIncomeMin)
	If dbYearlyIncomeMax <> "" Then dbYearlyIncomeMax = GetJapaneseYen(dbYearlyIncomeMax)
	If dbYearlyIncomeMin & dbYearlyIncomeMax <> "" Then
		If dbYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMin
		sYearlyIncome = sYearlyIncome & "&nbsp;～&nbsp;"
		If dbYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMax
	End If

	dbMonthlyIncomeMin = ChkStr(rRS.Collect("MonthlyIncomeMin"))
	dbMonthlyIncomeMax = ChkStr(rRS.Collect("MonthlyIncomeMax"))
	If dbMonthlyIncomeMin = "0" Then dbMonthlyIncomeMin = ""
	If dbMonthlyIncomeMax = "0" Then dbMonthlyIncomeMax = ""
	If dbMonthlyIncomeMin <> "" Then dbMonthlyIncomeMin = GetJapaneseYen(dbMonthlyIncomeMin)
	If dbMonthlyIncomeMax <> "" Then dbMonthlyIncomeMax = GetJapaneseYen(dbMonthlyIncomeMax)
	If dbMonthlyIncomeMin & dbMonthlyIncomeMax <> "" Then
		If dbMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & dbMonthlyIncomeMin
		sMonthlyIncome = sMonthlyIncome & "&nbsp;～&nbsp;"
		If dbMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & dbMonthlyIncomeMax
	End If

	dbDailyIncomeMin = ChkStr(rRS.Collect("DailyIncomeMin"))
	dbDailyIncomeMax = ChkStr(rRS.Collect("DailyIncomeMax"))
	If dbDailyIncomeMin = "0" Then dbDailyIncomeMin = ""
	If dbDailyIncomeMax = "0" Then dbDailyIncomeMax = ""
	If dbDailyIncomeMin <> "" Then dbDailyIncomeMin = GetJapaneseYen(dbDailyIncomeMin)
	If dbDailyIncomeMax <> "" Then dbDailyIncomeMax = GetJapaneseYen(dbDailyIncomeMax)
	If dbDailyIncomeMin & dbDailyIncomeMax <> "" Then
		If dbDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & dbDailyIncomeMin
		sDailyIncome = sDailyIncome & "&nbsp;～&nbsp;"
		If dbDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & dbDailyIncomeMax
	End If

	dbHourlyIncomeMin = ChkStr(rRS.Collect("HourlyIncomeMin"))
	dbHourlyIncomeMax = ChkStr(rRS.Collect("HourlyIncomeMax"))
	If dbHourlyIncomeMin = "0" Then dbHourlyIncomeMin = ""
	If dbHourlyIncomeMax = "0" Then dbHourlyIncomeMax = ""
	If dbHourlyIncomeMin <> "" Then dbHourlyIncomeMin = GetJapaneseYen(dbHourlyIncomeMin)
	If dbHourlyIncomeMax <> "" Then dbHourlyIncomeMax = GetJapaneseYen(dbHourlyIncomeMax)
	If dbHourlyIncomeMin & dbHourlyIncomeMax <> "" Then
		If dbHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMin
		sHourlyIncome = sHourlyIncome & "&nbsp;～&nbsp;"
		If dbHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMax
	End If

	dbPercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
	dbSalaryRemark = Replace(ChkStr(rRS.Collect("IncomeRemark")), vbCrLf, "<br>")
	dbSalaryRemark = Replace(dbSalaryRemark, vbCr, "<br>")
	dbSalaryRemark = Replace(dbSalaryRemark, vbLf, "<br>")
	sTrafficFee = ""
	dbTrafficFeeType = ChkStr(rRS.Collect("TrafficFeeType"))
	dbTrafficFeeMonth = ChkStr(rRS.Collect("MonthTrafficFee"))

	'歩合制
	If dbPercentagePay <> "" Then
		If dbPercentagePay = "1" Then
			dbPercentagePay = "あり"
		ElseIf dbPercentagePay = "0" Then
			dbPercentagePay = "なし"
		End If
	End If

	'交通費
	If ChkStr(rRS.Collect("NaviTrafficPayFlag")) = "1" Then 
		sTrafficFee = "交通費支給あり" & dbTrafficFeeType
		If IsNumber(dbTrafficFeeMonth, 0, False) = True Then
			sTrafficFee = sTrafficFee & "(" & FormatCanma(dbTrafficFeeMonth) & "円／月)"
		End If
	End If

	If sYearlyIncome & sMonthlyIncome & sDailyIncome & sHourlyIncome & dbPercentagePay & sTrafficFee & dbSalaryRemark <> "" Then flgDspSalary = True
	'</給与>

	'<時間>
	flgDspTime = False
	sWorkingTime = GetWorkingTime(rDB, rRS)
	dbWorkTimeRemark = ChkStr(rRS.Collect("WorkTimeRemark"))
	dbWorkTimeRemark = Replace(ChkStr(rRS.Collect("WorkTimeRemark")), vbCrLf, "<br>")
	dbWorkTimeRemark = Replace(dbWorkTimeRemark, vbCr, "<br>")
	dbWorkTimeRemark = Replace(dbWorkTimeRemark, vbLf, "<br>")

	If sWorkingTime & dbWorkTimeRemark <> "" Then flgDspTime = True
	'</時間>

	'<休日>
	flgDspHoliday = False
	dbWeeklyHolidayType = ChkStr(rRS.Collect("WeeklyHolidayTypeName"))
	dbHolidayRemark = ChkStr(rRS.Collect("HolidayRemark"))
	dbHolidayRemark = Replace(ChkStr(rRS.Collect("HolidayRemark")), vbCrLf, "<br>")
	dbHolidayRemark = Replace(dbHolidayRemark, vbCr, "<br>")
	dbHolidayRemark = Replace(dbHolidayRemark, vbLf, "<br>")

	If dbWeeklyHolidayType & dbHolidayRemark <> "" Then flgDspHoliday = True
	'</休日>

	'<募集人数>
	flgDspHumanNumber = False
	dbHumanNumber = ChkStr(rRS.Collect("HumanNumber"))

	If dbHumanNumber <> "" Then
		dbHumanNumber = dbHumanNumber & "人"
	End If

	If dbHumanNumber <> "" Then flgDspHumanNumber = True
	'</募集人数>

	'<勤務地>
	flgDspWorkingPlace = False

	iMaxRow = 0
	sWorkingPlace = ""
	sNearbyStationBlock = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		iMaxRow = oRS.RecordCount
		'<最寄駅>
		sSQL = "EXEC up_LstC_NearbyStation '" & dbOrderCode & "', '';"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then Set oRS2.ActiveConnection = Nothing
		'</最寄駅>
		'<最寄沿線>
		sSQL = "EXEC up_LstC_NearbyRailwayLine '" & rRS.Collect("OrderCode") & "','','';"
		flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then Set oRS3.ActiveConnection = Nothing
		'</最寄沿線>
	End If
	Do While GetRSState(oRS) = True
		dbWorkingPlaceSeq = ChkStr(oRS.Collect("WorkingPlaceSeq"))
		dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS.Collect("WorkingPlaceCity"))
		dbWorkingPlaceAddressAll = ChkStr(oRS.Collect("WorkingPlaceAddressAll"))
		dbWorkingPlaceSection = ChkStr(oRS.Collect("WorkingPlaceSection"))
		dbWorkingPlaceTelephoneNumber = ChkStr(oRS.Collect("WorkingPlaceTelephoneNumber"))
		dbMapFlag = ChkStr(oRS.Collect("MapFlag"))

		If sWorkingPlace <> "" And flgSOHOFC = True Then sWorkingPlace = sWorkingPlace & "、"

		'<勤務地>
		sWorkingPlace = sWorkingPlace & "<div"
		If flgSOHOFC = True Then sWorkingPlace = sWorkingPlace & " style=""display:inline;"""
		sWorkingPlace = sWorkingPlace & ">"
		If iMaxRow > 1 And flgSOHOFC = False Then sWorkingPlace = sWorkingPlace & "【勤務地" & dbWorkingPlaceSeq & "】"

		If dbOrderType <> "0" Then
			sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		ElseIf dbPlanTypeName = "mail" Then
			sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		Else
			sWorkingPlace = sWorkingPlace & dbWorkingPlaceAddressAll
			If dbWorkingPlaceSection & dbWorkingPlaceTelephoneNumber <> "" Then
				sWorkingPlace = sWorkingPlace & "("
				If dbWorkingPlaceSection <> "" Then sWorkingPlace = sWorkingPlace & dbWorkingPlaceSection
				If dbWorkingPlaceSection <> "" And dbWorkingPlaceTelephoneNumber <> "" Then sWorkingPlace = sWorkingPlace 
				If dbWorkingPlaceTelephoneNumber <> "" Then sWorkingPlace = sWorkingPlace '& "TEL:" & dbWorkingPlaceTelephoneNumber
				sWorkingPlace = sWorkingPlace & ")"
			End If
			If dbMapFlag = "1" Then sWorkingPlace = sWorkingPlace & "&nbsp;[<span style=""color:#0045f9;cursor:pointer;"" onclick=""open('" & HTTPS_CURRENTURL & "map/showmap.asp?ordercode=" & dbOrderCode & "&wpseq=" & dbWorkingPlaceSeq & "', 'map', 'width=700,height=650');"">地図</span>]"
		End If

		'<最寄駅>
		sNearbyStation = ""
		oRS2.Filter = "WorkingPlaceSeq = " & dbWorkingPlaceSeq
		If GetRSState(oRS2) = True Then
			sNearbyStation = GetNearbyStation(rDB, oRS2)
		End If
		oRS2.Filter = 0
		'</最寄駅>
		'<最寄沿線>
		sNearbyRailway = ""
		oRS3.Filter = "WorkingPlaceSeq = " & dbWorkingPlaceSeq
		If GetRSState(oRS3) = True Then
            'sNearbyRailway = GetNearbyRailway(rDB, oRS3)
			sNearbyRailway = GetNearbyRailway2(rDB, oRS3)
		End If
		oRS3.Filter = 0
		'</最寄沿線>
		If sNearbyStation <> "" Then
			sWorkingPlace = sWorkingPlace & "<p class=""m0"""
			If flgSOHOFC = True Then
				sWorkingPlace = sWorkingPlace & " style=""display:inline;"""
			Else
				sWorkingPlace = sWorkingPlace & " style=""padding-left:15px;"""
			End If
			sWorkingPlace = sWorkingPlace & ">"
			sWorkingPlace = sWorkingPlace & "[最寄駅]"
			sWorkingPlace = sWorkingPlace & sNearbyStation
			If flgSOHOFC = False Then sWorkingPlace = sWorkingPlace & "<br>"
			sWorkingPlace = sWorkingPlace & "[沿線]"
			sWorkingPlace = sWorkingPlace & sNearbyRailway
			sWorkingPlace = sWorkingPlace & "</p>"
		End If
		'</勤務地>

		sWorkingPlace = sWorkingPlace & "</div>"
		oRS.MoveNext
	Loop

	'転勤
	If (dbOrderType = "0" Or dbOrderType = "2") And dbCompanyKbn <> "4" Then
		'ﾘｽの派遣求人票 または 派遣会社の求人票の場合は表示しない

		dbTransfer = ChkStr(rRS.Collect("Transfer"))
		If dbTransfer <> "" Then
			If dbTransfer = "有" Then
				dbTransfer = "転勤あり"
			ElseIf dbTransfer = "無" Then
				dbTransfer = "転勤なし"
			End If
		End If
	End If
	If sWorkingPlace & sNearbyStationBlock & dbTransfer <> "" Then flgDspWorkingPlace = True
	'</勤務地>

flgLine = False


%>	
	<img src="/img/order/tab_detail_li.png" class="tab_img">
	<table class="detail_table">
    <tbody>


<% If sWorkingType <> "" Then
	If flgLine = True Then sHTML = sHTML & ""
		flgLine = True
 %>
     <tr>
    <th class="dborder_bottom">
     勤務形態
    </th>
    <td class="dborder_bottom">
	<p class="m0 get_job_type"><%= sWorkingType %></p>
	<% End If %>

	<% If dbTTPOrderFlag = "1" And sAfterWorkingType <> "" Then %>
	<p class="m0"><%= sAfterWorkingType %></p>
	<% End If %>
    
    <% If sWorkRange <> "" Then %>
    <p class="m0">※有期の場合：<%= sWorkRange %></p>
	<% End If %>
   
	<% If sWorkingType <> "" Then %>
    </td>
    </tr>

	<% End If %>

	<%
		If flgDspJobType = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">職種</th>
    <td class="dborder_bottom">
    
    <p class="m0"><strong><%= dbJobTypeDetail %></strong></p>
	<p class="m0"><%= sJobType %></p>
    
    
    </td>
    </tr>
    <% End If %>
    
	<%
		If flgDspSalary = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">給与</th>
    <td class="dborder_bottom">
    
   <% If sYearlyIncome <> "" Then %>
	<h5>年収</h5>
	<p class="m0"><%= sYearlyIncome %></p>
    <% End If %>
    
    <% If sMonthlyIncome <> "" Then %>
	<h5>月給</h5>
	<p class="m0"><%= sMonthlyIncome %></p>
    <% End If %>
    
    <% If sDailyIncome <> "" Then %>
	<h5>日給</h5>
	<p class="m0"><%= sDailyIncome %></p>
    <% End If %>
    
    <% If sHourlyIncome <> "" Then %>
	<h5>時給</h5>
	<p class="m0"><%= sHourlyIncome %></p>
    <% End If %>
    
    <% If dbPercentagePay <> "" Then %>
	<h5>歩合制</h5>
	<p class="m0"><%= dbPercentagePay %></p>
    <% End If %>
    
    <% If sTrafficFee <> "" Then %>
	<h5>交通費</h5>
	<p class="m0"><%= sTrafficFee %></p>
    <% End If %>
    
    <% If dbSalaryRemark <> "" Then %>
	<h5>給与備考</h5>
	<p class="m0"><%= dbSalaryRemark %></p>
    <% End If %>
    
    <% If sYearlyIncome & sMonthlyIncome & sDailyIncome & sHourlyIncome <> "" AND dbOrderCode <> "J0074418" Then %>
	<p class="m0">※最低額は条件に関係なく得られる額です。<br>(年収の最低額は条件に関係なく得られる月給の合計です。)</p>
    <% End If %>
    <% 'リスのキャリコン募集だけログをとる
    If dbOrderCode = "J0074418" Then
        Call RegMailFromAccess(GetForm("mailfromaccess", 2))
    End If 
    %>
    </td>
    </tr>
    <% End If %>
    
    <%
		If flgDspTime = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">時間</th>
    <td class="dborder_bottom">
    
    <% If sWorkingTime <> "" Then %>
	<h5>就業時間</h5>
	<p class="m0"><%= sWorkingTime %></p>
    <% End If %>
    
    <% If dbWorkTimeRemark <> "" Then %>
	<h5>就業時間備考</h5>
	<p class="m0"><%= dbWorkTimeRemark %></p>
    <% End If %>
       
    </td>
    </tr>
    <% End If %>
    
    <%
		If flgDspHoliday = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">休日</th>
    <td class="dborder_bottom">    
    
    <% If dbWeeklyHolidayType <> "" Then %>
	<h5>休日</h5>
	<p class="m0"><%= dbWeeklyHolidayType %></p>
    <% End If %>
    
    <% If dbHolidayRemark <> "" Then %>
	<h5>休日備考</h5>
	<p class="m0"><%= dbHolidayRemark %></p>
    <% End If %>
    
    </td>
    </tr>
    <% End If %>
    
    <%
		If flgDspHumanNumber = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">募集人数</th>
    <td class="dborder_bottom">
	<p class="m0"><%= dbHumanNumber %></p>

    </td>
    </tr>
    <% End If %>  
    
    <%
		If flgDspWorkingPlace = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
			%>
    <tr>
    <th class="dborder_bottom">勤務地</th>
    <td class="dborder_bottom">

	<% If sWorkingPlace <> "" Then %>
	<h5>住所</h5>
	<p class="m0"><%= sWorkingPlace %></p>
	<% If sNearbyStationBlock <> "" Then
				sHTML = sHTML & sNearbyStationBlock
			End If 
			End If
			%>
   <% If dbTransfer <> "" Then %>
	<h5>転勤</h5>
	<p class="m0"><%= dbTransfer %></p>
    <% End If %>

    </td>
    </tr>
    <% End If %>     
    
     
    </tbody>
    </table>

<div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div>  

<%

	

End Function

'******************************************************************************
'概　要：求人票の必要条件を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：しごとナビ/order/order_detail.asp
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'　　　：2008/11/12 LIS K.Kokubo ベスト・ベターパターン出力
'　　　：2010/10/18 LIS T.Ezaki  ベターパターンの非表示対応
'　　　：2012/03/12 LIS K.Kokubo 卒業年出力
'******************************************************************************
Function DspNeedCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	
	Dim dbOrderCode			'情報コード
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'求人票種類
	Dim sCompanyKbn			'企業区分
	Dim dbTempOrderFlag		'派遣案件フラグ
	Dim dbBestMatchStr		'ベストパターン
	Dim dbBetterMatchStr	'ベターパターン
	Dim sAge				'年齢制限
	Dim sAgeMin				'年齢下限
	Dim sAgeMax				'年齢上限
	Dim sAgeReasonFlag		'年齢理由フラグ
	Dim sAgeReason			'年齢理由
	Dim sAgeReasonDetail	'年齢制限理由詳細
	Dim sFEHistory			'学歴
	Dim dbGraduateYearMin	'卒業年下限
	Dim dbGraduateYearMax	'卒業年上限
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

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	dbOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	dbTempOrderFlag = rRS.Collect("TempOrderFlag")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'<ベスト・ベターパターン>
	'紹介・紹介予定派遣のみ
    '2014/05/12 必要条件の欄にマッチングポイントを追加するため、下記のif文に派遣（sOrderType = "1"）を追記：木村
	If sOrderType = "2" Or sOrderType = "3" Or sOrderType = "1" Then
		dbBestMatchStr = ChkStr(rRS.Collect("BestMatchStr"))
		dbBetterMatchStr = ChkStr(rRS.Collect("BetterMatchStr"))
	End If
	'</ベスト・ベターパターン>

	'******************************************************************************
	'年齢 start
	'------------------------------------------------------------------------------
	sAge = ""
	sAgeMin = ChkStr(rRS.Collect("AgeMin"))
	sAgeMax = ChkStr(rRS.Collect("AgeMax"))
	sAgeReasonFlag = ChkStr(rRS.Collect("AgeReasonFlag"))
	sAgeReason = ChkStr(rRS.Collect("AgeReason"))
	sAgeReasonDetail = Replace(ChkStr(rRS.Collect("AgeReasonDetail")), vbCrLf, "<br>")

	If dbTempOrderFlag = "1" Then
		sAge = "派遣案件のため、年齢掲載していません。<br>"
		sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[？]制限について</a>"
	ElseIf sAgeReasonFlag = "0" Or sAgeReasonFlag = "" Or (sAgeMin & sAgeMax = "") Then
		sAge = "年齢不問<br>"
		'sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[？]制限について</a>"
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
	dbGraduateYearMin = rRS.Collect("GraduateYearMin")
	dbGraduateYearMax = rRS.Collect("GraduateYearMax")

	If sFEHistory <> "" Then sFEHistory = sFEHistory & "卒以上"

	If dbGraduateYearMin + dbGraduateYearMax > 0 Then
		sFEHistory = sFEHistory & "<br>[卒業年] "
		If dbGraduateYearMin > 0 Then
			sFEHistory = sFEHistory & dbGraduateYearMin & "年卒"
		End If
		sFEHistory = sFEHistory & " ～ "
		If dbGraduateYearMax > 0 Then
			sFEHistory = sFEHistory & dbGraduateYearMax & "年卒"
		End If
	End If

	If sFEHistory <> "" Then DspNeedCondition = True
	'------------------------------------------------------------------------------
	'学歴 end
	'******************************************************************************

	'******************************************************************************
	'資格 start
	'------------------------------------------------------------------------------
	sLicense = GetLicense(rDB, rRS)
	sLicenseOther = GetOrderNote(rDB, rRS, "OtherLicense")
	sLicenseOther = Replace(sLicenseOther, vbCrLf, "<br>")
	flgLicense = False
	If sLicense & sLicenseOther <> "" Then
		flgLicense = True
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'資格 end
	'******************************************************************************

    '2014/04/25 歓迎資格追加 池田
	'******************************************************************************
	'歓迎資格 start
	'------------------------------------------------------------------------------
	Dim sLicense_want

    sLicense_want = GetLicense_Want(rDB, rRS)

	If sLicense_want <> "" Then
		flgLicense = True
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'歓迎資格 end
	'******************************************************************************


    'Dim sLicense_MustWant

    'If rRS.Collect("LicenseMustFlag") = "0" Then
    '    sLicense_MustWant = "※上記いずれかの資格を保有"
    'ElseIf rRS.Collect("LicenseMustFlag") = "1" Then
    '    sLicense_MustWant = "※上記全ての資格を保有"
    'Else
    '     sLicense_MustWant = "※ 必須資格条件未設定"
    'End If



	'******************************************************************************
	'スキル start
	'------------------------------------------------------------------------------
	sSkillOS = GetSkill(rDB, rRS, "OS")
	sSkillApp = GetSkill(rDB, rRS, "Application")
	sSkillDL = GetSkill(rDB, rRS, "DevelopmentLanguage")
	sSkillDB = GetSkill(rDB, rRS, "Database")
	sSkillOther = GetOrderNote(rDB, rRS, "OtherSkill")
	sSkillOther = Replace(sSkillOther, vbCrLf, "<br>")
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
		sOtherNote = Replace(GetOrderNote(rDB, rRS, "OtherNote"), vbCrLf, "<br>")
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'その他特記事項 end
	'******************************************************************************

	flgLine = False

%>	
	<img src="/img/order/tab_detail_ne.png" class="tab_img">
	<table class="detail_table">
	<tbody>
    <%
	If dbBestMatchStr & dbBetterMatchStr <> "" Then
		If flgLine = True Then Response.Write ""
		flgLine = True
	
	%>
    <tr>
    <th class="dborder_bottom">マッチングポイント
    <p class="smartNone">
    	[<span style="color:#0045F9;cursor:pointer;" onclick="window.open('<%= HTTPS_CURRENTURL %>infomation/matchingpoint.asp','matchingpoint','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=300');">？</span>]
    </p>
    </th>
    <td class="dborder_bottom">
   <% If dbBestMatchStr <> "" Then %>
    <h4>ベスト</h4>
    <% Response.Write "<p class=""m0"">" & Replace(dbBestMatchStr, vbCrLf, "<br>") %></p>
    <% End If %>
    </td>
    </tr>
    <% End If %>
   
    <%
	If flgLine = True Then Response.Write ""
	flgLine = True
	
	%>
    <tr>
    <th class="dborder_bottom">年齢</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sAge %></p>    
    </td>
    </tr>
    
    
    <%
		If sFEHistory <> "" Then
		If flgLine = True Then Response.Write ""
		flgLine = True
	
	%>
    <tr>
    <th class="dborder_bottom">希望学歴</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sFEHistory %></p>    
    </td>
    </tr>
    <% End If

	sClearSolid = " style=""border-top-width:0px;"""
	If flgLicense = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write ""
		flgLine = True

		If sLicense <> "" Then
%>

            <tr>
            <th class="dborder_bottom">資格</th>
            <td class="dborder_bottom">
            <h4>必須資格</h4>
            <p class="m0"><%= sLicense %></p>
            
            </td>
            </tr>
 <% 	End If 


        '2014/04/25 歓迎資格追加 池田
		If sLicense_Want <> "" Then
%>

<tr>
    <th class="dborder_bottom"></th>
    <td class="dborder_bottom">
    <h4>歓迎資格</h4>
    <p class="m0"><%= sLicense_Want %></p>
    </td>
    </tr>
 <% 	End If 
 

		If sLicenseOther <> "" Then
	%>
    
    <tr>
    <th class="dborder_bottom"></th>
    <td class="dborder_bottom">
    <h4>その他資格</h4>
     <p class="m0"><%= sLicenseOther %></p>
    </td>
    </tr>   
	<%   End If
	
	End If 
		
	sClearSolid = " style=""border-top-width:0px;"""
	If flgSkill = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write ""
		flgLine = True
		
		%>
        	
	<tr>
    <th class="dborder_bottom">スキル</th>
    <td class="dborder_bottom">
		<% If sSkillOS <> "" Then 
			
			Response.Write "<h5 class=""skill_h5"">OS</h5>" & vbCrLf
			Response.Write "<div>" & sSkillOS & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillApp <> "" Then
			Response.Write "<h5 class=""skill_h5"">アプリケーション</h5>" & vbCrLf
			Response.Write "<div>" & sSkillApp & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillDL <> "" Then
			Response.Write "<h5 class=""skill_h5"">開発言語</h5>" & vbCrLf
			Response.Write "<div>" & sSkillDL & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillDB <> "" Then
			Response.Write "<h5 class=""skill_h5"">データベース</h5>" & vbCrLf
			Response.Write "<div>" & sSkillDB & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillOther <> "" Then
			Response.Write "<h5 class=""skill_h5"">その他スキル</h5>" & vbCrLf
			Response.Write "<div><p class=""m0"">" & sSkillOther & "</p></div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If


		%>
    </td>
    </tr>   
	<% End If %>
    <%
		If sOtherNote <> "" Then
		If flgLine = True Then Response.Write ""
		flgLine = True

	%>
	<tr>
    <th class="dborder_bottom">特記事項</th>
    <td class="dborder_bottom">
     <p class="m0"><%= sOtherNote %></p>
    </td>
    </tr>   
	<%
		sClearSolid = ""
	End If


%>
    </tbody>
    </table>
<div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div>  

<%
End Function

'******************************************************************************
'概　要：求人票の応募情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'******************************************************************************
Function DspHowToEntry(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim dbOrderCode			'情報コード
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
	Dim dbWValueURL			'Ｗバリューの自社採用ページＵＲＬ
	Dim flgEntryInfo		'応募情報が有るか無いか [True]ある [False]ない
	Dim flgProcess			'選考手順が有るか無いか [True]ある [False]ない
	Dim sClearSolid
	Dim flgLine				'線引きフラグ

	DspHowToEntry = False

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderType = ChkStr(rRS.Collect("OrderType"))
	dbOrderCode = ChkStr(rRS.Collect("OrderCode"))
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
	'連絡先 start
	'------------------------------------------------------------------------------
	sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
	sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
	sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
	sLis = sCPersonName & "［リス株式会社" & sCSectionName & "］　" & sCTel & "<br>(この案件はリス株式会社が取りまとめています。)"
	DspHowToEntry = True
	'------------------------------------------------------------------------------
	'連絡先 end
	'******************************************************************************

	'******************************************************************************
	'Ｗバリューの自社採用ページＵＲＬ start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(rRS.Collect("WValueURL"))
	If dbWValueURL <> "" Then
		DspHowToEntry = True
	End If
	'------------------------------------------------------------------------------
	'Ｗバリューの自社採用ページＵＲＬ end
	'******************************************************************************

	flgLine = False
	
	%>
    <img src="/img/order/tab_detail_ji.png" class="tab_img">
	<table class="detail_table">
	<tbody>
    <tr>
    <th class="dborder_bottom">情報コード</th>
    <td class="dborder_bottom">
    <p class="m0"><%= dbOrderCode %></p>
    </td>
	</tr>    
    
    <% If flgEntryInfo = True Then %>
    <!--<tr>
    <th class="dborder_bottom">応募方法</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sEntryInfo %></p>
    </td>
	</tr>   -->   
    <% End If %>
    
    <% If flgProcess = True Then %>
    <tr>
    <th class="dborder_bottom">選考手順</th>
    <td class="dborder_bottom">
    
    	<table>
    		<tr>
            <% If sProcess1 <> "" Then %>
            	<td class="stepTd">ステップ1</td>
                <td><%= sProcess1 %></td>
            <% Else %>
            	<td class="stepTd">ステップ1</td>
                <td>しごとナビに登録またはログイン後、応募ボタンよりご応募頂き書類選考いたします。</td>
            <% End If %>
        	</tr>
            <% If sProcess2 <> "" Then %>
            <tr>
            	<td class="stepTd">▼</td>
                <td></td>
        	</tr>
            <tr>
            	<td class="stepTd">ステップ2</td>
                <td><%= sProcess2 %></td>
        	</tr>
            <% End If %>
            <% If sProcess3 <> "" Then %>
            <tr>
            	<td class="stepTd">▼</td>
                <td></td>
        	</tr>
            <tr>
            	<td class="stepTd">ステップ3</td>
                <td><%= sProcess3 %></td>
        	</tr>
            <% End If %>
            <% If sProcess4 <> "" Then %>
            <tr>
            	<td class="stepTd">▼</td>
                <td></td>
        	</tr>
            <tr>
            	<td class="stepTd">ステップ4</td>
                <td><%= sProcess4 %></td>
        	</tr>
			<% End If %>
        </table>
   
    </td>
	</tr>      
    <% End If %>
    
    <% If dbWValueURL <> "" Then %>
    <tr>
    <th class="dborder_bottom">自社採用ページ</th>
    <td class="dborder_bottom">
    <p class="m0"><a href="<%= dbWValueURL %>" target="_blank"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_wvalue.gif" border="0" alt="自社採用ページ"></a></p>
    </td>
	</tr>      
    <% End If %>
    
    
    </tbody>
    </table>
<div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div> 
<%

End Function


'******************************************************************************
'概　要：求人票の応募情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：
'履　歴：2013/09/09 LIS T.seki 作成
'******************************************************************************
Function DspHowToEntry2(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim dbOrderCode			'情報コード
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
	Dim dbWValueURL			'Ｗバリューの自社採用ページＵＲＬ
	Dim flgEntryInfo		'応募情報が有るか無いか [True]ある [False]ない
	Dim flgProcess			'選考手順が有るか無いか [True]ある [False]ない
	Dim sClearSolid
	Dim flgLine				'線引きフラグ

	DspHowToEntry2 = False

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderType = ChkStr(rRS.Collect("OrderType"))
	dbOrderCode = ChkStr(rRS.Collect("OrderCode"))
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
		DspHowToEntry2 = True
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
		DspHowToEntry2 = True
	End If
	'------------------------------------------------------------------------------
	'選考手順 end
	'******************************************************************************

	'******************************************************************************
	'連絡先 start
	'------------------------------------------------------------------------------
	sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
	sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
	sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
	sLis = sCPersonName & "［リス株式会社" & sCSectionName & "］　" & sCTel & "<br>(この案件はリス株式会社が取りまとめています。)"
	DspHowToEntry2 = True
	'------------------------------------------------------------------------------
	'連絡先 end
	'******************************************************************************

	'******************************************************************************
	'Ｗバリューの自社採用ページＵＲＬ start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(rRS.Collect("WValueURL"))
	If dbWValueURL <> "" Then
		DspHowToEntry2 = True
	End If
	'------------------------------------------------------------------------------
	'Ｗバリューの自社採用ページＵＲＬ end
	'******************************************************************************

	flgLine = False
	
	%>
   
	<table class="jCodeOnly">
    	<thead>
        	<td colspan="2">※求人に関するお問い合わせやお祝い金申請の際には、必ず下記の「情報コード」をお知らせください。「しごとナビを見た」とおっしゃって頂けますとスムースです。</td>
        </thead>
        <tbody>
            <tr>
                <th class="dborder_bottom">情報コード</th>
                <td class="dborder_bottom">
                	<p class="m0"><%= dbOrderCode %></p>
                </td>
            </tr>    
        </tbody>
    </table>
<%

End Function

'******************************************************************************
'概　要：求人票の担当者連絡先を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使用元：
'備　考：
'履　歴：2007/02/11 LIS K.Kokubo 作成
'　　　：2009/04/02 LIS K.Kokubo メール課金プランの場合は連絡先を非表示に
'******************************************************************************
Function DspContact(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim dbOrderCode			'情報コード
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
	Dim dbPlanTypeName
	Dim flgLine				'線引きフラグ

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	If sOrderType <> "0" Then Exit Function
	dbPlanTypeName = rRS.Collect("PlanTypeName")
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

	'Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
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

		If sCompanyKbn = "2" Then
			'人材会社の求人票の場合は「名前」＋「人材会社名」
			sPerson = sCPersonName & "&nbsp;(人材会社：" & sCompanyName & ")"
		Else
			'一般企業の求人票の場合は「名前」＋「カナ」
			sPerson = sCPersonName
			If sCPersonNameF <> "" Then sPerson = sPerson & "(" & sCPersonNameF & ")"
		End If
	End If

	sContact = ""
	If sCTel <> "" Then sContact = sContact & sCTel & "	<SPAN style='font-size:10px;'>　※電話等でのお問い合わせの際、「しごとナビを見た」と言うとスムーズです。</SPAN>"
	If sContact <> "" Then sContact = sContact & "<br>"
	If sCMail <> "" Then sContact = sContact & sCMail
	'------------------------------------------------------------------------------
	'仕事の連絡先
	'******************************************************************************

	flgLine = False
	
	%>
    <img src="/img/order/tab_detail_tn.png" class="tab_img">
    <table class="detail_table">
	<tbody>
    <tr>
    <th class="dborder_bottom">担当者情報</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sPerson %></p>
    </td>
    
	<% If sCSectionName <> "" Then %>
    <tr>
    <th class="dborder_bottom">担当部署</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sCSectionName %></p>
    </td>
	</tr>     
   	<% End If %>
    
   	<% If dbPlanTypeName <> "mail" Then %>
    <tr>
    <th class="dborder_bottom">連絡先</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sContact %></p>
    </td>
	</tr>     
   	<% End If %>
   
    </tbody>
    </table>
    <div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div> 
    
    <%

End Function

'******************************************************************************
'概　要：求人票詳細の先輩インタビューを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'備　考：
'使用元：しごとナビ/order/order_detail.asp
'履　歴：2008/01/30 LIS K.Kokubo
'******************************************************************************
Function DspElderInterview(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbSeq
	Dim dbProfile
	Dim dbQuestion
	Dim dbAnswer
	Dim dbPublicFlag
	Dim dbPictureFlag

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	sSQL = "EXEC up_LstC_ElderInterview '" & dbOrderCode & "', '1'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
%>
<h3 id="interview_h3">企業からのメッセージ</h3>
<div class="freeprblock">
<%
		Do While GetRSState(oRS) = True
			dbSeq = oRS.Collect("Seq")
			dbProfile = Replace(oRS.Collect("Profile"), vbCrLf, "<br>")
			dbQuestion = Replace(oRS.Collect("Question"), vbCrLf, "<br>")
			dbAnswer = Replace(oRS.Collect("Answer"), vbCrLf, "<br>")
			dbPublicFlag = oRS.Collect("PublicFlag")
			dbPictureFlag = oRS.Collect("PictureFlag")
%>
		
		<div class="interview">
        	
<%
			If dbPictureFlag = "1" Then
				'先輩写真有り
%>
			<h4 class="interview_h4"><%= dbProfile %></h4>
			<div>
				<img src="<%= HTTP_NAVI_CURRENTURL %>company/elderinterview/picture.asp?ordercode=<%= dbOrderCode %>&amp;seq=<%= dbSeq %>" alt="">
			</div>
			<h5 class="interview_p"><%= dbQuestion %></h5>
			<p><%= dbAnswer %></p>
			
			<br clear="both">
		</div>
<%
			Else
				'先輩写真無し
%>			
			<h4 class="interview_h4_no"><%= dbProfile %></h4>
            <h5 class="interview_h"><%= dbQuestion %></h5>
            <p><%= dbAnswer %></p>
            <br clear="both">
		</div>
<%
			End If
			oRS.MoveNext
		Loop
%>
</div>
<br>
<%
	End If
End Function

'******************************************************************************
'概　要：リスの案件担当者、コンサル所見を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
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
    Dim sEmployeeFrigana		'コンサルタント名フリガナ
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

	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")

	If sOrderType <> "0" Then
		'******************************************************************************
		'コンサルタント start
		'------------------------------------------------------------------------------
		'リス受注票の場合は「リス担当者名」＋「リス担当者カナ」
		sEmployeeCode = ChkStr(rRS.Collect("EmployeeCode"))
		sEmployeeName = ChkStr(rRS.Collect("EmployeeName"))
        sEmployeeFrigana = ChkStr(rRS.Collect("EmployeeFrigana"))
		sBranchName = ChkStr(rRS.Collect("LisDepartment"))
		sTel = ChkStr(rRS.Collect("LisTelephoneNumber"))

        '2017/05/16　川根さんのみ部署名変更ロジック by 計良
        '2015/09/14　川根さんのみ部署名変更ロジック by 池田
        '2014/09/11　小幡さんのみ部署名変更ロジック by tanizawa
        '-------------------------
        Dim sSQL : sSQL =""
        Dim flgQE
	Dim sError
        Dim oRS

        IF sEmployeeCode = "L0000488" Then '小幡さん
        	sSQL = "sp_GetDataJobType '" & rRS.Collect("OrderCode") & "'"
	        flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
            If GetRSState(rRS) = True Then
	            Do While GetRSState(oRS) = True
		            If Left(oRS.Collect("JobTypeCode"),2) = "13" Then
                        sBranchName = "メディカルチーム"
                        Exit Do
                    End IF
		            oRS.MoveNext
	            Loop
            End IF
            Call RSClose(oRS)
        End IF

        IF sEmployeeCode = "L0000381" Then '川根さん

        	sSQL = "sp_GetDataJobType '" & rRS.Collect("OrderCode") & "'"
	        flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
            If GetRSState(rRS) = True Then
	            Do While GetRSState(oRS) = True

		            If Left(oRS.Collect("JobTypeCode"),2) = "07" Then
                        sBranchName = "会計事務所・税理士法人・監査法人チーム"
                        Exit Do
                    Else 
                        sBranchName = "人材紹介チーム"

                    End IF
		            oRS.MoveNext
	            Loop
            End IF
            Call RSClose(oRS)
        End IF

        '-------------------------

		sImg = "<img src=""/consultant/consultantimage.asp?ec=" & sEmployeeCode & """ alt=""この求人情報を担当しているコンサルタント"" border=""1"" width=""180"" height=""180"" style=""border-color:#666666;"">"
		sComment = Replace(ChkStr(rRS.Collect("ConsultantComment")), vbCrLf, "<br>")
		sComment = Replace(sComment, vbCr, "<br>")
		sComment = Replace(sComment, vbLf, "<br>")
		sConsultantPublicFlag = ChkStr(rRS.Collect("ConsultantPublicFlag"))
		sPictureFlag = ChkStr(rRS.Collect("ConsultantPictureFlag"))

		'2016/06/22　木村：コンサル紹介ページへのリンク削除
		sConsultantLink = Split(sEmployeeName,"　")(0)
		'If sConsultantPublicFlag = "1" Then
		'	sConsultantLink = "<a href=""" & HTTP_NAVI_CURRENTURL & "consultant/consultantdetail.asp?ec=" & sEmployeeCode & """>" & sConsultantLink & "</a>"
		'End If
        if sEmployeeFrigana = "" Then
            sConsultantLink = Split(sEmployeeName,"　")(0)
        Else
            sConsultantLink = "<ruby><rb>" & Split(sEmployeeName,"　")(0) & "</rb><rp>（</rp><rt>" & Split(sEmployeeFrigana,"　")(0) & "</rt><rp>）</rp></ruby>"    
        End If
        

		sConsultantLink = sConsultantLink & "&nbsp;(人材会社：リス株式会社)"
		'------------------------------------------------------------------------------
		'コンサルタント end
		'******************************************************************************

		sTitle = "担当者連絡先"
		If sComment <> "" Then sTitle = "この求人情報を担当しているコンサルタントの所見"
	%>
    <img src="/img/order/tab_detail_cn.png" class="tab_img">
	<table class="detail_table">
	<tbody>
    <tr>
    <th class="dborder_bottom">コンサルタント</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sConsultantLink %></p>
    </td>
    <tr>
    <th class="dborder_bottom">担当部署</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sBranchName %></p>
    </td>
	</tr>     
    <tr>
    <th class="dborder_bottom">連絡先</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sTel %><span>※お問い合わせの際、上記「情報コード」と「しごとナビを見た」とおっしゃって下さるとスムーズです。</span></p>
    </td>
	</tr>     
   	<% End If %>
    
    <% If sComment <> "" Then %>
    <tr>
    <th class="dborder_bottom">所見</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sComment %></p>
    </td>
	</tr>     
   	<% End If %>
   
    </tbody>
    </table>
    <div class="to_top"><a class="stext_middle" href="#pagetop">▲ページTOPへ</a></div> 
    <%


End Function

'******************************************************************************
'概　要：最新メールを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
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
		sSQL = "up_DtlMailHistory_Order '" & vUserID & "', '" & rRS.Collect("OrderCode") & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sDateTime = GetDateStr(oRS.Collect("SendDay"), "/") & "　" & GetTimeStr(oRS.Collect("SendDay"), ":")
			sSubject = ChkStr(oRS.Collect("Subject"))
			sDetail = Replace(ChkStr(oRS.Collect("Body")), vbCrLf, "<br>")
			sDetail = Replace(sDetail, vbCr, "<br>")
			sDetail = Replace(sDetail, vbLf, "<br>")
			Response.Write "<h3 class=""sp"">最新の送信済みメール</h3>"
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			Response.Write "<div class=""category1""><h4>送信日時</h4></div>"
			Response.Write "<div class=""value1""><p class=""m0"">" & sDateTime & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			Response.Write "<div class=""category1""><h4>サブジェクト</h4></div>"
			Response.Write "<div class=""value1""><p class=""m0"">" & sSubject & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			Response.Write "<div class=""category1""><h4>内容</h4></div>"
			Response.Write "<div class=""value1""><p class=""m0"">" & sDetail & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<br>"
		End If
	End If

	Call RSClose(oRS)

	DspNewMail = True
End Function

'******************************************************************************
'概　要：求人票詳細ページの勤務形態部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
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
			Select Case oRS.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""haken_tr"">派遣とは</a>】</span>" 
				Case "002","003": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""shokai_tr"">人材紹介とは</a>】</span>" 
				Case "004": sWorkingType = sWorkingType & "<span class=""smartNone"">【<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>紹介予定派遣とは</a>】</span>" 
			End Select
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
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
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

    Dim count 
    count = 1
	Do While GetRSState(oRS) = True
		sJobType = sJobType & "(" & count & ") " & oRS.Collect("JobTypeName") & ""
        count = count + 1
		oRS.MoveNext
		If GetRSState(oRS) = True Then sJobType = sJobType & "<br>"
	Loop
	Call RSClose(oRS)

	GetJobType = sJobType
End Function

'******************************************************************************
'概　要：求人票詳細ページの勤務形態部分
'引　数：rDB	：接続中のDBConnection
'　　　：rRS	：up_DtlOrderで生成されたレコードセットオブジェクト
'備　考：
'更　新：2006/05/08 LIS K.Kokub 作成
'　　　：2009/11/17 LIS K.Kokubo FC,SOHO案件の場合は勤務時間を返さない
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
    Dim count
    count = 1
	If rRS.Collect("FCSOHOOrderFlag") = "0" Then
		sSQL = "sp_GetDataWorkingTime '" & rRS.Collect("OrderCode") & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			sWST = ChkStr(oRS.Collect("DspWorkStartTime"))
			sWET = ChkStr(oRS.Collect("DspWorkEndTime"))
			If sWST & sWET <> "" Then
				sWorkingTime = sWorkingTime & "(" & count & ") " & sWST & "～" & sWET
                count = count + 1
			End If
			oRS.MoveNext
			If GetRSState(oRS) = True And sWST & sWET <> "" Then sWorkingTime = sWorkingTime & "<br>"
		Loop
		Call RSClose(oRS)
	End If

	GetWorkingTime = sWorkingTime
End Function

'******************************************************************************
'概　要：求人票詳細ページの最寄駅部分
'引　数：rDB	：接続中のDBConnection
'　　　：rRS	：up_LstC_NearbyStationで生成されたレコードセットオブジェクト
'　　　：vWPSeq	：勤務地番号
'使　用：ナビ/include/func_order.asp
'備　考：
'履　歴：2006/05/08 LIS K.Kokubo 作成
'　　　：2008/10/22 LIS K.Kokubo 求人票勤務地複数化対応
'******************************************************************************
Function GetNearbyStation(ByRef rDB, ByRef rRS)
	Dim dbWorkingPlaceSeq
	Dim dbStationName
	Dim dbToStationTime
	Dim dbToStationRemark

	Dim idx
	Dim sStation
	Dim sToStation
	Dim iStation

	If GetRSState(rRS) = False Then Exit Function

	iStation = 0
	sStation = ""
	Do While GetRSState(rRS) = True
		dbWorkingPlaceSeq = rRS.Collect("WorkingPlaceSeq")
		dbStationName = ChkStr(rRS.Collect("StationName"))
		dbToStationTime = ChkStr(rRS.Collect("ToStationTime"))
		dbToStationRemark = ChkStr(rRS.Collect("ToStationRemark"))
		iStation = iStation + 1

		sToStation = ""
		If dbToStationTime <> "" Then sToStation = dbToStationTime & "分"
		If dbToStationRemark <> "" Then sToStation = dbToStationRemark & sToStation
		If sToStation <> "" Then sToStation = "(" & sToStation & ")"

		If sStation <> "" Then sStation = sStation & "/"
		sStation = sStation & dbStationName & "駅" & sToStation

		rRS.MoveNext
	Loop

	GetNearbyStation = sStation
End Function

'******************************************************************************
'概　要：求人票詳細ページの最寄沿線部分
'引　数：rDB	：接続中のDBConnection
'　　　：rRS	：up_LstC_NearbyRailwayLineで生成されたレコードセットオブジェクト
'使　用：ナビ/include/func_order.asp
'備　考：
'履　歴：2006/05/08 LIS K.Kokubo 作成
'　　　：2008/10/22 LIS K.Kokubo 求人票勤務地複数化対応
'******************************************************************************
Function GetNearbyRailway(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbWorkingPlaceSeq
	Dim dbRailwayLineName2

	Dim idx
	Dim iRowCnt
	Dim sRailway
	Dim iRailway

	If GetRSState(rRS) = False Then Exit Function

	iRowCnt = rRS.RecordCount
	iRailway = 0
	sRailway = ""
	Do While GetRSState(rRS) = True And iRailway < 3
		dbWorkingPlaceSeq = rRS.Collect("WorkingPlaceSeq")
		dbRailwayLineName2 = rRS.Collect("RailwayLineName2")
		iRailway = iRailway + 1

		If sRailway <> "" Then sRailway = sRailway & ","
		sRailway = sRailway & dbRailwayLineName2

		rRS.MoveNext
	Loop
	If iRowCnt > 3 Then sRailway = sRailway & "&nbsp;他"

	GetNearbyRailway = sRailway
End Function


'******************************************************************************
'概　要：求人票詳細ページの最寄沿線部分
'引　数：rDB	：接続中のDBConnection
'　　　：rRS	：up_LstC_NearbyRailwayLineで生成されたレコードセットオブジェクト
'使　用：ナビ/include/func_order.asp
'備　考：
'履　歴：2015/09/14 池田 元関数をコピーし、新幹線は表示しないように仕様追加
'******************************************************************************
Function GetNearbyRailway2(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbWorkingPlaceSeq
	Dim dbRailwayLineName2

	Dim idx
	Dim iRowCnt
	Dim sRailway
	Dim iRailway

	If GetRSState(rRS) = False Then Exit Function

	'iRowCnt = rRS.RecordCount

    iRowCnt = 0
	iRailway = 0
	sRailway = ""
	Do While GetRSState(rRS) = True And iRailway < 3
		dbWorkingPlaceSeq = rRS.Collect("WorkingPlaceSeq")
		dbRailwayLineName2 = rRS.Collect("RailwayLineName2")
		

		If sRailway <> "" Then sRailway = sRailway & ","
		
        If InStr(dbRailwayLineName2 ,"新幹線") Then
            '新幹線は表示しない
        Else
            sRailway = sRailway & dbRailwayLineName2

            iRailway = iRailway + 1
            iRowCnt = iRowCnt + 1
        End If

		rRS.MoveNext
	Loop
	If iRowCnt > 3 Then sRailway = sRailway & "&nbsp;他"

	GetNearbyRailway2 = sRailway
End Function


'******************************************************************************
'概　要：求人票詳細ページのスキル部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
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

		sSkill = sSkill & "<p style=""min-width:25%; max-width:50%; float:left; height:40px;""><span style=""color:#339933;"">■</span> " & oRS.Collect("SkillName")
		If ChkStr(oRS.Collect("Period")) <> "" Then
			sSkill = sSkill & "<br>　" & oRS.Collect("Period") & "年以上は尚可"
		End If
		sSkill = sSkill & "</p>"
		If iSkill Mod SKILLCOL = 0 Then sSkill = sSkill & ""

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
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
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

		sLicense = sLicense & "<p style=""width:50%; float:left;"">(" & iLicense & ") " & oRS.Collect("LicenseName") & "</p>"
		If iLicense Mod LICENSECOL = 0 Then sLicense = sLicense & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

    '2014/07/17 必須資格フラグ追加 たに
	'******************************************************************************
	'必須資格フラグ start
	'------------------------------------------------------------------------------
    If sLicense <> "" Then
        if iLicense > 1 Then
            If rRS.Collect("LicenseMustFlag") = "0" Then
                sLicense = sLicense & "<p style=""width:50%; float:left;"">※上記いずれかの資格を保有していること</p>"
            ElseIF rRS.Collect("LicenseMustFlag") = "1" Then
                sLicense = sLicense & "<p style=""width:50%; float:left;"">※上記全ての資格を保有していること</p>"
            End If
        End IF
    End If
    '------------------------------------------------------------------------------
	'必須資格フラグ end
	'******************************************************************************

	'中途半端で終わった場合の調整
	If sLicense <> "" And iLicense Mod LICENSECOL <> 0 Then sLicense = sLicense & "<br clear=""all"">"

	GetLicense = sLicense
End Function

'******************************************************************************
'概　要：求人票詳細ページの歓迎資格部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'作成者：Lis ikeda
'作成日：2014/04/25
'備　考：
'******************************************************************************
Function GetLicense_Want(ByRef rDB, ByRef rRS)
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

	sSQL = "sp_GetDataLicense_Want '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iLicense = iLicense + 1

		sLicense = sLicense & "<p style=""width:50%; float:left;"">(" & iLicense & ") " & oRS.Collect("LicenseName") & "</p>"
		If iLicense Mod LICENSECOL = 0 Then sLicense = sLicense & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'中途半端で終わった場合の調整
	If sLicense <> "" And iLicense Mod LICENSECOL <> 0 Then sLicense = sLicense & "<br clear=""all"">"

	GetLicense_Want = sLicense
End Function

'******************************************************************************
'概　要：求人票詳細ページのその他情報取得
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
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
Function GetOrderTitle(ByRef rDB, ByVal vOrderCode, ByRef rTitle, ByRef rKeywords, ByRef rDescription)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWorkingType
	
	Dim rRS

	sSQL = "EXEC up_DtlOrderTitle '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		'rTitle = ChkStr(oRS.Collect("JobTypeDetail")) & "&nbsp;" & ChkStr(oRS.Collect("PrefectureName"))
		rTitle = ChkStr(oRS.Collect("JobTypeDetail")) & "&nbsp;" & ChkStr(oRS.Collect("CatchCopy"))
		rKeywords = "求人情報,転職," & ChkStr(oRS.Collect("JobTypeDetail")) & "," & ChkStr(oRS.Collect("PrefectureName"))
		If ChkStr(oRS.Collect("JobTypeName")) <> "" Then rKeywords = rKeywords & "," & ChkStr(oRS.Collect("JobTypeName"))
		If ChkStr(oRS.Collect("WorkingTypeName")) <> "" Then rKeywords = rKeywords & "," & ChkStr(oRS.Collect("WorkingTypeName"))
		rDescription = "転職・求人情報：" & ChkStr(oRS.Collect("BusinessDetail"))
		If rDescription = "" Then rDescription = "転職・求人情報：" & ChkStr(oRS.Collect("JobTypeDetail"))
	End If
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
'履　歴：
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
			sSQL = "up_SearchRelationAccessOrder '" & vOrderCode & "'"
			sTitle = "この求人情報を見た人はこんな求人情報も見ています"
		Case "2"
			sSQL = "up_SearchHighRelationOrder '" & vOrderCode & "'"
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
'出　力：rJobTypeDetail		：具体的職種名
'　　　：rCompanyName		：企業名
'　　　：rImg				：企業イメージ
'　　　：rWorkingTypeIcon	：勤務形態アイコン
'　　　：rWorkingPlace		：勤務地
'　　　：rStation			：最寄駅 '2008/10/22 LIS K.Kokubo 不使用
'　　　：rYearlyIncome		：年収
'　　　：rMonthlyIncome		：月給
'　　　：rDailyIncome		：日給
'　　　：rHourlyIncome		：時給
'戻り値：
'備　考：
'履　歴：2007/05/31 LIS K.Kokubo 作成
'　　　：2008/10/22 LIS K.Kokubo 勤務地複数化による修正
'******************************************************************************
Function GetRecommendValues(ByRef rDB, ByRef rRS, ByVal vRCMD, ByRef rJobTypeDetail, ByRef rCompanyName, ByRef rImg, ByRef rWorkingTypeIcon, ByRef rWorkingPlace, ByRef rStation, ByRef rYearlyIncome, ByRef rMonthlyIncome, ByRef rDailyIncome, ByRef rHourlyIncome)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'情報コード
	Dim dbCompanyCode		'企業コード
	Dim dbOrderType			'受注区分
	Dim dbCompanyKbn		'会社区分
	Dim dbCompanyName		'企業名
	Dim dbCompanyNameF		'企業名カナ
	Dim dbCompanySpeciality	'企業名（特徴）
	Dim dbJobTypeDetail		'具体的職種名(altやtitleで出力する)
	Dim dbYearlyIncomeMin	'年収下限
	Dim dbYearlyIncomeMax	'年収上限
	Dim dbMonthlyIncomeMin	'月給下限
	Dim dbMonthlyIncomeMax	'月給上限
	Dim dbDailyIncomeMin	'日給下限
	Dim dbDailyIncomeMax	'日給上限
	Dim dbHourlyIncomeMin	'時給下限
	Dim dbHourlyIncomeMax	'時給上限
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName
	Dim dbWorkingPlaceCity
	Dim dbImageLimit

	Dim sViewJobTypeDetail	'求職者に見える具体的職種名(長い文字列はカットされる)
	Dim sYearlyIncome		'年収
	Dim sMonthlyIncome		'月給
	Dim sDailyIncome		'日給
	Dim sHourlyIncome		'時給
	Dim sWorkingTypeIcon	'勤務形態アイコン並び
	Dim sWorkingPlace		'勤務地
	Dim sImg				'画像URL

	Dim idx
	Dim sURL				'求人票詳細のURL
	Dim sAlign				'枠寄せ [vCols = 1]left [vCols = vMaxCols]right [それ以外]center

	If GetRSState(rRS) = False Then Exit Function

	sURL = HTTPS_CURRENTURL & "order/order_detail.asp"

	sSQL = "up_DtlOrder '" & rRS.Collect("OrderCode") & "', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'情報コード
	dbOrderCode = ChkStr(oRS.Collect("OrderCode"))
	'企業コード
	dbCompanyCode = ChkStr(oRS.Collect("CompanyCode"))
	'受注区分
	dbOrderType = ChkStr(oRS.Collect("OrderType"))
	'企業区分
	dbCompanyKbn = ChkStr(oRS.Collect("CompanyKbn"))
	'企業名, 企業名カナ
	dbCompanyName = ChkStr(oRS.Collect("CompanyName"))
	dbCompanyNameF = ChkStr(oRS.Collect("CompanyName_F"))
	dbCompanySpeciality = ChkStr(oRS.Collect("CompanySpeciality"))
	Call SetOrderCompanyName(dbCompanyName, dbCompanyNameF, dbOrderType, dbCompanyKbn, dbCompanySpeciality)
	'具体的職種名
	dbJobTypeDetail = ChkStr(oRS.Collect("JobTypeDetail"))
	sViewJobTypeDetail = dbJobTypeDetail
	If Len(sViewJobTypeDetail) > 14 Then sViewJobTypeDetail = Left(sViewJobTypeDetail, 14) & ".."
	'写真
	dbImageLimit = oRS.Collect("ImageLimit")

	'******************************************************************************
	'給与 start
	'------------------------------------------------------------------------------
	'年収
	dbYearlyIncomeMin = ChkStr(oRS.Collect("YearlyIncomeMin"))
	dbYearlyIncomeMax = ChkStr(oRS.Collect("YearlyIncomeMax"))
	If dbYearlyIncomeMin = "0" Then dbYearlyIncomeMin = ""
	If dbYearlyIncomeMax = "0" Then dbYearlyIncomeMax = ""
	If dbYearlyIncomeMin <> "" Then dbYearlyIncomeMin = GetJapaneseYen(dbYearlyIncomeMin)
	If dbYearlyIncomeMax <> "" Then dbYearlyIncomeMax = GetJapaneseYen(dbYearlyIncomeMax)
	If dbYearlyIncomeMin & dbYearlyIncomeMax <> "" Then
		If dbYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMin
		sYearlyIncome = sYearlyIncome & "&nbsp;～&nbsp;"
		If dbYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMax
	End If
	'月給
	dbMonthlyIncomeMin = ChkStr(oRS.Collect("MonthlyIncomeMin"))
	dbMonthlyIncomeMax = ChkStr(oRS.Collect("MonthlyIncomeMax"))
	If dbMonthlyIncomeMin = "0" Then dbMonthlyIncomeMin = ""
	If dbMonthlyIncomeMax = "0" Then dbMonthlyIncomeMax = ""
	If dbMonthlyIncomeMin <> "" Then dbMonthlyIncomeMin = GetJapaneseYen(dbMonthlyIncomeMin)
	If dbMonthlyIncomeMax <> "" Then dbMonthlyIncomeMax = GetJapaneseYen(dbMonthlyIncomeMax)
	If dbMonthlyIncomeMin & dbMonthlyIncomeMax <> "" Then
		If dbMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & dbMonthlyIncomeMin
		sMonthlyIncome = sMonthlyIncome & "&nbsp;～&nbsp;"
		If dbMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & dbMonthlyIncomeMax
	End If
	'日給
	dbDailyIncomeMin = ChkStr(oRS.Collect("DailyIncomeMin"))
	dbDailyIncomeMax = ChkStr(oRS.Collect("DailyIncomeMax"))
	If dbDailyIncomeMin = "0" Then dbDailyIncomeMin = ""
	If dbDailyIncomeMax = "0" Then dbDailyIncomeMax = ""
	If dbDailyIncomeMin <> "" Then dbDailyIncomeMin = GetJapaneseYen(dbDailyIncomeMin)
	If dbDailyIncomeMax <> "" Then dbDailyIncomeMax = GetJapaneseYen(dbDailyIncomeMax)
	If dbDailyIncomeMin & dbDailyIncomeMax <> "" Then
		If dbDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & dbDailyIncomeMin
		sDailyIncome = sDailyIncome & "&nbsp;～&nbsp;"
		If dbDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & dbDailyIncomeMax
	End If
	'時給
	dbHourlyIncomeMin = ChkStr(oRS.Collect("HourlyIncomeMin"))
	dbHourlyIncomeMax = ChkStr(oRS.Collect("HourlyIncomeMax"))
	If dbHourlyIncomeMin = "0" Then dbHourlyIncomeMin = ""
	If dbHourlyIncomeMax = "0" Then dbHourlyIncomeMax = ""
	If dbHourlyIncomeMin <> "" Then dbHourlyIncomeMin = GetJapaneseYen(dbHourlyIncomeMin)
	If dbHourlyIncomeMax <> "" Then dbHourlyIncomeMax = GetJapaneseYen(dbHourlyIncomeMax)
	If dbHourlyIncomeMin & dbHourlyIncomeMax <> "" Then
		If dbHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMin
		sHourlyIncome = sHourlyIncome & "&nbsp;～&nbsp;"
		If dbHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMax
	End If
	'------------------------------------------------------------------------------
	'給与 end
	'******************************************************************************

	'******************************************************************************
	'勤務形態アイコン start
	'------------------------------------------------------------------------------
	sWorkingTypeIcon = ""
	sSQL = "sp_GetListWorkingType '" & dbOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Select Case ChkStr(oRS.Collect("WorkingTypeCode"))
			Case "001": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/haken.gif"" alt=""派遣"" style=""margin-right:1px;"">"
			Case "002": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/seishain.gif"" alt=""正社員"" style=""margin-right:1px;"">"
			Case "003": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/keiyaku.gif"" alt=""契約社員"" style=""margin-right:1px;"">"
			Case "004": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/syoha.gif"" alt=""紹介予定派遣"" style=""margin-right:1px;"">"
			Case "005": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/arbeit.gif"" alt=""アルバイト・パート"" style=""margin-right:1px;"">"
			Case "006": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/soho.gif"" alt=""SOHO"" style=""margin-right:1px;"">"
			Case "007": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/fc.gif"" alt=""FC"" style=""margin-right:1px;"">"
		End Select
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'勤務形態アイコン end
	'******************************************************************************

	'******************************************************************************
	'画像 start
	'------------------------------------------------------------------------------
	sImg = ""
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If sImg = "" And ChkStr(oRS.Collect("PicNo1")) <> "" Then sImg = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			If sImg = "" And ChkStr(oRS.Collect("PicNo2")) <> "" Then sImg = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo2")
			If sImg = "" And ChkStr(oRS.Collect("PicNo3")) <> "" Then sImg = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo3")
			If sImg = "" And ChkStr(oRS.Collect("PicNo4")) <> "" Then sImg = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo4")
		End If
		Call RSClose(oRS)
	ElseIf dbImageLimit > 1 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If sImg = "" And ChkStr(oRS.Collect("OptionNo1")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo1")
			If sImg = "" And ChkStr(oRS.Collect("OptionNo2")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
			If sImg = "" And ChkStr(oRS.Collect("OptionNo3")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
			If sImg = "" And ChkStr(oRS.Collect("OptionNo4")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
		End If
	End If

	If sImg = "" Then sImg = "/img/no%20image.png"
	'sImg = "<img src=""" & sImg & """ alt=""" & dbCompanyName & """ width=""156"" height=""117"">"
	sImg = "<img src=""" & sImg & """ alt=""" & dbCompanyName & """ width=""88"" height=""66"" border=""0"" align=""left"" style=""margin:0px; padding:0px;"">"
	'------------------------------------------------------------------------------
	'画像 end
	'******************************************************************************

	'******************************************************************************
	'勤務地 start
	'------------------------------------------------------------------------------
	idx = 0
	sWorkingPlace = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True And idx < 3
		dbWorkingPlacePrefectureCode = ChkStr(oRS.Collect("WorkingPlacePrefectureCode"))
		dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS.Collect("WorkingPlaceCity"))

		'<勤務地>
		If sWorkingPlace <> "" Then sWorkingPlace = sWorkingPlace & "/"
		sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		'</勤務地>

		oRS.MoveNext
		idx = idx + 1
	Loop
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'最寄駅 end
	'******************************************************************************

	rJobTypeDetail = "<a href=""" & sURL & "?ordercode=" & dbOrderCode & "&amp;rcmd=" & vRCMD & """>" & sViewJobTypeDetail & "</a>"
	rCompanyName = dbCompanyName
	rImg = "<a href=""" & sURL & "?ordercode=" & dbOrderCode & "&amp;rcmd=" & vRCMD & """>" & sImg & "</a>"
	rWorkingTypeIcon = sWorkingTypeIcon
	rWorkingPlace = sWorkingPlace
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
				" '" & G_USERID & "'" & _
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
'戻り値：
'備　考：
'履　歴：2007/02/14 LIS K.Kokubo 作成
'　　　：2008/05/08 LIS K.Kokubo 特徴追加(シークレット求人)
'　　　：2008/08/19 LIS M.Hayashi 特徴追加
'　　　：2008/10/20 LIS K.Kokubo 勤務地複数化による修正
'　　　：2009/03/18 LIS K.Kokubo 特徴追加(ナビ無料化対応)
'******************************************************************************
Function GetImgOrderSpeciality(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName

	Dim sHTML
	Dim sWorkingCode

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	sHTML = ""
	'アクセス数が100を超えていれば「HOT」表示（リス安藤）
	If rRS.Collect("AccessCount") > 100 Then sHTML = sHTML & "<img src=""/img/c_HOT_green.gif"" alt=""人気"" width=""50"" height=""15"">&nbsp;"
	'UPDATEと今日から10日引いた日で「新着」表示(リス安藤)
	If rRS.Collect("Updateday") > NOW()-10 Then sHTML = sHTML & "<img src=""/img/c_NEW_green.gif"" alt=""新着"" width=""50"" height=""15"">&nbsp;"
	'未経験者ＯＫの場合、わかばマーク表示(リス安藤)
	If rRS.Collect("InexperiencedPersonFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp01=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/no_experience.gif"" alt=""未経験者歓迎"" width=""50"" height=""15""></a>&nbsp;"
	'Ｕターン・Ｉターン
	If rRS.Collect("UITurnFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp04=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/ui_turn.gif"" alt=""Ｕターン・Ｉターン"" width=""50"" height=""15""></a>&nbsp;"
	'語学を活かす仕事
	If rRS.Collect("UtilizeLanguageFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp02=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/linguistic_job.gif"" alt=""語学を活かす仕事"" width=""50"" height=""15""></a>&nbsp;"
	'年間休日120日以上
	If rRS.Collect("ManyHolidayFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp05=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/year_holidaycnt.gif"" alt=""年間休日120日以上"" width=""50"" height=""15""></a>&nbsp;"
	'2006/01/10 M.Hayashi ADD フレックスタイム制度あり
	If rRS.Collect("FlexTimeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp06=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_flextime.gif"" alt=""フレックスタイム制度あり"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NearStationFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp07=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_nearstation.gif"" alt=""駅近(徒歩5分以内)"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NoSmokingFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp08=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_nosmoking.gif"" alt=""禁煙・分煙"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NewlyBuiltFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp09=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_newlybuilt.gif"" alt=""新築ビル・オフィス(5年以内)"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("LandmarkFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp10=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_landmark.gif"" alt=""高層(15階以上)ビル"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("RenovationFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp11=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_renovation.gif"" alt=""リノベーションビル・オフィス(5年以内)"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DesignersFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp12=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_designers.gif"" alt=""デザイナーズビル・オフィス"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("CompanyCafeteriaFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp13=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_companycafeteria.gif"" alt=""社員食堂"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortOvertimeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp14=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_shortovertime.gif"" alt=""残業10h/月以内"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MaternityFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp15=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_maternity.gif"" alt=""産休・育休実績あり"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DressFreeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp16=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_dressfree.gif"" alt=""服装自由"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MammyFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp17=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_mammy.gif"" alt=""子育てママ歓迎"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("FixedTimeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp18=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_fixedtime.gif"" alt=""18時までに退社"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortTimeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp19=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_shorttime.gif"" alt=""1日6時間以内労働"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("HandicappedFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp20=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_handicapped.gif"" alt=""障害者歓迎"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentAllFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp21=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_rentallflag.gif"" alt=""住宅費用全額補助あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentPartFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp22=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_rentpartflag.gif"" alt=""住宅費用一部補助あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp23=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_mealsflag.gif"" alt=""食事・賄い付き案件"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsAssistanceFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp24=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_mealsassistanceflag.gif"" alt=""食事補助制度あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("TrainingCostFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp25=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_trainingcostflag.gif"" alt=""研修費助成制度あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("EntrepreneurCostFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp26=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_entrepreneurcostflag.gif"" alt=""起業機材補助制度あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MoneyFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp27=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_moneyflag.gif"" alt=""無利子・低利子補助制度あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("LandShopFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp28=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_landshopflag.gif"" alt=""土地・店舗等提供制度あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("FindJobFestiveFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp29=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_findjobfestiveflag.gif"" alt=""就職お祝い金制度あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("AppointmentFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp30=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_appointmentflag.gif"" alt=""正社員登用制度あり"" width=""50"" height=""15""></a>&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp31=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_societyinsuranceflag.gif"" alt=""社保完備"" width=""50"" height=""15""></a>&nbsp;"
	'2008/05/08 LIS K.Kokubo ADD シークレット求人
	If rRS.Collect("SecretFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order/secret.gif"" alt=""スカウトを受けた人だけが閲覧できる求人情報"" width=""50"" height=""15"">&nbsp;"

	'直接Yahoo!の検索からお仕事情報詳細ページへ来る人へアイコン表示
	If InStr(Request.ServerVariables("HTTP_REFERER"),"search.yahoo.co.jp/") <> 0 Then
		sSQL = "sp_GetDataWorkingType '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			sWorkingcode = oRS.Collect("WorkingTypecode")

			sHTML = sHTML & "<img src=""/img/order_detail_icon/icon_w" & sWorkingcode & ".gif"" alt=""派遣社員"" width=""50"" height=""15"">&nbsp;"

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
				sHTML = sHTML & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15"">&nbsp;"
			End If
		End If
		Call RSClose(oRS)
		'</勤務地>
	End If

	GetImgOrderSpeciality = sHTML
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
<div id="top_reg_button">

	<a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button1.gif" alt="履歴書登録して応募" border="0"></a>
	<a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button3.gif" alt="ログインして応募" border="0"></a>
	
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
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button3.gif" alt="ログインして応募" border="0"></a></div>
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button1.gif" alt="履歴書登録して応募" border="0"></a></div>
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
<div class="center">
	<p>
▼▼会員登録すれば応募や質問が可能になります！▼▼<BR>
応募のための履歴書も自動作成されます。</p>

	<div class="center left"><a href="<%= HTTPS_NAVI_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0"></a></div>
	<div class="center right"><a href="<%= HTTPS_NAVI_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0"></a></div>
	<br style="clear:both;">

</div>

<!--新SNSボタン-->
<div id="sns_button" class="smartNone">
<!-- #INCLUDE FILE="../include/social_bookmark.asp" -->
<div class="right">
<!--G+-->
<!-- +1 ボタン を表示したい位置に次のタグを貼り付けてください。 -->
<div class="g-plusone" data-size="tall" data-annotation="none"></div>

<!-- 最後の +1 ボタン タグの後に次のタグを貼り付けてください。 -->
<script type="text/javascript">
  window.___gcfg = {lang: 'ja'};

  (function() {
    var po = document.createElement('script'); po.type = 'text/javascript'; po.async = true;
    po.src = 'https://apis.google.com/js/plusone.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(po, s);
  })();
</script>
</div>

<!--facebook-->

<div class="fb-like" data-href="http://<%= Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") %>" data-send="false" data-layout="button_count" data-width="110" data-show-faces="false"></div>

<!--<script type="text/javascript">
    var url = encodeURIComponent(location.href);
    document.write('<iframe src="http://www.facebook.com/plugins/like.php?href=' + url + '&width=100&layout=button_count&show_faces=false&action=like&colorscheme=light&height=20" scrolling="no" frameborder="0" style="border:none; overflow:hidden;width:100px;height:20px;" allowTransparency="true"></iframe>');
</script>-->

<!--/facebook-->
<!--twitter-->
<a href="https://twitter.com/share" class="twitter-share-button" data-via="shigoto_navi" data-lang="ja" data-count="none">ツイート</a>
<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>

</div>
<!--/新SNSボタン-->
<br clear="both">

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
Sub DspBottomRegButton_OldPlan(ByVal vOrderCode,ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount,ByVal YearlyIncomeMin,ByVal MonthlyIncomeMin,ByVal DailyIncomeMin,ByVal HourlyIncomeMin)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic

	Dim JobTypeBigCode
	Dim JobTypeCode
	Dim WorkingTypeCode1
	Dim WorkingTypeCode2
	Dim WorkingTypeCode3
	Dim PrefectureCode
	
	If GetRSState(rRS) = False Then Exit Sub

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

	'雇用形態、勤務地、職種検索
			sSQL = "select CJT.jobtypecode, BJT.Bigclasscode from C_JobType AS CJT INNER JOIN B_JobType AS BJT ON CJT.JobTypeCode = BJT.AllConnectCode where CJT.id = '1' and CJT.OrderCode = '" & dbOrderCode & "';"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("Bigclasscode")) <> "" Then
					JobTypeBigCode = oRS.Collect("Bigclasscode")
				End If
				If ChkStr(oRS.Collect("jobtypecode")) <> "" Then
					JobTypeCode = oRS.Collect("jobtypecode")
				End If
			End If
			Call RSClose(oRS)

			sSQL = "select prefecturecode from c_workingplace where ordercode = '" & dbOrderCode & "';"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			PrefectureCode  = ""
			Do While GetRSState(oRS) = True
				If ChkStr(oRS.Collect("prefecturecode")) <> "" Then
					PrefectureCode = PrefectureCode & oRS.Collect("prefecturecode") & ","
				End If
				oRS.MoveNext
			Loop
			Call RSClose(oRS)
			PrefectureCode = Left(PrefectureCode, Len(PrefectureCode) -1)

			sSQL = "select CWT1.workingtypecode as workingtypecode1,CWT2.workingtypecode as workingtypecode2,CWT3.workingtypecode as workingtypecode3 from c_workingtype AS CWT1 "
			sSQL = sSQL & " left join c_workingtype AS CWT2 on CWT1.ordercode = '" & dbOrderCode & "' and CWT2.id = 2"
			sSQL = sSQL & " left join c_workingtype AS CWT3 on CWT2.ordercode = '" & dbOrderCode & "' and CWT3.id = 3"
			sSQL = sSQL & " where CWT3.ordercode = '" & dbOrderCode & "' and CWT1.id = 1;"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("workingtypecode1")) <> "" Then
					WorkingTypeCode1 = oRS.Collect("workingtypecode1")
				End If
				If ChkStr(oRS.Collect("workingtypecode2")) <> "" Then
					WorkingTypeCode2 = oRS.Collect("workingtypecode2")
				End If
				If ChkStr(oRS.Collect("workingtypecode3")) <> "" Then
					WorkingTypeCode3 = oRS.Collect("workingtypecode3")
				End If
			End If
			Call RSClose(oRS)

%>
<div class="center">
	<p>
▼▼会員登録すれば応募や質問が可能になります！▼▼<BR>
応募のための履歴書も自動作成されます。</p>

	<div class="center left">            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_reg_button03.png" alt="履歴書登録して応募" border="0"></a></div>
	<div class="center right"><a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_list.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_login_button03.png" alt="ログインして応募" border="0"></a></div>
	<br style="clear:both;">

</div>

<!--新SNSボタン-->
<div id="sns_button" class="smartNone">
<!-- #INCLUDE FILE="../include/social_bookmark.asp" -->
<div class="right">
<!--G+-->
<!-- +1 ボタン を表示したい位置に次のタグを貼り付けてください。 -->
<div class="g-plusone" data-size="tall" data-annotation="none"></div>

<!-- 最後の +1 ボタン タグの後に次のタグを貼り付けてください。 -->
<script type="text/javascript">
  window.___gcfg = {lang: 'ja'};

  (function() {
    var po = document.createElement('script'); po.type = 'text/javascript'; po.async = true;
    po.src = 'https://apis.google.com/js/plusone.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(po, s);
  })();
</script>
</div>

<!--facebook-->

<div class="fb-like" data-href="http://<%= Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") %>" data-send="false" data-layout="button_count" data-width="110" data-show-faces="false"></div>

<!--<script type="text/javascript">
    var url = encodeURIComponent(location.href);
    document.write('<iframe src="http://www.facebook.com/plugins/like.php?href=' + url + '&width=100&layout=button_count&show_faces=false&action=like&colorscheme=light&height=20" scrolling="no" frameborder="0" style="border:none; overflow:hidden;width:100px;height:20px;" allowTransparency="true"></iframe>');
</script>-->

<!--/facebook-->
<!--twitter-->
<a href="https://twitter.com/share" class="twitter-share-button" data-via="shigoto_navi" data-lang="ja" data-count="none">ツイート</a>
<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>

</div>
<!--/新SNSボタン-->
<br clear="both">

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
	<div align="center" style="float:left; width:300px;color:#C51035;">＜まだIDをお持ちでない方＞<br><a href="<%= HTTPS_NAVI_CURRENTURL %>resume/staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button1.gif" alt="履歴書登録して応募" border="0"></a></div>
	<div align="center" style="float:right; width:300px;color:#C51035;">＜すでにIDをお持ちの方＞<br><a href="<%= HTTPS_NAVI_CURRENTURL %>resume/login/login.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/resume/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button3.gif" alt="ログインして応募" border="0"></a></div>
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
'備　考：
'使用元：order/order_detail.asp
'履　歴：2007/05/08 LIS K.Kokubo 作成
'　　　：2009/05/19 LIS K.Kokubo 社内からのアクセスとS0018066のアクセスはログに残さない
'　　　：2009/06/01 LIS.T.Ezaki  パラメーター（uc）にスタッフコードが記載あればログに記録する
'******************************************************************************
Function AccessHistoryOrder(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	'社内からのアクセスと、たたろうさん(S0018066)からのアクセスはログに残さない
	If IsRE(G_IPADDRESS, "^192.168.", True) = False And vUserID <> "S0018066" Then
		If vUserType = "staff" Then
			sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & vUserID & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		ElseIf IsRE(Request.Cookies("id_memory"), "^S\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & Request.Cookies("id_memory") & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		ElseIf IsRE(GetForm("uc",2), "^S\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & GetForm("uc",2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
			sSQL = "update P_Userinfo set lastaccessday = getdate() where staffcode = '" & GetForm("uc",2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		End If
	End If
End Function

'******************************************************************************
'概　要：足跡ログの書き込み(新着メルマガ)
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_SearchOrder or 求人票詳細検索SQL で生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vOrderCode		：閲覧中求人票
'備　考：
'使用元：order/order_detail.asp
'履　歴：2014/09/11 LIS TANIZAWA 作成
'******************************************************************************
Function AccessHistoryOrderNEW(ByRef rDB, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	'社内からのアクセスと、たたろうさん(S0018066)からのアクセスはログに残さない
	If IsRE(G_IPADDRESS, "^192.168.", True) = False And vUserID <> "S0018066" Then
		If vUserType = "staff" Then
			sSQL = "up_Reg_LOG_NewOrderMailAccessHistory '" & vOrderCode & "', '" & vUserID & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		ElseIf IsRE(Request.Cookies("id_memory"), "^S\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "up_Reg_LOG_NewOrderMailAccessHistory '" & vOrderCode & "', '" & Request.Cookies("id_memory") & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		End If
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
	If GetRSState(oRS) = True Then
		AccessCountUp = oRS.Collect("AccessCount")
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：求人票の日別ＰＶのカウントアップ
'引　数：rDB		：接続中のDBConnection
'　　　：vOrderCode	：閲覧中求人票の情報コード
'備　考：
'使　用：order/order_detail.asp
'履　歴：2008/05/23 LIS K.Kokubo 作成
'******************************************************************************
Function PVCountUp(ByRef rDB, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	sSQL = "up_RegC_PV '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Call RSClose(oRS)
End Function

'*******************************************************************************
'概　要：全角半角が混じった文字列のバイト数を正確に返す(Webからの引用)
'引　数：string		:対象文字列
'戻り値：Interger	:対象文字列のバイト数
'作成日：2007/05/23 Lis Sotome
'履　歴：
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
'履　歴：
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
