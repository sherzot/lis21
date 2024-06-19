<%
'******************************************************************************
'概　要：検索条件を保持するクラス
'関　数：■Public
'　　　：GetSearchParam				：お仕事詳細検索ページへ渡すGETパラメータを生成して取得
'　　　：GetSQLOrderSearchDetail	：求人票詳細検索ＳＱＬを取得
'　　　：GetHtmlSearchCondition		：求人票詳細検索条件出力ＨＴＭＬを取得
'　　　：
'　　　：■Private
'　　　：Class_Initialize			：コンストラクタ
'　　　：SetNames					：コードに対応した名称をメンバ変数に設定
'　　　：ChkData					：メンバ変数の整合性をチェックして訂正
'　　　：
'備　考：■■■ 詳細検索用パラメータ （アドホックなＳＱＬ生成）
'　　　：sotf	：社内外案件検索フラグ
'　　　：snewf	：新着フラグ（１週間以内に掲載のあったもの）
'　　　：sjtbig1：希望職種大分類１
'　　　：sjt1	：希望職種１
'　　　：sjtbig2：希望職種大分類２
'　　　：sjt2	：希望職種２
'　　　：src	：希望沿線
'　　　：ssc	：希望駅
'　　　：spc	：希望都道府県
'　　　：sct	：希望市区郡
'　　　：swt1	：希望勤務形態１
'　　　：swt2	：希望勤務形態２
'　　　：swt3	：希望勤務形態３
'　　　：sit	：希望業種(カンマ区切り [XX,XX,XX])
'　　　：ssp01	：特徴（未経験歓迎）
'　　　：ssp02	：特徴（語学を活かす）
'　　　：ssp03	：特徴（派遣）※現在未使用
'　　　：ssp04	：特徴（ＵＩターン）
'　　　：ssp05	：特徴（休日１２０日以上）
'　　　：ssp06	：特徴（フレックスタイム）
'　　　：ssp07	：特徴（駅近）
'　　　：ssp08	：特徴（禁煙・分煙）
'　　　：ssp09	：特徴（新築ビル・オフィス）
'　　　：ssp10	：特徴（高層ビル（ランドマーク））
'　　　：ssp11	：特徴（リノベーションビル・オフィス）
'　　　：ssp12	：特徴（デザイナーズビル・オフィス）
'　　　：ssp13	：特徴（社員食堂）
'　　　：ssp14	：特徴（残業10h以内）
'　　　：ssp15	：特徴（産休・育休実績あり）
'　　　：ssp16	：特徴（服装自由）
'　　　：ssp17	：特徴（子育てママ歓迎）
'　　　：ssp18	：特徴（18時までに退社）
'　　　：ssp19	：特徴（1日6時間以内労働）
'　　　：ssp20	：特徴（障害者歓迎）
'　　　：ssp21	：特徴（住宅費用全額補助あり）
'　　　：ssp22	：特徴（住宅費用一部補助あり）
'　　　：ssp23	：特徴（食事・賄い付き案件）
'　　　：ssp24	：特徴（食事補助制度あり）
'　　　：ssp25	：特徴（研修費助成制度あり）
'　　　：ssp26	：特徴（起業機材補助制度あり）
'　　　：ssp27	：特徴（無利子・低利子補助制度あり）
'　　　：ssp28	：特徴（土地・店舗等提供制度あり）
'　　　：ssp29	：特徴（就職お祝い金制度あり）
'　　　：ssp30	：特徴（正社員登用制度あり）
'　　　：ssp31	：特徴（社保完備）
'　　　：snewf	：新着フラグ
'　　　：sppf	：歩合制
'　　　：syimin	：年収下限
'　　　：syimax	：年収上限
'　　　：smimin	：月給下限
'　　　：smimax	：月給上限
'　　　：sdimin	：日給下限
'　　　：sdimax	：日給上限
'　　　：shimin	：時給下限
'　　　：shimax	：時給上限
'　　　：swsh	：就業開始時間（時）
'　　　：swsm	：就業開始時間（分）
'　　　：sweh	：就業終了時間（時）
'　　　：swem	：就業終了時間（分）
'　　　：swht	：週休種類
'　　　：sage	：年齢
'　　　：sat	：契約期間
'　　　：slg1	：資格大分類
'　　　：slc1	：資格中分類
'　　　：sl1	：資格小分類
'　　　：sos1	：ＯＳ
'　　　：sap1	：アプリケーション
'　　　：sdl1	：開発言語
'　　　：sdb1	：データベース
'　　　：skw	：検索ワード
'　　　：skwflg	：検索ワードフラグ [1]OR [2]AND
'　　　：sst	：特徴ビット文字列(000000)
'　　　：poc	：情報コード（対象企業の求人票一覧用）
'　　　：soc	：情報コード（検索）
'　　　：socs	：情報コードCSV
'　　　：slocc	：社内案件の対象企業コード
'　　　：snewfkouko	：広告新着フラグ
'　　　：
'　　　：■■■ カンタン検索用パラメータ (ストアド up_SearchOrder 活用)
'　　　：jt		：職種大分類コード
'　　　：jt2	：職種コード
'　　　：ac		：エリアコード ※2012/02/28 LIS K.Kokubo 削除
'　　　：ac2	：都道府県コード
'　　　：wt		：勤務形態コード
'　　　：kw		：キーワード
'　　　：
'　　　：■■■ 情報ツール用
'　　　：boc	：前回表示情報コード
'　　　：
'　　　：■■■ 使用方法
'　　　：Dim oSOC
'　　　：Dim sSQL
'　　　：Set oSOC = New clsSearchOrderCondition	'生成された時点でパラメータとＰＯＳＴデータからＳＱＬが生成されている
'　　　：oSOC.Top = 100	'SELECT句で上限を設定
'　　　：sSQL = oSOC.GetSQLOrderSearchDetail()	'ＳＱＬを取得
'　　　：
'履　歴：2007/04/05 LIS K.Kokubo 作成
'　　　：2007/10/10 LIS K.Kokubo 情報ツール用変数追加
'　　　：2007/10/31 LIS K.Kokubo TOP ??? 用変数追加
'　　　：2008/01/15 LIS K.Kokubo パラメータ化クエリ化
'　　　：2008/03/26 LIS K.Kokubo 登録日検索追加
'　　　：2008/08/14 LIS M.Hayashi 特徴フラグ追加とフレックス移動
'　　　：2009/11/17 LIS K.Kokubo 給与や時間の変数に全角数字があった場合、半角数字に変換
'　　　：2010/10/08 LIS K.Kokubo 社内案件の対象企業コード追加
'　　　：2012/02/28 LIS K.Kokubo エリアコードの検索利用を廃止
'　　　：2012/03/12 LIS K.Kokubo 年齢検索廃止＆卒業年検索追加
'******************************************************************************
Class clsSearchOrderCondition
	'検索条件メンバ変数
	Public Top						'SELECTで取得する件数 (SELECT TOP ○ * FROM ～)
	PUblic SearchDetailFlag			'詳細検索フラグ
	Public OrderTypeFlag			'社内外案件検索フラグ
	Public NewFlag					'新着フラグ
	Public JobTypeBigCode1			'希望職種大分類１
	Public JobTypeCode1				'希望職種１
	Public JobTypeBigCode2			'希望職種大分類２
	Public JobTypeCode2				'希望職種２
	Public JobTypeBigCode3			'希望職種大分類３
	Public JobTypeCode3				'希望職種３
	Public RailwayLineCode			'希望沿線(カンマ区切り)
	Public StationCode				'希望駅(カンマ区切り)
	Public PrefectureCode			'希望都道府県(カンマ区切り)
	Public City						'希望市区郡(カンマ区切り)
	Public WorkingTypeCode1			'希望勤務形態１
	Public WorkingTypeCode2			'希望勤務形態２
	Public WorkingTypeCode3			'希望勤務形態３
	Public IndustryTypeCode			'希望業種(カンマ区切り [XX,XX,XX])
	Public PercentagePayFlag		'歩合制
	Public YearlyIncomeMin			'年収下限
	Public YearlyIncomeMax			'年収上限
	Public MonthlyIncomeMin			'月給下限
	Public MonthlyIncomeMax			'月給上限
	Public DailyIncomeMin			'日給下限
	Public DailyIncomeMax			'日給上限
	Public HourlyIncomeMin			'時給下限
	Public HourlyIncomeMax			'時給上限
	Public WorkStartHour			'就業開始時間（時）
	Public WorkStartMinute			'就業開始時間（分）
	Public WorkEndHour				'就業終了時間（時）
	Public WorkEndMinute			'就業終了時間（分）
	Public WeeklyHolidayType		'週休種類
	'Public Age						'年齢
	Public SchoolTypeCode			'卒業年検索（学歴）
	Public GraduateYear				'卒業年検索（卒業年）
	Public AgreementTerm			'契約期間
	Public LicenseCount				'資格件数
	Public LicenseGroupCode			'資格大分類
	Public LicenseCategoryCode		'資格中分類
	Public LicenseCode				'資格小分類
	Public OSCode					'ＯＳ（CSV）
	Public ApplicationCode			'アプリケーション（CSV）
	Public DevelopmentLanguageCode	'開発言語（CSV）
	Public DatabaseCode			'データベース（CSV）
	Public Keyword					'検索ワード
	Public KeywordFlag				'検索ワードフラグ [1]OR [2]AND
	Public PictureOrderCode			'情報コード（対象企業の求人票一覧用）
	Public OrderCode				'情報コード（検索） CSV可
	Public Specialty
	Public InexperiencedPersonFlag	'特徴（未経験歓迎）
	Public UtilizeLanguageFlag		'特徴（語学を活かす）
	Public TempFlag					'特徴（派遣）
	Public UITurnFlag				'特徴（ＵＩターン歓迎）
	Public ManyHolidayFlag			'特徴（休日１２０日以上）
	Public FlexFlag					'特徴（フレックス）
	Public NearStationFlag			'特徴（駅近）
	Public NoSmokingFlag			'特徴（禁煙・分煙）
	Public NewlyBuiltFlag			'特徴（新築）
	Public LandmarkFlag				'特徴（高層）
	Public RenovationFlag			'特徴（リノベーション）
	Public DesignersFlag			'特徴（デザイナーズ）
	Public CompanyCafeteriaFlag		'特徴（社員食堂）
	Public ShortOvertimeFlag		'特徴（短時間残業）
	Public MaternityFlag			'特徴（産休育休実績あり）
	Public DressFreeFlag			'特徴（服装自由）
	Public MammyFlag				'特徴（ママ歓迎）
	Public FixedTimeFlag			'特徴（18時までに退社）
	Public ShortTimeFlag			'特徴（短時間労働）
	Public HandicappedFlag			'特徴（障害者歓迎）
	Public RentAllFlag				'特徴（住宅費用全額補助あり）
	Public RentPartFlag				'特徴（住宅費用一部補助あり）
	Public MealsFlag				'特徴（食事・賄い付き案件）
	Public MealsAssistanceFlag		'特徴（食事補助制度あり）
	Public TrainingCostFlag			'特徴（研修費助成制度あり）
	Public EntrepreneurCostFlag		'特徴（起業機材補助制度あり）
	Public MoneyFlag				'特徴（無利子・低利子補助制度あり）
	Public LandShopFlag				'特徴（土地・店舗等提供制度あり）
	Public FindJobFestiveFlag		'特徴（就職お祝い金制度あり）
	Public AppointmentFlag			'特徴（正社員登用制度あり）
	Public SocietyInsuranceFlag		'特徴（社保完備）
	Public RegistDay				'登録日
	Public LISOrderCompanyCode		'社内案件の対象企業コード
    Public NewKoukokuFlag          	'広告新着フラグ
    Public FeatureFlag          	'特集検索フラグ

	'カンタン検索条件
	Public JT	'職種大分類コード
	Public JT2	'職種コード
	Public AC2	'都道府県コード
	Public WT	'勤務形態コード
	Public KW	'キーワード

	'ＴＯＰの写真から
	Public POC

	'沿線検索
	Public PC	'都道府県コード
	Public RC	'沿線コード
	Public SC	'駅コード

	'特集
	Public SP	'特集コード

	'情報ツール
	Public BOC	'前回表示時の最新情報コード

	'コード対応名称
	Public JobTypeBigName1	'希望職種大分類名称１
	Public JobTypeName1	'希望職種名称１
	Public JobTypeBigName2	'希望職種大分類名称２
	Public JobTypeName2	'希望職種名称２
	Public JobTypeBigName3	'希望職種大分類名称３
	Public JobTypeName3	'希望職種名称３
	Public RailwayLineName	'希望沿線名称
	Public StationName
	Public AreaName
	Public PrefectureName
	Public WorkingTypeName1
	Public WorkingTypeName2
	Public WorkingTypeName3
	Public IndustryTypeName	'業種名配列
	Public WeeklyHolidayTypeName
	Public OSName
	Public ApplicationName
	Public DevelopmentLanguageName
	Public DatabaseName
	Public SchoolTypeName
	Public LicenseGroupName
	Public LicenseCategoryName
	Public LicenseName

	'その他メンバ変数
	Public HtmlOrderSearch	'検索条件出力ＨＴＭＬ文
	Public SQLOrderSearch	'検索ＳＱＬ
	Public SQLWriteLog		'ログ書き込みＳＱＬ

	'******************************************************************************
	'概　要：コンストラクタ
	'作成者：Lis K.Kokubo
	'作成日：2007/04/04 Lis K.Kokubo
	'更　新：
	'備　考：
	'******************************************************************************
	Private Sub Class_Initialize()
		LicenseCount = 0

		'パラメータから検索条件を取得
		Call ReadParam()

		'データ整合性チェック
		Call ChkData()

		'コード対応名称取得
		Call SetNames()

		'求人票検索SQL生成
		SQLOrderSearch = GetSQLOrderSearchDetail()

		'ログ書き込みSQL生成
		SQLWriteLog = GetSQLWriteLog()

		'求人票検索条件出力ＨＴＭＬ文
		HtmlOrderSearch = GetHtmlSearchCondition()

		'Response.Write SQLOrderSearch
	End Sub

	'******************************************************************************
	'概　要：GETデータの読み込み
	'引　数：
	'備　考：
	'履　歴：2007/04/04 LIS K.Kokubo 作成
	'******************************************************************************
	Private Sub ReadParam()
		Dim idx

		If GetForm("sdf", 2) <> "" Then SearchDetailFlag = GetForm("sdf", 2)
		If GetForm("sotf", 2) <> "" Then OrderTypeFlag = GetForm("sotf", 2)
		If GetForm("snewf", 2) <> "" Then NewFlag = GetForm("snewf", 2)
		If GetForm("sjtbig1", 2) <> "" Then JobTypeBigCode1 = GetForm("sjtbig1", 2)
		If GetForm("sjt1", 2) <> "" Then JobTypeCode1 = GetForm("sjt1", 2)
		If GetForm("sjtbig2", 2) <> "" Then JobTypeBigCode2 = GetForm("sjtbig2", 2)
		If GetForm("sjt2", 2) <> "" Then JobTypeCode2 = GetForm("sjt2", 2)
		If GetForm("sjtbig3", 2) <> "" Then JobTypeBigCode3 = GetForm("sjtbig3", 2)
		If GetForm("sjt3", 2) <> "" Then JobTypeCode3 = GetForm("sjt3", 2)
		If GetForm("src", 2) <> "" Then RailwayLineCode = Replace(GetForm("src", 2)," ","")
		If GetForm("ssc", 2) <> "" Then StationCode = Replace(GetForm("ssc", 2)," ","")
		If GetForm("spc", 2) <> "" Then PrefectureCode = Replace(GetForm("spc", 2)," ","")
		If GetForm("sct", 2) <> "" Then City = GetForm("sct", 2)
		If GetForm("swt1", 2) <> "" Then WorkingTypeCode1 = GetForm("swt1", 2)
		If GetForm("swt2", 2) <> "" Then WorkingTypeCode2 = GetForm("swt2", 2)
		If GetForm("swt3", 2) <> "" Then WorkingTypeCode3 = GetForm("swt3", 2)
		If GetForm("sit", 2) <> "" Then IndustryTypeCode = GetForm("sit", 2)
		If GetForm("ssp01", 2) <> "" Then InexperiencedPersonFlag = GetForm("ssp01", 2)
		If GetForm("ssp02", 2) <> "" Then UtilizeLanguageFlag = GetForm("ssp02", 2)
		If GetForm("ssp03", 2) <> "" Then TempFlag = GetForm("ssp03", 2)
		If GetForm("ssp04", 2) <> "" Then UITurnFlag = GetForm("ssp04", 2)
		If GetForm("ssp05", 2) <> "" Then ManyHolidayFlag = GetForm("ssp05", 2)
		If GetForm("ssp06", 2) <> "" Then FlexFlag = GetForm("ssp06", 2)
		If GetForm("ssp07", 2) <> "" Then NearStationFlag = GetForm("ssp07", 2)
		If GetForm("ssp08", 2) <> "" Then NoSmokingFlag = GetForm("ssp08", 2)
		If GetForm("ssp09", 2) <> "" Then NewlyBuiltFlag = GetForm("ssp09", 2)
		If GetForm("ssp10", 2) <> "" Then LandmarkFlag = GetForm("ssp10", 2)
		If GetForm("ssp11", 2) <> "" Then RenovationFlag = GetForm("ssp11", 2)
		If GetForm("ssp12", 2) <> "" Then DesignersFlag = GetForm("ssp12", 2)
		If GetForm("ssp13", 2) <> "" Then CompanyCafeteriaFlag = GetForm("ssp13", 2)
		If GetForm("ssp14", 2) <> "" Then ShortOvertimeFlag = GetForm("ssp14", 2)
		If GetForm("ssp15", 2) <> "" Then MaternityFlag = GetForm("ssp15", 2)
		If GetForm("ssp16", 2) <> "" Then DressFreeFlag = GetForm("ssp16", 2)
		If GetForm("ssp17", 2) <> "" Then MammyFlag = GetForm("ssp17", 2)
		If GetForm("ssp18", 2) <> "" Then FixedTimeFlag = GetForm("ssp18", 2)
		If GetForm("ssp19", 2) <> "" Then ShortTimeFlag = GetForm("ssp19", 2)
		If GetForm("ssp20", 2) <> "" Then HandicappedFlag = GetForm("ssp20", 2)
		If GetForm("ssp21", 2) <> "" Then RentAllFlag = GetForm("ssp21", 2)
		If GetForm("ssp22", 2) <> "" Then RentPartFlag = GetForm("ssp22", 2)
		If GetForm("ssp23", 2) <> "" Then MealsFlag = GetForm("ssp23", 2)
		If GetForm("ssp24", 2) <> "" Then MealsAssistanceFlag = GetForm("ssp24", 2)
		If GetForm("ssp25", 2) <> "" Then TrainingCostFlag = GetForm("ssp25", 2)
		If GetForm("ssp26", 2) <> "" Then EntrepreneurCostFlag = GetForm("ssp26", 2)
		If GetForm("ssp27", 2) <> "" Then MoneyFlag = GetForm("ssp27", 2)
		If GetForm("ssp28", 2) <> "" Then LandShopFlag = GetForm("ssp28", 2)
		If GetForm("ssp29", 2) <> "" Then FindJobFestiveFlag = GetForm("ssp29", 2)
		If GetForm("ssp30", 2) <> "" Then AppointmentFlag = GetForm("ssp30", 2)
		If GetForm("ssp31", 2) <> "" Then SocietyInsuranceFlag = GetForm("ssp31", 2)
		If GetForm("sppf", 2) <> "" Then PercentagePayFlag = GetForm("sppf", 2)
		If GetForm("syimin", 2) <> "" Then YearlyIncomeMin = Replace(Replace(GetForm("syimin", 2),",",""),"万","0000")
		If GetForm("syimax", 2) <> "" Then YearlyIncomeMax = Replace(Replace(GetForm("syimax", 2),",",""),"万","0000")
		If GetForm("smimin", 2) <> "" Then MonthlyIncomeMin = GetForm("smimin", 2)
		If GetForm("smimax", 2) <> "" Then MonthlyIncomeMax = GetForm("smimax", 2)
		If GetForm("sdimin", 2) <> "" Then DailyIncomeMin = GetForm("sdimin", 2)
		If GetForm("sdimax", 2) <> "" Then DailyIncomeMax = GetForm("sdimax", 2)
		If GetForm("shimin", 2) <> "" Then HourlyIncomeMin = GetForm("shimin", 2)
		If GetForm("shimax", 2) <> "" Then HourlyIncomeMax = GetForm("shimax", 2)
		If GetForm("swsh", 2) <> "" Then WorkStartHour = GetForm("swsh", 2)
		If GetForm("swsm", 2) <> "" Then WorkStartMinute = GetForm("swsm", 2)
		If GetForm("sweh", 2) <> "" Then WorkEndHour = GetForm("sweh", 2)
		If GetForm("swem", 2) <> "" Then WorkEndMinute = GetForm("swem", 2)
		If GetForm("swht", 2) <> "" Then WeeklyHolidayType = GetForm("swht", 2)
		'If GetForm("sage", 2) <> "" Then Age = GetForm("sage", 2)
		If GetForm("sstc", 2) <> "" Then SchoolTypeCode = GetForm("sstc", 2)
		If GetForm("sgy", 2) <> "" Then GraduateYear = GetForm("sgy", 2)
		If GetForm("sat", 2) <> "" Then AgreementTerm = GetForm("sat", 2)
		If GetForm("slocc",2) <> "" Then LISOrderCompanyCode = GetForm("slocc",2)
        If GetForm("snewfkouko", 2) <> "" Then NewKoukokuFlag = GetForm("snewfkouko", 2)
        If GetForm("FeatureFlag", 2) <> "" Then FeatureFlag = GetForm("FeatureFlag", 2)
        

		'<資格>
		idx = 0
		Do While (IsEmpty(Request.Querystring("slg"&idx+1)) = False Or IsEmpty(Request.Querystring("slc"&idx+1)) = False Or IsEmpty(Request.Querystring("sl"&idx+1)) = False)
			idx = idx + 1
		Loop
		LicenseCount = idx
		ReDim LicenseGroupCode(LicenseCount)
		ReDim LicenseCategoryCode(LicenseCount)
		ReDim LicenseCode(LicenseCount)
		ReDim LicenseGroupName(LicenseCount)
		ReDim LicenseCategoryName(LicenseCount)
		ReDim LicenseName(LicenseCount)
		For idx = 0 To LicenseCount - 1
			LicenseGroupCode(idx) = GetForm("slg"&idx+1, 2)
			LicenseCategoryCode(idx) = GetForm("slc"&idx+1, 2)
			LicenseCode(idx) = Right(GetForm("sl"&idx+1, 2),2)
		Next
		'</資格>

		If GetForm("sos", 2) <> "" Then OSCode = Replace(GetForm("sos", 2)," ","")
		If GetForm("sap", 2) <> "" Then ApplicationCode = Replace(GetForm("sap", 2)," ","")
		If GetForm("sdl", 2) <> "" Then DevelopmentLanguageCode = Replace(GetForm("sdl", 2)," ","")
		If GetForm("sdb", 2) <> "" Then DatabaseCode = Replace(GetForm("sdb", 2)," ","")
		If GetForm("skw", 2) <> "" Then Keyword = GetForm("skw", 2)
		If GetForm("skwflg", 2) <> "" Then KeywordFlag = GetForm("skwflg", 2)
		If GetForm("sst", 2) <> "" Then Specialty = GetForm("sst", 2)
		If GetForm("poc", 2) <> "" Then PictureOrderCode = GetForm("poc", 2)
		If GetForm("soc", 2) <> "" Then OrderCode = GetForm("soc", 2)
		If GetForm("srd", 2) <> "" Then RegistDay = GetForm("srd", 2)

		If IsRE(GetForm("sst", 2), "^[01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01]$", True) = True Then
			If Mid(GetForm("sst", 2), 1, 1) = "1" Then InexperiencedPersonFlag = "1"
			If Mid(GetForm("sst", 2), 2, 1) = "1" Then UtilizeLanguageFlag = "1"
			If Mid(GetForm("sst", 2), 3, 1) = "1" Then TempFlag = "1"
			If Mid(GetForm("sst", 2), 4, 1) = "1" Then UITurnFlag = "1"
			If Mid(GetForm("sst", 2), 5, 1) = "1" Then ManyHolidayFlag = "1"
			If Mid(GetForm("sst", 2), 6, 1) = "1" Then FlexFlag = "1"
			If Mid(GetForm("sst", 2), 7, 1) = "1" Then NearStationFlag = "1"
			If Mid(GetForm("sst", 2), 8, 1) = "1" Then NoSmokingFlag = "1"
			If Mid(GetForm("sst", 2), 9, 1) = "1" Then NewlyBuiltFlag = "1"
			If Mid(GetForm("sst", 2), 10, 1) = "1" Then LandmarkFlag = "1"
			If Mid(GetForm("sst", 2), 11, 1) = "1" Then RenovationFlag = "1"
			If Mid(GetForm("sst", 2), 12, 1) = "1" Then DesignersFlag = "1"
			If Mid(GetForm("sst", 2), 13, 1) = "1" Then CompanyCafeteriaFlag = "1"
			If Mid(GetForm("sst", 2), 14, 1) = "1" Then ShortOvertimeFlag = "1"
			If Mid(GetForm("sst", 2), 15, 1) = "1" Then MaternityFlag = "1"
			If Mid(GetForm("sst", 2), 16, 1) = "1" Then DressFreeFlag = "1"
			If Mid(GetForm("sst", 2), 17, 1) = "1" Then MammyFlag = "1"
			If Mid(GetForm("sst", 2), 18, 1) = "1" Then FixedTimeFlag = "1"
			If Mid(GetForm("sst", 2), 19, 1) = "1" Then ShortTimeFlag = "1"
			If Mid(GetForm("sst", 2), 20, 1) = "1" Then HandicappedFlag = "1"
			If Mid(GetForm("sst", 2), 21, 1) = "1" Then RentAllFlag = "1"
			If Mid(GetForm("sst", 2), 22, 1) = "1" Then RentPartFlag = "1"
			If Mid(GetForm("sst", 2), 23, 1) = "1" Then MealsFlag = "1"
			If Mid(GetForm("sst", 2), 24, 1) = "1" Then MealsAssistanceFlag = "1"
			If Mid(GetForm("sst", 2), 25, 1) = "1" Then TrainingCostFlag = "1"
			If Mid(GetForm("sst", 2), 26, 1) = "1" Then EntrepreneurCostFlag = "1"
			If Mid(GetForm("sst", 2), 27, 1) = "1" Then MoneyFlag = "1"
			If Mid(GetForm("sst", 2), 28, 1) = "1" Then LandShopFlag = "1"
			If Mid(GetForm("sst", 2), 29, 1) = "1" Then FindJobFestiveFlag = "1"
			If Mid(GetForm("sst", 2), 30, 1) = "1" Then AppointmentFlag = "1"
			If Mid(GetForm("sst", 2), 31, 1) = "1" Then SocietyInsuranceFlag = "1"
		End If

		'TOP:職種大分類
		If GetForm("jt", 2) <> "" Then JobTypeBigCode1 = GetForm("jt", 2)
		'TOP:職種小分類
		If GetForm("jt2", 2) <> "" Then JobTypeCode1 = GetForm("jt2", 2)
		'TOP:勤務形態コード
		If GetForm("wt", 2) <> "" Then WorkingTypeCode1 = GetForm("wt", 2)
		'TOP:キーワード
		If GetForm("kw", 2) <> "" Then Keyword = GetForm("kw", 2)
		'特集
		If GetForm("sp", 2) <> "" Then SP = GetForm("sp", 2)

		'沿線検索（パラメータ）
		If GetForm("pc", 2) <> "" Then PC = GetForm("pc", 2)
		If GetForm("rc", 2) <> "" Then RC = GetForm("rc", 2)
		If GetForm("sc", 2) <> "" Then SC = GetForm("sc", 2)

		'情報ツール
		BOC = GetForm("boc", 2)
		If BOC <> "" Then SearchDetailFlag = "1"
	End Sub

	'******************************************************************************
	'概　要：パラメータ名とメンバ変数を紐付けて値を設定する
	'引　数：vKey	：
	'　　　：vValue	：
	'　　　：vFlag	：
	'備　考：
	'更　新：2010/11/06 LIS K.Kokubo
	'******************************************************************************
	Private Sub SetData_ParamPart(ByVal vKey, ByVal vValue)
		If Len(vValue) = 0 Then vValue = GetForm(vKey, 2)

		Select Case vKey
			Case "sdf": SearchDetailFlag = vValue
			Case "sotf": OrderTypeFlag = vValue
			Case "snewf": NewFlag = vValue
			Case "sjtbig1": JobTypeBigCode1 = vValue
			Case "sjt1": JobTypeCode1 = vValue
			Case "sjtbig2": JobTypeBigCode2 = vValue
			Case "sjt2": JobTypeCode2 = vValue
			Case "sjtbig3": JobTypeBigCode3 = vValue
			Case "sjt3": JobTypeCode3 = vValue
			Case "src": RailwayLineCode = Replace(vValue," ","")
			Case "ssc": StationCode = Replace(vValue," ","")
			Case "spc": PrefectureCode = Replace(vValue," ","")
			Case "sct": City = vValue
			Case "swt1": WorkingTypeCode1 = vValue
			Case "swt2": WorkingTypeCode2 = vValue
			Case "swt3": WorkingTypeCode3 = vValue
			Case "sit": IndustryTypeCode = vValue
			Case "ssp01": InexperiencedPersonFlag = vValue
			Case "ssp02": UtilizeLanguageFlag = vValue
			Case "ssp03": TempFlag = vValue
			Case "ssp04": UITurnFlag = vValue
			Case "ssp05": ManyHolidayFlag = vValue
			Case "ssp06": FlexFlag = vValue
			Case "ssp07": NearStationFlag = vValue
			Case "ssp08": NoSmokingFlag = vValue
			Case "ssp09": NewlyBuiltFlag = vValue
			Case "ssp10": LandmarkFlag = vValue
			Case "ssp11": RenovationFlag = vValue
			Case "ssp12": DesignersFlag = vValue
			Case "ssp13": CompanyCafeteriaFlag = vValue
			Case "ssp14": ShortOvertimeFlag = vValue
			Case "ssp15": MaternityFlag = vValue
			Case "ssp16": DressFreeFlag = vValue
			Case "ssp17": MammyFlag = vValue
			Case "ssp18": FixedTimeFlag = vValue
			Case "ssp19": ShortTimeFlag = vValue
			Case "ssp20": HandicappedFlag = vValue
			Case "ssp21": RentAllFlag = vValue
			Case "ssp22": RentPartFlag = vValue
			Case "ssp23": MealsFlag = vValue
			Case "ssp24": MealsAssistanceFlag = vValue
			Case "ssp25": TrainingCostFlag = vValue
			Case "ssp26": EntrepreneurCostFlag = vValue
			Case "ssp27": MoneyFlag = vValue
			Case "ssp28": LandShopFlag = vValue
			Case "ssp29": FindJobFestiveFlag = vValue
			Case "ssp30": AppointmentFlag = vValue
			Case "ssp31": SocietyInsuranceFlag = vValue
			Case "sppf": PercentagePayFlag = vValue
			Case "syimin": YearlyIncomeMin = vValue
			Case "syimax": YearlyIncomeMax = vValue
			Case "smimin": MonthlyIncomeMin = vValue
			Case "smimax": MonthlyIncomeMax = vValue
			Case "sdimin": DailyIncomeMin = vValue
			Case "sdimax": DailyIncomeMax = vValue
			Case "shimin": HourlyIncomeMin = vValue
			Case "shimax": HourlyIncomeMax = vValue
			Case "swsh": WorkStartHour = vValue
			Case "swsm": WorkStartMinute = vValue
			Case "sweh": WorkEndHour = vValue
			Case "swem": WorkEndMinute = vValue
			Case "swht": WeeklyHolidayType = vValue
			Case "sage": Age = vValue
			Case "sstc": SchoolTypeCode = vValue
			Case "sgy": GraduateYear = vValue
			Case "sat": AgreementTerm = vValue
			Case "slocc": LISOrderCompanyCode = vValue
			Case "sos": OSCode = vValue
			Case "sap": ApplicationCode = vValue
			Case "sdl": DevelopmentLanguageCode = vValue
			Case "sdb": DatabaseCode = vValue
			Case "skw": Keyword = vValue
			Case "skwflg": KeywordFlag = vValue
			Case "sst": Specialty = vValue
			Case "poc": PictureOrderCode = vValue
			Case "soc": OrderCode = vValue
			Case "srd": RegistDay = vValue
			Case "snewfkouko": NewKoukokuFlag = vValue
            Case "FeatureFlag": FeatureFlag = vValue
		End Select
	End Sub

	'******************************************************************************
	'概　要：パラメータ文字列からメンバ変数の設定
	'備　考：
	'履　歴：2010/11/06 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function SetData_Param(ByVal vParam)
		Dim idx
		Dim a1,a2

		If Len(vParam) = 0 Then Exit Function
		If Len(vParam) > 1 And Left(vParam,1) = "?" Then vParam = Mid(vParam, 2)

		If InStr(vParam,"&amp;") > 0 Then
			a1 = Split(vParam,"&amp;")
		Else
			a1 = Split(vParam,"&")
		End If

		For idx= LBound(a1) To UBound(a1)
			a2 = Split(a1(idx),"=")
			If UBound(a2) = 1 Then
				Call SetData_ParamPart(a2(0),a2(1))
			End If
		Next

		'<URLエンコードされている文字列をデコード>
		If City <> "" Then City = getURLDecode(HopeCity1,"sjis")
		If City <> "" Then City = getURLDecode(HopeCity2,"sjis")
		If KeyWord <> "" Then KeyWord = getURLDecode(KeyWord,"sjis")
		'</URLエンコードされている文字列をデコード>

		Call SetNames()
	End Function

	'******************************************************************************
	'概　要：コードに対応した名称を取得する
	'引　数：
	'備　考：
	'履　歴：2007/04/04 LIS K.Kokubo 作成
	'******************************************************************************
	Private Sub SetNames()
		Dim sSQL,oRS,flgQE,sError
		Dim idx,aValue,sXML

		'希望職種１
		If IsRE(JobTypeBigCode1, "^\d\d$", True) = True Then
			'大分類
			sSQL = "sp_GetListJobTypeBig '" & JobTypeBigCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeBigName1 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)

			'中分類
			If IsRE(JobTypeCode1, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode1, 2) & "', '" & Mid(JobTypeCode1, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeName1 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		Else
			'中分類のみ
			If IsRE(JobTypeCode1, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode1, 2) & "', '" & Mid(JobTypeCode1, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeBigName1 = ChkStr(oRS.Collect("BigClassName"))
					JobTypeName1 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		End If

		'希望職種２
		If IsRE(JobTypeBigCode2, "^\d\d$", True) = True Then
			'大分類
			sSQL = "sp_GetListJobTypeBig '" & JobTypeBigCode2 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeBigName2 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)

			'中分類
			If IsRE(JobTypeCode2, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode2, 2) & "', '" & Mid(JobTypeCode2, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeName2 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		Else
			'中分類のみ
			If IsRE(JobTypeCode2, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode2, 2) & "', '" & Mid(JobTypeCode2, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeBigName2 = ChkStr(oRS.Collect("BigClassName"))
					JobTypeName2 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		End If

		'希望職種３
		If IsRE(JobTypeBigCode3, "^\d\d$", True) = True Then
			'大分類
			sSQL = "sp_GetListJobTypeBig '" & JobTypeBigCode3 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeBigName3 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)

			'中分類
			If IsRE(JobTypeCode3, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode3, 2) & "', '" & Mid(JobTypeCode3, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeName3 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		Else
			'中分類のみ
			If IsRE(JobTypeCode3, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode3, 2) & "', '" & Mid(JobTypeCode3, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeBigName3 = ChkStr(oRS.Collect("BigClassName"))
					JobTypeName3 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		End If

		'希望沿線
		If RailwayLineCode <> "" Then
			aValue = Split(Replace(RailwayLineCode, " ", ""), ",")

			sXML = ""
			For idx = 0 To UBound(aValue)
				sXML = sXML & "<railwayline><railwaylinecode>" & aValue(idx) & "</railwaylinecode></railwayline>"
			Next
			sXML = "<root>" & sXML & "</root>"

			sSQL = "EXEC up_DtlRailwayLine_XML '" & sXML & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If RailwayLineName <> "" Then RailwayLineName = RailwayLineName & ","
				RailwayLineName = RailwayLineName & ChkStr(oRS.Collect("RailwayLineName2"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
		End If

		'希望駅
		If StationCode <> "" Then
			aValue = Split(Replace(StationCode, " ", ""), ",")

			sXML = ""
			For idx = 0 To UBound(aValue)
				sXML = sXML & "<station><stationcode>" & aValue(idx) & "</stationcode></station>"
			Next
			sXML = "<root>" & sXML & "</root>"

			sSQL = "EXEC up_DtlStation_XML '" & sXML & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If StationName <> "" Then StationName = StationName & ","
				StationName = StationName & ChkStr(oRS.Collect("StationName"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
		End If

		'都道府県
		If PrefectureCode <> "" Then
			aValue = Split(Replace(PrefectureCode, " ", ""), ",")

			sXML = ""
			For idx = 0 To UBound(aValue)
				sXML = sXML & "<prefecture><prefecturecode>" & aValue(idx) & "</prefecturecode></prefecture>"
			Next
			sXML = "<root>" & sXML & "</root>"

			sSQL = "EXEC up_DtlPrefecture_XML '" & sXML & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If PrefectureName <> "" Then PrefectureName = PrefectureName & ","
				PrefectureName = PrefectureName & ChkStr(oRS.Collect("PrefectureName"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
		End If

		'勤務形態１
		If WorkingTypeCode1 <> "" Then
			WorkingTypeName1 = GetDetail("WorkingType", WorkingTypeCode1)
		End If

		'勤務形態２
		If WorkingTypeCode2 <> "" Then
			WorkingTypeName2 = GetDetail("WorkingType", WorkingTypeCode2)
		End If

		'勤務形態３
		If WorkingTypeCode3 <> "" Then
			WorkingTypeName3 = GetDetail("WorkingType", WorkingTypeCode3)
		End If

		'業種
		If IndustryTypeCode <> "" Then
			aValue = Split(Replace(IndustryTypeCode, " ", ""), ",")

			IndustryTypeName = ""
			For idx = 0 To UBound(aValue)
				If IndustryTypeName <> "" Then IndustryTypeName = IndustryTypeName & ","
				IndustryTypeName = IndustryTypeName & GetDetail("IndustryType", aValue(idx))
			Next
		End If

		'週休種類
		If WeeklyHolidayType <> "" Then
			WeeklyHolidayTypeName = GetDetail("WeeklyHolidayType", WeeklyHolidayType)
		End If

		'ＯＳ
		If OSCode <> "" Then
			aValue = Split(Replace(OSCode, " ", ""), ",")
			For idx = 0 To UBound(aValue)
				If OSName <> "" Then OSName = OSName & ","
				OSName = OSName & GetDetail("OS", aValue(idx))
			Next
		End If

		'アプリケーション
		If ApplicationCode <> "" Then
			aValue = Split(Replace(ApplicationCode, " ", ""), ",")
			For idx = 0 To UBound(aValue)
				If ApplicationName <> "" Then ApplicationName = ApplicationName & ","
				ApplicationName = ApplicationName & GetDetail("Application", aValue(idx))
			Next
		End If

		'開発言語
		If DevelopmentLanguageCode <> "" Then
			aValue = Split(Replace(DevelopmentLanguageCode, " ", ""), ",")
			For idx = 0 To UBound(aValue)
				If DevelopmentLanguageName <> "" Then DevelopmentLanguageName = DevelopmentLanguageName & ","
				DevelopmentLanguageName = DevelopmentLanguageName & GetDetail("DevelopmentLanguage", aValue(idx))
			Next
		End If

		'データベース
		If DatabaseCode <> "" Then
			aValue = Split(Replace(DatabaseCode, " ", ""), ",")
			For idx = 0 To UBound(aValue)
				If DatabaseName <> "" Then DatabaseName = DatabaseName & ","
				DatabaseName = DatabaseName & GetDetail("Database", aValue(idx))
			Next
		End If

		'最終学歴
		If SchoolTypeCode <> "" Then
			'sSQL = "EXEC up_DtlSchoolType '" & SchoolType & "';"
			SchoolTypeName = GetDetail("SchoolType", SchoolTypeCode)
		End If

		'資格
		For idx = 0 To LicenseCount - 1
			If LicenseGroupCode(idx) <> "" Then
				'大分類
				sSQL = "sp_GetListLicenseGroup '" & LicenseGroupCode(idx) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					LicenseGroupName(idx) = ChkStr(oRS.Collect("GroupName"))
				End If
				Call RSClose(oRS)

				'中分類
				If IsRE(LicenseCategoryCode(idx), "^\d\d\d$", True) = True Then
					sSQL = "sp_GetListLicenseCategory '" & LicenseGroupCode(idx) & "', '" & LicenseCategoryCode(idx) & "'"
					flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
					If GetRSState(oRS) = True Then
						LicenseCategoryName(idx) = ChkStr(oRS.Collect("CategoryName"))
					End If
					Call RSClose(oRS)

					'小分類
					If IsRE(LicenseCode(idx), "^\d\d$", True) = True Then
						sSQL = "sp_GetListLicenseCode '" & LicenseGroupCode(idx) & "', '" & LicenseCategoryCode(idx) & "', '" & LicenseCode(idx) & "'"
						flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
						If GetRSState(oRS) = True Then
							LicenseName(idx) = ChkStr(oRS.Collect("Name"))
						End If
						Call RSClose(oRS)
					End If
				End If
			End If
		Next
	End Sub

	'******************************************************************************
	'概　要：全角数字を半角数字に変換
	'引　数：
	'備　考：
	'履　歴：2009/11/18 LIS K.Kokubo 作成
	'******************************************************************************
	Private Function ChgZenNum(ByVal vNum)
		Dim sChg
		ChgZenNum = ""

		sChg = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(CStr(vNum),"０","0"),"１","1"),"２","2"),"３","3"),"４","4"),"５","5"),"６","6"),"７","7"),"８","8"),"９","9")
		If IsRE(sChg,0,False) = False Then Exit Function
		ChgZenNum = sChg
	End Function

	'******************************************************************************
	'概　要：データの整合性をチェック
	'引　数：
	'備　考：
	'履　歴：2007/04/04 LIS K.Kokubo 作成
	'******************************************************************************
	Private Sub ChkData()
		Dim aValue
		Dim idx
		Dim tmp

		'おしごとチェッカー対応
		If IsRE(JobTypeCode1, "^\d\d$", True) = True Then
			JobTypeBigCode1 = JobTypeCode1
			JobTypeCode1 = ""
		End If

		'希望業種カンマ区切り
		IndustryTypeCode = Replace(IndustryTypeCode, " ", "")

		'就業開始時間
		If WorkStartHour <> "" Then
			If WorkStartMinute = "" Then WorkStartMinute = "00"
		ElseIf WorkStartMinute <> "" Then
			WorkStartMinute = ""
		End If

		'就業終了時間
		If WorkEndHour <> "" Then
			If WorkEndMinute = "" Then WorkEndMinute = "00"
		ElseIf WorkEndMinute <> "" Then
			WorkEndMinute = ""
		End If

		If JobTypeCode1 <> "" Then JobTypeBigCode1 = Left(JobTypeCode1, 2)
		If JobTypeCode2 <> "" Then JobTypeBigCode2 = Left(JobTypeCode2, 2)
		If JobTypeCode3 <> "" Then JobTypeBigCode3 = Left(JobTypeCode3, 2)

		'年齢
		'Age = ChgZenNum(Replace(Age,",",""))

		'卒業年
		If IsNumber(GraduateYear,0,False) Then
			If CInt(GraduateYear) < 1900 Or CInt(GraduateYear) > 2099 Then GraduateYear = ""
		End If

		'給与
		YearlyIncomeMin = Replace(ChgZenNum(Replace(YearlyIncomeMin,",","")),"万","0000")
		YearlyIncomeMax = Replace(ChgZenNum(Replace(YearlyIncomeMax,",","")),"万","0000")
		MonthlyIncomeMin = ChgZenNum(Replace(MonthlyIncomeMin,",",""))
		MonthlyIncomeMax = ChgZenNum(Replace(MonthlyIncomeMax,",",""))
		DailyIncomeMin = ChgZenNum(Replace(DailyIncomeMin,",",""))
		DailyIncomeMax = ChgZenNum(Replace(DailyIncomeMax,",",""))
		HourlyIncomeMin = ChgZenNum(Replace(HourlyIncomeMin,",",""))
		HourlyIncomeMax = ChgZenNum(Replace(HourlyIncomeMax,",",""))
		WorkStartHour = ChgZenNum(WorkStartHour)
		WorkEndHour = ChgZenNum(WorkEndHour)

		'情報コードCSV
		If InStr(OrderCode, ",") > 0 Then
			tmp = ""
			aValue = Split(OrderCode, ",")
			For idx = 0 To UBound(aValue)
				If tmp <> "" Then tmp = tmp & ","
				tmp = tmp & aValue(idx)
			Next
			OrderCode = tmp
		End If

		'特徴ビット文字列
		Specialty = ""
		If InexperiencedPersonFlag & UtilizeLanguageFlag & TempFlag & UITurnFlag & ManyHolidayFlag & FlexFlag & _
		NearStationFlag & NoSmokingFlag & NewlyBuiltFlag & LandmarkFlag & RenovationFlag & DesignersFlag & _
		CompanyCafeteriaFlag & ShortOvertimeFlag & MaternityFlag & DressFreeFlag & MammyFlag & FixedTimeFlag & _
		ShortTimeFlag & HandicappedFlag & RentAllFlag & RentPartFlag & MealsFlag & MealsAssistanceFlag & _
		TrainingCostFlag & EntrepreneurCostFlag & MoneyFlag & LandShopFlag & FindJobFestiveFlag & AppointmentFlag & SocietyInsuranceFlag <> "" Then
			If InexperiencedPersonFlag <> "" Then: Specialty = Specialty & InexperiencedPersonFlag: Else: Specialty = Specialty & "0": End If
			If UtilizeLanguageFlag <> "" Then: Specialty = Specialty & UtilizeLanguageFlag: Else: Specialty = Specialty & "0": End If
			If TempFlag <> "" Then: Specialty = Specialty & TempFlag: Else: Specialty = Specialty & "0": End If
			If UITurnFlag <> "" Then: Specialty = Specialty & UITurnFlag: Else: Specialty = Specialty & "0": End If
			If ManyHolidayFlag <> "" Then: Specialty = Specialty & ManyHolidayFlag: Else: Specialty = Specialty & "0": End If
			If FlexFlag <> "" Then: Specialty = Specialty & FlexFlag: Else: Specialty = Specialty & "0": End If
			If NearStationFlag <> "" Then: Specialty = Specialty & NearStationFlag: Else: Specialty = Specialty & "0": End If
			If NoSmokingFlag <> "" Then: Specialty = Specialty & NoSmokingFlag: Else: Specialty = Specialty & "0": End If
			If NewlyBuiltFlag <> "" Then: Specialty = Specialty & NewlyBuiltFlag: Else: Specialty = Specialty & "0": End If
			If LandmarkFlag <> "" Then: Specialty = Specialty & LandmarkFlag: Else: Specialty = Specialty & "0": End If
			If RenovationFlag <> "" Then: Specialty = Specialty & RenovationFlag: Else: Specialty = Specialty & "0": End If
			If DesignersFlag <> "" Then: Specialty = Specialty & DesignersFlag: Else: Specialty = Specialty & "0": End If
			If CompanyCafeteriaFlag <> "" Then: Specialty = Specialty & CompanyCafeteriaFlag: Else: Specialty = Specialty & "0": End If
			If ShortOvertimeFlag <> "" Then: Specialty = Specialty & ShortOvertimeFlag: Else: Specialty = Specialty & "0": End If
			If MaternityFlag <> "" Then: Specialty = Specialty & MaternityFlag: Else: Specialty = Specialty & "0": End If
			If DressFreeFlag <> "" Then: Specialty = Specialty & DressFreeFlag: Else: Specialty = Specialty & "0": End If
			If MammyFlag <> "" Then: Specialty = Specialty & MammyFlag: Else: Specialty = Specialty & "0": End If
			If FixedTimeFlag <> "" Then: Specialty = Specialty & FixedTimeFlag: Else: Specialty = Specialty & "0": End If
			If ShortTimeFlag <> "" Then: Specialty = Specialty & ShortTimeFlag: Else: Specialty = Specialty & "0": End If
			If HandicappedFlag <> "" Then: Specialty = Specialty & HandicappedFlag: Else: Specialty = Specialty & "0": End If
			If RentAllFlag <> "" Then: Specialty = Specialty & RentAllFlag: Else: Specialty = Specialty & "0": End If
			If RentPartFlag <> "" Then: Specialty = Specialty & RentPartFlag: Else: Specialty = Specialty & "0": End If
			If MealsFlag <> "" Then: Specialty = Specialty & MealsFlag: Else: Specialty = Specialty & "0": End If
			If MealsAssistanceFlag <> "" Then: Specialty = Specialty & MealsAssistanceFlag: Else: Specialty = Specialty & "0": End If
			If TrainingCostFlag <> "" Then: Specialty = Specialty & TrainingCostFlag: Else: Specialty = Specialty & "0": End If
			If EntrepreneurCostFlag <> "" Then: Specialty = Specialty & EntrepreneurCostFlag: Else: Specialty = Specialty & "0": End If
			If MoneyFlag <> "" Then: Specialty = Specialty & MoneyFlag: Else: Specialty = Specialty & "0": End If
			If LandShopFlag <> "" Then: Specialty = Specialty & LandShopFlag: Else: Specialty = Specialty & "0": End If
			If FindJobFestiveFlag <> "" Then: Specialty = Specialty & FindJobFestiveFlag: Else: Specialty = Specialty & "0": End If
			If AppointmentFlag <> "" Then: Specialty = Specialty & AppointmentFlag: Else: Specialty = Specialty & "0": End If
			If SocietyInsuranceFlag <> "" Then: Specialty = Specialty & SocietyInsuranceFlag: Else: Specialty = Specialty & "0": End If
		End If

		If JT <> "" Then JobTypeCode1 = JT
		If JT2 <> "" Then JobTypeCode1 = JT2
		If WT <> "" Then WorkingTypeCode1 = WT
		If KW <> "" Then Keyword = KW

		If PC <> "" Then PrefectureCode = PC
		If RC <> "" Then RailwayLineCode1 = RC
		If SC <> "" Then StationCode = SC
	End Sub

	'******************************************************************************
	'概　要：お仕事詳細検索ページへ渡すGETパラメータを生成して取得。
	'引　数：
	'備　考：■制限
	'　　　：パラメータを含むURLは、IEの制限が2048文字までであるので、それに合わせる。
	'履　歴：2007/04/04 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function GetSearchParam()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

		Dim sParam
		Dim idx

		GetSearchParam = ""

		If SearchDetailFlag <> "" Then sParam = sParam & "&sdf=" & SearchDetailFlag
		If OrderTypeFlag <> "" Then sParam = sParam & "&sotf=" & OrderTypeFlag
		If NewFlag <> "" Then sParam = sParam & "&snewf=" & NewFlag
		If JobTypeBigCode1 <> "" Then sParam = sParam & "&sjtbig1=" & JobTypeBigCode1
		If JobTypeCode1 <> "" Then sParam = sParam & "&sjt1=" & JobTypeCode1
		If JobTypeBigCode2 <> "" Then sParam = sParam & "&sjtbig2=" & JobTypeBigCode2
		If JobTypeCode2 <> "" Then sParam = sParam & "&sjt2=" & JobTypeCode2
		If JobTypeBigCode3 <> "" Then sParam = sParam & "&sjtbig3=" & JobTypeBigCode3
		If JobTypeCode3 <> "" Then sParam = sParam & "&sjt3=" & JobTypeCode3
		If RailwayLineCode <> "" Then sParam = sParam & "&src=" & RailwayLineCode
		If StationCode <> "" Then sParam = sParam & "&ssc=" & StationCode
		If PrefectureCode <> "" Then sParam = sParam & "&spc=" & PrefectureCode
		If City <> "" Then sParam = sParam & "&sct=" & Server.URLEncode(City)
		If WorkingTypeCode1 <> "" Then sParam = sParam & "&swt1=" & WorkingTypeCode1
		If WorkingTypeCode2 <> "" Then sParam = sParam & "&swt2=" & WorkingTypeCode2
		If WorkingTypeCode3 <> "" Then sParam = sParam & "&swt3=" & WorkingTypeCode3
		If IndustryTypeCode <> "" Then sParam = sParam & "&sit=" & IndustryTypeCode
		If InexperiencedPersonFlag <> "" Then sParam = sParam & "&ssp01=" & InexperiencedPersonFlag
		If UtilizeLanguageFlag <> "" Then sParam = sParam & "&ssp02=" & UtilizeLanguageFlag
		If TempFlag <> "" Then sParam = sParam & "&ssp03=" & TempFlag
		If UITurnFlag <> "" Then sParam = sParam & "&ssp04=" & UITurnFlag
		If ManyHolidayFlag <> "" Then sParam = sParam & "&ssp05=" & ManyHolidayFlag
		If FlexFlag <> "" Then sParam = sParam & "&ssp06=" & FlexFlag
		If NearStationFlag <> "" Then sParam = sParam & "&ssp07=" & NearStationFlag
		If NoSmokingFlag <> "" Then sParam = sParam & "&ssp08=" & NoSmokingFlag
		If NewlyBuiltFlag <> "" Then sParam = sParam & "&ssp09=" & NewlyBuiltFlag
		If LandmarkFlag <> "" Then sParam = sParam & "&ssp10=" & LandmarkFlag
		If RenovationFlag <> "" Then sParam = sParam & "&ssp11=" & RenovationFlag
		If DesignersFlag <> "" Then sParam = sParam & "&ssp12=" & DesignersFlag
		If CompanyCafeteriaFlag <> "" Then sParam = sParam & "&ssp13=" & CompanyCafeteriaFlag
		If ShortOvertimeFlag <> "" Then sParam = sParam & "&ssp14=" & ShortOvertimeFlag
		If MaternityFlag <> "" Then sParam = sParam & "&ssp15=" & MaternityFlag
		If DressFreeFlag <> "" Then sParam = sParam & "&ssp16=" & DressFreeFlag
		If MammyFlag <> "" Then sParam = sParam & "&ssp17=" & MammyFlag
		If FixedTimeFlag <> "" Then sParam = sParam & "&ssp18=" & FixedTimeFlag
		If ShortTimeFlag <> "" Then sParam = sParam & "&ssp19=" & ShortTimeFlag
		If HandicappedFlag <> "" Then sParam = sParam & "&ssp20=" & HandicappedFlag
		If RentAllFlag <> "" Then sParam = sParam & "&ssp21=" & RentAllFlag
		If RentPartFlag <> "" Then sParam = sParam & "&ssp22=" & RentPartFlag
		If MealsFlag <> "" Then sParam = sParam & "&ssp23=" & MealsFlag
		If MealsAssistanceFlag <> "" Then sParam = sParam & "&ssp24=" & MealsAssistanceFlag
		If TrainingCostFlag <> "" Then sParam = sParam & "&ssp25=" & TrainingCostFlag
		If EntrepreneurCostFlag <> "" Then sParam = sParam & "&ssp26=" & EntrepreneurCostFlag
		If MoneyFlag <> "" Then sParam = sParam & "&ssp27=" & MoneyFlag
		If LandShopFlag <> "" Then sParam = sParam & "&ssp28=" & LandShopFlag
		If FindJobFestiveFlag <> "" Then sParam = sParam & "&ssp29=" & FindJobFestiveFlag
		If AppointmentFlag <> "" Then sParam = sParam & "&ssp30=" & AppointmentFlag
		If SocietyInsuranceFlag <> "" Then sParam = sParam & "&ssp31=" & SocietyInsuranceFlag
		If PercentagePayFlag <> "" Then sParam = sParam & "&sppf=" & PercentagePayFlag
		If YearlyIncomeMin <> "" Then sParam = sParam & "&syimin=" & YearlyIncomeMin
		If YearlyIncomeMax <> "" Then sParam = sParam & "&syimax=" & YearlyIncomeMax
		If MonthlyIncomeMin <> "" Then sParam = sParam & "&smimin=" & MonthlyIncomeMin
		If MonthlyIncomeMax <> "" Then sParam = sParam & "&smimax=" & MonthlyIncomeMax
		If DailyIncomeMin <> "" Then sParam = sParam & "&sdimin=" & DailyIncomeMin
		If DailyIncomeMax <> "" Then sParam = sParam & "&sdimax=" & DailyIncomeMax
		If HourlyIncomeMin <> "" Then sParam = sParam & "&shimin=" & HourlyIncomeMin
		If HourlyIncomeMax <> "" Then sParam = sParam & "&shimax=" & HourlyIncomeMax
		If WorkStartHour <> "" Then sParam = sParam & "&swsh=" & WorkStartHour
		If WorkStartMinute <> "" Then sParam = sParam & "&swsm=" & WorkStartMinute
		If WorkEndHour <> "" Then sParam = sParam & "&sweh=" & WorkEndHour
		If WorkEndMinute <> "" Then sParam = sParam & "&swem=" & WorkEndMinute
		If WeeklyHolidayType <> "" Then sParam = sParam & "&swht=" & WeeklyHolidayType
		'If Age <> "" Then sParam = sParam & "&sage=" & Age
		If SchoolTypeCode <> "" Then sParam = sParam & "&sstc=" & SchoolTypeCode
		If GraduateYear <> "" Then sParam = sParam & "&sgy=" & GraduateYear
		If AgreementTerm <> "" Then sParam = sParam & "&sat=" & AgreementTerm
		If NewKoukokuFlag <> "" Then sParam = sParam & "&snewfkouko=" & NewKoukokuFlag
        If FeatureFlag <> "" Then sParam = sParam & "&FeatureFlag=" & FeatureFlag

		For idx = 0 To LicenseCount - 1
			If LicenseGroupCode(idx) <> "" Then
				sParam = sParam & "&slg"&idx+1 & "=" & LicenseGroupCode(idx)
				sParam = sParam & "&slc"&idx+1 & "=" & LicenseCategoryCode(idx)
				sParam = sParam & "&sl"&idx+1 & "=" & LicenseCode(idx)
			End If
		Next

		If OSCode <> "" Then sParam = sParam & "&sos=" & OSCode
		If ApplicationCode <> "" Then sParam = sParam & "&sap=" & ApplicationCode
		If DevelopmentLanguageCode <> "" Then sParam = sParam & "&sdl=" & DevelopmentLanguageCode
		If DatabaseCode <> "" Then sParam = sParam & "&sdb=" & DatabaseCode
		If Keyword <> "" Then sParam = sParam & "&skw=" & Server.URLEncode(Keyword)
		If KeywordFlag <> "" Then sParam = sParam & "&skwflg=" & KeywordFlag
		If PictureOrderCode <> "" Then sParam = sParam & "&poc=" & PictureOrderCode
		If OrderCode <> "" Then sParam = sParam & "&soc=" & OrderCode
		If Specialty <> "" Then sParam = sParam & "&sst=" & Specialty
		If SP <> "" Then sParam = sParam & "&sp=" & SP
		If RegistDay <> "" Then sParam = sParam & "&srd=" & RegistDay
		If LISOrderCompanyCode <> "" Then sParam = sParam & "&slocc=" & LISOrderCompanyCode

		If sParam <> "" Then
			'頭の&を？に変換
			sParam = "?" & Mid(sParam, 2)

			'ＩＥの仕様はパラメータの上限が２０４８バイト
			sParam = Left(sParam, 2048)
		End If

		GetSearchParam = Replace(sParam, "&", "&amp;")
	End Function

	'******************************************************************************
	'概　要：求人票詳細検索ＳＱＬを取得
	'引　数：
	'備　考：
	'履　歴：2007/04/04 LIS K.Kokubo 作成
	'******************************************************************************
	Function GetSQLOrderSearchDetail()
		Dim sSQL

		Dim sJoin
		Dim sWhere
		Dim sDeclare
		Dim sParams
		Dim iParamNo
		Dim iParamNo2
		Dim sFrom
		Dim sTemp
		Dim sTemp2
		Dim sTemp3
		Dim aValue
		Dim idx
		Dim sSearchCondition

		sJoin = ""
		sWhere = ""
		sDeclare = ""
		sParams = ""

		'データ整合性チェック
		Call ChkData()

		'******************************************************************************
		'社内外案件検索フラグ start
		'------------------------------------------------------------------------------
		If OrderTypeFlag <> "" Then
			If OrderTypeFlag = "0" Then
				'一般求人広告
				If sWhere <> "" Then sWhere = sWhere & "AND "
				sWhere = sWhere & "VWOC.OrderType = '0'" & vbCrLf
			ElseIf OrderTypeFlag = "1" Then
				'社内案件
				If sWhere <> "" Then sWhere = sWhere & "AND "
				sWhere = sWhere & "VWOC.OrderType > '0'" & vbCrLf
			End If
		End If
		'------------------------------------------------------------------------------
		'社内外案件検索フラグ end
		'******************************************************************************

		'******************************************************************************
		'新着フラグ start
		'------------------------------------------------------------------------------
		If NewFlag = "1" Then
			If sWhere <> "" Then sWhere = sWhere & "AND "
			sWhere = sWhere & "CONVERT(VARCHAR(8), VWOC.RegistDay, 112) >= DATEADD(DAY,  -9, CONVERT(DATETIME, CONVERT(VARCHAR(8), GETDATE(), 112))) "
			'ライセンスの掲載開始日も考慮するバージョンはコメントアウト
			'sWhere = sWhere & "(CASE WHEN VWOC.OrderType = '0' AND VWOC.RegistDay < VWOC.RiyoFromDate THEN VWOC.RiyoFromDate ELSE CONVERT(DATETIME, CONVERT(VARCHAR(8), VWOC.RegistDay, 112)) END) >= DATEADD(DAY, -6, CONVERT(DATETIME, CONVERT(VARCHAR(8), GETDATE(), 112))) "
		End If
		'------------------------------------------------------------------------------
		'社内外案件検索フラグ end
		'******************************************************************************
        '******************************************************************************
		'新着フラグ start
		'------------------------------------------------------------------------------
		If NewKoukokuFlag = "1" Then
			If sWhere <> "" Then sWhere = sWhere & "AND "
			sWhere = sWhere & "VWOC.OrderType = '0' AND (CONVERT(VARCHAR(8), VWOC.RegistDay, 112) >= DATEADD(DAY,  -30, CONVERT(DATETIME, CONVERT(VARCHAR(8), GETDATE(), 112)))) "
		End If
		'------------------------------------------------------------------------------
		'新着フラグ end
		'******************************************************************************

		'******************************************************************************
		'職種 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If JobTypeBigCode1 & JobTypeCode1 & JobTypeBigCode2 & JobTypeCode2 & JobTypeBigCode3 & JobTypeCode3 <> "" Then
			If JobTypeBigCode1 & JobTypeCode1 <> "" Then
				sTemp = JobTypeBigCode1
				If JobTypeCode1 <> "" Then sTemp = JobTypeCode1

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "A.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			If JobTypeBigCode2 & JobTypeCode2 <> "" Then
				sTemp = JobTypeBigCode2
				If JobTypeCode2 <> "" Then sTemp = JobTypeCode2

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "A.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			If JobTypeBigCode3 & JobTypeCode3 <> "" Then
				sTemp = JobTypeBigCode3
				If JobTypeCode3 <> "" Then sTemp = JobTypeCode3

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "A.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_JobType AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE (" & RTrim(sTemp2) & ") "
			sJoin = sJoin & ") AS CJT ON VWOC.OrderCode = CJT.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'職種 end
		'******************************************************************************

		'******************************************************************************
		'沿線 start
		'------------------------------------------------------------------------------
		sTemp = ""

		If RailwayLineCode <> "" Then
			aValue = Split(Replace(RailwayLineCode, " ", ""), ",")
			For iParamNo = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vRailwayLineCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vRailwayLineCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vRailwayLineCode" & iParamNo
			Next

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS A WITH(NOLOCK) "
			sJoin = sJoin & "INNER JOIN StationStop AS B WITH(NOLOCK) "
			sJoin = sJoin & "ON A.StationCode = B.StationCode "
			sJoin = sJoin & "INNER JOIN B_RailwayLine AS C WITH(NOLOCK) "
			sJoin = sJoin & "ON B.RailwayLineCode = C.RailwayLineCode "
			sJoin = sJoin & "AND C.RailwayLineCode IN (" & RTrim(sTemp) & ")"
			sJoin = sJoin & ") AS CRL "
			sJoin = sJoin & "ON VWOC.OrderCode = CRL.OrderCode" & vbCrLf
		End If

		'------------------------------------------------------------------------------
		'沿線 end
		'******************************************************************************

		'******************************************************************************
		'駅 start
		'------------------------------------------------------------------------------
		sTemp = ""

		If StationCode <> "" Then
			aValue = Split(Replace(StationCode, " ", ""), ",")
			For iParamNo = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vStationCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vStationCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vStationCode" & iParamNo
			Next

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE A.StationCode IN (" & sTemp & ")"
			sJoin = sJoin & ") AS CNS "
			sJoin = sJoin & "ON VWOC.OrderCode = CNS.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'駅 end
		'******************************************************************************

		'******************************************************************************
		'希望勤務地 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""

		If PrefectureCode <> "" Or City <> "" Then
			If PrefectureCode <> "" Then
				aValue = Split(Replace(PrefectureCode, " ", ""), ",")
				For iParamNo = LBound(aValue) To UBound(aValue)
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vPrefectureCode" & iParamNo & " VARCHAR(3)"
					sParams = sParams & ",@vPrefectureCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

					If sTemp <> "" Then sTemp = sTemp & ","
					sTemp = sTemp & "@vPrefectureCode" & iParamNo
				Next

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "A.PrefectureCode IN (" & sTemp & ") "
			End If

			If City <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vCity VARCHAR(200)"
				sParams = sParams & ",@vCity = N'" & City & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "A.City LIKE '%' + @vCity + '%' "
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_WorkingPlace AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE " & RTrim(sTemp2)
			sJoin = sJoin & ") AS CWP "
			sJoin = sJoin & "ON VWOC.OrderCode = CWP.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'希望勤務地 end
		'******************************************************************************

		'******************************************************************************
		'希望勤務形態 start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If WorkingTypeCode1 & WorkingTypeCode2 & WorkingTypeCode3 <> "" Then
			If WorkingTypeCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkingTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vWorkingTypeCode" & iParamNo & " = N'" & WorkingTypeCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vWorkingTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If WorkingTypeCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkingTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vWorkingTypeCode" & iParamNo & " = N'" & WorkingTypeCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vWorkingTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If WorkingTypeCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkingTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vWorkingTypeCode" & iParamNo & " = N'" & WorkingTypeCode3 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vWorkingTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_WorkingType AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE A.WorkingTypeCode IN (" & RTrim(sTemp) & ") "
			sJoin = sJoin & ") AS CWT "
			sJoin = sJoin & "ON VWOC.OrderCode = CWT.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'希望勤務形態 end
		'******************************************************************************

		'******************************************************************************
		'希望業種 start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If IndustryTypeCode <> "" Then
			aValue = Split(Replace(IndustryTypeCode, " ", ""), ",")
			For iParamNo = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vIndustryTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vIndustryTypeCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vIndustryTypeCode" & iParamNo
			Next

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT A.CompanyCode "
			sJoin = sJoin & "FROM CompanyInfo AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE A.IndustryType IN (" & RTrim(sTemp) & ") "
			sJoin = sJoin & ") AS CIDST "
			sJoin = sJoin & "ON VWOC.CompanyCode = CIDST.CompanyCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'希望業種 end
		'******************************************************************************

		'******************************************************************************
		'特徴 start
		'------------------------------------------------------------------------------
		'未経験歓迎、語学を活かす、UIターン、休日１２０日以上
		sTemp = ""

		If InexperiencedPersonFlag = "1" Or UtilizeLanguageFlag = "1" Or UITurnFlag = "1" Or ManyHolidayFlag = "1" Or _
		FlexFlag = "1" Or NearStationFlag = "1" Or NoSmokingFlag = "1" Or NewlyBuiltFlag = "1" Or LandmarkFlag = "1" Or _
		RenovationFlag = "1" Or DesignersFlag = "1" Or CompanyCafeteriaFlag = "1" Or ShortOvertimeFlag = "1" Or MaternityFlag = "1" Or _
		DressFreeFlag = "1" Or MammyFlag = "1" Or FixedTimeFlag = "1" Or ShortTimeFlag = "1" Or HandicappedFlag = "1" Or RentAllFlag = "1" Or _
		RentPartFlag = "1" Or MealsFlag = "1" Or MealsAssistanceFlag = "1" Or TrainingCostFlag = "1" Or EntrepreneurCostFlag = "1" Or _
		MoneyFlag = "1" Or LandShopFlag = "1" Or FindJobFestiveFlag = "1" Or AppointmentFlag = "1" Or SocietyInsuranceFlag = "1" Then
			'未経験歓迎
			If InexperiencedPersonFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.InexperiencedPersonFlag = '1' "
			End If

			'語学を活かす
			If UtilizeLanguageFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.UtilizeLanguageFlag = '1' "
			End If

			'UIターン
			If UITurnFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.UITurnFlag = '1' "
			End If

			'休日１２０日以上
			If ManyHolidayFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.ManyHolidayFlag = '1' "
			End If
			'フレックスタイム
			If FlexFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.FlexTimeFlag = '1' "
			End If
			'駅近
			If NearStationFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.NearStationFlag = '1' "
			End If
			'禁煙・分煙
			If NoSmokingFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.NoSmokingFlag = '1' "
			End If
			'新築ビル・オフィス
			If NewlyBuiltFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.NewlyBuiltFlag = '1' "
			End If
			'高層ビル
			If LandmarkFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.LandmarkFlag = '1' "
			End If
			'リノベーション
			If RenovationFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.RenovationFlag = '1' "
			End If
			'デザイナーズ
			If DesignersFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.DesignersFlag = '1' "
			End If
			'社員食堂
			If CompanyCafeteriaFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.CompanyCafeteriaFlag = '1' "
			End If
			'短時間残業
			If ShortOvertimeFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.ShortOvertimeFlag = '1' "
			End If
			'産休・育休
			If MaternityFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MaternityFlag = '1' "
			End If
			'服装自由
			If DressFreeFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.DressFreeFlag = '1' "
			End If
			'ママ歓迎
			If MammyFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MammyFlag = '1' "
			End If
			'18時までに退社
			If FixedTimeFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.FixedTimeFlag = '1' "
			End If
			'短時間労働
			If ShortTimeFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.ShortTimeFlag = '1' "
			End If
			'障害者歓迎
			If HandicappedFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.HandicappedFlag = '1' "
			End If
			'住宅費用全額補助あり
			If RentAllFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.RentAllFlag = '1' "
			End If
			'住宅費用一部補助あり
			If RentPartFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.RentPartFlag = '1' "
			End If
			'食事・賄い付き案件
			If MealsFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MealsFlag = '1' "
			End If
			'食事補助制度あり
			If MealsAssistanceFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MealsAssistanceFlag = '1' "
			End If
			'研修費助成制度あり
			If TrainingCostFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.TrainingCostFlag = '1' "
			End If
			'起業機材補助制度あり
			If EntrepreneurCostFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.EntrepreneurCostFlag = '1' "
			End If
			'無利子・低利子補助制度あり
			If MoneyFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MoneyFlag = '1' "
			End If
			'土地・店舗等提供制度あり
			If LandShopFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.LandShopFlag = '1' "
			End If
			'就職お祝い金制度あり
			If FindJobFestiveFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.FindJobFestiveFlag = '1' "
			End If
			'正社員登用制度あり
			If AppointmentFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.AppointmentFlag = '1' "
			End If
			'社保完備
			If SocietyInsuranceFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "EXISTS(SELECT * FROM C_Info AS A WHERE CSP.OrderCode = A.OrderCode AND EXISTS(SELECT * FROM CompanyInfo AS B WHERE A.CompanyCode = B.CompanyCode AND B.SocietyInsurance = 'ON') AND EXISTS(SELECT * FROM C_WorkingType AS C WHERE A.OrderCode = C.OrderCode AND C.WorkingTypeCode <= '005') AND NOT EXISTS(SELECT * FROM C_WorkingType AS D WHERE A.OrderCode = D.OrderCode AND D.WorkingTypeCode IN ('006','007'))) "
			End If

			sJoin = sJoin & "INNER JOIN C_SupplementInfo AS CSP WITH(NOLOCK) ON VWOC.OrderCode = CSP.OrderCode AND " & RTrim(sTemp) & vbCrLf
		End If

'		'派遣
'		If TempFlag = "1" Then
'			If InStr(sJoin, "INNER JOIN C_WorkingType AS CWT") = 0 Then sJoin = sJoin & "INNER JOIN C_WorkingType AS CWT ON CI.OrderCode = CWT.OrderCode "
'			If sWhere <> "" Then sWhere = sWhere & "AND "
'			sWhere = sWhere & "CWT.WorkingTypeCode IN ('001', '004') " & vbCrLf
'		End If
		'------------------------------------------------------------------------------
		'特徴 end
		'******************************************************************************

		'******************************************************************************
		'給与 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		If YearlyIncomeMin & YearlyIncomeMax & MonthlyIncomeMin & MonthlyIncomeMax & DailyIncomeMin & DailyIncomeMax & HourlyIncomeMin & HourlyIncomeMax & PercentagePayFlag <> "" Then
			'<年収>
			If YearlyIncomeMin & YearlyIncomeMax <> "" Then
				If YearlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vYearlyIncomeMin INT"
					sParams = sParams & ",@vYearlyIncomeMin = " & YearlyIncomeMin
				End If

				If YearlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vYearlyIncomeMax INT"
					sParams = sParams & ",@vYearlyIncomeMax = " & YearlyIncomeMax
				End If

				If sTemp <> "" Then sTemp = sTemp & "OR "
				If YearlyIncomeMin <> "" And YearlyIncomeMax <> "" Then
					'年収下限,上限両方の入力がある場合
					sTemp = sTemp & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND (A.YearlyIncomeMin BETWEEN @vYearlyIncomeMin AND @vYearlyIncomeMax)) OR (COALESCE(A.YearlyIncomeMax, 0) > 0 AND (A.YearlyIncomeMax BETWEEN @vYearlyIncomeMin AND @vYearlyIncomeMax))) "
				ElseIf YearlyIncomeMin <> "" Then
					'年収下限のみ入力がある場合
					sTemp = sTemp & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND A.YearlyIncomeMin >= @vYearlyIncomeMin) OR (COALESCE(A.YearlyIncomeMax, 0) > 0 AND A.YearlyIncomeMax >= @vYearlyIncomeMin)) "
				ElseIf YearlyIncomeMax <> "" Then
					'年収上限のみ入力がある場合
					sTemp = sTemp & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND A.YearlyIncomeMin <= @vYearlyIncomeMax) OR (COALESCE(A.YearlyIncomeMin, 0) = 0 AND COALESCE(A.YearlyIncomeMax, 0) > 0 AND A.YearlyIncomeMax <= @vYearlyIncomeMax)) "
				End If
			End If
			'</年収>

			'<月給>
			If MonthlyIncomeMin & MonthlyIncomeMax <> "" Then
				If MonthlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vMonthlyIncomeMin INT"
					sParams = sParams & ",@vMonthlyIncomeMin = " & MonthlyIncomeMin
				End If

				If MonthlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vMonthlyIncomeMax INT"
					sParams = sParams & ",@vMonthlyIncomeMax = " & MonthlyIncomeMax
				End If

				If sTemp <> "" Then sTemp = sTemp & "OR "
				If MonthlyIncomeMin <> "" And MonthlyIncomeMax <> "" Then
					'月給下限,上限両方の入力がある場合
					sTemp = sTemp & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND (A.MonthlyIncomeMin BETWEEN @vMonthlyIncomeMin AND @vMonthlyIncomeMax)) OR (COALESCE(A.MonthlyIncomeMax, 0) > 0 AND (A.MonthlyIncomeMax BETWEEN @vMonthlyIncomeMin AND @vMonthlyIncomeMax))) "
				ElseIf MonthlyIncomeMin <> "" Then
					'月給下限のみ入力がある場合
					sTemp = sTemp & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND A.MonthlyIncomeMin >= @vMonthlyIncomeMin) OR (COALESCE(A.MonthlyIncomeMax, 0) > 0 AND A.MonthlyIncomeMax >= @vMonthlyIncomeMin)) "
				ElseIf MonthlyIncomeMax <> "" Then
					'月給上限のみ入力がある場合
					sTemp = sTemp & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND A.MonthlyIncomeMin <= @vMonthlyIncomeMax) OR (COALESCE(A.MonthlyIncomeMin, 0) = 0 AND COALESCE(A.MonthlyIncomeMax, 0) > 0 AND A.MonthlyIncomeMax <= @vMonthlyIncomeMax)) "
				End If
			End If
			'</月給>

			'<日給>
			If DailyIncomeMin & DailyIncomeMax <> "" Then
				If DailyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDailyIncomeMin INT"
					sParams = sParams & ",@vDailyIncomeMin = " & DailyIncomeMin
				End If

				If DailyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDailyIncomeMax INT"
					sParams = sParams & ",@vDailyIncomeMax = " & DailyIncomeMax
				End If

				If sTemp <> "" Then sTemp = sTemp & "OR "
				If DailyIncomeMin <> "" And DailyIncomeMax <> "" Then
					'日給下限,上限両方の入力がある場合
					sTemp = sTemp & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND (A.DailyIncomeMin BETWEEN @vDailyIncomeMin AND @vDailyIncomeMax)) OR (COALESCE(A.DailyIncomeMax, 0) > 0 AND (A.DailyIncomeMax BETWEEN @vDailyIncomeMin AND @vDailyIncomeMax))) "
				ElseIf DailyIncomeMin <> "" Then
					'日給下限のみ入力がある場合
					sTemp = sTemp & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND A.DailyIncomeMin >= @vDailyIncomeMin) OR (COALESCE(A.DailyIncomeMax, 0) > 0 AND A.DailyIncomeMax >= @vDailyIncomeMin)) "
				ElseIf DailyIncomeMax <> "" Then
					'日給上限のみ入力がある場合
					sTemp = sTemp & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND A.DailyIncomeMin <= @vDailyIncomeMax) OR (COALESCE(A.DailyIncomeMin, 0) = 0 AND COALESCE(A.DailyIncomeMax, 0) > 0 AND A.DailyIncomeMax <= @vDailyIncomeMax)) "
				End If
			End If
			'</日給>

			'<時給>
			If HourlyIncomeMin & HourlyIncomeMax <> "" Then
				If HourlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHourlyIncomeMin INT"
					sParams = sParams & ",@vHourlyIncomeMin = " & HourlyIncomeMin
				End If

				If HourlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHourlyIncomeMax INT"
					sParams = sParams & ",@vHourlyIncomeMax = " & HourlyIncomeMax
				End If

				If sTemp <> "" Then sTemp = sTemp & "OR "
				If HourlyIncomeMin <> "" And HourlyIncomeMax <> "" Then
					'時給下限,上限両方の入力がある場合
					sTemp = sTemp & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND (A.HourlyIncomeMin BETWEEN @vHourlyIncomeMin AND @vHourlyIncomeMax)) OR (COALESCE(A.HourlyIncomeMax, 0) > 0 AND (A.HourlyIncomeMax BETWEEN @vHourlyIncomeMin AND @vHourlyIncomeMax))) "
				ElseIf HourlyIncomeMin <> "" Then
					'時給下限のみ入力がある場合
					sTemp = sTemp & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND A.HourlyIncomeMin >= @vHourlyIncomeMin) OR (COALESCE(A.HourlyIncomeMax, 0) > 0 AND A.HourlyIncomeMax >= @vHourlyIncomeMin)) "
				ElseIf HourlyIncomeMax <> "" Then
					'時給上限のみ入力がある場合
					sTemp = sTemp & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND A.HourlyIncomeMin <= @vHourlyIncomeMax) OR (COALESCE(A.HourlyIncomeMin, 0) = 0 AND COALESCE(A.HourlyIncomeMax, 0) > 0 AND A.HourlyIncomeMax <= @vHourlyIncomeMax)) "
				End If
			End If
			'</時給>

			If sTemp <> "" Then sTemp = "(" & sTemp & ") "

			'歩合制
			If PercentagePayFlag <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPercentagePayFlag VARCHAR(1)"
				sParams = sParams & ",@vPercentagePayFlag = N'" & PercentagePayFlag & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = "A.PercentagePayFlag = @vPercentagePayFlag "
			End If

			sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_Info AS A WHERE " & RTrim(sTemp) & ") AS CSLY ON VWOC.OrderCode = CSLY.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'給与 end
		'******************************************************************************

		'******************************************************************************
		'勤務開始・終了時間 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		If WorkStartHour & WorkEndHour <> "" Then
			If WorkStartHour <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkStartHour VARCHAR(2) "
				sParams = sParams & ",@vWorkStartHour = N'" & WorkStartHour & "'"

				If WorkStartMinute = "" Then WorkStartMinute = "00"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkStartMinute VARCHAR(2) "
				sParams = sParams & ",@vWorkStartMinute = N'" & WorkStartMinute & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "A.WorkStartTime >= @vWorkStartHour + @vWorkStartMinute "
			End If

			If WorkEndHour <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkEndHour VARCHAR(2) "
				sParams = sParams & ",@vWorkEndHour = N'" & WorkEndHour & "'"

				If WorkEndMinute = "" Then WorkEndMinute = "00"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkEndMinute VARCHAR(2) "
				sParams = sParams & ",@vWorkEndMinute = N'" & WorkEndMinute & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "A.WorkEndTime <= @vWorkEndHour + @vWorkEndMinute "
			End If

			If WorkStartHour <> "" And WorkEndHour <> "" Then
				If WorkStartHour < WorkEndHour Then
					'勤務開始時間 < 勤務終了時間の場合、夜間の業務時間を除くようにする
					sTemp2 = "AND A.WorkStartTime < A.WorkEndTime "
				End If
			End If

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_WorkingCondition AS A WITH(NOLOCK) WHERE " & sTemp & RTrim(sTemp2) & ") AS CWTM ON VWOC.OrderCode = CWTM.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'勤務開始・終了時間 end
		'******************************************************************************

		'******************************************************************************
		'週休 start
		'------------------------------------------------------------------------------
		sTemp = ""
		If WeeklyHolidayType <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vWeeklyHolidayType VARCHAR(3) "
			sParams = sParams & ",@vWeeklyHolidayType = N'" & WeeklyHolidayType & "'"

			sTemp = sTemp & "CWHT.WeeklyHolidayType = @vWeeklyHolidayType "

			sJoin = sJoin & "INNER JOIN C_Info AS CWHT WITH(NOLOCK) ON VWOC.OrderCode = CWHT.OrderCode AND " & RTrim(sTemp) & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'週休 end
		'******************************************************************************

		'******************************************************************************
		'年齢 start
		'------------------------------------------------------------------------------
		'sTemp = ""
		'If Age <> "" Then
		'	If sDeclare <> "" Then sDeclare = sDeclare & ","
		'	sDeclare = sDeclare & "@vAge INT "
		'	sParams = sParams & ",@vAge = " & Age

		'	sTemp = sTemp & "(@vAge BETWEEN ISNULL(CAGE.AgeMin, 0) AND ISNULL(CAGE.AgeMax, 255)) "

		'	sJoin = sJoin & "INNER JOIN C_Info AS CAGE WITH(NOLOCK) ON VWOC.OrderCode = CAGE.OrderCode AND " & RTrim(sTemp) & vbCrLf
		'End If
		'------------------------------------------------------------------------------
		' 年齢 end
		'******************************************************************************

		'******************************************************************************
		'卒業年検索 start
		'------------------------------------------------------------------------------
		'sTemp = ""
		If CStr(GraduateYear) <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vGraduateYear SMALLINT "
			sParams = sParams & ",@vGraduateYear = " & GraduateYear

			sTemp = sTemp & "(@vGraduateYear BETWEEN A.YearMin AND A.YearMax) "

			sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_GraduateYear AS A WITH(NOLOCK) WHERE " & RTrim(sTemp) & ") AS CGY ON VWOC.OrderCode = CGY.OrderCode " & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'卒業年検索 end
		'******************************************************************************

		'******************************************************************************
		'契約期間 start
		'------------------------------------------------------------------------------
		sTemp = ""
		If IsRE(AgreementTerm, "^[123]$", True) = True Then
			If AgreementTerm = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WITH(NOLOCK) WHERE WorkPeriod <= 1 UNION SELECT OrderCode FROM C_Undertake WITH(NOLOCK) WHERE WorkPeriod <= 1 UNION SELECT OrderCode FROM C_TTP WITH(NOLOCK) WHERE WorkPeriod <= 1) AS CAT ON VWOC.OrderCode = CAT.OrderCode" & vbCrLf
			ElseIf AgreementTerm = "2" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WITH(NOLOCK) WHERE WorkPeriod <= 2 UNION SELECT OrderCode FROM C_Undertake WITH(NOLOCK) WHERE WorkPeriod <= 2 UNION SELECT OrderCode FROM C_TTP WITH(NOLOCK) WHERE WorkPeriod <= 2) AS CAT ON VWOC.OrderCode = CAT.OrderCode" & vbCrLf
			ElseIf AgreementTerm = "3" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WITH(NOLOCK) WHERE WorkPeriod > 3 UNION SELECT OrderCode FROM C_Undertake WITH(NOLOCK) WHERE WorkPeriod > 3 UNION SELECT OrderCode FROM C_TTP WITH(NOLOCK) WHERE WorkPeriod > 3) AS CAT ON VWOC.OrderCode = CAT.OrderCode" & vbCrLf
			End If
		End If
		'------------------------------------------------------------------------------
		'契約期間 end
		'******************************************************************************

		'******************************************************************************
		'保有資格 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If LicenseCount > 0 Then
			For idx = 0 To LicenseCount - 1
				sTemp = ""

				If LicenseGroupCode(idx) <> "" Then
					'大分類
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLicenseGroupCode" & iParamNo & " VARCHAR(2)"
					sParams = sParams & ",@vLicenseGroupCode" & iParamNo & " = N'" & LicenseGroupCode(idx) & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "A.GroupCode = @vLicenseGroupCode" & iParamNo & " "

					'中分類
					If LicenseCategoryCode(idx) <> "" Then
						If sDeclare <> "" Then sDeclare = sDeclare & ","
						sDeclare = sDeclare & "@vLicenseCategoryCode" & iParamNo & " VARCHAR(3)"
						sParams = sParams & ",@vLicenseCategoryCode" & iParamNo & " = N'" & LicenseCategoryCode(idx) & "'"

						If sTemp <> "" Then sTemp = sTemp & "AND "
						sTemp = sTemp & "A.CategoryCode = @vLicenseCategoryCode" & iParamNo & " "
					End If

					'小分類
					If LicenseCode(idx) <> "" Then
						If sDeclare <> "" Then sDeclare = sDeclare & ","
						sDeclare = sDeclare & "@vLicenseCode" & iParamNo & " VARCHAR(2)"
						sParams = sParams & ",@vLicenseCode" & iParamNo & " = N'" & LicenseCode(idx) & "'"

						If sTemp <> "" Then sTemp = sTemp & "AND "
						sTemp = sTemp & "A.Code = @vLicenseCode" & iParamNo & " "
					End If

					If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
					sTemp2 = sTemp2 & "(" & Trim(sTemp) & ") "

					iParamNo = iParamNo + 1
				End If
			Next

			If sTemp2 <> "" Then sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_License AS A WITH(NOLOCK) WHERE " & RTrim(sTemp2) & ") AS CL ON VWOC.OrderCode = CL.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'保有資格 end
		'******************************************************************************

		'******************************************************************************
		'スキル start
		'------------------------------------------------------------------------------
		iParamNo2 = 1
		'OS
		sTemp = ""
		iParamNo = 1
		If OSCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo2 & " = N'OS'"

			aValue = Split(Replace(OSCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSkillCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vSkillCode" & iParamNo & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSkillCode" & iParamNo

				iParamNo = iParamNo + 1
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WITH(NOLOCK) WHERE A.CategoryCode = @vSkillCategoryCode" & iParamNo2 & " AND A.Code IN (" & Trim(sTemp) & ")) AS CSKL" & iParamNo2 & " ON VWOC.OrderCode = CSKL" & iParamNo2 & ".OrderCode" & vbCrLf
			iParamNo2 = iParamNo2 + 1
		End If

		'アプリケーション
		sTemp = ""
		If ApplicationCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo2 & " = N'Application'"

			aValue = Split(Replace(ApplicationCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSkillCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vSkillCode" & iParamNo & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSkillCode" & iParamNo

				iParamNo = iParamNo + 1
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WITH(NOLOCK) WHERE A.CategoryCode = @vSkillCategoryCode" & iParamNo2 & " AND A.Code IN (" & Trim(sTemp) & ")) AS CSKL" & iParamNo2 & " ON VWOC.OrderCode = CSKL" & iParamNo2 & ".OrderCode" & vbCrLf
			iParamNo2 = iParamNo2 + 1
		End If

		'開発言語
		sTemp = ""
		If DevelopmentLanguageCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo2 & " = N'DevelopmentLanguage'"

			aValue = Split(Replace(DevelopmentLanguageCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSkillCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vSkillCode" & iParamNo & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSkillCode" & iParamNo

				iParamNo = iParamNo + 1
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WITH(NOLOCK) WHERE A.CategoryCode = @vSkillCategoryCode" & iParamNo2 & " AND A.Code IN (" & Trim(sTemp) & ")) AS CSKL" & iParamNo2 & " ON VWOC.OrderCode = CSKL" & iParamNo2 & ".OrderCode" & vbCrLf
			iParamNo2 = iParamNo2 + 1
		End If

		'データベース
		sTemp = ""
		If DatabaseCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo2 & " = N'Database'"

			aValue = Split(Replace(DatabaseCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSkillCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vSkillCode" & iParamNo & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSkillCode" & iParamNo

				iParamNo = iParamNo + 1
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WITH(NOLOCK) WHERE A.CategoryCode = @vSkillCategoryCode" & iParamNo2 & " AND A.Code IN (" & Trim(sTemp) & ")) AS CSKL" & iParamNo2 & " ON VWOC.OrderCode = CSKL" & iParamNo2 & ".OrderCode" & vbCrLf
			iParamNo2 = iParamNo2 + 1
		End If
		'------------------------------------------------------------------------------
		'スキル end
		'******************************************************************************

		'******************************************************************************
		'キーワード start
		'------------------------------------------------------------------------------
		sTemp = ""
		If Keyword <> "" Then
			aValue = Split(Replace(Replace(Replace(Keyword, "(", "（"), ")", "）"), "　", " "), " ")
			For idx = LBound(aValue) To UBound(aValue)
				If sTemp <> "" Then
					If KeywordFlag = "1" Then
						sTemp = sTemp & " OR "
					ElseIf KeywordFlag = "2" Then
						sTemp = sTemp & " AND "
					Else
						sTemp = sTemp & " AND "
					End If
				End If
				sTemp = sTemp & "FORMSOF(THESAURUS, " & aValue(idx) & "*)"
			Next
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vKeyword VARCHAR(400)"
			sParams = sParams & ",@vKeyword = N'" & sTemp & "'"

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY A.OrderCode) AS Num, A.OrderCode FROM C_FullTextNavi AS A WITH(NOLOCK) WHERE CONTAINS(A.Text, @vKeyword)) AS CFTN ON VWOC.OrderCode = CFTN.OrderCode" & vbCrLf
            'sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY A.OrderCode) AS Num, A.OrderCode FROM C_FullTextNavi AS A WITH(NOLOCK) left join (SELECT A.OrderCode From C_info as A INNER JOIN CompanyInfo as B on A.CompanyCode = B.CompanyCode WHERE (b.CompanyName_K like @vKeyword OR b.CompanyName_F like @vKeyword)) as B ON A.OrderCode = B.OrderCode WHERE CONTAINS(A.Text, @vKeyword)) AS CFTN ON VWOC.OrderCode = CFTN.OrderCode" & vbCrLf
		
        End If
		'------------------------------------------------------------------------------
		'キーワード end
		'******************************************************************************

		'******************************************************************************
		'対象企業の求人票一覧用情報コード start
		'------------------------------------------------------------------------------
		If PictureOrderCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vPictureOrderCode VARCHAR(8) "
			sParams = sParams & ",@vPictureOrderCode = N'" & PictureOrderCode & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.CompanyCode FROM C_Info AS A WITH(NOLOCK) WHERE OrderCode = @vPictureOrderCode) AS CPOC ON VWOC.CompanyCode = CPOC.CompanyCode AND VWOC.OrderType = '0'" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'対象企業の求人票一覧用情報コード end
		'******************************************************************************

		'******************************************************************************
		'登録日 start
		'------------------------------------------------------------------------------
		If RegistDay <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vRegistDay VARCHAR(8) "
			sParams = sParams & ",@vRegistDay = N'" & RegistDay & "'"

			sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_Info AS A WITH(NOLOCK) WHERE RegistDay >= CONVERT(DATETIME, @vRegistDay)) AS CRD ON VWOC.OrderCode = CRD.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'登録日 end
		'******************************************************************************

		'******************************************************************************
		'前回表示時の最新情報コード start
		'------------------------------------------------------------------------------
		If BOC <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vBeforeOrderCode VARCHAR(8) "
			sParams = sParams & ",@vBeforeOrderCode = N'" & BOC & "'"

			If sWhere <> "" Then sWhere = sWhere & "AND "
			sWhere = sWhere & "VWOC.OrderCode > @vBeforeOrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'前回表示時の最新情報コード end
		'******************************************************************************

		'******************************************************************************
		'情報コードCSV start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If OrderCode <> "" Then
			aValue = Split(Replace(OrderCode, " ", ""), ",")
			For iParamNo = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOrderCode" & iParamNo & " CHAR(8)"
				sParams = sParams & ",@vOrderCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vOrderCode" & iParamNo
			Next

			If sWhere <> "" Then sWhere = sWhere & "AND "
			If UBound(aValue) = 0 Then
				sWhere = sWhere & "VWOC.OrderCode = " & sTemp & vbCrLf
			Else
				sWhere = sWhere & "VWOC.OrderCode IN (" & sTemp & ")" & vbCrLf
			End If
		End If
		'------------------------------------------------------------------------------
		'情報コードCSV end
		'******************************************************************************

		'<社内案件の対象企業>
		If LISOrderCompanyCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vLISOrderCompanyCode VARCHAR(8) "
			sParams = sParams & ",@vLISOrderCompanyCode = N'" & LISOrderCompanyCode & "'"

			If sWhere <> "" Then sWhere = sWhere & "AND "
			sWhere = sWhere & "VWOC.CompanyCode = @vLISOrderCompanyCode" & vbCrLf
		End If
		'</社内案件の対象企業>

		If CStr(Top) <> "" Then Top = "TOP " & Top & vbCrLf
		sSQL = ""
		sSQL = sSQL & "SELECT " & Top & "VWOC.OrderCode "
		sSQL = sSQL & ",VWOC.SortNum "
		sSQL = sSQL & ",VWOC.RegistDay ,VWOC.UpdateDay" & vbCrLf
		sSQL = sSQL & "FROM vw_OrderCode_PlusOld AS VWOC WITH(NOLOCK)" & vbCrLf
		sSQL = sSQL & sJoin
		If sWhere <> "" Then sSQL = sSQL & "WHERE " & sWhere
		sSQL = sSQL & "ORDER BY VWOC.SortNum ASC, VWOC.UpdateDay DESC"

        If FeatureFlag <> "" Then
            sSQL = ""
            sSQL = sSQL & "SELECT  " & vbCrLf
            sSQL = sSQL & "VWOC.OrderCode  " & vbCrLf
            sSQL = sSQL & ",VWOC.SortNum  " & vbCrLf
            sSQL = sSQL & ",VWOC.RegistDay  " & vbCrLf
            sSQL = sSQL & ",VWOC.UpdateDay  " & vbCrLf
            sSQL = sSQL & "FROM vw_OrderCode_PlusOld AS VWOC WITH(NOLOCK)  " & vbCrLf
            sSQL = sSQL & "INNER JOIN ( " & vbCrLf
            sSQL = sSQL & "SELECT DISTINCT A.OrderCode  " & vbCrLf
            sSQL = sSQL & "FROM C_JobType AS A WITH(NOLOCK)  " & vbCrLf
            If FeatureFlag = "1" Then
                sSQL = sSQL & "WHERE (A.JobTypeCode IN ('1302000','1325000','1326000','1308000','1319000','1312000','1318000','1311000','1399000'))) AS CJT ON VWOC.OrderCode = CJT.OrderCode  " & vbCrLf
            End If
            sSQL = sSQL & "ORDER BY VWOC.SortNum ASC, VWOC.UpdateDay DESC " & vbCrLf
        End If


		GetSQLOrderSearchDetail = ""
		GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & "/*ナビ・求人票詳細検索*/" & vbCrLf
		GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		If sSearchCondition <> "" Then
			sSearchCondition = "<table class=""pattern1"" border=""0"" style=""width:600px;""><thead><tr><th colspan=""2"" style=""width:588px;"">検索条件</th></tr></thead><tbody>" & sSearchCondition & "</tbody></table>"
		Else
			sSearchCondition = "なし"
		End If
'Response.Write GetSQLOrderSearchDetail
	End Function

	'******************************************************************************
	'概　要：求人のキーワード検索ＬＯＧ書き込みＳＱＬを取得
	'引　数：
	'備　考：
	'履　歴：2012/02/21 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function GetSQLWriteLog()
		Dim sSQL,sSN,sKW,sSiteType

		sSN = Request.ServerVariables("SERVER_NAME")
		If InStr(sSN,"www.shigotonavi.co.jp") + InStr(sSN,"www-b1.shigotonavi.co.jp") > 0 Then
			sSiteType = "1"
		ElseIf InStr(sSN,"m.shigotonavi.jp") + InStr(sSN,"m-b1.shigotonavi.jp") > 0 Then
			sSiteType = "2"
		ElseIf InStr(sSN,"www.a-rirekisyo.jp") + InStr(sSN,"www-b1.a-rirekisyo.jp") > 0 Then
			sSiteType = "3"
		End If

		sKW = KW
		If sKW = "" Then sKW = Keyword

		If sKW > "" Then
			sSQL = sSQL & "EXEC up_RegLOG_SearchOrderKeyword '" & G_USERID & "'"
			sSQL = sSQL & ",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'"
			sSQL = sSQL & ",'" & ChkSQLStr(Session.SessionID) & "'"
			sSQL = sSQL & ",'" & sSiteType & "'"
			sSQL = sSQL & ",'" & sKW & "';"
		End If

		GetSQLWriteLog = sSQL
	End Function

	'******************************************************************************
	'概　要：求人票詳細検索条件出力ＨＴＭＬを取得
	'引　数：
	'備　考：
	'履　歴：2007/04/04 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function GetHtmlSearchCondition()
		Dim sTemp
		Dim sTemp2
		Dim idx

		If SearchDetailFlag = "" Then Exit Function

		GetHtmlSearchCondition = ""

		'社内外案件検索フラグ
		sTemp = ""
		If OrderTypeFlag <> "" Then
			If OrderTypeFlag = "0" Then
				sTemp = "一般求人情報"
			ElseIf OrderTypeFlag = "1" Then
				sTemp = "リスの紹介・派遣情報"
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("一般／リス区分",sTemp)
		End If

		'新着フラグ
		sTemp = ""
		If NewFlag = "1" or NewKoukokuFlag = "1" Then
			sTemp = "新着情報"
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("新着区分",sTemp)
		End If

		'職種
		sTemp2 = ""
		If JobTypeBigCode1 & JobTypeCode1 & JobTypeBigCode2 & JobTypeCode2 & JobTypeBigCode3 & JobTypeCode3 <> "" Then
			sTemp = ""
			If JobTypeBigCode1 & JobTypeCode1 <> "" Then
				sTemp = sTemp & JobTypeName1
				If sTemp = "" And JobTypeBigName1 <> "" Then sTemp = sTemp & JobTypeBigName1

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "　"
				sTemp2 = sTemp2 & sTemp
			End If

			sTemp = ""
			If JobTypeBigCode2 & JobTypeCode2 <> "" Then
				sTemp = sTemp & JobTypeName2
				If sTemp = "" And JobTypeBigName2 <> "" Then sTemp = sTemp & JobTypeBigName2

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "　"
				sTemp2 = sTemp2 & sTemp
			End If

			sTemp = ""
			If JobTypeBigCode3 & JobTypeCode3 <> "" Then
				sTemp = sTemp & JobTypeName3
				If sTemp = "" And JobTypeBigName3 <> "" Then sTemp = sTemp & JobTypeBigName3

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "　"
				sTemp2 = sTemp2 & sTemp
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("職種",sTemp2)
		End If

		'勤務地
		sTemp = ""
		If PrefectureCode & City & RailwayLineCode & RailwayLineCode <> "" Then
			'エリア
			sTemp = sTemp & AreaName

			'都道府県
			If PrefectureName <> "" Then
				sTemp = sTemp & "　"
				sTemp = sTemp & PrefectureName
			End If

			'市区郡
			If City <> "" Then
				sTemp = sTemp & "　"
				sTemp = sTemp & City
			End If

			'沿線
			If RailwayLineCode <> "" Then
				sTemp = sTemp & "　"
				sTemp = sTemp & RailwayLineName
			End If

			'駅
			If StationCode <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "　"
				sTemp = sTemp & StationName & "駅"
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("勤務地",sTemp)
		End If

		'勤務形態
		sTemp = ""
		If WorkingTypeCode1 & WorkingTypeCode2 & WorkingTypeCode3 <> "" Then
			If WorkingTypeCode1 <> "" Then sTemp = sTemp & WorkingTypeName1
			If WorkingTypeCode2 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "　"
				sTemp = sTemp & WorkingTypeName2
			End If
			If WorkingTypeCode3 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "　"
				sTemp = sTemp & WorkingTypeName3
			End If
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("勤務形態",sTemp)
		End If

		'業種
		sTemp = ""
		If IndustryTypeCode <> "" Then
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("業種",IndustryTypeName)
		End If

		'給与
		sTemp = ""
		If YearlyIncomeMin & YearlyIncomeMax & MonthlyIncomeMin & MonthlyIncomeMax & DailyIncomeMin & DailyIncomeMax & HourlyIncomeMin & HourlyIncomeMax & PercentagePayFlag <> "" Then
			If PercentagePayFlag = "1" Then
				sTemp = sTemp & "歩合制あり"
			ElseIf PercentagePayFlag = "0" Then
				sTemp = sTemp & "歩合制なし"
			End If
			If YearlyIncomeMin & YearlyIncomeMax <> "" Then
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "年収：" & YearlyIncomeMin & "～" & YearlyIncomeMax
			End If
			If MonthlyIncomeMin & YearlyIncomeMax <> "" Then
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "月給：" & MonthlyIncomeMin & "～" & MonthlyIncomeMax
			End If
			If DailyIncomeMin & DailyIncomeMax <> "" Then
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "日給：" & DailyIncomeMin & "～" & DailyIncomeMax
			End If
			If HourlyIncomeMin & HourlyIncomeMax <> "" Then
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "時給：" & HourlyIncomeMin & "～" & HourlyIncomeMax
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("給与",sTemp)
		End If

		'特徴
		sTemp = ""
		If InexperiencedPersonFlag & UtilizeLanguageFlag & TempFlag & UITurnFlag & ManyHolidayFlag & FlexFlag & _
		NearStationFlag & NoSmokingFlag & NewlyBuiltFlag & LandmarkFlag & RenovationFlag & DesignersFlag & _
		CompanyCafeteriaFlag & ShortOvertimeFlag & MaternityFlag & DressFreeFlag & MammyFlag & FixedTimeFlag & _
		ShortTimeFlag & HandicappedFlag & RentAllFlag & RentPartFlag & MealsFlag & MealsAssistanceFlag & _
		TrainingCostFlag & EntrepreneurCostFlag & MoneyFlag & LandShopFlag & FindJobFestiveFlag & AppointmentFlag & SocietyInsuranceFlag <> "" Then
			If InexperiencedPersonFlag = "1" Then sTemp = sTemp & "「未経験者ＯＫ」"
			If UtilizeLanguageFlag = "1" Then sTemp = sTemp & "「語学を活かす」"
			If TempFlag = "1" Then sTemp = sTemp & "「派遣」"
			If UITurnFlag = "1" Then sTemp = sTemp & "「ＵＩターン歓迎」"
			If ManyHolidayFlag = "1" Then sTemp = sTemp & "「休日１２０日以上」"
			If FlexFlag = "1" Then sTemp = sTemp & "「フレックス」"
			If NearStationFlag = "1" Then sTemp = sTemp & "「駅近(徒歩5分以内)」"
			If NoSmokingFlag = "1" Then sTemp = sTemp & "「禁煙・分煙」"
			If NewlyBuiltFlag = "1" Then sTemp = sTemp & "「新築ビル・オフィス(5年以内)」"
			If LandmarkFlag = "1" Then sTemp = sTemp & "「高層(15階以上)ビル」"
			If RenovationFlag = "1" Then sTemp = sTemp & "「リノベーションビル・オフィス(5年以内)」"
			If DesignersFlag = "1" Then sTemp = sTemp & "「デザイナーズビル・オフィス」"
			If CompanyCafeteriaFlag = "1" Then sTemp = sTemp & "「社員食堂」"
			If ShortOvertimeFlag = "1" Then sTemp = sTemp & "「残業10h/月以内」"
			If MaternityFlag = "1" Then sTemp = sTemp & "「産休・育休実績あり」"
			If DressFreeFlag = "1" Then sTemp = sTemp & "「服装自由」"
			If MammyFlag = "1" Then sTemp = sTemp & "「子育てママ歓迎」"
			If FixedTimeFlag = "1" Then sTemp = sTemp & "「18時までに退社」"
			If ShortTimeFlag = "1" Then sTemp = sTemp & "「1日6時間以内労働」"
			If HandicappedFlag = "1" Then sTemp = sTemp & "「障害者歓迎」"
			If RentAllFlag = "1" Then sTemp = sTemp & "「住宅費用全額補助あり」"
			If RentPartFlag = "1" Then sTemp = sTemp & "「住宅費用一部補助あり」"
			If MealsFlag = "1" Then sTemp = sTemp & "「食事・賄い付き案件」"
			If MealsAssistanceFlag = "1" Then sTemp = sTemp & "「食事補助制度あり」"
			If TrainingCostFlag = "1" Then sTemp = sTemp & "「研修費助成制度あり」"
			If EntrepreneurCostFlag = "1" Then sTemp = sTemp & "「起業機材補助制度あり」"
			If MoneyFlag = "1" Then sTemp = sTemp & "「無利子・低利子補助制度あり」"
			If LandShopFlag = "1" Then sTemp = sTemp & "「土地・店舗等提供制度あり」"
			If FindJobFestiveFlag = "1" Then sTemp = sTemp & "「就職お祝い金制度あり」"
			If AppointmentFlag = "1" Then sTemp = sTemp & "「正社員登用制度あり」"
			If SocietyInsuranceFlag = "1" Then sTemp = sTemp & "「社保完備」"

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("特徴",sTemp)
		End If

		'就業時間
		sTemp = ""
		If WorkStartHour & WorkStartMinute & WorkEndHour & WorkEndMinute <> "" Then
			If WorkStartHour & WorkStartMinute <> "" Then sTemp = sTemp & "就業開始時間：" & WorkStartHour & ":" & WorkStartMinute & "&nbsp;以降"
			If WorkEndHour & WorkEndMinute <> "" And sTemp <> "" Then sTemp = sTemp & ","
			If WorkEndHour & WorkEndMinute <> "" Then sTemp = sTemp & "就業終了時間：" & WorkEndHour & ":" & WorkEndMinute & "&nbsp;以前"

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("就業時間",sTemp)
		End If

		'週休種類
		sTemp = ""
		If WeeklyHolidayType <> "" Then
			sTemp = sTemp & WeeklyHolidayTypeName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("週休種類",sTemp)
		End If

		'年齢
		'sTemp = ""
		'If Age <> "" Then
		'	sTemp = sTemp & Age & "歳を含む"
		'	GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("年齢",sTemp)
		'End If

		'卒業年
		sTemp = ""
		If SchoolTypeName & GraduateYear <> "" Then
			sTemp = SchoolTypeName & "　" & GraduateYear & "年卒"
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("学歴",sTemp)
		End If

		'契約期間
		sTemp = ""
		If AgreementTerm <> "" Then
			If AgreementTerm = "1" Then
				sTemp = "～１ヶ月"
			ElseIf AgreementTerm = "2" Then
				sTemp = "～２ヶ月"
			ElseIf AgreementTerm = "3" Then
				sTemp = "３ヶ月以上"
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("契約期間",sTemp)
		End If

		'資格
		sTemp = ""
		If LicenseCount > 0 Then
			For idx = 0 To LicenseCount - 1
				If sTemp <> "" Then sTemp = sTemp & ","

				If LicenseName(idx) <> "" Then
					sTemp = sTemp & LicenseName(idx)
				ElseIf LicenseCategoryName(idx) <> "" Then
					sTemp = sTemp & LicenseCategoryName(idx)
				ElseIf LicenseGroupName(idx) <> "" Then
					sTemp = sTemp & LicenseGroupName(idx)
				End If
			Next

			If sTemp <> "" Then GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("資格",sTemp)
		End If

		'ＯＳ
		sTemp = ""
		If OSName <> "" Then
			sTemp = sTemp & OSName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("ＯＳ",sTemp)
		End If

		'アプリケーション
		sTemp = ""
		If ApplicationName <> "" Then
			sTemp = sTemp & ApplicationName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("アプリケーション",sTemp)
		End If

		'開発言語
		sTemp = ""
		If DevelopmentLanguageName <> "" Then
			sTemp = sTemp & DevelopmentLanguageName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("開発言語",sTemp)
		End If

		'データベース
		sTemp = ""
		If DatabaseName <> "" Then
			sTemp = sTemp & DatabaseName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("データベース",sTemp)
		End If

		'キーワード
		sTemp = ""
		If Keyword <> "" Then
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("キーワード",Keyword)
		End If

		'情報コード（検索）
		If OrderCode <> "" Then
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("情報コード",OrderCode)
		End If

		If GetHtmlSearchCondition <> "" Then
			'GetHtmlSearchCondition = "<table class=""pattern1"" border=""0"" style=""width:600px;""><colgroup><col style=""width:138px;""><col style=""width:439px;""></colgroup><thead><tr><th colspan=""2"" style=""width:588px;"">検索条件</th></tr></thead><tbody>" & GetHtmlSearchCondition & "</tbody></table>"
			GetHtmlSearchCondition = "<div class=""description"">" & GetHtmlSearchCondition & "</div>"
		End If

	End Function

	Private Function GetHtmlSearchConditionTable(ByVal vKey, ByVal vValue)
		'GetHtmlSearchConditionTable = "<tr><th>" & vKey & "</th><td>" & vValue & "</td></tr>"
		GetHtmlSearchConditionTable = "【"&vKey&"】&nbsp;" & vValue & "<br>"
	End Function
End Class
%>
