<%
'******************************************************************************
'概　要：検索条件を保持するクラス
'関　数：■Public
'　　　：GetSearchParam				：お仕事詳細検索ページへ渡すGETパラメータを生成して取得
'　　　：DspConditionHidden			：お仕事詳細検索の条件hiddenを出力する
'　　　：GetSQLOrderSearchDetail	：求人票詳細検索ＳＱＬを取得
'　　　：GetSQLWriteLog				：求人票検索ＬＯＧ書き込みＳＱＬを取得
'　　　：GetHtmlSearchCondition		：求人票詳細検索条件出力ＨＴＭＬを取得
'　　　：
'　　　：■Private
'　　　：Class_Initialize			：コンストラクタ
'　　　：SetNames					：コードに対応した名称をメンバ変数に設定
'　　　：ChkSQLType					：カンタン検索か詳細検索かを判断してflgEasySearchを設定
'　　　：ChkData					：メンバ変数の整合性をチェックして訂正
'　　　：
'備　考：■■■ 詳細検索用パラメータ （アドホックなＳＱＬ生成）
'　　　：sjtbig1：希望職種大分類１
'　　　：sjt1	：希望職種１
'　　　：sjtbig2：希望職種大分類２
'　　　：sjt2	：希望職種２
'　　　：src1	：希望沿線１
'　　　：src2	：希望沿線２
'　　　：ssc1	：希望駅１
'　　　：ssc2	：希望駅２
'　　　：sac1	：希望エリア１
'　　　：spc1	：希望都道府県１
'　　　：sct1	：希望市区郡１
'　　　：sac2	：希望エリア２
'　　　：spc2	：希望都道府県２
'　　　：sct2	：希望市区郡２
'　　　：swt1	：希望勤務形態１
'　　　：swt2	：希望勤務形態２
'　　　：swt3	：希望勤務形態３
'　　　：sit	：希望業種(カンマ区切り [XX,XX,XX])
'　　　：sppf	：歩合制
'　　　：syi	：年収
'　　　：smi	：月給
'　　　：sdi	：日給
'　　　：shi	：時給
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
'　　　：soc	：情報コード（検索）
'　　　：
'　　　：■■■ カンタン検索用パラメータ (ストアド up_SearchOrder 活用)
'　　　：jt		：職種大分類コード
'　　　：jt2	：職種コード
'　　　：ac		：エリアコード
'　　　：ac2	：都道府県コード
'　　　：wt		：勤務形態コード
'　　　：kw		：キーワード
'　　　：
'　　　：■■■ 情報ツール用
'　　　：boc	：前回表示情報コード
'　　　：
'　　　：■■■ 詳細検索用ＰＯＳＴデータ （アドホックなＳＱＬ生成）
'　　　：CONF_SearchHopeJobTypeBigCode1			：希望職種大分類１
'　　　：CONF_SearchHopeJobTypeCode1			：希望職種１
'　　　：CONF_SearchHopeJobTypeBigCode2			：希望職種大分類２
'　　　：CONF_SearchHopeJobTypeCode2			：希望職種２
'　　　：CONF_SearchRailwayLineCode1			：希望沿線１
'　　　：CONF_SearchRailwayLineCode2			：希望沿線２
'　　　：CONF_SearchStationCode1				：希望駅１
'　　　：CONF_SearchStationCode2				：希望駅２
'　　　：CONF_SearchAreaCode1					：希望エリア１
'　　　：CONF_SearchPrefectureCode1				：希望都道府県１
'　　　：CONF_SearchCity1						：希望市区郡１
'　　　：CONF_SearchAreaCode2					：希望エリア２
'　　　：CONF_SearchPrefectureCode2				：希望都道府県２
'　　　：CONF_SearchCity2						：希望市区郡２
'　　　：CONF_SearchHopeWorkingTypeCode1		：希望勤務形態１
'　　　：CONF_SearchHopeWorkingTypeCode2		：希望勤務形態２
'　　　：CONF_SearchHopeWorkingTypeCode3		：希望勤務形態３
'　　　：CONF_SearchHopeIndustryTypeCode		：希望業種(カンマ区切り [XX,XX,XX])
'　　　：CONF_SearchPercentagePayFlag			：歩合制
'　　　：CONF_SearchYearlyIncome				：年収
'　　　：CONF_SearchMonthlyIncome				：月給
'　　　：CONF_SearchDailyIncome					：日給
'　　　：CONF_SearchHourlyIncome				：時給
'　　　：CONF_SearchWorkStartHour				：就業開始時間（時）
'　　　：CONF_SearchWorkStartMinute				：就業開始時間（分）
'　　　：CONF_SearchWorkEndHour					：就業終了時間（時）
'　　　：CONF_SearchWorkEndMinute				：就業終了時間（分）
'　　　：CONF_SearchWeeklyHolidayType			：週休種類
'　　　：CONF_SearchAge							：年齢
'　　　：CONF_SearchAgreementTerm				：契約期間
'　　　：CONF_SearchLicenseGroupCode1			：資格大分類
'　　　：CONF_SearchLicenseCategoryCode1		：資格中分類
'　　　：CONF_SearchLicenseCode1				：資格小分類
'　　　：CONF_SearchOSCode1						：ＯＳ
'　　　：CONF_SearchApplicationCode1			：アプリケーション
'　　　：CONF_SearchDevelopmentLanguageCode1	：開発言語
'　　　：CONF_SearchDatabaseCode1				：データベース
'　　　：CONF_SearchKeyword						：検索ワード
'　　　：CONF_SearchKeywordFlag					：検索ワードフラグ [1]OR [2]AND
'　　　：CONF_SearchOrderCode					：情報コード（検索）
'　　　：CONF_SearchInexperiencedPersonFlag		：特徴（未経験歓迎）
'　　　：CONF_SearchUtilizeLanguageFlag			：特徴（語学を活かす）
'　　　：CONF_SearchTempFlag					：特徴（派遣）※現在未使用
'　　　：CONF_SearchUITurnFlag					：特徴（ＵＩターン）
'　　　：CONF_SearchManyHolidayFlag				：特徴（休日１２０日以上）
'　　　：CONF_SearchFlexFlag					：特徴（フレックスタイム）
'　　　：CONF_SP								：特集コード（検索では使わない。パラメータ生成用に保持する。）
'　　　：
'　　　：■■■ カンタン検索用ＰＯＳＴデータ (ストアド up_SearchOrder 活用)
'　　　：CONF_JT	：職種大分類コード
'　　　：CONF_JT2	：職種コード
'　　　：CONF_AC	：エリアコード
'　　　：CONF_AC2	：都道府県コード
'　　　：CONF_WT	：勤務形態コード
'　　　：CONF_ST1	：特徴（未経験歓迎）
'　　　：CONF_ST2	：特徴（語学を活かす）
'　　　：CONF_ST3	：特徴（派遣）※現在未使用
'　　　：CONF_ST4	：特徴（ＵＩターン）
'　　　：CONF_ST5	：特徴（休日１２０日以上）
'　　　：CONF_ST6	：特徴（フレックスタイム）
'　　　：CONF_KW	：キーワード
'　　　：
'　　　：■■■ 使用方法
'　　　：Dim oSOC
'　　　：Dim sSQL
'　　　：Set oSOC = New clsSearchOrderCondition	'生成された時点でパラメータとＰＯＳＴデータからＳＱＬが生成されている
'　　　：oSOC.Top = 100	'SELECT句で上限を設定
'　　　：sSQL = oSOC.GetSQLOrderSearchDetail()	'ＳＱＬを取得
'　　　：
'更　新：2007/04/05 LIS K.Kokubo 作成
'　　　：2007/10/10 LIS K.Kokubo 情報ツール用変数追加
'　　　：2007/10/31 LIS K.Kokubo TOP ??? 用変数追加
'　　　：2008/01/15 LIS K.Kokubo パラメータ化クエリ化
'******************************************************************************
Class clsSearchOrderCondition
	'検索条件メンバ変数
	Public Top						'SELECTで取得する件数 (SELECT TOP ○ * FROM ～)
	Public JobTypeBigCode1			'希望職種大分類１
	Public JobTypeCode1				'希望職種１
	Public JobTypeBigCode2			'希望職種大分類２
	Public JobTypeCode2				'希望職種２
	Public RailwayLineCode1			'希望沿線１
	Public RailwayLineCode2			'希望沿線２
	Public StationCode1				'希望駅１
	Public StationCode2				'希望駅２
	Public AreaCode1				'希望エリア１
	Public PrefectureCode1			'希望都道府県１
	Public City1					'希望市区郡１
	Public AreaCode2				'希望エリア２
	Public PrefectureCode2			'希望都道府県２
	Public City2					'希望市区郡２
	Public WorkingTypeCode1			'希望勤務形態１
	Public WorkingTypeCode2			'希望勤務形態２
	Public WorkingTypeCode3			'希望勤務形態３
	Public IndustryTypeCode			'希望業種(カンマ区切り [XX,XX,XX])
	Public IndustryTypeCode1		'希望業種１
	Public IndustryTypeCode2		'希望業種２
	Public IndustryTypeCode3		'希望業種３
	Public PercentagePayFlag		'歩合制
	Public YearlyIncome				'年収
	Public MonthlyIncome			'月給
	Public DailyIncome				'日給
	Public HourlyIncome				'時給
	Public WorkStartHour			'就業開始時間（時）
	Public WorkStartMinute			'就業開始時間（分）
	Public WorkEndHour				'就業終了時間（時）
	Public WorkEndMinute			'就業終了時間（分）
	Public WeeklyHolidayType		'週休種類
	Public Age						'年齢
	Public AgreementTerm			'契約期間
	Public LicenseGroupCode1		'資格大分類
	Public LicenseCategoryCode1		'資格中分類
	Public LicenseCode1				'資格小分類
	Public OSCode1					'ＯＳ
	Public OACode1
	Public ApplicationCode1			'アプリケーション
	Public DevelopmentLanguageCode1	'開発言語
	Public DatabaseCode1			'データベース
	Public Keyword					'検索ワード
	Public KeywordFlag				'検索ワードフラグ [1]OR [2]AND
	Public OrderCode				'情報コード（検索）
	Public Specialty
	Public InexperiencedPersonFlag	'特徴（）
	Public UtilizeLanguageFlag		'特徴（）
	Public TempFlag					'特徴（派遣）
	Public UITurnFlag				'特徴（ＵＩターン歓迎）
	Public ManyHolidayFlag			'特徴（休日１２０日以上）
	Public FlexFlag					'特徴（フレックス）

	'カンタン検索条件
	Public JT	'職種大分類コード
	Public JT2	'職種コード
	Public AC	'エリアコード
	Public AC2	'都道府県コード
	Public WT	'勤務形態コード
	Public ST
	Public ST1	'特徴
	Public ST2	'特徴
	Public ST3	'特徴
	Public ST4	'特徴
	Public ST5	'特徴
	Public ST6	'特徴
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
	Public RailwayLineName1	'希望沿線名称１
	Public RailwayLineName2	'希望沿線名称２
	Public StationName1
	Public StationName2
	Public AreaName1
	Public AreaName2
	Public PrefectureName1
	Public PrefectureName2
	Public WorkingTypeName1
	Public WorkingTypeName2
	Public WorkingTypeName3
	Public IndustryTypeName1
	Public IndustryTypeName2
	Public IndustryTypeName3
	Public WeeklyHolidayTypeName
	Public OSName1
	Public ApplicationName1
	Public DevelopmentLanguageName1
	Public DatabaseName1
	Public LicenseGroupName1	'資格大分類名称１
	Public LicenseCategoryName1	'資格中分類名称１
	Public LicenseName1		'資格名称１

	'その他メンバ変数
	Public flgEasySearch	'カンタン検索フラグ [True]カンタン検索 [False]詳細検索
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
		'FORMデータから検索条件を取得
		JobTypeBigCode1 = GetForm("CONF_SearchHopeJobTypeBigCode1", 1)
		JobTypeCode1 = GetForm("CONF_SearchHopeJobTypeCode1", 1)
		JobTypeBigCode2 = GetForm("CONF_SearchHopeJobTypeBigCode2", 1)
		JobTypeCode2 = GetForm("CONF_SearchHopeJobTypeCode2", 1)
		RailwayLineCode1 = GetForm("CONF_SearchRailwayLineCode1", 1)
		RailwayLineCode2 = GetForm("CONF_SearchRailwayLineCode2", 1)
		StationCode1 = GetForm("CONF_SearchStationCode1", 1)
		StationCode2 = GetForm("CONF_SearchStationCode2", 1)
		AreaCode1 = GetForm("CONF_SearchAreaCode1", 1)
		PrefectureCode1 = GetForm("CONF_SearchPrefectureCode1", 1)
		City1 = GetForm("CONF_SearchCity1", 1)
		AreaCode2 = GetForm("CONF_SearchAreaCode2", 1)
		PrefectureCode2 = GetForm("CONF_SearchPrefectureCode2", 1)
		City2 = GetForm("CONF_SearchCity2", 1)
		WorkingTypeCode1 = GetForm("CONF_SearchHopeWorkingTypeCode1", 1)
		WorkingTypeCode2 = GetForm("CONF_SearchHopeWorkingTypeCode2", 1)
		WorkingTypeCode3 = GetForm("CONF_SearchHopeWorkingTypeCode3", 1)
		IndustryTypeCode = GetForm("CONF_SearchHopeIndustryTypeCode", 1)
		PercentagePayFlag = GetForm("CONF_SearchPercentagePayFlag", 1)
		YearlyIncome = GetForm("CONF_SearchYearlyIncome", 1)
		MonthlyIncome = GetForm("CONF_SearchMonthlyIncome", 1)
		DailyIncome = GetForm("CONF_SearchDailyIncome", 1)
		HourlyIncome = GetForm("CONF_SearchHourlyIncome", 1)
		WorkStartHour = GetForm("CONF_SearchWorkStartHour", 1)
		WorkStartMinute = GetForm("CONF_SearchWorkStartMinute", 1)
		WorkEndHour = GetForm("CONF_SearchWorkEndHour", 1)
		WorkEndMinute = GetForm("CONF_SearchWorkEndMinute", 1)
		WeeklyHolidayType = GetForm("CONF_SearchWeeklyHolidayType", 1)
		Age = GetForm("CONF_SearchAge", 1)
		AgreementTerm = GetForm("CONF_SearchAgreementTerm", 1)
		LicenseGroupCode1 = GetForm("CONF_SearchLicenseGroupCode1", 1)
		LicenseCategoryCode1 = GetForm("CONF_SearchLicenseCategoryCode1", 1)
		LicenseCode1 = GetForm("CONF_SearchLicenseCode1", 1)
		OSCode1 = GetForm("CONF_SearchOSCode1", 1)
		ApplicationCode1 = GetForm("CONF_SearchApplicationCode1", 1)
		DevelopmentLanguageCode1 = GetForm("CONF_SearchDevelopmentLanguageCode1", 1)
		DatabaseCode1 = GetForm("CONF_SearchDatabaseCode1", 1)
		Keyword = GetForm("CONF_SearchKeyword", 1)
		KeywordFlag = GetForm("CONF_SearchKeywordFlag", 1)
		OrderCode = GetForm("CONF_SearchOrderCode", 1)
		InexperiencedPersonFlag = GetForm("CONF_SearchInexperiencedPersonFlag", 1)
		UtilizeLanguageFlag = GetForm("CONF_SearchUtilizeLanguageFlag", 1)
		TempFlag = GetForm("CONF_SearchTempFlag", 1)
		UITurnFlag = GetForm("CONF_SearchUITurnFlag", 1)
		ManyHolidayFlag = GetForm("CONF_SearchManyHolidayFlag", 1)
		FlexFlag = GetForm("CONF_SearchFlexFlag", 1)
		SP = GetForm("CONF_SP", 1)

		'パラメータから検索条件を取得
		If GetForm("sjtbig1", 2) <> "" Then JobTypeBigCode1 = GetForm("sjtbig1", 2)
		If GetForm("sjt1", 2) <> "" Then JobTypeCode1 = GetForm("sjt1", 2)
		If GetForm("sjtbig2", 2) <> "" Then JobTypeBigCode2 = GetForm("sjtbig2", 2)
		If GetForm("sjt2", 2) <> "" Then JobTypeCode2 = GetForm("sjt2", 2)
		If GetForm("src1", 2) <> "" Then RailwayLineCode1 = GetForm("src1", 2)
		If GetForm("src2", 2) <> "" Then RailwayLineCode2 = GetForm("src2", 2)
		If GetForm("ssc1", 2) <> "" Then StationCode1 = GetForm("ssc1", 2)
		If GetForm("ssc2", 2) <> "" Then StationCode2 = GetForm("ssc2", 2)
		If GetForm("sac1", 2) <> "" Then AreaCode1 = GetForm("sac1", 2)
		If GetForm("spc1", 2) <> "" Then PrefectureCode1 = GetForm("spc1", 2)
		If GetForm("sct1", 2) <> "" Then City1 = GetForm("sct1", 2)
		If GetForm("sac2", 2) <> "" Then AreaCode2 = GetForm("sac2", 2)
		If GetForm("spc2", 2) <> "" Then PrefectureCode2 = GetForm("spc2", 2)
		If GetForm("sct2", 2) <> "" Then City2 = GetForm("sct2", 2)
		If GetForm("swt1", 2) <> "" Then WorkingTypeCode1 = GetForm("swt1", 2)
		If GetForm("swt2", 2) <> "" Then WorkingTypeCode2 = GetForm("swt2", 2)
		If GetForm("swt3", 2) <> "" Then WorkingTypeCode3 = GetForm("swt3", 2)
		If GetForm("sit", 2) <> "" Then IndustryTypeCode = GetForm("sit", 2)
		If GetForm("sppf", 2) <> "" Then PercentagePayFlag = GetForm("sppf", 2)
		If GetForm("syi", 2) <> "" Then YearlyIncome = GetForm("syi", 2)
		If GetForm("smi", 2) <> "" Then MonthlyIncome = GetForm("smi", 2)
		If GetForm("sdi", 2) <> "" Then DailyIncome = GetForm("sdi", 2)
		If GetForm("shi", 2) <> "" Then HourlyIncome = GetForm("shi", 2)
		If GetForm("swsh", 2) <> "" Then WorkStartHour = GetForm("swsh", 2)
		If GetForm("swsm", 2) <> "" Then WorkStartMinute = GetForm("swsm", 2)
		If GetForm("sweh", 2) <> "" Then WorkEndHour = GetForm("sweh", 2)
		If GetForm("swem", 2) <> "" Then WorkEndMinute = GetForm("swem", 2)
		If GetForm("swht", 2) <> "" Then WeeklyHolidayType = GetForm("swht", 2)
		If GetForm("sage", 2) <> "" Then Age = GetForm("sage", 2)
		If GetForm("sat", 2) <> "" Then AgreementTerm = GetForm("sat", 2)
		If GetForm("slg1", 2) <> "" Then LicenseGroupCode1 = GetForm("slg1", 2)
		If GetForm("slc1", 2) <> "" Then LicenseCategoryCode1 = GetForm("slc1", 2)
		If GetForm("sl1", 2) <> "" Then LicenseCode1 = GetForm("sl1", 2)
		If GetForm("sos1", 2) <> "" Then OSCode1 = GetForm("sos1", 2)
		If GetForm("sap1", 2) <> "" Then ApplicationCode1 = GetForm("sap1", 2)
		If GetForm("sdl1", 2) <> "" Then DevelopmentLanguageCode1 = GetForm("sdl1", 2)
		If GetForm("sdb1", 2) <> "" Then DatabaseCode1 = GetForm("sdb1", 2)
		If GetForm("skw", 2) <> "" Then Keyword = GetForm("skw", 2)
		If GetForm("skwflag", 2) <> "" Then KeywordFlag = GetForm("skwflag", 2)
		If GetForm("sst", 2) <> "" Then Specialty = GetForm("sst", 2)
		If GetForm("soc", 2) <> "" Then OrderCode = GetForm("soc", 2)

		If IsRE(GetForm("sst", 2), "^[01][01][01][01][01][01]$", True) = True Then
			If Mid(GetForm("sst", 2), 1, 1) = "1" Then InexperiencedPersonFlag = "1"
			If Mid(GetForm("sst", 2), 2, 1) = "1" Then UtilizeLanguageFlag = "1"
			If Mid(GetForm("sst", 2), 3, 1) = "1" Then TempFlag = "1"
			If Mid(GetForm("sst", 2), 4, 1) = "1" Then UITurnFlag = "1"
			If Mid(GetForm("sst", 2), 5, 1) = "1" Then ManyHolidayFlag = "1"
			If Mid(GetForm("sst", 2), 6, 1) = "1" Then FlexFlag = "1"
		End If

		'特徴ビット文字列
		Specialty = ""
		If InexperiencedPersonFlag & UtilizeLanguageFlag & TempFlag & UITurnFlag & ManyHolidayFlag & FlexFlag <> "" Then
			If InexperiencedPersonFlag <> "" Then: Specialty = Specialty & InexperiencedPersonFlag: Else: Specialty = Specialty & "0": End If
			If UtilizeLanguageFlag <> "" Then: Specialty = Specialty & UtilizeLanguageFlag: Else: Specialty = Specialty & "0": End If
			If TempFlag <> "" Then: Specialty = Specialty & TempFlag: Else: Specialty = Specialty & "0": End If
			If UITurnFlag <> "" Then: Specialty = Specialty & UITurnFlag: Else: Specialty = Specialty & "0": End If
			If ManyHolidayFlag <> "" Then: Specialty = Specialty & ManyHolidayFlag: Else: Specialty = Specialty & "0": End If
			If FlexFlag <> "" Then: Specialty = Specialty & FlexFlag: Else: Specialty = Specialty & "0": End If
		End If

		'希望業種
		If IndustryTypeCode <> "" Then
			IndustryTypeCode = Replace(IndustryTypeCode, " ", "")
			Dim aHITC
			Dim idx

			aHITC = Split(IndustryTypeCode, ",")
			For idx = 0 To UBound(aHITC)
				Select Case idx
					Case 0:	IndustryTypeCode1 = aHITC(idx)
					Case 1:	IndustryTypeCode2 = aHITC(idx)
					Case 2:	IndustryTypeCode3 = aHITC(idx)
				End Select
			Next
		End If

		'カンタン検索条件取得（FORMデータ）
		JT = GetForm("CONF_JT", 1)
		JT2 = GetForm("CONF_JT2", 1)
		AC = GetForm("CONF_AC", 1)
		AC2 = GetForm("CONF_AC2", 1)
		WT = GetForm("CONF_WT", 1)
		ST1 = GetForm("CONF_ST1", 1)
		ST2 = GetForm("CONF_ST2", 1)
		ST3 = GetForm("CONF_ST3", 1)
		ST4 = GetForm("CONF_ST4", 1)
		ST5 = GetForm("CONF_ST5", 1)
		ST6 = GetForm("CONF_ST6", 1)
		KW = GetForm("CONF_KW", 1)

		'特徴ビット文字列
		ST = ""
		If ST1 & ST2 & ST3 & ST4 & ST5 & ST6 <> "" Then
			ST = ""
			If ST1 <> "" Then: ST = ST & ST1: Else: ST = ST & "0": End If
			If ST2 <> "" Then: ST = ST & ST2: Else: ST = ST & "0": End If
			If ST3 <> "" Then: ST = ST & ST3: Else: ST = ST & "0": End If
			If ST4 <> "" Then: ST = ST & ST4: Else: ST = ST & "0": End If
			If ST5 <> "" Then: ST = ST & ST5: Else: ST = ST & "0": End If
			If ST6 <> "" Then: ST = ST & ST6: Else: ST = ST & "0": End If

			Specialty = ST
		End If

		'ＴＯＰから
		POC = GetForm("poc", 2)

		If POC <> "" Then OrderCode = POC

		'特集
		If GetForm("sp", 2) <> "" Then SP = GetForm("sp", 2)

		'カンタン検索条件取得（パラメータ）
		If GetForm("jt", 2) <> "" Then JT = GetForm("jt", 2)
		If GetForm("jt2", 2) <> "" Then JT2 = GetForm("jt2", 2)
		If GetForm("ac", 2) <> "" Then AC = GetForm("ac", 2)
		If GetForm("ac2", 2) <> "" Then AC2 = GetForm("ac2", 2)
		If GetForm("wt", 2) <> "" Then WT = GetForm("wt", 2)
		If GetForm("kw", 2) <> "" Then KW = GetForm("kw", 2)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST1 = Mid(GetForm("st", 2), 1, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST2 = Mid(GetForm("st", 2), 2, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST3 = Mid(GetForm("st", 2), 3, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST4 = Mid(GetForm("st", 2), 4, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST5 = Mid(GetForm("st", 2), 5, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST6 = Mid(GetForm("st", 2), 6, 1)

		If JT <> "" Then JobTypeCode1 = JT
		If JT2 <> "" Then JobTypeCode1 = JT2
		If AC <> "" Then AreaCode1 = AC
		If AC2 <> "" Then PrefectureCode1 = AC2
		If WT <> "" Then WorkingTypeCode1 = WT
		If KW <> "" Then Keyword = KW

		'沿線検索（FORMデータ）
		PC = GetForm("CONF_PC", 1)
		RC = GetForm("CONF_RC", 1)
		SC = GetForm("CONF_SC", 1)

		'沿線検索（パラメータ）
		If GetForm("pc", 2) <> "" Then PC = GetForm("pc", 2)
		If GetForm("rc", 2) <> "" Then RC = GetForm("rc", 2)
		If GetForm("sc", 2) <> "" Then SC = GetForm("sc", 2)

		If PC <> "" Then PrefectureCode1 = PC
		If RC <> "" Then RailwayLineCode1 = RC
		If SC <> "" Then StationCode1 = SC

		'情報ツール
		BOC = GetForm("boc", 2)

		'**********************************************************************
		'補正 start
		'----------------------------------------------------------------------
		If AC = "" And AC2 <> "" Then
			sSQL = "up_GetListPrefecture '', '" & AC2 & "', ''"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then AC = oRS.Collect("AreaCode")
			Call RSClose(oRS)
		End If
		'----------------------------------------------------------------------
		'補正 end
		'**********************************************************************

		'データ整合性チェック
		Call ChkData()

		'カンタン検索・詳細検索判定
		Call ChkSQLType()

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
	'概　要：コードに対応した名称を取得する
	'作成者：Lis K.Kokubo
	'作成日：2007/04/04 Lis K.Kokubo
	'更　新：
	'備　考：
	'******************************************************************************
	Private Sub SetNames()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

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
		End If

		'希望沿線１
		If RailwayLineCode1 <> "" Then
			sSQL = "up_GetRailwayLineName '" & RailwayLineCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				RailwayLineName1 = ChkStr(oRS.Collect("RailwayLineName"))
			End If
			Call RSClose(oRS)
		End If
		'希望沿線２
		If RailwayLineCode2 <> "" Then
			sSQL = "up_GetRailwayLineName '" & RailwayLineCode2 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				RailwayLineName2 = ChkStr(oRS.Collect("RailwayLineName"))
			End If
			Call RSClose(oRS)
		End If

		'希望駅１
		If StationCode1 <> "" Then
			sSQL = "up_GetStationName '" & StationCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				StationName1 = ChkStr(oRS.Collect("StationName"))
			End If
			Call RSClose(oRS)
		End If
		'希望駅２
		If StationCode2 <> "" Then
			sSQL = "up_GetStationName '" & StationCode2 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				StationName2 = ChkStr(oRS.Collect("StationName"))
			End If
			Call RSClose(oRS)
		End If

		'エリア１
		If AreaCode1 <> "" Then
			AreaName1 = GetDetail("Area", AreaCode1)
		End If

		'エリア２
		If AreaCode2 <> "" Then
			AreaName2 = GetDetail("Area", AreaCode2)
		End If

		'都道府県１
		If PrefectureCode1 <> "" Then
			PrefectureName1 = GetDetail("Prefecture", PrefectureCode1)

			If AreaCode1 = "" Then
				sSQL = "SELECT A.AreaCode, B.AreaName FROM Area AS A WITH(NOLOCK) INNER JOIN vw_Area AS B ON A.AreaCode = B.AreaCode WHERE A.PrefectureCode = '" & PrefectureCode1 & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					AreaCode1 = ChkStr(oRS.Collect("AreaCode"))
					AreaName1 = ChkStr(oRS.Collect("AreaName"))
				End If
				Call RSClose(oRS)
			End If
		End If

		'都道府県２
		If PrefectureCode2 <> "" Then
			PrefectureName2 = GetDetail("Prefecture", PrefectureCode2)

			If AreaCode2 = "" Then
				sSQL = "SELECT A.AreaCode, B.AreaName FROM Area AS A WITH(NOLOCK) INNER JOIN vw_Area AS B ON A.AreaCode = B.AreaCode WHERE A.PrefectureCode = '" & PrefectureCode2 & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					AreaCode2 = ChkStr(oRS.Collect("AreaCode"))
					AreaName2 = ChkStr(oRS.Collect("AreaName"))
				End If
				Call RSClose(oRS)
			End If
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

		'業種１
		If IndustryTypeCode1 <> "" Then
			IndustryTypeName1 = GetDetail("IndustryType", IndustryTypeCode1)
		End If

		'業種２
		If IndustryTypeCode2 <> "" Then
			IndustryTypeName2 = GetDetail("IndustryType", IndustryTypeCode2)
		End If

		'業種３
		If IndustryTypeCode3 <> "" Then
			IndustryTypeName3 = GetDetail("IndustryType", IndustryTypeCode3)
		End If

		'週休種類
		If WeeklyHolidayType <> "" Then
			WeeklyHolidayTypeName = GetDetail("WeeklyHolidayType", WeeklyHolidayType)
		End If

		'ＯＳ
		If OSCode1 <> "" Then
			OSName1 = GetDetail("OS", OSCode1)
		End If

		'アプリケーション
		If ApplicationCode1 <> "" Then
			ApplicationName1 = GetDetail("Application", ApplicationCode1)
		End If

		'開発言語
		If DevelopmentLanguageCode1 <> "" Then
			DevelopmentLanguageName1 = GetDetail("DevelopmentLanguage", DevelopmentLanguageCode1)
		End If

		'データベース
		If DatabaseCode1 <> "" Then
			DatabaseName1 = GetDetail("Database", DatabaseCode1)
		End If

		'資格
		If IsRE(LicenseGroupCode1, "^\d\d$", True) = True Then
			'大分類
			sSQL = "sp_GetListLicenseGroup '" & LicenseGroupCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				LicenseGroupName1 = ChkStr(oRS.Collect("GroupName"))
			End If
			Call RSClose(oRS)

			'中分類
			If IsRE(LicenseCategoryCode1, "^\d\d\d$", True) = True Then
				sSQL = "sp_GetListLicenseCategory '" & LicenseGroupCode1 & "', '" & LicenseCategoryCode1 & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					LicenseCategoryName1 = ChkStr(oRS.Collect("CategoryName"))
				End If
				Call RSClose(oRS)

				'大分類
				If IsRE(LicenseCode1, "^\d\d$", True) = True Then
					sSQL = "sp_GetListLicense '" & LicenseGroupCode1 & "', '" & LicenseCategoryCode1 & "', '" & LicenseCode1 & "'"
					flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
					If GetRSState(oRS) = True Then
						LicenseName1 = ChkStr(oRS.Collect("Name"))
					End If
					Call RSClose(oRS)
				End If
			End If
		End If
	End Sub

	'******************************************************************************
	'概　要：カンタン検索か詳細検索かを判断してflgEasySearchを設定
	'備　考：
	'更　新：2007/11/01 LIS K.Kokubo 作成
	'******************************************************************************
	Private Sub ChkSQLType()
		'カンタン検索・詳細検索判定
		If JT & JT2 & AC & AC2 & WT & ST1 & ST2 & ST3 & ST4 & ST5 & ST6 & PC & RC & SC & KW <> "" Then
			flgEasySearch = True
		Else
			flgEasySearch = False
		End If
	End Sub

	'******************************************************************************
	'概　要：データの整合性をチェック
	'作成者：Lis K.Kokubo
	'作成日：2007/04/17 Lis K.Kokubo
	'更　新：
	'備　考：
	'******************************************************************************
	Private Sub ChkData()
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
	End Sub

	'******************************************************************************
	'概　要：お仕事詳細検索の条件hiddenを出力する
	'作成者：Lis K.Kokubo
	'作成日：2007/04/04 Lis K.Kokubo
	'更　新：
	'備　考：
	'******************************************************************************
	Public Sub DspConditionHidden()
		Response.Write "<input name=""CONF_SearchHopeJobTypeBigCode1"" type=""hidden"" value=""" & JobTypeBigCode1 & """>"
		Response.Write "<input name=""CONF_SearchHopeJobTypeCode1"" type=""hidden"" value=""" & JobTypeCode1 & """>"
		Response.Write "<input name=""CONF_SearchHopeJobTypeBigCode2"" type=""hidden"" value=""" & JobTypeBigCode2 & """>"
		Response.Write "<input name=""CONF_SearchHopeJobTypeCode2"" type=""hidden"" value=""" & JobTypeCode2 & """>"
		Response.Write "<input name=""CONF_SearchRailwayLineCode1"" type=""hidden"" value=""" & RailwayLineCode1 & """>"
		Response.Write "<input name=""CONF_SearchRailwayLineCode2"" type=""hidden"" value=""" & RailwayLineCode2 & """>"
		Response.Write "<input name=""CONF_SearchStationCode1"" type=""hidden"" value=""" & StationCode1 & """>"
		Response.Write "<input name=""CONF_SearchStationCode2"" type=""hidden"" value=""" & StationCode2 & """>"
		Response.Write "<input name=""CONF_SearchAreaCode1"" type=""hidden"" value=""" & AreaCode1 & """>"
		Response.Write "<input name=""CONF_SearchPrefectureCode1"" type=""hidden"" value=""" & PrefectureCode1 & """>"
		Response.Write "<input name=""CONF_SearchCity1"" type=""hidden"" value=""" & City1 & """>"
		Response.Write "<input name=""CONF_SearchAreaCode2"" type=""hidden"" value=""" & AreaCode2 & """>"
		Response.Write "<input name=""CONF_SearchPrefectureCode2"" type=""hidden"" value=""" & PrefectureCode2 & """>"
		Response.Write "<input name=""CONF_SearchCity2"" type=""hidden"" value=""" & City2 & """>"
		Response.Write "<input name=""CONF_SearchHopeWorkingTypeCode1"" type=""hidden"" value=""" & WorkingTypeCode1 & """>"
		Response.Write "<input name=""CONF_SearchHopeWorkingTypeCode2"" type=""hidden"" value=""" & WorkingTypeCode2 & """>"
		Response.Write "<input name=""CONF_SearchHopeWorkingTypeCode3"" type=""hidden"" value=""" & WorkingTypeCode3 & """>"
		Response.Write "<input name=""CONF_SearchHopeIndustryTypeCode"" type=""hidden"" value=""" & IndustryTypeCode & """>"
		Response.Write "<input name=""CONF_SearchPercentagePayFlag"" type=""hidden"" value=""" & PercentagePayFlag & """>"
		Response.Write "<input name=""CONF_SearchYearlyIncome"" type=""hidden"" value=""" & YearlyIncome & """>"
		Response.Write "<input name=""CONF_SearchMonthlyIncome"" type=""hidden"" value=""" & MonthlyIncome & """>"
		Response.Write "<input name=""CONF_SearchDailyIncome"" type=""hidden"" value=""" & DailyIncome & """>"
		Response.Write "<input name=""CONF_SearchHourlyIncome"" type=""hidden"" value=""" & HourlyIncome & """>"
		Response.Write "<input name=""CONF_SearchWorkStartHour"" type=""hidden"" value=""" & WorkStartHour & """>"
		Response.Write "<input name=""CONF_SearchWorkStartMinute"" type=""hidden"" value=""" & WorkStartMinute & """>"
		Response.Write "<input name=""CONF_SearchWorkEndHour"" type=""hidden"" value=""" & WorkEndHour & """>"
		Response.Write "<input name=""CONF_SearchWorkEndMinute"" type=""hidden"" value=""" & WorkEndMinute & """>"
		Response.Write "<input name=""CONF_SearchWeeklyHolidayType"" type=""hidden"" value=""" & WeeklyHolidayType & """>"
		Response.Write "<input name=""CONF_SearchAge"" type=""hidden"" value=""" & Age & """>"
		Response.Write "<input name=""CONF_SearchAgreementTerm"" type=""hidden"" value=""" & AgreementTerm & """>"
		Response.Write "<input name=""CONF_SearchLicenseGroupCode1"" type=""hidden"" value=""" & LicenseGroupCode1 & """>"
		Response.Write "<input name=""CONF_SearchLicenseCategoryCode1"" type=""hidden"" value=""" & LicenseCategoryCode1 & """>"
		Response.Write "<input name=""CONF_SearchLicenseCode1"" type=""hidden"" value=""" & LicenseCode1 & """>"
		Response.Write "<input name=""CONF_SearchOSCode1"" type=""hidden"" value=""" & OSCode1 & """>"
		Response.Write "<input name=""CONF_SearchApplicationCode1"" type=""hidden"" value=""" & ApplicationCode1 & """>"
		Response.Write "<input name=""CONF_SearchDevelopmentLanguageCode1"" type=""hidden"" value=""" & DevelopmentLanguageCode1 & """>"
		Response.Write "<input name=""CONF_SearchDatabaseCode1"" type=""hidden"" value=""" & DatabaseCode1 & """>"
		Response.Write "<input name=""CONF_SearchKeyword"" type=""hidden"" value=""" & Keyword & """>"
		Response.Write "<input name=""CONF_SearchKeywordFlag"" type=""hidden"" value=""" & KeywordFlag & """>"
		Response.Write "<input name=""CONF_SearchOrderCode"" type=""hidden"" value=""" & OrderCode & """>"
		'沿線検索
		Response.Write "<input name=""CONF_PC"" type=""hidden"" value=""" & PC & """>"
		Response.Write "<input name=""CONF_RC"" type=""hidden"" value=""" & RC & """>"
		Response.Write "<input name=""CONF_SC"" type=""hidden"" value=""" & SC & """>"
		'特集
		Response.Write "<input name=""CONF_SP"" type=""hidden"" value=""" & SP & """>"
	End Sub

	'******************************************************************************
	'概　要：お仕事詳細検索ページへ渡すGETパラメータを生成して取得。
	'作成者：Lis Kokubo
	'作成日：2007/02/27
	'引　数：
	'備　考：■制限
	'　　　：パラメータを含むURLは、IEの制限が2048文字までであるので、それに合わせる。
	'******************************************************************************
	Public Function GetSearchParam()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

		Dim odSC
		Dim odResult
		Dim idxKey
		Dim aKeys
		Dim aValues

		GetSearchParam = ""

		If JobTypeBigCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sjtbig1=" & JobTypeBigCode1
		If JobTypeCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sjt1=" & JobTypeCode1
		If JobTypeBigCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;sjtbig2=" & JobTypeBigCode2
		If JobTypeCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;sjt2=" & JobTypeCode2
		If RailwayLineCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;src1=" & RailwayLineCode1
		If RailwayLineCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;src2=" & RailwayLineCode2
		If StationCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;ssc1=" & StationCode1
		If StationCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;ssc2=" & StationCode2
		If AreaCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sac1=" & AreaCode1
		If PrefectureCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;spc1=" & PrefectureCode1
		If City1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sct1=" & Server.URLEncode(City1)
		If AreaCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;sac2=" & AreaCode2
		If PrefectureCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;spc2=" & PrefectureCode2
		If City2 <> "" Then GetSearchParam = GetSearchParam & "&amp;sct2=" & Server.URLEncode(City2)
		If WorkingTypeCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;swt1=" & WorkingTypeCode1
		If WorkingTypeCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;swt2=" & WorkingTypeCode2
		If WorkingTypeCode3 <> "" Then GetSearchParam = GetSearchParam & "&amp;swt3=" & WorkingTypeCode3
		If IndustryTypeCode <> "" Then GetSearchParam = GetSearchParam & "&amp;sit=" & IndustryTypeCode
		If PercentagePayFlag <> "" Then GetSearchParam = GetSearchParam & "&amp;sppf=" & PercentagePayFlag
		If YearlyIncome <> "" Then GetSearchParam = GetSearchParam & "&amp;syi=" & YearlyIncome
		If MonthlyIncome <> "" Then GetSearchParam = GetSearchParam & "&amp;smi=" & MonthlyIncome
		If DailyIncome <> "" Then GetSearchParam = GetSearchParam & "&amp;sdi=" & DailyIncome
		If HourlyIncome <> "" Then GetSearchParam = GetSearchParam & "&amp;shi=" & HourlyIncome
		If WorkStartHour <> "" Then GetSearchParam = GetSearchParam & "&amp;swsh=" & WorkStartHour
		If WorkStartMinute <> "" Then GetSearchParam = GetSearchParam & "&amp;swsm=" & WorkStartMinute
		If WorkEndHour <> "" Then GetSearchParam = GetSearchParam & "&amp;sweh=" & WorkEndHour
		If WorkEndMinute <> "" Then GetSearchParam = GetSearchParam & "&amp;swem=" & WorkEndMinute
		If WeeklyHolidayType <> "" Then GetSearchParam = GetSearchParam & "&amp;swht=" & WeeklyHolidayType
		If Age <> "" Then GetSearchParam = GetSearchParam & "&amp;sage=" & Age
		If AgreementTerm <> "" Then GetSearchParam = GetSearchParam & "&amp;sat=" & AgreementTerm
		If LicenseGroupCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;slg1=" & LicenseGroupCode1
		If LicenseCategoryCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;slc1=" & LicenseCategoryCode1
		If LicenseCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sl1=" & LicenseCode1
		If OSCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sos1=" & OSCode1
		If ApplicationCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sap1=" & ApplicationCode1
		If DevelopmentLanguageCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sdl1=" & DevelopmentLanguageCode1
		If DatabaseCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sdb1=" & DatabaseCode1
		If Keyword <> "" Then GetSearchParam = GetSearchParam & "&amp;skw=" & Server.URLEncode(Keyword)
		If KeywordFlag <> "" Then GetSearchParam = GetSearchParam & "&amp;skwflag=" & KeywordFlag
		If OrderCode <> "" Then GetSearchParam = GetSearchParam & "&amp;soc=" & OrderCode
		If Specialty <> "" Then GetSearchParam = GetSearchParam & "&amp;sst=" & Specialty
		If SP <> "" Then GetSearchParam = GetSearchParam & "&amp;sp=" & SP

		If GetSearchParam <> "" Then
			'頭の&amp;を？に変換
			GetSearchParam = "?" & Mid(GetSearchParam, 6)

			'ＩＥの仕様はパラメータの上限が２０４８バイト
			GetSearchParam = Left(GetSearchParam, 2048)
		End If
	End Function

	'******************************************************************************
	'概　要：求人票詳細検索ＳＱＬを取得
	'作成者：Lis Kokubo
	'作成日：2007/04/04
	'引　数：
	'備　考：
	'******************************************************************************
	Function GetSQLOrderSearchDetail()
		Dim sJoin		: sJoin = ""
		Dim sWhere		: sWhere = ""
		Dim sDeclare	: sDeclare = ""
		Dim sParams		: sParams = ""
		Dim iParamNo
		Dim sFrom
		Dim sTemp
		Dim sTemp2
		Dim sTemp3
		Dim aValue
		Dim idx
		Dim sSearchCondition

		'データ整合性チェック
		Call ChkData()
		'カンタン検索・詳細検索判定
		Call ChkSQLType()

		'******************************************************************************
		'職種 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If JobTypeBigCode1 & JobTypeCode1 & JobTypeBigCode2 & JobTypeCode2 <> "" Then
			If JobTypeBigCode1 & JobTypeCode1 <> "" Then
				sTemp = JobTypeBigCode1
				If JobTypeCode1 <> "" Then sTemp = JobTypeCode1

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "CJT.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			If JobTypeBigCode2 & JobTypeCode2 <> "" Then
				sTemp = JobTypeBigCode2
				If JobTypeCode2 <> "" Then sTemp = JobTypeCode2

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "CJT.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CJT.OrderCode FROM C_JobType AS CJT WHERE (" & sTemp2 & ")) AS CJT ON VWOC.OrderCode = CJT.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'職種 end
		'******************************************************************************

		'******************************************************************************
		'沿線 start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If RailwayLineCode1 & RailwayLineCode2 <> "" Then
			If RailwayLineCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vRailwayLineCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vRailwayLineCode" & iParamNo & " = N'" & RailwayLineCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vRailwayLineCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If RailwayLineCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vRailwayLineCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vRailwayLineCode" & iParamNo & " = N'" & RailwayLineCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vRailwayLineCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT CNS.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS CNS "
			sJoin = sJoin & "INNER JOIN StationStop AS SS "
			sJoin = sJoin & "ON CNS.StationCode = SS.StationCode "
			sJoin = sJoin & "INNER JOIN B_RailwayLine AS BRL "
			sJoin = sJoin & "ON SS.RailwayLineCode = BRL.RailwayLineCode "
			sJoin = sJoin & "AND BRL.RailwayLineCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CRL "
			sJoin = sJoin & "ON VWOC.OrderCode = CRL.OrderCode "
		End If

		'------------------------------------------------------------------------------
		'沿線 end
		'******************************************************************************

		'******************************************************************************
		'駅 start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If StationCode1 & StationCode2 <> "" Then
			If StationCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vStationCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vStationCode" & iParamNo & " = N'" & StationCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vStationCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If StationCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vStationCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vStationCode" & iParamNo & " = N'" & StationCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vStationCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT CNS.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS CNS "
			sJoin = sJoin & "WHERE CNS.StationCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CNS "
			sJoin = sJoin & "ON VWOC.OrderCode = CNS.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'駅 end
		'******************************************************************************

		'******************************************************************************
		'希望勤務地 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If AreaCode1 & AreaCode2 <> "" Then
			sTemp = ""
			If AreaCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAreaCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAreaCode" & iParamNo & " = N'" & AreaCode1 & "'"

				sTemp = "AREA.AreaCode = @vAreaCode" & iParamNo & " "

				If PrefectureCode1 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vPrefectureCode" & iParamNo & " VARCHAR(3)"
					sParams = sParams & ",@vPrefectureCode" & iParamNo & " = N'" & PrefectureCode1 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CWP.WorkingPlacePrefectureCode = @vPrefectureCode" & iParamNo & " "
				End If

				If PrefectureCode1 <> "" And City1 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vCity" & iParamNo & " VARCHAR(100)"
					sParams = sParams & ",@vCity" & iParamNo & " = N'" & City1 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CWP.WorkingPlaceCity LIKE '%' + @vCity" & iParamNo & " + '%' "
				End If

				iParamNo = iParamNo + 1
			End If

			If sTemp <> "" Then
				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = "(" & sTemp & ") "
			End If

			sTemp = ""
			If AreaCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAreaCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAreaCode" & iParamNo & " = N'" & AreaCode2 & "'"

				sTemp = "AREA.AreaCode = @vAreaCode" & iParamNo & " "

				If PrefectureCode2 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vPrefectureCode" & iParamNo & " VARCHAR(3)"
					sParams = sParams & ",@vPrefectureCode" & iParamNo & " = N'" & PrefectureCode2 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CWP.WorkingPlacePrefectureCode = @vPrefectureCode" & iParamNo & " "
				End If

				If PrefectureCode2 <> "" And City2 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vCity" & iParamNo & " VARCHAR(200)"
					sParams = sParams & ",@vCity" & iParamNo & " = N'" & City2 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CWP.WorkingPlaceCity LIKE '%' + @vCity" & iParamNo & " + '%' "
				End If

				iParamNo = iParamNo + 1
			End If

			If sTemp <> "" Then
				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "(" & sTemp & ") "
			End If

			sJoin = sJoin & "INNER JOIN ( "
			sJoin = sJoin & "SELECT DISTINCT CWP.OrderCode "
			sJoin = sJoin & "FROM C_Info AS CWP "
			sJoin = sJoin & "INNER JOIN Area AS AREA ON CWP.WorkingPlacePrefectureCode = AREA.PrefectureCode "
			sJoin = sJoin & "WHERE " & sTemp2 & " "
			sJoin = sJoin & ") AS CWP "
			sJoin = sJoin & "ON VWOC.OrderCode = CWP.OrderCode "
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

			sJoin = sJoin & "INNER JOIN ( "
			sJoin = sJoin & "SELECT DISTINCT CWT.OrderCode "
			sJoin = sJoin & "FROM C_WorkingType AS CWT "
			sJoin = sJoin & "WHERE CWT.WorkingTypeCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CWT "
			sJoin = sJoin & "ON VWOC.OrderCode = CWT.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'希望勤務形態 end
		'******************************************************************************

		'******************************************************************************
		'希望業種 start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If IndustryTypeCode1 & IndustryTypeCode2 & IndustryTypeCode3 <> "" Then
			If IndustryTypeCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vIndustryTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vIndustryTypeCode" & iParamNo & " = N'" & IndustryTypeCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vIndustryTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If IndustryTypeCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vIndustryTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vIndustryTypeCode" & iParamNo & " = N'" & IndustryTypeCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vIndustryTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If IndustryTypeCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vIndustryTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vIndustryTypeCode" & iParamNo & " = N'" & IndustryTypeCode3 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vIndustryTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ( "
			sJoin = sJoin & "SELECT CIDST.CompanyCode "
			sJoin = sJoin & "FROM CompanyInfo AS CIDST "
			sJoin = sJoin & "WHERE CIDST.IndustryType IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CIDST "
			sJoin = sJoin & "ON VWOC.CompanyCode = CIDST.CompanyCode "
		End If
		'------------------------------------------------------------------------------
		'希望業種 end
		'******************************************************************************

		'******************************************************************************
		'特徴 start
		'------------------------------------------------------------------------------
		'未経験歓迎、語学を活かす、UIターン、休日１２０日以上
		sTemp = ""
		If InexperiencedPersonFlag = "1" Or UtilizeLanguageFlag = "1" Or UITurnFlag = "1" Or ManyHolidayFlag = "1" Then
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

			sJoin = sJoin & "INNER JOIN C_SupplementInfo AS CSP ON VWOC.OrderCode = CSP.OrderCode AND " & sTemp & " "
		End If

		'フレックスタイム
		sTemp = ""
		If FlexFlag = "1" Then
			sJoin = sJoin & "INNER JOIN CompanyInfo AS CMPFLEX ON VWOC.CompanyCode = CMPFLEX.CompanyCode AND CMPFLEX.CompanyKbn = '1' AND CMPFLEX.FlexTime = 'ON' "
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
		If YearlyIncome & MonthlyIncome & DailyIncome & HourlyIncome & PercentagePayFlag <> "" Then
			If YearlyIncome <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vYearlyIncome INT "
				sParams = sParams & ",@vYearlyIncome = " & YearlyIncome

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "CSLY.YearlyIncomeMin >= @vYearlyIncome "
			End If

			If MonthlyIncome <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vMonthlyIncome INT "
				sParams = sParams & ",@vMonthlyIncome = " & MonthlyIncome

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "CSLY.MonthlyIncomeMin >= @vMonthlyIncome "
			End If

			If DailyIncome <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDailyIncome INT "
				sParams = sParams & ",@vDailyIncome = " & DailyIncome

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "CSLY.DailyIncomeMin >= @vDailyIncome "
			End If

			If HourlyIncome <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vHourlyIncome INT "
				sParams = sParams & ",@vHourlyIncome = " & HourlyIncome

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "CSLY.HourlyIncomeMin >= @vHourlyIncome "
			End If

			If sTemp <> "" Then sTemp = "(" & sTemp & ") "

			'歩合制
			If PercentagePayFlag <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPercentagePayFlag VARCHAR(1)"
				sParams = sParams & ",@vPercentagePayFlag = N'" & PercentagePayFlag & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = "CSLY.PercentagePayFlag = @vPercentagePayFlag "
			End If

			sJoin = sJoin & "INNER JOIN C_Info AS CSLY ON VWOC.OrderCode = CSLY.OrderCode AND " & sTemp & " "
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
				sTemp = sTemp & "CWTM.WorkStartTime >= @vWorkStartHour + @vWorkStartMinute "
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
				sTemp = sTemp & "CWTM.WorkEndTime <= @vWorkEndHour + @vWorkEndMinute "
			End If

			If WorkStartHour <> "" And WorkEndHour <> "" Then
				If WorkStartHour < WorkEndHour Then
					'勤務開始時間 < 勤務終了時間の場合、夜間の業務時間を除くようにする
					sTemp2 = "AND CWTM.WorkStartTime < CWTM.WorkEndTime "
				End If
			End If

			sJoin = sJoin & "INNER JOIN C_WorkingCondition AS CWTM ON VWOC.OrderCode = CWTM.OrderCode AND " & sTemp & sTemp2
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

			sJoin = sJoin & "INNER JOIN C_Info AS CWHT ON VWOC.OrderCode = CWHT.OrderCode AND " & sTemp
		End If
		'------------------------------------------------------------------------------
		'週休 end
		'******************************************************************************

		'******************************************************************************
		'年齢 start
		'------------------------------------------------------------------------------
		sTemp = ""
		If Age <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vAge INT "
			sParams = sParams & ",@vAge = " & Age

			sTemp = sTemp & "(@vAge BETWEEN ISNULL(CAGE.AgeMin, 0) AND ISNULL(CAGE.AgeMax, 255)) "

			sJoin = sJoin & "INNER JOIN C_Info AS CAGE ON VWOC.OrderCode = CAGE.OrderCode AND " & sTemp
		End If
		'------------------------------------------------------------------------------
		' 年齢 end
		'******************************************************************************

		'******************************************************************************
		'契約期間 start
		'------------------------------------------------------------------------------
		sTemp = ""
		If IsRE(AgreementTerm, "^[123]$", True) = True Then
			If AgreementTerm = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WHERE WorkPeriod <= 1 UNION SELECT OrderCode FROM C_Undertake WHERE WorkPeriod <= 1 UNION SELECT OrderCode FROM C_TTP WHERE WorkPeriod <= 1) AS CAT ON VWOC.OrderCode = CAT.OrderCode "
			ElseIf AgreementTerm = "2" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WHERE WorkPeriod <= 2 UNION SELECT OrderCode FROM C_Undertake WHERE WorkPeriod <= 2 UNION SELECT OrderCode FROM C_TTP WHERE WorkPeriod <= 2) AS CAT ON VWOC.OrderCode = CAT.OrderCode "
			ElseIf AgreementTerm = "3" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WHERE WorkPeriod > 3 UNION SELECT OrderCode FROM C_Undertake WHERE WorkPeriod > 3 UNION SELECT OrderCode FROM C_TTP WHERE WorkPeriod > 3) AS CAT ON VWOC.OrderCode = CAT.OrderCode "
			End If
		End If
		'------------------------------------------------------------------------------
		'契約期間 end
		'******************************************************************************

		'******************************************************************************
		'保有資格 start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If LicenseGroupCode1 <> "" Then
			'大分類
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vLicenseGroupCode" & iParamNo & " VARCHAR(2)"
			sParams = sParams & ",@vLicenseGroupCode" & iParamNo & " = N'" & LicenseGroupCode1 & "'"

			If sTemp <> "" Then sTemp = sTemp & "AND "
			sTemp = sTemp & "CL.GroupCode = @vLicenseGroupCode" & iParamNo & " "

			If LicenseCategoryCode1 <> "" Then
				'中分類
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLicenseCategoryCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vLicenseCategoryCode" & iParamNo & " = N'" & LicenseCategoryCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CL.CategoryCode = @vLicenseCategoryCode" & iParamNo & " "

				If LicenseCode1 <> "" Then
					'小分類
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLicenseCode" & iParamNo & " VARCHAR(2)"
					sParams = sParams & ",@vLicenseCode" & iParamNo & " = N'" & LicenseCode1 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CL.Code = @vLicenseCode" & iParamNo & " "
				End If
			End If

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CL.OrderCode FROM C_License AS CL WHERE " & sTemp & ") AS CL ON VWOC.OrderCode = CL.OrderCode "
			iParamNo = iParamNo + 1
		End If
		'------------------------------------------------------------------------------
		'保有資格 end
		'******************************************************************************

		'******************************************************************************
		'スキル start
		'------------------------------------------------------------------------------
		'OS
		sTemp = ""
		iParamNo = 0
		If OSCode1 <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo & " VARCHAR(20), @vSkillCode" & iParamNo & " VARCHAR(3) "
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo & " = N'OS',@vSkillCode" & iParamNo & " = N'" & OSCode1 & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CSKL.OrderCode FROM C_Skill AS CSKL WHERE CSKL.CategoryCode = @vSkillCategoryCode" & iParamNo & " AND CSKL.Code = @vSkillCode" & iParamNo & ") AS CSKL" & iParamNo & " ON VWOC.OrderCode = CSKL" & iParamNo & ".OrderCode "
			iParamNo = iParamNo + 1
		End If

		'アプリケーション
		sTemp = ""
		If ApplicationCode1 <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo & " VARCHAR(20), @vSkillCode" & iParamNo & " VARCHAR(3) "
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo & " = N'Application',@vSkillCode" & iParamNo & " = N'" & ApplicationCode1 & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CSKL.OrderCode FROM C_Skill AS CSKL WHERE CSKL.CategoryCode = @vSkillCategoryCode" & iParamNo & " AND CSKL.Code = @vSkillCode" & iParamNo & ") AS CSKL" & iParamNo & " ON VWOC.OrderCode = CSKL" & iParamNo & ".OrderCode "
			iParamNo = iParamNo + 1
		End If

		'開発言語
		sTemp = ""
		If DevelopmentLanguageCode1 <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo & " VARCHAR(20), @vSkillCode" & iParamNo & " VARCHAR(3) "
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo & " = N'DevelopmentLanguage',@vSkillCode" & iParamNo & " = N'" & DevelopmentLanguageCode1 & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CSKL.OrderCode FROM C_Skill AS CSKL WHERE CSKL.CategoryCode = @vSkillCategoryCode" & iParamNo & " AND CSKL.Code = @vSkillCode" & iParamNo & ") AS CSKL" & iParamNo & " ON VWOC.OrderCode = CSKL" & iParamNo & ".OrderCode "
			iParamNo = iParamNo + 1
		End If

		'データベース
		sTemp = ""
		If DatabaseCode1 <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo & " VARCHAR(20), @vSkillCode" & iParamNo & " VARCHAR(3) "
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo & " = N'Database',@vSkillCode" & iParamNo & " = N'" & DatabaseCode1 & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CSKL.OrderCode FROM C_Skill AS CSKL WHERE CSKL.CategoryCode = @vSkillCategoryCode" & iParamNo & " AND CSKL.Code = @vSkillCode" & iParamNo & ") AS CSKL" & iParamNo & " ON VWOC.OrderCode = CSKL" & iParamNo & ".OrderCode "
			iParamNo = iParamNo + 1
		End If
		'------------------------------------------------------------------------------
		'スキル end
		'******************************************************************************

		'******************************************************************************
		'キーワード start
		'------------------------------------------------------------------------------
		sTemp = ""
		If Keyword <> "" Then
			aValue = Split(Replace(Keyword, "　", " "), " ")
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

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY CFTN.OrderCode) AS Num, CFTN.OrderCode FROM C_FullTextNavi AS CFTN WHERE CONTAINS(CFTN.Text, @vKeyword)) AS CFTN ON VWOC.OrderCode = CFTN.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'キーワード end
		'******************************************************************************

		'******************************************************************************
		'情報コード start
		'------------------------------------------------------------------------------
		If OrderCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vOrderCode VARCHAR(8) "
			sParams = sParams & ",@vOrderCode = N'" & OrderCode & "'"

			sJoin = ""
			sWhere = "WHERE VWOC.OrderCode = @vOrderCode "
		End If
		'------------------------------------------------------------------------------
		'情報コード end
		'******************************************************************************

		'******************************************************************************
		'前回表示時の最新情報コード start
		'------------------------------------------------------------------------------
		If BOC <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vBeforeOrderCode VARCHAR(8) "
			sParams = sParams & ",@vBeforeOrderCode = N'" & BOC & "'"

			sWhere = "WHERE VWOC.OrderCode > @vBeforeOrderCode "
		End If
		'------------------------------------------------------------------------------
		'前回表示時の最新情報コード end
		'******************************************************************************

		If flgEasySearch = False And sJoin & sWhere <> "" Then
			If CStr(Top) <> "" Then Top = "TOP " & Top
			GetSQLOrderSearchDetail = "" & _
				"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & _
				"SELECT " & Top & " VWOC.OrderCode " & _
				",VWOC.SortNum " & _
				",VWOC.UpdateDay " & _
				"FROM vw_OrderCode AS VWOC " & _
				sJoin & _
				sWhere & _
				"ORDER BY VWOC.SortNum ASC, VWOC.UpdateDay DESC "

			GetSQLOrderSearchDetail = "" & _
				"/*ナビ・求人票詳細検索*/ " & _
				"EXEC sp_executesql N'" & Replace(GetSQLOrderSearchDetail, "'", "''") & "'"
			If sDeclare <> "" Then GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & ",N'" & sDeclare & "'" & sParams
		Else
			GetSQLOrderSearchDetail = GetSQLOrderSearchEasy()
		End If

		If sSearchCondition <> "" Then
			sSearchCondition = "<table class=""pattern1"" border=""0"" style=""width:600px;""><thead><tr><th colspan=""2"" style=""width:588px;"">検索条件</th></tr></thead><tbody>" & sSearchCondition & "</tbody></table>"
		Else
			sSearchCondition = "なし"
		End If
	End Function

	'******************************************************************************
	'概　要：求人票検索ＬＯＧ書き込みＳＱＬを取得
	'作成者：Lis Kokubo
	'作成日：2007/04/04
	'引　数：
	'備　考：
	'******************************************************************************
	Public Function GetSQLWriteLog()
		Dim sTmpJT
		sTmpJT = JT
		If JT2 = "" Then sTmpJT = JT2

		If flgEasySearch = True Then
			'カンタン検索ログ
			GetSQLWriteLog = "EXEC up_Reg_LOG_SearchOrder '" & G_USERID & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'" & _
				",'" & ChkSQLStr(Session.SessionID) & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("URL")) & "?" & ChkSQLStr(Request.ServerVariables("QUERY_STRING")) & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("HTTP_REFERER")) & "'" & _
				",'" & sTmpJT & "'" & _
				",'" & WT & "'" & _
				",'" & AC & "'" & _
				",'" & AC2 & "'" & _
				",'" & Specialty & "'" & _
				",''" & _
				",'" & RC & "'" & _
				",'" & SC & "'" & _
				",'" & KW & "'" & _
				",'" & Replace(SQLOrderSearch, "'", "''") & "'"
		Else
			'詳細検索ログ
			GetSQLWriteLog = "EXEC up_Reg_LOG_SearchOrderDetail '" & G_USERID & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'" & _
				",'" & ChkSQLStr(Session.SessionID) & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("URL")) & "?" & ChkSQLStr(Request.ServerVariables("QUERY_STRING")) & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("HTTP_REFERER")) & "'" & _
				",'" & JobTypeCode1 & "'" & _
				",'" & JobTypeCode2 & "'" & _
				",'" & RailwayLineCode1 & "'" & _
				",'" & StationCode1 & "'" & _
				",'" & RailwayLineCode2 & "'" & _
				",'" & StationCode2 & "'" & _
				",'" & AreaCode1 & "'" & _
				",'" & PrefectureCode1 & "'" & _
				",'" & City1 & "'" & _
				",'" & AreaCode2 & "'" & _
				",'" & PrefectureCode2 & "'" & _
				",'" & City2 & "'" & _
				",'" & WorkingTypeCode1 & "'" & _
				",'" & WorkingTypeCode2 & "'" & _
				",'" & WorkingTypeCode3 & "'" & _
				",'" & IndustryTypeCode1 & "'" & _
				",'" & IndustryTypeCode2 & "'" & _
				",'" & IndustryTypeCode3 & "'" & _
				",'" & PercentagePayFlag & "'" & _
				",'" & YearlyIncome & "'" & _
				",'" & MonthlyIncome & "'" & _
				",'" & DailyIncome & "'" & _
				",'" & HourlyIncome & "'" & _
				",'" & WorkStartHour & WorkStartMinute & "'" & _
				",'" & WorkEndHour & WorkEndMinute & "'" & _
				",'" & WeeklyHolidayType & "'" & _
				",'" & Age & "'" & _
				",'" & AgreementTerm & "'" & _
				",'" & LicenseGroupCode1 & "'" & _
				",'" & LicenseCategoryCode1 & "'" & _
				",'" & LicenseCode1 & "'" & _
				",'" & OSCode1 & "'" & _
				",'" & ApplicationCode1 & "'" & _
				",'" & DevelopmentLanguageCode1 & "'" & _
				",'" & DatabaseCode1 & "'" & _
				",'" & InexperiencedPersonFlag & "'" & _
				",'" & UtilizeLanguageFlag & "'" & _
				",'" & UITurnFlag & "'" & _
				",'" & ManyHolidayFlag & "'" & _
				",'" & FlexFlag & "'" & _
				",'" & Keyword & "'" & _
				",'" & Replace(SQLOrderSearch, "'", "''") & "'"
		End If
	End Function

	'******************************************************************************
	'概　要：求人票カンタン検索ＳＱＬを取得
	'作成者：Lis Kokubo
	'作成日：2007/04/04
	'引　数：
	'備　考：
	'******************************************************************************
	Function GetSQLOrderSearchEasy()
		Dim sJoin		: sJoin = ""
		Dim sWhere		: sWhere = ""
		Dim sDeclare	: sDeclare = ""
		Dim sParams		: sParams = ""
		Dim iParamNo
		Dim sFrom
		Dim sTemp
		Dim sTemp2
		Dim sTemp3
		Dim aValue
		Dim idx
		Dim sSearchCondition

		GetSQLOrderSearchEasy = ""

		'******************************************************************************
		'職種 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If JT & JT2 <> "" Then
			If JT2 <> "" Then
				sTemp = JobTypeCode2
				If Len(JobTypeCode2) < 7 Then sTemp = sTemp & "%"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & JT2 & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "CJT.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " "

				iParamNo = iParamNo + 1
			ElseIf JT <> "" Then
				sTemp = JobTypeCode1
				If Len(JobTypeCode1) < 7 Then sTemp = sTemp & "%"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & JT & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "CJT.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CJT.OrderCode FROM C_JobType AS CJT WHERE (" & sTemp2 & ")) AS CJT ON VWOC.OrderCode = CJT.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'職種 end
		'******************************************************************************

		'******************************************************************************
		'希望勤務地 start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If AC & AC2 <> "" Then
			sTemp = ""
			If AC <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAreaCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAreaCode" & iParamNo & " = N'" & AC & "'"

				sTemp = "AREA.AreaCode = @vAreaCode" & iParamNo & " "
			End If

			If AC2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPrefectureCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vPrefectureCode" & iParamNo & " = N'" & AC2 & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CWP.WorkingPlacePrefectureCode = @vPrefectureCode" & iParamNo & " "
			End If

			iParamNo = iParamNo + 1

			sJoin = sJoin & "INNER JOIN ( "
			sJoin = sJoin & "SELECT DISTINCT CWP.OrderCode "
			sJoin = sJoin & "FROM C_Info AS CWP "
			sJoin = sJoin & "INNER JOIN Area AS AREA ON CWP.WorkingPlacePrefectureCode = AREA.PrefectureCode "
			sJoin = sJoin & "WHERE " & sTemp & " "
			sJoin = sJoin & ") AS CWP "
			sJoin = sJoin & "ON VWOC.OrderCode = CWP.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'希望勤務地 end
		'******************************************************************************

		'******************************************************************************
		'特徴 start
		'------------------------------------------------------------------------------
		'未経験歓迎、語学を活かす、UIターン、休日１２０日以上
		sTemp = ""
		If Len(ST) >= 6 Then
			'未経験歓迎
			If Mid(ST, 1, 1) = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.InexperiencedPersonFlag = '1' "
			End If

			'語学を活かす
			If Mid(ST, 2, 1) = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.UtilizeLanguageFlag = '1' "
			End If

			'UIターン
			If Mid(ST, 4, 1) = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.UITurnFlag = '1' "
			End If

			'休日１２０日以上
			If Mid(ST, 5, 1) = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.ManyHolidayFlag = '1' "
			End If

			sJoin = sJoin & "INNER JOIN C_SupplementInfo AS CSP ON VWOC.OrderCode = CSP.OrderCode AND " & sTemp & " "

			'フレックスタイム
			sTemp = ""
			If Mid(ST, 6, 1) = "1" Then
				sJoin = sJoin & "INNER JOIN CompanyInfo AS CMPFLEX ON VWOC.CompanyCode = CMPFLEX.CompanyCode AND CMPFLEX.CompanyKbn = '1' AND CMPFLEX.FlexTime = 'ON' "
			End If

'			'派遣
'			If Mid(ST, 3, 1) = "1" Then
'				If InStr(sJoin, "INNER JOIN C_WorkingType AS CWT") = 0 Then sJoin = sJoin & "INNER JOIN C_WorkingType AS CWT ON CI.OrderCode = CWT.OrderCode "
'				If sWhere <> "" Then sWhere = sWhere & "AND "
'				sWhere = sWhere & "CWT.WorkingTypeCode IN ('001', '004') "
'			End If
		End If
		'------------------------------------------------------------------------------
		'特徴 end
		'******************************************************************************

		'******************************************************************************
		'沿線 start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If RC <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vRailwayLineCode" & iParamNo & " VARCHAR(7)"
			sParams = sParams & ",@vRailwayLineCode" & iParamNo & " = N'" & RC & "'"

			If sTemp <> "" Then sTemp = sTemp & ","
			sTemp = sTemp & "@vRailwayLineCode" & iParamNo

			iParamNo = iParamNo + 1

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT CNS.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS CNS "
			sJoin = sJoin & "INNER JOIN StationStop AS SS "
			sJoin = sJoin & "ON CNS.StationCode = SS.StationCode "
			sJoin = sJoin & "INNER JOIN B_RailwayLine AS BRL "
			sJoin = sJoin & "ON SS.RailwayLineCode = BRL.RailwayLineCode "
			sJoin = sJoin & "AND BRL.RailwayLineCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CRL "
			sJoin = sJoin & "ON VWOC.OrderCode = CRL.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'沿線 end
		'******************************************************************************

		'******************************************************************************
		'駅 start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If SC <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vStationCode" & iParamNo & " VARCHAR(7)"
			sParams = sParams & ",@vStationCode" & iParamNo & " = N'" & SC & "'"

			If sTemp <> "" Then sTemp = sTemp & ","
			sTemp = sTemp & "@vStationCode" & iParamNo

			iParamNo = iParamNo + 1

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT CNS.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS CNS "
			sJoin = sJoin & "WHERE CNS.StationCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CNS "
			sJoin = sJoin & "ON VWOC.OrderCode = CNS.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'駅 end
		'******************************************************************************

		'******************************************************************************
		'キーワード start
		'------------------------------------------------------------------------------
		sTemp = ""
		If KW <> "" Then
			aValue = Split(Replace(KW, "　", " "), " ")
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

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY CFTN.OrderCode) AS Num, CFTN.OrderCode FROM C_FullTextNavi AS CFTN WHERE CONTAINS(CFTN.Text, @vKeyword)) AS CFTN ON VWOC.OrderCode = CFTN.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'キーワード end
		'******************************************************************************

		If CStr(Top) <> "" Then Top = "TOP " & Top
		GetSQLOrderSearchEasy = "" & _
			"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & _
			"SELECT " & Top & " VWOC.OrderCode " & _
			",VWOC.SortNum " & _
			",VWOC.UpdateDay " & _
			"FROM vw_OrderCode AS VWOC " & _
			sJoin & _
			sWhere & _
			"ORDER BY VWOC.SortNum ASC, VWOC.UpdateDay DESC "

		GetSQLOrderSearchEasy = "" & _
			"/*ナビ・求人票カンタン検索*/ " & _
			"EXEC sp_executesql N'" & Replace(GetSQLOrderSearchEasy, "'", "''") & "'"
		If sDeclare <> "" Then GetSQLOrderSearchEasy = GetSQLOrderSearchEasy & ",N'" & sDeclare & "'" & sParams
	End Function

	'******************************************************************************
	'概　要：求人票詳細検索条件出力ＨＴＭＬを取得
	'作成者：Lis Kokubo
	'作成日：2007/04/04
	'引　数：
	'備　考：
	'******************************************************************************
	Public Function GetHtmlSearchCondition()
		Dim sTemp
		Dim sTemp2

		If flgEasySearch = True Then Exit Function

		GetHtmlSearchCondition = ""

		'職種
		sTemp2 = ""
		If JobTypeBigCode1 & JobTypeCode1 & JobTypeBigCode2 & JobTypeCode2 <> "" Then
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
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">職種</th><td style=""width:439px;"">" & sTemp2 & "</td></tr>"
		End If

		'勤務地
		sTemp = ""
		If AreaCode1 & PrefectureCode1 & City1 & RailwayLineCode1 & RailwayLineCode1 & AreaCode2 & PrefectureCode2 & City2 & RailwayLineCode2 & RailwayLineCode2 <> "" Then
			If AreaCode1 & PrefectureCode1 & City1 & RailwayLineCode1 & RailwayLineCode1 <> "" Then
				'エリア
				sTemp = sTemp & AreaName1

				'都道府県
				If PrefectureCode1 <> "" Then
					sTemp = sTemp & "　"
					sTemp = sTemp & PrefectureName1

					'市区郡
					If City1 <> "" Then
						sTemp = sTemp & "　"
						sTemp = sTemp & City1
					End If

					'沿線
					If RailwayLineCode1 <> "" Then
						sTemp = sTemp & "　"
						sTemp = sTemp & RailwayLineName1
					End If

					'駅
					If RailwayLineCode2 <> "" Then
						If sTemp <> "" Then sTemp = sTemp & "　"
						sTemp = sTemp & StationName1 & "駅"
					End If
				End If
			End If

			If AreaCode2 & PrefectureCode2 & City2 & RailwayLineCode2 & RailwayLineCode2 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				'エリア
				sTemp = sTemp & AreaName2

				'都道府県
				If PrefectureCode2 <> "" Then
					sTemp = sTemp & "　"
					sTemp = sTemp & PrefectureName2

					'市区郡
					If City2 <> "" Then
						sTemp = sTemp & "　"
						sTemp = sTemp & City2
					End If

					'沿線
					If RailwayLineCode2 <> "" Then
						sTemp = sTemp & "　"
						sTemp = sTemp & RailwayLineName2
					End If

					'駅
					If RailwayLineCode2 <> "" Then
						If sTemp <> "" Then sTemp = sTemp & "　"
						sTemp = sTemp & StationName2 & "駅"
					End If
				End If
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">勤務地</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
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
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">勤務形態</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'業種
		sTemp = ""
		If IndustryTypeCode1 & IndustryTypeCode2 & IndustryTypeCode3 <> "" Then
			If IndustryTypeCode1 <> "" Then sTemp = sTemp & IndustryTypeName1
			If IndustryTypeCode2 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "　"
				sTemp = sTemp & IndustryTypeName2
			End If
			If IndustryTypeCode3 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "　"
				sTemp = sTemp & IndustryTypeName3
			End If
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">業種</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'歩合制
		sTemp = ""
		If PercentagePayFlag & YearlyIncome & MonthlyIncome & DailyIncome & HourlyIncome <> "" Then
			If PercentagePayFlag = "1" Then
				sTemp = sTemp & "歩合制あり"
			ElseIf PercentagePayFlag = "0" Then
				sTemp = sTemp & "歩合制なし"
			End If
			If YearlyIncome <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				sTemp = sTemp & "年収：" & YearlyIncome & "～"
			End If
			If MonthlyIncome <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				sTemp = sTemp & "月収：" & MonthlyIncome & "～"
			End If
			If DailyIncome <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				sTemp = sTemp & "日給：" & DailyIncome & "～"
			End If
			If HourlyIncome <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				sTemp = sTemp & "時給：" & HourlyIncome & "～"
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">給与</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'特徴
		sTemp = ""
		If InexperiencedPersonFlag & UtilizeLanguageFlag & TempFlag & UITurnFlag & ManyHolidayFlag & FlexFlag <> "" Then
			If InexperiencedPersonFlag = "1" Then sTemp = sTemp & "「未経験者ＯＫ」"
			If UtilizeLanguageFlag = "1" Then sTemp = sTemp & "「語学を活かす」"
			If TempFlag = "1" Then sTemp = sTemp & "「派遣」"
			If UITurnFlag = "1" Then sTemp = sTemp & "「ＵＩターン歓迎」"
			If ManyHolidayFlag = "1" Then sTemp = sTemp & "「休日１２０日以上」"
			If FlexFlag = "1" Then sTemp = sTemp & "「フレックス」"

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">特徴</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'就業時間
		sTemp = ""
		If WorkStartHour & WorkStartMinute & WorkEndHour & WorkEndMinute <> "" Then
			If WorkStartHour & WorkStartMinute <> "" Then sTemp = sTemp & "就業開始時間：" & WorkStartHour & ":" & WorkStartMinute & "&nbsp;以降"
			If WorkEndHour & WorkEndMinute <> "" And sTemp <> "" Then sTemp = sTemp & "<br>"
			If WorkEndHour & WorkEndMinute <> "" Then sTemp = sTemp & "就業終了時間：" & WorkEndHour & ":" & WorkEndMinute & "&nbsp;以前"

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">就業時間</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'週休種類
		sTemp = ""
		If WeeklyHolidayType <> "" Then
			sTemp = sTemp & WeeklyHolidayTypeName
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">週休種類</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'年齢
		sTemp = ""
		If Age <> "" Then
			sTemp = sTemp & Age & "歳を含む"
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">年齢</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
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

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">契約期間</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'資格
		sTemp = ""
		If LicenseGroupCode1 & LicenseCategoryCode1 & LicenseName1 <> "" Then
			sTemp = LicenseName1
			If sTemp = "" And LicenseCategoryName1 <> "" Then sTemp = LicenseCategoryName1 & "関連"
			If sTemp = "" And LicenseGroupName1 <> "" Then sTemp = LicenseGroupName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">資格</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'ＯＳ
		sTemp = ""
		If OSName1 <> "" Then
			sTemp = sTemp & OSName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">ＯＳ</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'アプリケーション
		sTemp = ""
		If ApplicationName1 <> "" Then
			sTemp = sTemp & ApplicationName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">アプリケーション</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'開発言語
		sTemp = ""
		If DevelopmentLanguageName1 <> "" Then
			sTemp = sTemp & DevelopmentLanguageName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">開発言語</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'データベース
		sTemp = ""
		If DatabaseName1 <> "" Then
			sTemp = sTemp & DatabaseName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">データベース</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'キーワード
		sTemp = ""
		If Keyword <> "" Then
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">キーワード</th><td style=""width:439px;"">" & Keyword & "</td></tr>"
		End If

		'情報コード（検索）
		If OrderCode <> "" Then
			GetHtmlSearchCondition = "<tr><th style=""width:138px;"">情報コード</th><td style=""width:439px;"">" & OrderCode & "</td></tr>"
		End If

		If GetHtmlSearchCondition <> "" Then
			GetHtmlSearchCondition = "<table class=""pattern1"" border=""0"" style=""width:600px;""><thead><tr><th colspan=""2"" style=""width:588px;"">検索条件</th></tr></thead><tbody>" & GetHtmlSearchCondition & "</tbody></table>"
		End If

	End Function
End Class
%>
