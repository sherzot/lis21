<%
'******************************************************************************
'概　要：企業テーブルにデータをInsert, Updateする時に
'　　　：formで飛んできたデータを格納するためのクラス群
'備　考：事前に commonfunc.asp をインクルードしておくこと！
'更　新：2006/05/13 LIS K.Kokubo 作成
'******************************************************************************
%>
<%
'******************************************************************************
'名　称：clsCompanyInfo
'概　要：formで飛んできたCompanyInfoテーブル用のデータを持つためのクラス
'備　考：
'更　新：2006/03/24 LIS K.Kokubo 作成
'　　　：2008/06/05 LIS K.Kokubo GetData関数作成
'　　　：2008/06/05 LIS K.Kokubo ChkData関数作成
'　　　：2008/08/14 LIS M.Hayashi 特徴フラグの追加とフレックス移動
'　　　：2009/01/05 LIS K.Kokubo 福利厚生備考追加
'　　　：2010/01/06 LIS K.Kokubo 会社の雰囲気追加
'******************************************************************************
Class clsCompanyInfo
	Public CompanyCode
	Public CompanyKbn
	Public CompanyName_K
	Public CompanyName_F
	Public EstablishYear
	Public IndustryType
	Public CapitalAmount
	Public ForeinCapital
	Public ListClass
	Public AllEmployeeNum
	Public ManEmployeeNum
	Public WomanEmployeeNum
	Public HomepageAddress
	Public Post_U
	Public Post_L
	Public PrefectureCode
	Public City_K
	Public City_F
	Public Town
	Public Address
	Public TelephoneNumber
	Public StationCode1
	Public StationName1
	Public CompanySyudan1_1
	Public WorkOrBus1
	Public CompanySyudan1_2
	Public WorkBusTime1
	Public StationCode2
	Public StationName2
	Public CompanySyudan2_1
	Public WorkOrBus2
	Public CompanySyudan2_2
	Public WorkBusTime2
	Public SocietyInsurance
	Public Sanatorium
	Public EnterprisePension
	Public WealthShape
	Public StockOption
	Public RetirementPay
	Public ResidencePay
	Public FamilyPay
	Public EmployeeDormitory
	Public CompanyHouse
	Public NewEmployeeTraining
	Public OverseasTraining
	Public OtherTraning
	'Public FlexTime	'2008/08/14 Lis林 DEL
	Public WelfareProgramRemark '2009/01/05 LIS K.Kokubo ADD
	Public BusinessContents
	Public CompanyPR
	Public Simebi
	Public ContactPersonName
	Public Tanto1Yakusyoku
	Public Tanto2Name
	Public Tanto2Yakusyoku
	Public MailAddr
	Public NewJobMail
	Public DemandPrefectureCode
	Public DemandCity_K
	Public DemandCity_F
	Public DemandTown
	Public DemandAddress
	Public DemandSectionName
	Public DemandPersonName
	Public Atmosphere
	Public IsData
	Public MaxIndex
	Public Err
	Public ErrStyle

	'******************************************************************************
	'概　要：clsCompanyInfoクラスの初期化関数
	'引　数：
	'戻り値：×
	'備　考：
	'更　新：2006/03/24 LIS K.Kokubo 作成
	'******************************************************************************
	Private Sub Class_Initialize()
		MaxIndex = -1
		IsData = False

		Err = ""

		Set ErrStyle = Server.CreateObject("Scripting.Dictionary")
		ErrStyle.CompareMode = 1
	End Sub

	'******************************************************************************
	'概　要：データ取得
	'引　数：
	'戻り値：×
	'備　考：POSTデータを読み取り、各プロパティにデータを設定する
	'更　新：2008/06/05 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function GetData()
		If GetForm("CONF_CompanyCode", 1) <> "" Then CompanyCode = GetForm("CONF_CompanyCode", 1)
		If GetForm("CONF_CompanyKbn", 1) <> "" Then CompanyKbn = GetForm("CONF_CompanyKbn", 1)
		If GetForm("CONF_CompanyName_K", 1) <> "" Then CompanyName_K = GetForm("CONF_CompanyName_K", 1)
		If GetForm("CONF_CompanyName_F", 1) <> "" Then CompanyName_F = GetForm("CONF_CompanyName_F", 1)
		If GetForm("CONF_EstablishYear", 1) <> "" Then EstablishYear = GetForm("CONF_EstablishYear", 1)
		If GetForm("CONF_IndustryType", 1) <> "" Then IndustryType = GetForm("CONF_IndustryType", 1)
		If GetForm("CONF_CapitalAmount", 1) <> "" Then CapitalAmount = GetForm("CONF_CapitalAmount", 1)
		If GetForm("CONF_ForeinCapital", 1) <> "" Then ForeinCapital = GetForm("CONF_ForeinCapital", 1)
		If GetForm("CONF_ListClass", 1) <> "" Then ListClass = GetForm("CONF_ListClass", 1)
		If GetForm("CONF_AllEmployeeNum", 1) <> "" Then AllEmployeeNum = GetForm("CONF_AllEmployeeNum", 1)
		If GetForm("CONF_ManEmployeeNum", 1) <> "" Then ManEmployeeNum = GetForm("CONF_ManEmployeeNum", 1)
		If GetForm("CONF_WomanEmployeeNum", 1) <> "" Then WomanEmployeeNum = GetForm("CONF_WomanEmployeeNum", 1)
		If GetForm("CONF_HomepageAddress", 1) <> "" Then HomepageAddress = GetForm("CONF_HomepageAddress", 1)
		If GetForm("CONF_Post_U", 1) <> "" Then Post_U = GetForm("CONF_Post_U", 1)
		If GetForm("CONF_Post_L", 1) <> "" Then Post_L = GetForm("CONF_Post_L", 1)
		If GetForm("CONF_PrefectureCode", 1) <> "" Then PrefectureCode = GetForm("CONF_PrefectureCode", 1)
		If GetForm("CONF_City_K", 1) <> "" Then City_K = GetForm("CONF_City_K", 1)
		If GetForm("CONF_City_F", 1) <> "" Then City_F = GetForm("CONF_City_F", 1)
		If GetForm("CONF_Town", 1) <> "" Then Town = GetForm("CONF_Town", 1)
		If GetForm("CONF_Address", 1) <> "" Then Address = GetForm("CONF_Address", 1)
		If GetForm("CONF_TelephoneNumber", 1) <> "" Then TelephoneNumber = GetForm("CONF_TelephoneNumber", 1)
		If GetForm("CONF_StationCode1", 1) <> "" Then StationCode1 = GetForm("CONF_StationCode1", 1)
		If GetForm("CONF_StationName1", 1) <> "" Then StationName1 = GetForm("CONF_StationName1", 1)
		If GetForm("CONF_CompanySyudan1_1", 1) <> "" Then CompanySyudan1_1 = GetForm("CONF_CompanySyudan1_1", 1)
		If GetForm("CONF_WorkOrBus1", 1) <> "" Then WorkOrBus1 = GetForm("CONF_WorkOrBus1", 1)
		If GetForm("CONF_CompanySyudan1_2", 1) <> "" Then CompanySyudan1_2 = GetForm("CONF_CompanySyudan1_2", 1)
		If GetForm("CONF_WorkBusTime1", 1) <> "" Then WorkBusTime1 = GetForm("CONF_WorkBusTime1", 1)
		If GetForm("CONF_StationCode2", 1) <> "" Then StationCode2 = GetForm("CONF_StationCode2", 1)
		If GetForm("CONF_StationName2", 1) <> "" Then StationName2 = GetForm("CONF_StationName2", 1)
		If GetForm("CONF_CompanySyudan2_1", 1) <> "" Then CompanySyudan2_1 = GetForm("CONF_CompanySyudan2_1", 1)
		If GetForm("CONF_WorkOrBus2", 1) <> "" Then WorkOrBus2 = GetForm("CONF_WorkOrBus2", 1)
		If GetForm("CONF_CompanySyudan2_2", 1) <> "" Then CompanySyudan2_2 = GetForm("CONF_CompanySyudan2_2", 1)
		If GetForm("CONF_WorkBusTime2", 1) <> "" Then WorkBusTime2 = GetForm("CONF_WorkBusTime2", 1)
		If GetForm("CONF_SocietyInsurance", 1) <> "" Then SocietyInsurance = GetForm("CONF_SocietyInsurance", 1)
		If GetForm("CONF_Sanatorium", 1) <> "" Then Sanatorium = GetForm("CONF_Sanatorium", 1)
		If GetForm("CONF_EnterprisePension", 1) <> "" Then EnterprisePension = GetForm("CONF_EnterprisePension", 1)
		If GetForm("CONF_WealthShape", 1) <> "" Then WealthShape = GetForm("CONF_WealthShape", 1)
		If GetForm("CONF_StockOption", 1) <> "" Then StockOption = GetForm("CONF_StockOption", 1)
		If GetForm("CONF_RetirementPay", 1) <> "" Then RetirementPay = GetForm("CONF_RetirementPay", 1)
		If GetForm("CONF_ResidencePay", 1) <> "" Then ResidencePay = GetForm("CONF_ResidencePay", 1)
		If GetForm("CONF_FamilyPay", 1) <> "" Then FamilyPay = GetForm("CONF_FamilyPay", 1)
		If GetForm("CONF_EmployeeDormitory", 1) <> "" Then EmployeeDormitory = GetForm("CONF_EmployeeDormitory", 1)
		If GetForm("CONF_CompanyHouse", 1) <> "" Then CompanyHouse = GetForm("CONF_CompanyHouse", 1)
		If GetForm("CONF_NewEmployeeTraining", 1) <> "" Then NewEmployeeTraining = GetForm("CONF_NewEmployeeTraining", 1)
		If GetForm("CONF_OverseasTraining", 1) <> "" Then OverseasTraining = GetForm("CONF_OverseasTraining", 1)
		If GetForm("CONF_OtherTraning", 1) <> "" Then OtherTraning = GetForm("CONF_OtherTraning", 1)
		'If GetForm("CONF_FlexTime", 1) <> "" Then FlexTime = GetForm("CONF_FlexTime", 1)	'08/08/14 Lis林 DEL
		If GetForm("CONF_WelfareProgramRemark", 1) <> "" Then WelfareProgramRemark = GetForm("CONF_WelfareProgramRemark", 1)
		If GetForm("CONF_BusinessContents", 1) <> "" Then BusinessContents = GetForm("CONF_BusinessContents", 1)
		If GetForm("CONF_CompanyPR", 1) <> "" Then CompanyPR = GetForm("CONF_CompanyPR", 1)
		If GetForm("CONF_Simebi", 1) <> "" Then Simebi = GetForm("CONF_Simebi", 1)
		If GetForm("CONF_ContactPersonName", 1) <> "" Then ContactPersonName = GetForm("CONF_ContactPersonName", 1)
		If GetForm("CONF_Tanto1Yakusyoku", 1) <> "" Then Tanto1Yakusyoku = GetForm("CONF_Tanto1Yakusyoku", 1)
		If GetForm("CONF_Tanto2Name", 1) <> "" Then Tanto2Name = GetForm("CONF_Tanto2Name", 1)
		If GetForm("CONF_Tanto2Yakusyoku", 1) <> "" Then Tanto2Yakusyoku = GetForm("CONF_Tanto2Yakusyoku", 1)
		If GetForm("CONF_MailAddr", 1) <> "" Then MailAddr = GetForm("CONF_MailAddr", 1)
		If GetForm("CONF_NewJobMail", 1) <> "" Then NewJobMail = GetForm("CONF_NewJobMail", 1)
		If GetForm("CONF_DemandPrefectureCode", 1) <> "" Then DemandPrefectureCode = GetForm("CONF_DemandPrefectureCode", 1)
		If GetForm("CONF_DemandCity_K", 1) <> "" Then DemandCity_K = GetForm("CONF_DemandCity_K", 1)
		If GetForm("CONF_DemandCity_F", 1) <> "" Then DemandCity_F = GetForm("CONF_DemandCity_F", 1)
		If GetForm("CONF_DemandTown", 1) <> "" Then DemandTown = GetForm("CONF_DemandTown", 1)
		If GetForm("CONF_DemandAddress", 1) <> "" Then DemandAddress = GetForm("CONF_DemandAddress", 1)
		If GetForm("CONF_DemandSectionName", 1) <> "" Then DemandSectionName = GetForm("CONF_DemandSectionName", 1)
		If GetForm("CONF_DemandPersonName", 1) <> "" Then DemandPersonName = GetForm("CONF_DemandPersonName", 1)
		If GetForm("CONF_Atmosphere", 1) <> "" Then Atmosphere = GetForm("CONF_Atmosphere", 1)
	End Function

	'******************************************************************************
	'概　要：データの整合性チェック
	'引　数：
	'戻り値：×
	'備　考：エラー内容をErrプロパティに書き込み
	'更　新：2008/06/05 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function ChkData()
		IsData = False

		'企業名
		If CompanyName_K = "" Or ChkLen(CompanyName_K, 100) = False Then
			Call DicAdd(ErrStyle, "CompanyName_K", "background-color:#ffff00;")
			Err = Err & "企業名は半角１文字、全角２文字と数えて１００文字までです。<br>"
		End If
		'企業名カナ
		If CompanyName_F = "" Or ChkLen(CompanyName_F, 80) = False Then
			Call DicAdd(ErrStyle, "CompanyName_F", "background-color:#ffff00;")
			Err = Err & "企業名カナは半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'締め日
		If Simebi <> "" And ChkInt(Simebi) = True Then
			If CInt(Simebi) < 1 Or CInt(Simebi) > 31 Then
				Call DicAdd(ErrStyle, "Simebi", "background-color:#ffff00;")
				Err = Err & "締め日に半角数字で正しい日を入力して下さい。<br>"
			End If
		ElseIf Simebi <> "" And ChkInt(Simebi) = False Then
			Call DicAdd(ErrStyle, "Simebi", "background-color:#ffff00;")
			Err = Err & "締め日に半角数字で正しい日を入力して下さい。<br>"
		End If
		'設立年度
		If EstablishYear <> "" And IsDate(EstablishYear & "/01/01") = False Then
			Call DicAdd(ErrStyle, "EstablishYear", "background-color:#ffff00;")
			Err = Err & "設立年度に半角数字で正しい年を入力して下さい。<br>"
		End If
		'資本金
		If CapitalAmount <> "" And ChkLen(CapitalAmount, 40) = False Then
			Call DicAdd(ErrStyle, "CapitalAmount", "background-color:#ffff00;")
			Err = Err & "資本金は半角１文字、全角２文字と数えて４０文字までです。<br>"
		End If
		'外資
		If ForeinCapital <> "" And ChkLen(ForeinCapital, 12) = False Then
			Call DicAdd(ErrStyle, "ForeinCapital", "background-color:#ffff00;")
			Err = Err & "外資は半角１文字、全角２文字と数えて１２文字までです。<br>"
		End If
		'株式
		If ListClass <> "" And ChkLen(ListClass, 40) = False Then
			Call DicAdd(ErrStyle, "ListClass", "background-color:#ffff00;")
			Err = Err & "株式は半角１文字、全角２文字と数えて４０文字までです。<br>"
		End If
		'全社員数
		If AllEmployeeNum <> "" And IsRE(AllEmployeeNum, "^\d*$", False) = False Then
			Call DicAdd(ErrStyle, "AllEmployeeNum", "background-color:#ffff00;")
			Err = Err & "全社員数は半角数字で１２桁までです。<br>"
		End If
		'男性数
		If ManEmployeeNum <> "" And IsRE(ManEmployeeNum, "^\d*$", False) = False Then
			Call DicAdd(ErrStyle, "ManEmployeeNum", "background-color:#ffff00;")
			Err = Err & "男性社員数は半角数字で１２桁までです。<br>"
		End If
		'女性数
		If WomanEmployeeNum <> "" And IsRE(WomanEmployeeNum, "^\d*$", False) = False Then
			Call DicAdd(ErrStyle, "WomanEmployeeNum", "background-color:#ffff00;")
			Err = Err & "女性社員数は半角数字で１２桁までです。<br>"
		End If
		'ホームページ
		If WomanEmployeeNum <> "" And IsRE(WomanEmployeeNum, "^\d*$", False) = False Then
			Call DicAdd(ErrStyle, "WomanEmployeeNum", "background-color:#ffff00;")
			Err = Err & "ホームページは半角１文字、全角２文字と数えて１００文字までです。<br>"
		End If
		'郵便番号
		If IsRE(Post_U & Post_L, "^\d\d\d\d\d\d\d$", False) = False Then
			Call DicAdd(ErrStyle, "Post_U", "background-color:#ffff00;")
			Call DicAdd(ErrStyle, "Post_L", "background-color:#ffff00;")
			Err = Err & "正しい郵便番号を半角数字で入力して下さい。<br>"
		End If
		'都道府県
		If IsRE(PrefectureCode, "^\d\d\d$", False) = False Then
			Call DicAdd(ErrStyle, "PrefectureCode", "background-color:#ffff00;")
			Err = Err & "都道府県を選択して下さい。<br>"
		End If
		'市区郡
		If City_K <> "" And ChkLen(City_K, 80) = False Then
			Call DicAdd(ErrStyle, "City_K", "background-color:#ffff00;")
			Err = Err & "市区郡は半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'市区郡カナ
		If City_F <> "" And ChkLen(City_F, 80) = False Then
			Call DicAdd(ErrStyle, "City_F", "background-color:#ffff00;")
			Err = Err & "市区郡カナは半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'町村
		If Town <> "" And ChkLen(Town, 80) = False Then
			Call DicAdd(ErrStyle, "Town", "background-color:#ffff00;")
			Err = Err & "町村は半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'番地等
		If Address <> "" And ChkLen(Address, 80) = False Then
			Call DicAdd(ErrStyle, "Address", "background-color:#ffff00;")
			Err = Err & "番地等は半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'移動手段１
		If CompanySyudan1_1 <> "" And ChkLen(CompanySyudan1_1, 20) = False Then
			Call DicAdd(ErrStyle, "CompanySyudan1_1", "background-color:#ffff00;")
			Err = Err & "移動手段１は半角１文字、全角２文字と数えて２０文字までです。<br>"
		End If
		'移動手段１の時間
		If WorkOrBus1 <> "" And ChkLen(WorkOrBus1, 3) = False Then
			Call DicAdd(ErrStyle, "WorkOrBus1", "background-color:#ffff00;")
			Err = Err & "移動手段１の時間は半角数字で３桁までです。<br>"
		End If
		'移動手段２
		If CompanySyudan1_2 <> "" And ChkLen(CompanySyudan1_2, 20) = False Then
			Call DicAdd(ErrStyle, "CompanySyudan1_2", "background-color:#ffff00;")
			Err = Err & "移動手段２は半角１文字、全角２文字と数えて２０文字までです。<br>"
		End If
		'移動手段２の時間
		If WorkBusTime1 <> "" And ChkLen(WorkBusTime1, 3) = False Then
			Call DicAdd(ErrStyle, "WorkBusTime1", "background-color:#ffff00;")
			Err = Err & "移動手段２の時間は半角数字で３桁までです。<br>"
		End If
		'移動手段１
		If CompanySyudan2_1 <> "" And ChkLen(CompanySyudan2_1, 20) = False Then
			Call DicAdd(ErrStyle, "CompanySyudan2_1", "background-color:#ffff00;")
			Err = Err & "移動手段１は半角１文字、全角２文字と数えて２０文字までです。<br>"
		End If
		'移動手段１の時間
		If WorkOrBus2 <> "" And ChkLen(WorkOrBus2, 3) = False Then
			Call DicAdd(ErrStyle, "WorkOrBus2", "background-color:#ffff00;")
			Err = Err & "移動手段１の時間は半角数字で３桁までです。<br>"
		End If
		'移動手段２
		If CompanySyudan2_2 <> "" And ChkLen(CompanySyudan2_2, 20) = False Then
			Call DicAdd(ErrStyle, "CompanySyudan2_2", "background-color:#ffff00;")
			Err = Err & "移動手段２は半角１文字、全角２文字と数えて２０文字までです。<br>"
		End If
		'移動手段２の時間
		If WorkBusTime2 <> "" And ChkLen(WorkBusTime2, 3) = False Then
			Call DicAdd(ErrStyle, "WorkBusTime2", "background-color:#ffff00;")
			Err = Err & "移動手段２の時間は半角数字で３桁までです。<br>"
		End If
		'福利厚生備考
		If WelfareProgramRemark <> "" And ChkLen(WelfareProgramRemark, 100) = False Then
			Call DicAdd(ErrStyle, "WelfareProgramRemark", "background-color:#ffff00;")
			Err = Err & "福利厚生備考は半角１文字、全角２文字と数えて１００文字までです。<br>"
		End If
		'事業内容
		If BusinessContents <> "" And ChkLen(BusinessContents, 1000) = False Then
			Call DicAdd(ErrStyle, "BusinessContents", "background-color:#ffff00;")
			Err = Err & "事業内容は半角１文字、全角２文字と数えて１０００文字までです。<br>"
		End If
		'会社案内
		If CompanyPR <> "" And ChkLen(CompanyPR, 1000) = False Then
			Call DicAdd(ErrStyle, "CompanyPR", "background-color:#ffff00;")
			Err = Err & "会社案内は半角１文字、全角２文字と数えて１０００文字までです。<br>"
		End If
		'窓口担当者名１
		If ContactPersonName <> "" And ChkLen(ContactPersonName, 40) = False Then
			Call DicAdd(ErrStyle, "ContactPersonName", "background-color:#ffff00;")
			Err = Err & "窓口担当者名は半角１文字、全角２文字と数えて４０文字までです。<br>"
		End If
		'窓口担当者役職１
		If Tanto1Yakusyoku <> "" And ChkLen(Tanto1Yakusyoku, 40) = False Then
			Call DicAdd(ErrStyle, "Tanto1Yakusyoku", "background-color:#ffff00;")
			Err = Err & "窓口担当者役職は半角１文字、全角２文字と数えて４０文字までです。<br>"
		End If
		'窓口担当者名２
		If Tanto2Name <> "" And ChkLen(Tanto2Name, 40) = False Then
			Call DicAdd(ErrStyle, "Tanto2Name", "background-color:#ffff00;")
			Err = Err & "窓口担当者名は半角１文字、全角２文字と数えて４０文字までです。<br>"
		End If
		'窓口担当者役職２
		If Tanto2Yakusyoku <> "" And ChkLen(Tanto2Yakusyoku, 40) = False Then
			Call DicAdd(ErrStyle, "Tanto2Yakusyoku", "background-color:#ffff00;")
			Err = Err & "窓口担当者役職は半角１文字、全角２文字と数えて４０文字までです。<br>"
		End If
		'電話番号
		If IsTel(TelephoneNumber, "0") = False Then
			Call DicAdd(ErrStyle, "TelephoneNumber", "background-color:#ffff00;")
			Err = Err & "正しい電話番号を半角数字とハイフン - で入力して下さい。<br>"
		End If
		'メールアドレス
		If MailAddr <> "" And IsMailAddress(MailAddr) = False Then
			Call DicAdd(ErrStyle, "MailAddr", "background-color:#ffff00;")
			Err = Err & "正しいメールアドレスを半角、５０文字以内で入力して下さい。<br>"
		End If
		'請求書送付先市区郡
		If DemandCity_K <> "" And ChkLen(DemandCity_K, 80) = False Then
			Call DicAdd(ErrStyle, "DemandCity_K", "background-color:#ffff00;")
			Err = Err & "請求書送付先市区郡は半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'請求書送付先市区郡カナ
		If DemandCity_F <> "" And ChkLen(DemandCity_F, 80) = False Then
			Call DicAdd(ErrStyle, "DemandCity_F", "background-color:#ffff00;")
			Err = Err & "請求書送付先市区郡カナは半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'請求書送付先町村
		If DemandTown <> "" And ChkLen(DemandTown, 80) = False Then
			Call DicAdd(ErrStyle, "DemandTown", "background-color:#ffff00;")
			Err = Err & "請求書送付先町村は半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'請求書送付先番地等
		If DemandAddress <> "" And ChkLen(DemandAddress, 80) = False Then
			Call DicAdd(ErrStyle, "DemandAddress", "background-color:#ffff00;")
			Err = Err & "請求書送付先市区郡カナは半角１文字、全角２文字と数えて８０文字までです。<br>"
		End If
		'請求書送付先部署
		If DemandSectionName <> "" And ChkLen(DemandSectionName, 40) = False Then
			Call DicAdd(ErrStyle, "DemandSectionName", "background-color:#ffff00;")
			Err = Err & "請求書送付先部署は半角１文字、全角２文字と数えて４０文字までです。<br>"
		End If
		'請求書送付先担当者
		If DemandPersonName <> "" And ChkLen(DemandPersonName, 40) = False Then
			Call DicAdd(ErrStyle, "DemandPersonName", "background-color:#ffff00;")
			Err = Err & "請求書送付先担当者名は半角１文字、全角２文字と数えて４０文字までです。<br>"
		End If
		'会社の雰囲気
		If Atmosphere <> "" And ChkLen(Atmosphere, 500) = False Then
			Call DicAdd(ErrStyle, "DemandPersonName", "background-color:#ffff00;")
			Err = Err & "会社の雰囲気は半角１文字、全角２文字と数えて５００文字までです。<br>"
		End If

		If Err = "" Then IsData = True
	End Function

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_CompanyInfo 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(vCompanyCode)
		GetRegSQL = ""
		GetRegSQL = GetRegSQL & "EXEC up_RegCompanyInfo_Navi"
		GetRegSQL = GetRegSQL & " '" & vCompanyCode & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyKbn & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyName_K & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyName_F & "'"
		GetRegSQL = GetRegSQL & ",'" & EstablishYear & "'"
		GetRegSQL = GetRegSQL & ",'" & IndustryType & "'"
		GetRegSQL = GetRegSQL & ",'" & CapitalAmount & "'"
		GetRegSQL = GetRegSQL & ",'" & ForeinCapital & "'"
		GetRegSQL = GetRegSQL & ",'" & ListClass & "'"
		GetRegSQL = GetRegSQL & ",'" & AllEmployeeNum & "'"
		GetRegSQL = GetRegSQL & ",'" & ManEmployeeNum & "'"
		GetRegSQL = GetRegSQL & ",'" & WomanEmployeeNum & "'"
		GetRegSQL = GetRegSQL & ",'" & HomepageAddress & "'"
		GetRegSQL = GetRegSQL & ",'" & Post_U & "'"
		GetRegSQL = GetRegSQL & ",'" & Post_L & "'"
		GetRegSQL = GetRegSQL & ",'" & PrefectureCode & "'"
		GetRegSQL = GetRegSQL & ",'" & City_K & "'"
		GetRegSQL = GetRegSQL & ",'" & City_F & "'"
		GetRegSQL = GetRegSQL & ",'" & Town & "'"
		GetRegSQL = GetRegSQL & ",'" & Address & "'"
		GetRegSQL = GetRegSQL & ",'" & TelephoneNumber & "'"
		GetRegSQL = GetRegSQL & ",'" & StationCode1 & "'"
		GetRegSQL = GetRegSQL & ",'" & StationName1 & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanySyudan1_1 & "'"
		GetRegSQL = GetRegSQL & ",'" & WorkOrBus1 & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanySyudan1_2 & "'"
		GetRegSQL = GetRegSQL & ",'" & WorkBusTime1 & "'"
		GetRegSQL = GetRegSQL & ",'" & StationCode2 & "'"
		GetRegSQL = GetRegSQL & ",'" & StationName2 & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanySyudan2_1 & "'"
		GetRegSQL = GetRegSQL & ",'" & WorkOrBus2 & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanySyudan2_2 & "'"
		GetRegSQL = GetRegSQL & ",'" & WorkBusTime2 & "'"
		GetRegSQL = GetRegSQL & ",'" & SocietyInsurance & "'"
		GetRegSQL = GetRegSQL & ",'" & Sanatorium & "'"
		GetRegSQL = GetRegSQL & ",'" & EnterprisePension & "'"
		GetRegSQL = GetRegSQL & ",'" & WealthShape & "'"
		GetRegSQL = GetRegSQL & ",'" & StockOption & "'"
		GetRegSQL = GetRegSQL & ",'" & RetirementPay & "'"
		GetRegSQL = GetRegSQL & ",'" & ResidencePay & "'"
		GetRegSQL = GetRegSQL & ",'" & FamilyPay & "'"
		GetRegSQL = GetRegSQL & ",'" & EmployeeDormitory & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyHouse & "'"
		GetRegSQL = GetRegSQL & ",'" & NewEmployeeTraining & "'"
		GetRegSQL = GetRegSQL & ",'" & OverseasTraining & "'"
		GetRegSQL = GetRegSQL & ",'" & OtherTraning & "'"
		'GetRegSQL = GetRegSQL & ",'" & FlexTime & "'" '08/08/14 Lis林 DEL
		GetRegSQL = GetRegSQL & ",'" & WelfareProgramRemark & "'"
		GetRegSQL = GetRegSQL & ",'" & BusinessContents & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyPR & "'"
		GetRegSQL = GetRegSQL & ",'" & Simebi & "'"
		GetRegSQL = GetRegSQL & ",'" & ContactPersonName & "'"
		GetRegSQL = GetRegSQL & ",'" & Tanto1Yakusyoku & "'"
		GetRegSQL = GetRegSQL & ",'" & Tanto2Name & "'"
		GetRegSQL = GetRegSQL & ",'" & Tanto2Yakusyoku & "'"
		GetRegSQL = GetRegSQL & ",'" & MailAddr & "'"
		GetRegSQL = GetRegSQL & ",'" & NewJobMail & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandPrefectureCode & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandCity_K & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandCity_F & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandTown & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandAddress & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandSectionName & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandPersonName & "'"
		GetRegSQL = GetRegSQL & ",'" & Atmosphere & "'"
	End Function
End Class
%>
