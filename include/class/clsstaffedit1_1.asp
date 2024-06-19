<%
'******************************************************************************
'概　要：基本情報（/staff/edit1_1.asp）登録用のクラス
'備　考：■メンバ関数
'　　　：SetData
'　　　：ChkData
'　　　：GetRegSQL
'　　　：DiffData
'更　新：2008/04/22 LIS K.Kokubo
'******************************************************************************
Class clsStaffEdit1_1
	'登録用データ
	Public Name_1					'姓
	Public Name_2					'名
	Public Name_F_1					'セイ
	Public Name_F_2					'メイ
	Public Birthday					'誕生日
	Public Sex						'性別
	Public Post_U					'住所：郵便番号上３桁
	Public Post_L					'住所：郵便番号下４桁
	Public PrefectureCode			'住所：都道府県コード
	Public PrefectureName			'住所：都道府県名
	Public City						'住所：市区郡
	Public City_F					'住所：市区郡カナ
	Public Town						'住所：町村
	Public Town_F					'住所：町村カナ
	Public Address					'住所：番地など
	Public Address_F				'住所：番地などカナ
	Public HomeTelephoneNumber		'家TEL
	Public PortableTelephoneNumber	'携帯
	Public FaxNumber				'FAX
	Public MailAddress				'ＰＣメールアドレス
	Public MailAddress2				'ＰＣメールアドレス確認
	Public PortableMailAddress		'携帯メールアドレス
	Public PortableMailAddress2		'携帯メールアドレス確認
	Public HomeContactFlag			'希望連絡先フラグ：家TEL
	Public PortableContactFlag		'希望連絡先フラグ：携帯
	Public FaxContactFlag			'希望連絡先フラグ：FAX
	Public MailContactFlag			'希望連絡先フラグ：メール
	Public NoticeMailFlag			'メール連絡先フラグ
	Public UrgencyPost_U			'緊急連絡先：郵便番号上３桁
	Public UrgencyPost_L			'緊急連絡先：郵便番号下４桁
	Public UrgencyAddress			'緊急連絡先：住所
	Public UrgencyAddress_F			'緊急連絡先：住所カナ
	Public UrgencyTelephoneNumber	'緊急連絡先：TEL
	Public URL						'ホームページ
	'エラー処理用
	Public Err							'エラー文言
	Public ErrName_1					'姓
	Public ErrName_2					'名
	Public ErrName_F_1					'セイ
	Public ErrName_F_2					'メイ
	Public ErrBirthday					'誕生日
	Public ErrSex						'性別
	Public ErrPost_U					'住所：郵便番号上３桁
	Public ErrPost_L					'住所：郵便番号下４桁
	Public ErrPrefectureCode			'住所：都道府県コード
	Public ErrCity						'住所：市区郡
	Public ErrCity_F					'住所：市区郡カナ
	Public ErrTown						'住所：町村
	Public ErrTown_F					'住所：町村カナ
	Public ErrAddress					'住所：番地など
	Public ErrAddress_F					'住所：番地などカナ
	Public ErrHomeTelephoneNumber		'家TEL
	Public ErrPortableTelephoneNumber	'携帯
	Public ErrFaxNumber					'FAX
	Public ErrMailAddress				'ＰＣメールアドレス
	Public ErrMailAddress2				'ＰＣメールアドレス確認
	Public ErrPortableMailAddress		'携帯メールアドレス
	Public ErrPortableMailAddress2		'携帯メールアドレス確認
	Public ErrHomeContactFlag			'希望連絡先フラグ：家TEL
	Public ErrPortableContactFlag		'希望連絡先フラグ：携帯
	Public ErrFaxContactFlag			'希望連絡先フラグ：FAX
	Public ErrMailContactFlag			'希望連絡先フラグ：メール
	Public ErrNoticeMailFlag			'メール連絡先フラグ
	Public ErrUrgencyPost_U				'緊急連絡先：郵便番号上３桁
	Public ErrUrgencyPost_L				'緊急連絡先：郵便番号下４桁
	Public ErrUrgencyAddress			'緊急連絡先：住所
	Public ErrUrgencyAddress_F			'緊急連絡先：住所カナ
	Public ErrUrgencyTelephoneNumber	'緊急連絡先：TEL
	Public ErrURL						'ホームページ

	'******************************************************************************
	'概　要：入力データチェック
	'引　数：
	'備　考：
	'更　新：2008/04/22 LIS K.Kokubo
	'******************************************************************************
	Public Function SetData(ByVal vStaffCode)
		Dim sSQL
		Dim oRS
		Dim sError
		Dim flgQE

		Dim dbName
		Dim dbName_F

		sSQL = "sp_GetDetailStaff '" & vStaffCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbName = oRS.Collect("Name")
			If InStr(dbName, "　") <> 0 Then
				Name_1 = Mid(dbName, 1, InStr(dbName, "　") - 1)
				Name_2 = Mid(dbName, InStr(dbName, "　") + 1)
			Else
				Name_1 = dbName
			End If

			dbName_F = oRS.Collect("Name_F")
			If InStr(dbName_F, "　") <> 0 Then
				Name_F_1 = Mid(dbName_F, 1, InStr(dbName_F, "　") - 1)
				Name_F_2 = Mid(dbName_F, InStr(dbName_F, "　") + 1)
			Else
				Name_F_1 = dbName_F
			End If

			Post_U = oRS.Collect("Post_U")
			Post_L = oRS.Collect("Post_L")
			Birthday = GetDateStr(oRS.Collect("Birthday"), "")
			Sex = oRS.Collect("SexType")
			PrefectureCode = oRS.Collect("PrefectureCode")
			PrefectureName = oRS.Collect("PrefectureName")
			City = oRS.Collect("City")
			City_F = oRS.Collect("City_F")
			Town = oRS.Collect("Town")
			Town_F = oRS.Collect("Town_F")
			Address = oRS.Collect("Address")
			Address_F = oRS.Collect("Address_F")
			HomeTelephoneNumber = oRS.Collect("HomeTelephoneNumber")
			PortableTelephoneNumber = oRS.Collect("PortableTelephoneNumber")
			FaxNumber = oRS.Collect("FaxNumber")
			MailAddress = oRS.Collect("MailAddress")
			MailAddress2 = oRS.Collect("MailAddress")
			PortableMailAddress = oRS.Collect("PortableMailAddress")
			PortableMailAddress2 = oRS.Collect("PortableMailAddress")
			HomeContactFlag = oRS.Collect("HomeContactFlag")
			PortableContactFlag = oRS.Collect("PortableContactFlag")
			FaxContactFlag = oRS.Collect("FaxContactFlag")
			MailContactFlag = oRS.Collect("MailContactFlag")
			NoticeMailFlag = ChkStr(oRS.Collect("NoticeMailFlag"))
			UrgencyPost_U = oRS.Collect("UrgencyPost_U")
			UrgencyPost_L = oRS.Collect("UrgencyPost_L")
			UrgencyAddress = oRS.Collect("UrgencyAddress")
			UrgencyAddress_F = oRS.Collect("UrgencyAddress_F")
			UrgencyTelephoneNumber = oRS.Collect("UrgencyTelephoneNumber")
			URL = oRS.Collect("URL")
		End If
		Call RSClose(oRS)
	End Function

	'******************************************************************************
	'概　要：入力データチェック
	'引　数：
	'備　考：
	'更　新：2008/04/22 LIS K.Kokubo
	'******************************************************************************
	Public Function ChkData()
		Dim sStyle
		Dim flgReg

		sStyle = "background-color:#ffffcc;"
		flgReg = True
		Err = ""

		'姓チェック
		If Name_1 = "" Then
			Err = Err & "「姓」は必須です。全角で入力してください。<br>"
			ErrName_1 = sStyle
			flgReg = False
		ElseIf IsRE(Name_1, "[\w\f\n\r\t\v,.*@!""#$%&'()-=~|\[\]\\?+;:{}]", False) = True Or ChkLen(Name_1, 40) = False Then
			Err = Err & "「姓」は全角で２０文字以内で入力してください。<br>"
			ErrName_1 = sStyle
			flgReg = False
		End If

		'名チェック
		If Name_2 = "" Then
			Err = Err & "「名」は必須です。全角で入力してください。<br>"
			ErrName_2 = sStyle
			flgReg = False
		ElseIf IsRE(Name_2, "[\w\f\n\r\t\v,.*@!""#$%&'()-=~|\[\]\\?+;:{}]", False) = True Or ChkLen(Name_2, 40) = False Then
			Err = Err & "「名」は全角で２０文字以内で入力してください。<br>"
			ErrName_2 = sStyle
			flgReg = False
		End If

		'セイチェック
		If Name_F_1 = "" Then
			Err = Err & "「セイ」は必須です。全角カナで入力してください。<br>"
			ErrName_F_1 = sStyle
			flgReg = False
		ElseIf ChkKana(Name_F_1) = False Or ChkLen(Name_F_1, 40) = False Then
			Err = Err & "「セイ」は全角カナで入力してください。<br>"
			ErrName_F_1 = sStyle
			flgReg = False
		End If

		'メイチェック
		If Name_F_2 = "" Then
			Err = Err & "「メイ」は必須です。全角カナで入力してください。<br>"
			ErrName_F_2 = sStyle
			flgReg = False
		ElseIf ChkKana(Name_F_2) = False Or ChkLen(Name_F_2, 40) = False Then
			Err = Err & "「メイ」は全角カナで入力してください。<br>"
			ErrName_F_2 = sStyle
			flgReg = False
		End If

		'性別チェック
		If Not(Sex = "1" Or Sex = "2") Then
			Err = Err & "性別をチェックしてください。<br>"
			ErrSex = sStyle
			flgReg = False
		End If

		'郵便番号チェック
		If Post_U & Post_L = "" Then
			Err = Err & "現住所の郵便番号を入力してください。<br>"
			ErrPost_U = sStyle
			ErrPost_L = sStyle
			flgReg = False
		ElseIf IsRE(Post_U & Post_L, "^\d\d\d\d\d\d\d$", False) = True Then 
			sSQL = "/* 求職者基本情報編集時の郵便番号チェック */ "
			sSQL = "EXEC up_DtlZip '" & Post_U & "', '" & Post_L & "'"

			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = False Then
				Err = Err & "現住所の郵便番号は存在しません。<br>"
				ErrPost_U = sStyle
				ErrPost_L = sStyle
				flgReg = False
			End If
			Call RSClose(oRS)
		Else
			Err = Err & "現住所の郵便番号は半角数字で入力してください。<br>"
			ErrPost_U = sStyle
			ErrPost_L = sStyle
			flgReg = False
		End If

		'都道府県コード
		If IsRE(PrefectureCode, "^\d\d\d$", False) = False Then
			Err = Err & "都道府県を選択してください。<br>"
			ErrPrefectureCode = sStyle
			flgReg = False
		End If

		'市区郡
		If City = "" Then
			Err = Err & "市区郡は必須です。全角５０文字以内で入力してください。<br>"
			ErrCity = sStyle
			flgReg = False
		ElseIf ChkLen(City, 100) = False Then
			Err = Err & "市区郡の文字数が制限数を超えています。全角５０文字以内で入力してください。<br>"
			ErrCity = sStyle
			flgReg = False
		End If

		'市区郡カナ
		If City_F <> "" Then
			If ChkLen(City_F, 100) = False Then
				Err = Err & "市区郡カナの文字数が制限数を超えています。全角５０文字以内で入力してください。<br>"
				ErrCity_F = sStyle
				flgReg = False
			End If
		End If

		'町村
		If Town <> "" Then
			If ChkLen(Town, 100) = False Then
				Err = Err & "町村の文字数が制限数を超えています。全角５０文字以内で入力してください。<br>"
				ErrTown = sStyle
				flgReg = False
			End If
		End If

		'町村カナ
		If Town_F <> "" Then
			If ChkLen(Town_F, 100) = False Then
				Err = Err & "町村カナの文字数が制限数を超えています。全角５０文字以内で入力してください。<br>"
				ErrTown_F = sStyle
				flgReg = False
			End If
		End If

		'番地等
		If Address <> "" Then
			If ChkLen(Address, 100) = False Then
				Err = Err & "番地等の文字数が制限数を超えています。全角５０文字以内で入力してください。<br>"
				ErrAddress = sStyle
				flgReg = False
			End If
		End If

		'番地等カナ
		If Address_F <> "" Then
			If ChkLen(Address_F, 100) = False Then
				Err = Err & "番地等カナの文字数が制限数を超えています。全角５０文字以内で入力してください。<br>"
				ErrAddress_F = sStyle
				flgReg = False
			End If
		End If

		'家TEL
		If HomeTelephoneNumber <> "" Then
			If IsTel(HomeTelephoneNumber, "1") = False Then
				Err = Err & "ＰＣメールアドレスに誤りがあります。<br>"
				ErrHomeTelephoneNumber = sStyle
				flgReg = False
			End If
		End If

		'ＰＣメールアドレス
		If MailAddress = "" Then
			Err = Err & "ＰＣメールアドレスは必須です。<br>"
			ErrMailAddress = sStyle
			flgReg = False
		ElseIf IsMailAddress(MailAddress) = False Then
			Err = Err & "ＰＣメールアドレスに誤りがあります。<br>"
			ErrMailAddress = sStyle
			flgReg = False
		End If
		'ＰＣメールアドレス確認
		If MailAddress <> MailAddress2 Then
			Err = Err & "ＰＣメールアドレスが確認のものとで相違しています。<br>"
			ErrMailAddress2 = sStyle
			flgReg = False
		End If

		'携帯メールアドレス
		If PortableMailAddress <> "" Then
			If IsMailAddress(PortableMailAddress) = False Then
				Err = Err & "携帯メールアドレスに誤りがあります。<br>"
				ErrPortableMailAddress = sStyle
				flgReg = False
			End If
		End If
		'携帯メールアドレス確認
		If PortableMailAddress <> PortableMailAddress2 Then
			Err = Err & "携帯メールアドレスが確認のものとで相違しています。<br>"
				ErrPortableMailAddress2 = sStyle
			flgReg = False
		End If

		'希望連絡方法チェック
		If HomeContactFlag <> "" And HomeContactFlag <> "1" Then HomeContactFlag = ""
		If PortableContactFlag <> "" And PortableContactFlag <> "1" Then PortableContactFlag = ""
		If FaxContactFlag <> "" And FaxContactFlag <> "1" Then FaxContactFlag = ""
		If MailContactFlag <> "" And MailContactFlag <> "1" Then MailContactFlag = ""

		'緊急連絡先郵便番号チェック
		If UrgencyPost_U & UrgencyPost_L <> "" Then
			If IsRE(UrgencyPost_U & UrgencyPost_L, "^\d\d\d\d\d\d\d$", False) = True Then 
				sSQL = "/* 求職者基本情報編集時の郵便番号チェック */ "
				sSQL = "EXEC up_DtlZip '" & UrgencyPost_U & "', '" & UrgencyPost_L & "'"

				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = False Then
					Err = Err & "緊急連絡先の郵便番号は存在しません。<br>"
					ErrUrgencyPost_U = sStyle
					ErrUrgencyPost_L = sStyle
					flgReg = False
				End If
				Call RSClose(oRS)
			End If
		End If
		'緊急連絡先住所チェック
		If UrgencyAddress <> "" Then
			If ChkLen(UrgencyAddress, 200) = False Then
				Err = Err & "緊急連絡先の住所は全角２００文字以内で入力してください。<br>"
				ErrUrgencyAddress = sStyle
				flgReg = False
			End If
		End If
		'緊急連絡先住所カナチェック
		If UrgencyAddress_F <> "" Then
			If ChkLen(UrgencyAddress_F, 200) = False Then
				Err = Err & "緊急連絡先の住所カナは全角２００文字以内で入力してください。<br>"
				ErrUrgencyAddress_F = sStyle
				flgReg = False
			End If
		End If
		'緊急連絡先TELチェック
		If UrgencyTelephoneNumber <> "" Then
			If IsTel(UrgencyTelephoneNumber, "0") = False Then
				Err = Err & "緊急連絡先の電話番号が不正です。正しい電話番号を入力してください。<br>"
				ErrUrgencyTelephoneNumber = sStyle
				flgReg = False
			End If
		End If

		'ホームページチェック
		If URL <> "" Then
			If ChkLen(URL, 200) = False Then
				Err = Err & "ホームページは半角で１００文字以内で入力してください。<br>"
				ErrURL = sStyle
				flgReg = False
			End If
		End If

		ChkData = flgReg
	End Function

	'******************************************************************************
	'概　要：登録ＳＱＬ取得
	'引　数：vStaffCode	：求職者コード
	'備　考：
	'更　新：2008/04/22 LIS K.Kokubo
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim sSQL

		sSQL = "EXEC up_RegStaff_Edit1_1"
		sSQL = sSQL & " '" & vStaffCode & "'"
		sSQL = sSQL & ",'" & Name_1 & "　" & Name_2 & "'"
		sSQL = sSQL & ",'" & Name_F_1 & "　" & Name_F_2 & "'"
		sSQL = sSQL & ",'" & Name_1 & Name_2 & "'"
		sSQL = sSQL & ",'" & Name_F_1 & Name_F_2 & "'"
		sSQL = sSQL & ",'" & Post_U & "'"
		sSQL = sSQL & ",'" & Post_L & "'"
		sSQL = sSQL & ",'" & PrefectureCode & "'"
		sSQL = sSQL & ",'" & City & "'"
		sSQL = sSQL & ",'" & City_F & "'"
		sSQL = sSQL & ",'" & Town & "'"
		sSQL = sSQL & ",'" & Town_F & "'"
		sSQL = sSQL & ",'" & Address & "'"
		sSQL = sSQL & ",'" & Address_F & "'"
		sSQL = sSQL & ",'" & HomeTelephoneNumber & "'"
		sSQL = sSQL & ",'" & PortableTelephoneNumber & "'"
		sSQL = sSQL & ",'" & FaxNumber & "'"
		sSQL = sSQL & ",'" & MailAddress & "'"
		sSQL = sSQL & ",'" & PortableMailAddress & "'"
		sSQL = sSQL & ",'" & NoticeMailFlag & "'"
		sSQL = sSQL & ",'" & UrgencyPost_U & "'"
		sSQL = sSQL & ",'" & UrgencyPost_L & "'"
		sSQL = sSQL & ",'" & UrgencyAddress & "'"
		sSQL = sSQL & ",'" & UrgencyAddress_F & "'"
		sSQL = sSQL & ",'" & UrgencyTelephoneNumber & "'"
		sSQL = sSQL & ",'" & HomeContactFlag & "'"
		sSQL = sSQL & ",'" & PortableContactFlag & "'"
		sSQL = sSQL & ",'" & FaxContactFlag & "'"
		sSQL = sSQL & ",'" & MailContactFlag & "'"
		sSQL = sSQL & ",'" & Birthday & "'"
		sSQL = sSQL & ",'" & Sex & "'"
		sSQL = sSQL & ",'" & URL & "'"

		GetRegSQL = sSQL
	End Function

	'******************************************************************************
	'概　要：更新前後のデータの差異をメール(リスたま：登録情報更新通知機能)
	'引　数：vStaffCode	：求職者コード
	'備　考：基本情報,現住所,連絡先が変わった場合のみ
	'更　新：2008/04/22 LIS K.Kokubo
	'******************************************************************************
	Public Function DiffData(ByRef rSE)
		Dim sChg
		Dim sBef
		Dim sAft

		sChg = ""

		'基本情報
		sBef = ""
		sAft = ""
		If Name_1 & Name_2 & Name_F_1 & Name_F_2 & Birthday & Sex _
		<> rSE.Name_1 & rSE.Name_2 & rSE.Name_F_1 & rSE.Name_F_2 & rSE.Birthday & rSE.Sex Then
			If Name_1 & Name_2 <> rSE.Name_1 & rSE.Name_2 Then
				sBef = sBef & "[名　　前]" & rSE.Name_1 & rSE.Name_2 & vbCrLf
				sAft = sAft & "[名　　前]" & Name_1 & Name_2 & vbCrLf
			End If

			If Name_F_1 & Name_F_2 <> rSE.Name_F_1 & rSE.Name_F_2 Then
				sBef = sBef & "[名前カナ]" & rSE.Name_F_1 & rSE.Name_F_2 & vbCrLf
				sAft = sAft & "[名前カナ]" & Name_F_1 & Name_F_2 & vbCrLf
			End If

			If Birthday <> rSE.Birthday Then
				sBef = sBef & "[生年月日]" & rSE.Birthday & vbCrLf
				sAft = sAft & "[生年月日]" & Birthday & vbCrLf
			End If

			If Sex <> rSE.Sex Then
				sBef = sBef & "[性　　別]" & rSE.Sex & vbCrLf
				sAft = sAft & "[性　　別]" & Sex & vbCrLf
			End If

			sChg = sChg & "---------- 基本情報 ----------" & vbCrLf
			sChg = sChg & sBef
			sChg = sChg & "↓" & vbCrLf
			sChg = sChg & sAft
		End If

		'現住所
		sBef = ""
		sAft = ""
		If Post_U & Post_L & PrefectureCode & City & City_F & Town & Town_F & Address & Address_F _
		<> rSE.Post_U & rSE.Post_L & rSE.PrefectureCode & rSE.City & rSE.City_F & rSE.Town & rSE.Town_F & rSE.Address & rSE.Address_F Then
			If Post_U & Post_L <> rSE.Post_U & rSE.Post_L Then
				sBef = sBef & "[郵便番号]" & rSE.Post_U & "-" & rSE.Post_L & vbCrLf
				sAft = sAft & "[郵便番号]" & Post_U & "-" & Post_L & vbCrLf
			End If

			If PrefectureName & City & Town & Address <> rSE.PrefectureCode & rSE.City & rSE.Town & rSE.Address Then
				sBef = sBef & "[住　　所]" & rSE.PrefectureName & rSE.City & rSE.Town & rSE.Address & vbCrLf
				sAft = sAft & "[住　　所]" & PrefectureName & City & Town & Address & vbCrLf
			End If

			If City_F & Town_F & Address_F <> rSE.City_F & rSE.Town_F & rSE.Address_F Then
				sBef = sBef & "[住所カナ]" & rSE.City_F & rSE.Town_F & rSE.Address_F & vbCrLf
				sAft = sAft & "[住所カナ]" & City_F & Town_F & Address_F & vbCrLf
			End If

			sChg = sChg & "---------- 現住所 ----------" & vbCrLf
			sChg = sChg & sBef
			sChg = sChg & "↓" & vbCrLf
			sChg = sChg & sAft
		End If

		'連絡先
		sBef = ""
		sAft = ""
		If HomeTelephoneNumber & PortableTelephoneNumber & FaxNumber & MailAddress & PortableMailAddress _
		<> rSE.HomeTelephoneNumber & rSE.PortableTelephoneNumber & rSE.FaxNumber & rSE.MailAddress & rSE.PortableMailAddress _
		Or HomeContactFlag <> rSE.HomeContactFlag Or PortableContactFlag <> rSE.PortableContactFlag Or FaxContactFlag <> rSE.FaxContactFlag Or MailContactFlag <> rSE.MailContactFlag Then
			If HomeTelephoneNumber <> rSE.HomeTelephoneNumber Then
				sBef = sBef & "[Ｔ Ｅ Ｌ]" & rSE.HomeTelephoneNumber & vbCrLf
				sAft = sAft & "[Ｔ Ｅ Ｌ]" & HomeTelephoneNumber & vbCrLf
			End If

			If PortableTelephoneNumber <> rSE.PortableTelephoneNumber Then
				sBef = sBef & "[携　　帯]" & rSE.PortableTelephoneNumber & vbCrLf
				sAft = sAft & "[携　　帯]" & PortableTelephoneNumber & vbCrLf
			End If

			If FaxNumber <> rSE.FaxNumber Then
				sBef = sBef & "[Ｆ Ａ Ｘ]" & rSE.FaxNumber & vbCrLf
				sAft = sAft & "[Ｆ Ａ Ｘ]" & FaxNumber & vbCrLf
			End If

			If MailAddress <> rSE.MailAddress Then
				sBef = sBef & "[ＰＣMAIL]" & rSE.MailAddress & vbCrLf
				sAft = sAft & "[ＰＣMAIL]" & MailAddress & vbCrLf
			End If

			If PortableMailAddress <> rSE.PortableMailAddress Then
				sBef = sBef & "[携帯MAIL]" & rSE.PortableMailAddress & vbCrLf
				sAft = sAft & "[携帯MAIL]" & PortableMailAddress & vbCrLf
			End If

			If HomeContactFlag <> rSE.HomeContactFlag Or PortableContactFlag <> rSE.PortableContactFlag Or FaxContactFlag <> rSE.FaxContactFlag Or MailContactFlag <> rSE.MailContactFlag Then
				sBef = sBef & "[連絡方法]"
				If rSE.HomeContactFlag = "1" Then sBef = sBef & "家,"
				If rSE.PortableContactFlag = "1" Then sBef = sBef & "携帯,"
				If rSE.FaxContactFlag = "1" Then sBef = sBef & "FAX,"
				If rSE.MailContactFlag = "1" Then sBef = sBef & "メール,"
				sBef = sBef & vbCrLf

				sAft = sAft & "[連絡方法]"
				If HomeContactFlag = "1" Then sAft = sAft & "家,"
				If PortableContactFlag = "1" Then sAft = sAft & "携帯,"
				If FaxContactFlag = "1" Then sAft = sAft & "FAX,"
				If MailContactFlag = "1" Then sAft = sAft & "メール,"
				sAft = sAft & vbCrLf
			End If

			sChg = sChg & "---------- 連絡先 ----------" & vbCrLf
			sChg = sChg & sBef
			sChg = sChg & "↓" & vbCrLf
			sChg = sChg & sAft
			sChg = sChg & vbCrLf
		End If

		DiffData = sChg
	End Function
End Class
%>
