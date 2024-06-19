<%
'******************************************************************************
'名　称：clsP_Info
'概　要：formで飛んできたP_Infoテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_Info
	Public StaffCode
	Public Name
	Public Name_F
	Public SearchName
	Public SearchName_F
	Public OldName
	Public Birthday
	Public Sex
	Public MarriageFlag
	Public Post_U
	Public Post_L
	Public PrefectureCode
	Public City
	Public City_F
	Public Town
	Public Town_F
	Public Address
	Public Address_F
	Public LivingType
	Public HomeTelephoneNumber
	Public CountryTelephoneNumber
	Public PortableTelephoneNumber
	Public FaxNumber
	Public MailAddress
	Public PortableMailAddress
	Public UrgencyPost_U
	Public UrgencyPost_L
	Public UrgencyAddress
	Public UrgencyAddress_F
	Public UrgencyTelephoneNumber
	Public URL
	Public InfoSourceType
	Public InfoSourceDay
	Public InfoSourceOther
	Public DependentFlag
	Public DependentNumber
	Public SpouseFlag
	Public CurrentCompanyName
	Public CurrentCompanyName_F
	Public SocietyInsuranceIn
	Public SocietyInsuranceLoss
	Public EmployInsuranceIn
	Public EmployInsuranceLoss
	Public IsData
	Public MaxIndex
	Public Err
	Public ErrStyle

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_Info クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1
	End Sub

	Public Function ChkData()
		'値チェック
		Err = ""
		Set ErrStyle = Server.CreateObject("Scripting.Dictionary")
		ErrStyle.CompareMode = 1

		'名前
		If Name = "" Or ChkLen(Name, 100) = False Then
			ErrStyle("Name_1") = "background-color:#ffff00;"
			ErrStyle("Name_2") = "background-color:#ffff00;"
			Err = Err & "名前が不正です。<br>"
		End If

		'名前カナ
		If Name_F = "" Or ChkLen(Name_F, 100) = False Then
			ErrStyle("Name_F_1") = "background-color:#ffff00;"
			ErrStyle("Name_F_2") = "background-color:#ffff00;"
			Err = Err & "名前カナが不正です。<br>"
		End If

		'誕生日
		If Birthday <> "" And IsDay(Birthday) = False Then
			ErrStyle("Birthday") = "background-color:#ffff00;"
			Err = Err & "誕生日の日付が不正です。<br>"
		End If

		'性別
		If Sex <> "" And IsRE(Sex, "^[12]$", True) = False Then
			ErrStyle("Sex") = "background-color:#ffff00;"
			Err = Err & "性別を入力してください。<br>"
		End If

		'
		If MarriageFlag <> "" And IsFlag(MarriageFlag) = False Then
			ErrStyle("MarriageFlag") = "background-color:#ffff00;"
			Err = Err & "扶養が不正です。<br>"
		End If

		'郵便番号
		If Post_U & Post_L <> "" And IsNumber(Post_U & Post_L, 7, False) = False Then
			ErrStyle("Post_U") = "background-color:#ffff00;"
			ErrStyle("Post_L") = "background-color:#ffff00;"
			Err = Err & "郵便番号が不正です。<br>"
		End If

		'都道府県コード
		If PrefectureCode <> "" And IsNumber(PrefectureCode, 3, False) = False Then
			ErrStyle("PrefectureCode") = "background-color:#ffff00;"
			Err = Err & "住所の都道府県が不正です。<br>"
		End If

		'同居人種類
		If LivingType <> "" And IsRE(LivingType, "^[1234]$", True) = False Then
			ErrStyle("PrefectureCode") = "background-color:#ffff00;"
			Err = Err & "同居人種類が不正です。<br>"
		End If

		'電話番号
		If HomeTelephoneNumber <> "" And IsNumber(HomeTelephoneNumber, 0, False) = False Then
			ErrStyle("HomeTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "電話番号が不正です。<br>"
		End If

		'実家電話番号
		If CountryTelephoneNumber <> "" And IsNumber(CountryTelephoneNumber, 0, False) = False Then
			ErrStyle("CountryTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "実家電話番号が不正です。<br>"
		End If

		'携帯番号
		If PortableTelephoneNumber <> "" And IsNumber(PortableTelephoneNumber, 0, False) = False Then
			ErrStyle("PortableTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "携帯番号が不正です。<br>"
		End If

		'FAX
		If FaxNumber <> "" And IsNumber(FaxNumber, 0, False) = False Then
			ErrStyle("PortableTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "FAX番号が不正です。<br>"
		End If

		'緊急連絡先郵便番号
		If UrgencyPost_U & UrgencyPost_L <> "" And IsNumber(UrgencyPost_U & UrgencyPost_L, 7, False) = False Then
			ErrStyle("UrgencyPost_U") = "background-color:#ffff00;"
			ErrStyle("UrgencyPost_L") = "background-color:#ffff00;"
			Err = Err & "緊急連絡先郵便番号が不正です。<br>"
		End If

		'緊急連絡先住所
		If UrgencyAddress <> "" And ChkLen(UrgencyAddress, 200) = False Then
			ErrStyle("UrgencyAddress") = "background-color:#ffff00;"
			Err = Err & "緊急連絡先住所の文字数が多すぎます。全角で１００文字まで入力できます。<br>"
		End If

		'緊急連絡先住所カナ
		If UrgencyAddress_F <> "" And ChkLen(UrgencyAddress_F, 200) = False Then
			ErrStyle("UrgencyAddress_F") = "background-color:#ffff00;"
			Err = Err & "緊急連絡先住所カナの文字数が多すぎます。全角で１００文字まで入力できます。<br>"
		End If

		'緊急連絡先電話番号
		If UrgencyTelephoneNumber <> "" And IsTel(UrgencyTelephoneNumber, 0) = False Then
			ErrStyle("UrgencyTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "緊急連絡先電話番号が不正です。<br>"
		End If

		'情報源種類
		If InfoSourceType <> "" And IsNumber(InfoSourceType, 3, False) = False Then
			ErrStyle("InfoSourceType") = "background-color:#ffff00;"
			Err = Err & "情報元種類が不正です。<br>"
		End If

		'情報源日付
		If InfoSourceDay <> "" And IsDay(InfoSourceDay) = False Then
			ErrStyle("InfoSourceDay") = "background-color:#ffff00;"
			Err = Err & "情報元日付が不正です。<br>"
		End If

		'扶養フラグ
		If DependentFlag <> "" And IsFlag(DependentFlag) = False Then
			ErrStyle("DependentFlag") = "background-color:#ffff00;"
			Err = Err & "扶養が不正です。<br>"
		End If

		'扶養人数
		If DependentNumber <> "" And IsNumber(DependentNumber, 0, False) = False Then
			ErrStyle("DependentNumber") = "background-color:#ffff00;"
			Err = Err & "扶養人数が不正です。<br>"
		End If

		'配偶者扶養フラグ
		If SpouseFlag <> "" And IsFlag(SpouseFlag) = False Then
			ErrStyle("SpouseFlag") = "background-color:#ffff00;"
			Err = Err & "配偶者扶養が不正です。<br>"
		End If
	End Function

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_Info 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "up_Reg_P_Info '" & vStaffCode & "'" & _
			",'S'" & _
			",'" & Name & "'" & _
			",'" & Name_F & "'" & _
			",'" & SearchName & "'" & _
			",'" & SearchName_F & "'" & _
			",'" & OldName & "'" & _
			",'" & Birthday & "'" & _
			",'" & Sex & "'" & _
			",'" & MarriageFlag & "'" & _
			",'" & Post_U & "'" & _
			",'" & Post_L & "'" & _
			",'" & PrefectureCode & "'" & _
			",'" & City & "'" & _
			",'" & City_F & "'" & _
			",'" & Town & "'" & _
			",'" & Town_F & "'" & _
			",'" & Address & "'" & _
			",'" & Address_F & "'" & _
			",'" & LivingType & "'" & _
			",'" & HomeTelephoneNumber & "'" & _
			",'" & CountryTelephoneNumber & "'" & _
			",'" & PortableTelephoneNumber & "'" & _
			",'" & FaxNumber & "'" & _
			",'" & MailAddress & "'" & _
			",'" & PortableMailAddress & "'" & _
			",'" & UrgencyPost_U & "'" & _
			",'" & UrgencyPost_L & "'" & _
			",'" & UrgencyAddress & "'" & _
			",'" & UrgencyAddress_F & "'" & _
			",'" & UrgencyTelephoneNumber & "'" & _
			",'" & URL & "'" & _
			",'" & InfoSourceType & "'" & _
			",'" & InfoSourceDay & "'" & _
			",'" & InfoSourceOther & "'" & _
			",'" & DependentFlag & "'" & _
			",'" & DependentNumber & "'" & _
			",'" & SpouseFlag & "'" & _
			",'" & CurrentCompanyName & "'" & _
			",'" & CurrentCompanyName_F & "'" & _
			",'" & SocietyInsuranceIn & "'" & _
			",'" & SocietyInsuranceLoss & "'" & _
			",'" & EmployInsuranceIn & "'" & _
			",'" & EmployInsuranceLoss & "'" & vbCrLf
	End Function
End Class
%>
