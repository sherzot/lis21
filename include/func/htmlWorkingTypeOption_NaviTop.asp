<%
'*******************************************************************************
'概　要：しごとナビTOPのカンタン検索で使用する勤務形態一覧の<option></option>を取得
'引　数：vCode			：チェック中のコード
'　　　：vAttribute		：optionの追加属性
'戻り値：String
'備　考：
'履　歴：2011/02/04 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlWorkingTypeOption_NaviTop(ByVal vCode, ByVal vAttribute)
	Dim sHTML
	Dim aCode(9)

	sHTML = ""

	Select Case vCode
		Case "001": aCode(1) = " selected"
		Case "002": aCode(2) = " selected"
		Case "003": aCode(3) = " selected"
		Case "004": aCode(4) = " selected"
		Case "005": aCode(5) = " selected"
		Case "006": aCode(6) = " selected"
		Case "007": aCode(7) = " selected"
		Case "009": aCode(8) = " selected"
		Case "100": aCode(9) = " selected"
	End Select

	If vAttribute <> "" Then vAttribute = " " & vAttribute

	sHTML = sHTML & "<option value=""001""" & aCode(1) & ">派遣</option>"
	sHTML = sHTML & "<option value=""002""" & aCode(2) & ">正社員</option>"
	sHTML = sHTML & "<option value=""003""" & aCode(3) & ">契約社員</option>"
	sHTML = sHTML & "<option value=""004""" & aCode(4) & ">紹介予定派遣</option>"
	sHTML = sHTML & "<option value=""005""" & aCode(5) & ">パート・アルバイト</option>"
	sHTML = sHTML & "<option value=""006""" & aCode(6) & ">SOHO(在宅・副業)</option>"
	sHTML = sHTML & "<option value=""007""" & aCode(7) & ">FC・代理店</option>"
	sHTML = sHTML & "<option value=""009""" & aCode(8) & ">経営者・役員</option>"
	sHTML = sHTML & "<option value=""100""" & aCode(9) & ">新卒</option>"

	htmlWorkingTypeOption_NaviTop = sHTML
End Function
%>
