<%
'*******************************************************************************
'概　要：リファラ解析タグ取得
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2010/05/11 LIS K.Kokubo 作成
'*******************************************************************************
Function scrRefAll()
	Dim sScript

	Dim id
	Dim refer

	id = request.querystring("id")	'ID所得

	If IsNumeric(id) = True Then
		If id = 1 Then
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume1.js""></script>"
		ElseIf id = 2 Then 'Overture「履歴書」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume2.js""></script>"
		ElseIf id = 3 Then
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume3.js""></script>"
		ElseIf id = 4 Then
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume4.js""></script>"
		ElseIf id = 5 Then
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume5.js""></script>"
		ElseIf id = 6 Then 'Google「転職」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume6.js""></script>"
		ElseIf id = 7 Then 'Google「就職」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume7.js""></script>"
		ElseIf id = 8 Then 'Google「求人」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume8.js""></script>"
		ElseIf id = 9 Then 'Google「面接」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume9.js""></script>"
		ElseIf id = 10 Then '毎日インタラクティブメール
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume10.js""></script>"
		ElseIf id = 11 Then 'Overture[職務経歴書] 2003/09/12ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume11.js""></script>"
		ElseIf id = 12 Then 'JListing[履歴書] 2004/04/28ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume12.js""></script>"
		ElseIf id = 13 Then 'JListing[職務経歴書] 2004/04/28ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume13.js""></script>"
		ElseIf id = 14 Then 'Overture[面接] 2004/07/22ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume14.js""></script>"
		ElseIf id = 15 Then 'Overture[退職] 2004/07/22ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume15.js""></script>"
		ElseIf id = 16 Then 'Overture[履歴書の書き方] 2004/07/22ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume16.js""></script>"
		ElseIf id = 17 Then 'プレジデントビジョン 2004/09/3配信
			Session("ref") = "president"
		ElseIf id = 18 Then 'JINZAIメルマガ 2004/09/6配信
			Session("ref") = "jinzai-mail"
		ElseIf id = 19 Then 'Overture「人材」 2004/09/16
			Session("ref") = "overture_01"
		ElseIf id = 20 Then 'Overture「その他」 2004/09/16
			Session("ref") = "overture_02"
		ElseIf id = 21 Then 'サイボウズ　テキスト広告 2004/09/20～26
			Session("ref") = "cybozu"
		ElseIf id = 22 Then 'しごと情報ネット 2004/10/22
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume17.js""></script>"
		ElseIf id = 23 Then 'e-words 2004/11/24
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume18.js""></script>"
		ElseIf id = 24 Then 'Overture【12月追加C向け】 2004/12/22
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume19.js""></script>"
		ElseIf id = 25 Then 'MSNのオススメ
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume25.js""></script>"
		ElseIf id = 26 Then 'Adwords新追加履歴書
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume1.js""></script>"
			'sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume26.js""></script>"
		ElseIf id = 27 Then 'oveture新追加履歴書
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume2.js""></script>"
			'sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume27.js""></script>"
		ElseIf id = 28 Then 'oveture「志望動機」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_column_1_ot.js""></script>"
		ElseIf id = 29 Then 'adwords「志望動機」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_column_1_aw.js""></script>"
		ElseIf id = 30 Then 'overture「職種別ページ」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_shoku_search_1_ot.js""></script>"
		ElseIf id = 31 Then 'adwords「職種別ページ」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_shoku_search_1_aw.js""></script>"
		ElseIf id = 32 Then 'Jリスティング「志望動機」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_column_1_jlisting.js""></script>"
		ElseIf id = 33 Then 'overture「退職願」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_retire_want_ot.js""></script>"
		ElseIf id = 34 Then 'adword「退職願」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_retire_want_ad.js""></script>"
		ElseIf id = 35 Then 'overture「自己PR」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_Mypr_ot.js""></script>"
		ElseIf id = 36 Then 'adword「自己PR」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_Mypr_ad.js""></script>"
		ElseIf id = 37 Then 'overture「ニート」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_neet_ot.js""></script>"
		ElseIf id = 38 Then 'adword「ニート」
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_neet_ad.js""></script>"
		Else 
			refer = Request.ServerVariables("HTTP_REFERER") 
			If Left(refer,22) <> "http://www.shigotonavi" Then
				If Left(refer,23) <> "https://www.shigotonavi" Then
					If Left(refer,22) <> "www.shigotonavi.co.jp/" Then
						If refer <> "" Then
							sScript = "<script type=""text/javascript"" src=""/java-script/refer_c_hajime.js""></script>"
						End If
					End If
				End If
			End If
		End If
	End If

	scrRefAll = sScript & vbCrLf
End Function
%>
