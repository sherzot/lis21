<%
'*******************************************************************************
'概　要：タブIndexを取得
'引　数：
'戻り値：String
'備　考：
'履　歴：2010/05/13 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlTabIndex(ByVal vURL,ByVal vUserType,ByVal vHeadComment)
	Dim sHTML
	Dim aPattern
	Dim aLink(5)
	Dim iTabIndexType
	Dim sBorderColor

	aPattern = Array(1,1,1,1,1,1,1,1)
	vURL = LCase(vURL)

	iTabIndexType = getTabIndexType(vURL)
	sBorderColor = "#66cc33"
	If iTabIndexType = 5 Or iTabIndexType = 7 Then sBorderColor = "#0033cc"

	Select Case iTabIndexType
	'はじめての方タブ
	Case 0: aPattern(0) = 2
	'求人を探すタブ
	Case 1: aPattern(1) = 2
	'便利ツールタブ
	Case 2: aPattern(2) = 2
	'転職サポートタブ
	Case 3: aPattern(3) = 2
	'コミュニティタブ
	Case 4: aPattern(4) = 2
	'採用ご担当者タブ
	Case 5: aPattern(5) = 2
	'登録会員タブ
	Case 6: aPattern(6) = 2
	'ログイン企業タブ
	Case 7: aPattern(7) = 2
	'Case Else: aPattern(1) = 2
	End Select

'	If vUserType = "staff" Then
'		aLink(0) = "<a href=""" & HTTPS_CURRENTURL & "tab/index6.asp"" title=""ようこそ！　しごとナビへ！　あなたのためのメニュー一覧です。""><img src=""/img/common/tab_index7_" & aPattern(6) & ".png"" alt=""ようこそ！　しごとナビへ！　あなたのためのメニュー一覧です。"" border=""0""></a>"
'	ElseIf vUserType = "company" Then
'		aLink(0) = "<a href=""" & HTTPS_CURRENTURL & "tab/index7.asp"" title=""ようこそ！　しごとナビへ！　あなたのためのメニュー一覧です。""><img src=""/img/common/tab_index8_" & aPattern(7) & ".png"" alt=""ようこそ！　しごとナビへ！　あなたのためのメニュー一覧です。"" border=""0""></a>"
'	Else
'		aLink(0) = "<a href=""" & HTTPS_CURRENTURL & "tab/index1.asp"" title=""会員登録（履歴書登録）するメリットとは？""><img src=""/img/common/tab_index1_" & aPattern(0) & ".png"" alt=""会員登録（履歴書登録）するメリットとは？"" border=""0""></a>"
'	End If

	aLink(0) = "<a href=""" & HTTPS_CURRENTURL & """>HOME</a>"
	aLink(1) = "<a href=""" & HTTPS_CURRENTURL & "search/"">しごとを探す</a>"	
	'aLink(2) = "<a href=""" & HTTPS_CURRENTURL & "koryu/"">交流</a>"
	aLink(3) = "<a href=""" & HTTPS_CURRENTURL & "manabu/"">学ぶ</a>"
	aLink(4) = "<a href=""" & HTTPS_CURRENTURL & "link/"">リンク</a>"
	aLink(5) = "<a href=""" & HTTPS_CURRENTURL & "company_AB/"">採用企業</a>"
	


	sHTML = ""
	sHTML = sHTML & "<nav class=""headtab"">"
	sHTML = sHTML & "<div id=""line_on"">"	
	sHTML = sHTML & "<ul>"
	sHTML = sHTML & "<li id=""gmenu01"">" & aLink(1) & "</li>"
	'sHTML = sHTML & "<li id=""gmenu02"">" & aLink(2) & "</li>"
	sHTML = sHTML & "<li id=""gmenu03"">" & aLink(3) & "</li>"
	sHTML = sHTML & "<li id=""gmenu04"">" & aLink(4) & "</li>"
	sHTML = sHTML & "<li id=""gmenu05"">" & aLink(5) & "</li>"
	sHTML = sHTML & "<li id=""gmenu06"">" & aLink(0) & "</li>"
	sHTML = sHTML & "</ul>"
	sHTML = sHTML & "<img src=""/img/border.gif"" id=""slide_border""></div>"
	
	sHTML = sHTML & "</nav>"

	htmlTabIndex = sHTML
End Function
%>
