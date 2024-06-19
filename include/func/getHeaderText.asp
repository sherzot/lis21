<%
'*******************************************************************************
'概　要：タブIndexを取得
'引　数：
'戻り値：String
'備　考：
'履　歴：2011/03/16 LIS K.Kokubo 作成
'*******************************************************************************
Function getHeaderText(ByVal vHeadType,ByVal vURL)
	Dim sText

	vURL = LCase(vURL)

	If vHeadType = 0 Then 'トップ
		sText = "転職サイト「しごとナビ」。正社員・派遣の求人情報はもちろん、プロによる貴方に適した転職サポートをご提供しています！"
	ElseIf vHeadType = 1 Then '求職者
		sText = "転職活動・求職活動の方々に最適な求人情報と履歴書ツールを提供しています"
	ElseIf vHeadType = 2 Or vHeadType = 4 Then '企業
		sText = "企業の人材雇用を幅広くサポートしております。（求人広告、人材派遣、人材紹介）"
	ElseIf vHeadType = 3 Then '共用
		sText = "求職活動の方々に最適な求人情報と履歴書ツールを提供しています"
	End If

	If InStr(vURL,"/staff/s_careersheet.asp") > 0 Then
		sText = "職務経歴書作成・印刷｜しごとナビの簡単便利な職務経歴書の自動作成・印刷"
	ElseIf InStr(vURL,"/order/special/ad/0001/") > 0 Then
		sText = "SE転職特集｜SEの転職・求人情報や転職支援サービスを行っています。"
	ElseIf InStr(vURL,"/order/special/tg/0004/") > 0 Then
		sText = "臨床検査技師 求人特集｜臨床検査技師の資格を活かした求人のご紹介と、無料の転職支援サービスのご案内。"
	ElseIf InStr(vURL,"/order/special/tg/0005/") > 0 Then
		sText = "英語を活かせる派遣特集｜英語に特化した派遣スタッフさんの登録会を全国で行っています。"
	ElseIf InStr(vURL,"/order/special/sz/0001/") > 0 Then
		sText = "静岡 転職特集｜地域に特化した求人の提供と、無料の転職支援サービスを行っています。"
	ElseIf InStr(vURL,"/order/special/ng/0002/") > 0 Then
		sText = "名古屋 派遣特集｜地域に特化した派遣スタッフさんの登録会を全国で行っています。"
	ElseIf InStr(vURL,"/order/special/or/0001/") > 0 Then
		sText = "DTP 求人特集｜DTPオペレータ,デザイナーに特化した求人と、無料の転職支援サービスのご案内。"
	ElseIf InStr(vURL,"/order/special/oy/0001/") > 0 Then
		sText = "岡山 求人特集｜地域に特化した求人の提供と、無料の転職支援サービスを行っています。"
	ElseIf InStr(vURL,"/order/special/hr/0001/") > 0 Then
		sText = "広島 転職特集｜地域に特化した求人の提供と、無料の転職支援サービスを行っています。"
	End If

	If vHeadType <> 0 Then
		sText = "<strong style=""color:#666666;"">" & sText & "</strong>"
	Else
		sText = "<b style=""color:#666666;"">" & sText & "</b>"
	End If

	getHeaderText = sText
End Function
%>
