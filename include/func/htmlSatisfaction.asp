<%
'*******************************************************************************
'概　要：企業満足度アンケートコンテンツの「企業のアンケート回答,リス回答」部分のHTMLを取得
'引　数：vSeq				：アンケート回
'　　　：vAnswerDay			：アンケート回答日(日付型)
'　　　：vSatisfactionPoint	：満足度ポイント
'　　　：vOpinion			：意見
'　　　：vLISAnswer			：リス回答内容
'　　　：vQuestionnaireID	：企業満足度アンケートID
'戻り値：String
'備　考：
'更　新：2009/06/01 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlSatisfaction(ByVal vSeq, ByVal vAnswerDay, ByVal vSatisfactionPoint, ByVal vOpinion, ByVal vLISAnswer, ByVal vQuestionnaireID)
	Dim sHTML

	sHTML = ""

	'<企業回答部分>
	sHTML = sHTML & "<div style=""border-bottom:1px dashed #ccc; margin:0 0 5px;"">"
	'タイトル
	sHTML = sHTML & "<div style=""border-bottom: 3px double #999999;"">"
	sHTML = sHTML & "<div style=""float:left;width:74%;background-color:transparent;"">"
	sHTML = sHTML & "<div style=""padding:4px;font-size:16px;"">"
	Select Case vSatisfactionPoint
		Case 1: sHTML = sHTML & "<span style=""font-weight:bold;"">満足度：<span style=""color:gold;"">★</span></span>　かなり不満"
		Case 2: sHTML = sHTML & "<span style=""font-weight:bold;"">満足度：<span style=""color:gold;"">★★</span></span>　どちらかというと不満"
		Case 3: sHTML = sHTML & "<span style=""font-weight:bold;"">満足度：<span style=""color:gold;"">★★★</span></span>　ふつう"
		Case 4: sHTML = sHTML & "<span style=""font-weight:bold;"">満足度：<span style=""color:gold;"">★★★★</span></span>　どちらかというと満足"
		Case 5: sHTML = sHTML & "<span style=""font-weight:bold;"">満足度：<span style=""color:gold;"">★★★★★</span></span>　かなり満足"
	End Select
	sHTML = sHTML & "&nbsp;(第" & vSeq & "回)"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""float:right;width:25%;background-color:transparent;text-align:right;""><div style=""padding:4px;"">回答日：" & GetDateStr(vAnswerDay, "/") & "</div></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	sHTML = sHTML & "</div>"
	'内容
	sHTML = sHTML & "<div style=""padding:4px;""><b>L</b><p class=""m0"" style=""float:right;width:705px;"">" & Replace(vOpinion, vbCrLf, "<br>") & "</p></div>"
	sHTML = sHTML & "<div clear=""both""></div></div>"
	'</企業回答部分>

	sHTML = sHTML & "<div style=""margin:0 0 5px 15px; border-bottom:1px dashed #ccc;"">"

	'<リス回答部分>
	If vLISAnswer <> "" Then
		'タイトル

		sHTML = sHTML & "<div style=""padding:4px;font-size:16px;font-weight: bold;"">しごとナビからの回答</div>"

		'内容
		sHTML = sHTML & "<div style=""padding:4px;""><b>L</b><p class=""m0"" style=""float:right;width:690px;"">" & Replace(vLISAnswer, vbCrLf, "<br>") & "</p></div>"
		sHTML = sHTML & "<div clear=""both""></div>"
	End If
	'</リス回答部分>

	sHTML = sHTML & "</div>"

	htmlSatisfaction = sHTML
End Function
%>
