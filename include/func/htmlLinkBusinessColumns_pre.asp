<%
'*******************************************************************************
'概　要：しごとナビビジネスコラム - プレゼンテーションカテゴリの各コラムリンクHTMLを取得
'引　数：
'戻り値：String
'備　考：
'履　歴：2010/09/03 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlLinkBusinessColumns_pre(vUserID,vCurrentURL)
	Dim sHTML

	Dim aDate(3)
	Dim idx
	Dim sTitle
	Dim iNo

	iNo = 0

	aDate(0) = CDate("2010/11/01")
	For idx = 1 To UBound(aDate)
		aDate(idx) = aDate(idx-1)+7
	Next


	sTitle = "準備"
	iNo = iNo + 1
	sHTML = sHTML & "<div style=""text-align:right;""><p class=""m0"">"
	sHTML = sHTML & "&nbsp;≫&nbsp;"
	If Date >= aDate(iNo-1) Then
		sHTML = sHTML & "<a href=""/s_contents/businesscolumns/pre" & Right("0"&iNo,2) & ".asp"">" & sTitle & "</a>"
		If Date < aDate(iNo-1)+7 Then sHTML = sHTML & "<img src=""/img/new_title/new2.gif"" alt=""NEW!"">"
	Else
		sHTML = sHTML & sTitle
	End If
	sHTML = sHTML & "</p></div>"

	sTitle = "プレゼン方法"
	iNo = iNo + 1
	sHTML = sHTML & "<div style=""text-align:right;""><p class=""m0"">"
	sHTML = sHTML & "&nbsp;≫&nbsp;"
	If Date >= aDate(iNo-1) Then
		sHTML = sHTML & "<a href=""/s_contents/businesscolumns/pre" & Right("0"&iNo,2) & ".asp"">" & sTitle & "</a>"
		If Date < aDate(iNo-1)+7 Then sHTML = sHTML & "<img src=""/img/new_title/new2.gif"" alt=""NEW!"">"
	Else
		sHTML = sHTML & sTitle
	End If
	sHTML = sHTML & "</p></div>"

	sTitle = "会議"
	iNo = iNo + 1
	sHTML = sHTML & "<div style=""text-align:right;""><p class=""m0"">"
	sHTML = sHTML & "&nbsp;≫&nbsp;"
	If Date >= aDate(iNo-1) Then
		sHTML = sHTML & "<a href=""/s_contents/businesscolumns/pre" & Right("0"&iNo,2) & ".asp"">" & sTitle & "</a>"
		If Date < aDate(iNo-1)+7 Then sHTML = sHTML & "<img src=""/img/new_title/new2.gif"" alt=""NEW!"">"
	Else
		sHTML = sHTML & sTitle
	End If
	sHTML = sHTML & "</p></div>"

	sHTML = sHTML & "<div style=""clear:both;""></div>"


	htmlLinkBusinessColumns_pre = sHTML
End Function
%>
