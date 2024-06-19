<%
'*******************************************************************************
'概　要：HTML文書のDOCTYPE～bodyタグまでを取得
'引　数：vHTML	：</body>の直前に挿入するHTML
'出　力：
'戻り値：String
'備　考：
'履　歴：2010/05/11 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlFooter(ByVal vHTML)
	Dim sHTML

	sHTML = vbCrLf & vHTML & vbCrLf & "</body></html>"

	htmlFooter = sHTML
End Function
%>
