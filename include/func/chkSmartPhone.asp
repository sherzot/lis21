<%
'*******************************************************************************
'概　要：スマートフォンのアクセスかどうかをチェック
'引　数：
'戻り値：Boolean
'備　考：
'履　歴：2011/06/30 LIS K.Kokubo 作成
'*******************************************************************************
Function chkSmartPhone(ByVal vUserAgent)
	Dim oRE
	Dim sPattern

	sPattern = "((iPhone)|(Android)|(BlackBerry)|(Symbian))"

	Set oRE = New RegExp
	oRE.IgnoreCase = True
	oRE.Pattern = sPattern

	chkSmartPhone = oRE.Test(vUserAgent)
End Function
%>
