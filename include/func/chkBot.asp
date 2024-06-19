<%
'*******************************************************************************
'概　要：ボットのアクセスかどうかをチェック
'引　数：
'戻り値：Boolean
'備　考：
'更　新：2010/08/31 LIS K.Kokubo 作成
'*******************************************************************************
Function chkBot(ByVal vUserAgent)
	Dim oRE
	Dim sPattern

	sPattern = "(bitlybot)|(Twitterbot)|(PycURL)|(Twingly Recon)|(\$_agentname)|(PEAR HTTP_Request)|(PostRank)|(Python-urllib)|" & _
		"(Yahoo! Slurp)|(Googlebot)|(Butterfly)|(kmbot)|(mxbot)|(TweetmemeBot)|(OneRiot)|(NjuiceBot)"

	Set oRE = New RegExp
	oRE.IgnoreCase = True
	oRE.Pattern = sPattern

	chkBot = oRE.Test(vUserAgent)
End Function
%>
