<%
'*******************************************************************************
'�T�@�v�F�{�b�g�̃A�N�Z�X���ǂ������`�F�b�N
'���@���F
'�߂�l�FBoolean
'���@�l�F
'�X�@�V�F2010/08/31 LIS K.Kokubo �쐬
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
