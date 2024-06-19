<%
'*******************************************************************************
'概　要：人材紹介ツイッターのFollow Me バッヂを表示するjavascript
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2010/08/31 LIS K.Kokubo 作成
'*******************************************************************************
Function scrIntroTwitterFollowBadge()
	Dim sScript

	If Request.ServerVariables("HTTPS") = "off" Then
		sScript = vbCrLf
		sScript = sScript & "<!-- twitter follow badge by go2web20 -->"
		sScript = sScript & "<script src=""http://www.go2web20.net/twitterfollowbadge/1.0/badge.js"" type=""text/javascript""></script>"
		sScript = sScript & "<script type=""text/javascript"" charset=""utf-8""><!--" & vbCrLf
		sScript = sScript & "tfb.account = 'jinzai_navi';"
		sScript = sScript & "tfb.label = 'follow-me';"
		sScript = sScript & "tfb.color = '#5573b7';"
		sScript = sScript & "tfb.side = 'r';"
		sScript = sScript & "tfb.top = 136;"
		sScript = sScript & "tfb.showbadge();" & vbCrLf
		sScript = sScript & "--></script>"
		sScript = sScript & "<!-- end of twitter follow badge -->" & vbCrLf
	End If

	scrIntroTwitterFollowBadge = sScript
End Function
%>
