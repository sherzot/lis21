<%
'*******************************************************************************
'�T�@�v�FTwitter��Follow Me �o�b�a��\������javascript
'���@���F
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2010/06/24 LIS K.Kokubo �쐬
'*******************************************************************************
Function scrTwitterFollowBadge()
	Dim sScript

	If Request.ServerVariables("HTTPS") = "off" Then
		sScript = vbCrLf
		sScript = sScript & "<!-- twitter follow badge by go2web20 -->"
		sScript = sScript & "<script src=""http://www.go2web20.net/twitterfollowbadge/1.0/badge.js"" type=""text/javascript""></script>"
		sScript = sScript & "<script type=""text/javascript"" charset=""utf-8""><!--" & vbCrLf
		sScript = sScript & "tfb.account = 'shigoto_navi';"
		sScript = sScript & "tfb.label = 'follow-me';"
		sScript = sScript & "tfb.color = '#67cb35';"
		sScript = sScript & "tfb.side = 'r';"
		sScript = sScript & "tfb.top = 136;"
		sScript = sScript & "tfb.showbadge();" & vbCrLf
		sScript = sScript & "--></script>"
		sScript = sScript & "<!-- end of twitter follow badge -->" & vbCrLf
	End If

	scrTwitterFollowBadge = sScript
End Function
%>
