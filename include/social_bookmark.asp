<% 'ソーシャルブックマーク st

dim allurl
allurl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
	if Request.ServerVariables("QUERY_STRING") <> "" then
		allurl = allurl & "?" & Request.ServerVariables("QUERY_STRING")
	end if

Response.Write"<div class=""left sb_waku"">" & VbCrlf

' Buzzurlブックマーク
'Response.Write"<a class=""sb_button"" href=""http://buzzurl.jp/entry/http://" & allurl & """><img src=""../../img/top/bookmark_b_buzzurl_s.gif"" border=""0"" alt=""""></a>"

' livedoorクリップブックマーク
'Response.Write"<a class=""sb_button"" href=""javascript:window.location='http://clip.livedoor.com/redirect?link=" & "http://" & allurl & "&amp;title='+escape(document.title)"" rel=""nofollow"" title=""livedoorクリップに登録""><img src=""../../img/top/bookmark_livedoor_clip.gif"" width=""16"" height=""16"" alt=""livedoorクリップに登録"" border=""0""></a>"

' FC2ブックマーク　閉鎖
'Response.Write"	<a href=""javascript:location.href='http://bookmark.fc2.com/user/post?url=" & "http://" & allurl & "&amp;title='+encodeURIComponent(document.title)"" title=""FC2ブックマークに登録""><img alt=""FC2ブックマーク"" src=""../../img/top/bookmark_fc2_clip.gif"" width=""16"" height=""16"" border=""0""> FC2ブクマ</a>"

'del.icio.usブックマーク　休止
'Response.Write"	<IMG SRC=""../../img/top/delicious_add.gif"" alt=""""><a href=""javascript:window.location='http://del.icio.us/post?url=" & "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "&amp;title='+escape (document.title);"">del.icio.us</a>"
Response.Write"</div>"

'ソーシャルブックマーク　ed %>