<% '�\�[�V�����u�b�N�}�[�N st

dim allurl
allurl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
	if Request.ServerVariables("QUERY_STRING") <> "" then
		allurl = allurl & "?" & Request.ServerVariables("QUERY_STRING")
	end if

Response.Write"<div class=""left sb_waku"">" & VbCrlf

' Buzzurl�u�b�N�}�[�N
'Response.Write"<a class=""sb_button"" href=""http://buzzurl.jp/entry/http://" & allurl & """><img src=""../../img/top/bookmark_b_buzzurl_s.gif"" border=""0"" alt=""""></a>"

' livedoor�N���b�v�u�b�N�}�[�N
'Response.Write"<a class=""sb_button"" href=""javascript:window.location='http://clip.livedoor.com/redirect?link=" & "http://" & allurl & "&amp;title='+escape(document.title)"" rel=""nofollow"" title=""livedoor�N���b�v�ɓo�^""><img src=""../../img/top/bookmark_livedoor_clip.gif"" width=""16"" height=""16"" alt=""livedoor�N���b�v�ɓo�^"" border=""0""></a>"

' FC2�u�b�N�}�[�N�@��
'Response.Write"	<a href=""javascript:location.href='http://bookmark.fc2.com/user/post?url=" & "http://" & allurl & "&amp;title='+encodeURIComponent(document.title)"" title=""FC2�u�b�N�}�[�N�ɓo�^""><img alt=""FC2�u�b�N�}�[�N"" src=""../../img/top/bookmark_fc2_clip.gif"" width=""16"" height=""16"" border=""0""> FC2�u�N�}</a>"

'del.icio.us�u�b�N�}�[�N�@�x�~
'Response.Write"	<IMG SRC=""../../img/top/delicious_add.gif"" alt=""""><a href=""javascript:window.location='http://del.icio.us/post?url=" & "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "&amp;title='+escape (document.title);"">del.icio.us</a>"
Response.Write"</div>"

'�\�[�V�����u�b�N�}�[�N�@ed %>