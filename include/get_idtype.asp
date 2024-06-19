<%
Dim sJumpParam_get_idtype
Dim sQS_get_idtype

'セッション情報が存在しない場合は、ログイン画面を表示する
If Session("usertype") = "" Then
	sJumpParam_get_idtype = "JUMP_URL_FLAG=True"
	sJumpParam_get_idtype = sJumpParam_get_idtype & "&JUMP_URL=" & Request.ServerVariables("URL")
	For Each sQS_get_idtype In Request.QueryString
		sJumpParam_get_idtype = sJumpParam_get_idtype & "&" & sQS_get_idtype & "=" & Request.QueryString(sQS_get_idtype)
	Next

	If InStr(LCase(Request.ServerVariables("URL")),"/company/") > 0 Then
		'企業の場合
		Response.Redirect ("/login_menu.asp?" & sJumpParam_get_idtype)
	Else
		'求職者の場合
		Response.Redirect ("/index.asp?" & sJumpParam_get_idtype)
	End if
	Response.Flush
End If
%>
