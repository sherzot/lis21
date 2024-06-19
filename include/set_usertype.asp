<%
Dim USERTYPE	: USERTYPE = Session("usertype")
Dim USERID		: USERID = Session("userid")
Dim COMPANYTYPE	: COMPANYTYPE = Session("companytype")
Dim UserPass
Dim sJumpParam_set_usertype
Dim sQS_set_usertype

userid = session("userid")

'セッション情報が存在しない場合は、ログイン画面を表示する
If UserPass = true Then
Else
	If session("usertype") = "" Then
		sJumpParam_set_usertype = "jump_url_flag=true"
		sJumpParam_set_usertype = sJumpParam_set_usertype & "&jump_url=" & Request.ServerVariables("URL")
		for each sQS_set_usertype in Request.QueryString
			sJumpParam_set_usertype = sJumpParam_set_usertype & "&" & sQS_set_usertype & "=" & Request.QueryString(sQS_set_usertype)
		next
		'response.write sJumpParam_set_usertype
	
	If Instr(Request.ServerVariables("URL"),"/company/") > 0 Or Instr(Request.ServerVariables("URL"),"/company/") > 0 Then
		'企業の場合
		Response.Redirect (HTTP_CURRENTURL & "login_menu.asp?" & sJumpParam_set_usertype)
	Else
		'求職者の場合
		Response.Redirect (HTTP_CURRENTURL & "index.asp?" & sJumpParam_set_usertype)
	End If
		Response.End
	End If
End If
%>
