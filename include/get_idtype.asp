<%
Dim sJumpParam_get_idtype
Dim sQS_get_idtype

'�Z�b�V������񂪑��݂��Ȃ��ꍇ�́A���O�C����ʂ�\������
If Session("usertype") = "" Then
	sJumpParam_get_idtype = "JUMP_URL_FLAG=True"
	sJumpParam_get_idtype = sJumpParam_get_idtype & "&JUMP_URL=" & Request.ServerVariables("URL")
	For Each sQS_get_idtype In Request.QueryString
		sJumpParam_get_idtype = sJumpParam_get_idtype & "&" & sQS_get_idtype & "=" & Request.QueryString(sQS_get_idtype)
	Next

	If InStr(LCase(Request.ServerVariables("URL")),"/company/") > 0 Then
		'��Ƃ̏ꍇ
		Response.Redirect ("/login_menu.asp?" & sJumpParam_get_idtype)
	Else
		'���E�҂̏ꍇ
		Response.Redirect ("/index.asp?" & sJumpParam_get_idtype)
	End if
	Response.Flush
End If
%>
