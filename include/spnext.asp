<%@ Language=VBScript CodePage=932 %>
<!-- #INCLUDE VIRTUAL="/config/personnel.asp" -->
<!-- #INCLUDE VIRTUAL="/include/connect.asp" -->
<!-- #INCLUDE VIRTUAL="/include/commonfunc.asp" -->
<%
response.Write "<html>"
response.write "<head>"
response.write "<style type=""text/css"">"
response.write "H1{font-size:small;color:red;}"
response.write "p{margin:0px;padding:2px;font-size:12px}"
response.write "form{margin:0px;padding:0px;}"
response.write "</style>"
response.write "</head>"
response.write "<body style=""padding-top:46px;background-image:url(/img/sidemenu/smartphonebunner.gif);"">"

Dim Commit	:	Commit = false
Dim ErrMsg	:	ErrMsg = ""
Dim oRS,sSQL,sErr

if IsMailAddress(GetForm("mailaddress",1)) = true then
	
	sSubject = "�y�����ƃi�r�z�X�}�[�g�t�H���T�C�gURL"
	sBody = "�����p���肪�Ƃ��������܂��B�����ƃi�r�̃X�}�[�g�t�H���T�C�g�ւ́A����URL����A�N�Z�X�����肢�v���܂��B" & vbCrLf
	sBody = sBody & "http://sp.shigotonavi.jp/?ml=1" & vbCrLf
	sBody = sBody & "���{���[���Ɋo�����Ȃ��ꍇ�́A�폜���Ă��������B"
	if SndMail(MAIL_SERVER,trim(GetForm("mailaddress",1)),"info@shigotonavi.jp",sSubject,sBody,"") = true then
		Commit = true
	else
		ErrMsg = "���[���̑��M�Ɏ��s���܂����B"
	end if
elseif trim(GetForm("mailaddress",1)) <> "" then
	ErrMsg = "�A�h���X������������܂���B"
end if

if Commit = true then
	response.write "<p>"
	response.write "���w���̃A�h���X�ցA���ē����[���𑗐M�v���܂����B"
	response.write "</p>"
else
	response.write "<p>"
	response.write "�����p�̕��́A�ȉ��Ƀ��[���A�h���X����͂��u���M�v�������Ă��������B"
	response.Write "</p>"
	response.write "<form action=""spnext.asp?reg=1"" Method=""post"">"
	response.write "<input type=""text"" name=""mailaddress"" value="""
	response.write GetForm("mailaddress",1)
	response.write """ placeholder=""account@domain.jp"" style=""width:100px;"">"
	response.write "<input type=""submit"" value=""���M"">"
	if ErrMsg <> "" then
		response.write "<span style=""color:red;font-size:x-small"">"
		response.write ErrMsg
		response.write "</span>"
	end if
	response.write "</form>"

end if
%>
</body>
</html>