<%@ Language="VBScript" CodePage="932" %>
<%

	'## �ϐ��錾������
	Option Explicit
	'On Error Resume Next
	'## ��������������܂Ńy�[�W���o�͂��Ȃ�
	Response.buffer = True
	'## �y�[�W���L���b�V�����Ȃ�
	Response.Expires = 0
	
%>
<!-- #INCLUDE VIRTUAL="/config/personnel.asp" -->
<!-- #INCLUDE VIRTUAL="/include/connect.asp" -->
<!-- #INCLUDE VIRTUAL="/include/commonfunc.asp" -->
<!-- #INCLUDE VIRTUAL="/include/set_usertype.asp" -->
<!-- #INCLUDE VIRTUAL="/include/func_navigation.asp" -->

<%
Dim sPageTitle
Dim sPageKeyword
Dim sPageDescription
Dim sAddHead
Dim sBodyAttribute

sPageTitle = "�V�K�t�@�C���̓o�^"
sPageKeyword = "�V�K�t�@�C���̓o�^"
sPageDescription = "�V�K�t�@�C���̓o�^"
sAddHead = "<link rel=""stylesheet"" type=""text/css"" href=""/css/style_main.css"">"

Response.Write htmlHeader(CURRENTURL,sPageTitle,sPageKeyword,sPageDescription,sAddHead,False,False,False,True,sBodyAttribute)
%>
</head><body>
<%
	Call NaviHeader(1)'0�i�g�b�v�j1�i���E�ҁj2�i��Ɓj3�i���L�j
%>
<div>
	<p style="font-size:20px;">
		�V�K�t�@�C���̓o�^
	</p>

	<p style="font-size:16px;">
		�V�K�t�@�C���̓o�^���������܂����B<br>�v���r���[�͂��q�l�̃��[���A�h���X���Ɂu�o�^�����ʒm�v�̃��[�������m�F���������B
	</p>

	<p>
		<button onclick="location.href='https://www.shigotonavi.co.jp/staff/resume_print.asp'"><span>�߂�</span></button>
	</p>
</div>
<%
Call NaviSidemenu(1)
Call NaviFooter()
Response.Write htmlFooter("")
%>
</body></html>
