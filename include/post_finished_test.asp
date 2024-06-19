<%@ Language="VBScript" CodePage="932" %>
<%

	'## 変数宣言を強制
	Option Explicit
	'On Error Resume Next
	'## 処理が完了するまでページを出力しない
	Response.buffer = True
	'## ページをキャッシュしない
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

sPageTitle = "新規ファイルの登録"
sPageKeyword = "新規ファイルの登録"
sPageDescription = "新規ファイルの登録"
sAddHead = "<link rel=""stylesheet"" type=""text/css"" href=""/css/style_main.css"">"

Response.Write htmlHeader(CURRENTURL,sPageTitle,sPageKeyword,sPageDescription,sAddHead,False,False,False,True,sBodyAttribute)
%>
</head><body>
<%
	Call NaviHeader(1)'0（トップ）1（求職者）2（企業）3（共有）
%>
<div>
	<p style="font-size:20px;">
		新規ファイルの登録
	</p>

	<p style="font-size:16px;">
		新規ファイルの登録が完了しました。<br>プレビューはお客様のメールアドレス宛に「登録完了通知」のメールをご確認ください。
	</p>

	<p>
		<button onclick="location.href='https://www.shigotonavi.co.jp/staff/resume_print.asp'"><span>戻る</span></button>
	</p>
</div>
<%
Call NaviSidemenu(1)
Call NaviFooter()
Response.Write htmlFooter("")
%>
</body></html>
