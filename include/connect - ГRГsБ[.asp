<%

'SQLサーバーにコネクトする。
'※ 現在、ODBCドライバ経由でコネクションしているが、
'   ADOで直接コネクトする方法に変更するかを検討
'※ 現在、単なるインクルードファイルになっているがこれをfunction化するかを検討
%>

<%
'If Request.ServerVariables("URL") <> "/maintenance/index.asp" _
'And (Now >= "2015/04/19 09:00:00" And Now <= "2015/04/19 23:00:00") Then
	'Response.Status="302 Found"
'	If Now <= "2015/04/19 23:00:00" Then
	'	Response.AddHeader "Retry-After","Sun, 19 Apr 2015 23:00:00 GMT"
'	End If
'If InStr(Request.ServerVariables("REMOTE_HOST"),"192.168.") = 0 Then
	'Response.redirect "http://www.shigotonavi.co.jp/maintenance/"
'End IF
'End If

Dim dbconn
Dim sServer
Dim sLoginID
Dim sPassword
Dim sDBName

sServer = DBCNSERVERNAME
sLoginID = DBCNLOGINID
sPassword = DBCNPASSWORD
sDBName = DBCNDBNAME

Set dbconn = Server.CreateObject("ADODB.Connection")

dbconn.commandtimeout = 600  '秒
dbconn.connectiontimeout = 600  '秒

'dbconn.Provider = "SQLOLEDB"
'dbconn.ConnectionString = "User ID=sa;Password=;" &_
'						  "Data Source=william;" &_
'						  "Initial Catalog=Person"

'dbconn.ConnectionString = "DRIVER=SQL Server" &_
'						  ";SERVER=" & sServer &_
'						  ";UID=" & sLoginID &_
'						  ";PWD=" & sPassword &_
'						  ";DATABASE=" & sDBName
						  
dbconn.ConnectionString = "Provider=SQLOLEDB;" &_
						  "Password=" & sPassword &_
						  ";Persist Security Info=True" &_
						  ";User ID=" & sLoginID &_
						  ";Initial Catalog=" & sDBName &_
						  ";Data Source=" & sServer & _
						  ";Application Name=SHIGOTONAVI"
dbconn.Open
dbconn.CursorLocation = 3

If Err.Number <> 0 Then
	'エラーメール時の処理
	Call SQLServerStop()
Elseif Application("MailFlag") = "1" Then
	'共有変数削除
	Application.Contents.Remove ("MailFlag")
End if

Function SQLServerStop()
	Dim bobj
	Dim mailto
	Dim rc

	If Application("MailFlag") = "" or Application("MailFlag") = 0 Then
		Set bobj = Server.CreateObject("basp21")

		mailto = ""
		'送信先メールアドレス
		mailto = mailto & "tetsuya-e@docomo.ne.jp" & vbtab			'江崎
		mailto = mailto & "munekyun.nice.guy@ezweb.ne.jp" & vbtab	'小久保
		rc = bobj.SendMail("153.153.150.22",mailto,"info@shigotonavi.jp","【緊急!】SQLサーバー停止",Err.Description,"")
		Application("MailFlag") = 1
	End if
	Response.Redirect ("http://www.shigotonavi.co.jp/underconst.html")
End Function
%>