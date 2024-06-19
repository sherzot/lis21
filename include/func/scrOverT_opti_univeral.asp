<%
'*******************************************************************************
'概　要：オーバーチュアタグ取得
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2010/05/11 LIS K.Kokubo 作成
'*******************************************************************************
Function scrOverT_opti_univeral()
	Dim sScript

'	sScript = vbCrLf
'	sScript = sScript & "<SCRIPT TYPE=""text/javascript"" LANGUAGE=""JavaScript"">" & vbCrLf
'	sScript = sScript & "<!--" & vbCrLf
'sScript = sScript & "/*" & vbCrLf
'	sScript = sScript & "var pm_tagname    = 'universalTag.txt';" & vbCrLf
'	sScript = sScript & "var pm_tagversion = '1.4';" & vbCrLf
'	sScript = sScript & "var pm_accountid  = '09FIKJ4M44S4GM2B78CVG78KJ4';" & vbCrLf
'	sScript = sScript & "var pm_scripthost = 'srv.perf.overture.com';" & vbCrLf
'	sScript = sScript & "var pm_customargs = '';" & vbCrLf
'	sScript = sScript & "var pm_querystr = '?' + 'ver=' + pm_tagversion + '&aid=' + pm_accountid + pm_customargs;" & vbCrLf
'	sScript = sScript & "var pm_tag = '<SCR' + 'IPT LANGUAGE=""JavaScript"" ' + 'SRC=//' + pm_scripthost + '/collweb/ScriptServlet' + pm_querystr + '><' + '/SCRIPT>';" & vbCrLf
'	sScript = sScript & "document.write(pm_tag);" & vbCrLf
'sScript = sScript & "*/" & vbCrLf
'	sScript = sScript & "// -->" & vbCrLf
'	sScript = sScript & "</SCRIPT>" & vbCrLf

	scrOverT_opti_univeral = sScript
End Function
%>
