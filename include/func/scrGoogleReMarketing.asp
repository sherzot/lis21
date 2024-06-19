<%
'*******************************************************************************
'概　要：Googleリマーケティング：TOPページに貼るタグ取得
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2011/01/28 LIS K.Kokubo 作成
'*******************************************************************************
Function scrGoogleRemarketing()
	Dim sScript

	sScript = vbCrLf
	sScript = sScript & "<!-- Google Code for &#12522;&#12510;&#12540;&#12465;&#12486;&#12451;&#12531;&#12464;_Top&#12506;&#12540;&#12472;&#29992;&#12479;&#12464; Remarketing List -->" & vbCrLf
	sScript = sScript & "<script type=""text/javascript"">" & vbCrLf
	sScript = sScript & "/* <![CDATA[ */" & vbCrLf
	sScript = sScript & "var google_conversion_id = 1070319369;" & vbCrLf
	sScript = sScript & "var google_conversion_language = ""en"";" & vbCrLf
	sScript = sScript & "var google_conversion_format = ""3"";" & vbCrLf
	sScript = sScript & "var google_conversion_color = ""666666"";" & vbCrLf
	sScript = sScript & "var google_conversion_label = ""AFWuCNmS9wEQiY6v_gM"";" & vbCrLf
	sScript = sScript & "var google_conversion_value = 0;" & vbCrLf
	sScript = sScript & "/* ]]> */" & vbCrLf
	sScript = sScript & "</script>" & vbCrLf
	sScript = sScript & "<script type=""text/javascript"" src=""https://www.googleadservices.com/pagead/conversion.js"">" & vbCrLf
	sScript = sScript & "</script>" & vbCrLf
	sScript = sScript & "<noscript>" & vbCrLf
	sScript = sScript & "<div style=""display:inline;"">" & vbCrLf
	'sScript = sScript & "<img height=""1"" width=""1"" style=""border-style:none;"" alt="""" src=""https://www.googleadservices.com/pagead/conversion/1070319369/?label=AFWuCNmS9wEQiY6v_gM&amp;guid=ON&amp;script=0""/>" & vbCrLf
	sScript = sScript & "<img height=""1"" width=""1"" style=""border-style:none;"" alt="""" src=""https://www.googleadservices.com/pagead/conversion/1070319369/?label=AFWuCNmS9wEQiY6v_gM&amp;guid=ON&amp;script=0"">" & vbCrLf
	sScript = sScript & "</div>" & vbCrLf
	sScript = sScript & "</noscript>" & vbCrLf

	scrGoogleRemarketing = sScript
End Function
%>
