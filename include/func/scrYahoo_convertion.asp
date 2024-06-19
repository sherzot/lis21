<%
'*******************************************************************************
'概　要：Yahoo!リスティングコンバージョンタグ取得
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2011/10/19 LIS K.Kokubo 作成
'*******************************************************************************
Function scrYahoo_convertion()
	Dim sScript

	sScript = sScript & "<script type=""text/javascript"">" & vbCrLf
	sScript = sScript & "/* <![CDATA[ */" & vbCrLf
	sScript = sScript & "var yahoo_conversion_id = 1000012858;" & vbCrLf
	sScript = sScript & "var yahoo_conversion_label = ""PnbpCL6UswIQiuyM5AM"";" & vbCrLf
	sScript = sScript & "var yahoo_conversion_value = 0;" & vbCrLf
	sScript = sScript & "/* ]]> */" & vbCrLf
	sScript = sScript & "</script>" & vbCrLf
	sScript = sScript & "<script type=""text/javascript"" src=""https://s.yimg.jp/images/listing/tool/cv/conversion.js"">" & vbCrLf
	sScript = sScript & "</script>" & vbCrLf
	sScript = sScript & "<noscript>" & vbCrLf
	sScript = sScript & "<div style=""display:inline;"">" & vbCrLf
	sScript = sScript & "<img height=""1"" width=""1"" style=""border-style:none;"" alt="""" src=""https://b91.yahoo.co.jp/pagead/conversion/1000012858/?label=PnbpCL6UswIQiuyM5AM&amp;guid=ON&amp;script=0&amp;disvt=true""/>" & vbCrLf
	sScript = sScript & "</div>" & vbCrLf
	sScript = sScript & "</noscript>" & vbCrLf

	scrYahoo_convertion = sScript
End Function
%>
