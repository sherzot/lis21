<%
	Dim sUserAgent : sUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
	if InStr(sUserAgent,"DoCoMo") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	Elseif InStr(sUserAgent,"KDDI") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	Elseif InStr(sUserAgent,"au") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	Elseif InStr(sUserAgent,"J-PHONE") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	Elseif InStr(sUserAgent,"Vodafone") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	Elseif InStr(sUserAgent,"SoftBank") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	Elseif InStr(sUserAgent,"UP.Browser") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	'Elseif InStr(sUserAgent,"WILLCOM") > 0 then
	Elseif InStr(sUserAgent,"DDIPOCKET") > 0 or InStr(sUserAgent,"WILLCOM") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	Elseif Left(sUserAgent,8) = "MOT-C980" or Left(sUserAgent,8) = "MOT-V980" then
		Response.Redirect "http://m.shigotonavi.jp/"
	Elseif InStr(sUserAgent,"Googlebot-Mobile") > 0 then
		Response.Redirect "http://m.shigotonavi.jp/"
	End If
%>