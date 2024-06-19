<%
'*******************************************************************************
'概　要：外部サーバのXMLドキュメントを取得
'引　数：vURL	：XMLドキュメントのURL
'出　力：
'戻り値：Object (MSXML2.DOMDocument)
'備　考：
'更　新：2009/01/09 LIS K.Kokubo 作成
'*******************************************************************************
Function getXMLDocument(ByRef rXML, ByVal vURL)
	Dim childLength
	Dim iRet

	Set rXML = Server.CreateObject("MSXML2.DOMDocument")
	rXML.async = False
	rXML.setProperty "ServerHTTPRequest", True

	getXMLDocument = rXML.Load(vURL)
End Function
%>
