<%
'*******************************************************************************
'�T�@�v�F�O���T�[�o��XML�h�L�������g���擾
'���@���FvURL	�FXML�h�L�������g��URL
'�o�@�́F
'�߂�l�FObject (MSXML2.DOMDocument)
'���@�l�F
'�X�@�V�F2009/01/09 LIS K.Kokubo �쐬
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
