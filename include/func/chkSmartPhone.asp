<%
'*******************************************************************************
'�T�@�v�F�X�}�[�g�t�H���̃A�N�Z�X���ǂ������`�F�b�N
'���@���F
'�߂�l�FBoolean
'���@�l�F
'���@���F2011/06/30 LIS K.Kokubo �쐬
'*******************************************************************************
Function chkSmartPhone(ByVal vUserAgent)
	Dim oRE
	Dim sPattern

	sPattern = "((iPhone)|(Android)|(BlackBerry)|(Symbian))"

	Set oRE = New RegExp
	oRE.IgnoreCase = True
	oRE.Pattern = sPattern

	chkSmartPhone = oRE.Test(vUserAgent)
End Function
%>
