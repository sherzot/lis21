<%
'*******************************************************************************
'�T�@�v�FURLEncode��������f�R�[�h
'���@���FvText		�F�f�R�[�h�Ώە�����
'�@�@�@�FvCharset	�F�L�����N�^�[�Z�b�g ["sjis"]["utf8"]["euc"]["jis"]
'�o�@�́F
'�߂�l�FString		�F�f�R�[�h���ꂽ������
'���@�l�F
'�X�@�V�F2009/08/06 LIS K.Kokubo �쐬
'*******************************************************************************
Function getURLDecode(ByVal vText, ByVal vCharSet)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	sSQL = "SELECT dbo.func_GetCharactor('" & vText & "','sjis') AS Str;"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
	If GetRSState(oRS) = True Then
		getURLDecode = oRS.Collect("Str")
	End If
	Call RSClose(oRS)
End Function
%>
