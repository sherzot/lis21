<%
'*******************************************************************************
'�T�@�v�FAPI�L�[�\�����݊������̃��[�����e
'���@���FvAPIKey�FAPI�L�[
'�@�@�@�FrSbj	�F[OUTPUT]���[������
'�@�@�@�FrBdy	�F[OUTPUT]���[���{��
'�߂�l�FBoolean
'���@�l�F
'���@���F2012/02/13 LIS K.Kokubo �쐬
'*******************************************************************************
Function setMail_APIUser(ByVal vAPIKey,ByRef rSbj,ByRef rBdy)
	Dim sSbj,sBdy

	setMail_APIUser = False

	sSbj = "�������ƃi�r�������ƃi�r���l����API��API�L�[�\�����ݓ��e�m�F"

	sBdy = ""
	sBdy = sBdy & "�����ƃi�r���l����API�̂����p���肪�Ƃ��������܂��B" & vbCrLf
	sBdy = sBdy & "�����ƃi�r�T�|�[�g�ł��B" & vbCrLf & vbCrLf

	sBdy = sBdy & "API�L�[�𔭍s���܂����̂ŁA���LURL���N���b�N���Ċm�肵�Ă��������B" & vbCrLf & vbCrLf

	sBdy = sBdy & HTTPS_CURRENTURL & "api/approval/?key=" & vAPIKey & vbCrLf & vbCrLf

	sBdy = sBdy & "����������������������������������������������������������������������" & vbCrLf
	sBdy = sBdy & "�����ׂĂ��Ȃ���u�����ƃi�r�v�i���X������Ёj" & vbCrLf
	sBdy = sBdy & HTTP_CURRENTURL & vbCrLf
	sBdy = sBdy & "���₢���킹�Flis@lis21.co.jp" & vbCrLf

	rSbj = sSbj
	rBdy = sBdy
End Function
%>
