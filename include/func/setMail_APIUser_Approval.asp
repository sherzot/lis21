<%
'*******************************************************************************
'�T�@�v�FAPI�L�[���F�������̃��[�����e
'���@���FvAPIKey�FAPI�L�[
'�@�@�@�FrSbj	�F[OUTPUT]���[������
'�@�@�@�FrBdy	�F[OUTPUT]���[���{��
'�߂�l�FBoolean
'���@�l�F
'���@���F2012/02/13 LIS K.Kokubo �쐬
'*******************************************************************************
Function setMail_APIUser_Approval(ByVal vAPIKey,ByRef rSbj,ByRef rBdy)
	Dim sSbj,sBdy

	setMail_APIUser_Approval = False

	sSbj = "�������ƃi�r�������ƃi�r���l����API��API�L�[���F����"

	sBdy = ""
	sBdy = sBdy & "�����ƃi�r���l����API�̂����p���肪�Ƃ��������܂��B" & vbCrLf
	sBdy = sBdy & "�����ƃi�r�T�|�[�g�ł��B" & vbCrLf & vbCrLf

	sBdy = sBdy & "API�L�[�̏��F���������܂����B" & vbCrLf
	sBdy = sBdy & "���Ȃ���API�L�[�͈ȉ��̒ʂ�ł��B" & vbCrLf & vbCrLf

	sBdy = sBdy & vAPIKey & vbCrLf & vbCrLf

	sBdy = sBdy & "���̃��[���͑�؂ɕۊǂ��Ă��������B" & vbCrLf & vbCrLf

	sBdy = sBdy & "����������������������������������������������������������������������" & vbCrLf
	sBdy = sBdy & "�����ׂĂ��Ȃ���u�����ƃi�r�v�i���X������Ёj" & vbCrLf
	sBdy = sBdy & HTTP_CURRENTURL & vbCrLf
	sBdy = sBdy & "���₢���킹�Flis@lis21.co.jp" & vbCrLf

	rSbj = sSbj
	rBdy = sBdy

	setMail_APIUser_Approval = True
End Function
%>
