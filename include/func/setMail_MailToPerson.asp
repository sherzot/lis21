<%
'*******************************************************************************
'�T�@�v�F��Ƃ���̃��[�����M�ʒm���[���̕����𐶐�
'���@���FvMailType		�F���M�惁�[����� [pc][mobile]
'�@�@�@�FvSubject		�F���[�������ɓo�^��������
'�@�@�@�FvBody			�F���[�������ɓo�^�����{��
'�@�@�@�FvStaffCode		�F���M�拁�E�҃R�[�h
'�@�@�@�FvStaffName		�F���M�拁�E�Җ�
'�@�@�@�FvCompanyName	�F���M����Ɩ�
'�@�@�@�FvOrderCode		�F���[���ɕR�Â������R�[�h
'�@�@�@�FvJobTypeDetail	�F���[���ɕR�Â������l�̋�̓I�E�햼
'�@�@�@�FrSubject		�F[OUTPUT]����
'�@�@�@�FrBody			�F[OUTPUT]�{��
'�߂�l�FBoolean
'���@�l�F
'�X�@�V�F2009/07/02 LIS K.Kokubo �쐬
'*******************************************************************************
Function setMail_MailToPerson(ByVal vMailType, ByVal vSubject, ByVal vBody, ByVal vStaffCode, ByVal vStaffName, ByVal vCompanyName, ByVal vOrderCode, ByVal vJobTypeDetail, ByRef rSubject, ByRef rBody)
	setMail_MailToPerson = False

	If vMailType = "pc" Then
		rSubject = GetMailSubject(vJobTypeDetail)
		rBody = GetMailBodyCompany(vOrderCode, vSubject, vBody, vStaffCode, vStaffName, vCompanyName, vJobTypeDetail, "1")

		setMail_MailToPerson = True
	ElseIf vMailType = "mobile" Then
		rSubject = "[�����ƃi�r]��Ƃ���Ұْ��M"

		rBody = ""
		rBody = rBody & "���������p���肪�Ƃ��������܂��B" & vbCrLf
		rBody = rBody & "�������l���E��Ģ�����ƃi�r�(���X�������)�ł��B" & vbCrLf
		rBody = rBody & "������ƃi�r���o�C�����ʂ��āA���l��Ƃ���M����Ұق��͂��܂����B" & vbCrLf
		rBody = rBody & "������ƃi�r���۸޲݂��āAҰٓ��e�����m�F�������B" & vbCrLf & vbCrLf
		rBody = rBody & "���g�є�(��޲�)�̏ꍇ:������ƃi�r���o�C���TOP����۸޲݁�My�߰�ނ̢Ұٗ����ݸ��د�" & vbCrLf
		rBody = rBody & HTTP_NAVI_MOBILE & vbCrLf
		rBody = rBody & "��PC�ł̏ꍇ:������ƃi�r�TOP����۸޲݁�۸޲��ƭ��̢Ұٗ����ݸ��د�" & vbCrLf
		rBody = rBody & "-----------------"  & vbCrLf
		rBody = rBody & "������Ұق͎������MҰق̂��߁A�ԐM�ł��܂���B�����Ӊ������B" & vbCrLf
		rBody = rBody & "-----------------"  & vbCrLf
		rBody = rBody & "���X�������" & MAIL_LIS & vbCrLf
		rBody = rBody & "�����ƃi�r���o�C���F" & HTTP_NAVI_MOBILE & vbCrLf
		rBody = rBody & "�����ƃi�r�F" & HTTP_CURRENTURL

		setMail_MailToPerson = True
	End If
End Function
%>
