<%
'*******************************************************************************
'�T�@�v�F��Ƃ���̃��[�����M�ʒm���[���̕����𐶐�
'���@���FvCompanyName	�F�ʒm���Ɩ�
'�@�@�@�FvPersonName	�F�ʒm���Ƌ��l�S����
'�@�@�@�FvLinkOrder		�F�Ώۋ��l�[�ւ̃����N
'�@�@�@�FvLinkProfile	�F�}�b�`���O�l�ނ̃v���t�B�[���ւ̃����N�Q
'�@�@�@�FrSubject		�F[OUTPUT]����
'�@�@�@�FrBody			�F[OUTPUT]�{��
'�߂�l�FBoolean
'���@�l�F
'�X�@�V�F2009/08/14 LIS K.Kokubo �쐬
'*******************************************************************************
Function setMail_SpMchNotice(ByVal vCompanyName, ByVal vPersonName, ByVal vLinkOrder, ByVal vLinkProfile, ByRef rSubject, ByRef rBody)
	setMail_SpMchNotice = False

	rSubject = ""
	rSubject = rSubject & "�������ƃi�r���K�ޏ����Ƀ}�b�`�������E�҂̂��m�点"

	rBody = ""
	rBody = rBody & ""
	rBody = rBody & vCompanyName & vbCrLf
	rBody = rBody & vPersonName & "�l" & vbCrLf
	rBody = rBody & vbCrLf

���������p���肪�Ƃ��������܂��B
�������l���E�T�C�g�u�����ƃi�r�v�i���X������Ёj�ł��B

���l�̓K�ޏ����Ƀ}�b�`�������E�҂�����܂����̂ł��A���������܂��B

End Function
%>
