<%
'*******************************************************************************
'�T�@�v�F�^�uIndex���擾
'���@���F
'�߂�l�FString
'���@�l�F
'���@���F2011/03/16 LIS K.Kokubo �쐬
'*******************************************************************************
Function getHeaderText(ByVal vHeadType,ByVal vURL)
	Dim sText

	vURL = LCase(vURL)

	If vHeadType = 0 Then '�g�b�v
		sText = "�]�E�T�C�g�u�����ƃi�r�v�B���Ј��E�h���̋��l���͂������A�v���ɂ��M���ɓK�����]�E�T�|�[�g�����񋟂��Ă��܂��I"
	ElseIf vHeadType = 1 Then '���E��
		sText = "�]�E�����E���E�����̕��X�ɍœK�ȋ��l���Ɨ������c�[����񋟂��Ă��܂�"
	ElseIf vHeadType = 2 Or vHeadType = 4 Then '���
		sText = "��Ƃ̐l�ތٗp�𕝍L���T�|�[�g���Ă���܂��B�i���l�L���A�l�ޔh���A�l�ޏЉ�j"
	ElseIf vHeadType = 3 Then '���p
		sText = "���E�����̕��X�ɍœK�ȋ��l���Ɨ������c�[����񋟂��Ă��܂�"
	End If

	If InStr(vURL,"/staff/s_careersheet.asp") > 0 Then
		sText = "�E���o�����쐬�E����b�����ƃi�r�̊ȒP�֗��ȐE���o�����̎����쐬�E���"
	ElseIf InStr(vURL,"/order/special/ad/0001/") > 0 Then
		sText = "SE�]�E���W�bSE�̓]�E�E���l����]�E�x���T�[�r�X���s���Ă��܂��B"
	ElseIf InStr(vURL,"/order/special/tg/0004/") > 0 Then
		sText = "�Տ������Z�t ���l���W�b�Տ������Z�t�̎��i�������������l�̂��Љ�ƁA�����̓]�E�x���T�[�r�X�̂��ē��B"
	ElseIf InStr(vURL,"/order/special/tg/0005/") > 0 Then
		sText = "�p�����������h�����W�b�p��ɓ��������h���X�^�b�t����̓o�^���S���ōs���Ă��܂��B"
	ElseIf InStr(vURL,"/order/special/sz/0001/") > 0 Then
		sText = "�É� �]�E���W�b�n��ɓ����������l�̒񋟂ƁA�����̓]�E�x���T�[�r�X���s���Ă��܂��B"
	ElseIf InStr(vURL,"/order/special/ng/0002/") > 0 Then
		sText = "���É� �h�����W�b�n��ɓ��������h���X�^�b�t����̓o�^���S���ōs���Ă��܂��B"
	ElseIf InStr(vURL,"/order/special/or/0001/") > 0 Then
		sText = "DTP ���l���W�bDTP�I�y���[�^,�f�U�C�i�[�ɓ����������l�ƁA�����̓]�E�x���T�[�r�X�̂��ē��B"
	ElseIf InStr(vURL,"/order/special/oy/0001/") > 0 Then
		sText = "���R ���l���W�b�n��ɓ����������l�̒񋟂ƁA�����̓]�E�x���T�[�r�X���s���Ă��܂��B"
	ElseIf InStr(vURL,"/order/special/hr/0001/") > 0 Then
		sText = "�L�� �]�E���W�b�n��ɓ����������l�̒񋟂ƁA�����̓]�E�x���T�[�r�X���s���Ă��܂��B"
	End If

	If vHeadType <> 0 Then
		sText = "<strong style=""color:#666666;"">" & sText & "</strong>"
	Else
		sText = "<b style=""color:#666666;"">" & sText & "</b>"
	End If

	getHeaderText = sText
End Function
%>
