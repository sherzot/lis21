<%
'*******************************************************************************
'�T�@�v�F�N����̃A�h�o�C�X�g�s�l�k���擾
'���@���FvAgeMin	�F�N���
'�@�@�@�FvAgeMax	�F�N����
'�߂�l�FString
'���@�l�F
'���@���F2010/11/12 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlAgeAdvice(ByVal vAgeMin,ByVal vAgeMax)
	Dim sHTML

	sHTML = sHTML & "<p style=""margin-bottom:5px;"">"
	sHTML = sHTML & "�y�N������n�j�ȏꍇ�̃A�h�o�C�X�z<br>"
	sHTML = sHTML & "�E<span style=""color:#ff0000;"">���Ј���W</span>���A<span style=""color:#ff0000;"">��N�܂ł̔N����W</span>���Ă���B�i�����̋L�ڂ́~�j<br>"
	sHTML = sHTML & "�@�����R�၄�U�O�Ζ����̕����W�i��N���U�O�΁j<br>"
	sHTML = sHTML & "�E<span style=""color:#ff0000;"">���Ј���W</span>���A<span style=""color:#ff0000;"">�E�ƌo���s��</span>���A<span style=""color:#ff0000;"">�V�K�w���҂Ɠ����̏���</span>�ł��邱�ƁB�i�����̋L�ڂ́~�j<br>"
	sHTML = sHTML & "�@�����R�၄�R�T�Ζ����̕����W�i�E���o���s��j<br>"
	sHTML = sHTML & "����L�̃P�[�X�ȊO�̏ꍇ�́A�N������F�߂��Ȃ��Ƌ^���������ǂ��ł��B<br>"
	sHTML = sHTML & "</p>"
	sHTML = sHTML & "<p style=""margin-bottom:5px;"">"
	sHTML = sHTML & "�y�ǂ�����_���ȃP�[�X�z<br>"
	sHTML = sHTML & "�E<span style=""color:#ff0000;"">�L���J���_��i�_��Ј��Ȃǁj</span>�̏ꍇ�͂قƂ�ǂ̏ꍇ�N������ł��܂���B<br>"
	If vAgeMin <> "" And vAgeMax <> "" Then
		sHTML = sHTML & "�E<span style=""color:#ff0000;"">�N��̏���E��������</span>�̋L�ڂ�����A<span style=""color:#ff0000;"">30�΁`49�΂̊Ԃɔ[�܂��Ă��Ȃ�</span>�ꍇ�A�قƂ�ǂ̃P�[�X�łł��܂���B<br>"
		sHTML = sHTML & "�@��30�΁`49�΂ɔ[�܂��Ă��Ă��u�R���̃��v�ɓK�����Ă��Ȃ��ꍇ�̓A�E�g�ł��B<br>"
	End If
	If vAgeMin <> "" Then
		sHTML = sHTML & "�E<span style=""color:#ff0000;"">�N��̉���</span>������ꍇ�́A�قƂ�ǂ̃P�[�X�łł��܂���B<br>"
		sHTML = sHTML & "�@����O...�J����@�Ȃǂɂ��N���������d���̏ꍇ�i�P�W�Έȏ�Ȃǁj<br>"
	End If
	If vAgeMax <> "" Then
		sHTML = sHTML & "�E<span style=""color:#ff0000;"">�N��̏��</span>������ꍇ��<span style=""color:#ff0000;"">�o���҂�D��</span>���Ă���ꍇ�́A�قƂ�ǂ̃P�[�X�łł��܂���B<br>"
		sHTML = sHTML & "�@��<span style=""color:#ff0000;"">�����o�����K�v�Ȏ��i</span>�̏��L�҂�D�����Ă���ꍇ���_���ł��B<br>"
		sHTML = sHTML & "�@����O...���Ј���W���A�E�ƌo���s�₩�A�V�K�w���҂Ɠ����̏����ł��邱�ƁB<br>"
	End If
	sHTML = sHTML & "</p>"

	htmlAgeAdvice = sHTML
End Function
%>
