<%
'*******************************************************************************
'�T�@�v�F�L�����A�J�E���Z���[�֑��k�̓��e�m�F���[���̕����𐶐�
'���@���FvMailType		�F���M�惁�[����� [1]PC[2]MOBILE
'�@�@�@�FvName			�F���E�Җ�
'�@�@�@�FvTitle			�F���k�^�C�g��
'�@�@�@�FvBody			�F���k�{��
'�@�@�@�FrSubject		�F[OUTPUT]���[������
'�@�@�@�FrBody			�F[OUTPUT]���[���{��
'�߂�l�FBoolean
'���@�l�F
'���@���F2011/11/29 LIS K.Kokubo �쐬
'*******************************************************************************
Function setMail_FBConsul(ByVal vMailType,ByVal vName,ByVal vTitle,ByVal vBody,ByRef rSubject,ByRef rBody)
	Dim sSubject,sBody

	setMail_FBConsul = False

	If CStr(vMailType) = "1" Then
		sSubject = "�������ƃi�r���L�����A�J�E���Z���[�ւ̑��k���󂯕t���܂���"
		sBody = ""
		sBody = sBody & vName & "�@�l" & vbCrLf & vbCrLf

		sBody = sBody & "�����p���肪�Ƃ��������܂��B" & vbCrLf
		sBody = sBody & "�����ƃi�r�^�c�����ǂł��B" & vbCrLf & vbCrLf

		sBody = sBody & "�L�����A�J�E���Z���[�ւ̑��k���󂯕t���܂����̂ł��m�点�v���܂��B" & vbCrLf
		sBody = sBody & "�����k�𒸂������e�͉��L�̒ʂ�ł��B" & vbCrLf & vbCrLf

		sBody = sBody & "----------------------------------------------------------------------" & vbCrLf
		sBody = sBody & "�y���k�^�C�g���z" & vbCrLf
		sBody = sBody & vTitle & vbCrLf & vbCrLf
		sBody = sBody & "�y�{���z" & vbCrLf
		sBody = sBody & vBody & vbCrLf
		sBody = sBody & "----------------------------------------------------------------------" & vbCrLf & vbCrLf

		sBody = sBody & "����������������������������������������������������������������������" & vbCrLf
		sBody = sBody & "�����ׂĂ��Ȃ���u�����ƃi�r�v�i���X������Ёj" & vbCrLf
		sBody = sBody & HTTP_CURRENTURL & vbCrLf
		sBody = sBody & "�������ƃi�rFacebook�y�[�W" & vbCrLf
		sBody = sBody & HTTP_FB & vbCrLf
		sBody = sBody & "���₢���킹�Flis@lis21.co.jp" & vbCrLf

		setMail_FBConsul = True
	ElseIf CStr(vMailType) = "2" Then
		sSubject = "����������ށ���ر��ݾװ�ւ̑��k���󂯕t���܂���"
		sBody = ""
		sBody = sBody & vName & "�@�l" & vbCrLf & vbCrLf

		sBody = sBody & "�����p���肪�Ƃ��������܂��B" & vbCrLf
		sBody = sBody & "��������މ^�c�����ǂł��B" & vbCrLf & vbCrLf

		sBody = sBody & "��ر��ݾװ�ւ̑��k���󂯕t���܂����̂ł��m�点�v���܂��B" & vbCrLf
		sBody = sBody & "�����k�𒸂������e�͉��L�̒ʂ�ł��B" & vbCrLf & vbCrLf

		sBody = sBody & "------------------------------" & vbCrLf
		sBody = sBody & "�y���k���فz" & vbCrLf
		sBody = sBody & vTitle & vbCrLf & vbCrLf
		sBody = sBody & "�y�{���z" & vbCrLf
		sBody = sBody & vBody & vbCrLf
		sBody = sBody & "------------------------------" & vbCrLf & vbCrLf

		sBody = sBody & "������������������������������" & vbCrLf
		sBody = sBody & "�����������(ؽ�������)" & vbCrLf
		sBody = sBody & HTTP_CURRENTURL & vbCrLf
		sBody = sBody & "�����������Facebook�߰��" & vbCrLf
		sBody = sBody & HTTP_FB & vbCrLf
		sBody = sBody & "���₢���킹�Flis@lis21.co.jp" & vbCrLf

		setMail_FBConsul = True
	End If

	rSubject = sSubject
	rBody = sBody
End Function
%>
