<%
'*******************************************************************************
'�T�@�v�F�����ƃi�rTOP�̃J���^�������Ŏg�p����Ζ��`�Ԉꗗ��<option></option>���擾
'���@���FvCode			�F�`�F�b�N���̃R�[�h
'�@�@�@�FvAttribute		�Foption�̒ǉ�����
'�߂�l�FString
'���@�l�F
'���@���F2011/02/04 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlWorkingTypeOption_NaviTop(ByVal vCode, ByVal vAttribute)
	Dim sHTML
	Dim aCode(9)

	sHTML = ""

	Select Case vCode
		Case "001": aCode(1) = " selected"
		Case "002": aCode(2) = " selected"
		Case "003": aCode(3) = " selected"
		Case "004": aCode(4) = " selected"
		Case "005": aCode(5) = " selected"
		Case "006": aCode(6) = " selected"
		Case "007": aCode(7) = " selected"
		Case "009": aCode(8) = " selected"
		Case "100": aCode(9) = " selected"
	End Select

	If vAttribute <> "" Then vAttribute = " " & vAttribute

	sHTML = sHTML & "<option value=""001""" & aCode(1) & ">�h��</option>"
	sHTML = sHTML & "<option value=""002""" & aCode(2) & ">���Ј�</option>"
	sHTML = sHTML & "<option value=""003""" & aCode(3) & ">�_��Ј�</option>"
	sHTML = sHTML & "<option value=""004""" & aCode(4) & ">�Љ�\��h��</option>"
	sHTML = sHTML & "<option value=""005""" & aCode(5) & ">�p�[�g�E�A���o�C�g</option>"
	sHTML = sHTML & "<option value=""006""" & aCode(6) & ">SOHO(�ݑ�E����)</option>"
	sHTML = sHTML & "<option value=""007""" & aCode(7) & ">FC�E�㗝�X</option>"
	sHTML = sHTML & "<option value=""009""" & aCode(8) & ">�o�c�ҁE����</option>"
	sHTML = sHTML & "<option value=""100""" & aCode(9) & ">�V��</option>"

	htmlWorkingTypeOption_NaviTop = sHTML
End Function
%>
