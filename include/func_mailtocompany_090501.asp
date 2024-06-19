<%
'******************************************************************************
'�T�@�v�F���[���̏������擾
'��@�ҁF2007/06/25 Lis K.Kokubo
'���@���FvType				�F���[�����M���� ["1"]��ʂ̋��l�L�� ["2"]���X�Č� ["3"]���l�[���R�t���Ă��Ȃ����X�Č�
'�@�@�@�FvStaffCode			�F���[�����M�����E�҃R�[�h
'�@�@�@�FvCompanyName		�F���[�����M���Ɩ�
'�@�@�@�FvOrderCode			�F���[���ɕR�Â������R�[�h
'�@�@�@�FvContactPersonName	�F���[�����M��Č��S����
'�@�@�@�FvSubject			�F���[������
'�@�@�@�FvBody				�F���[�����e
'�@�@�@�FvMailURL			�F���E�ҏ��Q�Ɛ�t�q�k
'�@�@�@�FvHeader			�F���[�����e�ɒǉ�����w�b�_
'�@�@�@�FvFooter			�F���[�����e�ɒǉ�����t�b�^
'�߂�l�F
'���@�l�F
'�g�p���F�����ƃi�r/staff/mailtocompany.asp
'�X�@�V�F
'******************************************************************************
Function GetMailBodyStaff(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter)
	Dim sBody
	Dim iLen

	sBody = ""
	Select Case vType
		Case "1":
			'��ʂ̋��l�L��
			'��Ɩ��{���l�[�S���Җ��{�l
			sBody = vCompanyname & "�@" & vContactPersonName & "�l" & vbCrLf & vbCrLf
		Case "2":
			'���X�Č�
			'���X�Ј����{�l
			sBody = vContactPersonName & "�l" & vbCrLf & vbCrLf
			vMailURL = "http://bi.lis21.co.jp/staff/staff_MailHistory.asp?Mail=pop&newopen=1"
		Case Else:
			'���l�[���R�t���Ă��Ȃ�
			If Left(vCompanyCode, 1) = "L" Then
				vMailURL = "http://bi.lis21.co.jp/staff/staff_MailHistory.asp?Mail=pop&newopen=1"
			End If
	End Select
	sBody = sBody & vHeader & vbCrLf & vMailURL & vbCrLf

	'���[�����e�\������
	iLen = Len(vBody) * 0.3
	sBody = sBody & vbCrLf & vbCrLf & _
		"-----------------------���@���[�����@��-------------------------" & vbCrLf & _
		"�y���M�҃R�[�h�z" & vStaffCode & vbCrLf & _
		"�y�Ώۏ��R�[�h�z" & vOrderCode & vbCrLf & _
		"�y���[���^�C�g���z" & vSubject & vbCrLf & _
		"�y���[�����e�z" & vbCrLf & vbCrLf & Left(vBody, iLen) & "..." & vbCrLf & _
		"------------------------------------------------------------------" & vbCrLf
	sBody = sBody & vFooter

	GetMailBodyStaff = sBody
End Function

'******************************************************************************
'�T�@�v�F���[���̏������擾
'��@�ҁF2007/06/25 Lis K.Kokubo
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserID	�F���O�C�����̃��[�U���
'�߂�l�F
'���@�l�F
'�g�p���F�����ƃi�r/staff/mailtoperson.asp
'�X�@�V�F
'******************************************************************************
Function GetMailSignatureStaff(ByRef rDB, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	sSQL = "sp_GetDataMailSignatureStaff '" & vUserID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	sSignature = ""
	If GetRSState(oRS) = True Then
		sSignature = sSignature & "------------------------------" & vbCrLf
		sSignature = sSignature & "�Z���F" & oRS.Collect("Prefecture") & oRS.Collect("City") & oRS.Collect("Town") & oRS.Collect("Address") & vbCrLf
		sSignature = sSignature & "�����F" & oRS.Collect("Name") & vbCrLf
		If oRS.Collect("HomeContactFlag") = "1" Then
			sSignature = sSignature & "����F" & oRS.Collect("HomeTelephoneNumber") & vbCrLf
		End If
		If oRS.Collect("PortableContactFlag") = "1" Then
			sSignature = sSignature & "�g�сF" & oRS.Collect("PortableTelephoneNumber") & vbCrLf
		End If
		If oRS.Collect("FaxContactFlag") = "1" Then
			sSignature = sSignature & "FAX �F" & oRS.Collect("FaxNumber") & vbCrLf
		End If
		If oRS.Collect("MailContactFlag") = "1" Then
			sSignature = sSignature & "Mail�F" & oRS.Collect("MailAddress") & vbCrLf
		End If
	End If
	Call RSClose(oRS)

	GetMailSignatureStaff = sSignature
End Function

'******************************************************************************
'�T�@�v�F���E�҂̃��[���e���v���[�g�擾
'��@�ҁF2007/06/25 Lis K.Kokubo
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserType	�F���O�C�����̃��[�U���
'�@�@�@�FvSEQ		�F�ԍ�
'�߂�l�F
'���@�l�F
'�g�p���F�����ƃi�r/staff/mailtocompany.asp
'�X�@�V�F
'******************************************************************************
Function GetStaffMailTemplateOptionHtml(ByRef rDB, ByVal vUserType, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sSelected

	GetStaffMailTemplateOptionHtml = ""

	sSQL = "sp_GetDataMailTemplate '1'"

	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sSelected = ""
		If CStr(oRS.Collect("Cd")) = CStr(vSEQ) Then sSelected = "selected=""true"""
		GetStaffMailTemplateOptionHtml = GetStaffMailTemplateOptionHtml & _
			"<option value=""" & oRS.Collect("Cd") & """ " & sSelected & ">" & oRS.Collect("Title") & "</option>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function
%>
