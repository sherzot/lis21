<%
'**********************************************************************************************************************
'�T�@�v�F��Ƒ����[���쐬��� �����ƃi�r/company/mailtoperson.asp
'�@�@�@�F��L�y�[�W�ŏo�͗p�̊֐��Q�����̃t�@�C���ɗp�ӂ���B
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/config/constant.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���[���@������
'�@�@�@�FGetMailSubject					�F���[���̌����𐶐����Ď擾
'�@�@�@�FGetMailBodyCompany				�F���[���̓��e�𐶐����Ď擾
'�@�@�@�FGetMailSignature				�F���[���̏������擾
'�@�@�@�FGetNaviMailTemplateOptionHtml	�F�i�r���[���e���v���[�g�擾
'�@�@�@�FGetMailTemplateOptionHtml		�F���l�[���[���e���v���[�g�擾
'�@�@�@�FRegMail						�F���[�����c�a�ɓo�^����
'�@�@�@�FMailToPerson					�F���[���𑗐M����
'**********************************************************************************************************************
%>
<!-- #INCLUDE VIRTUAL="/include/func/setMail_MailToPerson.asp" -->
<%
'******************************************************************************
'�T�@�v�F���[���̌����𐶐����Ď擾
'��@�ҁF2007/06/22 Lis K.Kokubo
'���@���FvJobTypeDetail	�F��̓I�E�햼
'�߂�l�F
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'�X�@�V�F
'******************************************************************************
Function GetMailSubject(ByVal vJobTypeDetail)
	If Len(vJobTypeDetail) <= 15 Then
		GetMailSubject = "�������ƃi�r���X�J�E�g�E�A�����[�����M�̂��m�点�I�^�u" & Left(vJobTypeDetail,15) & "�v�̂��d��"
	Else
		GetMailSubject = "�������ƃi�r���X�J�E�g�E�A�����[�����M�̂��m�点�I�^�u" & Left(vJobTypeDetail,15) & "...�v�̂��d��"
	End If
End Function

'******************************************************************************
'�T�@�v�F���[���̓��e�𐶐����Ď擾
'���@���FvOrderCode		�F���[���ɕt�����鋁�l�[�̏��R�[�h
'�@�@�@�FvSubject		�F���[���̌���
'�@�@�@�FvBody			�F���[���̓��e
'�@�@�@�FvStaffCode		�F���[����M�����E�҃R�[�h
'�@�@�@�FvStaffName		�F���[����M�����E�Җ�
'�@�@�@�FvCompanyName	�F���[�����M����Ɩ�
'�@�@�@�FvJobTypeDetail	�F��̓I�E�햼
'�߂�l�F
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'���@���F2007/06/22 LIS K.Kokubo �쐬
'�@�@�@�F2009/07/02 LIS K.Kokubo �X�J�E�g���[�����ǐ��\�����폜�B���ǒʒm���[�����o���Ă��邽�߂ɕs�v�B
'******************************************************************************
Function GetMailBodyCompany(ByVal vOrderCode, ByVal vSubject, ByVal vBody, ByVal vStaffCode, ByVal vStaffName, ByVal vCompanyName, ByVal vJobTypeDetail, ByVal vType)
	Dim sBody
	Dim iLen
	Dim idx

	GetMailBodyCompany = ""

	'�{��
	sBody = ""
	If vStaffName <> "" Then sBody = vStaffName & "�@�l"  & vbCrLf & vbCrLf
	sBody = sBody & MAIL_FROM_COMPANY_BODY & vbCrLf
	sBody = sBody & MAIL_URL_STAFF & "?si=" & vStaffCode & vbCrLf

	'���[�����e�\������
	iLen = Len(vBody) * 0.3
	sBody = sBody & vbCrLf & _
		"----------------------���@�ŐV�z�M���@��-----------------------" & vbCrLf & _
		"�y��Ж��z�@"

	If vType = "2" Then
			sBody = sBody & "���X�������"
	ElseIf vType = "1" Then
			sBody = sBody & vCompanyName	'CC_CompanyName_K
	Else
			sBody = sBody & vCompanyName
	End If

	sBody = sBody & vbCrLf & _
		"�y�d�����e�z�@" & vJobTypeDetail & "(" & vOrderCode & ")" & vbCrLf & _
		"-----------------------���@���[�����@��-------------------------" & vbCrLf & _
		"�y���[���^�C�g���z�@" & vSubject & vbCrLf & _
		"�y���[�����e�z" & vbCrLf & Left(vBody, iLen) & "..." & vbCrLf & _
		"----------------------------------------------------------------" & vbCrLf

	sBody = sBody & MAIL_FROM_COMPANY_FOOTER

	GetMailBodyCompany = sBody
End Function

'******************************************************************************
'�T�@�v�F���[���̏������擾
'��@�ҁF2007/06/22 Lis K.Kokubo
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserType	�F���O�C�����̃��[�U���
'�@�@�@�FvOrderCode	�F���[���ɕt�����鋁�l�[�̏��R�[�h
'�߂�l�F
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'�X�@�V�F
'******************************************************************************
Function GetMailSignature(ByRef rDB, ByVal vUserType, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	GetMailSignature = ""
	If G_USERTYPE <> "staff" Then
		sSQL = "sp_GetDataMailSignatureCompany '" & vOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

		If GetRSState(oRS) = True Then
			GetMailSignature = GetMailSignature & "------------------------------" & vbCrLf
			GetMailSignature = GetMailSignature & "��Ж��@�@�F" & oRS.Collect("CompanyName") & vbCrLf
			GetMailSignature = GetMailSignature & "�S���ҕ����F" & oRS.Collect("SectionName") & vbCrLf
			GetMailSignature = GetMailSignature & "�d�b�ԍ��@�F" & oRS.Collect("TelephoneNumber") & vbCrLf
			GetMailSignature = GetMailSignature & "�S���Ҏ����F" & oRS.Collect("PersonName") & vbCrLf
			GetMailSignature = GetMailSignature & "�S����Mail�F" & oRS.Collect("MailAddress") & vbCrLf
		End If
		Call RSClose(oRS)
	Else
		'�����p���̎擾
		sSQL = "sp_GetDataMailSignatureStaff '" & G_USERID & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

		If GetRSState(oRS) = True Then
			GetMailSignature = GetMailSignature & "------------------------------" & vbCrLf & _
				"�Z���F" & oRS.Collect("Prefecture") & oRS.Collect("City") & oRS.Collect("Town") & oRS.Collect("Address") & _
				"�����F" & oRS.Collect("Name")

			If oRS.Collect("HomeContactFlag") = "1" Then
				GetMailSignature = GetMailSignature & "����F" & oRS.Collect("HomeTelephoneNumber")
			End If
			If oRS.Collect("PortableContactFlag") = "1" Then
				GetMailSignature = GetMailSignature & "�g�сF" & oRS.Collect("PortableTelephoneNumber")
			End If
			If oRS.Collect("FaxContactFlag") = "1" Then
				GetMailSignature = GetMailSignature & "FAX �F" & oRS.Collect("FaxNumber")
			End If
			If oRS.Collect("MailContactFlag") = "1" Then
				GetMailSignature = GetMailSignature & "Mail�F" & oRS.Collect("MailAddress")
			End If
		End If
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'�T�@�v�F�i�r���[���e���v���[�g�擾
'��@�ҁF2007/06/22 Lis K.Kokubo
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserType	�F���O�C�����̃��[�U���
'�@�@�@�FvNaviSEQ	�F�ԍ�
'�߂�l�F
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'�X�@�V�F
'******************************************************************************
Function GetNaviMailTemplateOptionHtml(ByRef rDB, ByVal vUserType, ByVal vNaviSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sSelected

	GetNaviMailTemplateOptionHtml = ""

	sSQL = "sp_GetDataMailTemplate '2'"

	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sSelected = ""
		If CStr(oRS.Collect("Cd")) = CStr(vNaviSEQ) Then sSelected = "selected=""true"""
		GetNaviMailTemplateOptionHtml = GetNaviMailTemplateOptionHtml & _
			"<option value=""" & oRS.Collect("Cd") & """ " & sSelected & ">" & oRS.Collect("Title") & "</option>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���l�[���[���e���v���[�g�擾
'��@�ҁF2007/06/22 Lis K.Kokubo
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserCode	�F���O�C�����̃��[�U�h�c
'�@�@�@�FvOrderCode	�F���[���ɕt�����鋁�l�[�̏��R�[�h
'�@�@�@�FvSEQ		�F�ԍ�
'�߂�l�F
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'�X�@�V�F
'******************************************************************************
Function GetMailTemplateOptionHtml(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sSelected

	GetMailTemplateOptionHtml = ""

	sSQL = "up_GetListMailTemplate '" & vUserCode & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	Do While GetRSState(oRS) = True
		sSelected = ""
		If CStr(oRS.Collect("SEQ")) = CStr(vSEQ) Then sSelected = "selected=""true"""
		GetMailTemplateOptionHtml = GetMailTemplateOptionHtml & _
			"<option value=""" & oRS.Collect("SEQ") & """ " & sSelected & ">" & oRS.Collect("Subject") & "</option>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���[�����c�a�ɓo�^����
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserCode	�F���O�C�����̃��[�U�h�c
'�@�@�@�FvOrderCode	�F���[���ɕt�����鋁�l�[�̏��R�[�h
'�@�@�@�FvSEQ		�F�ԍ�
'�߂�l�FBoolean	�F[True]���[�����o�^�ł��� [False]���[���̓o�^�ŃG���[����������
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'���@���F2007/06/22 LIS K.Kokubo �쐬
'�@�@�@�F2011/01/05 LIS K.Kokubo Basp.SendMail �� SndMail
'******************************************************************************
Function RegMail(ByRef rDB, ByVal vID, ByVal vUserID, ByVal vReceiverCode, ByVal vSubject, ByVal vBody, ByVal vOrderCode, ByVal vSenderEvaluation, ByVal vSenderRemark, ByVal vReceiverEvaluation, ByVal vReceiverRemark, ByVal vSenderDelFlag, ByVal vReceiverDelFlag, ByVal vAnswerFlag, ByVal vPayFlag)
	On Error Resume Next
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sSessionValue:	sSessionValue = ""
	Dim sFormValue:	sFormValue = ""
	Dim idx
	Dim sMsg

	sSQL = ""
	sSQL = sSQL & "/* �����ƃi�r ���[���o�^ */" & vbCrLf
	sSQL = sSQL & "up_RegMailHistory"
	sSQL = sSQL & " '" & vID & "'"
	sSQL = sSQL & ",'" & vUserID & "'"
	sSQL = sSQL & ",'" & vReceiverCode & "'"
	sSQL = sSQL & ",'" & vSubject & "'"
	sSQL = sSQL & ",'" & vBody & "'"
	sSQL = sSQL & ",'" & vOrderCode & "'"
	sSQL = sSQL & ",'" & vSenderEvaluation & "'"
	sSQL = sSQL & ",'" & vSenderRemark & "'"
	sSQL = sSQL & ",'" & vReceiverEvaluation & "'"
	sSQL = sSQL & ",'" & vReceiverRemark & "'"
	sSQL = sSQL & ",'" & vSenderDelFlag & "'"
	sSQL = sSQL & ",'" & vReceiverDelFlag & "'"
	sSQL = sSQL & ",'" & vAnswerFlag & "'"
	sSQL = sSQL & ",'" & vPayFlag & "';"

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If flgQE = True Then
		RegMail = True

		If G_PLANTYPE = "mail" And vPayFlag = "1" Then
			sSQL = ""
			sSQL = sSQL & "/* �ۋ����[�����M�|�C���g */" & vbCrLf
			sSQL = sSQL & "EXEC up_RegCMPNaviPoint '" & vUserID & "','','003','" & GetDateStr(Date,"") & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		End If
	Else
		Session("err") = "�X�J�E�g���[���̑��M�Ɏ��s���܂����B<br>���萔�ł����{�����m�F�̏�A�ēx�����M�������B"

		sSQL = "EXEC up_Reg_LOG_Error '" & G_USERID & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'" & _
			",'" & ChkSQLStr(Session.SessionID) & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("URL")) & "?" & ChkSQLStr(Request.ServerVariables("QUERY_STRING")) & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("HTTP_REFERER")) & "'" & _
			",'" & ChkSQLStr(sSQL) & "'" & _
			",'" & ChkSQLStr(Err.Source & vbCrLf & Err.Description) & "'"

		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

		For idx = 1 To Session.Contents.Count
			sSessionValue = sSessionValue & "�y" & Session.Contents.Key(idx) & "�z"
			sSessionValue = sSessionValue & Session.Contents(idx) & vbCrLf
		Next

		For idx = 1 To Request.Form.Count
			sFormValue = sFormValue & "�y" & Request.Form.Key(idx) & "�z"
			sFormValue = sFormValue & Request.Form(idx) & vbCrLf
		Next

		sMsg = "UserID     �F" & G_USERID & vbCrLf & _
			"IPAddress  �F" & Request.ServerVariables("REMOTE_ADDR") & vbCrLf & _
			"UserAgent  �F" & Request.ServerVariables("HTTP_USER_AGENT") & vbCrLf & _
			"Referer    �F" & Request.ServerVariables("HTTP_REFERER") & vbCrLf & _
			"Page       �F" & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING") & vbCrLf & _
			"������������������������������������������������������������" & vbCrLf & _
			"Error Page �F" & vbCrLf & _
			Session("errorpagereferer") & vbCrLf & "��" & vbCrLf & _
			Session("errorpage") & vbCrLf & _
			"������������������������������������������������������������" & vbCrLf & _
			"Description�F" & vbCrLf & Err.Description & vbCrLf & _
			"������������������������������������������������������������" & vbCrLf & _
			"Session    �F" & vbCrLf & sSessionValue & vbCrLf & _
			"������������������������������������������������������������" & vbCrLf & _
			"Post       �F" & vbCrLf & sFormValue & vbCrLf

		Call SndMail(Cnt_MailServer, "kisui@lis21.co.jp", "info@shigotonavi.jp", "�y�����ƃi�r �G���[�z", sMsg, "")

		RegMail = False
	End If

	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���[���𑗐M����
'���@���FvMailServer
'�@�@�@�FvFrom
'�@�@�@�FvUserCode
'�@�@�@�FvID
'�@�@�@�FvStaffCode
'�@�@�@�FvOrderCode
'�@�@�@�FvSubject
'�@�@�@�FvBody
'�@�@�@�FvPayFlag
'�߂�l�FBoolean	�F[True]���[�����o�^�ł��� [False]���[���̓o�^�ŃG���[����������
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'���@���F2007/06/22 LIS K.Kokubo
'�@�@�@�F2011/01/05 LIS K.Kokubo Basp.SendMail �� SndMail
'******************************************************************************
Function MailToPerson(ByVal vMailServer, ByVal vFrom, ByVal vUserCode, ByVal vID, ByVal vStaffCode, ByVal vOrderCode, ByVal vSubject, ByVal vBody, ByVal vPayFlag)
	'���E�҂Ɍ����ă��[���𑗐M����ꍇ�Ɏg�p����o�f�ł��B
	Dim MOBILE_URL:			MOBILE_URL = "http://m.shigotonavi.jp/"
	Dim MOBILE_URL_SSL:		MOBILE_URL_SSL = "https://m.shigotonavi.jp/"
	Dim PC_URL:				PC_URL = HTTP_CURRENTURL
	Dim LIS_MAILADDRESS:	LIS_MAILADDRESS = "lis@lis21.co.jp"

	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim sRes
	Dim sTo
	Dim sSubject
	Dim sBody
	Dim sPortableSubject
	Dim sPortableBody

	Dim flgRegMail					'���[��DB�o�^���������t���O�F[True]�o�^���� [False]�G���[
	Dim sType						'���[�����M����Ƃ̎�� ["1"]���X�ȊO�̊�� ["2"]���X
	Dim sCompanyName				'���[�����M���̊�Ɩ�
	Dim sStaffName					'���[����M���̋��E�Җ�
	Dim sStaffMailAddress			'���[����M�����E�҂̂o�b���[��
	Dim sStaffPortableMailAddress	'���[����M�����E�҂̃P�[�^�C���[��
	Dim sNoticeMailFlag				'���[�����M��t���O
	Dim sJobTypeDetail				'���[���ɕt�����鋁�l�[�̋�̓I�E�햼

	MailToPerson = False

	'���[�����e�Ŏg�p
	sSQL = "sp_GetDataMailToStaff '" & G_USERID & "', '" & vStaffCode & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
		sCompanyName = ChkStr(oRS.Collect("CompanyName"))
		sStaffName = ChkStr(oRS.Collect("ReceiverName"))
		sStaffMailAddress = ChkStr(oRS.Collect("ReceiverMailAddress"))
		sStaffPortableMailAddress = ChkStr(oRS.Collect("ReceiverPortableMailAddress"))
		sNoticeMailFlag = ChkStr(oRS.Collect("NoticeMailFlag"))
		sType = ChkStr(oRS.Collect("Type"))
		sJobTypeDetail = ChkStr(oRS.Collect("JobTypeDetail"))
	End If
	Call RSClose(oRS)

	If sStaffMailAddress & sStaffPortableMailAddress <> "" Then
		'���[�����M�����e�[�u���Ɋi�[
		flgRegMail = RegMail(dbconn, vID, G_USERID, vStaffCode, vSubject, vBody, vOrderCode, "", "", "", "", "", "", "", vPayFlag)
		MailToPerson = flgRegMail

		'���[��������o�^�̏ꍇ�̂݃��[���̑��M����������B
		If flgRegMail = True Then

			'***************************************************************************
			'���[�����M start
			'---------------------------------------------------------------------------
			'�o�b���[�����M
			If sNoticeMailFlag = "0" Or sNoticeMailFlag = "1" Then
				sSubject = GetMailSubject(sJobTypeDetail)
				sBody = GetMailBodyCompany(vOrderCode, vSubject, vBody, vStaffCode, sStaffName, sCompanyName, sJobTypeDetail, sType)

				sTo = sStaffMailAddress		'���M�惁�[���A�h���X
				If Len(sTo) > 0 Then sRes = SndMail(vMailServer, sTo, vFrom, sSubject, sBody, "")
			End If

			'�P�[�^�C���[�����M
			If sNoticeMailFlag = "0" Or sNoticeMailFlag = "2" Then
				sPortableSubject = "[�����ƃi�r]��Ƃ���Ұْ��M"
				sPortableBody = "���������p���肪�Ƃ��������܂��B" & vbCrLf & _
					"�������l���E��Ģ�����ƃi�r�(���X�������)�ł��B" & vbCrLf & _
					"������ƃi�r���o�C�����ʂ��āA���l��Ƃ���M����Ұق��͂��܂����B" & vbCrLf & _
					"������ƃi�r���۸޲݂��āAҰٓ��e�����m�F�������B" & vbCrLf & vbCrLf & _
					"���g�є�(��޲�)�̏ꍇ:������ƃi�r���o�C���TOP����۸޲݁�My�߰�ނ̢Ұٗ����ݸ��د�" & vbCrLf & _
					MOBILE_URL & vbCrLf & _
					"��PC�ł̏ꍇ:������ƃi�r�TOP����۸޲݁�۸޲��ƭ��̢Ұٗ����ݸ��د�" & vbCrLf & _
					"-----------------"  & vbCrLf & _
					"������Ұق͎������MҰق̂��߁A�ԐM�ł��܂���B�����Ӊ������B" & vbCrLf & _
					"-----------------"  & vbCrLf & _
					"���X�������" & LIS_MAILADDRESS & vbCrLf & _
					"�����ƃi�r���o�C���F" & MOBILE_URL & vbCrLf & _
					"�����ƃi�r�F" & PC_URL

				sTo = sStaffPortableMailAddress
				If Len(sTo) > 0 Then sRes = SndMail(vMailServer, sTo, vFrom, sPortableSubject, sPortableBody, "")
			End If
			'---------------------------------------------------------------------------
			'���[�����M end
			'***************************************************************************
		End If
	End If
End Function
%>
