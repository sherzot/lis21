<%
'**********************************************************************************************************************
'�T�@�v�F��Ə������̊֐��Q
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@�l�擾�p�@������
'�@�@�@�FGetNaviContact			�F�Ώۊ�Ƃ��A�Ώۋ��E�҂Ɖߋ��Ƀ��[���̂���肪���邩�ǂ������`�F�b�N�B
'�@�@�@�FChkNaviRcvMail			�F�Ώۊ�Ƃ��A�Ώۋ��E�҂���ߋ��Ƀ��[�����󂯎�������Ƃ����邩�ǂ������`�F�b�N�B�����C�Z���X�؂��Ƀ��[�����M�\���ǂ����𔻒�
'�@�@�@�FChkMailAble			�F���[�����M�ۂ��擾
'�@�@�@�F�������@�o�͗p�@������
'�@�@�@�FDspScoutLimit			�F��Ƃ̃��C�Z���X�󋵃e�[�u���̕\��, �X�J�E�g�ێ擾
'�@�@�@�FGetHtmlMyMenuAdvManager�F��ƃ}�C���j���[�̋��l�L���Ǘ������g�s�l�k�擾
'�@�@�@�FGetHtmlMyMenuMyOrders	�F��ƃ}�C���j���[�̎��Ћ��l�[�ꗗ
'�@�@�@�F�������@�`�F�b�N�p�@������
'�@�@�@�FChkEditOrder			�F���l�[�o�^�����ۃ`�F�b�N
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F�Ώۊ�Ƃ��A�Ώۋ��E�҂���ߋ��Ƀ��[�����󂯎�������Ƃ����邩�ǂ������`�F�b�N�B
'���@���FrDB			�F
'�@�@�@�FvCompanyCode	�F��ƃR�[�h
'�@�@�@�FvStaffCode		�F���E�҃R�[�h
'�߂�l�FBoolean		�F[True]�X�J�E�g�\ [False]�X�J�E�g�s��
'���@�l�F���C�Z���X�؂��Ƀ��[�����M�\���ǂ����𔻒�
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'�X�@�V�F2008/06/06 LIS K.kokubo �쐬
'******************************************************************************
Function GetNaviContact(ByRef rDB, ByVal vCompanyCode, ByVal vStaffCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	GetNaviContact = False

	sSQL = "up_ChkNaviContact '" & vCompanyCode & "', '" & vStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ContactFlag") = "1" Then GetNaviContact = True
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F�Ώۊ�Ƃ��A�Ώۋ��E�҂���ߋ��Ƀ��[�����󂯎�������Ƃ����邩�ǂ������`�F�b�N�B
'���@���FrDB			�F
'�@�@�@�FvCompanyCode	�F��ƃR�[�h
'�@�@�@�FvStaffCode		�F���E�҃R�[�h
'�߂�l�FBoolean		�F[True]�X�J�E�g�\ [False]�X�J�E�g�s��
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/18
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'******************************************************************************
Function ChkNaviRcvMail(ByRef rDB, ByVal vCompanyCode, ByVal vStaffCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	ChkNaviRcvMail = False

	sSQL = "EXEC up_ChkNaviRcvMail '" & vCompanyCode & "', '" & vStaffCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("RcvFlag") = "1" Then ChkNaviRcvMail = True
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F�Ώۊ�Ƃ��A�Ώۋ��E�҂փ��[�����M�\���ۂ����`�F�b�N�B
'���@���FrDB			�F
'�@�@�@�FvCompanyCode	�F��ƃR�[�h
'�@�@�@�FvStaffCode		�F���E�҃R�[�h
'�@�@�@�FrMailAbleFlag	�F[OUTPUT]���[�����M�ۃt���O [True]�� [False]�s��
'�@�@�@�FrScoutFlag		�F[OUTPUT]�X�J�E�g�t���O [True]�X�J�E�g [False]�X�J�E�g�łȂ�
'�߂�l�FString			�F���[�����M���ӕ���
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'���@�l�F
'���@���F2009/03/23 LIS K.Kokubo �쐬
'******************************************************************************
Function ChkMailAble(ByRef rDB, ByVal vCompanyCode, ByVal vStaffCode, ByRef rMailAbleFlag, ByRef rScoutFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbMailAbleFlag
	Dim dbScoutFlag
	Dim dbReceiveFlag
	Dim dbLicenseStatus
	Dim dbScoutLimitOverFlag
	Dim dbPlanTypeName

	rMailAbleFlag = False
	rScoutFlag = True
	If vCompanyCode = "" Or vStaffCode = "" Then Exit Function

	sSQL = ""
	sSQL = sSQL & "/* �����ƃi�r ���[�����M��,�X�J�E�g�t���O�擾 */" & vbCrLf
	sSQL = sSQL & "EXEC up_ChkMailAble '" & vCompanyCode & "', '" & vStaffCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbMailAbleFlag = oRS.Collect("MailAbleFlag")
		dbScoutFlag = oRS.Collect("ScoutFlag")
		dbReceiveFlag = oRS.Collect("ReceiveFlag")
		dbLicenseStatus = oRS.Collect("LicenseStatus")
		dbScoutLimitOverFlag = oRS.Collect("ScoutLimitOverFlag")
		dbPlanTypeName = oRS.Collect("PlanTypeName")

		If dbMailAbleFlag = "1" Then rMailAbleFlag = True
		If dbScoutFlag = "0" Then rScoutFlag = False

		If dbLicenseStatus = "public" Then
			'���C�Z���X���L��(���s��<=�{��<=�f�ڏI����)
			If dbPlanTypeName = "mail" Then
				'���[���ۋ��v����
			Else
				'���[���ۋ��v�����ȊO�̏ꍇ�́u�X�J�E�g�v�̍l���𓱓�
				If dbScoutFlag = "1" Then
					'�X�J�E�g�Ώ�
					If dbScoutLimitOverFlag = "0" Then
						'�X�J�E�g�����n�j
						ChkMailAble = "���X�J�E�g�̑ΏۂƂȂ�l���ł��B���[���𑗂�ƃX�J�E�g���[���̑��M���ɐ����܂��B"
					Else
						'�X�J�E�g�����m�n
						ChkMailAble = "���X�J�E�g�̐������𒴂��Ă��邽�߁A���̋��E�҂ւ̃��[�����M�͂ł��܂���B"
					End If
				Else
					'��X�J�E�g�Ώ�
					ChkMailAble = "���ߋ��Ƀ��[���̂��Ƃ�̎��т��L�鋁�E�҂ł��B���[���𑗂��Ă��X�J�E�g���[���̑��M���ɐ����܂���B"
				End If
			End If
		ElseIf dbLicenseStatus = "valid" Then
			'���C�Z���X���L�������ǁA�f�ړ��͂܂�(���s��<=�{��<=�f�ڊJ�n��)
			ChkMailAble = "���f�ڊJ�n���ɒB���Ă��Ȃ����߁A�܂����[���𑗐M���邱�Ƃ͂ł��܂���B"
		ElseIf dbLicenseStatus = "mailread" Then
			'���C�Z���X�͐؂�Ă��邪���[���{�����Ԓ�(�f�ڏI����<=�{��<=�f�ڏI����+7��)
			If dbReceiveFlag = "0" Then
				'���[���̎�M���т�����(�X�J�E�g�ΏۂƂ͕ʕ��ł���_�ɒ���)
				ChkMailAble = "�����[���̉{���\���Ԓ��̏ꍇ�A���[���̎�M���т̖������E�҂փ��[���𑗐M���邱�Ƃ͂ł��܂���B"
			Else
				'���[���̎�M���т��L��(���[�����M�\)
				ChkMailAble = "���ߋ��ɂ��Ƃ�̎��т��L�鋁�E�҂ł��B���C�Z���X�؂�ł������[���{�����Ԓ��ł���΃��[���\�ł��B"
			End If
		Else
			'���C�Z���X�؂�
			ChkMailAble = "�����C�Z���X���؂�Ă��邽�߁A���[���𑗐M���邱�Ƃ͂ł��܂���B"
		End If
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F��Ƃ̃��C�Z���X�󋵃e�[�u���̕\��, �X�J�E�g�ێ擾
'���@���FrDB			�F
'�@�@�@�FvCompanyCode	�F
'�@�@�@�FvDspOrderFlag	�F
'�߂�l�FBoolean		�F[True]�X�J�E�g�\ [False]�X�J�E�g�s��
'���@�l�F
'�g�p���F�����ƃi�r/company/c_login.asp
'�@�@�@�F�����ƃi�r/dispatch/d_login.asp
'�X�@�V�F2007/02/13 LIS K.Kokubo �쐬
'******************************************************************************
Function DspScoutLimit(ByRef rDB, ByVal vCompanyCode, ByVal vDspOrderFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbScoutCnt
	Dim dbScoutLimit
	Dim dbMailUnitPrice
	Dim dbShimeFrom
	Dim dbShimeTo
	Dim dbHakouDate
	Dim dbRiyoFromDate
	Dim dbRiyoToDate
	Dim dbPermitOrderCnt
	Dim dbPublicOrderCnt
	Dim dbNotPublicOrderCnt
	Dim dbHalfwayOrderCnt
	Dim dbDspRiyoToDate
	Dim dbPointRemainder
	Dim dbPointWaiting
	Dim dbUsePoint
	Dim dbMailSendPayFlag
	Dim dbMailReceivePayFlag
	Dim dbMchFlag
	Dim dbSpMchFlag
	Dim dbPaySendMailPrice
	Dim dbPayReceiveMailPrice
	Dim dbPayMchPrice
	Dim dbPaySpMchNoticePrice
	Dim dbPaySpMchResponsePrice
	Dim dbPaySendMailCnt
	Dim dbPayReceiveMailCnt
	Dim dbPayMchCnt
	Dim dbPaySpMchNoticeCnt
	Dim dbPaySpMchResponseCnt

	Dim sHTML
	Dim sHTMLPrice
	Dim iScoutAble
	Dim sShimeFrom
	Dim sShimeTo
	Dim sHakouDate
	Dim sRiyoFromDate
	Dim sRiyoToDate
	Dim sDspRiyoToDate
	Dim iPrice

	sHTML = ""
	DspScoutLimit = False

	sSQL = sSQL & "/* �����ƃi�r ��Ƃ̗��p�󋵎擾 */" & vbCrLf
	sSQL = sSQL & "EXEC up_DtlUseStatusCompany_Advertisement '" & vCompanyCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbScoutCnt = ChkStr(oRS.Collect("ScoutCnt"))
		dbScoutLimit = ChkStr(oRS.Collect("ScoutLimit"))
		dbMailUnitPrice = oRS.Collect("MailUnitPrice")
		dbShimeFrom = oRS.Collect("ShimeFrom")
		dbShimeTo = oRS.Collect("ShimeTo")
		dbHakouDate = ChkStr(oRS.Collect("HakouDate"))
		dbRiyoFromDate = ChkStr(oRS.Collect("RiyoFromDate"))
		dbRiyoToDate = ChkStr(oRS.Collect("RiyoToDate"))
		dbPermitOrderCnt = ChkStr(oRS.Collect("PermitOrderCnt"))
		dbPublicOrderCnt = ChkStr(oRS.Collect("PublicOrderCnt"))
		dbNotPublicOrderCnt = ChkStr(oRS.Collect("NotPublicOrderCnt"))
		dbHalfwayOrderCnt = ChkStr(oRS.Collect("HalfwayOrderCnt"))
		dbDspRiyoToDate = ChkStr(oRS.Collect("DspRiyoToDate"))
		dbPointRemainder = ChkStr(oRS.Collect("PointRemainder"))
		dbPointWaiting = oRS.Collect("PointWaiting")
		dbUsePoint = oRS.Collect("UsePoint")
		dbMailSendPayFlag = oRS.Collect("MailSendPayFlag")
		dbMailReceivePayFlag = oRS.Collect("MailReceivePayFlag")
		dbMchFlag = oRS.Collect("MchFlag")
		dbSpMchFlag = oRS.Collect("SpMchFlag")
		dbPaySendMailPrice = oRS.Collect("PaySendMailPrice")
		dbPayReceiveMailPrice = oRS.Collect("PayReceiveMailPrice")
		dbPayMchPrice = oRS.Collect("PayMchPrice")
		dbPaySpMchNoticePrice = oRS.Collect("PaySpMchNoticePrice")
		dbPaySpMchResponsePrice = oRS.Collect("PaySpMchResponsePrice")
		dbPaySendMailCnt = oRS.Collect("PaySendMailCnt")
		dbPayReceiveMailCnt = oRS.Collect("PayReceiveMailCnt")
		dbPayMchCnt = oRS.Collect("PayMchCnt")
		dbPaySpMchNoticeCnt = oRS.Collect("PaySpMchNoticeCnt")
		dbPaySpMchResponseCnt = oRS.Collect("PaySpMchResponseCnt")
	End If
	Call RSClose(oRS)

	iScoutAble = dbScoutLimit - dbScoutCnt
	sShimeFrom = Year(dbShimeFrom) & "�N" & Month(dbShimeFrom) & "��" & Day(dbShimeFrom) & "��"
	sShimeTo = Year(dbShimeTo) & "�N" & Month(dbShimeTo) & "��" & Day(dbShimeTo) & "��"
	sHakouDate = Year(dbHakouDate) & "�N" & Month(dbHakouDate) & "��" & Day(dbHakouDate) & "��"
	sRiyoFromDate = Year(dbRiyoFromDate) & "�N" & Month(dbRiyoFromDate) & "��" & Day(dbRiyoFromDate) & "��"
	sRiyoToDate = Year(dbRiyoToDate) & "�N" & Month(dbRiyoToDate) & "��" & Day(dbRiyoToDate) & "��"
	sDspRiyoToDate = Year(dbDspRiyoToDate) & "�N" & Month(dbDspRiyoToDate) & "��" & Day(dbDspRiyoToDate) & "��"

	DspScoutLimit = True

	sHTML = sHTML & "<div style=""margin-bottom:15px;"">"
	sHTML = sHTML & "<table style=""margin:0px;"">"
	sHTML = sHTML & "<colgroup>"
	sHTML = sHTML & "<col style=""width:112px;padding:3px;background-color:#e8e8ff;""></col>"
	sHTML = sHTML & "<col style=""width:473px;padding:3px;""></col>"
	sHTML = sHTML & "</colgroup>"
	sHTML = sHTML & "<tbody>"

	iPrice = (dbPaySendMailPrice + dbPayReceiveMailPrice + dbPayMchPrice + dbPaySpMchNoticePrice + dbPaySpMchResponsePrice) - dbUsePoint * 100

	'<����>
	If InStr(dbMailSendPayFlag & dbMailReceivePayFlag & dbMchFlag & dbSpMchFlag, "1") > 0 Then
		sHTMLPrice = ""
		sHTMLPrice = sHTMLPrice & "<div style=""float:left;width:39%;"">"
		sHTMLPrice = sHTMLPrice & "���݂̗����v&nbsp;:&nbsp;<b><span style=""color:#ff0000;"">" & GetJapaneseYen(iPrice) & "</span></b>"
		If dbUsePoint > 0 Then sHTMLPrice = sHTMLPrice & "&nbsp;(" & dbUsePoint & "pt���p�F" & dbUsePoint * 100 & "�~����)"
		sHTMLPrice = sHTMLPrice & "</div>"

		sHTMLPrice = sHTMLPrice & "<div style=""float:right;width:59%;"">"
		If (dbMailSendPayFlag = "1" Or dbMailReceivePayFlag = "1") Or (dbPaySendMailCnt + dbPayReceiveMailCnt > 0) Then
			sHTMLPrice = sHTMLPrice & "�ۋ����[�����M��&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPaySendMailCnt & "��</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPaySendMailPrice) & "</span></b>)<br>"
			sHTMLPrice = sHTMLPrice & "�ۋ����[����M��&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPayReceiveMailCnt & "��</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPayReceiveMailPrice) & "</span></b>)<br>"
		End If

		If dbMchFlag = "1" Or dbPayMchCnt > 0 Then
			sHTMLPrice = sHTMLPrice & "�}�b�`���O�l�މ��吔&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPayMchCnt & "��</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPayMchPrice) & "</span></b>)<br>"
		End If

		If dbSpMchFlag = "1" Or (dbPaySpMchNoticeCnt + dbPaySpMchResponseCnt > 0) Then
			sHTMLPrice = sHTMLPrice & "�ʒm���[����&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPaySpMchNoticeCnt & "��</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPaySpMchNoticePrice) & "</span></b>)<br>"
			sHTMLPrice = sHTMLPrice & "���吔&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPaySpMchResponseCnt & "��</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPaySpMchResponsePrice) & "</span></b>)<br>"
		End If

		sHTMLPrice = sHTMLPrice & "</div>"
		sHTMLPrice = sHTMLPrice & "<div style=""clear:both;""></div>"
		sHTMLPrice = sHTMLPrice & "<div class=""line1""></div>"
	End If
	'</����>

	If G_PLANTYPE = "mail" Then

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">�ۋ���</th>"
		sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"

		sHTML = sHTML & sHTMLPrice

		sHTML = sHTML & "<span style=""font-size:10px;"">���Z�o���ԁF" & sShimeFrom & "&nbsp;�`&nbsp;" & sShimeTo & "�Y</span>"
		If Date < "2009/08/01" Then sHTML = sHTML & "<br>��<span style=""color:#ff0000;"">�W���P���ȍ~�̎�M���[�����ۋ������悤�ɂȂ�܂��B</span>"

		If vDspOrderFlag = True Then
			'���[���ۋ��v�����̏ꍇ�͗��p�󋵂ւ̃����N��\��
			sHTML = sHTML & "<p class=""m0""><a href=""/company/license/mailplan_status.asp"">���ۋ��󋵂̉ߋ������m�F</a>&nbsp;...&nbsp;�ߋ��̗����̖��ׂȂǂ��m�F�ł��܂��B</p>"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	Else
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">�X�J�E�g���[��</th>"
		sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"
		sHTML = sHTML & "�X�J�E�g�\���F<b><span style=""color:#0000ff;"">" & iScoutAble & "��</span></b>&nbsp;&nbsp;"
		sHTML = sHTML & "�X�J�E�g���M���F" & dbScoutCnt & "���^�ő�" & dbScoutLimit & "���܂�"

		If sHTMLPrice <> "" Then
			sHTML = sHTML & "<div class=""line1""></div>"
			sHTML = sHTML & sHTMLPrice
		End If

		If dbMchFlag = "1" Or dbSpMchFlag = "1" Then
			sHTML = sHTML & "<p class=""m0""><a href=""/company/license/mailplan_status.asp"">���ۋ��󋵂̉ߋ������m�F</a>&nbsp;...&nbsp;���z���p���ȊO�̉ߋ��̗����̖��ׂȂǂ��m�F�ł��܂��B</p>"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	If vDspOrderFlag = True Then
		If G_PLANTYPE = "mail" Then
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">�|�C���g��</th>"
			sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"
			sHTML = sHTML & "���|�C���g&nbsp;:&nbsp;<b><span style=""color:Red;"">" & dbPointWaiting + dbPointRemainder & "pt</span></b>&nbsp;&nbsp;"
			sHTML = sHTML & "�i���p�\�|�C���g&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPointRemainder & "pt</span></b>�j<br>"
			sHTML = sHTML & "<span style=""font-size:10px;"">���|�C���g�͔���������Q������ɗ��p�\�|�C���g�ƂȂ�܂��B</span>"
			sHTML = sHTML & "<p class=""m0""><a href=""/company/point/"" style="""">���|�C���g�Ǘ�</a>&nbsp;...&nbsp;�|�C���g�̎c���̊m�F�◘�p�\���E������ł��܂��B</p>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "</tr>"
		End If

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">�f�ڊ���</th>"
		sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"
		sHTML = sHTML & sRiyoFromDate & "�`" & sDspRiyoToDate
		If G_PLANTYPE = "mail" Then
			sHTML = sHTML & "<br>"
			sHTML = sHTML & "<span style=""font-size:10px;"">�����O�C��������ƁA�f�ڏI���������̓����Q������́Y��(����)�ɍX�V����܂��B</span>"
		End If
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">���l�[���J��</th>"
		sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"
		sHTML = sHTML & "<p class=""m0"">"
		sHTML = sHTML & "�f�ڒ�(&nbsp;" & dbPublicOrderCnt & "��&nbsp;)&nbsp;&nbsp;"
		sHTML = sHTML & "��f��(&nbsp;" & dbNotPublicOrderCnt & "��&nbsp;)&nbsp;&nbsp;"
		sHTML = sHTML & "�R����(&nbsp;" & dbPermitOrderCnt & "��&nbsp;)&nbsp;&nbsp;"
		sHTML = sHTML & "�쐬��(&nbsp;" & dbHalfwayOrderCnt & "��&nbsp;)&nbsp;&nbsp;"
		sHTML = sHTML & "</p>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"
	sHTML = sHTML & "</div>"

	Response.Write sHTML
End Function

'******************************************************************************
'�T�@�v�F��ƃ}�C���j���[�̋��l�L���Ǘ������g�s�l�k�擾
'���@���FrDB			�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvCompanyCode	�F��ƃR�[�h
'�@�@�@�FvCompanyKbn	�F��Ǝ��
'�@�@�@�FvLicenseFlag	�F���C�Z���X�󋵃t���O ["1"]���p��
'�@�@�@�FvCollectionCnt	�F�f�ڒ����l�[����
'�߂�l�FString
'�g�p���F�����ƃi�r/company/c_login.asp
'�@�@�@�F�����ƃi�r/dispatch/d_login.asp
'���@�l�F
'�X�@�V�F2007/12/04 LIS K.Kokubo �쐬
'�@�@�@�F2008/01/31 LIS K.Kokubo �v���`�i�v�����A�S�[���h�v�����ŊO������̃A�N�Z�X�̏ꍇ�ł́u���l�[�̐V�K�쐬�v�����N���\��
'�@�@�@�F2009/03/11 LIS K.Kokubo �u���l�[�̐V�K�쐬�v�����N��S�Ẵv�����ɊJ��
'�@�@�@�F2009/06/25 LIS K.Kokubo ���l�[���������ۑ��ǉ�
'�@�@�@�F2009/07/02 LIS K.Kokubo �ꊇ���[���Ǘ��ǉ�
'******************************************************************************
Function GetHtmlMyMenuAdvManager(ByRef rDB, ByVal vCompanyCode, ByVal vCompanyKbn, ByVal vLicenseFlag, ByVal vPlanType, ByVal vCollectionCnt)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim dbCnt
	Dim sHTML
	Dim flgNew

	sHTML = ""

	If vLicenseFlag = "1" Then
		'�i�r���l�L�����C�Z���X�𗘗p��
		sHTML = sHTML & "<table border=""0"" style=""width:100%;"">"

		'sHTML = sHTML & "<tr>"
		'sHTML = sHTML & "<td colspan=""2"" bgcolor=""#666699""><font color=""#FFFFFF"">���l�L��</font></td>"
		'sHTML = sHTML & "</tr>"

		'<�u���l�[�̐V�K�쐬�v�����N�\��>
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"

		flgNew = True
		sSQL = "SELECT COUNT(*) AS Cnt FROM C_Info AS A INNER JOIN C_SupplementInfo AS B ON A.OrderCode = B.OrderCode WHERE RegistCommit = '0' AND A.CompanyCode = '" & G_USERID & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbCnt = oRS.Collect("Cnt")
			If dbCnt >= 5 Then flgNew = False
		End If
		If flgNew = True Then
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/orderedit/new.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
		Else
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td bgcolor=""#E8E8FF"" style=""border-bottom:1px solid #ffffff;"">"

		If flgNew = True Then
			'�쐬�r�����l�[�����K��ȓ��Ȃ�V�K�쐬�̃����N��\��
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/orderedit/new.asp"">���l�[�̐V�K�쐬</a>"
		Else
			'�쐬�r�����l�[�����K��I�[�o�[�Ȃ�V�K�쐬�̃����N��\��
			sHTML = sHTML & "<p class=""m0"">���l�[�̐V�K�쐬</p>"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""border-bottom:1px solid #ffffff;"">"
		If dbCnt >= 5 Then
			sHTML = sHTML & "<p class=""m0"" style=""line-height:16px;"">���쐬�r���̋��l�[��<b>�T��</b>�܂Ŏ������ł��܂��B�V���ɋ��l�[���쐬�������ꍇ�́A�쐬�r���̋��l�[�ꗗ�́u�ҏW�v�{�^����苁�l�[���m�肷�邩�A�u�폜�v�{�^���ō폜����K�v������܂��B</p>"
		Else
			sHTML = sHTML & "�V�������l�[���쐬���܂��B�y<a href=""http://jinzai.shigotonavi.co.jp/joboffer/make_advertisement.asp"" target=""blank_"">���l�[�쐬�̃|�C���g</a>�z<br>"
		End If
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		'</�u���l�[�̐V�K�쐬�v�����N�\��>

		If vCollectionCnt = 0 Then
			'��W���̋��l�������ꍇ
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"">"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td bgcolor=""#E8E8FF"" style=""border-bottom:1px solid #ffffff;"">"
			sHTML = sHTML & "�|�|�|"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "���݁A���l��W���Ă��鋁�l�[���������߁A���E�Ҍ����͂����p�ł��܂���B<br>"
			sHTML = sHTML & "���炽�ɋ��l�[���쐬����������΁A���ЂɂĊm�F�̏�A�����p���������܂��B"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "</tr>"
		Else
			'��W���̋��l������ꍇ

			'���l�[�̃R�s�[�쐬
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">���l�[�̃R�s�[�쐬</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "��Ђ̊����̋��l�[�����ɐV�������l�[���쐬�ł��܂��B"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "</tr>"

			'���l�[�C��
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">���l�[�C��</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>���݌�Ђɂĕ�W���Ă��鋁�l�[�̌����ƏC�����ł��܂��B</td>"
			sHTML = sHTML & "</tr>"

			'���E�Ҍ���
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">���E�҂̌����ƃX�J�E�g</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>���E�҂��������A�X�J�E�g�ł��܂��B</td>"
			sHTML = sHTML & "</tr>"

			'���E�Ҍ��������Ǘ�
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/searchstaffcondition/list.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/searchstaffcondition/list.asp"">���E�Ҍ��������Ǘ�</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>�ۑ��������E�Ҍ����������폜�E���̕ύX���܂��B</td>"
			sHTML = sHTML & "</tr>"
		End If

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td colspan=""3"">"
		sHTML = sHTML & "<div class=""line1""></div>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"


		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/company_reg1.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/company_reg1.asp"">���Џ��X�V</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>���l���ȊO�̕����A��ЊT�v�̕ҏW���ł��܂��B</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">��Ǝʐ^�f��</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "��ƏЉ�ɗ��p����A��\�I�ȉ摜�i���S�Ȃǁj��o�^�ł��܂��B"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">���l�[�p�摜�X�g�b�N</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "���l�[�ɕ����摜���ڂ���ꍇ�́A�����ŉ摜��o�^���Ă����܂��B"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">���[������</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>���E�҂Ƃ̃��[��������A�̗p�̐i���Ǘ����\�B</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/lumpmail/list.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/lumpmail/list.asp"">�ꊇ���[���Ǘ�</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""border-bottom:1px solid #ffffff;"">�ꊇ���[���̗\��󋵂̊m�F�A�ꊇ���[���̍쐬�E���M���\�B</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "mailtemplate/manager.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "mailtemplate/manager.asp"">���[���e���v���[�g�Ǘ�</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>���[�����쐬����ۂɗ��p�ł��鐗�`���Ǘ����܂��B�y<a href=""/company/c_scout3point.asp"">�X�J�E�g���[���쐬�̃|�C���g</a>�z</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</table>"
	End If
	'******************************************************************************
	'** ���l�L�� end
	'******************************************************************************

	GetHtmlMyMenuAdvManager = sHTML
End Function

'******************************************************************************
'�T�@�v�F���m��̋��l�[�ꗗ
'���@���FrDB			�F�ڑ����c�a
'�@�@�@�FvCompanyCode	�F��ƃR�[�h
'�߂�l�FString			�F���Ћ��l�[�ꗗ�g�s�l�k
'���@�l�F
'�g�p���F�����ƃi�r/company/mailtoperson.asp
'�X�@�V�F2008/11/04 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHtmlUnCommitOrders(ByRef rDB, ByVal vCompanyCode)
	'<�ϐ��錾>
	Dim sHTML

	Dim dbOrderCode
	Dim dbJobTypeDetail
	Dim dbUpdateDay
	'</�ϐ��錾>

	'<�ϐ�������>
	sHTML = ""
	'</�ϐ�������>

	sSQL = "SELECT A.OrderCode, A.JobTypeDetail, A.UpdateDay FROM C_Info AS A INNER JOIN C_SupplementInfo AS B ON A.OrderCode = B.OrderCode WHERE RegistCommit = '0' AND A.CompanyCode = '" & G_USERID & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sHTML = sHTML & "<table class=""pattern3"" style=""width:100%;"">"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""width:150px"">�X�V����</th>"
		sHTML = sHTML & "<th style=""width:150px"">�L������</th>"
		sHTML = sHTML & "<th>�E��</th>"
		sHTML = sHTML &"</tr>"
		sHTML = sHTML &"</thead>"
		sHTML = sHTML & "<tbody>"

		Do While GetRSState(oRS) = True
			dbOrderCode = oRS.Collect("OrderCode")
			dbJobTypeDetail = oRS.Collect("JobTypeDetail")
			dbUpdateDay = oRS.Collect("UpdateDay")

			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>" & GetDateStr(dbUpdateDay, "/") & "<br>" & GetTimeStr(dbUpdateDay, ":") & "</td>"
			sHTML = sHTML & "<td>" & GetDateStr(DateAdd("d", 6, dbUpdateDay), "/") & "<br>00:00:00</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<form action=""/company/orderedit/base.asp?ordercode="& dbOrderCode & """ method=""post"" style=""display:inline;""><input class=""btn1"" type=""submit"" value=""�ҏW""></form>&nbsp;"
			sHTML = sHTML & "<form action="""" method=""post"" style=""display:inline;"" onsubmit=""return confirm('�폜���܂����H');""><input name=""frmdelordercode"" type=""hidden"" value=""" & dbOrderCode & """><input class=""btn1"" type=""submit"" value=""�폜""></form>&nbsp;"
			sHTML = sHTML & "<input type=""text"" value=""" & dbJobTypeDetail & """ style=""width:350px; border-width:0px; background-color:transparent;"">"
			sHTML = sHTML &"</td>"
			sHTML = sHTML &"</tr>"

			oRS.MoveNext
		Loop

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""width:60px; padding:0px; border-width:0px;""></td>"
		sHTML = sHTML & "<td style=""width:60px; padding:0px; border-width:0px;""></td>"
		sHTML = sHTML & "<td style=""width:480px; padding:0px; border-width:0px;""></td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
	End If

	GetHtmlUnCommitOrders = sHTML
End Function

'******************************************************************************
'�T�@�v�F��ƃ}�C���j���[�̎��Ћ��l�[�ꗗ
'���@���FrDB			�F�ڑ����c�a
'�@�@�@�FvUserType		�F���O�C�����[�U���
'�@�@�@�FvCompanyCode	�F��ƃR�[�h
'�@�@�@�FvPageSize		�F�P�y�[�W������̍ő�o�͌���
'�@�@�@�FvPage			�F�o�̓y�[�W
'�@�@�@�FvSort			�F�f�[�^���ёւ�
'�@�@�@�FvPersonName	�F�i���݋��l�S����
'�߂�l�FString			�F���Ћ��l�[�ꗗ�g�s�l�k
'���@�l�F
'�g�p���F�����ƃi�r/company/c_login.asp
'���@���F2007/02/18 LIS K.Kokubo �쐬
'�@�@�@�F2009/05/22 LIS K.Kokubo ���Ћ��l�ꗗ�e�[�u���̃e�[�u���w�b�_���\���B�i���ǃ��[���A���ǋ��E�҂̐������ǂ܂�Ă��Ȃ������΍�j
'******************************************************************************
Function GetHtmlMyMenuMyOrders(ByRef rDB, ByVal vCompanyCode, ByVal vUserType, ByVal vPageSize, ByVal vPage, ByVal vSort, ByVal vPersonName)
	'<�ϐ��錾>
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim oRS3
	Dim flgQE
	Dim sError

	Dim dbOrderCode		'���R�[�h
	Dim dbJobTypeDetail	'��̓I�E�햼
	Dim dbPersonName	'���l�S����
	Dim dbMailCnt		'���ǃ��[����
	Dim dbStaffCnt		'�V�����E�Ґ�
	Dim dbPublicFlag	'�f�ڏ�ԃt���O ["1"]�f�ڒ� ["0"]��f��
	Dim dbSecretFlag	'�V�[�N���b�g���l�t���O ["1"]�V�[�N���b�g���l ["0"]��ʌ��J���l
	Dim dbSearchName	'����������
	Dim dbSearchParam	'���������p�����[�^

	Dim sHTML
	Dim sHTML2
	Dim sPageControl
	Dim sURL
	Dim sParam
	Dim sJobTypeDetail
	Dim sMailCnt
	Dim sStaffCnt
	Dim iRow			'���Ћ��l�[�̏o�͒����R�[�h�ԍ�
	Dim sOnChange
	'�ۑ����������擾�p
	Dim tmpAbsolutePage
	Dim sXML
	'</�ϐ��錾>

	'�t�q�k
	If vUserType = "company" Then
		sURL = "/company/c_login.asp"
	ElseIf vUserType = "dispatch" Then
		sURL = "/dispatch/d_login.asp"
	End If

	sParam = ""
	'���ёւ��p�����[�^
	If vSort <> "" Then
		If sParam <> "" Then sParam = sParam & "&amp;"
		sParam = sParam & "sort=" & vSort
	End If
	If vPersonName <> "" Then
		If sParam <> "" Then sParam = sParam & "&amp;"
		sParam = sParam & "pn=" & Server.URLEncode(vPersonName)
	End If
	If sParam <> "" Then sParam = "?" & sParam

	'���Ћ��l�[�ꗗ
	sHTML = ""
	sSQL = "up_LstMyMenuMyOrders '" & vCompanyCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	oRS.PageSize = vPageSize
	If vPersonName <> "" Then oRS.Filter = "PersonName = '" & vPersonName & "'"
	If GetRSState(oRS) = True Then
		If vPage <> "" Then oRS.AbsolutePage = vPage
		Select Case vSort
			Case Else: oRS.Sort = "PublicFlag DESC, OrderCode DESC"
		End Select
	End If

	'�y�[�W�R���g���[���擾
	sPageControl = GetHtmlPageControlParam(rDB, oRS, vPageSize, vPage, sURL & sParam, "myorder")

	'<�ۑ����������擾>
	tmpAbsolutePage = oRS.AbsolutePage
	iRow = 1
	sXML = "<root>"
	Do While GetRSState(oRS) = True And iRow <= vPageSize
		dbOrderCode = oRS.Collect("OrderCode")
		sXML = sXML & "<order><ordercode>" & dbOrderCode & "</ordercode></order>"
		oRS.MoveNext
	Loop
	sXML = sXML & "</root>"

	sSQL = ""
	sSQL = sSQL & "/* �ۑ����������擾 */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstCMPSearchStaffCondition_XML '" & sXML & "';"
	flgQE = QUERYEXE(dbconn, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		Set oRS2.ActiveConnection = Nothing
	End If
	oRS.AbsolutePage = tmpAbsolutePage
	'</�ۑ����������擾>

	iRow = 1
	Do While GetRSState(oRS) = True And iRow <= vPageSize
		dbOrderCode = oRS.Collect("OrderCode")
		dbJobTypeDetail = oRS.Collect("JobTypeDetail")
		dbPersonName = oRS.Collect("PersonName")
		dbPublicFlag = oRS.Collect("PublicFlag")
		dbSecretFlag = oRS.Collect("SecretFlag")

		sJobTypeDetail = dbJobTypeDetail
		If Len(sJobTypeDetail) > 29 Then sJobTypeDetail = Left(sJobTypeDetail, 29) & "..."

		'���ǃ��[����
		sSQL = "up_CntMyMenuNotReadMail '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then
			dbMailCnt = oRS3.Collect("Cnt")
			If dbMailCnt > 0 Then
				sMailCnt = "<div class=""iconred"">���ǃ��[��</div>&nbsp;"
				sMailCnt = sMailCnt & "<a href=""" & HTTP_CURRENTURL & "company/mailhistory_company.asp?soc=" & dbOrderCode & """>" & dbMailCnt & "��</a>"
			Else
				sMailCnt = "<div class=""icongray"">���ǃ��[��</div>&nbsp;<span style=""color:#999999;"">�Ȃ�</span>"
			End If
		End If
		Call RSClose(oRS3)

		'�V�����E�Ґ�
		sSQL = "up_CntMyMenuNewStaff '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then
			dbStaffCnt = oRS3.Collect("Cnt")
			If dbStaffCnt > 0 Then
				sStaffCnt = "<div class=""iconred"">���ǋ��E��</div>&nbsp;"
				sStaffCnt = sStaffCnt & "<a href=""" & HTTP_CURRENTURL & "staff/person_list.asp?ordercode=" & dbOrderCode & "&amp;rdfrom=" & GetDateStr(DateAdd("d", -9, Date), "") & """>" & dbStaffCnt & "�l</a>"
			Else
				sStaffCnt = "<div class=""icongray"">���ǋ��E��</div>&nbsp;<span style=""color:#999999;"">�Ȃ�</span>"
			End If
		End If
		Call RSClose(oRS3)

		sHTML = sHTML & "<tr>"
		'<���l>
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "order/order_detail.asp?ordercode=" & dbOrderCode & """>" & sJobTypeDetail & "</a>&nbsp;"
		sHTML = sHTML & "(�S���F" & dbPersonName & ")<br>"
		If dbPublicFlag = "1" Then
			sHTML = sHTML & "<div class=""iconredbg"">�f�ڒ�</div>&nbsp;"
		ElseIf dbPublicFlag = "0" Then
			sHTML = sHTML & "<div class=""icongraybg"">��f��</div>&nbsp;"
		End If
		If dbSecretFlag = "1" Then sHTML = sHTML & "<div class=""icongraybg"">SECRET</div>&nbsp;"
		sHTML = sHTML & sMailCnt & "&nbsp;"
		sHTML = sHTML & sStaffCnt

		sHTML = sHTML & "</td>"
		'</���l>

		'<���E�Ҍ���>
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<div style=""margin-bottom:3px;"">"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""��������"" onclick=""location.href='/staff/person_list.asp?ordercode=" & dbOrderCode & "';"">&nbsp;"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ڍ׌���"" onclick=""location.href='/staff/person_search_detail.asp?ordercode=" & dbOrderCode & "&amp;setdata=1';"">"
		sHTML = sHTML & "</div>"

		'<�ۑ���������>
		If GetRSState(oRS2) = True Then
			oRS2.Filter = 0
			oRS2.Filter = "OrderCode = '" & dbOrderCode & "'"
			If GetRSState(oRS2) = True Then
				sHTML = sHTML & "<div class=""line1""></div>"
				sHTML = sHTML & "<ul>"
				Do While GetRSState(oRS2) = True
					dbSearchName = oRS2.Collect("SearchName")
					dbSearchParam = oRS2.Collect("SearchParam")

					sHTML = sHTML & "<li>�E<a href=""" & HTTP_CURRENTURL & "staff/person_list.asp?ordercode=" & dbOrderCode & "&amp;" & dbSearchParam & """>" & dbSearchName & "</a></li>"

					oRS2.MoveNext
				Loop
				sHTML = sHTML & "</ul>"
			End If
			'</�ۑ���������>
		End If
		sHTML = sHTML & "</td>"
		'</���E�Ҍ���>
		sHTML = sHTML & "</tr>"

		iRow = iRow + 1
		oRS.MoveNext
	Loop
	Call RSClose(oRS2)
	Call RSClose(oRS)

	sHTML2 = ""
	If sHTML <> "" Then
		sOnChange = "location.href='" & sURL & "?sort=" & vSort & "&amp;pn=' + escape(this.value) + '#myorder';"
		sHTML2 = sHTML2 & "<div id=""myorder""></div>"
		sHTML2 = sHTML2 & sPageControl & vbCrLf
		sHTML2 = sHTML2 & "<table class=""pattern3"" border=""0"" style=""width:100%;"">"
		'sHTML2 = sHTML2 & "<thead>"
		'sHTML2 = sHTML2 & "<tr>"
		'sHTML2 = sHTML2 & "<th colspan=""3"">���Ћ��l�[�ꗗ</th>"
		'sHTML2 = sHTML2 & "</tr>"
		'sHTML2 = sHTML2 & "</thead>"
		sHTML2 = sHTML2 & "<thead>"
		sHTML2 = sHTML2 & "<tr>"
		sHTML2 = sHTML2 & "<th style=""width:388px;"">�E��&nbsp;<select name=""pn"" onchange=""" & sOnChange & """><option value="""">--���l�S��--</option>" & GetContactPersonNameOptionHtml(vCompanyCode, vPersonName) & "</select></th>"
		sHTML2 = sHTML2 & "<th style=""width:189px;"">���E�Ҍ���</th>"
		sHTML2 = sHTML2 & "</tr>"
		sHTML2 = sHTML2 & "</thead>"
		sHTML2 = sHTML2 & "<tbody>"
		sHTML2 = sHTML2 & sHTML
		sHTML2 = sHTML2 & "</tbody>"
		sHTML2 = sHTML2 & "</table>" & vbCrLf
		sHTML2 = sHTML2 & sPageControl & vbCrLf
	End If

	GetHtmlMyMenuMyOrders = sHTML2
End Function

'******************************************************************************
'�T�@�v�F���l�[�o�^�����ۃ`�F�b�N
'���@���FvOrderCode	�F���R�[�h
'�@�@�@�FvUserID	�F���O�C�������[�U�R�[�h
'�@�@�@�FvUseFlag	�F���O�C������Ƃ̃��C�Z���X�̗L���t���O
'�߂�l�FBoolean	�F[True]���l�[�o�^�\ [False]���l�[�o�^�s��
'���@�l�F
'�g�p���F�����ƃi�r/company/order/edit1.asp
'�X�@�V�F2008/10/08 LIS K.kokubo �쐬
'******************************************************************************
Function ChkEditOrder(ByVal vOrderCode, ByVal vUserID, ByVal vUseFlag)
	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim dbCheck
	Dim dbLicenseFlag

'	If vOrderCode = "" Then Exit Function

	'<���C�Z���X�؂�̓}�C���j���[�փ��_�C���N�g>
	sSQL = "EXEC up_DtlNaviLicense_Now '" & vUserID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("LicenseType1Flag") <> "1" Then Exit Function
	End If
	Call RSClose(oRS)
	'</���C�Z���X�؂�̓}�C���j���[�փ��_�C���N�g>

	'<���O�C�����̊�Ƃ̏��R�[�h���ǂ������`�F�b�N>
	sSQL = "sp_ChkCompanyOrder '" & vUserID & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbCheck = oRS.Collect("CheckFlag")
		dbLicenseFlag = oRS.Collect("LicenseFlag")
	End If
	Call RSClose(oRS)
	If vOrderCode = "" Then dbCheck = "1"
	If dbCheck = "0" And dbLicenseFlag = "0" Then Exit Function
	'</���O�C�����̊�Ƃ̏��R�[�h���ǂ������`�F�b�N>

	ChkEditOrder = True
End Function
%>
