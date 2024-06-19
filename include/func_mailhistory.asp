<%
'**********************************************************************************************************************
'�T�@�v�F���[���ꗗ�y�[�W	/staff/mailhistory_person_entity.asp
'�@�@�@�F��L�y�[�W�ŏo�͗p�̊֐��Q�����̃t�@�C���ɗp�ӂ���B
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���[���ꗗ�y�[�W�o�͗p�@������
'�@�@�@�FGetHtmlMailPageControlParam		�F���[���ꗗ�y�[�W�̃y�[�W�R���g���[���̂g�s�l�k���擾
'�@�@�@�FGetHtmlMailHistory					�F���[���ꗗ�g�s�l�k�擾
'�@�@�@�FGetHtmlMailHistoryDetailStaff		�F�X�^�b�t�̃��[���ꗗ�̂P�s�g�s�l�k���擾
'�@�@�@�FGetHtmlMailHistoryDetailCompany	�F��Ƃ̃��[���ꗗ�̂P�s�g�s�l�k���擾
'�@�@�@�FGetHtmlMailControl					�F�폜�{�^���A�X�V�{�^���o��
'�@�@�@�FGetHtmlMailSearch					�F���[�������o��
'�@�@�@�FDspNoMail							�F���[���������ꍇ�̏o��
'�@�@�@�F
'�@�@�@�F�������@���[���Ǘ��@������
'�@�@�@�FUpdMailListStf	�F���[���ꗗ����]���Ɣ��l�̍X�V
'�@�@�@�FDelMailStf		�F���[���̍폜
'�@�@�@�F
'�@�@�@�F�������@���[���֘A�l�擾�@������
'�@�@�@�FGetSortImg		�F���ёւ��̃{�^�����擾
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F���[���ꗗ�y�[�W�̃y�[�W�R���g���[���̂g�s�l�k���擾
'���@���FrDB		�F�ڑ����c�a�R�l�N�V����
'�@�@�@�FrRS		�Fup_SearchMail �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType	�F���O�C�������[�U���
'�@�@�@�FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvPageSize	�F�ő�\���\���[����
'�@�@�@�FvPage		�F���݂̃y�[�W
'�@�@�@�FvURL		�F
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlMailPageControlParam(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vMode, ByVal vPageSize, ByVal vPage, ByVal vURL)
	Dim iPage
	Dim iPageSize
	Dim iMaxPage
	Dim iStartPage
	Dim iEndPage
	Dim idxPage
	Dim sAnc
	Dim sAncRcvSnd
	Dim sHtml

	GetHtmlMailPageControlParam = ""
	sHtml = ""

	If GetRSState(rRS) = False Then Exit Function

	If InStr(vURL, "?") > 0 Then
		sAnc = vURL & "&amp;"
	Else
		sAnc = vURL & "?"
	End If

	If vUserType = "staff" Then
		sAncRcvSnd = HTTP_CURRENTURL & "staff/mailhistory_person.asp"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sAncRcvSnd = HTTP_CURRENTURL & "company/mailhistory_company.asp"
	End If

	iPage = CInt(vPage)
	iPageSize = CInt(vPageSize)

	rRS.PageSize = iPageSize
	iMaxPage = rRS.PageCount

	'�͈͊O�̃y�[�W�w��΍�
	If iPage > iMaxPage Then iPage = iMaxPage
	If iPage < 1 Then iPage = 1

	rRS.AbsolutePage = iPage

	'�\���J�n�y�[�W�ԍ����w��
	iStartPage = iPage - 2
	If iStartPage < 1 Then iStartPage = 1

	'�\���I���y�[�W�ԍ����w��
	iEndPage = iPage + 2
	If iEndPage - iStartPage < 4 Then iEndPage = 5
	If iEndPage > iMaxPage Then iEndPage = iMaxPage

	sHtml = sHtml & "<div class=""mail_joken"">"

	If vMode = "1" Then
		sHtml = sHtml & "<h3>���M��</h3>"
		sHtml = sHtml & "<div>"
		sHtml = sHtml & "<a href=""" & sAncRcvSnd & "?mode=0""><span>��M��</span></a><span>���M��</span>" & vbCrLf
		sHtml = sHtml & "<a href=""#mail_under""><span>���M���[������</span></a></div>"
	Else
		sHtml = sHtml & "<h3>��M��</h3>"
		sHtml = sHtml & "<div>"
		sHtml = sHtml & "<span>��M��</span><a href=""" & sAncRcvSnd & "?mode=1""><span>���M��</span></a>" & vbCrLf
		sHtml = sHtml & "<a href=""#mail_under""><span>��M���[������</span></a></div>"
	End If

	sHtml = sHtml & "</div><br clear=""both"">" & vbCrLf
		sHtml = sHtml & "<div>" & vbCrLf
	sHtml = sHtml & "<div class=""left"">" & vbCrLf

	If CInt(iStartPage) <> 1 Then sHtml = sHtml & "�c"
	For idxPage = iStartPage To iEndPage	'�y�[�W�ԍ���\��
		sHtml = sHtml & "�@"
		If idxPage = CInt(iPage) Then		'�w��y�[�W�̕\��
			sHtml = sHtml & "[" & idxPage & "]"
		Else
			sHtml = sHtml & "<a href=""" & sAnc & "page=" & idxPage & """>" & idxPage & "</a>"
		End If
	Next
	If iEndPage < iMaxPage Then sHtml = sHtml & "�@�c"

	sHtml = sHtml & "</div>" & vbCrLf

	
	sHtml = sHtml & "<div class=""right"">" & rRS.RecordCount & "���q�b�g�F" & iPage & "/" & iMaxPage & "�y�[�W��</div>" & vbCrLf
	sHtml = sHtml & "<div style=""clear:both;""></div>" & vbCrLf
	sHtml = sHtml & "</div>" & vbCrLf

	GetHtmlMailPageControlParam = sHtml
End Function


'******************************************************************************
'�T�@�v�F���[���ꗗ�y�[�W�̉��̃y�[�W�R���g���[���̂g�s�l�k���擾

'******************************************************************************
Function GetHtmlMailPageControlParam2(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vMode, ByVal vPageSize, ByVal vPage, ByVal vURL)
	Dim iPage
	Dim iPageSize
	Dim iMaxPage
	Dim iStartPage
	Dim iEndPage
	Dim idxPage
	Dim sAnc
	Dim sAncRcvSnd
	Dim sHtml

	GetHtmlMailPageControlParam2 = ""
	sHtml = ""

	If GetRSState(rRS) = False Then Exit Function

	If InStr(vURL, "?") > 0 Then
		sAnc = vURL & "&amp;"
	Else
		sAnc = vURL & "?"
	End If

	If vUserType = "staff" Then
		sAncRcvSnd = HTTP_CURRENTURL & "staff/mailhistory_person.asp"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sAncRcvSnd = HTTP_CURRENTURL & "company/mailhistory_company.asp"
	End If

	iPage = CInt(vPage)
	iPageSize = CInt(vPageSize)

	rRS.PageSize = iPageSize
	iMaxPage = rRS.PageCount

	'�͈͊O�̃y�[�W�w��΍�
	If iPage > iMaxPage Then iPage = iMaxPage
	If iPage < 1 Then iPage = 1

	rRS.AbsolutePage = iPage

	'�\���J�n�y�[�W�ԍ����w��
	iStartPage = iPage - 2
	If iStartPage < 1 Then iStartPage = 1

	'�\���I���y�[�W�ԍ����w��
	iEndPage = iPage + 2
	If iEndPage - iStartPage < 4 Then iEndPage = 5
	If iEndPage > iMaxPage Then iEndPage = iMaxPage


		sHtml = sHtml & "<div>" & vbCrLf
	sHtml = sHtml & "<div class=""left"">" & vbCrLf

	If CInt(iStartPage) <> 1 Then sHtml = sHtml & "�c"
	For idxPage = iStartPage To iEndPage	'�y�[�W�ԍ���\��
		sHtml = sHtml & "�@"
		If idxPage = CInt(iPage) Then		'�w��y�[�W�̕\��
			sHtml = sHtml & "[" & idxPage & "]"
		Else
			sHtml = sHtml & "<a href=""" & sAnc & "page=" & idxPage & """>" & idxPage & "</a>"
		End If
	Next
	If iEndPage < iMaxPage Then sHtml = sHtml & "�@�c"

	sHtml = sHtml & "</div>" & vbCrLf

	
	sHtml = sHtml & "<div class=""right"">" & rRS.RecordCount & "���q�b�g�F" & iPage & "/" & iMaxPage & "�y�[�W��</div>" & vbCrLf
	sHtml = sHtml & "<div style=""clear:both;""></div>" & vbCrLf
	sHtml = sHtml & "</div>" & vbCrLf

	GetHtmlMailPageControlParam2 = sHtml
End Function

'******************************************************************************
'�T�@�v�F���[���ꗗ�y�[�W�̃��[���ꗗ���o��
'���@���FrRS			�Fup_SearchMail �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F
'�@�@�@�FvPage			�F���݂̃y�[�W
'�@�@�@�FvMode			�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvSort			�F���ёւ� ["0"]���M���~�� ["1"]���M������
'�@�@�@�FvPageSize		�F�ő�\���\���[����
'�@�@�@�FvParamToDetail	�F���[�������p�p�����[�^�i���[���ڍׂֈ����p�����߂̂��́j
'�@�@�@�FvParamToSort	�F�\�[�g�p�p�����[�^�ipage, sort �p�����[�^�����������́j
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlMailHistory(ByRef rDB, ByRef rRS, ByRef vUserType, ByVal vPage, ByVal vMode, ByVal vSort, ByVal vPageSize, ByVal vParamToDetail, ByVal vParamToSort)
	Dim iLine				'�o�͒����[����
	Dim iPageSize			'�P�y�[�W�ɏo�͂��郁�[�����̍ő�l
	Dim sEvaluationField	'�]�� ��M���[�h�̏ꍇ�F["ReceiverEvaluation"] ���M���[�h�̏ꍇ�F["SenderEvaluation"]
	Dim sRemarkField		'���l ��M���[�h�̏ꍇ�F["ReceiverRemark"] ���M���[�h�̏ꍇ�F["SenderRemark"]
	Dim sTableClass
	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

	If vUserType = "staff" Then
		sTableClass = "pattern1"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sTableClass = "pattern8"
	End If

	'���ёւ�
	Select Case vSort
		Case "0": rRS.Sort = "SendDay DESC"
		Case "1": rRS.Sort = "SendDay ASC"
		Case "2": oRS.Sort = "OrderCode DESC"
		Case "3": oRS.Sort = "OrderCode ASC"
		Case "4": oRS.Sort = "ReceiverCode DESC"
		Case "5": oRS.Sort = "ReceiverCode ASC"
		Case "6": oRS.Sort = "SenderCode DESC"
		Case "7": oRS.Sort = "SenderCode ASC"
	End Select

	iPageSize = vPageSize
	rRS.PageSize = iPageSize

	'�͈͊O�̃y�[�W�w��΍�
	If vPage > rRS.PageCount Then vPage = rRS.PageCount
	If vPage < 1 Then vPage = 1

	rRS.AbsolutePage = vPage

	sHTML = ""
	sHTML = sHTML & "<table class=""pattern1 mailHisTable smartNone"" border=""0"" cellspacing=""0"">"
	sHTML = sHTML & "<thead>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th class=""mail_delete"">�폜</th>"

	If vMode = "0" Then
		'��M�����
		sHTML = sHTML & "<th class=""mail_person"">���o�l(���R�[�h)�^����</th>"
		sHTML = sHTML & "<th class=""mail_day"">��M��" & GetSortImg(vUserType, "0", vSort, vMode, vParamToSort) & "&nbsp;" & GetSortImg(vUserType, "1", vSort, vMode, vParamToSort) & "</th>"	
	Else
		'���M�����
		sHTML = sHTML & "<th class=""mail_person"">����(���R�[�h)�^����</th>"
		sHTML = sHTML & "<th class=""mail_day"">���M��" & GetSortImg(vUserType, "0", vSort, vMode, vParamToSort) & "&nbsp;" & GetSortImg(vUserType, "1", vSort, vMode, vParamToSort) & "</th>"
	End If
	
	sHTML = sHTML & "<th class=""mail_memo"">���l</th>"
	sHTML = sHTML & "<th class=""mail_point"">�d�v�x</th>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</thead>"
	
	
	sHTML = sHTML & "<tfoot>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th class=""mail_delete"">�폜</th>"

	If vMode = "0" Then
		'��M�����
		sHTML = sHTML & "<th class=""mail_person"">���o�l(���R�[�h)�^����</th>"
		sHTML = sHTML & "<th class=""mail_day"">��M��" & GetSortImg(vUserType, "0", vSort, vMode, vParamToSort) & "&nbsp;" & GetSortImg(vUserType, "1", vSort, vMode, vParamToSort) & "</th>"	
	Else
		'���M�����
		sHTML = sHTML & "<th class=""mail_person"">����(���R�[�h)�^����</th>"
		sHTML = sHTML & "<th class=""mail_day"">���M��" & GetSortImg(vUserType, "0", vSort, vMode, vParamToSort) & "&nbsp;" & GetSortImg(vUserType, "1", vSort, vMode, vParamToSort) & "</th>"
	End If
	
	sHTML = sHTML & "<th class=""mail_memo"">���l</th>"
	sHTML = sHTML & "<th class=""mail_point"">�d�v�x</th>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</tfoot>"
	
	sHTML = sHTML & "<tbody>"

	iLine = 1
	Do While (GetRSState(rRS) = True And iLine <= vPageSize)
		If vUserType = "staff" Then sHTML = sHTML & GetHtmlMailHistoryDetailStaff(rDB, rRS, vUserType, vMode, vParamToDetail)
		If vUserType = "company" Or vUserType = "dispatch" Then sHTML = sHTML & GetHtmlMailHistoryDetailCompany(rDB, rRS, vUserType, vMode, vParamToDetail)

		iLine = iLine + 1
		rRS.MoveNext
	Loop
	

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>" & vbCrLf

	GetHtmlMailHistory = sHTML
End Function

'******************************************************************************
'�T�@�v�F�X�^�b�t�̃��[���ꗗ�y�[�W�̂P�s�g�s�l�k���擾
'���@���FrDB		�F�ڑ����̂c�a�I�u�W�F�N�g
'�@�@�@�FrRS		�Fup_SearchMail �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType	�F���O�C�������[�U���
'�@�@�@�FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlMailHistoryDetailStaff(ByRef rDB, ByRef rRS, ByRef vUserType, ByVal vMode, ByVal vParam)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim iID
	Dim sAnswerNGFlag
	Dim sSendDay		'����M��
	Dim sImgMailState	'���[���J���E�ԐM�E���J���C���[�W
	Dim sName			'�Ώۂ̊�Ɩ��E���X�Ј���
	Dim sOpenDay		'�J������
	Dim sSubject		'���[������
	Dim sAncOrder		'���l�[�ւ̃����N
	Dim sAncMailDetail	'���[���ڍׂւ̃����N
	Dim sEvaluation		'�]��
	Dim sRemark			'���l
	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function
	If vUserType <> "staff" Then Exit Function

	iID = rRS.Collect("ID")

	sSQL = "up_GetDetailMail '" & iID & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	'���[���ԐM��
	If oRS.Collect("SuspensionFlag") = "1" Or oRS.Collect("ErasureFlag") = "1" Then
		sAnswerNGFlag = "1"	'�ԐM�s��
	Else
		sAnswerNGFlag = "0"	'�ԐM��
	End If

	'���[������M��
	sSendDay = GetDateStr(oRS.Collect("SendDay"), "/") & "<br><span>" & GetTimeStr(oRS.Collect("SendDay"), ":") &"</span>"

	'��M���[�h���́A���[���̏�ԉ摜��\��
	If vMode = "0" Then
		If oRS.Collect("AnswerFlag") = "1" Then
			'�ԐM��
			sImgMailState = "<img src=""/img/common/mailre.gif"" alt=""�ԐM��"">&nbsp;"
		ElseIf ChkStr(oRS.Collect("OpenDay")) <> "" Then
			'�J����
			sImgMailState = "<img src=""/img/common/mailkai.gif"" alt=""�J����"">&nbsp;"
		Else
			'���J��
			sImgMailState = "<img src=""/img/common/mailhei.gif"" alt=""���J��"">&nbsp;"
		End If
	End If

	'�Ώۂ̊�Ɩ��E���X�Ј�����\��
	If vMode = "1" Then
		'���M�����
		sName = oRS.Collect("ReceiverCompanyName")
	Else
		'��M�����
		sName = oRS.Collect("SenderCompanyName")
	End If

	'�J������
	If ChkStr(oRS.Collect("OpenDay")) <> "" Then
		sOpenDay = "<span class=""moji812"">�y�J���F" & ChkStr(oRS.Collect("OpenDay")) & "�z</span><br>"
	End If

	'���[������
	If oRS.Collect("Subject") <> "" Then
		sSubject = oRS.Collect("Subject")
	Else
		sSubject = "�^�C�g���Ȃ�"
	End If

	'���l�[�ւ̃����N
	If Left(oRS.Collect("OrderCode"),1) = "J" Then
		sAncOrder = sAncOrder & "&nbsp;("
		If ChkOrderDsp(rDB, oRS.Collect("OrderCode"), G_USERID) = True Then
			'�f�ڒ��̋��l�[�̏ꍇ�͋��l�[�ւ̃����N
			sAncOrder = sAncOrder & "<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & oRS.Collect("OrderCode") & "</a>"
		Else
			'��f�ڂ̋��l�[�̏ꍇ�͏��R�[�h�̂�
			sAncOrder = sAncOrder & oRS.Collect("OrderCode")
		End If
		If oRS.Collect("SecretFlag") = "1" Then sAncOrder = sAncOrder & "&nbsp;<img src=""/img/order/secret.gif"" alt=""�X�J�E�g���󂯂��l�������{���ł��鋁�l���"" border=""0"">"
		sAncOrder = sAncOrder & ")"
		sAncOrder = sAncOrder & "<br>"
	Else
		sAncOrder = "<br>"
	End If

	'�]���E���l
	If vMode = "1" Then
		sEvaluation = ChkStr(oRS.Collect("SenderEvaluation"))
		sRemark = ChkStr(oRS.Collect("SenderRemark"))
	Else
		sEvaluation = ChkStr(oRS.Collect("ReceiverEvaluation"))
		sRemark = ChkStr(oRS.Collect("ReceiverRemark"))
	End If

	sAncMailDetail = HTTPS_CURRENTURL & "staff/mail_detail_person.asp"
	If vParam <> "" Then
		sAncMailDetail = sAncMailDetail & vParam & "&amp;id=" & iID
	Else
		sAncMailDetail = sAncMailDetail & "?id=" & iID
	End If

	sHTML = ""
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th class=""delCheck""><input type=""checkbox"" name=""delflag"" value=""" & iID & """></th>"
	sHTML = sHTML & "<td class=""fromWho"">" & sImgMailState & sName &  sAncOrder & sOpenDay & "<a href=""" & sAncMailDetail & """>" & sSubject & "</a></td>"
	sHTML = sHTML & "<td class=""mailDay"">" & sSendDay & "</td>"
	sHTML = sHTML & "<td class=""textareaMemo"">"
	sHTML = sHTML & "<input type=""text"" name=""CONF_Remark" & iID & """ value=""" & sRemark & """>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td class=""weightCheck"">"
	sHTML = sHTML & "<select name=""CONF_Evaluation" & iID & """>"
	sHTML = sHTML & "<option value=""""></option>"
	If sEvaluation = "A" Then: sHTML = sHTML & "<option value=""A"" selected>�`</option>": Else: sHTML = sHTML & "<option value=""A"">�`</option>": End If
	If sEvaluation = "B" Then: sHTML = sHTML & "<option value=""B"" selected>�a</option>": Else: sHTML = sHTML & "<option value=""B"">�a</option>": End If
	If sEvaluation = "C" Then: sHTML = sHTML & "<option value=""C"" selected>�b</option>": Else: sHTML = sHTML & "<option value=""C"">�b</option>": End If
	If sEvaluation = "D" Then: sHTML = sHTML & "<option value=""D"" selected>�c</option>": Else: sHTML = sHTML & "<option value=""D"">�c</option>": End If
	If sEvaluation = "E" Then: sHTML = sHTML & "<option value=""E"" selected>�d</option>": Else: sHTML = sHTML & "<option value=""E"">�d</option>": End If
	sHTML = sHTML & "</select>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	GetHtmlMailHistoryDetailStaff = sHTML
End Function

'******************************************************************************
'�T�@�v�F��Ƃ̃��[���ꗗ�y�[�W�̂P�s�g�s�l�k���擾
'���@���FrDB		�F�ڑ����̂c�a�I�u�W�F�N�g
'�@�@�@�FrRS		�Fup_SearchMail �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType	�F���O�C�������[�U���
'�@�@�@�FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvParam		�F���[�������p�����[�^
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'���@�l�F
'�X�@�V�F2007/03/02 LIS K.Kokubo �쐬
'�@�@�@�F2008/09/09 LIS K.Kokubo ���E�Җ����͑Ή�
'******************************************************************************
Function GetHtmlMailHistoryDetailCompany(ByRef rDB, ByRef rRS, ByRef vUserType, ByVal vMode, ByVal vParam)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim dbStaffName

	Dim iID
	Dim sStaffCode
	Dim sAnswerNGFlag
	Dim sSendDay		'����M��
	Dim sImgMailState	'���[���J���E�ԐM�E���J���C���[�W
	Dim sName			'�Ώۂ̊�Ɩ��E���X�Ј���
	Dim sOpenDay		'�J������
	Dim sSubject		'���[������
	Dim sAncOrder		'���l�[�ւ̃����N
	Dim sAncStaff		'���E�҃v���t�B�[���ւ̃����N
	Dim sAncProgress	'�i���󋵂ւ̃����N
	Dim sAncMailDetail	'���[���ڍׂւ̃����N
	Dim sEvaluation		'�]��
	Dim sRemark			'���l
	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function
	If Not(vUserType = "company" Or vUserType = "dispatch") Then Exit Function

	iID = rRS.Collect("ID")

	sSQL = "up_GetDetailMail '" & iID & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'���E�҃R�[�h�̎擾
	If vMode = "1" Then
		'���M���[�h�̏ꍇ�͎�M�҃R�[�h�����E�҃R�[�h
		sStaffCode = oRS.Collect("ReceiverCode")
	Else
		'��M���[�h�̏ꍇ�͑��M�҃R�[�h�����E�҃R�[�h
		sStaffCode = oRS.Collect("SenderCode")
	End If

	'���[���ԐM��
	If oRS.Collect("SuspensionFlag") = "1" Or oRS.Collect("ErasureFlag") = "1" Then
		sAnswerNGFlag = "1"	'�ԐM�s��
	Else
		sAnswerNGFlag = "0"	'�ԐM��
	End If

	'���[������M��
	sSendDay = GetDateStr(oRS.Collect("SendDay"), "/") & "<br>" & GetTimeStr(oRS.Collect("SendDay"), ":")

	'��M���[�h���́A���[���̏�ԉ摜��\��
	If vMode = "0" Then
		If oRS.Collect("AnswerFlag") = "1" Then
			'�ԐM��
			sImgMailState = "<img src=""/img/common/mailre.gif"" alt=""�ԐM��"">&nbsp;"
		ElseIf ChkStr(oRS.Collect("OpenDay")) <> "" Then
			'�J����
			sImgMailState = "<img src=""/img/common/mailkai.gif"" alt=""�J����"">&nbsp;"
		Else
			'���J��
			sImgMailState = "<img src=""/img/common/mailhei.gif"" alt=""���J��"">&nbsp;"
		End If
	End If

	'�Ώۂ̊�Ɩ��E���X�Ј�����\��
	If vMode = "1" Then
		'���M�����
		sName = oRS.Collect("ReceiverCompanyName")
	Else
		'��M�����
		sName = oRS.Collect("SenderCompanyName")
	End If

	'���l�[�ւ̃����N
	sAncOrder = "(<a href=""/order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & oRS.Collect("OrderCode") & "</a>)"

	'�J������
	If ChkStr(oRS.Collect("OpenDay")) <> "" Then
		sOpenDay = "<span class=""moji812"">�y�J���F" & ChkStr(oRS.Collect("OpenDay")) & "�z</span><br>"
	End If

	'���[������
	If oRS.Collect("Subject") <> "" Then
		sSubject = oRS.Collect("Subject")
	Else
		sSubject = "�^�C�g���Ȃ�"
	End If

	'���E�҃v���t�B�[���ւ̃����N
	If Left(sStaffCode,1) = "S" Then
		'<���E�Җ��擾>
		sSQL = "EXEC up_DtlCMPStaffName '" & G_USERID & "', '" & sStaffCode & "'"
		flgQE = QUERYEXE(dbconn, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			dbStaffName = ChkStr(oRS2.Collect("StaffName"))
		End If
		Call RSClose(oRS2)
		'</���E�Җ��擾>

		If sAnswerNGFlag = "0" Then
			'�ԐM�\
			sAncStaff = "/staff/person_detail.asp?staffcode=" & sStaffCode
			If oRS.Collect("OrderPublicFlag") = "1" Then sAncStaff = sAncStaff & "&amp;ordercode=" & oRS.Collect("OrderCode")
			sAncStaff = "<a href=""" & sAncStaff & """>"
			If dbStaffName <> "" Then
				sAncStaff = sAncStaff & dbStaffName
			Else
				sAncStaff = sAncStaff & sStaffCode
			End If
			sAncStaff = sAncStaff & "</a>"
		Else
			'�ԐM�s�̏ꍇ�͏��R�[�h�̂�
			If dbStaffName <> "" Then
				sAncStaff = dbStaffName
			Else
				sAncStaff = sStaffCode
			End If
		End If

		If vParam <> "" Then
			sAncProgress = "<a href=""" & HTTPS_CURRENTURL & "company/mailhistory_progress.asp" & vParam & "&amp;staffcode=" & sStaffCode & "&amp;ordercode=" & oRS.Collect("OrderCode") & """>�i���m�F</a><br>"
		Else
			sAncProgress = "<a href=""" & HTTPS_CURRENTURL & "company/mailhistory_progress.asp?staffcode=" & sStaffCode & "&amp;ordercode=" & oRS.Collect("OrderCode") & """>�i���m�F</a><br>"
		End If
	Else
		'��ƃR���{
		'Response.Write "<a href=""../ad/ad_detail.asp?advertcode=" & oRS.Collect("OrderCode") & """>" & oRS.Collect("OrderCode") & "</a>"
	End If

	'�]���E���l
	If vMode = "1" Then
		sEvaluation = ChkStr(oRS.Collect("SenderEvaluation"))
		sRemark = ChkStr(oRS.Collect("SenderRemark"))
	Else
		sEvaluation = ChkStr(oRS.Collect("ReceiverEvaluation"))
		sRemark = ChkStr(oRS.Collect("ReceiverRemark"))
	End If

	'���[���ڍׂւ̃����N
	sAncMailDetail = HTTPS_CURRENTURL & "company/mail_detail_company.asp"
	If vParam <> "" Then
		sAncMailDetail = sAncMailDetail & vParam & "&amp;id=" & iID
	Else
		sAncMailDetail = sAncMailDetail & "?id=" & iID
	End If

	sHTML = ""
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th><input type=""checkbox"" name=""delflag"" value=""" & iID & """></th>"
	sHTML = sHTML & "<td>" & sImgMailState & sAncStaff & "&nbsp;" & sAncOrder & "&nbsp;-&nbsp;" & sAncProgress & "<a href=""" & sAncMailDetail & """>" & sSubject & "</a></td>"
	sHTML = sHTML & "<td>" & sSendDay & "</td>"
	sHTML = sHTML & "<td>"
	sHTML = sHTML & "<input type=""text"" name=""CONF_Remark" & iID & """ value=""" & sRemark & """ size=""20"">"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td>"
	sHTML = sHTML & "<select name=""CONF_Evaluation" & iID & """>"
	sHTML = sHTML & "<option value=""""></option>"
	If sEvaluation = "A" Then: sHTML = sHTML & "<option value=""A"" selected>�`</option>": Else: sHTML = sHTML & "<option value=""A"">�`</option>": End If
	If sEvaluation = "B" Then: sHTML = sHTML & "<option value=""B"" selected>�a</option>": Else: sHTML = sHTML & "<option value=""B"">�a</option>": End If
	If sEvaluation = "C" Then: sHTML = sHTML & "<option value=""C"" selected>�b</option>": Else: sHTML = sHTML & "<option value=""C"">�b</option>": End If
	If sEvaluation = "D" Then: sHTML = sHTML & "<option value=""D"" selected>�c</option>": Else: sHTML = sHTML & "<option value=""D"">�c</option>": End If
	If sEvaluation = "E" Then: sHTML = sHTML & "<option value=""E"" selected>�d</option>": Else: sHTML = sHTML & "<option value=""E"">�d</option>": End If
	sHTML = sHTML & "</select>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	GetHtmlMailHistoryDetailCompany = sHTML
End Function

'******************************************************************************
'�T�@�v�F���[���ꗗ�y�[�W�̊Ǘ��{�^�����o��
'���@���FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlMailControl(ByVal vMode)
	Dim sHTML
	Dim sOnClickDel	'�폜�{�^�����N���b�N�����ۂ�javascript
	Dim sOnClickUpd	'�X�V�{�^�����N���b�N�����ۂ�javascript

	sOnClickDel = "if(ChkInput(getElementsByName('delflag'), 'checkbox', '1', '�폜���郁�[���Ƀ`�F�b�N���Ă��������B') == true){if(confirm('�`�F�b�N�������[�����폜���܂����H') == true){this.form.hdndelflag.value='1';this.form.submit();}}"
	sOnClickUpd = "this.form.hdnupdflag.value='1';this.form.submit();"

	sHTML = ""
	sHTML = sHTML & "<br>"
	sHTML = sHTML & "<div id=""mail_input"">"
	sHTML = sHTML & "<input type=""button"" name=""DeleteData"" value=""�I���������[�����폜"" onclick=""" & sOnClickDel & """>"
	sHTML = sHTML & "<input id=""hdndelflag"" type=""hidden"" name=""frmdelflag"" value="""">"
	sHTML = sHTML & "<div>"
	sHTML = sHTML & "<input type=""button"" name=""Update"" value=""���l�E�d�v�x���X�V"" onclick=""" & sOnClickUpd & """>"
	sHTML = sHTML & "<input id=""hdnupdflag"" type=""hidden"" name=""frmupdflag"" value="""">"
	sHTML = sHTML & "<br><span>���u���l�v�u�d�v�x�v�͂����g�̃����@�\<br>�Ƃ��Ă��g�����������܂��B</span>"
	sHTML = sHTML & "</div></div>"
	sHTML = sHTML & "<br clear=""both"">"
	

	GetHtmlMailControl = sHTML
End Function

'******************************************************************************
'�T�@�v�F���[���ꗗ�y�[�W�̃��[���������o��
'���@���FvUserID	�F���[�U�h�c
'�@�@�@�FvUserType	�F���O�C�������[�U���
'�@�@�@�FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvSort		�F���݂̕��ёւ��t���O
'�@�@�@�FrSMC		�FclsSearchMailCondition�̃C���X�^���X
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'�@�@�@�F�i�r/company/mailhistory_company.asp
'���@�l�F
'���@���F2007/03/02 LIS K.Kokubo �쐬
'�@�@�@�F2009/07/30 LIS K.Kokubo �X�J�E�g�A�v���[�`�����ǉ�
'******************************************************************************
Function GetHtmlMailSearch(ByVal vUserID, ByVal vUserType, ByVal vMode, ByVal vSort, ByRef rSMC)
	Dim sHTML
	Dim sTableClass
	Dim sChecked

	If vUserType = "staff" Then
		sTableClass = "pattern1"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sTableClass = "pattern3"
	End If

	sHTML = ""
	sHTML = sHTML & "<div id=""mail_search"" class=""smartNone"">"
	sHTML = sHTML & "<form id=""frmsearchmail"" action="""" method=""get"">"
	sHTML = sHTML & "<input type=""hidden"" name=""mode"" value=""" & vMode & """>"
	sHTML = sHTML & "<input type=""hidden"" name=""sort"" value=""" & vSort & """>"


	If vMode = "1" Then
	%>
	<h3 id="mail_under" class="smartNone">���M���[������</h3>
    <%
	ElseIf vMode = "0" Then
	%>
	<h3 id="mail_under" class="smartNone">��M���[������</h3>
    <%
    End If



	If vUserType = "company" Or vUserType = "dispatch" Then


		sHTML = sHTML & "<p>���l�S����</p>"
		sHTML = sHTML & "<select name=""mcpn"">"
		sHTML = sHTML & "<option value="""">--�S����--</option>"
		sHTML = sHTML & GetMailContactPersonNameOptionHtml(vUserID, vMode, rSMC.MailContactPersonName)
		sHTML = sHTML & "</select>"

		'<�X�J�E�g,�A�v���[�`>

		If vMode = "1" Then
			sHTML = sHTML & "<p�X�J�E�g���[��</p>"
		Else
			sHTML = sHTML & "<p>�A�v���[�`���[��</p>"
		End If


		sChecked = ""
		If rSMC.ScoutApproachFlag = "1" Then sChecked = " checked"
		sHTML = sHTML & "<label><input name=""ssaf"" type=""radio"" value=""1""" & sChecked & ">"
		If vMode = "1" Then
			sHTML = sHTML & "�X�J�E�g���[���̂�"
		Else
			sHTML = sHTML & "�A�v���[�`���[���̂�"
		End If
		sHTML = sHTML & "</label>&nbsp;"

		sChecked = ""
		If rSMC.ScoutApproachFlag = "0" Then sChecked = " checked"
		sHTML = sHTML & "<label><input name=""ssaf"" type=""radio"" value=""0""" & sChecked & ">"
		If vMode = "1" Then
			sHTML = sHTML & "�X�J�E�g���[���ȊO"
		Else
			sHTML = sHTML & "�A�v���[�`���[���ȊO"
		End If
		sHTML = sHTML & "</label>&nbsp;"

		sChecked = ""
		If rSMC.ScoutApproachFlag = "" Then sChecked = " checked"
		sHTML = sHTML & "<label><input name=""ssaf"" type=""radio"" value=""""" & sChecked & ">�w�肵�Ȃ�</label>"

		'</�X�J�E�g,�A�v���[�`>

		'<���ǃ��[��>
		If vMode = "0" Then

			sHTML = sHTML & "<p>�J����</p>"


			sChecked = ""
			If rSMC.NotOpenFlag = "1" Then sChecked = " checked"
			sHTML = sHTML & "<label><input type=""radio"" name=""snof"" value=""1""" & sChecked & ">���ǃ��[���̂�</label>&nbsp;"

			sChecked = ""
			If rSMC.NotOpenFlag = "0" Then sChecked = " checked"
			sHTML = sHTML & "<label><input type=""radio"" name=""snof"" value=""0""" & sChecked & ">���ǃ��[���̂�</label>&nbsp;"

			sChecked = ""
			If Not(rSMC.NotOpenFlag = "1" Or rSMC.NotOpenFlag = "0") Then sChecked = " checked"
			sHTML = sHTML & "<label><input type=""radio"" name=""snof"" value=""""" & sChecked & ">�w�肵�Ȃ�</label>"


		End If
		'</���ǃ��[��>
	End If


	sHTML = sHTML & "<p><b>���t</b>�Ō���</p>"
	sHTML = sHTML & "<input class=""num8"" type=""text"" name=""sdf"" maxlength=""8"" value=""" & rSMC.DayFrom & """>"
	sHTML = sHTML & "�`"
	sHTML = sHTML & "<input class=""num8"" type=""text"" name=""sdt"" maxlength=""8"" value=""" & rSMC.DayTo & """>�i��F20020101�j"
	sHTML = sHTML & "<p><b>���R�[�h</b>�Ō���</p>"
	sHTML = sHTML & "<input class=""alpha8"" type=""text"" name=""soc"" value=""" & rSMC.OrderCode & """ size=""15"">&nbsp;&nbsp;"
	
	If vUserType = "company" Or vUserType = "dispatch" Then
		If vMode = "1" Then
			sHTML = sHTML & "��M�҃R�[�h"
		Else
			sHTML = sHTML & "���M�҃R�[�h"
		End If
		sHTML = sHTML & "�F<input class=""alpha8"" type=""text"" name=""sc"" value=""" & rSMC.SearchCode & """>"
	End If
	
	sHTML = sHTML & "<p><b>�]��,����,���e,���l</b>�Ō���</p>"
	sHTML = sHTML & "<select name=""se"">"
	sHTML = sHTML & "<option value="""">--�]��--</option>"
	If rSMC.Evaluation = "A" Then: sHTML = sHTML & "<option value=""A"" selected>�`</option>": Else: sHTML = sHTML & "<option value=""A"">�`</option>": End If
	If rSMC.Evaluation = "B" Then: sHTML = sHTML & "<option value=""B"" selected>�a</option>": Else: sHTML = sHTML & "<option value=""B"">�a</option>": End If
	If rSMC.Evaluation = "C" Then: sHTML = sHTML & "<option value=""C"" selected>�b</option>": Else: sHTML = sHTML & "<option value=""C"">�b</option>": End If
	If rSMC.Evaluation = "D" Then: sHTML = sHTML & "<option value=""D"" selected>�c</option>": Else: sHTML = sHTML & "<option value=""D"">�c</option>": End If
	If rSMC.Evaluation = "E" Then: sHTML = sHTML & "<option value=""E"" selected>�d</option>": Else: sHTML = sHTML & "<option value=""E"">�d</option>": End If
	sHTML = sHTML & "</select>"
	sHTML = sHTML & "<input type=""text"" name=""skwd"" value=""" & rSMC.Keyword & """ maxlength=""50"" style=""width:300px;"">"
	sHTML = sHTML & "<br>"
	sHTML = sHTML & "<div align=""center""><input type=""submit"" value=""���̏����Ō�������""></div>"
	sHTML = sHTML & "</form>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & vbCrLf

	GetHtmlMailSearch = sHTML
End Function

'******************************************************************************
'�T�@�v�F���[�����Ȃ��ꍇ�̏o��
'���@���FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlNoMail(ByVal vUserType, ByVal vMode)
	Dim sHTML
	Dim sAnc

	If vUserType = "staff" Then
		sAnc = "./mailhistory_person.asp"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sAnc = "./mailhistory_company.asp"
	End If

	sHTML = ""
	If vMode = "1" Then
		sHTML = sHTML & "<div style=""width:100%; text-align:center;"">"
		sHTML = sHTML & "<a href=""" & sAnc & "?mode=0"">��M��</a>&nbsp;&nbsp;���M��<br>"
		sHTML = sHTML & "���M���[���͂���܂���"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & vbCrLf
	Else
		sHTML = sHTML & "<div style=""width:100%; text-align:center;"">"
		sHTML = sHTML & "��M��&nbsp;&nbsp;<a href=""" & sAnc & "?mode=1"">���M��</a><br>"
		sHTML = sHTML & "��M���[���͂���܂���"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & vbCrLf
	End If

	GetHtmlNoMail = sHTML
End Function

'******************************************************************************
'�T�@�v�F���[���ꗗ����]���Ɣ��l�̍X�V
'���@���FvUserID	�F���O�C�����̃��[�UID
'�@�@�@�FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'******************************************************************************
Function UpdMailList(ByVal vUserID, ByVal vMode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sEvaluation
	Dim sRemark
	Dim sID
	Dim idx

	sSQL = ""
	For idx = 1 To Request.Form.Count
		sID = Mid(Request.Form.Key(idx), Len("CONF_Remark") + 1)
		If IsNumber(sID, 0, False) = True Then
			sEvaluation = GetForm("CONF_Evaluation" & sID, 1)
			sRemark = GetForm("CONF_Remark" & sID, 1)
			sSQL = sSQL & _
				"EXEC sp_Reg_MailUpdate '" & sID & "', '" & vUserID & "', '" & vMode & "', '" & sEvaluation & "', '" & sRemark & "'" & vbCrLf
		End If
	Next
	If sSQL <> "" Then flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
End Function

'******************************************************************************
'�T�@�v�F���[���̍폜
'���@���FvUserID	�F���O�C�����̃��[�UID
'�@�@�@�FvDelFlag	�F�폜�Ώۃ��[��ID�@[XXXX,XXXXX,XXXXXX,�c]
'�@�@�@�FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'******************************************************************************
Function DelMailList(ByVal vUserID, ByVal vDelFlag, ByVal vMode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim aID
	Dim idx

	aID = Split(vDelFlag, ",")

	sSQL = ""
	For idx = 0 To UBound(aID)
		If IsNumber(aID(idx), 0, False) = False Then sSQL = "": Exit For
		sSQL = sSQL & "EXEC sp_Reg_MailDeleteFlag '" & aID(idx) & "', '" & vUserID & "', '" & vMode & "'" & vbCrLf
	Next
	If sSQL <> "" Then flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
End Function

'******************************************************************************
'�T�@�v�F���ёւ��̃{�^�����擾
'���@���FvUserType	�F
'�@�@�@�FvMySortNo	�F�{�^�����g�Ɋ��蓖�Ă��Ă���\�[�g�l
'�@�@�@�FvSortNo	�F���ёւ� ["0"]���M���~�� ["1"]���M������
'�@�@�@�FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvParam		�F���[�����������p�����[�^
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetSortImg(ByVal vUserType, ByVal vMySortNo, ByVal vSortNo, ByVal vMode, ByVal vParam)
	Dim sImg		'���ёւ��p�{�^���̃C���[�W�t�@�C��
	Dim sAlt		'���ёւ��悤�{�^����ALT
	Dim sOnClick	'���ёւ��p�{�^�����N���b�N�����Ƃ���javascript
	Dim sURL

	'******************************************************************************
	'���ёւ��p�̐ݒ�
	'------------------------------------------------------------------------------
	sImg = "/img/common/sort"

	If vMySortNo = vSortNo Then
		'���݂̕��ёւ������A���M�̕��ёւ����̏ꍇ�̃C���[�W
		sImg = sImg & "1"
	Else
		sImg = sImg & "4"
	End If

	If vUserType = "company" Or vUserType = "dispatch" Then
		sURL = HTTP_CURRENTURL & "company/mailhistory_company.asp"
	ElseIf vUserType = "staff" Then
		sURL = HTTP_CURRENTURL & "staff/mailhistory_person.asp"
	End If

	If CInt(vMySortNo) Mod 2 = 0 Then
		'�~��
		sImg = sImg & "1"
		sAlt = "�~��"
		If vParam <> "" Then
			sURL = sURL & vParam & "&amp;sort=0"
		Else
			sURL = sURL & "?sort=0"
		End If
	Else
		'����
		sImg = sImg & "2"
		sAlt = "����"
		If vParam <> "" Then
			sURL = sURL & vParam & "&amp;sort=1"
		Else
			sURL = sURL & "?sort=1"
		End If
	End If

	sImg = sImg & ".gif"
	'------------------------------------------------------------------------------
	'���ёւ��p�̐ݒ�
	'******************************************************************************

	If vMySortNo = vSortNo Then
		'���݂̕��ёւ������A���M�̕��ёւ����̏ꍇ�̃C���[�W
		GetSortImg = "<img src=""" & sImg & """ alt=""" & sAlt & """ border=""0"">"
	Else
		GetSortImg = "<a href=""" & sURL & """><img src=""" & sImg & """ alt=""" & sAlt & """ border=""0""></a>"
	End If
End Function
%>
