<%
'**********************************************************************************************************************
'�T�@�v�F���[���e���v���[�g�Ǘ���� /mailtemplate/manager.asp
'�@�@�@�F�v���t�B�[�� /staff/person_detail.asp
'�@�@�@�F��L�y�[�W�ŏo�͗p�̊֐��Q�����̃t�@�C���ɗp�ӂ���B
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���[���e���v���[�g�Ǘ���ʏo�͗p�@������
'�@�@�@�FDspMailTemplateList				�F���[���e���v���[�g�ꗗ�����o��
'�@�@�@�FDspMailTemplateListOne				�F���[���e���v���[�g�ꗗ�̌X�̃e���v���[�g���o��
'�@�@�@�FGetMailTemplateLink				�F�Ώۋ��l���̑S���[���e���v���[�g�̕ҏW�y�[�W�ւ̃����N���擾
'�@�@�@�F
'�@�@�@�F�������@���[���e���v���[�g���́E�m�F�E�o�^������ʏo�͗p�@������
'�@�@�@�FDspInputMailTemplate				�F���[���e���v���[�g�̓��͉�ʏo��
'�@�@�@�FDspConfMailTemplate				�F���[���e���v���[�g�̓��͓��e�m�F��ʏo��
'�@�@�@�FDspRegMailTemplate					�F���[���e���v���[�g�c�a�o�^�`�o�^�����E���s��ʏo��
'�@�@�@�F
'�@�@�@�F�������@���[���e���v���[�g�Q�Ɖ�ʏo�͗p�@������
'�@�@�@�FDspMailTemplateRefList				�F���[���e���v���[�g�Q�Ɖ�ʂ̃e���v���[�g�ꗗ���o��
'�@�@�@�FDspMailTemplateRefListOne			�F���[���e���v���[�g�Q�Ɖ�ʂ̈ꗗ�̌X�̃e���v���[�g���o��
'�@�@�@�FGetMailTemplateRefLink				�F���[���e���v���[�g�Q�Ɖ�ʂ̌X�̃e���v���[�g�̓��e�m�F�̃y�[�W�ւ̃����N���擾
'�@�@�@�FDspMailTemplateRefCopy				�F���[���e���v���[�g�Q�Ɖ�ʂ̃e���v���[�g�ڍׁ��R�s�[�{�^�����o��
'�@�@�@�FGetContactPersonNameMTOptionHtml	�F���[���e���v���[�g��ێ����鋁�l�[�̒S���҈ꗗ�� <option></option> �`���Ŏ擾
'�@�@�@�F
'�@�@�@�F�������@���[���e���v���[�g�폜��ʏo�͗p�@������
'�@�@�@�FDspConfDeleteMailTemplate			�F���[���e���v���[�g�̍폜�m�F��ʏo��
'�@�@�@�FDspDeleteMailTemplate				�F���[���e���v���[�g�c�a�폜�`�폜������ʏo��
'�@�@�@�F
'�@�@�@�F�������@���[���e���v���[�g�̃R�s�[�쐬��ʏo�͗p�@������
'�@�@�@�FDspCopyMailTemplate				�F���[���e���v���[�g�R�s�[��ʂ̃R�s�[�����o��
'�@�@�@�FDspCopyMailTemplateList			�F���[���e���v���[�g�R�s�[��ʂ̈ꗗ�����o��
'�@�@�@�FDspMailTemplateListOne2			�F���[���e���v���[�g�R�s�[��ʂ̃R�s�[�拁�l�[�ꗗ���o��
'�@�@�@�FDspRegCopyMailTemplate				�F���[���e���v���[�g�R�s�[�c�a�o�^�`�o�^������ʏo��
'�@�@�@�F
'�@�@�@�F�������@���[���e���v���[�g�c�a�����@������
'�@�@�@�FRegMailTemplate					�F���[���e���v���[�g�̓o�^����
'�@�@�@�FDelMailTemplate					�F���[���e���v���[�g�̍폜����
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�ꗗ�����o��
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FvUserCode			�F���O�C�������[�U
'�@�@�@�FvContactPersonName	�F���l�S���҃t�B���^
'�@�@�@�FvPageSize			�F�P�y�[�W������̕\������
'�@�@�@�FvPage				�F�\�����y�[�W
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/manager.asp
'******************************************************************************
Function DspMailTemplateList(ByRef rDB, ByVal vUserCode, ByVal vContactPersonName, ByVal vPageSize, ByVal vPage)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sHtmlPageCtrl
	Dim iRow
	Dim sFilterPerson	'���l�S���҈ꗗ�i<option></option> ��ێ�����j

	sSQL = "up_GetMyOrder '" & vUserCode & "', '0', '', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'���l�S���҈ꗗ���擾
	If GetRSState(oRS) = True Then
		sFilterPerson = GetContactPersonNameOptionHtml(G_USERID, vContactPersonName)
	End If

	'���l�S���҂ōi����
	If GetRSState(oRS) = True And vContactPersonName <> "" Then
		If vContactPersonName <> "" Then oRS.Filter = "ContactPersonName = '" & vContactPersonName & "'"
	End If

	'�y�[�W�R���g���[���擾
	If GetRSState(oRS) = True Then
		sHtmlPageCtrl = GetPageControlHtml(rDB, oRS, vPageSize, vPage)
	End If

	Response.Write sHtmlPageCtrl
%>
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="3">���[���e���v���[�g�ꗗ</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:68px; text-align:center;">���R�[�h</th>
				<th style="width:109px;">
					<select id="contactpersonname" name="frmcontactpersonname" style="width:109px;" onchange="ChgPage(1);">
						<option value="">(���l�S����)</otpion>
						<%= sFilterPerson %>
					</select>
				</th>
				<th style="width:389px;">���[���e���v���[�g���</th>
			</tr>
<%
	iRow = 1
	If GetRSState(oRS) = True Then
		Do While GetRSState(oRS) = True And iRow <= vPageSize
			Call DspMailTemplateListOne(rDB, oRS, G_USERID)
			oRS.MoveNext
			iRow = iRow + 1
		Loop
	Else
%>
			<tr>
				<td colspan="3">���l�[������܂���B</td>
			</tr>
<%
	End If
%>
		</tbody>
	</table>
<%
	Response.Write sHtmlPageCtrl
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�ꗗ�̌X�̃e���v���[�g���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_GetMyOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/manager.asp
'******************************************************************************
Function DspMailTemplateListOne(ByRef rDB, ByRef rRS, ByVal vUserCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sContactPersonName
	Dim iMTCnt
	Dim sAnc				'���l�[�ւ̃����N

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = ChkStr(rRS.Collect("OrderCode"))
	sContactPersonName = ChkStr(rRS.Collect("ContactPersonName"))

	iMTCnt = 0
	sSQL = "up_ChkMyMailTemplate '" & vUserCode & "', '" & sOrderCode & "', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		iMTCnt = oRS.Collect("MailTemplateCnt")
	End If

	sAnc = "<a href=""" & HTTPS_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & sOrderCode & """>" & sOrderCode & "</a>"
%>
			<tr>
				<td style="text-align:center;"><%= sAnc %></td>
				<td><%= sContactPersonName %></td>
				<td>
					<%= GetMailTemplateLink(rDB, vUserCode, sOrderCode, iMTCnt) %>
				</td>
			</tr>
<%
End Function

'******************************************************************************
'�T�@�v�F�Ώۋ��l���̑S���[���e���v���[�g�̕ҏW�y�[�W�ւ̃����N���擾
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_GetMyOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvOrderCode		�F
'�@�@�@�FvMTCnt			�F���݂̃��[���e���v���[�g����
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/manager.asp
'******************************************************************************
Function GetMailTemplateLink(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vMTCnt)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim iSEQ
	Dim sMTTName
	Dim sSubject

	Dim sParam
	Dim sEditLink
	Dim sEditButton
	Dim sDelButton

	If vMTCnt < 5 Then
		GetMailTemplateLink = "<input type=""button"" value=""�V�K�쐬"" onclick=""location.href = '" & HTTPS_NAVI_CURRENTURL & "mailtemplate/regist.asp?ordercode=" & vOrderCode & "';""><br>"
	Else
		GetMailTemplateLink = "<p class=""m0"" style=""font-size:10px; color:#ff0000;"">�e���v���[�g�̐�������ɒB���Ă��܂��B</p>"
	End If

	sSQL = "up_GetListMailTemplate '" & vUserCode & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iSEQ = oRS.Collect("SEQ")
		sMTTName = ChkStr(oRS.Collect("MailTemplateTypeName"))
		sSubject = ChkStr(oRS.Collect("Subject"))
		If Len(sSubject) > 20 Then sSubject = Left(sSubject, 20) & "..."

		sParam = "?ordercode=" & vOrderCode & "&amp;seq=" & iSEQ

		sEditButton = "<input type=""button"" value=""��߰"" style=""width:35px;"" onclick=""location.href='" & HTTPS_NAVI_CURRENTURL & "mailtemplate/copy.asp" & sParam & "';"">"
		sDelButton = "<input type=""button"" value=""�폜"" style=""width:35px;"" onclick=""location.href='" & HTTPS_NAVI_CURRENTURL & "mailtemplate/delete.asp" & sParam & "';"">"

		sEditLink = "<div style=""float:left; width:100px;"">" & sMTTName & "</div>"
		sEditLink = sEditLink & "<div style=""float:left; width:230px;""><a href=""" & HTTPS_NAVI_CURRENTURL & "mailtemplate/regist.asp" & sParam & """>" & sSubject & "</a></div>"
		sEditLink = sEditLink & "<div style=""float:left; width:80px;"">" & sEditButton & "&nbsp;" & sDelButton & "</div>"
		sEditLink = sEditLink & "<div style=""clear:both;""></div>"

		GetMailTemplateLink = GetMailTemplateLink & "<div style="" border-bottom:1px dashed #ccc; margin:2px 0; padding:2px 0;"">" & sEditLink & "</div>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�̓��͉�ʏo��
'���@���FvModeText				�F"�V�K�쐬" or "�ҏW"
'�@�@�@�FvOrderCode				�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h
'�@�@�@�FvSEQ					�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h�̘A��
'�@�@�@�FvMailTemplateTypeCode	�F���[���e���v���[�g��ރR�[�h
'�@�@�@�FvSubject				�F����
'�@�@�@�FvBody					�F���e
'�@�@�@�FrErrStyle				�F�f�B�N�V���i���F���̓G���[���̃X�^�C���V�[�g��ێ���������
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DspInputMailTemplate(ByVal vModeText, ByVal vOrderCode, ByVal vSEQ, ByVal vMailTemplateTypeCode, ByVal vSubject, ByVal vBody, ByRef rErrStyle)
%>
	<form id="frmmailtemplate" action="<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/regist.asp?ordercode=<%= vOrderCode %>&amp;seq=<%= vSEQ %>" method="post">
	<input id="regmode" name="frmregmode" type="hidden" value="1">

	<p>
		<input type="button" value="�����[���e���v���[�g�̓��e���R�s�[" style="width:200px;" onclick="window.open('<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/copyreference.asp', 'mywindow6', 'width=635, height=500, menubar=no, toolbar=no, scrollbars=yes');">
		<span style="font-size:10px;">�E�E�E���̋��l�[�̃��[���e���v���[�g�̓��e���R�s�[���Ĕ��f������</span>
	</p>

	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">���[���e���v���[�g<%= vModeText %></th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">���R�[�h</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">�e���v���[�g���</th>
				<td style="<%= rErrStyle("frmmailtemplatetypecode") %>">
					<%= GetRadioHtml("MailTemplateType", vMailTemplateTypeCode, "frmmailtemplatetypecode") %>
				</td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">
					���[������<br>
					<a href="<%= HTTP_NAVI_CURRENTURL %>counter.htm" target="_brank" style="font-size:10px;">�������J�E���g</a>
				</th>
				<td><input id="subject" name="frmsubject" type="text" value="<%= vSubject %>" style="width:100%;<%= rErrStyle("frmsubject") %>"><br><p class="m0">���S�p�T�O�����ȓ�</p></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">
					���[�����e<br>
					<a href="<%= HTTP_NAVI_CURRENTURL %>counter.htm" target="_brank" style="font-size:10px;">�������J�E���g</a>
				</th>
				<td><textarea id="body" name="frmbody" style="width:100%; height:400px;<%= rErrStyle("frmbody") %>"><%= vBody %></textarea><br><p class="m0">���S�p�Q�O�O�O�����ȓ�</p></td>
			</tr>
		</tbody>
	</table>
	<br>
<%
	If vSEQ <> "" Then
		'�ҏW��ʂ̏ꍇ�́A�o�^�{�^���A�R�s�[�{�^���A�폜�{�^����z�u����B
%>
	<div style="width:600px;">
		<div align="center" style="float:left; width:150px; border-bottom:1px dotted #666666;"><input type="button" value="�m�@�F" style="width:80px;" onclick="document.forms.frmmailtemplate.submit();"></div>
		<div align="left" style="float:left; width:450px; border-bottom:1px dotted #666666;">���ҏW�������e���m�F���A�o�^���܂��B</div>
		<div style="clear:both;"></div>
	</div>
	<div style="width:600px; padding-top:10px;">
		<div align="center" style="float:left; width:150px; border-bottom:1px dotted #666666;"><input type="button" value="�R�s�[" style="width:80px;" onclick="location.href='<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/copy.asp?ordercode=<%= qsOrderCode %>&amp;seq=<%= qsSEQ %>';"></div>
		<div align="left" style="float:left; width:450px; border-bottom:1px dotted #666666;">�����̃e���v���[�g�i<span style="color:#ff0000;">�ҏW�O</span>�j���A���̋��l�[�ɃR�s�[���܂��B</div>
		<div style="clear:both;"></div>
	</div>
	<div style="width:600px;">
		<div align="center" style="float:left; width:150px; border-bottom:1px dotted #666666;"><input type="button" value="��@��" style="width:80px;" onclick="location.href='<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/delete.asp?ordercode=<%= qsOrderCode %>&amp;seq=<%= qsSEQ %>';"></div>
		<div align="left" style="float:left; width:450px; border-bottom:1px dotted #666666;">�����̃e���v���[�g���폜���܂��B</div>
		<div style="clear:both;"></div>
	</div>
	<br>
<%
	Else
		'�V�K�쐬�̏ꍇ�́A�o�^�{�^���݂̂�z�u����B
%>
	<div align="center"><input type="submit" value="�m�@�F"></div>
<%
	End If
%>
	</form>
<%
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�̓��͓��e�m�F��ʏo��
'���@���FvModeText				�F"�V�K�쐬" or "�ҏW"
'�@�@�@�FvOrderCode				�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h
'�@�@�@�FvSEQ					�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h�̘A��
'�@�@�@�FvMailTemplateTypeCode	�F���[���e���v���[�g��ރR�[�h
'�@�@�@�FvSubject				�F����
'�@�@�@�FvBody					�F���e
'�@�@�@�FrErrStyle				�F�f�B�N�V���i���F���̓G���[���̃X�^�C���V�[�g��ێ���������
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DspConfMailTemplate(ByVal vModeText, ByVal vOrderCode, ByVal vSEQ, ByVal vMailTemplateTypeCode, ByVal vSubject, ByVal vBody)
	Session("regmailtemplate") = "1"
%>
	<form id="frmmailtemplate" action="<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/regist.asp?ordercode=<%= vOrderCode %>&amp;seq=<%= vSEQ %>" method="post">
	<input id="regmode" name="frmregmode" type="hidden" value="2">
	<input id="mailtemplatetypecode" name="frmmailtemplatetypecode" type="hidden" value="<%= vMailTemplateTypeCode %>">
	<input id="subject" name="frmsubject" type="hidden" value="<%= vSubject %>">
	<input id="body" name="frmbody" type="hidden" value="<%= vBody %>">
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">���[���e���v���[�g<%= vModeText %></th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">���R�[�h</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">�e���v���[�g���</th>
				<td><%= GetDetail("MailTemplateType", vMailTemplateTypeCode) %></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">���[������</th>
				<td><%= ChgSQLtoView(vSubject) %></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">
					���[�����e
				</th>
				<td><%= ChgSQLtoView(vBody) %></td>
			</tr>
		</tbody>
	</table>
	<br>
	<div align="center"><input type="submit" value="�o�@�^"></div>
	</form>
<%
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g���͉�ʂ̃R�s�[���I�𕔕��̃e���v���[�g�ꗗ���o��
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FvUserCode			�F���O�C�������[�U
'�@�@�@�FvContactPersonName	�F���l�S���҃t�B���^
'�@�@�@�FvPageSize			�F�P�y�[�W������̕\������
'�@�@�@�FvPage				�F�\�����y�[�W
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DspMailTemplateRefList(ByRef rDB, ByVal vUserCode, ByVal vContactPersonName, ByVal vPageSize, ByVal vPage)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sHtmlPageCtrl
	Dim iRow
	Dim sFilterPerson	'���l�S���҈ꗗ�i<option></option> ��ێ�����j

	sSQL = "up_GetListMailTemplateExists '" & vUserCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'���l�S���҈ꗗ���擾
	If GetRSState(oRS) = True Then
		sFilterPerson = GetContactPersonNameMTOptionHtml(rDB, G_USERID, vContactPersonName)
	End If

	'���l�S���҂ōi����
	If GetRSState(oRS) = True And vContactPersonName <> "" Then
		If vContactPersonName <> "" Then oRS.Filter = "ContactPersonName = '" & vContactPersonName & "'"
	End If

	'�y�[�W�R���g���[���擾
	If GetRSState(oRS) = True Then
		sHtmlPageCtrl = GetPageControlHtml(rDB, oRS, vPageSize, vPage)
	End If

	Response.Write sHtmlPageCtrl
%>
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="3">�Q�Ƃ��郁�[���e���v���[�g��I�����Ă�������</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:68px; text-align:center;">���R�[�h</th>
				<th style="width:109px;">
					<select id="contactpersonname" name="frmcontactpersonname" style="width:109px;" onchange="ChgPage(1);">
						<option value="">(���l�S����)</otpion>
						<%= sFilterPerson %>
					</select>
				</th>
				<th style="width:389px;">���[���e���v���[�g���</th>
			</tr>
<%
	iRow = 1
	If GetRSState(oRS) = True Then
		Do While GetRSState(oRS) = True And iRow <= vPageSize
			Call DspMailTemplateRefListOne(rDB, oRS, G_USERID)
			oRS.MoveNext
			iRow = iRow + 1
		Loop
	Else
%>
			<tr>
				<td colspan="3">���l�[������܂���B</td>
			</tr>
<%
	End If
%>
		</tbody>
	</table>
<%
	Response.Write sHtmlPageCtrl
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�ꗗ�̌X�̃e���v���[�g���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_GetMyOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/manager.asp
'******************************************************************************
Function DspMailTemplateRefListOne(ByRef rDB, ByRef rRS, ByVal vUserCode)
	Dim sOrderCode
	Dim sContactPersonName
	Dim sAnc				'���l�[�ւ̃����N

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = ChkStr(rRS.Collect("OrderCode"))
	sContactPersonName = ChkStr(rRS.Collect("ContactPersonName"))

	sAnc = "<a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & sOrderCode & """ target=""_brank"">" & sOrderCode & "</a>"
%>
			<tr>
				<td style="text-align:center;"><%= sAnc %></td>
				<td><%= sContactPersonName %></td>
				<td>
					<%= GetMailTemplateRefLink(rDB, vUserCode, sOrderCode) %>
				</td>
			</tr>
<%
End Function

'******************************************************************************
'�T�@�v�F�Ώۋ��l���̑S���[���e���v���[�g�̕ҏW�y�[�W�ւ̃����N���擾
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_GetMyOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/manager.asp
'******************************************************************************
Function GetMailTemplateRefLink(ByRef rDB, ByVal vUserCode, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim iSEQ
	Dim sMTTName
	Dim sSubject

	Dim sParam
	Dim sEditLink

	GetMailTemplateRefLink = ""

	sSQL = "up_GetListMailTemplate '" & vUserCode & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iSEQ = oRS.Collect("SEQ")
		sMTTName = ChkStr(oRS.Collect("MailTemplateTypeName"))
		sSubject = ChkStr(oRS.Collect("Subject"))
		If Len(sSubject) > 25 Then sSubject = Left(sSubject, 25) & "..."

		sParam = "?ordercode=" & vOrderCode & "&amp;seq=" & iSEQ & "&amp;conf=1"
		sEditLink = "<div style=""float:left; width:80px;"">" & sMTTName & "</div>" & _
			"<div style=""float:left; width:309px;"">" & "<a href=""" & HTTPS_NAVI_CURRENTURL & "mailtemplate/copyreference.asp" & sParam & """>" & sSubject & "</a></div>" & _
			"<div style=""clear:both;""></div>"

		GetMailTemplateRefLink = GetMailTemplateRefLink & _
			"<div>" & sEditLink & "</div>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�̓��͓��e�m�F��ʏo��
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserCode	�F���O�C�������[�U�R�[�h
'�@�@�@�FvOrderCode	�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h
'�@�@�@�FvSEQ		�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h�̘A��
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DspMailTemplateRefCopy(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sMailTemplateTypeCode
	Dim sSubject
	Dim sBody

	DspMailTemplateRefCopy = False

	sSQL = "up_GetDetailMailTemplate '" & vUserCode & "', '" & vOrderCode & "', '" & vSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	sMailTemplateTypeCode = ChkStr(oRS.Collect("MailTemplateTypeCode"))
	sSubject = ChkStr(oRS.Collect("Subject"))
	sBody = ChkStr(oRS.Collect("Body"))
%>
	<input id="mailtemplatetypecode" name="frmmailtemplatetypecode" type="hidden" value="<%= sMailTemplateTypeCode %>">
	<input id="subject" name="frmsubject" type="hidden" value="<%= sSubject %>">
	<input id="body" name="frmbody" type="hidden" value="<%= sBody %>">
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">�R�s�[���郁�[���e���v���[�g</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">���R�[�h</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th style="text-align:center;">�e���v���[�g���</th>
				<td><%= GetDetail("MailTemplateType", sMailTemplateTypeCode) %></td>
			</tr>
			<tr>
				<th style="text-align:center;">���[������</th>
				<td><%= ChgSQLtoView(sSubject) %></td>
			</tr>
			<tr>
				<th style="text-align:center;">
					���[�����e
				</th>
				<td><%= ChgSQLtoView(sBody) %></td>
			</tr>
		</tbody>
	</table>
	<br>
	<div align="center"><input type="button" value="�R�s�[" onclick="setMailTemplateCopy();"></div>

<script type="text/javascript" language="javascript">
//<!--
function setMailTemplateCopy(){
	var oopfrm = opener.document.forms.frmmailtemplate;
	var oopmttcode = opener.document.getElementsByName('frmmailtemplatetypecode');

	//get data
	var smttcode = document.getElementById('mailtemplatetypecode').value
	var ssubject = document.getElementById('subject').value
	var sbody = document.getElementById('body').value

	//set openr
	for(var idx = 0; oopmttcode[idx] != null; idx++){
		if(oopmttcode[idx].value == smttcode){
			oopmttcode[idx].checked = true;
			break;
		}
	}
	oopfrm.subject.value = ssubject;
	oopfrm.body.value = sbody;
	close();
}
//-->
</script>
<%
	DspMailTemplateRefCopy = True
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g��ێ����鋁�l�[�̒S���҈ꗗ�� <option></option> �`���Ŏ擾
'���@���FrDB		�F�ڑ����̂c�a�I�u�W�F�N�g
'�@�@�@�FvUserCode	�F���O�C�������[�U�R�[�h
'�@�@�@�FvPersonName�F�i�荞�ދ��l�[�̒S���Җ�
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function GetContactPersonNameMTOptionHtml(ByRef rDB, ByVal vUserCode, ByVal vPersonName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	GetContactPersonNameMTOptionHtml = ""

	sSQL = "up_GetListContactPersonNameMT '" & vUserCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		If oRS.Collect("PersonName") = vPersonName Then
			GetContactPersonNameMTOptionHtml = GetContactPersonNameMTOptionHtml & _
				"<option value=""" & oRS.Collect("PersonName") & """ selected>" & oRS.Collect("PersonName") & "</option>"
		Else
			GetContactPersonNameMTOptionHtml = GetContactPersonNameMTOptionHtml & _
				"<option value=""" & oRS.Collect("PersonName") & """>" & oRS.Collect("PersonName") & "</option>"
		End If
		oRS.MoveNext
	Loop
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�c�a�o�^�`�o�^�����E���s��ʏo��
'���@���FvModeText				�F"�V�K�쐬" or "�ҏW"
'�@�@�@�FvUserCode				�F���O�C�������[�U�R�[�h
'�@�@�@�FvOrderCode				�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h
'�@�@�@�FvSEQ					�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h�̘A��
'�@�@�@�FvMailTemplateTypeCode	�F���[���e���v���[�g��ރR�[�h
'�@�@�@�FvSubject				�F����
'�@�@�@�FvBody					�F���e
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DspRegMailTemplate(ByRef rDB, ByVal vModeText, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ, ByVal vMailTemplateTypeCode, ByVal vSubject, ByVal vBody)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim flgReg

	flgReg = False
	If Session("regmailtemplate") = "1" Then
		'�o�^����
		flgReg = RegMailTemplate(rDB, vUserCode, vOrderCode, vSEQ, vMailTemplateTypeCode, vSubject, vBody)
	End If
	Session.Contents.Remove("regmailtemplate")

	If flgReg = True Then
		Response.Write "<p><b>���[���e���v���[�g��o�^���܂����B</b></p>"
	Else
		Response.Write "<p><b>���[���e���v���[�g�̓o�^��<span style=""color:#ff0000;"">���s</span>���܂����B</b></p>"
	End If

	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & vOrderCode & """>���l�[�ڍ׃y�[�W��</a></p>"
	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "mailtemplate/manager.asp"">���[���e���v���[�g�Ǘ��y�[�W��</a></p>"
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�̓��͓��e�m�F��ʏo��
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserCode	�F���O�C�������[�U�R�[�h
'�@�@�@�FvOrderCode	�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h
'�@�@�@�FvSEQ		�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h�̘A��
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DspConfDeleteMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sMailTemplateTypeCode
	Dim sSubject
	Dim sBody

	DspConfDeleteMailTemplate = False

	sSQL = "up_GetDetailMailTemplate '" & vUserCode & "', '" & vOrderCode & "', '" & vSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	sMailTemplateTypeCode = ChkStr(oRS.Collect("MailTemplateTypeCode"))
	sSubject = ChkStr(oRS.Collect("Subject"))
	sBody = ChkStr(oRS.Collect("Body"))
%>
	<form id="frmmailtemplate" action="<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/delete.asp?ordercode=<%= vOrderCode %>&amp;seq=<%= vSEQ %>" method="post">
	<input id="regmode" name="frmregmode" type="hidden" value="1">
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">���[���e���v���[�g�폜�m�F</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">���R�[�h</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th style="text-align:center;">�e���v���[�g���</th>
				<td><%= GetDetail("MailTemplateType", sMailTemplateTypeCode) %></td>
			</tr>
			<tr>
				<th style="text-align:center;">���[������</th>
				<td><%= sSubject %></td>
			</tr>
			<tr>
				<th style="text-align:center;">
					���[�����e<br>
				</th>
				<td><%= Replace(sBody, vbCrLf, "<br>") %></td>
			</tr>
		</tbody>
	</table>
	<br>
	<div align="center"><input type="submit" value="��@��"></div>
	</form>
<%
	DspConfDeleteMailTemplate = True
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�̃R�s�[�����o��
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FvUserCode			�F���O�C�������[�U
'�@�@�@�FvContactPersonName	�F���l�S���҃t�B���^
'�@�@�@�FvPageSize			�F�P�y�[�W������̕\������
'�@�@�@�FvPage				�F�\�����y�[�W
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/18
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/manager.asp
'******************************************************************************
Function DspCopyMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sMailTemplateTypeCode
	Dim sSubject
	Dim sBody

	DspCopyMailTemplate = False

	sSQL = "up_GetDetailMailTemplate '" & vUserCode & "', '" & vOrderCode & "', '" & vSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	sMailTemplateTypeCode = ChkStr(oRS.Collect("MailTemplateTypeCode"))
	sSubject = ChkStr(oRS.Collect("Subject"))
	sBody = ChkStr(oRS.Collect("Body"))
%>
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">�R�s�[�����[���e���v���[�g</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">���R�[�h</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th style="text-align:center;">�e���v���[�g���</th>
				<td><%= GetDetail("MailTemplateType", sMailTemplateTypeCode) %></td>
			</tr>
			<tr>
				<th style="text-align:center;">���[������</th>
				<td><%= sSubject %></td>
			</tr>
			<tr>
				<th style="text-align:center;">
					���[�����e<br>
				</th>
				<td><%= Replace(sBody, vbCrLf, "<br>") %></td>
			</tr>
		</tbody>
	</table>
	<br>
<%
	DspCopyMailTemplate = True
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�R�s�[��ʂ̈ꗗ�����o��
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FvUserCode			�F���O�C�������[�U
'�@�@�@�FvContactPersonName	�F���l�S���҃t�B���^
'�@�@�@�FvPageSize			�F�P�y�[�W������̕\������
'�@�@�@�FvPage				�F�\�����y�[�W
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/manager.asp
'******************************************************************************
Function DspCopyMailTemplateList(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ, ByVal vPageSize, ByVal vPage, ByVal vContactPersonName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dicOrderDetail	'�f�B�N�V���i���F���l�[�ڍ�
	Dim sOrderCode		'���R�[�h
	Dim sFilterPerson	'���l�[�ꗗ�̃��R�[�h�Z�b�g�����l�S���҂ōi�荞�ރt�B���^�[
	Dim sHtmlPageCtrl	'�y�[�W�R���g���[����HTML
	Dim iRow

	sSQL = "up_GetMyOrder '" & vUserCode & "', '0', '', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'���l�S���҈ꗗ���擾
	If GetRSState(oRS) = True Then
		sFilterPerson = GetContactPersonNameMTOptionHtml(rDB, G_USERID, vContactPersonName)
	End If

	'���l�S���҂ōi����
	If GetRSState(oRS) = True And vContactPersonName <> "" Then
		If vContactPersonName <> "" Then oRS.Filter = "ContactPersonName = '" & vContactPersonName & "'"
	End If

	'�y�[�W�R���g���[���擾
	If GetRSState(oRS) = True Then
		sHtmlPageCtrl = GetPageControlHtml(rDB, oRS, vPageSize, vPage)
		sHtmlPageCtrl = "<div style=""border-top:1px dotted #666666; border-bottom:1px dotted #666666;"">" & sHtmlPageCtrl & "</div>"
	End If
%>
<select id="contactpersonname" name="frmcontactpersonname" style="margin-bottom:5px;" onchange="this.form.frmpage.value='1'; ChgPage(1);">
	<option value="">----- �R�s�[�拁�l�[�̒S���� -----</otpion>
	<%= sFilterPerson %>
</select>
<%

	If GetRSState(oRS) = True Then
		iRow = 1
		Response.Write sHtmlPageCtrl

		Do While GetRSState(oRS) = True And iRow <= 10
			sOrderCode = oRS.Collect("OrderCode")
			Set dicOrderDetail = GetDicOrderDetail(rDB, sOrderCode)
			Call DspMailTemplateListOne2(dicOrderDetail, vOrderCode, vSEQ)
			oRS.MoveNext
			iRow = iRow + 1
			Set dicOrderDetail = Nothing
		Loop

		Response.Write sHtmlPageCtrl
	Else
		Response.Write "<p>���l�[������܂���B</p>"
	End If

	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�R�s�[��ʂ̃R�s�[�拁�l�[�ꗗ���o��
'���@���FrDic		�F
'�@�@�@�FvOrderCode	�F
'�@�@�@�FvSEQ		�F
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/18
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/manager.asp
'******************************************************************************
Function DspMailTemplateListOne2(ByRef rDic, ByVal vOrderCode, ByVal vSEQ)
	Dim sSalary

	DspMailTemplateListOne2 = False

	If IsObject(rDic) = False Then Exit Function
	If Len(rDic("OrderCode")) = 0 Then Exit Function

	sSalary = ""
%>
<table class="cw" border="0" style="margin:2px 0px;">
	<tbody>
	<tr>
		<td style="width:125px; padding-right:5px; vertical-align:top;">
			��&nbsp;
<%
	If rDic("MailTemplateCnt") < 5 Then
%>
			<a href="<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/copy.asp?ordercode=<%= vOrderCode %>&amp;seq=<%= vSEQ %>&amp;copyto=<%= rDic("OrderCode") %>">���̋��l�[�փR�s�[</a>
<%
	Else
%>
			<span style="font-size:10px; color:#ff0000;">�������ɒB���Ă��܂�</span>
<%
	End If
%>
		</td>
		<td style="width:480px; vertical-align:top;"><%= rDic("OrderCode") & "&nbsp;�F&nbsp;" & rDic("JobTypeDetail") & rDic("WorkingType") %></td>
	</tr>
	</tbody>
</table>
<%
	DspMailTemplateListOne2 = True
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�R�s�[�c�a�o�^�`�o�^������ʏo��
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserCode	�F���O�C�������[�U�R�[�h
'�@�@�@�FvOrderCode	�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h
'�@�@�@�FvSEQ		�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h�̘A��
'�@�@�@�FvCopyTo	�F�R�s�[����R�[�h
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DspRegCopyMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ, ByVal vCopyTo)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim flgReg

	sSQL = "up_GetDetailMailTemplate '" & G_USERID & "', '" & qsOrderCode & "', '" & qsSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sMailTemplateTypeCode = oRS.Collect("MailTemplateTypeCode")
		sSubject = oRS.Collect("Subject")
		sBody = oRS.Collect("Body")
	End If

	flgReg = False
	If Session("regmailtemplate") = "1" Then
		'�o�^����
		flgReg = RegMailTemplate(rDB, vUserCode, vCopyTo, "", sMailTemplateTypeCode, sSubject, sBody)
	End If
	Session.Contents.Remove("regmailtemplate")

	If flgReg = True Then
		Response.Write "<p><b>���[���e���v���[�g���R�s�[���܂����B</b></p>"
	Else
		Response.Write "<p><b>���[���e���v���[�g�̃R�s�[��<span style=""color:#ff0000;"">���s</span>���܂����B</b></p>"
	End If

	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & qsOrderCode & """>���l�[�ڍ׃y�[�W��</a></p>"
	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "mailtemplate/manager.asp"">���[���e���v���[�g�Ǘ��y�[�W��</a></p>"
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�c�a�폜�`�폜������ʏo��
'���@���FrDB		�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvUserCode	�F���O�C�������[�U�R�[�h
'�@�@�@�FvOrderCode	�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h
'�@�@�@�FvSEQ		�F���[���e���v���[�g�쐬�Ώۂ̏��R�[�h�̘A��
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DspDeleteMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim flgDel

	flgDel = DelMailTemplate(rDB, vUserCode, vOrderCode, vSEQ)

	If flgDel = True Then
		Response.Write "<p><b>���[���e���v���[�g���폜���܂����B</b></p>"
	Else
		Response.Write "<p><b>���[���e���v���[�g�̍폜��<span style=""color:#ff0000;"">���s</span>���܂����B</b></p>"
	End If

	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & vOrderCode & """>���l�[�ڍ׃y�[�W��</a></p>"
	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "mailtemplate/manager.asp"">���[���e���v���[�g�Ǘ��y�[�W��</a></p>"

	DspDeleteMailTemplate = True
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�̓o�^����
'���@���FrDB					�F�ڑ�����DBConnection
'�@�@�@�FvUserCode				�F
'�@�@�@�FvOrderCode				�F
'�@�@�@�FvSEQ					�F
'�@�@�@�FvMailTemplateTypeCode	�F
'�@�@�@�FvSubject				�F
'�@�@�@�FsBody					�F
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function RegMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ, ByVal vMailTemplateTypeCode, ByVal vSubject, ByVal vBody)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim flgData

	RegMailTemplate = False

	sSQL = "up_Reg_C_MailTemplate '" & vOrderCode & "', '" & vSEQ & "', '" & vMailTemplateTypeCode & "', '" & ChkSQLStr(vSubject) & "', '" & ChkSQLStr(vBody) & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If flgQE = True Then RegMailTemplate = True
End Function

'******************************************************************************
'�T�@�v�F���[���e���v���[�g�̍폜����
'���@���FrDB					�F�ڑ�����DBConnection
'�@�@�@�FvUserCode				�F
'�@�@�@�FvOrderCode				�F
'�@�@�@�FvSEQ					�F
'�@�@�@�FvMailTemplateTypeCode	�F
'�@�@�@�FvSubject				�F
'�@�@�@�FsBody					�F
'�쐬�ҁFLis Kokubo
'�쐬���F2007/06/15
'���@�l�F
'�g�p���F�����ƃi�r/mailtemplate/regist.asp
'******************************************************************************
Function DelMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim flgData

	DelMailTemplate = False

	sSQL = "up_Del_C_MailTemplate '" & vUserCode & "', '" & vOrderCode & "', '" & vSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If flgQE = True Then DelMailTemplate = True
End Function
%>
