<%
'**********************************************************************************************************************
'�T�@�v�F���[���ڍ׃y�[�W /staff/mailhistory_person_entity.asp
'�@�@�@�F��L�y�[�W�ŏo�͗p�̊֐��Q�����̃t�@�C���ɗp�ӂ���B
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���[���ڍ׃y�[�W�o�͗p�@������
'�@�@�@�FDspMailReturnBtn	�F�ԐM�{�^���o��
'�@�@�@�FDspMailDetail		�F���[���ڍׂ��o��
'�@�@�@�FDspNoMailDetail	�F���[���������ꍇ�̕����o��
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F�ԐM�{�^�����o��
'���@���FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvAnswerNG	�F�ԐM�� ["0"]�ԐM�� ["1"]�ԐM�s��
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/maildetail_person_entity.asp
'******************************************************************************
Function DspMailReturnBtnCustom(ByVal vMode, ByVal vAnswerNG)
	If vMode <> "1" Then
		If vAnswerNG = "0" Then
			Response.Write "<br><div align=""center""><input type=""button"" value=""�ԁ@�M"" onclick=""SendAnswer();""></div><br>"
		Else
			Response.Write "<b>���̋��l�[�͌f�ڂ��I�����Ă���A�A������邱�Ƃ��ł��܂���B</b><br><br>"
		End If
	End If
End Function

'******************************************************************************
'�T�@�v�F���[���ڍׂ��o��
'���@���FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvAnswerNG	�F�ԐM�� ["0"]�ԐM�� ["1"]�ԐM�s��
'�쐬�ҁFLis Kokubo
'�쐬���F2007/03/02
'���@�l�F
'�g�p���F�i�r/staff/maildetail_person_entity.asp
'******************************************************************************
Function DspMailDetailCustom(ByRef rRS, ByVal vMode, ByVal vAnswerNG)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
%>
<table class="pattern1" border="0" style="width:600px;">
	<thead>
	<tr>
		<th colspan="2" style="width:588px;"><%= rRS.Collect("Subject") %></th>
	</tr>
	</thead>
	<tbody>
	<tr>
<%
	If vMode = "1" Then
		'���M��ʂ̏ꍇ�͈���
%>
		<th style="width:188px;">����</th>
<%
	Else
		'�J���ς݂ɂ���
		sSQL = "sp_Reg_MailOpenDay '" & rRS.Collect("ID") & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		'��M��ʂ̏ꍇ�͍��o�l
%>
		<th style="width:188px;">���o�l</th>
<%
	End If
%>
		<td style="width:389px;">
<%
	Response.Write rRS.Collect("CompanyName")
	If sAnswerNG = "0" And Trim(rRS.Collect("OrderCode")) <> "" Then
		Response.Write "�@�@<a href=""../order/order_detail.asp?OrderCode=" & rRS.Collect("OrderCode") & """>���l�[�ڍ�</a>"
	ElseIf sAnswerNG = "1" Then
		Response.Write "�@�@" & rRS.Collect("OrderCode")
	End If
%>
		</td>
	</tr>
	<tr>
		<th>���e</th>
		<td>
			<textarea rows="20" cols="66" readonly style="height:300px;"><%= rRS.Collect("Body") %></textarea>
		</td>
	</tr>
	</tbody>
</table>
<br>
<%
End Function

Function DspNoMailDetail()
	Response.Write "<b>�w�肳�ꂽ���[���͑��݂��Ȃ����A�폜����Ă��܂��B</b><br><br>"
End Function
%>
