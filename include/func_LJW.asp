<%
'**********************************************************************************************************************
'�T�@�v�F�k�h�r�W���[�i���Ŏg�p����֐��Q
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���[���ꗗ�y�[�W�o�͗p�@������
'�@�@�@�FGetHtmlJNLCharge			�F�k�h�r�W���[�i���̖⍇���S���g�s�l�k���擾
'�@�@�@�FGetHtmlJNLInquiryBody		�F�k�h�r�W���[�i���̖⍇�����e�g�s�l�k���擾
'�@�@�@�FGetMailBodyToCompany		�F�k�h�r�W���[�i���̖⍇���T���N�X���[��
'�@�@�@�FGetMailBodyToLis			�F�k�h�r�W���[�i���̖⍇����t�ʒm���[��
'**********************************************************************************************************************


'******************************************************************************
'�T�@�v�F�k�h�r�W���[�i���̖⍇���S���g�s�l�k���擾
'���@���FvBranchName	�F���_��
'�@�@�@�FvEmployeeName	�F�Ј���
'�@�@�@�FvTel			�F���_�d�b�ԍ�
'�@�@�@�FvFax			�F���_FAX�ԍ�
'���@�l�F
'�g�p���F�i�r/mailservice/lisjournal/inquiry.asp
'�X�@�V�F2007/08/30 LIS K.Kokubo
'******************************************************************************
Function GetHtmlJNLCharge(ByVal vBranchName, ByVal vEmployeeName, ByVal vTel, ByVal vFax)
	Dim sHTML

	sHTML = ""
	sHTML = sHTML & "<div style=""margin-top:25px;"">"
	sHTML = sHTML & "<table border=""0"" style=""width:600px;"">"
	sHTML = sHTML & "<tbody>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""width:100px; padding:5px 0px; border-bottom:1px dashed #666666;"">�S����</td>"
	sHTML = sHTML & "<td style=""width:500px; padding:5px 0px; border-bottom:1px dashed #666666;"">���X������Ё@" & vBranchName & "�@" & vEmployeeName & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""width:100px; padding:5px 0px; border-bottom:1px dashed #666666;"">�A����</td>"
	sHTML = sHTML & "<td style=""width:500px; padding:5px 0px; border-bottom:1px dashed #666666;"">TEL:" & vTel & "�@FAX:" & vFAX & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & vbCrLf

	GetHtmlJNLCharge = sHTML
End Function

'******************************************************************************
'�T�@�v�F�k�h�r�W���[�i���̖⍇�����e�g�s�l�k���擾
'���@���FvBody	�F�⍇�����e
'���@�l�F
'�g�p���F�i�r/mailservice/lisjournal/inquiry.asp
'�X�@�V�F2007/08/30 LIS K.Kokubo
'******************************************************************************
Function GetHtmlJNLInquiryBody(ByVal vBody)
	Dim sHTML

	sHTML = ""
	If Len(vBody) > 0 Then
		sHTML = sHTML & "<table class=""pattern1"" border=""0"" style=""width:600px;"">"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th>������E���v�]�Ȃ�</th>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</thead>"
		sHTML = sHTML & "<tbody>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>" & Replace(vBody, vbCrLf, "<br>") & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table><br>"
		sHTML = sHTML & vbCrLf
	End If

	GetHtmlJNLInquiryBody = sHTML
End Function

'******************************************************************************
'�T�@�v�F�k�h�r�W���[�i���̖⍇���T���N�X���[��
'���@���F
'���@�l�F
'�g�@�p�F�i�r/mailservice/lisjournal/inquiry.asp
'�X�@�V�F2007/08/30 LIS K.Kokubo �쐬
'******************************************************************************
Function GetMailBodyToCompany(ByVal vCompanyName, ByVal vCertify, ByVal vBranchName, ByVal vTel, ByVal vFax)
	Dim sBody

	sBody = ""
	sBody = sBody & vCompanyName & "�@�l" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�k�h�r�W���[�i���v�ւ��₢���킹�����܂��Đ��ɂ��肪�Ƃ��������܂��B" & vbCrLf
	sBody = sBody & "���ВS������̘A�������҂����������B" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�����₢���킹���������E��" & vbCrLf
	sBody = sBody & HTTP_NAVI_CURRENTURL & "LJW/inquiry.asp?certify=" & vCertify & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�����ВS��" & vbCrLf
	sBody = sBody & "���@���F" & vBranchName & vbCrLf
	sBody = sBody & "�A����FTEL " & vTel & "�@FAX " & vFax & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�E�]�E�T�C�g�u�����ƃi�r�v�@�@�@�@�@�@�@http://www.shigotonavi.co.jp/" & vbCrLf
	sBody = sBody & "�@�@�@�@�@�@�u�����ƃi�r���o�C���v�@�@�@http://m.shigotonavi.jp/" & vbCrLf
	sBody = sBody & "�@�l����p�T�C�g�u�����ƃi�r�l�ލ̗p�v�@http://jinzai.shigotonavi.co.jp/" & vbCrLf
	sBody = sBody & "�E�l�ޔh����l�ޏЉ�ȂǁA�l�ނɊւ���e�킲���k�����󂯂������܂��B" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "���X������Ёm�����l�ރT�[�r�X�n" & vbCrLf
	sBody = sBody & "��163-0825�@�����s�V�h�搼�V�h�Q���ڂS�ԂP�� �V�h�m�r�r���Q�T�K" & vbCrLf
	sBody = sBody & "03-5909-4120" & vbCrLf
	sBody = sBody & "lis@lis21.co.jp" & vbCrLf

	GetMailBodyToCompany = sBody
End Function

'******************************************************************************
'�T�@�v�F�k�h�r�W���[�i���̖⍇����t�ʒm���[��
'���@���FvCertify			�F�⍇���R�[�h
'�@�@�@�FvCompanyName		�F�⍇����Ɩ�
'�@�@�@�FvLisBranchName		�F���X�S�����_���@[1]�ǉ�
'�@�@�@�FvLisEmployeeName	�F���X�S���Җ��@[1]�ǉ�
'���@�l�F
'******************************************************************************
Function GetMailBodyToLis(ByVal vCertify, ByVal vCompanyName, ByVal vLisBranchName, ByVal vLisEmployeeName, ByVal vStaffLisBranchName, ByVal vStaffLisEmployeeName, _
    ByVal vDeliveryDay, _
    ByVal vCompanyCode)

	Dim sBody

	sBody = ""
	sBody = sBody & "�k�h�r�W���[�i������Ƃ���⍇��������܂����B" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "���⍇�����" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & vCompanyName & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�����X�S����" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�@��ƒS���@�F" & vLisBranchname & "(" & vLisEmployeeName & ")" & vbCrLf
	sBody = sBody & "�@���E�ҒS���F" & vStaffLisBranchname & "(" & vStaffLisEmployeeName & ")" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "���⍇�����e" & vbCrLf
	sBody = sBody & "�@�ȉ��̃����N���炨�₢���킹�̏ڍׂ����邱�Ƃ��ł��܂��B" & vbCrLf
    sBody = sBody & "�@" & HTTP_BI_CURRENTURL & "LJW/DeliveryRecord/InquiryDetail.asp?DeliveryDay=" & vDeliveryDay & "&companycode=" & vCompanyCode

	GetMailBodyToLis = sBody
End Function
%>
