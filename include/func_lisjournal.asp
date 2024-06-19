<%
'**********************************************************************************************************************
'�T�@�v�F�k�h�r�W���[�i���Ŏg�p����֐��Q
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���[���ꗗ�y�[�W�o�͗p�@������
'�@�@�@�FChgLisJournalWorkStartDay	�F�k�h�r�W���[�i���̋Ζ��J�n�\����ϊ�
'�@�@�@�FGetHtmlJNLStaffDetail		�F�k�h�r�W���[�i���̋��E�҈�l���̃e�[�u�����擾
'�@�@�@�FGetHtmlJNLCharge			�F�k�h�r�W���[�i���̖⍇���S���g�s�l�k���擾
'�@�@�@�FGetHtmlJNLInquiryBody		�F�k�h�r�W���[�i���̖⍇�����e�g�s�l�k���擾
'�@�@�@�FGetMailBodyToCompany		�F�k�h�r�W���[�i���̖⍇���T���N�X���[��
'�@�@�@�FGetMailBodyToLis			�F�k�h�r�W���[�i���̖⍇����t�ʒm���[��
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F�k�h�r�W���[�i���̋Ζ��J�n�\����ϊ�
'���@���F
'���@�l�F
'�g�@�p�F�i�r/mailservice/lisjournal/inquiry.asp
'�X�@�V�F2007/09/10 LIS K.Kokubo �쐬
'******************************************************************************
Function ChgLisJournalWorkStartDay(ByVal vBranchCode, ByVal vWorkStartDay)
	On Error Resume Next
	Dim dWorkStartDay
	Dim sWorkStartDay

	ChgLisJournalWorkStartDay = ""

	If IsDate(vWorkStartDay) = True Then
		dWorkStartDay = CDate(vWorkStartDay)
		sWorkStartDay = Year(dWorkStartDay) & "/" & Month(dWorkStartDay) & "/" & Day(dWorkStartDay)

		If DateDiff("d", dWorkStartDay, Date) < 0 Then
			ChgLisJournalWorkStartDay = Year(dWorkStartDay) & "�N" & Month(dWorkStartDay) & "��" & Day(dWorkStartDay) & "��"
		Else
			ChgLisJournalWorkStartDay = "�����\"
		End If
	ElseIf vBranchCode = "OR" Then
		ChgLisJournalWorkStartDay = "���Г������k"
	Else
		ChgLisJournalWorkStartDay = "��������"
	End If
End Function

'******************************************************************************
'�T�@�v�F�k�h�r�W���[�i���̋��E�҈�l���̃e�[�u���s���擾
'���@���FrDB		�F�c�a�ڑ�
'�@�@�@�FrJNLStaff	�F�z�M�ԍ�
'�@�@�@�FvFlag		�F�⍇�������N�\���t���O ["1"]�\�� [<>""]��\��
'�g�@�p�F�i�r/mailservice/lisjournal/inquiry.asp
'���@�l�F
'�X�@�V�F2007/08/21 LIS K.Kokubo
'�@�@�@�F2009/01/14 LIS K.Kokubo �k�h�r�W���[�i�����E�҃N���X�������ɂ��ď����擾����悤�ɕύX
'�@�@�@�F2009/02/17 LIS K.Kokubo �ύX ��]�E��A��]�N���A�w���̒ǉ��Ή�
'******************************************************************************
Function GetHtmlJNLStaffDetail(ByRef rDB, ByRef rJNLStaff, ByVal vFlag)
	Dim sHTML
	Dim sInquiry

	If vFlag = "1" Then
		sInquiry = "<form action="""" method=""post"" style=""display:inline;""><input type=""submit"" value=""�z�M����""><input type=""hidden"" name=""frmdelstaffcode"" value=""" & rJNLStaff.StaffCode & """></form>"
	End If

	sHTML = sHTML & "<table border=""0"" style=""width:600px; border-collapse:collapse;"">"
	sHTML = sHTML & "<thead>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</thead>"
	sHTML = sHTML & "<tbody>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th colspan=""4"" style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; border-width:1px 0px 1px 0px; font-size:12px; line-height:18px; text-align:left; font-family:'�l�r �o�S�V�b�N';"">"
	If rJNLStaff.HopeJobType <> "" Then
		sHTML = sHTML & rJNLStaff.HopeJobType & "��]&nbsp;(" & rJNLStaff.StaffCode & ")<br>"
	Else
		sHTMl = sHTML & "(" & rJNLStaff.StaffCode & ")&nbsp;"
	End If

	If rJNLStaff.Age <> "" Then
		sHTML = sHTML & rJNLStaff.Age & "�E" & rJNLStaff.Sex & "�E" & rJNLStaff.Address
	Else
		sHTML = sHTML & rJNLStaff.Address
	End If

	sHTML = sHTML & "</th>"
	sHTML = sHTML & "<th colspan=""2"" style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; border-width:1px 1px 1px 0px; font-size:12px; line-height:18px; text-align:right; font-family:'�l�r �o�S�V�b�N';"">" & sInquiry & "</th>"
	sHTML = sHTML & "</tr>"

	If rJNLStaff.BranchCode <> "TG" And rJNLStaff.Age <> "" Then
		'�{�Љc�Ɠ������̏ꍇ�́u�Ζ��J�n�\��v���\��
		'�h���p�k�h�r�W���[�i���̏ꍇ�͔�\���i�N���j
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'�l�r �o�S�V�b�N';"">�Ζ��J�n�\���</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'�l�r �o�S�V�b�N';"">" & rJNLStaff.ChgWorkStartDay() & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'�l�r �o�S�V�b�N';"">���E�R�����g</th>"
	sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'�l�r �o�S�V�b�N';"">" & rJNLStaff.CounselingView & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'�l�r �o�S�V�b�N';"">�o��(�N��)</th>"
	sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'�l�r �o�S�V�b�N';"">" & rJNLStaff.CareerHistory & "</td>"
	sHTML = sHTML & "</tr>"
	If rJNLStaff.Skill <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'�l�r �o�S�V�b�N';"">�X�L��</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'�l�r �o�S�V�b�N';"">" & rJNLStaff.Skill & "</td>"
		sHTML = sHTML & "</tr>"
	End If
	If rJNLStaff.RecentConditions <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'�l�r �o�S�V�b�N';"">�A�E������</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'�l�r �o�S�V�b�N';"">" & rJNLStaff.RecentConditions & "</td>"
		sHTML = sHTML & "</tr>"
	End If
	If rJNLStaff.Hope <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'�l�r �o�S�V�b�N';"">��]�E�������</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'�l�r �o�S�V�b�N';"">" & rJNLStaff.Hope & "</td>"
		sHTML = sHTML & "</tr>"
	End If
	If rJNLStaff.HopeYearlyIncome <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'�l�r �o�S�V�b�N';"">��]�N��</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'�l�r �o�S�V�b�N';"">" & rJNLStaff.HopeYearlyIncome & "</td>"
		sHTML = sHTML & "</tr>"
	End If
	If rJNLStaff.EducateHistory <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'�l�r �o�S�V�b�N';"">�w��</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'�l�r �o�S�V�b�N';"">" & rJNLStaff.EducateHistory & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table><br>"

	GetHtmlJNLStaffDetail = sHTML
End Function

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
Function GetMailBodyToCompany(ByVal vCompanyName, ByVal vPersonName, ByVal vCertify, ByVal vBranchName, ByVal vTel, ByVal vFax)
	Dim sBody

	sBody = ""
	sBody = sBody & vCompanyName & "�@" & vPersonName & "�l" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�k�h�r�W���[�i���ւ��₢���킹�����܂��Đ��ɂ��肪�Ƃ��������܂��B" & vbCrLf
	sBody = sBody & "���ВS������̘A�������҂����������B" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�����₢���킹���������E��" & vbCrLf
	sBody = sBody & HTTP_NAVI_CURRENTURL & "mailservice/lisjournal/inquiry.asp?certify=" & vCertify & vbCrLf
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
'�@�@�@�FvPersonName		�F�⍇���S���Җ�
'�@�@�@�FvLisBranchName		�F���X�S�����_���@[1]�ǉ�
'�@�@�@�FvLisEmployeeName	�F���X�S���Җ��@[1]�ǉ�
'���@�l�F
'�g�@�p�F�i�r/mailservice/lisjournal/inquiry.asp
'�X�@�V�F2007/08/30 LIS K.Kokubo �쐬
'�@�@�@�F2007/09/14 LIS K.Kokubo [1]�@���E�ҒS������ǉ�
'******************************************************************************
Function GetMailBodyToLis(ByVal vCertify, ByVal vCompanyName, ByVal vPersonName, ByVal vLisBranchName, ByVal vLisEmployeeName, ByVal vStaffLisBranchName, ByVal vStaffLisEmployeeName)
	Dim sBody

	sBody = ""
	sBody = sBody & "�k�h�r�W���[�i������Ƃ���⍇��������܂����B" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "���⍇�����" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & vCompanyName & "(" & vPersonName & ")" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�����X�S����" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "�@��ƒS���@�F" & vLisBranchname & "(" & vLisEmployeeName & ")" & vbCrLf
	sBody = sBody & "�@���E�ҒS���F" & vStaffLisBranchname & "(" & vStaffLisEmployeeName & ")" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "���⍇�����e" & vbCrLf
	sBody = sBody & "�@�ȉ��̃����N���炨�₢���킹�̏ڍׂ����邱�Ƃ��ł��܂��B" & vbCrLf
	sBody = sBody & "�@" & HTTP_BI_CURRENTURL & "mailservice/lisjournal/inquiry/detail.asp?certify=" & vCertify & vbCrLf

	GetMailBodyToLis = sBody
End Function
%>
