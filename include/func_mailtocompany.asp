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


Function GetMailBodyStaff2(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter)
	Dim sBody2
	Dim iLen
	Dim sName
	Dim sName_1
	Dim sName_2
	Dim sName_F
	Dim sName_F_1
	Dim sName_F_2
	Dim sOperateClassWebCode
	Dim sHopeWorkStartDay

	
	
	
	sSQL = "EXEC up_DtlStaff '" & G_USERID & "';"
flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
If GetRSState(oRS) = True Then
	sOperateClassWebCode = oRS.Collect("OperateClassWebCode")
	sHopeWorkStartDay = GetDateStr(oRS.Collect("HopeWorkStartDay"), "")
	sName = oRS.Collect("Name")
	If InStr(sName, "�@") <> 0 Then
		sName_1 = Mid(sName, 1, InStr(sName, "�@") - 1)
		sName_2 = Mid(sName, InStr(sName, "�@") + 1)
	Else
		sName_1 = sName
	End If
	sName_F = oRS.Collect("Name_F")
	If InStr(sName_F, "�@") <> 0 Then
		sName_F_1 = Mid(sName_F, 1, InStr(sName_F, "�@") - 1)
		sName_F_2 = Mid(sName_F, InStr(sName_F, "�@") + 1)
	Else
		sName_F_1 = sName_F
	End If
	
End If


	sBody =  sName_1 & "" & sName_2 & "�@�l" & vbCrLf & vbCrLf & _
					 "���̓x�́w�����ƃi�r�x�̓]�E���������p�����܂��āA" & vbCrLf & _
					 "�܂��Ƃɂ��肪�Ƃ��������܂��B" & vbCrLf & vbCrLf & _
					 "�ȉ��̓��e�ł����傪�����������܂����B"& vbCrLf & vbCrLf & _
					 "�S���҂�育�A�������グ�܂��̂ŁA"& vbCrLf & _
					 "�����΂炭���҂����������B"& vbCrLf
		
	'���[�����e�\������
	iLen = Len(vBody) * 1
	sBody = sBody & vbCrLf & vbCrLf & _
		"-----------------------���@��������@��-------------------------" & vbCrLf & _
		"�y���l���R�[�h�z" & vbCrLf & _
		vOrderCode & vbCrLf & vbCrLf & _
		"�y���l�y�[�W�z" & vbCrLf & _
		"http://www.shigotonavi.co.jp/order/order_detail.asp?ordercode="& vOrderCode & vbCrLf & vbCrLf & _
		"�y���[�����e�z"& vbCrLf & _
		"�^�C�g���F" & vbCrLf & _
		 vSubject & vbCrLf & vbCrLf & _
		 "�{���F" & vbCrLf & _
		 Left(vBody, iLen) &  vbCrLf & _
		"-------------------------------------------------------------" & vbCrLf & vbCrLf & _
		"�����傠�肪�Ƃ��������܂��B" & vbCrLf & vbCrLf & _
		"���R�[�h�y" & vOrderCode & "�z�ւ̂�������󂯕t���܂����B" & vbCrLf & vbCrLf & vbCrLf & _
		"�s�����Ӂt" & vbCrLf & _
		"2�T�Ԉȏ�Ԏ��������ꍇ�A������e�̏C���E����������]��������" & vbCrLf & vbCrLf & _
		"https://www.shigotonavi.co.jp/staff/access.asp" & vbCrLf &vbCrLf & _
		"�̂��⍇�킹�t�H�[�����A�y���l���R�[�h�z���L�ڂ̂������A���������B" & vbCrLf
	sBody = sBody & vFooter

	GetMailBodyStaff2 = sBody2

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
	'2009/05/01�@�������ɍ��킹�ď��������������ɂ��Ă݂��@����
		'sSignature = sSignature & "�Z���F" & oRS.Collect("Prefecture") & oRS.Collect("City") & oRS.Collect("Town") & oRS.Collect("Address") & vbCrLf
		sSignature = sSignature & "�����F" & oRS.Collect("Name") & vbCrLf
		'If oRS.Collect("HomeContactFlag") = "1" Then
		'	sSignature = sSignature & "����F" & oRS.Collect("HomeTelephoneNumber") & vbCrLf
		'End If
		'If oRS.Collect("PortableContactFlag") = "1" Then
		'	sSignature = sSignature & "�g�сF" & oRS.Collect("PortableTelephoneNumber") & vbCrLf
		'End If
		'If oRS.Collect("FaxContactFlag") = "1" Then
		'	sSignature = sSignature & "FAX �F" & oRS.Collect("FaxNumber") & vbCrLf
		'End If
		'If oRS.Collect("MailContactFlag") = "1" Then
		'	sSignature = sSignature & "Mail�F" & oRS.Collect("MailAddress") & vbCrLf
		'End If
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

'******************************************************************************
'�T�@�v�F���僁�[�����ǂ������`�F�b�N
'���@���FvText	�F���肷�镶����
'�߂�l�F
'���@�l�F�}�b�`���O�l�މ���,�K�ޑҋ@�}�b�`���O�l�މ���̔���Ɏg�p����
'���@���F2009/08/13 LIS K.Kokubo
'******************************************************************************
Function ChkMailResponse(ByVal vText)
	Dim sTrueText
	Dim sFalseText

	ChkMailResponse = False

	sTrueText = ""
	sTrueText = sTrueText & "(����)"
	sTrueText = sTrueText & "|(��x���b����������������������)" '���o�C���e���v���[�g:���l�։���
	sTrueText = sTrueText & "|(���b���f�����������܂�)" '���o�C���e���v���[�g:�ԐM(����)

	sFalseText = ""
	sFalseText = sFalseText & "(����)"
	sFalseText = sFalseText & "|(���l�ɉ��傳���Ē���)"

	If IsRE(vText, sTrueText, True) = True And IsRE(vText, sFalseText, True) = False Then 
		ChkMailResponse = True
	End If
End Function

'NEO�̃X�^�b�t�p�ԐM���[���{���쐬
Function GetMailBodyStaffNEO(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter)
	Dim sBody
	Dim iLen
	Dim sName
	Dim sName_1
	Dim sName_2
	Dim sName_F
	Dim sName_F_1
	Dim sName_F_2
	Dim sOperateClassWebCode
	Dim sHopeWorkStartDay

	
	sSQL = "EXEC up_DtlStaff '" & G_USERID & "';"
    flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
    If GetRSState(oRS) = True Then
	    sOperateClassWebCode = oRS.Collect("OperateClassWebCode")
	    sHopeWorkStartDay = GetDateStr(oRS.Collect("HopeWorkStartDay"), "")
	    sName = oRS.Collect("Name")
	    If InStr(sName, "�@") <> 0 Then
	    	sName_1 = Mid(sName, 1, InStr(sName, "�@") - 1)
	    	sName_2 = Mid(sName, InStr(sName, "�@") + 1)
	    Else
	    	sName_1 = sName
	    End If
	    sName_F = oRS.Collect("Name_F")
	    If InStr(sName_F, "�@") <> 0 Then
	    	sName_F_1 = Mid(sName_F, 1, InStr(sName_F, "�@") - 1)
	    	sName_F_2 = Mid(sName_F, InStr(sName_F, "�@") + 1)
	    Else
	    	sName_F_1 = sName_F
	    End If
	
    End If
    Call RSClose(oRS)
    
    sBody = ""
    sSQL = "EXEC sp_GetData_NEO_C_Reply '" & vOrderCode & "';"
    flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
    If GetRSState(oRS) = True Then
        If oRS.Collect("ReplyFlag") = "0" Then
            If oRS.Collect("Reply") <> "" Then
                sBody = oRS.Collect("Reply")
                sBody = sBody & vbCrLf & vbCrLf & vbCrLf
                sBody = sBody & vbCrLf & vbCrLf & _
		        "-----------------------���@��������@��-------------------------" & vbCrLf & _
		        "�y��Ɩ��z" & vbCrLf & _
		        vCompanyName & vbCrLf & vbCrLf & _
	    	    "�y���l�y�[�W�z" & vbCrLf & _
		        "http://www.shigotonavi.co.jp/order/order_detail.asp?ordercode="& vOrderCode & vbCrLf & vbCrLf & _
	    	    "�y������e�z"& vbCrLf & _
		        Left(vBody, iLen) &  vbCrLf & _
		        "-------------------------------------------------------------" & vbCrLf & vbCrLf & vbCrLf & _
                "������������������������������������������������������������������" & vbCrLf & vbCrLf & _
                "���u�����ƃi�r�v��G�|�C���g�𒙂߂悤�I" & vbCrLf & vbCrLf & _
                "------------------------------------------------------------------" & vbCrLf & vbCrLf & _
                "�u�����ƃi�r�v�ł́A���߂��̕��ЃI�t�B�X�Ŗʒk�o�^�����ꂽ��A"& vbCrLf & _
                "�Ώۂ̌f�ڋ��l�œ]�E�����܂�ƁA" & vbCrLf & _
                "�y�ő�P��G�|�C���g�z���v���[���g�������܂��I" & vbCrLf & _
                "���|�C���g�t�^�ɂ͏���������܂��B" & vbCrLf & vbCrLf & _
                "�l�X�Ȃ����p���Ώۂł��̂ł��Ѓ`�F�b�N���Ă��������B" & vbCrLf & _
                "�����p�̌�ɁyG�|�C���g�\���z�����Y��Ȃ��I" & vbCrLf & vbCrLf & _
                "�ڂ����͂�����" & vbCrLf & _
                "https://www.shigotonavi.co.jp/point/pr/"& vbCrLf & vbCrLf & _
                "������������������������������������������������������������������" & vbCrLf & vbCrLf & vbCrLf & _
		        "" & vCompanyName & "�ւ̂�������󂯕t���܂����B" & vbCrLf & vbCrLf & vbCrLf & _
		        "�����傠�肪�Ƃ��������܂����B" & vbCrLf & vbCrLf & vbCrLf & _

                "��������̃A�h���X�͑��M��p�ƂȂ��Ă���܂��B" & vbCrLf & _
                "�����Ђւ̂��A���́ulis@lis21.co.jp�v�ɂ��肢�������܂��B" & vbCrLf
                sBody = sBody & "-------------------------------" & vbCrLf
                sBody = sBody & "�͂��炭�l�̃\�[�V�����R�~���j�e�B�[�u�����ƃi�r�v" & vbCrLf
                sBody = sBody & "�^�c��ЁF���X�������" & vbCrLf
                sBody = sBody & "http://www.shigotonavi.co.jp/" & vbCrLf
                sBody = sBody & "���₢���킹�Flis@lis21.co.jp" & vbCrLf
            End If
        End If
    End If
    Call RSClose(oRS)

    if sBody = "" Then
	    sBody =  sName_1 & "" & sName_2 & "�@�l" & vbCrLf & vbCrLf & _
					 "���̓x�́w�����ƃi�r�x�̓]�E���������p�����܂��āA" & vbCrLf & _
					 "�܂��Ƃɂ��肪�Ƃ��������܂��B" & vbCrLf & vbCrLf & _
					 "�ȉ��̓��e�ł����傪�����������܂����B"& vbCrLf & vbCrLf

		sBody = sBody & "�ʏ�R���ȓ��Ɋ�ƒS���҂���A�����������܂��̂ŁA�����΂炭���҂����������B" & vbCrLf
        sBody = sBody & "�iGW��N���N�n�Ȃǂ̑�^�A�x���ɂ́A�ԓ����x���Ȃ邱�Ƃ�����܂��j" & vbCrLf & vbCrLf

        '���I�Ɋ�ƒS���҂̓d�b�ԍ��ƃ��[���A�h���X�h���C����}��
        Dim aryStrings
        sSQL = "EXEC sp_GetData_C_Contact '" & vOrderCode & "';"
        flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
        If GetRSState(oRS) = True Then
            aryStrings = Split(oRS.Collect("MailAddress"), "@")
            IF ChkStr(oRS.Collect("TelNumber")) <> "" Then
                sBody = sBody & "��Ƃ���̘A���́A�y" & oRS.Collect("TelNumber") & "�z����̂��d�b�A" & vbCrLf
            Else
                sBody = sBody & "��Ƃ���̘A���́A"
            End if
            sBody = sBody & "�y@" & aryStrings(1) & "�z����̃��[���ƂȂ�܂��B" & vbCrLf

            IF ChkStr(oRS.Collect("TelNumber")) <> "" Then
                sBody = sBody & "���M���ۂ�A���f���[�������Ƃ���Ȃ��悤�A�����ӂ��������B" & vbCrLf
            Else
                sBody = sBody & "���f���[�������Ƃ���Ȃ��悤�A�����ӂ��������B" & vbCrLf
            End if
            sBody = sBody & "�i�g�у��[���A�h���X�����o�^����Ă�����́A���[����M�ݒ��Ԃ����m�F���������j" & vbCrLf & vbCrLf
        End If
        Call RSClose(oRS)

        sBody = sBody & "�܂��A���Ђ��󋵊m�F�������Ă������������������܂��B" & vbCrLf
        sBody = sBody & "�y�u�����ƃi�r�v�G�[�W�F���gNEO�^�c�����ǁz����͂����[���ɂ́A���ԐM��������悤���肢�������܂��B"

	    '���[�����e�\������
	    iLen = Len(vBody) * 1
    	sBody = sBody & vbCrLf & vbCrLf & _
		    "-----------------------���@��������@��-------------------------" & vbCrLf & _
		    "�y��Ɩ��z" & vbCrLf & _
		    vCompanyName & vbCrLf & vbCrLf & _
	    	"�y���l�y�[�W�z" & vbCrLf & _
		    "http://www.shigotonavi.co.jp/order/order_detail.asp?ordercode="& vOrderCode & vbCrLf & vbCrLf & _
	    	"�y������e�z"& vbCrLf & _
		    Left(vBody, iLen) &  vbCrLf & _
		    "-------------------------------------------------------------" & vbCrLf & vbCrLf & _
            "�����嗚���ɂ���" & vbCrLf & _
            "�����ƃi�r�Ƀ��O�C�����AMy�y�[�W �� My���j���[ �� ����ꗗ ����m�F�o���܂��B" & vbCrLf & vbCrLf & _
                "������������������������������������������������������������������" & vbCrLf & vbCrLf & _
                "���u�����ƃi�r�v��G�|�C���g�𒙂߂悤�I" & vbCrLf & vbCrLf & _
                "------------------------------------------------------------------" & vbCrLf & vbCrLf & _
                "�u�����ƃi�r�v�ł́A���߂��̕��ЃI�t�B�X�Ŗʒk�o�^�����ꂽ��A"& vbCrLf & _
                "�Ώۂ̌f�ڋ��l�œ]�E�����܂�ƁA" & vbCrLf & _
                "�y�ő�P��G�|�C���g�z���v���[���g�������܂��I" & vbCrLf & _
                "���|�C���g�t�^�ɂ͏���������܂��B" & vbCrLf & vbCrLf & _
                "�l�X�Ȃ����p���Ώۂł��̂ł��Ѓ`�F�b�N���Ă��������B" & vbCrLf & _
                "�����p�̌�ɁyG�|�C���g�\���z�����Y��Ȃ��I" & vbCrLf & vbCrLf & _
                "�ڂ����͂�����" & vbCrLf & _
                "https://www.shigotonavi.co.jp/point/pr/"& vbCrLf & vbCrLf & _
                "������������������������������������������������������������������" & vbCrLf & vbCrLf & vbCrLf & _
		    "" & vCompanyName & "�ւ̂�������󂯕t���܂����B" & vbCrLf & vbCrLf & vbCrLf & _
		    "�����傠�肪�Ƃ��������܂����B" & vbCrLf & vbCrLf & vbCrLf & _
            "��������̃A�h���X�͑��M��p�ƂȂ��Ă���܂��B" & vbCrLf & _
            "�����Ђւ̂��A���́ulis@lis21.co.jp�v�ɂ��肢�������܂��B" & vbCrLf
	    sBody = sBody & vFooter

    End if

	GetMailBodyStaffNEO = sBody

End Function

'NEO�L���̗p�ԐM���[���{���쐬
Function GetMailBodyCompanyNEO(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter,ByVal vfrmsubject)
	Dim sBody
	Dim iLen

	sBody = ""
	'��ʂ̋��l�L��
	'��Ɩ��{���l�[�S���Җ��{�l
	sBody = vCompanyname & vbCrLf & vContactPersonName & "�l" & vbCrLf & vbCrLf

	'���[�����e�\������
	iLen = Len(vBody) * 0.3
	sBody = sBody & vbCrLf & _
		"�����b�ɂȂ��Ă���܂��B" & vbCrLf & _
        "�u�����ƃi�r�v�G�[�W�F���gNEO�����ǂł��B" & vbCrLf & vbCrLf & _
        "���f�́u�����ƃi�r�v�������p���������܂��āA" & vbCrLf & _
        "���ɂ��肪�Ƃ��������܂��B" & vbCrLf & vbCrLf & _
        "��Ђ̋��l���ɁA���E�҂��牞�傪����܂����̂ł��m�点�v���܂��B" & vbCrLf & vbCrLf & vbCrLf & _
        "�y" & vfrmsubject & "�z"  & vbCrLf & vbCrLf & _
        "����ҏ��̏ڍׂɂ��܂��ẮA" & vbCrLf & _
        "��p�̊Ǘ���ʂփ��O�C����A�B" & vbCrLf & _
        "�u����Ǘ��v�́u����҈ꗗ�v���" & vbCrLf & _
        "���m�F���������܂��B" & vbCrLf & vbCrLf & _
        "�Ǘ���ʂ͂�����" & vbCrLf & _
        "https://www.shigotonavi.co.jp/management/login.asp" & vbCrLf & vbCrLf & vbCrLf & _
        "�� ������ ��" & vbCrLf & _
        "����҂ւ̑Ή��ɂ���āA��ƃC���[�W�ɒ������邱�Ƃ��������܂��B" & vbCrLf & _
        "����҂ւ̂��₢���킹�Ή��≞��ԐM�A�ʐڑΉ��A" & vbCrLf & _
        "�̗p�̍��ےʒm�Ȃǂɂ��܂��ẮA" & vbCrLf & _
        "���ӂ����߂Đv���ɂ��Ή����������܂��悤���肢�v���܂��B" & vbCrLf & vbCrLf & vbCrLf & _
        "����Ƃ��u�����ƃi�r�v����낵�����肢�������܂��B" & vbCrLf & vbCrLf & vbCrLf & _
        "�͂��炭�l�̃\�[�V�����R�~���j�e�B�[�u�����ƃi�r�v" & vbCrLf & _
        "http://www.shigotonavi.co.jp/" & vbCrLf & vbCrLf & _
        "��������̃A�h���X�͑��M��p�ƂȂ��Ă���܂��B" & vbCrLf & _
        "�����Ђւ̂��A���́ulis@lis21.co.jp�v�ɂ��肢�������܂��B" & vbCrLf & _
        "��������������������������������������������������������������" & vbCrLf & _
        "�����s���@���X�������" & vbCrLf & _
        "��163-0825�@�����s�V�h�搼�V�h4-2-21 �V�hNS�r��25�K" & vbCrLf & _
        "���⍇�킹�Flis@lis21.co.jp" & vbCrLf & _
        "��������������������������������������������������������������" & vbCrLf & _
        "�l�b�g���[�N�F�����i�V�h�j�E�Q�n�E�É��E���É��E���E�L���E���R" & vbCrLf

	GetMailBodyCompanyNEO = sBody

End Function
%>
