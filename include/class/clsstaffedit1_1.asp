<%
'******************************************************************************
'�T�@�v�F��{���i/staff/edit1_1.asp�j�o�^�p�̃N���X
'���@�l�F�������o�֐�
'�@�@�@�FSetData
'�@�@�@�FChkData
'�@�@�@�FGetRegSQL
'�@�@�@�FDiffData
'�X�@�V�F2008/04/22 LIS K.Kokubo
'******************************************************************************
Class clsStaffEdit1_1
	'�o�^�p�f�[�^
	Public Name_1					'��
	Public Name_2					'��
	Public Name_F_1					'�Z�C
	Public Name_F_2					'���C
	Public Birthday					'�a����
	Public Sex						'����
	Public Post_U					'�Z���F�X�֔ԍ���R��
	Public Post_L					'�Z���F�X�֔ԍ����S��
	Public PrefectureCode			'�Z���F�s���{���R�[�h
	Public PrefectureName			'�Z���F�s���{����
	Public City						'�Z���F�s��S
	Public City_F					'�Z���F�s��S�J�i
	Public Town						'�Z���F����
	Public Town_F					'�Z���F�����J�i
	Public Address					'�Z���F�Ԓn�Ȃ�
	Public Address_F				'�Z���F�Ԓn�ȂǃJ�i
	Public HomeTelephoneNumber		'��TEL
	Public PortableTelephoneNumber	'�g��
	Public FaxNumber				'FAX
	Public MailAddress				'�o�b���[���A�h���X
	Public MailAddress2				'�o�b���[���A�h���X�m�F
	Public PortableMailAddress		'�g�у��[���A�h���X
	Public PortableMailAddress2		'�g�у��[���A�h���X�m�F
	Public HomeContactFlag			'��]�A����t���O�F��TEL
	Public PortableContactFlag		'��]�A����t���O�F�g��
	Public FaxContactFlag			'��]�A����t���O�FFAX
	Public MailContactFlag			'��]�A����t���O�F���[��
	Public NoticeMailFlag			'���[���A����t���O
	Public UrgencyPost_U			'�ً}�A����F�X�֔ԍ���R��
	Public UrgencyPost_L			'�ً}�A����F�X�֔ԍ����S��
	Public UrgencyAddress			'�ً}�A����F�Z��
	Public UrgencyAddress_F			'�ً}�A����F�Z���J�i
	Public UrgencyTelephoneNumber	'�ً}�A����FTEL
	Public URL						'�z�[���y�[�W
	'�G���[�����p
	Public Err							'�G���[����
	Public ErrName_1					'��
	Public ErrName_2					'��
	Public ErrName_F_1					'�Z�C
	Public ErrName_F_2					'���C
	Public ErrBirthday					'�a����
	Public ErrSex						'����
	Public ErrPost_U					'�Z���F�X�֔ԍ���R��
	Public ErrPost_L					'�Z���F�X�֔ԍ����S��
	Public ErrPrefectureCode			'�Z���F�s���{���R�[�h
	Public ErrCity						'�Z���F�s��S
	Public ErrCity_F					'�Z���F�s��S�J�i
	Public ErrTown						'�Z���F����
	Public ErrTown_F					'�Z���F�����J�i
	Public ErrAddress					'�Z���F�Ԓn�Ȃ�
	Public ErrAddress_F					'�Z���F�Ԓn�ȂǃJ�i
	Public ErrHomeTelephoneNumber		'��TEL
	Public ErrPortableTelephoneNumber	'�g��
	Public ErrFaxNumber					'FAX
	Public ErrMailAddress				'�o�b���[���A�h���X
	Public ErrMailAddress2				'�o�b���[���A�h���X�m�F
	Public ErrPortableMailAddress		'�g�у��[���A�h���X
	Public ErrPortableMailAddress2		'�g�у��[���A�h���X�m�F
	Public ErrHomeContactFlag			'��]�A����t���O�F��TEL
	Public ErrPortableContactFlag		'��]�A����t���O�F�g��
	Public ErrFaxContactFlag			'��]�A����t���O�FFAX
	Public ErrMailContactFlag			'��]�A����t���O�F���[��
	Public ErrNoticeMailFlag			'���[���A����t���O
	Public ErrUrgencyPost_U				'�ً}�A����F�X�֔ԍ���R��
	Public ErrUrgencyPost_L				'�ً}�A����F�X�֔ԍ����S��
	Public ErrUrgencyAddress			'�ً}�A����F�Z��
	Public ErrUrgencyAddress_F			'�ً}�A����F�Z���J�i
	Public ErrUrgencyTelephoneNumber	'�ً}�A����FTEL
	Public ErrURL						'�z�[���y�[�W

	'******************************************************************************
	'�T�@�v�F���̓f�[�^�`�F�b�N
	'���@���F
	'���@�l�F
	'�X�@�V�F2008/04/22 LIS K.Kokubo
	'******************************************************************************
	Public Function SetData(ByVal vStaffCode)
		Dim sSQL
		Dim oRS
		Dim sError
		Dim flgQE

		Dim dbName
		Dim dbName_F

		sSQL = "sp_GetDetailStaff '" & vStaffCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbName = oRS.Collect("Name")
			If InStr(dbName, "�@") <> 0 Then
				Name_1 = Mid(dbName, 1, InStr(dbName, "�@") - 1)
				Name_2 = Mid(dbName, InStr(dbName, "�@") + 1)
			Else
				Name_1 = dbName
			End If

			dbName_F = oRS.Collect("Name_F")
			If InStr(dbName_F, "�@") <> 0 Then
				Name_F_1 = Mid(dbName_F, 1, InStr(dbName_F, "�@") - 1)
				Name_F_2 = Mid(dbName_F, InStr(dbName_F, "�@") + 1)
			Else
				Name_F_1 = dbName_F
			End If

			Post_U = oRS.Collect("Post_U")
			Post_L = oRS.Collect("Post_L")
			Birthday = GetDateStr(oRS.Collect("Birthday"), "")
			Sex = oRS.Collect("SexType")
			PrefectureCode = oRS.Collect("PrefectureCode")
			PrefectureName = oRS.Collect("PrefectureName")
			City = oRS.Collect("City")
			City_F = oRS.Collect("City_F")
			Town = oRS.Collect("Town")
			Town_F = oRS.Collect("Town_F")
			Address = oRS.Collect("Address")
			Address_F = oRS.Collect("Address_F")
			HomeTelephoneNumber = oRS.Collect("HomeTelephoneNumber")
			PortableTelephoneNumber = oRS.Collect("PortableTelephoneNumber")
			FaxNumber = oRS.Collect("FaxNumber")
			MailAddress = oRS.Collect("MailAddress")
			MailAddress2 = oRS.Collect("MailAddress")
			PortableMailAddress = oRS.Collect("PortableMailAddress")
			PortableMailAddress2 = oRS.Collect("PortableMailAddress")
			HomeContactFlag = oRS.Collect("HomeContactFlag")
			PortableContactFlag = oRS.Collect("PortableContactFlag")
			FaxContactFlag = oRS.Collect("FaxContactFlag")
			MailContactFlag = oRS.Collect("MailContactFlag")
			NoticeMailFlag = ChkStr(oRS.Collect("NoticeMailFlag"))
			UrgencyPost_U = oRS.Collect("UrgencyPost_U")
			UrgencyPost_L = oRS.Collect("UrgencyPost_L")
			UrgencyAddress = oRS.Collect("UrgencyAddress")
			UrgencyAddress_F = oRS.Collect("UrgencyAddress_F")
			UrgencyTelephoneNumber = oRS.Collect("UrgencyTelephoneNumber")
			URL = oRS.Collect("URL")
		End If
		Call RSClose(oRS)
	End Function

	'******************************************************************************
	'�T�@�v�F���̓f�[�^�`�F�b�N
	'���@���F
	'���@�l�F
	'�X�@�V�F2008/04/22 LIS K.Kokubo
	'******************************************************************************
	Public Function ChkData()
		Dim sStyle
		Dim flgReg

		sStyle = "background-color:#ffffcc;"
		flgReg = True
		Err = ""

		'���`�F�b�N
		If Name_1 = "" Then
			Err = Err & "�u���v�͕K�{�ł��B�S�p�œ��͂��Ă��������B<br>"
			ErrName_1 = sStyle
			flgReg = False
		ElseIf IsRE(Name_1, "[\w\f\n\r\t\v,.*@!""#$%&'()-=~|\[\]\\?+;:{}]", False) = True Or ChkLen(Name_1, 40) = False Then
			Err = Err & "�u���v�͑S�p�łQ�O�����ȓ��œ��͂��Ă��������B<br>"
			ErrName_1 = sStyle
			flgReg = False
		End If

		'���`�F�b�N
		If Name_2 = "" Then
			Err = Err & "�u���v�͕K�{�ł��B�S�p�œ��͂��Ă��������B<br>"
			ErrName_2 = sStyle
			flgReg = False
		ElseIf IsRE(Name_2, "[\w\f\n\r\t\v,.*@!""#$%&'()-=~|\[\]\\?+;:{}]", False) = True Or ChkLen(Name_2, 40) = False Then
			Err = Err & "�u���v�͑S�p�łQ�O�����ȓ��œ��͂��Ă��������B<br>"
			ErrName_2 = sStyle
			flgReg = False
		End If

		'�Z�C�`�F�b�N
		If Name_F_1 = "" Then
			Err = Err & "�u�Z�C�v�͕K�{�ł��B�S�p�J�i�œ��͂��Ă��������B<br>"
			ErrName_F_1 = sStyle
			flgReg = False
		ElseIf ChkKana(Name_F_1) = False Or ChkLen(Name_F_1, 40) = False Then
			Err = Err & "�u�Z�C�v�͑S�p�J�i�œ��͂��Ă��������B<br>"
			ErrName_F_1 = sStyle
			flgReg = False
		End If

		'���C�`�F�b�N
		If Name_F_2 = "" Then
			Err = Err & "�u���C�v�͕K�{�ł��B�S�p�J�i�œ��͂��Ă��������B<br>"
			ErrName_F_2 = sStyle
			flgReg = False
		ElseIf ChkKana(Name_F_2) = False Or ChkLen(Name_F_2, 40) = False Then
			Err = Err & "�u���C�v�͑S�p�J�i�œ��͂��Ă��������B<br>"
			ErrName_F_2 = sStyle
			flgReg = False
		End If

		'���ʃ`�F�b�N
		If Not(Sex = "1" Or Sex = "2") Then
			Err = Err & "���ʂ��`�F�b�N���Ă��������B<br>"
			ErrSex = sStyle
			flgReg = False
		End If

		'�X�֔ԍ��`�F�b�N
		If Post_U & Post_L = "" Then
			Err = Err & "���Z���̗X�֔ԍ�����͂��Ă��������B<br>"
			ErrPost_U = sStyle
			ErrPost_L = sStyle
			flgReg = False
		ElseIf IsRE(Post_U & Post_L, "^\d\d\d\d\d\d\d$", False) = True Then 
			sSQL = "/* ���E�Ҋ�{���ҏW���̗X�֔ԍ��`�F�b�N */ "
			sSQL = "EXEC up_DtlZip '" & Post_U & "', '" & Post_L & "'"

			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = False Then
				Err = Err & "���Z���̗X�֔ԍ��͑��݂��܂���B<br>"
				ErrPost_U = sStyle
				ErrPost_L = sStyle
				flgReg = False
			End If
			Call RSClose(oRS)
		Else
			Err = Err & "���Z���̗X�֔ԍ��͔��p�����œ��͂��Ă��������B<br>"
			ErrPost_U = sStyle
			ErrPost_L = sStyle
			flgReg = False
		End If

		'�s���{���R�[�h
		If IsRE(PrefectureCode, "^\d\d\d$", False) = False Then
			Err = Err & "�s���{����I�����Ă��������B<br>"
			ErrPrefectureCode = sStyle
			flgReg = False
		End If

		'�s��S
		If City = "" Then
			Err = Err & "�s��S�͕K�{�ł��B�S�p�T�O�����ȓ��œ��͂��Ă��������B<br>"
			ErrCity = sStyle
			flgReg = False
		ElseIf ChkLen(City, 100) = False Then
			Err = Err & "�s��S�̕��������������𒴂��Ă��܂��B�S�p�T�O�����ȓ��œ��͂��Ă��������B<br>"
			ErrCity = sStyle
			flgReg = False
		End If

		'�s��S�J�i
		If City_F <> "" Then
			If ChkLen(City_F, 100) = False Then
				Err = Err & "�s��S�J�i�̕��������������𒴂��Ă��܂��B�S�p�T�O�����ȓ��œ��͂��Ă��������B<br>"
				ErrCity_F = sStyle
				flgReg = False
			End If
		End If

		'����
		If Town <> "" Then
			If ChkLen(Town, 100) = False Then
				Err = Err & "�����̕��������������𒴂��Ă��܂��B�S�p�T�O�����ȓ��œ��͂��Ă��������B<br>"
				ErrTown = sStyle
				flgReg = False
			End If
		End If

		'�����J�i
		If Town_F <> "" Then
			If ChkLen(Town_F, 100) = False Then
				Err = Err & "�����J�i�̕��������������𒴂��Ă��܂��B�S�p�T�O�����ȓ��œ��͂��Ă��������B<br>"
				ErrTown_F = sStyle
				flgReg = False
			End If
		End If

		'�Ԓn��
		If Address <> "" Then
			If ChkLen(Address, 100) = False Then
				Err = Err & "�Ԓn���̕��������������𒴂��Ă��܂��B�S�p�T�O�����ȓ��œ��͂��Ă��������B<br>"
				ErrAddress = sStyle
				flgReg = False
			End If
		End If

		'�Ԓn���J�i
		If Address_F <> "" Then
			If ChkLen(Address_F, 100) = False Then
				Err = Err & "�Ԓn���J�i�̕��������������𒴂��Ă��܂��B�S�p�T�O�����ȓ��œ��͂��Ă��������B<br>"
				ErrAddress_F = sStyle
				flgReg = False
			End If
		End If

		'��TEL
		If HomeTelephoneNumber <> "" Then
			If IsTel(HomeTelephoneNumber, "1") = False Then
				Err = Err & "�o�b���[���A�h���X�Ɍ�肪����܂��B<br>"
				ErrHomeTelephoneNumber = sStyle
				flgReg = False
			End If
		End If

		'�o�b���[���A�h���X
		If MailAddress = "" Then
			Err = Err & "�o�b���[���A�h���X�͕K�{�ł��B<br>"
			ErrMailAddress = sStyle
			flgReg = False
		ElseIf IsMailAddress(MailAddress) = False Then
			Err = Err & "�o�b���[���A�h���X�Ɍ�肪����܂��B<br>"
			ErrMailAddress = sStyle
			flgReg = False
		End If
		'�o�b���[���A�h���X�m�F
		If MailAddress <> MailAddress2 Then
			Err = Err & "�o�b���[���A�h���X���m�F�̂��̂Ƃő��Ⴕ�Ă��܂��B<br>"
			ErrMailAddress2 = sStyle
			flgReg = False
		End If

		'�g�у��[���A�h���X
		If PortableMailAddress <> "" Then
			If IsMailAddress(PortableMailAddress) = False Then
				Err = Err & "�g�у��[���A�h���X�Ɍ�肪����܂��B<br>"
				ErrPortableMailAddress = sStyle
				flgReg = False
			End If
		End If
		'�g�у��[���A�h���X�m�F
		If PortableMailAddress <> PortableMailAddress2 Then
			Err = Err & "�g�у��[���A�h���X���m�F�̂��̂Ƃő��Ⴕ�Ă��܂��B<br>"
				ErrPortableMailAddress2 = sStyle
			flgReg = False
		End If

		'��]�A�����@�`�F�b�N
		If HomeContactFlag <> "" And HomeContactFlag <> "1" Then HomeContactFlag = ""
		If PortableContactFlag <> "" And PortableContactFlag <> "1" Then PortableContactFlag = ""
		If FaxContactFlag <> "" And FaxContactFlag <> "1" Then FaxContactFlag = ""
		If MailContactFlag <> "" And MailContactFlag <> "1" Then MailContactFlag = ""

		'�ً}�A����X�֔ԍ��`�F�b�N
		If UrgencyPost_U & UrgencyPost_L <> "" Then
			If IsRE(UrgencyPost_U & UrgencyPost_L, "^\d\d\d\d\d\d\d$", False) = True Then 
				sSQL = "/* ���E�Ҋ�{���ҏW���̗X�֔ԍ��`�F�b�N */ "
				sSQL = "EXEC up_DtlZip '" & UrgencyPost_U & "', '" & UrgencyPost_L & "'"

				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = False Then
					Err = Err & "�ً}�A����̗X�֔ԍ��͑��݂��܂���B<br>"
					ErrUrgencyPost_U = sStyle
					ErrUrgencyPost_L = sStyle
					flgReg = False
				End If
				Call RSClose(oRS)
			End If
		End If
		'�ً}�A����Z���`�F�b�N
		If UrgencyAddress <> "" Then
			If ChkLen(UrgencyAddress, 200) = False Then
				Err = Err & "�ً}�A����̏Z���͑S�p�Q�O�O�����ȓ��œ��͂��Ă��������B<br>"
				ErrUrgencyAddress = sStyle
				flgReg = False
			End If
		End If
		'�ً}�A����Z���J�i�`�F�b�N
		If UrgencyAddress_F <> "" Then
			If ChkLen(UrgencyAddress_F, 200) = False Then
				Err = Err & "�ً}�A����̏Z���J�i�͑S�p�Q�O�O�����ȓ��œ��͂��Ă��������B<br>"
				ErrUrgencyAddress_F = sStyle
				flgReg = False
			End If
		End If
		'�ً}�A����TEL�`�F�b�N
		If UrgencyTelephoneNumber <> "" Then
			If IsTel(UrgencyTelephoneNumber, "0") = False Then
				Err = Err & "�ً}�A����̓d�b�ԍ����s���ł��B�������d�b�ԍ�����͂��Ă��������B<br>"
				ErrUrgencyTelephoneNumber = sStyle
				flgReg = False
			End If
		End If

		'�z�[���y�[�W�`�F�b�N
		If URL <> "" Then
			If ChkLen(URL, 200) = False Then
				Err = Err & "�z�[���y�[�W�͔��p�łP�O�O�����ȓ��œ��͂��Ă��������B<br>"
				ErrURL = sStyle
				flgReg = False
			End If
		End If

		ChkData = flgReg
	End Function

	'******************************************************************************
	'�T�@�v�F�o�^�r�p�k�擾
	'���@���FvStaffCode	�F���E�҃R�[�h
	'���@�l�F
	'�X�@�V�F2008/04/22 LIS K.Kokubo
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim sSQL

		sSQL = "EXEC up_RegStaff_Edit1_1"
		sSQL = sSQL & " '" & vStaffCode & "'"
		sSQL = sSQL & ",'" & Name_1 & "�@" & Name_2 & "'"
		sSQL = sSQL & ",'" & Name_F_1 & "�@" & Name_F_2 & "'"
		sSQL = sSQL & ",'" & Name_1 & Name_2 & "'"
		sSQL = sSQL & ",'" & Name_F_1 & Name_F_2 & "'"
		sSQL = sSQL & ",'" & Post_U & "'"
		sSQL = sSQL & ",'" & Post_L & "'"
		sSQL = sSQL & ",'" & PrefectureCode & "'"
		sSQL = sSQL & ",'" & City & "'"
		sSQL = sSQL & ",'" & City_F & "'"
		sSQL = sSQL & ",'" & Town & "'"
		sSQL = sSQL & ",'" & Town_F & "'"
		sSQL = sSQL & ",'" & Address & "'"
		sSQL = sSQL & ",'" & Address_F & "'"
		sSQL = sSQL & ",'" & HomeTelephoneNumber & "'"
		sSQL = sSQL & ",'" & PortableTelephoneNumber & "'"
		sSQL = sSQL & ",'" & FaxNumber & "'"
		sSQL = sSQL & ",'" & MailAddress & "'"
		sSQL = sSQL & ",'" & PortableMailAddress & "'"
		sSQL = sSQL & ",'" & NoticeMailFlag & "'"
		sSQL = sSQL & ",'" & UrgencyPost_U & "'"
		sSQL = sSQL & ",'" & UrgencyPost_L & "'"
		sSQL = sSQL & ",'" & UrgencyAddress & "'"
		sSQL = sSQL & ",'" & UrgencyAddress_F & "'"
		sSQL = sSQL & ",'" & UrgencyTelephoneNumber & "'"
		sSQL = sSQL & ",'" & HomeContactFlag & "'"
		sSQL = sSQL & ",'" & PortableContactFlag & "'"
		sSQL = sSQL & ",'" & FaxContactFlag & "'"
		sSQL = sSQL & ",'" & MailContactFlag & "'"
		sSQL = sSQL & ",'" & Birthday & "'"
		sSQL = sSQL & ",'" & Sex & "'"
		sSQL = sSQL & ",'" & URL & "'"

		GetRegSQL = sSQL
	End Function

	'******************************************************************************
	'�T�@�v�F�X�V�O��̃f�[�^�̍��ق����[��(���X���܁F�o�^���X�V�ʒm�@�\)
	'���@���FvStaffCode	�F���E�҃R�[�h
	'���@�l�F��{���,���Z��,�A���悪�ς�����ꍇ�̂�
	'�X�@�V�F2008/04/22 LIS K.Kokubo
	'******************************************************************************
	Public Function DiffData(ByRef rSE)
		Dim sChg
		Dim sBef
		Dim sAft

		sChg = ""

		'��{���
		sBef = ""
		sAft = ""
		If Name_1 & Name_2 & Name_F_1 & Name_F_2 & Birthday & Sex _
		<> rSE.Name_1 & rSE.Name_2 & rSE.Name_F_1 & rSE.Name_F_2 & rSE.Birthday & rSE.Sex Then
			If Name_1 & Name_2 <> rSE.Name_1 & rSE.Name_2 Then
				sBef = sBef & "[���@�@�O]" & rSE.Name_1 & rSE.Name_2 & vbCrLf
				sAft = sAft & "[���@�@�O]" & Name_1 & Name_2 & vbCrLf
			End If

			If Name_F_1 & Name_F_2 <> rSE.Name_F_1 & rSE.Name_F_2 Then
				sBef = sBef & "[���O�J�i]" & rSE.Name_F_1 & rSE.Name_F_2 & vbCrLf
				sAft = sAft & "[���O�J�i]" & Name_F_1 & Name_F_2 & vbCrLf
			End If

			If Birthday <> rSE.Birthday Then
				sBef = sBef & "[���N����]" & rSE.Birthday & vbCrLf
				sAft = sAft & "[���N����]" & Birthday & vbCrLf
			End If

			If Sex <> rSE.Sex Then
				sBef = sBef & "[���@�@��]" & rSE.Sex & vbCrLf
				sAft = sAft & "[���@�@��]" & Sex & vbCrLf
			End If

			sChg = sChg & "---------- ��{��� ----------" & vbCrLf
			sChg = sChg & sBef
			sChg = sChg & "��" & vbCrLf
			sChg = sChg & sAft
		End If

		'���Z��
		sBef = ""
		sAft = ""
		If Post_U & Post_L & PrefectureCode & City & City_F & Town & Town_F & Address & Address_F _
		<> rSE.Post_U & rSE.Post_L & rSE.PrefectureCode & rSE.City & rSE.City_F & rSE.Town & rSE.Town_F & rSE.Address & rSE.Address_F Then
			If Post_U & Post_L <> rSE.Post_U & rSE.Post_L Then
				sBef = sBef & "[�X�֔ԍ�]" & rSE.Post_U & "-" & rSE.Post_L & vbCrLf
				sAft = sAft & "[�X�֔ԍ�]" & Post_U & "-" & Post_L & vbCrLf
			End If

			If PrefectureName & City & Town & Address <> rSE.PrefectureCode & rSE.City & rSE.Town & rSE.Address Then
				sBef = sBef & "[�Z�@�@��]" & rSE.PrefectureName & rSE.City & rSE.Town & rSE.Address & vbCrLf
				sAft = sAft & "[�Z�@�@��]" & PrefectureName & City & Town & Address & vbCrLf
			End If

			If City_F & Town_F & Address_F <> rSE.City_F & rSE.Town_F & rSE.Address_F Then
				sBef = sBef & "[�Z���J�i]" & rSE.City_F & rSE.Town_F & rSE.Address_F & vbCrLf
				sAft = sAft & "[�Z���J�i]" & City_F & Town_F & Address_F & vbCrLf
			End If

			sChg = sChg & "---------- ���Z�� ----------" & vbCrLf
			sChg = sChg & sBef
			sChg = sChg & "��" & vbCrLf
			sChg = sChg & sAft
		End If

		'�A����
		sBef = ""
		sAft = ""
		If HomeTelephoneNumber & PortableTelephoneNumber & FaxNumber & MailAddress & PortableMailAddress _
		<> rSE.HomeTelephoneNumber & rSE.PortableTelephoneNumber & rSE.FaxNumber & rSE.MailAddress & rSE.PortableMailAddress _
		Or HomeContactFlag <> rSE.HomeContactFlag Or PortableContactFlag <> rSE.PortableContactFlag Or FaxContactFlag <> rSE.FaxContactFlag Or MailContactFlag <> rSE.MailContactFlag Then
			If HomeTelephoneNumber <> rSE.HomeTelephoneNumber Then
				sBef = sBef & "[�s �d �k]" & rSE.HomeTelephoneNumber & vbCrLf
				sAft = sAft & "[�s �d �k]" & HomeTelephoneNumber & vbCrLf
			End If

			If PortableTelephoneNumber <> rSE.PortableTelephoneNumber Then
				sBef = sBef & "[�g�@�@��]" & rSE.PortableTelephoneNumber & vbCrLf
				sAft = sAft & "[�g�@�@��]" & PortableTelephoneNumber & vbCrLf
			End If

			If FaxNumber <> rSE.FaxNumber Then
				sBef = sBef & "[�e �` �w]" & rSE.FaxNumber & vbCrLf
				sAft = sAft & "[�e �` �w]" & FaxNumber & vbCrLf
			End If

			If MailAddress <> rSE.MailAddress Then
				sBef = sBef & "[�o�bMAIL]" & rSE.MailAddress & vbCrLf
				sAft = sAft & "[�o�bMAIL]" & MailAddress & vbCrLf
			End If

			If PortableMailAddress <> rSE.PortableMailAddress Then
				sBef = sBef & "[�g��MAIL]" & rSE.PortableMailAddress & vbCrLf
				sAft = sAft & "[�g��MAIL]" & PortableMailAddress & vbCrLf
			End If

			If HomeContactFlag <> rSE.HomeContactFlag Or PortableContactFlag <> rSE.PortableContactFlag Or FaxContactFlag <> rSE.FaxContactFlag Or MailContactFlag <> rSE.MailContactFlag Then
				sBef = sBef & "[�A�����@]"
				If rSE.HomeContactFlag = "1" Then sBef = sBef & "��,"
				If rSE.PortableContactFlag = "1" Then sBef = sBef & "�g��,"
				If rSE.FaxContactFlag = "1" Then sBef = sBef & "FAX,"
				If rSE.MailContactFlag = "1" Then sBef = sBef & "���[��,"
				sBef = sBef & vbCrLf

				sAft = sAft & "[�A�����@]"
				If HomeContactFlag = "1" Then sAft = sAft & "��,"
				If PortableContactFlag = "1" Then sAft = sAft & "�g��,"
				If FaxContactFlag = "1" Then sAft = sAft & "FAX,"
				If MailContactFlag = "1" Then sAft = sAft & "���[��,"
				sAft = sAft & vbCrLf
			End If

			sChg = sChg & "---------- �A���� ----------" & vbCrLf
			sChg = sChg & sBef
			sChg = sChg & "��" & vbCrLf
			sChg = sChg & sAft
			sChg = sChg & vbCrLf
		End If

		DiffData = sChg
	End Function
End Class
%>
