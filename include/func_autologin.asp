<%
'**********************************************************************************************************************
'�T�@�v�F�������O�C��
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���FAutoLogin	�F�������O�C��
'**********************************************************************************************************************

'���̃��W���[����INCLUDE�Ɠ����Ɏ������O�C�����s
Call AutoLogin()

'******************************************************************************
'�T�@�v�F�������O�C��
'���@���F
'���@�l�F
'�g�p���F�i�r/
'�X�@�V�F2008/05/23 LIS K.Kokubo
'******************************************************************************
Function AutoLogin()
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sCertify
	Dim sRedirectURL

	If Session("userid") = "" Then	'G_USERID�͂܂��l���ݒ肳��Ă��Ȃ��\��������̂Ŏg�p�s��
		If Session("autologinflag") <> "0" Then
			sCertify = Request.Cookies("certify")
			If sCertify <> "" Then
				If Request.ServerVariables("HTTPS") <> "on" Then
					sRedirectURL = "https://" & G_WEBSERVERNAME & Request.ServerVariables("URL")
					If G_QUERYSTRING <> "" Then sRedirectURL = sRedirectURL &  "?" & G_QUERYSTRING
					Response.Redirect sRedirectURL
				End If
			End If
		End If

		If sCertify <> "" Then
			sSQL = "EXEC up_ChkNaviLogin_Auto '" & sCertify & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				Session("usertype") = ChkStr(oRS.Collect("UserType"))
				Session("userid") = ChkStr(oRS.Collect("LoginID"))
				Session("password") = oRS.Collect("Password")
				If Session("usertype") = "company" Then
					Session("companykbn") = oRS.Collect("CompanyKbn")
					Session("plantype") = oRS.Collect("PlanTypeName") '2008/01/24 LIS K.Kokubo �ǉ�
					Session("applicationcode") = ChkStr(oRS.Collect("ApplicationCode")) '2008/06/04 LIS K.Kokubo �ǉ�
					Session("useflag") = oRS.Collect("UseFlag") '2008/06/04 LIS K.Kokubo �ǉ�
					Session("publicflag") = oRS.Collect("PublicFlag") '2008/06/06 LIS K.Kokubo �ǉ�
					Session("mailreadflag") = oRS.Collect("MailReadFlag") '2008/06/06 LIS K.Kokubo �ǉ�
					Session("imagelimit") = oRS.Collect("ImageLimit") '2009/03/10 LIS K.Kokubo �ǉ�
					Session("interviewflag") = oRS.Collect("InterviewFlag") '2009/03/11 LIS K.Kokubo �ǉ�
					Session("temppermitflag") = oRS.Collect("TempPermitFlag") '2009/03/17 LIS K.Kokubo �ǉ�
					Session("intropermitflag") = oRS.Collect("IntroPermitFlag") '2009/03/17 LIS K.Kokubo �ǉ�
					If oRS.Collect("UseFlag") = "0" Then
						Session("oldapplicationcode") = oRS.Collect("OldApplicationCode")
						Session("oldplantype") = oRS.Collect("OldPlanTypeName")
						Session("oldimagelimit") = oRS.Collect("OldImageLimit") '2009/03/10 LIS K.Kokubo �ǉ�
						Session("oldinterviewflag") = oRS.Collect("OldInterviewFlag") '2009/03/11 LIS K.Kokubo �ǉ�
					End If
				End If

				'�O���[�o���ϐ��ɑ��
				G_USERID = Session("userid")
				G_USERTYPE = Session("usertype")
				G_COMPANYKBN = Session("companykbn")
				G_PLANTYPE = Session("plantype")
				G_APPLICATIONCODE = Session("applicationcode")
				G_OLDAPPLICATIONCODE = Session("oldapplicationcode")
				G_OLDPLANTYPE = Session("oldplantype")
				G_USEFLAG = Session("useflag")
				G_PUBLICFLAG = Session("publicflag")
				G_MAILREADFLAG = Session("mailreadflag")
			Else
				'��ŏ����������N�b�L�[�̉\��������̂ŁA���̃Z�b�V�����̓��O�C���ł���܂ł͎������O�C�����Ȃ��悤�ɂ���
				'Response.Cookies("certify") = ""
				'Session("autologinflag") = "0"
			End If
			Call RSClose(oRS)

			If G_USERID <> "" Then
				'Cookies�̍X�V
				Call AutoLogin_WriteCookies(G_USERID, G_WEBSERVERNAME)

				'�ŏI���O�C�����X�V
				sSQL = "sp_Reg_LastAccessDay '" & G_USERID & "','1'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				Application("navistatus_cnt_login") = Application("navistatus_cnt_login") + 1
			End If
		Else
			Session("autologinflag") = "0"
		End If
	End If
End Function

'******************************************************************************
'�T�@�v�F�������O�C���̔F�؃R�[�h���N�b�L�[�ɏ�������
'���@���FvUserID
'�@�@�@�FvDomainName
'���@�l�F
'�g�p���F�i�r/
'�X�@�V�F2008/05/27 LIS K.Kokubo
'******************************************************************************
Function AutoLogin_WriteCookies(ByVal vUserID, ByVal vDomainName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sCertify

	sSQL = "up_RegAutoLogin '" & vUserID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sCertify = oRS.Collect("Certify")

		'�N�b�L�[�̗��p�\�T�C�g
		Response.Cookies("certify").Domain = vDomainName
		'�L��������90����(3����)
		Response.Cookies("certify").Expires = Date + 89
		'�F�ؕ�����
		Response.Cookies("certify") = sCertify
		Session.Contents.Remove("autologinflag")
	End If
End Function
%>
