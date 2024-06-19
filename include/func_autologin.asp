<%
'**********************************************************************************************************************
'概　要：自動ログイン
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：AutoLogin	：自動ログイン
'**********************************************************************************************************************

'このモジュールをINCLUDEと同時に自動ログイン実行
Call AutoLogin()

'******************************************************************************
'概　要：自動ログイン
'引　数：
'備　考：
'使用元：ナビ/
'更　新：2008/05/23 LIS K.Kokubo
'******************************************************************************
Function AutoLogin()
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sCertify
	Dim sRedirectURL

	If Session("userid") = "" Then	'G_USERIDはまだ値が設定されていない可能性があるので使用不可
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
					Session("plantype") = oRS.Collect("PlanTypeName") '2008/01/24 LIS K.Kokubo 追加
					Session("applicationcode") = ChkStr(oRS.Collect("ApplicationCode")) '2008/06/04 LIS K.Kokubo 追加
					Session("useflag") = oRS.Collect("UseFlag") '2008/06/04 LIS K.Kokubo 追加
					Session("publicflag") = oRS.Collect("PublicFlag") '2008/06/06 LIS K.Kokubo 追加
					Session("mailreadflag") = oRS.Collect("MailReadFlag") '2008/06/06 LIS K.Kokubo 追加
					Session("imagelimit") = oRS.Collect("ImageLimit") '2009/03/10 LIS K.Kokubo 追加
					Session("interviewflag") = oRS.Collect("InterviewFlag") '2009/03/11 LIS K.Kokubo 追加
					Session("temppermitflag") = oRS.Collect("TempPermitFlag") '2009/03/17 LIS K.Kokubo 追加
					Session("intropermitflag") = oRS.Collect("IntroPermitFlag") '2009/03/17 LIS K.Kokubo 追加
					If oRS.Collect("UseFlag") = "0" Then
						Session("oldapplicationcode") = oRS.Collect("OldApplicationCode")
						Session("oldplantype") = oRS.Collect("OldPlanTypeName")
						Session("oldimagelimit") = oRS.Collect("OldImageLimit") '2009/03/10 LIS K.Kokubo 追加
						Session("oldinterviewflag") = oRS.Collect("OldInterviewFlag") '2009/03/11 LIS K.Kokubo 追加
					End If
				End If

				'グローバル変数に代入
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
				'手で書き換えたクッキーの可能性があるので、このセッションはログインできるまでは自動ログインしないようにする
				'Response.Cookies("certify") = ""
				'Session("autologinflag") = "0"
			End If
			Call RSClose(oRS)

			If G_USERID <> "" Then
				'Cookiesの更新
				Call AutoLogin_WriteCookies(G_USERID, G_WEBSERVERNAME)

				'最終ログイン日更新
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
'概　要：自動ログインの認証コードをクッキーに書き込み
'引　数：vUserID
'　　　：vDomainName
'備　考：
'使用元：ナビ/
'更　新：2008/05/27 LIS K.Kokubo
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

		'クッキーの利用可能サイト
		Response.Cookies("certify").Domain = vDomainName
		'有効期限は90日間(3ヶ月)
		Response.Cookies("certify").Expires = Date + 89
		'認証文字列
		Response.Cookies("certify") = sCertify
		Session.Contents.Remove("autologinflag")
	End If
End Function
%>
