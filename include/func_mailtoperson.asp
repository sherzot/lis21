<%
'**********************************************************************************************************************
'概　要：企業側メール作成画面 しごとナビ/company/mailtoperson.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/config/constant.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　メール　■■■
'　　　：GetMailSubject					：メールの件名を生成して取得
'　　　：GetMailBodyCompany				：メールの内容を生成して取得
'　　　：GetMailSignature				：メールの署名を取得
'　　　：GetNaviMailTemplateOptionHtml	：ナビメールテンプレート取得
'　　　：GetMailTemplateOptionHtml		：求人票メールテンプレート取得
'　　　：RegMail						：メールをＤＢに登録する
'　　　：MailToPerson					：メールを送信する
'**********************************************************************************************************************
%>
<!-- #INCLUDE VIRTUAL="/include/func/setMail_MailToPerson.asp" -->
<%
'******************************************************************************
'概　要：メールの件名を生成して取得
'作　者：2007/06/22 Lis K.Kokubo
'引　数：vJobTypeDetail	：具体的職種名
'戻り値：
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'更　新：
'******************************************************************************
Function GetMailSubject(ByVal vJobTypeDetail)
	If Len(vJobTypeDetail) <= 15 Then
		GetMailSubject = "■しごとナビ■スカウト・連絡メール着信のお知らせ！／「" & Left(vJobTypeDetail,15) & "」のお仕事"
	Else
		GetMailSubject = "■しごとナビ■スカウト・連絡メール着信のお知らせ！／「" & Left(vJobTypeDetail,15) & "...」のお仕事"
	End If
End Function

'******************************************************************************
'概　要：メールの内容を生成して取得
'引　数：vOrderCode		：メールに付随する求人票の情報コード
'　　　：vSubject		：メールの件名
'　　　：vBody			：メールの内容
'　　　：vStaffCode		：メール受信側求職者コード
'　　　：vStaffName		：メール受信側求職者名
'　　　：vCompanyName	：メール送信側企業名
'　　　：vJobTypeDetail	：具体的職種名
'戻り値：
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'履　歴：2007/06/22 LIS K.Kokubo 作成
'　　　：2009/07/02 LIS K.Kokubo スカウトメール未読数表示を削除。未読通知メールを出しているために不要。
'******************************************************************************
Function GetMailBodyCompany(ByVal vOrderCode, ByVal vSubject, ByVal vBody, ByVal vStaffCode, ByVal vStaffName, ByVal vCompanyName, ByVal vJobTypeDetail, ByVal vType)
	Dim sBody
	Dim iLen
	Dim idx

	GetMailBodyCompany = ""

	'本文
	sBody = ""
	If vStaffName <> "" Then sBody = vStaffName & "　様"  & vbCrLf & vbCrLf
	sBody = sBody & MAIL_FROM_COMPANY_BODY & vbCrLf
	sBody = sBody & MAIL_URL_STAFF & "?si=" & vStaffCode & vbCrLf

	'メール内容表示処理
	iLen = Len(vBody) * 0.3
	sBody = sBody & vbCrLf & _
		"----------------------■　最新配信情報　■-----------------------" & vbCrLf & _
		"【会社名】　"

	If vType = "2" Then
			sBody = sBody & "リス株式会社"
	ElseIf vType = "1" Then
			sBody = sBody & vCompanyName	'CC_CompanyName_K
	Else
			sBody = sBody & vCompanyName
	End If

	sBody = sBody & vbCrLf & _
		"【仕事内容】　" & vJobTypeDetail & "(" & vOrderCode & ")" & vbCrLf & _
		"-----------------------■　メール情報　■-------------------------" & vbCrLf & _
		"【メールタイトル】　" & vSubject & vbCrLf & _
		"【メール内容】" & vbCrLf & Left(vBody, iLen) & "..." & vbCrLf & _
		"----------------------------------------------------------------" & vbCrLf

	sBody = sBody & MAIL_FROM_COMPANY_FOOTER

	GetMailBodyCompany = sBody
End Function

'******************************************************************************
'概　要：メールの署名を取得
'作　者：2007/06/22 Lis K.Kokubo
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserType	：ログイン中のユーザ種類
'　　　：vOrderCode	：メールに付随する求人票の情報コード
'戻り値：
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'更　新：
'******************************************************************************
Function GetMailSignature(ByRef rDB, ByVal vUserType, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	GetMailSignature = ""
	If G_USERTYPE <> "staff" Then
		sSQL = "sp_GetDataMailSignatureCompany '" & vOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

		If GetRSState(oRS) = True Then
			GetMailSignature = GetMailSignature & "------------------------------" & vbCrLf
			GetMailSignature = GetMailSignature & "会社名　　：" & oRS.Collect("CompanyName") & vbCrLf
			GetMailSignature = GetMailSignature & "担当者部署：" & oRS.Collect("SectionName") & vbCrLf
			GetMailSignature = GetMailSignature & "電話番号　：" & oRS.Collect("TelephoneNumber") & vbCrLf
			GetMailSignature = GetMailSignature & "担当者氏名：" & oRS.Collect("PersonName") & vbCrLf
			GetMailSignature = GetMailSignature & "担当者Mail：" & oRS.Collect("MailAddress") & vbCrLf
		End If
		Call RSClose(oRS)
	Else
		'署名用情報の取得
		sSQL = "sp_GetDataMailSignatureStaff '" & G_USERID & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

		If GetRSState(oRS) = True Then
			GetMailSignature = GetMailSignature & "------------------------------" & vbCrLf & _
				"住所：" & oRS.Collect("Prefecture") & oRS.Collect("City") & oRS.Collect("Town") & oRS.Collect("Address") & _
				"氏名：" & oRS.Collect("Name")

			If oRS.Collect("HomeContactFlag") = "1" Then
				GetMailSignature = GetMailSignature & "自宅：" & oRS.Collect("HomeTelephoneNumber")
			End If
			If oRS.Collect("PortableContactFlag") = "1" Then
				GetMailSignature = GetMailSignature & "携帯：" & oRS.Collect("PortableTelephoneNumber")
			End If
			If oRS.Collect("FaxContactFlag") = "1" Then
				GetMailSignature = GetMailSignature & "FAX ：" & oRS.Collect("FaxNumber")
			End If
			If oRS.Collect("MailContactFlag") = "1" Then
				GetMailSignature = GetMailSignature & "Mail：" & oRS.Collect("MailAddress")
			End If
		End If
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'概　要：ナビメールテンプレート取得
'作　者：2007/06/22 Lis K.Kokubo
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserType	：ログイン中のユーザ種類
'　　　：vNaviSEQ	：番号
'戻り値：
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'更　新：
'******************************************************************************
Function GetNaviMailTemplateOptionHtml(ByRef rDB, ByVal vUserType, ByVal vNaviSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sSelected

	GetNaviMailTemplateOptionHtml = ""

	sSQL = "sp_GetDataMailTemplate '2'"

	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sSelected = ""
		If CStr(oRS.Collect("Cd")) = CStr(vNaviSEQ) Then sSelected = "selected=""true"""
		GetNaviMailTemplateOptionHtml = GetNaviMailTemplateOptionHtml & _
			"<option value=""" & oRS.Collect("Cd") & """ " & sSelected & ">" & oRS.Collect("Title") & "</option>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：求人票メールテンプレート取得
'作　者：2007/06/22 Lis K.Kokubo
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserCode	：ログイン中のユーザＩＤ
'　　　：vOrderCode	：メールに付随する求人票の情報コード
'　　　：vSEQ		：番号
'戻り値：
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'更　新：
'******************************************************************************
Function GetMailTemplateOptionHtml(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sSelected

	GetMailTemplateOptionHtml = ""

	sSQL = "up_GetListMailTemplate '" & vUserCode & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	Do While GetRSState(oRS) = True
		sSelected = ""
		If CStr(oRS.Collect("SEQ")) = CStr(vSEQ) Then sSelected = "selected=""true"""
		GetMailTemplateOptionHtml = GetMailTemplateOptionHtml & _
			"<option value=""" & oRS.Collect("SEQ") & """ " & sSelected & ">" & oRS.Collect("Subject") & "</option>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：メールをＤＢに登録する
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserCode	：ログイン中のユーザＩＤ
'　　　：vOrderCode	：メールに付随する求人票の情報コード
'　　　：vSEQ		：番号
'戻り値：Boolean	：[True]メールが登録できた [False]メールの登録でエラーが発生した
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'履　歴：2007/06/22 LIS K.Kokubo 作成
'　　　：2011/01/05 LIS K.Kokubo Basp.SendMail → SndMail
'******************************************************************************
Function RegMail(ByRef rDB, ByVal vID, ByVal vUserID, ByVal vReceiverCode, ByVal vSubject, ByVal vBody, ByVal vOrderCode, ByVal vSenderEvaluation, ByVal vSenderRemark, ByVal vReceiverEvaluation, ByVal vReceiverRemark, ByVal vSenderDelFlag, ByVal vReceiverDelFlag, ByVal vAnswerFlag, ByVal vPayFlag)
	On Error Resume Next
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sSessionValue:	sSessionValue = ""
	Dim sFormValue:	sFormValue = ""
	Dim idx
	Dim sMsg

	sSQL = ""
	sSQL = sSQL & "/* しごとナビ メール登録 */" & vbCrLf
	sSQL = sSQL & "up_RegMailHistory"
	sSQL = sSQL & " '" & vID & "'"
	sSQL = sSQL & ",'" & vUserID & "'"
	sSQL = sSQL & ",'" & vReceiverCode & "'"
	sSQL = sSQL & ",'" & vSubject & "'"
	sSQL = sSQL & ",'" & vBody & "'"
	sSQL = sSQL & ",'" & vOrderCode & "'"
	sSQL = sSQL & ",'" & vSenderEvaluation & "'"
	sSQL = sSQL & ",'" & vSenderRemark & "'"
	sSQL = sSQL & ",'" & vReceiverEvaluation & "'"
	sSQL = sSQL & ",'" & vReceiverRemark & "'"
	sSQL = sSQL & ",'" & vSenderDelFlag & "'"
	sSQL = sSQL & ",'" & vReceiverDelFlag & "'"
	sSQL = sSQL & ",'" & vAnswerFlag & "'"
	sSQL = sSQL & ",'" & vPayFlag & "';"

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If flgQE = True Then
		RegMail = True

		If G_PLANTYPE = "mail" And vPayFlag = "1" Then
			sSQL = ""
			sSQL = sSQL & "/* 課金メール送信ポイント */" & vbCrLf
			sSQL = sSQL & "EXEC up_RegCMPNaviPoint '" & vUserID & "','','003','" & GetDateStr(Date,"") & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		End If
	Else
		Session("err") = "スカウトメールの送信に失敗しました。<br>お手数ですが本文を確認の上、再度ご送信下さい。"

		sSQL = "EXEC up_Reg_LOG_Error '" & G_USERID & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'" & _
			",'" & ChkSQLStr(Session.SessionID) & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("URL")) & "?" & ChkSQLStr(Request.ServerVariables("QUERY_STRING")) & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("HTTP_REFERER")) & "'" & _
			",'" & ChkSQLStr(sSQL) & "'" & _
			",'" & ChkSQLStr(Err.Source & vbCrLf & Err.Description) & "'"

		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

		For idx = 1 To Session.Contents.Count
			sSessionValue = sSessionValue & "【" & Session.Contents.Key(idx) & "】"
			sSessionValue = sSessionValue & Session.Contents(idx) & vbCrLf
		Next

		For idx = 1 To Request.Form.Count
			sFormValue = sFormValue & "【" & Request.Form.Key(idx) & "】"
			sFormValue = sFormValue & Request.Form(idx) & vbCrLf
		Next

		sMsg = "UserID     ：" & G_USERID & vbCrLf & _
			"IPAddress  ：" & Request.ServerVariables("REMOTE_ADDR") & vbCrLf & _
			"UserAgent  ：" & Request.ServerVariables("HTTP_USER_AGENT") & vbCrLf & _
			"Referer    ：" & Request.ServerVariables("HTTP_REFERER") & vbCrLf & _
			"Page       ：" & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING") & vbCrLf & _
			"──────────────────────────────" & vbCrLf & _
			"Error Page ：" & vbCrLf & _
			Session("errorpagereferer") & vbCrLf & "↓" & vbCrLf & _
			Session("errorpage") & vbCrLf & _
			"──────────────────────────────" & vbCrLf & _
			"Description：" & vbCrLf & Err.Description & vbCrLf & _
			"──────────────────────────────" & vbCrLf & _
			"Session    ：" & vbCrLf & sSessionValue & vbCrLf & _
			"──────────────────────────────" & vbCrLf & _
			"Post       ：" & vbCrLf & sFormValue & vbCrLf

		Call SndMail(Cnt_MailServer, "kisui@lis21.co.jp", "info@shigotonavi.jp", "【しごとナビ エラー】", sMsg, "")

		RegMail = False
	End If

	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：メールを送信する
'引　数：vMailServer
'　　　：vFrom
'　　　：vUserCode
'　　　：vID
'　　　：vStaffCode
'　　　：vOrderCode
'　　　：vSubject
'　　　：vBody
'　　　：vPayFlag
'戻り値：Boolean	：[True]メールが登録できた [False]メールの登録でエラーが発生した
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'履　歴：2007/06/22 LIS K.Kokubo
'　　　：2011/01/05 LIS K.Kokubo Basp.SendMail → SndMail
'******************************************************************************
Function MailToPerson(ByVal vMailServer, ByVal vFrom, ByVal vUserCode, ByVal vID, ByVal vStaffCode, ByVal vOrderCode, ByVal vSubject, ByVal vBody, ByVal vPayFlag)
	'求職者に向けてメールを送信する場合に使用するＰＧです。
	Dim MOBILE_URL:			MOBILE_URL = "http://m.shigotonavi.jp/"
	Dim MOBILE_URL_SSL:		MOBILE_URL_SSL = "https://m.shigotonavi.jp/"
	Dim PC_URL:				PC_URL = HTTP_CURRENTURL
	Dim LIS_MAILADDRESS:	LIS_MAILADDRESS = "lis@lis21.co.jp"

	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim sRes
	Dim sTo
	Dim sSubject
	Dim sBody
	Dim sPortableSubject
	Dim sPortableBody

	Dim flgRegMail					'メールDB登録処理完了フラグ：[True]登録完了 [False]エラー
	Dim sType						'メール送信側企業の種類 ["1"]リス以外の企業 ["2"]リス
	Dim sCompanyName				'メール送信側の企業名
	Dim sStaffName					'メール受信側の求職者名
	Dim sStaffMailAddress			'メール受信側求職者のＰＣメール
	Dim sStaffPortableMailAddress	'メール受信側求職者のケータイメール
	Dim sNoticeMailFlag				'メール送信先フラグ
	Dim sJobTypeDetail				'メールに付随する求人票の具体的職種名

	MailToPerson = False

	'メール内容で使用
	sSQL = "sp_GetDataMailToStaff '" & G_USERID & "', '" & vStaffCode & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
		sCompanyName = ChkStr(oRS.Collect("CompanyName"))
		sStaffName = ChkStr(oRS.Collect("ReceiverName"))
		sStaffMailAddress = ChkStr(oRS.Collect("ReceiverMailAddress"))
		sStaffPortableMailAddress = ChkStr(oRS.Collect("ReceiverPortableMailAddress"))
		sNoticeMailFlag = ChkStr(oRS.Collect("NoticeMailFlag"))
		sType = ChkStr(oRS.Collect("Type"))
		sJobTypeDetail = ChkStr(oRS.Collect("JobTypeDetail"))
	End If
	Call RSClose(oRS)

	If sStaffMailAddress & sStaffPortableMailAddress <> "" Then
		'メール送信情報をテーブルに格納
		flgRegMail = RegMail(dbconn, vID, G_USERID, vStaffCode, vSubject, vBody, vOrderCode, "", "", "", "", "", "", "", vPayFlag)
		MailToPerson = flgRegMail

		'メールが正常登録の場合のみメールの送信処理をする。
		If flgRegMail = True Then

			'***************************************************************************
			'メール送信 start
			'---------------------------------------------------------------------------
			'ＰＣメール送信
			If sNoticeMailFlag = "0" Or sNoticeMailFlag = "1" Then
				sSubject = GetMailSubject(sJobTypeDetail)
				sBody = GetMailBodyCompany(vOrderCode, vSubject, vBody, vStaffCode, sStaffName, sCompanyName, sJobTypeDetail, sType)

				sTo = sStaffMailAddress		'送信先メールアドレス
				If Len(sTo) > 0 Then sRes = SndMail(vMailServer, sTo, vFrom, sSubject, sBody, "")
			End If

			'ケータイメール送信
			If sNoticeMailFlag = "0" Or sNoticeMailFlag = "2" Then
				sPortableSubject = "[しごとナビ]企業からﾒｰﾙ着信"
				sPortableBody = "いつもご利用ありがとうございます。" & vbCrLf & _
					"総合求人求職ｻｲﾄ｢しごとナビ｣(リス株式会社)です。" & vbCrLf & _
					"｢しごとナビモバイル｣を通じて、求人企業から貴方へﾒｰﾙが届きました。" & vbCrLf & _
					"｢しごとナビ｣にﾛｸﾞｲﾝして、ﾒｰﾙ内容をご確認下さい。" & vbCrLf & vbCrLf & _
					"※携帯版(ﾓﾊﾞｲﾙ)の場合:｢しごとナビモバイル｣TOPからﾛｸﾞｲﾝ⇒Myﾍﾟｰｼﾞの｢ﾒｰﾙ履歴｣ﾘﾝｸをｸﾘｯｸ" & vbCrLf & _
					MOBILE_URL & vbCrLf & _
					"※PC版の場合:｢しごとナビ｣TOPからﾛｸﾞｲﾝ⇒ﾛｸﾞｲﾝﾒﾆｭｰの｢ﾒｰﾙ履歴｣ﾘﾝｸをｸﾘｯｸ" & vbCrLf & _
					"-----------------"  & vbCrLf & _
					"※このﾒｰﾙは自動送信ﾒｰﾙのため、返信できません。ご注意下さい。" & vbCrLf & _
					"-----------------"  & vbCrLf & _
					"リス株式会社" & LIS_MAILADDRESS & vbCrLf & _
					"しごとナビモバイル：" & MOBILE_URL & vbCrLf & _
					"しごとナビ：" & PC_URL

				sTo = sStaffPortableMailAddress
				If Len(sTo) > 0 Then sRes = SndMail(vMailServer, sTo, vFrom, sPortableSubject, sPortableBody, "")
			End If
			'---------------------------------------------------------------------------
			'メール送信 end
			'***************************************************************************
		End If
	End If
End Function
%>
