<%
'*******************************************************************************
'概　要：企業からのメール着信通知メールの文言を生成
'引　数：vMailType		：送信先メール種類 [pc][mobile]
'　　　：vSubject		：メール履歴に登録した件名
'　　　：vBody			：メール履歴に登録した本文
'　　　：vStaffCode		：送信先求職者コード
'　　　：vStaffName		：送信先求職者名
'　　　：vCompanyName	：送信元企業名
'　　　：vOrderCode		：メールに紐づいた情報コード
'　　　：vJobTypeDetail	：メールに紐づいた求人の具体的職種名
'　　　：rSubject		：[OUTPUT]件名
'　　　：rBody			：[OUTPUT]本文
'戻り値：Boolean
'備　考：
'更　新：2009/07/02 LIS K.Kokubo 作成
'*******************************************************************************
Function setMail_MailToPerson(ByVal vMailType, ByVal vSubject, ByVal vBody, ByVal vStaffCode, ByVal vStaffName, ByVal vCompanyName, ByVal vOrderCode, ByVal vJobTypeDetail, ByRef rSubject, ByRef rBody)
	setMail_MailToPerson = False

	If vMailType = "pc" Then
		rSubject = GetMailSubject(vJobTypeDetail)
		rBody = GetMailBodyCompany(vOrderCode, vSubject, vBody, vStaffCode, vStaffName, vCompanyName, vJobTypeDetail, "1")

		setMail_MailToPerson = True
	ElseIf vMailType = "mobile" Then
		rSubject = "[しごとナビ]企業からﾒｰﾙ着信"

		rBody = ""
		rBody = rBody & "いつもご利用ありがとうございます。" & vbCrLf
		rBody = rBody & "総合求人求職ｻｲﾄ｢しごとナビ｣(リス株式会社)です。" & vbCrLf
		rBody = rBody & "｢しごとナビモバイル｣を通じて、求人企業から貴方へﾒｰﾙが届きました。" & vbCrLf
		rBody = rBody & "｢しごとナビ｣にﾛｸﾞｲﾝして、ﾒｰﾙ内容をご確認下さい。" & vbCrLf & vbCrLf
		rBody = rBody & "※携帯版(ﾓﾊﾞｲﾙ)の場合:｢しごとナビモバイル｣TOPからﾛｸﾞｲﾝ⇒Myﾍﾟｰｼﾞの｢ﾒｰﾙ履歴｣ﾘﾝｸをｸﾘｯｸ" & vbCrLf
		rBody = rBody & HTTP_NAVI_MOBILE & vbCrLf
		rBody = rBody & "※PC版の場合:｢しごとナビ｣TOPからﾛｸﾞｲﾝ⇒ﾛｸﾞｲﾝﾒﾆｭｰの｢ﾒｰﾙ履歴｣ﾘﾝｸをｸﾘｯｸ" & vbCrLf
		rBody = rBody & "-----------------"  & vbCrLf
		rBody = rBody & "※このﾒｰﾙは自動送信ﾒｰﾙのため、返信できません。ご注意下さい。" & vbCrLf
		rBody = rBody & "-----------------"  & vbCrLf
		rBody = rBody & "リス株式会社" & MAIL_LIS & vbCrLf
		rBody = rBody & "しごとナビモバイル：" & HTTP_NAVI_MOBILE & vbCrLf
		rBody = rBody & "しごとナビ：" & HTTP_CURRENTURL

		setMail_MailToPerson = True
	End If
End Function
%>
