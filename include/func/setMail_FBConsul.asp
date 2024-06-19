<%
'*******************************************************************************
'概　要：キャリアカウンセラーへ相談の内容確認メールの文言を生成
'引　数：vMailType		：送信先メール種類 [1]PC[2]MOBILE
'　　　：vName			：求職者名
'　　　：vTitle			：相談タイトル
'　　　：vBody			：相談本文
'　　　：rSubject		：[OUTPUT]メール件名
'　　　：rBody			：[OUTPUT]メール本文
'戻り値：Boolean
'備　考：
'履　歴：2011/11/29 LIS K.Kokubo 作成
'*******************************************************************************
Function setMail_FBConsul(ByVal vMailType,ByVal vName,ByVal vTitle,ByVal vBody,ByRef rSubject,ByRef rBody)
	Dim sSubject,sBody

	setMail_FBConsul = False

	If CStr(vMailType) = "1" Then
		sSubject = "■しごとナビ■キャリアカウンセラーへの相談を受け付けました"
		sBody = ""
		sBody = sBody & vName & "　様" & vbCrLf & vbCrLf

		sBody = sBody & "ご利用ありがとうございます。" & vbCrLf
		sBody = sBody & "しごとナビ運営事務局です。" & vbCrLf & vbCrLf

		sBody = sBody & "キャリアカウンセラーへの相談を受け付けましたのでお知らせ致します。" & vbCrLf
		sBody = sBody & "ご相談を頂いた内容は下記の通りです。" & vbCrLf & vbCrLf

		sBody = sBody & "----------------------------------------------------------------------" & vbCrLf
		sBody = sBody & "【相談タイトル】" & vbCrLf
		sBody = sBody & vTitle & vbCrLf & vbCrLf
		sBody = sBody & "【本文】" & vbCrLf
		sBody = sBody & vBody & vbCrLf
		sBody = sBody & "----------------------------------------------------------------------" & vbCrLf & vbCrLf

		sBody = sBody & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
		sBody = sBody & "■すべてがつながる「しごとナビ」（リス株式会社）" & vbCrLf
		sBody = sBody & HTTP_CURRENTURL & vbCrLf
		sBody = sBody & "■しごとナビFacebookページ" & vbCrLf
		sBody = sBody & HTTP_FB & vbCrLf
		sBody = sBody & "お問い合わせ：lis@lis21.co.jp" & vbCrLf

		setMail_FBConsul = True
	ElseIf CStr(vMailType) = "2" Then
		sSubject = "■しごとﾅﾋﾞ■ｷｬﾘｱｶｳﾝｾﾗｰへの相談を受け付けました"
		sBody = ""
		sBody = sBody & vName & "　様" & vbCrLf & vbCrLf

		sBody = sBody & "ご利用ありがとうございます。" & vbCrLf
		sBody = sBody & "しごとﾅﾋﾞ運営事務局です。" & vbCrLf & vbCrLf

		sBody = sBody & "ｷｬﾘｱｶｳﾝｾﾗｰへの相談を受け付けましたのでお知らせ致します。" & vbCrLf
		sBody = sBody & "ご相談を頂いた内容は下記の通りです。" & vbCrLf & vbCrLf

		sBody = sBody & "------------------------------" & vbCrLf
		sBody = sBody & "【相談ﾀｲﾄﾙ】" & vbCrLf
		sBody = sBody & vTitle & vbCrLf & vbCrLf
		sBody = sBody & "【本文】" & vbCrLf
		sBody = sBody & vBody & vbCrLf
		sBody = sBody & "------------------------------" & vbCrLf & vbCrLf

		sBody = sBody & "━━━━━━━━━━━━━━━" & vbCrLf
		sBody = sBody & "■しごとﾅﾋﾞ(ﾘｽ株式会社)" & vbCrLf
		sBody = sBody & HTTP_CURRENTURL & vbCrLf
		sBody = sBody & "■しごとﾅﾋﾞFacebookﾍﾟｰｼﾞ" & vbCrLf
		sBody = sBody & HTTP_FB & vbCrLf
		sBody = sBody & "お問い合わせ：lis@lis21.co.jp" & vbCrLf

		setMail_FBConsul = True
	End If

	rSubject = sSubject
	rBody = sBody
End Function
%>
