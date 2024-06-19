<%
'*******************************************************************************
'概　要：APIキー申し込み完了時のメール内容
'引　数：vAPIKey：APIキー
'　　　：rSbj	：[OUTPUT]メール件名
'　　　：rBdy	：[OUTPUT]メール本文
'戻り値：Boolean
'備　考：
'履　歴：2012/02/13 LIS K.Kokubo 作成
'*******************************************************************************
Function setMail_APIUser(ByVal vAPIKey,ByRef rSbj,ByRef rBdy)
	Dim sSbj,sBdy

	setMail_APIUser = False

	sSbj = "■しごとナビ■しごとナビ求人検索APIのAPIキー申し込み内容確認"

	sBdy = ""
	sBdy = sBdy & "しごとナビ求人検索APIのご利用ありがとうございます。" & vbCrLf
	sBdy = sBdy & "しごとナビサポートです。" & vbCrLf & vbCrLf

	sBdy = sBdy & "APIキーを発行しましたので、下記URLをクリックして確定してください。" & vbCrLf & vbCrLf

	sBdy = sBdy & HTTPS_CURRENTURL & "api/approval/?key=" & vAPIKey & vbCrLf & vbCrLf

	sBdy = sBdy & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
	sBdy = sBdy & "■すべてがつながる「しごとナビ」（リス株式会社）" & vbCrLf
	sBdy = sBdy & HTTP_CURRENTURL & vbCrLf
	sBdy = sBdy & "お問い合わせ：lis@lis21.co.jp" & vbCrLf

	rSbj = sSbj
	rBdy = sBdy
End Function
%>
