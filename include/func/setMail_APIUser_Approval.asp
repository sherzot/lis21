<%
'*******************************************************************************
'概　要：APIキー承認完了時のメール内容
'引　数：vAPIKey：APIキー
'　　　：rSbj	：[OUTPUT]メール件名
'　　　：rBdy	：[OUTPUT]メール本文
'戻り値：Boolean
'備　考：
'履　歴：2012/02/13 LIS K.Kokubo 作成
'*******************************************************************************
Function setMail_APIUser_Approval(ByVal vAPIKey,ByRef rSbj,ByRef rBdy)
	Dim sSbj,sBdy

	setMail_APIUser_Approval = False

	sSbj = "■しごとナビ■しごとナビ求人検索APIのAPIキー承認完了"

	sBdy = ""
	sBdy = sBdy & "しごとナビ求人検索APIのご利用ありがとうございます。" & vbCrLf
	sBdy = sBdy & "しごとナビサポートです。" & vbCrLf & vbCrLf

	sBdy = sBdy & "APIキーの承認が完了しました。" & vbCrLf
	sBdy = sBdy & "あなたのAPIキーは以下の通りです。" & vbCrLf & vbCrLf

	sBdy = sBdy & vAPIKey & vbCrLf & vbCrLf

	sBdy = sBdy & "このメールは大切に保管してください。" & vbCrLf & vbCrLf

	sBdy = sBdy & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
	sBdy = sBdy & "■すべてがつながる「しごとナビ」（リス株式会社）" & vbCrLf
	sBdy = sBdy & HTTP_CURRENTURL & vbCrLf
	sBdy = sBdy & "お問い合わせ：lis@lis21.co.jp" & vbCrLf

	rSbj = sSbj
	rBdy = sBdy

	setMail_APIUser_Approval = True
End Function
%>
