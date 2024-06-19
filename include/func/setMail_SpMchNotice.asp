<%
'*******************************************************************************
'概　要：企業からのメール着信通知メールの文言を生成
'引　数：vCompanyName	：通知先企業名
'　　　：vPersonName	：通知先企業求人担当者
'　　　：vLinkOrder		：対象求人票へのリンク
'　　　：vLinkProfile	：マッチング人材のプロフィールへのリンク群
'　　　：rSubject		：[OUTPUT]件名
'　　　：rBody			：[OUTPUT]本文
'戻り値：Boolean
'備　考：
'更　新：2009/08/14 LIS K.Kokubo 作成
'*******************************************************************************
Function setMail_SpMchNotice(ByVal vCompanyName, ByVal vPersonName, ByVal vLinkOrder, ByVal vLinkProfile, ByRef rSubject, ByRef rBody)
	setMail_SpMchNotice = False

	rSubject = ""
	rSubject = rSubject & "■しごとナビ■適材条件にマッチした求職者のお知らせ"

	rBody = ""
	rBody = rBody & ""
	rBody = rBody & vCompanyName & vbCrLf
	rBody = rBody & vPersonName & "様" & vbCrLf
	rBody = rBody & vbCrLf

いつもご利用ありがとうございます。
総合求人求職サイト「しごとナビ」（リス株式会社）です。

求人の適材条件にマッチした求職者がおりましたのでご連絡いたします。

End Function
%>
