<%
'**********************************************************************************************************************
'概　要：ＬＩＳジャーナルで使用する関数群
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　メール一覧ページ出力用　■■■
'　　　：GetHtmlJNLCharge			：ＬＩＳジャーナルの問合せ担当ＨＴＭＬを取得
'　　　：GetHtmlJNLInquiryBody		：ＬＩＳジャーナルの問合せ内容ＨＴＭＬを取得
'　　　：GetMailBodyToCompany		：ＬＩＳジャーナルの問合せサンクスメール
'　　　：GetMailBodyToLis			：ＬＩＳジャーナルの問合せ受付通知メール
'**********************************************************************************************************************


'******************************************************************************
'概　要：ＬＩＳジャーナルの問合せ担当ＨＴＭＬを取得
'引　数：vBranchName	：拠点名
'　　　：vEmployeeName	：社員名
'　　　：vTel			：拠点電話番号
'　　　：vFax			：拠点FAX番号
'備　考：
'使用元：ナビ/mailservice/lisjournal/inquiry.asp
'更　新：2007/08/30 LIS K.Kokubo
'******************************************************************************
Function GetHtmlJNLCharge(ByVal vBranchName, ByVal vEmployeeName, ByVal vTel, ByVal vFax)
	Dim sHTML

	sHTML = ""
	sHTML = sHTML & "<div style=""margin-top:25px;"">"
	sHTML = sHTML & "<table border=""0"" style=""width:600px;"">"
	sHTML = sHTML & "<tbody>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""width:100px; padding:5px 0px; border-bottom:1px dashed #666666;"">担当者</td>"
	sHTML = sHTML & "<td style=""width:500px; padding:5px 0px; border-bottom:1px dashed #666666;"">リス株式会社　" & vBranchName & "　" & vEmployeeName & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""width:100px; padding:5px 0px; border-bottom:1px dashed #666666;"">連絡先</td>"
	sHTML = sHTML & "<td style=""width:500px; padding:5px 0px; border-bottom:1px dashed #666666;"">TEL:" & vTel & "　FAX:" & vFAX & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & vbCrLf

	GetHtmlJNLCharge = sHTML
End Function

'******************************************************************************
'概　要：ＬＩＳジャーナルの問合せ内容ＨＴＭＬを取得
'引　数：vBody	：問合せ内容
'備　考：
'使用元：ナビ/mailservice/lisjournal/inquiry.asp
'更　新：2007/08/30 LIS K.Kokubo
'******************************************************************************
Function GetHtmlJNLInquiryBody(ByVal vBody)
	Dim sHTML

	sHTML = ""
	If Len(vBody) > 0 Then
		sHTML = sHTML & "<table class=""pattern1"" border=""0"" style=""width:600px;"">"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th>ご質問・ご要望など</th>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</thead>"
		sHTML = sHTML & "<tbody>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>" & Replace(vBody, vbCrLf, "<br>") & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table><br>"
		sHTML = sHTML & vbCrLf
	End If

	GetHtmlJNLInquiryBody = sHTML
End Function

'******************************************************************************
'概　要：ＬＩＳジャーナルの問合せサンクスメール
'引　数：
'備　考：
'使　用：ナビ/mailservice/lisjournal/inquiry.asp
'更　新：2007/08/30 LIS K.Kokubo 作成
'******************************************************************************
Function GetMailBodyToCompany(ByVal vCompanyName, ByVal vCertify, ByVal vBranchName, ByVal vTel, ByVal vFax)
	Dim sBody

	sBody = ""
	sBody = sBody & vCompanyName & "　様" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "――――――――――――――――――――――――――――――――" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "ＬＩＳジャーナルＷへお問い合わせ頂きまして誠にありがとうございます。" & vbCrLf
	sBody = sBody & "弊社担当からの連絡をお待ちください。" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■お問い合わせ頂いた求職者" & vbCrLf
	sBody = sBody & HTTP_NAVI_CURRENTURL & "LJW/inquiry.asp?certify=" & vCertify & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■弊社担当" & vbCrLf
	sBody = sBody & "部　署：" & vBranchName & vbCrLf
	sBody = sBody & "連絡先：TEL " & vTel & "　FAX " & vFax & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "――――――――――――――――――――――――――――――――" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "・転職サイト「しごとナビ」　　　　　　　http://www.shigotonavi.co.jp/" & vbCrLf
	sBody = sBody & "　　　　　　「しごとナビモバイル」　　　http://m.shigotonavi.jp/" & vbCrLf
	sBody = sBody & "　人事専用サイト「しごとナビ人材採用」　http://jinzai.shigotonavi.co.jp/" & vbCrLf
	sBody = sBody & "・人材派遣や人材紹介など、人材に関する各種ご相談もお受けいたします。" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "リス株式会社［総合人材サービス］" & vbCrLf
	sBody = sBody & "〒163-0825　東京都新宿区西新宿２丁目４番１号 新宿ＮＳビル２５階" & vbCrLf
	sBody = sBody & "03-5909-4120" & vbCrLf
	sBody = sBody & "lis@lis21.co.jp" & vbCrLf

	GetMailBodyToCompany = sBody
End Function

'******************************************************************************
'概　要：ＬＩＳジャーナルの問合せ受付通知メール
'引　数：vCertify			：問合せコード
'　　　：vCompanyName		：問合せ企業名
'　　　：vLisBranchName		：リス担当拠点名　[1]追加
'　　　：vLisEmployeeName	：リス担当者名　[1]追加
'備　考：
'******************************************************************************
Function GetMailBodyToLis(ByVal vCertify, ByVal vCompanyName, ByVal vLisBranchName, ByVal vLisEmployeeName, ByVal vStaffLisBranchName, ByVal vStaffLisEmployeeName, _
    ByVal vDeliveryDay, _
    ByVal vCompanyCode)

	Dim sBody

	sBody = ""
	sBody = sBody & "ＬＩＳジャーナルより企業から問合せがありました。" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "――――――――――――――――――――――――――――――――" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■問合せ企業" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & vCompanyName & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■リス担当者" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "　企業担当　：" & vLisBranchname & "(" & vLisEmployeeName & ")" & vbCrLf
	sBody = sBody & "　求職者担当：" & vStaffLisBranchname & "(" & vStaffLisEmployeeName & ")" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■問合せ内容" & vbCrLf
	sBody = sBody & "　以下のリンクからお問い合わせの詳細を見ることができます。" & vbCrLf
    sBody = sBody & "　" & HTTP_BI_CURRENTURL & "LJW/DeliveryRecord/InquiryDetail.asp?DeliveryDay=" & vDeliveryDay & "&companycode=" & vCompanyCode

	GetMailBodyToLis = sBody
End Function
%>
