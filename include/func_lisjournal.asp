<%
'**********************************************************************************************************************
'概　要：ＬＩＳジャーナルで使用する関数群
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　メール一覧ページ出力用　■■■
'　　　：ChgLisJournalWorkStartDay	：ＬＩＳジャーナルの勤務開始予定日変換
'　　　：GetHtmlJNLStaffDetail		：ＬＩＳジャーナルの求職者一人分のテーブルを取得
'　　　：GetHtmlJNLCharge			：ＬＩＳジャーナルの問合せ担当ＨＴＭＬを取得
'　　　：GetHtmlJNLInquiryBody		：ＬＩＳジャーナルの問合せ内容ＨＴＭＬを取得
'　　　：GetMailBodyToCompany		：ＬＩＳジャーナルの問合せサンクスメール
'　　　：GetMailBodyToLis			：ＬＩＳジャーナルの問合せ受付通知メール
'**********************************************************************************************************************

'******************************************************************************
'概　要：ＬＩＳジャーナルの勤務開始予定日変換
'引　数：
'備　考：
'使　用：ナビ/mailservice/lisjournal/inquiry.asp
'更　新：2007/09/10 LIS K.Kokubo 作成
'******************************************************************************
Function ChgLisJournalWorkStartDay(ByVal vBranchCode, ByVal vWorkStartDay)
	On Error Resume Next
	Dim dWorkStartDay
	Dim sWorkStartDay

	ChgLisJournalWorkStartDay = ""

	If IsDate(vWorkStartDay) = True Then
		dWorkStartDay = CDate(vWorkStartDay)
		sWorkStartDay = Year(dWorkStartDay) & "/" & Month(dWorkStartDay) & "/" & Day(dWorkStartDay)

		If DateDiff("d", dWorkStartDay, Date) < 0 Then
			ChgLisJournalWorkStartDay = Year(dWorkStartDay) & "年" & Month(dWorkStartDay) & "月" & Day(dWorkStartDay) & "日"
		Else
			ChgLisJournalWorkStartDay = "即日可能"
		End If
	ElseIf vBranchCode = "OR" Then
		ChgLisJournalWorkStartDay = "入社日ご相談"
	Else
		ChgLisJournalWorkStartDay = "条件次第"
	End If
End Function

'******************************************************************************
'概　要：ＬＩＳジャーナルの求職者一人分のテーブル行を取得
'引　数：rDB		：ＤＢ接続
'　　　：rJNLStaff	：配信番号
'　　　：vFlag		：問合せリンク表示フラグ ["1"]表示 [<>""]非表示
'使　用：ナビ/mailservice/lisjournal/inquiry.asp
'備　考：
'更　新：2007/08/21 LIS K.Kokubo
'　　　：2009/01/14 LIS K.Kokubo ＬＩＳジャーナル求職者クラスを引数にして情報を取得するように変更
'　　　：2009/02/17 LIS K.Kokubo 変更 希望職種、希望年収、学歴の追加対応
'******************************************************************************
Function GetHtmlJNLStaffDetail(ByRef rDB, ByRef rJNLStaff, ByVal vFlag)
	Dim sHTML
	Dim sInquiry

	If vFlag = "1" Then
		sInquiry = "<form action="""" method=""post"" style=""display:inline;""><input type=""submit"" value=""配信解除""><input type=""hidden"" name=""frmdelstaffcode"" value=""" & rJNLStaff.StaffCode & """></form>"
	End If

	sHTML = sHTML & "<table border=""0"" style=""width:600px; border-collapse:collapse;"">"
	sHTML = sHTML & "<thead>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "<th style=""width:100px; padding:0px;""></th>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</thead>"
	sHTML = sHTML & "<tbody>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th colspan=""4"" style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; border-width:1px 0px 1px 0px; font-size:12px; line-height:18px; text-align:left; font-family:'ＭＳ Ｐゴシック';"">"
	If rJNLStaff.HopeJobType <> "" Then
		sHTML = sHTML & rJNLStaff.HopeJobType & "希望&nbsp;(" & rJNLStaff.StaffCode & ")<br>"
	Else
		sHTMl = sHTML & "(" & rJNLStaff.StaffCode & ")&nbsp;"
	End If

	If rJNLStaff.Age <> "" Then
		sHTML = sHTML & rJNLStaff.Age & "・" & rJNLStaff.Sex & "・" & rJNLStaff.Address
	Else
		sHTML = sHTML & rJNLStaff.Address
	End If

	sHTML = sHTML & "</th>"
	sHTML = sHTML & "<th colspan=""2"" style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; border-width:1px 1px 1px 0px; font-size:12px; line-height:18px; text-align:right; font-family:'ＭＳ Ｐゴシック';"">" & sInquiry & "</th>"
	sHTML = sHTML & "</tr>"

	If rJNLStaff.BranchCode <> "TG" And rJNLStaff.Age <> "" Then
		'本社営業統括部の場合は「勤務開始予定」を非表示
		'派遣用ＬＩＳジャーナルの場合は非表示（年齢が空）
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'ＭＳ Ｐゴシック';"">勤務開始予定日</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'ＭＳ Ｐゴシック';"">" & rJNLStaff.ChgWorkStartDay() & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'ＭＳ Ｐゴシック';"">推薦コメント</th>"
	sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'ＭＳ Ｐゴシック';"">" & rJNLStaff.CounselingView & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'ＭＳ Ｐゴシック';"">経験(年数)</th>"
	sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'ＭＳ Ｐゴシック';"">" & rJNLStaff.CareerHistory & "</td>"
	sHTML = sHTML & "</tr>"
	If rJNLStaff.Skill <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'ＭＳ Ｐゴシック';"">スキル</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'ＭＳ Ｐゴシック';"">" & rJNLStaff.Skill & "</td>"
		sHTML = sHTML & "</tr>"
	End If
	If rJNLStaff.RecentConditions <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'ＭＳ Ｐゴシック';"">就職活動状況</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'ＭＳ Ｐゴシック';"">" & rJNLStaff.RecentConditions & "</td>"
		sHTML = sHTML & "</tr>"
	End If
	If rJNLStaff.Hope <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'ＭＳ Ｐゴシック';"">希望・こだわり</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'ＭＳ Ｐゴシック';"">" & rJNLStaff.Hope & "</td>"
		sHTML = sHTML & "</tr>"
	End If
	If rJNLStaff.HopeYearlyIncome <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'ＭＳ Ｐゴシック';"">希望年収</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'ＭＳ Ｐゴシック';"">" & rJNLStaff.HopeYearlyIncome & "</td>"
		sHTML = sHTML & "</tr>"
	End If
	If rJNLStaff.EducateHistory <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""padding:3px; border:1px solid #000000; color:#333333; background-color:#bceaa8; border-color:#ccffcc #8fbb79 #8fbb79 #ccffcc; font-size:12px; line-height:18px; text-align:center; font-family:'ＭＳ Ｐゴシック';"">学歴</th>"
		sHTML = sHTML & "<td colspan=""5"" style=""width:493px; padding:3px; border:1px solid #000000; border-color: #cccccc #999999 #999999 #cccccc; background-color:#f2f1e3; font-size:12px; line-height:18px; font-family:'ＭＳ Ｐゴシック';"">" & rJNLStaff.EducateHistory & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table><br>"

	GetHtmlJNLStaffDetail = sHTML
End Function

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
Function GetMailBodyToCompany(ByVal vCompanyName, ByVal vPersonName, ByVal vCertify, ByVal vBranchName, ByVal vTel, ByVal vFax)
	Dim sBody

	sBody = ""
	sBody = sBody & vCompanyName & "　" & vPersonName & "様" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "――――――――――――――――――――――――――――――――" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "ＬＩＳジャーナルへお問い合わせ頂きまして誠にありがとうございます。" & vbCrLf
	sBody = sBody & "弊社担当からの連絡をお待ちください。" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■お問い合わせ頂いた求職者" & vbCrLf
	sBody = sBody & HTTP_NAVI_CURRENTURL & "mailservice/lisjournal/inquiry.asp?certify=" & vCertify & vbCrLf
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
'　　　：vPersonName		：問合せ担当者名
'　　　：vLisBranchName		：リス担当拠点名　[1]追加
'　　　：vLisEmployeeName	：リス担当者名　[1]追加
'備　考：
'使　用：ナビ/mailservice/lisjournal/inquiry.asp
'更　新：2007/08/30 LIS K.Kokubo 作成
'　　　：2007/09/14 LIS K.Kokubo [1]　求職者担当情報を追加
'******************************************************************************
Function GetMailBodyToLis(ByVal vCertify, ByVal vCompanyName, ByVal vPersonName, ByVal vLisBranchName, ByVal vLisEmployeeName, ByVal vStaffLisBranchName, ByVal vStaffLisEmployeeName)
	Dim sBody

	sBody = ""
	sBody = sBody & "ＬＩＳジャーナルより企業から問合せがありました。" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "――――――――――――――――――――――――――――――――" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■問合せ企業" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & vCompanyName & "(" & vPersonName & ")" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■リス担当者" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "　企業担当　：" & vLisBranchname & "(" & vLisEmployeeName & ")" & vbCrLf
	sBody = sBody & "　求職者担当：" & vStaffLisBranchname & "(" & vStaffLisEmployeeName & ")" & vbCrLf
	sBody = sBody & vbCrLf
	sBody = sBody & "■問合せ内容" & vbCrLf
	sBody = sBody & "　以下のリンクからお問い合わせの詳細を見ることができます。" & vbCrLf
	sBody = sBody & "　" & HTTP_BI_CURRENTURL & "mailservice/lisjournal/inquiry/detail.asp?certify=" & vCertify & vbCrLf

	GetMailBodyToLis = sBody
End Function
%>
