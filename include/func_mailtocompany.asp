<%
'******************************************************************************
'概　要：メールの署名を取得
'作　者：2007/06/25 Lis K.Kokubo
'引　数：vType				：メール送信先種別 ["1"]一般の求人広告 ["2"]リス案件 ["3"]求人票が紐付いていないリス案件
'　　　：vStaffCode			：メール送信元求職者コード
'　　　：vCompanyName		：メール送信先企業名
'　　　：vOrderCode			：メールに紐づいた情報コード
'　　　：vContactPersonName	：メール送信先案件担当者
'　　　：vSubject			：メール件名
'　　　：vBody				：メール内容
'　　　：vMailURL			：求職者情報参照先ＵＲＬ
'　　　：vHeader			：メール内容に追加するヘッダ
'　　　：vFooter			：メール内容に追加するフッタ
'戻り値：
'備　考：
'使用元：しごとナビ/staff/mailtocompany.asp
'更　新：
'******************************************************************************
Function GetMailBodyStaff(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter)
	Dim sBody
	Dim iLen

	sBody = ""
	Select Case vType
		Case "1":
			'一般の求人広告
			'企業名＋求人票担当者名＋様
			sBody = vCompanyname & "　" & vContactPersonName & "様" & vbCrLf & vbCrLf
		Case "2":
			'リス案件
			'リス社員名＋様
			sBody = vContactPersonName & "様" & vbCrLf & vbCrLf
			vMailURL = "http://bi.lis21.co.jp/staff/staff_MailHistory.asp?Mail=pop&newopen=1"
		Case Else:
			'求人票が紐付いていない
			If Left(vCompanyCode, 1) = "L" Then
				vMailURL = "http://bi.lis21.co.jp/staff/staff_MailHistory.asp?Mail=pop&newopen=1"
			End If
	End Select
	sBody = sBody & vHeader & vbCrLf & vMailURL & vbCrLf

	'メール内容表示処理
	iLen = Len(vBody) * 0.3
	sBody = sBody & vbCrLf & vbCrLf & _
		"-----------------------■　メール情報　■-------------------------" & vbCrLf & _
		"【送信者コード】" & vStaffCode & vbCrLf & _
		"【対象情報コード】" & vOrderCode & vbCrLf & _
		"【メールタイトル】" & vSubject & vbCrLf & _
		"【メール内容】" & vbCrLf & vbCrLf & Left(vBody, iLen) & "..." & vbCrLf & _
		"------------------------------------------------------------------" & vbCrLf
	sBody = sBody & vFooter

	GetMailBodyStaff = sBody

End Function


Function GetMailBodyStaff2(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter)
	Dim sBody2
	Dim iLen
	Dim sName
	Dim sName_1
	Dim sName_2
	Dim sName_F
	Dim sName_F_1
	Dim sName_F_2
	Dim sOperateClassWebCode
	Dim sHopeWorkStartDay

	
	
	
	sSQL = "EXEC up_DtlStaff '" & G_USERID & "';"
flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
If GetRSState(oRS) = True Then
	sOperateClassWebCode = oRS.Collect("OperateClassWebCode")
	sHopeWorkStartDay = GetDateStr(oRS.Collect("HopeWorkStartDay"), "")
	sName = oRS.Collect("Name")
	If InStr(sName, "　") <> 0 Then
		sName_1 = Mid(sName, 1, InStr(sName, "　") - 1)
		sName_2 = Mid(sName, InStr(sName, "　") + 1)
	Else
		sName_1 = sName
	End If
	sName_F = oRS.Collect("Name_F")
	If InStr(sName_F, "　") <> 0 Then
		sName_F_1 = Mid(sName_F, 1, InStr(sName_F, "　") - 1)
		sName_F_2 = Mid(sName_F, InStr(sName_F, "　") + 1)
	Else
		sName_F_1 = sName_F
	End If
	
End If


	sBody =  sName_1 & "" & sName_2 & "　様" & vbCrLf & vbCrLf & _
					 "この度は『しごとナビ』の転職情報をご利用頂きまして、" & vbCrLf & _
					 "まことにありがとうございます。" & vbCrLf & vbCrLf & _
					 "以下の内容でご応募が完了いたしました。"& vbCrLf & vbCrLf & _
					 "担当者よりご連絡差し上げますので、"& vbCrLf & _
					 "今しばらくお待ちください。"& vbCrLf
		
	'メール内容表示処理
	iLen = Len(vBody) * 1
	sBody = sBody & vbCrLf & vbCrLf & _
		"-----------------------■　ご応募情報　■-------------------------" & vbCrLf & _
		"【求人情報コード】" & vbCrLf & _
		vOrderCode & vbCrLf & vbCrLf & _
		"【求人ページ】" & vbCrLf & _
		"http://www.shigotonavi.co.jp/order/order_detail.asp?ordercode="& vOrderCode & vbCrLf & vbCrLf & _
		"【メール内容】"& vbCrLf & _
		"タイトル：" & vbCrLf & _
		 vSubject & vbCrLf & vbCrLf & _
		 "本文：" & vbCrLf & _
		 Left(vBody, iLen) &  vbCrLf & _
		"-------------------------------------------------------------" & vbCrLf & vbCrLf & _
		"ご応募ありがとうございます。" & vbCrLf & vbCrLf & _
		"情報コード【" & vOrderCode & "】へのご応募を受け付けました。" & vbCrLf & vbCrLf & vbCrLf & _
		"《ご注意》" & vbCrLf & _
		"2週間以上返事が無い場合、応募内容の修正・取り消しを希望される方は" & vbCrLf & vbCrLf & _
		"https://www.shigotonavi.co.jp/staff/access.asp" & vbCrLf &vbCrLf & _
		"のお問合わせフォームより、【求人情報コード】を記載のうえご連絡下さい。" & vbCrLf
	sBody = sBody & vFooter

	GetMailBodyStaff2 = sBody2

End Function



'******************************************************************************
'概　要：メールの署名を取得
'作　者：2007/06/25 Lis K.Kokubo
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserID	：ログイン中のユーザ種類
'戻り値：
'備　考：
'使用元：しごとナビ/staff/mailtoperson.asp
'更　新：
'******************************************************************************
Function GetMailSignatureStaff(ByRef rDB, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	sSQL = "sp_GetDataMailSignatureStaff '" & vUserID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	sSignature = ""
	If GetRSState(oRS) = True Then
		sSignature = sSignature & "------------------------------" & vbCrLf
	'2009/05/01　無料化に合わせて署名を氏名だけにしてみた　安藤
		'sSignature = sSignature & "住所：" & oRS.Collect("Prefecture") & oRS.Collect("City") & oRS.Collect("Town") & oRS.Collect("Address") & vbCrLf
		sSignature = sSignature & "氏名：" & oRS.Collect("Name") & vbCrLf
		'If oRS.Collect("HomeContactFlag") = "1" Then
		'	sSignature = sSignature & "自宅：" & oRS.Collect("HomeTelephoneNumber") & vbCrLf
		'End If
		'If oRS.Collect("PortableContactFlag") = "1" Then
		'	sSignature = sSignature & "携帯：" & oRS.Collect("PortableTelephoneNumber") & vbCrLf
		'End If
		'If oRS.Collect("FaxContactFlag") = "1" Then
		'	sSignature = sSignature & "FAX ：" & oRS.Collect("FaxNumber") & vbCrLf
		'End If
		'If oRS.Collect("MailContactFlag") = "1" Then
		'	sSignature = sSignature & "Mail：" & oRS.Collect("MailAddress") & vbCrLf
		'End If
	End If
	Call RSClose(oRS)

	GetMailSignatureStaff = sSignature
End Function

'******************************************************************************
'概　要：求職者のメールテンプレート取得
'作　者：2007/06/25 Lis K.Kokubo
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserType	：ログイン中のユーザ種類
'　　　：vSEQ		：番号
'戻り値：
'備　考：
'使用元：しごとナビ/staff/mailtocompany.asp
'更　新：
'******************************************************************************
Function GetStaffMailTemplateOptionHtml(ByRef rDB, ByVal vUserType, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sSelected

	GetStaffMailTemplateOptionHtml = ""

	sSQL = "sp_GetDataMailTemplate '1'"

	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sSelected = ""
		If CStr(oRS.Collect("Cd")) = CStr(vSEQ) Then sSelected = "selected=""true"""
		GetStaffMailTemplateOptionHtml = GetStaffMailTemplateOptionHtml & _
			"<option value=""" & oRS.Collect("Cd") & """ " & sSelected & ">" & oRS.Collect("Title") & "</option>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：応募メールかどうかをチェック
'引　数：vText	：判定する文字列
'戻り値：
'備　考：マッチング人材応募,適材待機マッチング人材応募の判定に使用する
'履　歴：2009/08/13 LIS K.Kokubo
'******************************************************************************
Function ChkMailResponse(ByVal vText)
	Dim sTrueText
	Dim sFalseText

	ChkMailResponse = False

	sTrueText = ""
	sTrueText = sTrueText & "(応募)"
	sTrueText = sTrueText & "|(一度お話をお聞かせいただきたく)" 'モバイルテンプレート:求人へ応募
	sTrueText = sTrueText & "|(お話を伺いたく存じます)" 'モバイルテンプレート:返信(応募)

	sFalseText = ""
	sFalseText = sFalseText & "(辞退)"
	sFalseText = sFalseText & "|(求人に応募させて頂き)"

	If IsRE(vText, sTrueText, True) = True And IsRE(vText, sFalseText, True) = False Then 
		ChkMailResponse = True
	End If
End Function

'NEOのスタッフ用返信メール本文作成
Function GetMailBodyStaffNEO(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter)
	Dim sBody
	Dim iLen
	Dim sName
	Dim sName_1
	Dim sName_2
	Dim sName_F
	Dim sName_F_1
	Dim sName_F_2
	Dim sOperateClassWebCode
	Dim sHopeWorkStartDay

	
	sSQL = "EXEC up_DtlStaff '" & G_USERID & "';"
    flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
    If GetRSState(oRS) = True Then
	    sOperateClassWebCode = oRS.Collect("OperateClassWebCode")
	    sHopeWorkStartDay = GetDateStr(oRS.Collect("HopeWorkStartDay"), "")
	    sName = oRS.Collect("Name")
	    If InStr(sName, "　") <> 0 Then
	    	sName_1 = Mid(sName, 1, InStr(sName, "　") - 1)
	    	sName_2 = Mid(sName, InStr(sName, "　") + 1)
	    Else
	    	sName_1 = sName
	    End If
	    sName_F = oRS.Collect("Name_F")
	    If InStr(sName_F, "　") <> 0 Then
	    	sName_F_1 = Mid(sName_F, 1, InStr(sName_F, "　") - 1)
	    	sName_F_2 = Mid(sName_F, InStr(sName_F, "　") + 1)
	    Else
	    	sName_F_1 = sName_F
	    End If
	
    End If
    Call RSClose(oRS)
    
    sBody = ""
    sSQL = "EXEC sp_GetData_NEO_C_Reply '" & vOrderCode & "';"
    flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
    If GetRSState(oRS) = True Then
        If oRS.Collect("ReplyFlag") = "0" Then
            If oRS.Collect("Reply") <> "" Then
                sBody = oRS.Collect("Reply")
                sBody = sBody & vbCrLf & vbCrLf & vbCrLf
                sBody = sBody & vbCrLf & vbCrLf & _
		        "-----------------------■　ご応募情報　■-------------------------" & vbCrLf & _
		        "【企業名】" & vbCrLf & _
		        vCompanyName & vbCrLf & vbCrLf & _
	    	    "【求人ページ】" & vbCrLf & _
		        "http://www.shigotonavi.co.jp/order/order_detail.asp?ordercode="& vOrderCode & vbCrLf & vbCrLf & _
	    	    "【応募内容】"& vbCrLf & _
		        Left(vBody, iLen) &  vbCrLf & _
		        "-------------------------------------------------------------" & vbCrLf & vbCrLf & vbCrLf & _
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & _
                "■「しごとナビ」でGポイントを貯めよう！" & vbCrLf & vbCrLf & _
                "------------------------------------------------------------------" & vbCrLf & vbCrLf & _
                "「しごとナビ」では、お近くの弊社オフィスで面談登録をされたり、"& vbCrLf & _
                "対象の掲載求人で転職が決まると、" & vbCrLf & _
                "【最大１万Gポイント】をプレゼントいたします！" & vbCrLf & _
                "※ポイント付与には条件があります。" & vbCrLf & vbCrLf & _
                "様々なご利用が対象ですのでぜひチェックしてください。" & vbCrLf & _
                "ご利用の後に【Gポイント申請】もお忘れなく！" & vbCrLf & vbCrLf & _
                "詳しくはこちら" & vbCrLf & _
                "https://www.shigotonavi.co.jp/point/pr/"& vbCrLf & vbCrLf & _
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & vbCrLf & _
		        "" & vCompanyName & "へのご応募を受け付けました。" & vbCrLf & vbCrLf & vbCrLf & _
		        "ご応募ありがとうございました。" & vbCrLf & vbCrLf & vbCrLf & _

                "※こちらのアドレスは送信専用となっております。" & vbCrLf & _
                "※弊社へのご連絡は「lis@lis21.co.jp」にお願いいたします。" & vbCrLf
                sBody = sBody & "-------------------------------" & vbCrLf
                sBody = sBody & "はたらく人のソーシャルコミュニティー「しごとナビ」" & vbCrLf
                sBody = sBody & "運営会社：リス株式会社" & vbCrLf
                sBody = sBody & "http://www.shigotonavi.co.jp/" & vbCrLf
                sBody = sBody & "お問い合わせ：lis@lis21.co.jp" & vbCrLf
            End If
        End If
    End If
    Call RSClose(oRS)

    if sBody = "" Then
	    sBody =  sName_1 & "" & sName_2 & "　様" & vbCrLf & vbCrLf & _
					 "この度は『しごとナビ』の転職情報をご利用頂きまして、" & vbCrLf & _
					 "まことにありがとうございます。" & vbCrLf & vbCrLf & _
					 "以下の内容でご応募が完了いたしました。"& vbCrLf & vbCrLf

		sBody = sBody & "通常３日以内に企業担当者から連絡がございますので、今しばらくお待ちください。" & vbCrLf
        sBody = sBody & "（GWや年末年始などの大型連休時には、返答が遅くなることもあります）" & vbCrLf & vbCrLf

        '動的に企業担当者の電話番号とメールアドレスドメインを挿入
        Dim aryStrings
        sSQL = "EXEC sp_GetData_C_Contact '" & vOrderCode & "';"
        flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
        If GetRSState(oRS) = True Then
            aryStrings = Split(oRS.Collect("MailAddress"), "@")
            IF ChkStr(oRS.Collect("TelNumber")) <> "" Then
                sBody = sBody & "企業からの連絡は、【" & oRS.Collect("TelNumber") & "】からのお電話、" & vbCrLf
            Else
                sBody = sBody & "企業からの連絡は、"
            End if
            sBody = sBody & "【@" & aryStrings(1) & "】からのメールとなります。" & vbCrLf

            IF ChkStr(oRS.Collect("TelNumber")) <> "" Then
                sBody = sBody & "着信拒否や、迷惑メール扱いとされないよう、ご注意ください。" & vbCrLf
            Else
                sBody = sBody & "迷惑メール扱いとされないよう、ご注意ください。" & vbCrLf
            End if
            sBody = sBody & "（携帯メールアドレスをご登録されている方は、メール受信設定状態をご確認ください）" & vbCrLf & vbCrLf
        End If
        Call RSClose(oRS)

        sBody = sBody & "また、弊社より状況確認をさせていただく事もございます。" & vbCrLf
        sBody = sBody & "【「しごとナビ」エージェントNEO運営事務局】から届くメールには、ご返信くださるようお願いいたします。"

	    'メール内容表示処理
	    iLen = Len(vBody) * 1
    	sBody = sBody & vbCrLf & vbCrLf & _
		    "-----------------------■　ご応募情報　■-------------------------" & vbCrLf & _
		    "【企業名】" & vbCrLf & _
		    vCompanyName & vbCrLf & vbCrLf & _
	    	"【求人ページ】" & vbCrLf & _
		    "http://www.shigotonavi.co.jp/order/order_detail.asp?ordercode="& vOrderCode & vbCrLf & vbCrLf & _
	    	"【応募内容】"& vbCrLf & _
		    Left(vBody, iLen) &  vbCrLf & _
		    "-------------------------------------------------------------" & vbCrLf & vbCrLf & _
            "■応募履歴について" & vbCrLf & _
            "しごとナビにログインし、Myページ → Myメニュー → 応募一覧 から確認出来ます。" & vbCrLf & vbCrLf & _
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & _
                "■「しごとナビ」でGポイントを貯めよう！" & vbCrLf & vbCrLf & _
                "------------------------------------------------------------------" & vbCrLf & vbCrLf & _
                "「しごとナビ」では、お近くの弊社オフィスで面談登録をされたり、"& vbCrLf & _
                "対象の掲載求人で転職が決まると、" & vbCrLf & _
                "【最大１万Gポイント】をプレゼントいたします！" & vbCrLf & _
                "※ポイント付与には条件があります。" & vbCrLf & vbCrLf & _
                "様々なご利用が対象ですのでぜひチェックしてください。" & vbCrLf & _
                "ご利用の後に【Gポイント申請】もお忘れなく！" & vbCrLf & vbCrLf & _
                "詳しくはこちら" & vbCrLf & _
                "https://www.shigotonavi.co.jp/point/pr/"& vbCrLf & vbCrLf & _
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & vbCrLf & _
		    "" & vCompanyName & "へのご応募を受け付けました。" & vbCrLf & vbCrLf & vbCrLf & _
		    "ご応募ありがとうございました。" & vbCrLf & vbCrLf & vbCrLf & _
            "※こちらのアドレスは送信専用となっております。" & vbCrLf & _
            "※弊社へのご連絡は「lis@lis21.co.jp」にお願いいたします。" & vbCrLf
	    sBody = sBody & vFooter

    End if

	GetMailBodyStaffNEO = sBody

End Function

'NEO広告の用返信メール本文作成
Function GetMailBodyCompanyNEO(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter,ByVal vfrmsubject)
	Dim sBody
	Dim iLen

	sBody = ""
	'一般の求人広告
	'企業名＋求人票担当者名＋様
	sBody = vCompanyname & vbCrLf & vContactPersonName & "様" & vbCrLf & vbCrLf

	'メール内容表示処理
	iLen = Len(vBody) * 0.3
	sBody = sBody & vbCrLf & _
		"お世話になっております。" & vbCrLf & _
        "「しごとナビ」エージェントNEO事務局です。" & vbCrLf & vbCrLf & _
        "平素は「しごとナビ」をご利用くださいまして、" & vbCrLf & _
        "誠にありがとうございます。" & vbCrLf & vbCrLf & _
        "御社の求人情報に、求職者から応募がありましたのでお知らせ致します。" & vbCrLf & vbCrLf & vbCrLf & _
        "【" & vfrmsubject & "】"  & vbCrLf & vbCrLf & _
        "応募者情報の詳細につきましては、" & vbCrLf & _
        "専用の管理画面へログイン後、。" & vbCrLf & _
        "「応募管理」の「応募者一覧」より" & vbCrLf & _
        "ご確認いただけます。" & vbCrLf & vbCrLf & _
        "管理画面はこちら" & vbCrLf & _
        "https://www.shigotonavi.co.jp/management/login.asp" & vbCrLf & vbCrLf & vbCrLf & _
        "※ ご注意 ※" & vbCrLf & _
        "応募者への対応によって、企業イメージに直結することもございます。" & vbCrLf & _
        "応募者へのお問い合わせ対応や応募返信、面接対応、" & vbCrLf & _
        "採用の合否通知などにつきましては、" & vbCrLf & _
        "誠意を込めて迅速にご対応くださいますようお願い致します。" & vbCrLf & vbCrLf & vbCrLf & _
        "今後とも「しごとナビ」をよろしくお願いいたします。" & vbCrLf & vbCrLf & vbCrLf & _
        "はたらく人のソーシャルコミュニティー「しごとナビ」" & vbCrLf & _
        "http://www.shigotonavi.co.jp/" & vbCrLf & vbCrLf & _
        "※こちらのアドレスは送信専用となっております。" & vbCrLf & _
        "※弊社へのご連絡は「lis@lis21.co.jp」にお願いいたします。" & vbCrLf & _
        "───────────────────────────────" & vbCrLf & _
        "■発行元　リス株式会社" & vbCrLf & _
        "〒163-0825　東京都新宿区西新宿4-2-21 新宿NSビル25階" & vbCrLf & _
        "お問合わせ：lis@lis21.co.jp" & vbCrLf & _
        "───────────────────────────────" & vbCrLf & _
        "ネットワーク：東京（新宿）・群馬・静岡・名古屋・大阪・広島・岡山" & vbCrLf

	GetMailBodyCompanyNEO = sBody

End Function
%>
