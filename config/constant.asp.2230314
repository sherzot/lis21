<%
'************************************************
'*	固定変数宣言ファイル						*
'************************************************

'*****メール情報*****
Const Cnt_MailServer = "172.16.1.39"		'メールサーバー名
Const Cnt_LisMailAddress = "しごとナビ・リス <lis@lis21.co.jp>"		'リス様代表メールアドレス
Const Cnt_NaviMailAddress = "しごとナビ <info@shigotonavi.jp>"		'しごとナビ通知用メールアドレス

'******************************************************************************
'** アプローチメールにて使用(staff/mailtocompany.asp)
'******************************************************************************
Dim MAIL_URL_STAFF: MAIL_URL_STAFF = "https://www.shigotonavi.co.jp/staff/mailhistory_person.asp"	'求職者メール管理ページ
Dim MAIL_URL_COMPANY: MAIL_URL_COMPANY = "https://www.shigotonavi.co.jp/company/mailhistory_company.asp"	'求人企業メール管理ページ
Dim MAIL_URL_TALENT: MAIL_URL_TALENT = "https://www.shigotonavi.co.jp/talent/mailhistory_talent.asp"		'人材保有企業メール管理ページ

'************************************************
'*	メール送信時の参照先URL						*
'************************************************

'************************************************
'*	求職者(staff)								*
'************************************************
'新規スタッフ登録時および更新時にリス殿へメール送信（派遣の場合のみ）
'(LISPROJE/staff/RegisterStaff_sendmail.asp)
Const Cnt_RegistMail_StaffURL = "http://bi-b1.lis21.co.jp/INCLUDE/Staff_detail.asp"
'↑本番投入前にCnt_RegistMail_StaffURLは必ず
'"http://bi.lis21.co.jp/INCLUDE/Staff_detail.asp"（本番用）にしてください。(派遣仮登録者の入力通知メール用アドレス)

'************************************************
'*	求人企業(company)							*
'************************************************

'************************************************
'*	人材保有企業(talent)						*
'************************************************
'人材保有企業が会社情報を更新した際にリス殿へメール送信
'(LISPROJE/talent/t_company_regist_sendmail.asp)
'@@@2003/02/27 TASc Uda Del		Cnt_RegistMail_TCompanyURL = "https://www.shigotonavi.co.jp/talent/t_company_regist.asp"

'人材保有企業が人材情報を更新した際にリス殿へメール送信
'(LISPROJE/talent/t_person_reg_sendmail.asp)
'@@@2003/02/27 TASC Uda Del		Cnt_RegistMail_TPersonURL = "https://www.shigotonavi.co.jp/talent/t_person_reg1l.asp"

'******************************************************************************
'** スタッフ登録確認メール(staff/person_reg1_register.asp)
'******************************************************************************
'タイトル
Const MAIL_STAFFREG_SUBJECT = "【しごとナビ】ご登録完了のご案内 ※履歴書作成の手順も掲載"
'メール本文
Dim MAIL_STAFFREG_BODY: MAIL_STAFFREG_BODY = "" & _
	"求人サイト「しごとナビ」を運営しているリス株式会社です。" & vbCrLf & _
	"この度はしごとナビへのご登録ありがとうございました。" & vbCrLf & _
	"貴方様の登録は、無事完了致しました。" & vbCrLf & _
	vbCrLf & _
	"今後、「しごとナビ」のサービスメニューをご利用いただくため、" & vbCrLf & _
	"下記のＩＤとパスワードを発行致します。大切に保管してください。" & vbCrLf & _
	vbCrLf & _
	"【貴方様のＩＤ】━━━━━━━━━━━━━━━━━━━━━━━━━━━"
'メールフッタ
Dim MAIL_STAFFREG_FOOTER: MAIL_STAFFREG_FOOTER = "" & _
	"■「休止・退会」について" & vbCrLf & vbCrLf & _
	"しごとナビのご利用が必要な場合は、上記IDとパスワードでログイン後、" & vbCrLf & _
	"メニュー内の「休止・退会」を押して下さい。" & vbCrLf & vbCrLf & _
	"しごとナビに関してご不明な点がございましたら、" & vbCrLf & _
	"お手数ではございますが下記までメールにてお問合せください。" & vbCrLf & vbCrLf & _
	"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & _
	"はたらく人のソーシャルコミュニティー「しごとナビ」" & vbCrLf & _
	"運営会社：リス株式会社" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"お問い合わせ：lis@lis21.co.jp"

'******************************************************************************
'** アフィリエイト登録者用認証メール(staff/person_reg1_register.asp)
'******************************************************************************
Dim MAIL_STAFFREG_AFFILIATE_HEADER: MAIL_STAFFREG_AFFILIATE_HEADER = "" & _
	"求人ｻｲﾄ「しごとﾅﾋﾞ」のﾘｽ株式会社です。" & vbCrLf & _
	"この度はしごとﾅﾋﾞへのご登録ありがとうございました。" & vbCrLf & _
	"下記のURLをｸﾘｯｸすると、ﾒｰﾙの認証が完了いたします。" & vbCrLf & _
	"--------------------" & vbCrLf & vbCrLf & _
	"認証確定ページ↓" & vbCrLf

Dim MAIL_STAFFREG_AFFILIATE_FOOTER: MAIL_STAFFREG_AFFILIATE_FOOTER = "" & _
	"--------------------" & vbCrLf & _
	"はたらく人のソーシャルコミュニティー「しごとナビ」" & vbCrLf & _
	"運営会社：リス株式会社" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"お問い合わせ：lis@lis21.co.jp"

'******************************************************************************
'** 【求職者から求人企業へのアプローチメール (staff/mailtocompany.aspにて使用)】
'******************************************************************************
'タイトル
Const MAIL_FROM_STAFF_SUBJECT = "【しごとナビ】求職者からメールが届きました"
'メール本文
Dim MAIL_FROM_STAFF_BODY: MAIL_FROM_STAFF_BODY = "" & _
	"いつも「しごとナビ」をご利用くださいましてありがとうございます。" & vbCrLf & vbCrLf & _
	"「しごとナビ」を運営しておりますリス株式会社です。" & vbCrLf & vbCrLf & vbCrLf & _
	"貴社の求人情報へ求職者から応募がありましたのでお知らせ致します。" & vbCrLf & vbCrLf & vbCrLf & _
	"ご応募内容は、下記のURLよりご覧下さい。" & vbCrLf & vbCrLf & _
	"※求職者の方へご連絡等のご対応を宜しくお願い申し上げます。" & vbCrLf & vbCrLf & vbCrLf & _
	"求職者からのメール内容はこちらからご確認ください" & vbCrLf & _
	"↓↓"
'メールフッタ
Dim MAIL_FROM_STAFF_FOOTER: MAIL_FROM_STAFF_FOOTER = "" & vbCrLf & vbCrLf & _
	"-------------------------------" & vbCrLf & _
	"はたらく人のソーシャルコミュニティー「しごとナビ」" & vbCrLf & _
	"運営会社：リス株式会社" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"お問い合わせ：lis@lis21.co.jp"

'******************************************************************************
'** 【求職者から求人企業へのアプローチメール (staff/mailtocompany.aspにて使用)】
'******************************************************************************
'タイトル
Const MAIL_FROM_COMPANY_SUBJECT = "【しごとナビ】求人企業からメールが届きました"
'メール本文
Dim MAIL_FROM_COMPANY_BODY: MAIL_FROM_COMPANY_BODY = "" & _
	"いつもご利用いただきましてありがとうございます。" & vbCrLf & _
	"ご登録いただいております「しごとナビ」（リス株式会社）です。"  & vbCrLf & vbCrLf & _
	"先ほど、求人企業からスカウト・連絡メールをお預かりしました。" & vbCrLf & _
	"￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣" & vbCrLf & _
	"■メールをご覧頂くには下のURL→ログイン→メール管理→メールタイトル" & vbCrLf & _
	"　をクリックして下さい。" & vbCrLf & _
	"■返信は、上記方法でメールをご覧になり「返信」ボタンから、" & vbCrLf & _
	"　メールを作成・送信できます。" & vbCrLf & _
	"■企業からの大切なメールです。ぜひお返事をお願いいたします。" & vbCrLf & vbCrLf & _
	"●気になる「メール内容」と「返信」はこちらをクリックして下さい！"
'メールフッタ
Dim MAIL_FROM_COMPANY_FOOTER: MAIL_FROM_COMPANY_FOOTER = "" & vbCrLf & vbCrLf & _
	"▲このメールに直接返信されても企業には届きませんのでご注意下さい。▲" & vbCrLf & vbCrLf & _
	"◎スカウトが少ない！そんな時は、登録情報が足りないことがあります。" & vbCrLf & _
	"　もう一度登録情報を見直し、追加して下さい。" & vbCrLf & _
	"◎企業へ直接アプローチもできます。ログイン→仕事を検索・表示し、" & vbCrLf & _
	"　「この会社へメールを送信する」ボタンからメールを書き、積極的に" & vbCrLf & _
	"　アプローチして下さい。" & vbCrLf & _
	"◎しごとナビのリスでは、お仕事のご紹介も行なっております。" & vbCrLf & _
	"　お仕事があった際、電話やメールにて連絡する最速サービスもしております。" & vbCrLf & _
	"◎スカウト連絡メールは、しごとナビをご利用のすべての皆様に届く" & vbCrLf & _
	"　チャンスがあります。" & vbCrLf & vbCrLf & _
	"・IDやパスワードなどが分からない時はこちらから入手できます↓" & vbCrLf & _
	"https://www.shigotonavi.co.jp/staff/passwordreminder.asp" & vbCrLf & vbCrLf & _
	"■--------------------------------------------" & vbCrLf & _
	"はたらく人のソーシャルコミュニティー「しごとナビ」" & vbCrLf & _
	"運営会社：リス株式会社" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"お問い合わせ：lis@lis21.co.jp"

'【企業間コラボでのアプローチメール(/LISPROJE/sendmail_company.aspにて使用)】
'タイトル
Const Cnt_sendmail_Advert_Subject = "【しごとナビ】コラボレーションのメールが届きました。"
'メール本文
Dim Cnt_sendmail_Advert_Body
Cnt_sendmail_Advert_Body = "いつもご利用ありがとうございます。" & vbCrLf & _
	"しごとナビです。" & vbCrLf & vbCrLf & _
	"しごとナビのコラボ機能を利用して貴社へメールが届いております。" & vbCrLf & _
	"ご確認の上、お返事などのご対応の程、宜しくお願い申し上げます。" & vbCrLf & _
	"メール内容はこちらから↓"
'メールフッタ
Dim Cnt_sendmail_Advert_Fut
Cnt_sendmail_Advert_Fut = vbCrLf & vbCrLf & "-------------------------------" & vbCrLf & _
	"はたらく人のソーシャルコミュニティー「しごとナビ」" & vbCrLf & _
	"運営会社：リス株式会社" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"お問い合わせ：lis@lis21.co.jp"


'【求職者コラボでのアプローチメール】
'タイトル
Const Cnt_sendmail_Advert2_Subject = "【しごとナビ】コラボレーションのメールが届きました。"
'メール本文
Dim Cnt_sendmail_Advert2_Body
Cnt_sendmail_Advert2_Body = "いつもご利用ありがとうございます。" & vbCrLf & _
	"しごとナビです。" & vbCrLf & _
	"しごとナビをお使いの求職者（もしくは人材保有企業）からあなたへ、" & vbCrLf & _
	"しごとナビコラボレーション機能を利用してメールが届いております。" & vbCrLf & _
	"ご確認の上、お返事などのご対応の程、宜しくお願い申し上げます。" & vbCrLf & _
	"メール内容はこちらから↓"
'メールフッタ
Dim Cnt_sendmail_Advert2_Fut
Cnt_sendmail_Advert2_Fut = vbCrLf & vbCrLf & "※コラボレーション機能" & vbCrLf & _
	"「協業」「共同開発」の意味。しごとナビではＳＯＨＯの方や、" & vbCrLf & _
	"独立志向の方、人材保有企業という仕事を求めていながら、" & vbCrLf & _
	"一方で共同で作業を行なう方などが求人活動を行なうこともできます。" & vbCrLf & _
	"その機能をコラボレーション機能と呼んでおります。" & vbCrLf & vbCrLf & _
	vbCrLf & _
	"はたらく人のソーシャルコミュニティー「しごとナビ」" & vbCrLf & _
	"運営会社：リス株式会社" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"お問い合わせ：lis@lis21.co.jp"


'【求職者コラボレーション申込完了メール(/LISPROJE/Collabo_Entry_Reg.aspにて使用)】
'タイトル
Const Cnt_Collabo_Entry_Subject = "【しごとナビ】求職者コラボレーション申込のお知らせ"
'メール本文
Dim Cnt_Collabo_Entry_Body
Cnt_Collabo_Entry_Body = "求職者コラボレーションの申込がありました。" & vbCrLf & vbCrLf & _
	"担当の方は「社内システム」より" & vbCrLf & _
	"「しごとナビ管理」→「ライセンス情報保守」で" & vbCrLf & _
	"ライセンス発行を行なってください。" & vbCrLf & _
	"（現時点ではライセンスデータは入っていません）" & vbCrLf & _
	"発行後、広告申込書を送付してください。" & vbCrLf

'【求職者コラボレーション継続メール(/LISPROJE/Collabo_Entry_Reg.aspにて使用)】
'タイトル
Const Cnt_Collabo_keizoku_Subject = "【しごとナビ】求職者コラボレーション継続申込"
'メール本文
Dim Cnt_Collabo_keizoku_Body
Cnt_Collabo_keizoku_Body = "求職者コラボレーションの継続申込がありました。" & vbCrLf & _
	"すでに継続ライセンスは発行済みですので、確認と申込書の送付をお願いいたします。" & vbCrLf & vbCrLf & _
	"すでにライセンス情報は入っております。" & vbCrLf & _
	"担当の方は情報確認後、広告申込書を送付してください。" & vbCrLf & _
	"「社内システム」より「しごとナビ管理」→「ライセンス情報保守」" & vbCrLf & _
	"で確認が可能です。" & vbCrLf


'【企業間コラボレーション申込完了メール(/LISPROJE/company/Collabo_Entry_Reg.aspにて使用)】
'タイトル
Const Cnt_Company_Collabo_Entry_Subject = "【しごとナビ】ビジネスコラボレーション申込のお知らせ"
'メール本文
Dim Cnt_Company_Collabo_Entry_Body
Cnt_Company_Collabo_Entry_Body = "ビジネスコラボレーションの申込がありました。" & vbCrLf & vbCrLf & _
	"担当の方は「社内システム」より" & vbCrLf & _
	"「しごとナビ管理」→「ライセンス情報保守」で" & vbCrLf & _
	"ライセンス発行を行なってください。" & vbCrLf & _
	"（現時点ではライセンスデータは入っていません）" & vbCrLf & _
	"発行後、広告申込書を送付してください。" & vbCrLf

'【企業間コラボレーション継続メール(/LISPROJE/company/Collabo_Entry_Reg.aspにて使用)】
'タイトル
Const Cnt_Company_Collabo_keizoku_Subject = "【しごとナビ】ビジネスコラボレーション継続申込のお知らせ"
'メール本文
Dim Cnt_Company_Collabo_keizoku_Body
Cnt_Company_Collabo_keizoku_Body = "ビジネスコラボレーションの継続申込がありました。" & vbCrLf & _
	"すでに継続ライセンスは発行済みですので、確認と申込書送付をお願いいたします。" & vbCrLf & vbCrLf & _
	"すでにライセンス情報は入っております。" & vbCrLf & _
	"担当の方は情報確認後、広告申込書を送付してください。" & vbCrLf & _
	"「社内システム」より「しごとナビ管理」→「ライセンス情報保守」" & vbCrLf & _
	"で確認が可能です。" & vbCrLf

'【求人広告継続利用】
'タイトル
Const Cnt_Company_keizoku_Subject = "【しごとナビ】求人広告利用継続申し込みのお知らせ"
'メール本文
Dim Cnt_Company_keizoku_Body
Cnt_Company_keizoku_Body = "求人広告の利用継続の申込がありました。" & vbCrLf & _
	"すでに継続ライセンスは発行済みですので、確認と申込書送付をお願いいたします。" & vbCrLf & vbCrLf & _
	"すでにライセンス情報は入っております。" & vbCrLf & _
	"担当の方は情報確認後、広告申込書を送付してください。" & vbCrLf & _
	"「社内システム」より「しごとナビ管理」→「ライセンス情報保守」" & vbCrLf & _
	"で確認が可能です。" & vbCrLf


'配信中止ＵＲＬ
Dim Cnt_Jinzai_Stop_URL
Cnt_Jinzai_Stop_URL = HTTP_CURRENTURL & "jinzai/jinzai_stop_reg.asp"

'【メールマガジン登録確認メール(/LISPROJECT/JINZAI/Jinzai_Reg.aspにて使用)】
'タイトル
Dim Cnt_Jinzai_Entry_Subject
Cnt_Jinzai_Entry_Subject = "無料メルマガ「ＪＩＮＺＡＩ」登録確認"
'メール本文
Dim Cnt_Jinzai_Entry_Body
Cnt_Jinzai_Entry_Body = "ご利用ありがとうございます。" & vbCrLf & _
	"無料メルマガ「ＪＩＮＺＡＩ」です。" & vbCrLf & vbCrLf & _
	"メールマガジンの登録確認メールです。" & vbCrLf & _
	"次の内容でメールマガジンを登録いたしました。" & vbCrLf
'メールフッタ
Dim Cnt_Jinzai_Entry_Fut
Cnt_Jinzai_Entry_Fut = vbCrLf & "-------------------------------" & vbCrLf & _
	"お問い合わせ：lis@lis21.co.jp"

'【しごとナビ登録申込メール(/LISPROJECT/JINZAI/Jinzai_EntryNavi_Reg.aspにて使用)】
'タイトル
Dim Cnt_Jinzai_Navi_Subject
Cnt_Jinzai_Navi_Subject 	= "【メルマガしごとナビ申込み】"
'メール本文
Dim Cnt_Jinzai_Navi_Body
Cnt_Jinzai_Navi_Body = "メールマガジンの購読企業からしごとナビ登録の" & vbCrLf & _
	"申込みがありました。" & vbCrLf & vbCrLf & _
	"担当の方はこの企業にアプローチを行ってください。" & vbCrLf & _
	"情報収集後、「社内システム」より" & vbCrLf & _
	"「しごとナビ管理」→「利用者情報保守」で企業情報登録、" & vbCrLf & _
	"　　　　　　　　　→「ライセンス情報保守」でライセンス発行、" & vbCrLf & _
	"　　　　　　　　　→「認証ＩＤ発行」と処理を進め、" & vbCrLf & _
	"企業へ利用開始を連絡してください。" & vbCrLf & vbCrLf & _
	"【申込のあった企業はこちら↓】" & vbCrLf
'メールフッタ
Dim Cnt_Jinzai_Navi_Fut
Cnt_Jinzai_Navi_Fut = ""

'【メールマガジン登録確認メール(/LISPROJECT/JINZAI/Jinzai_Reg.aspにて使用)】
	'タイトル
Dim Cnt_Jinzai_GetID_Subject
Cnt_Jinzai_GetID_Subject = "無料メルマガ「ＪＩＮＺＡＩ」ＩＤの送付"
	'メール本文
Dim Cnt_Jinzai_GetID_Body
Cnt_Jinzai_GetID_Body = vbCrLf & "ご登録ありがとうございます。" & vbCrLf & _
	"人材相場情報「ＪＩＮＺＡＩ」です。" & vbCrLf & vbCrLf & _
	"以下の通りご登録を受け付け致しましたのでＩＤをご送付" & vbCrLf & _
	"させていただきます。" & vbCrLf & vbCrLf & _
	"■貴方様のＩＤ："
	'メールフッタ
Dim Cnt_Jinzai_GetID_Fut
Cnt_Jinzai_GetID_Fut = vbCrLf & "上記ＩＤを以下の方法でご設定下さい。" & vbCrLf & vbCrLf & _
	"1.「ＪＩＮＺＡＩ！」を立ち上げる" & vbCrLf & _
	"2.画面左上の「ファイル」→「設定」をクリック" & vbCrLf & _
	"3.メールマガジンＩＤに上記ＩＤを登録する" & vbCrLf & vbCrLf & _
	"以上で完了です。" & vbCrLf & vbCrLf & _
	"是非、御社に適任の優秀な人材をご獲得下さい。" & vbCrLf & vbCrLf & _
	"-------------------------------" & vbCrLf & _
	"お問い合わせ" & vbCrLf & _
	"リス株式会社　Web戦略室" & vbCrLf & _
	"メールアドレス　lis@lis21.co.jp" & vbCrLf

'【ソフト登録確認メール(/LISPROJECT/JINZAI/JinzaiSoft_Reg2.aspにて使用)】
Dim Cnt_JinzaiSoft_Entry_Body
Cnt_JinzaiSoft_Entry_Body = "ご利用ありがとうございます。" & vbCrLf & _
"「ＪＩＮＺＡＩ！」です。" & vbCrLf & vbCrLf & _
"条件登録確認メールです。" & vbCrLf & _
"以下の内容で登録いたしました。" & vbCrLf

'*****メール履歴*****
'1ページ当りの最大表示件数(/LISPROJE/company/mail_company.asp)
Const Cnt_DispNum = 20

'*** ライセンス情報・継続 ***
'(/LISPROJE/License_Continue.asp、License_ContinueEnd.aspにて使用)
Const License_DispNum					= 12	'1ページ当りの最大表示件数
Const License_Zeiritu					= 5		'税率
Const License_CompanySimebi			= 31	'企業の締め日（締め日未入力の場合使用）
Const License_PersonSimebi			= 20	'求職者の締め日


'***企業検索***
Const Company_List_DispNum			= 10	'1ページあたりの表示最大件数
Const Company_List_ShowPageNum 		= 10	'指定数分ページ番号を表示
Const Company_DspModePop				= 1		'ポップアップ
Const Company_DspModeSam				= 2		'同じ画面
Const Company_JyucyuFlag				= 1		'受注関係の画面より会社選択画面へ移行した場合のフラグ


'*** 求職者登録 ***
'求職者登録において、希望勤務形態に派遣が選択されていた場合使用(/LISPROJE/staff/person_reg5l.aspにて使用)
Const Cnt_Haken_OparateClass_Com = "100"
Const Cnt_Haken_OparateClass_ComMoji = "未面接"
%>
