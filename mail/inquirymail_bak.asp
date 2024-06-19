<%@ Language=VBScript CodePage=932 %>
<% Option Explicit %>
<%
'******************************************************************************
'概　要：メール送信画面
'備　考：
'更　新：2024/05/17 LIS katayama 作成（旧シテムからの改修）
'　　　：2008/03/13 LIS K.katayama メール送信時, 
'　　　：2008/05/07 LIS K.katayama 問い合わせメール

'　　　：2011/01/05 LIS K.Kokubo Basp.SendMail → SndMail
'******************************************************************************
%>

<!-- #INCLUDE VIRTUAL="/include/commonfunc.asp" -->
<html lang="ja">
<head>
		<meta charset="sjis">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>問い合わせメールテスト</title>
	</head>
<body>


<%

		''susername = Request.Form("username")
		'メール送信処理
		sTo = "mkatayama@lis21.co.jp" ''sReceiverMailAddress	'送信先メールアドレス
		
		sFrom = "lis@lis21.co.jp" ''Cnt_NaviMailAddress	'送信元メールアドレス
		
		'タイトル
		sSubject = "リスHPテストメール" ''MAIL_FROM_STAFF_SUBJECT

		'本文
		'' sBody = "testmail: username= " & VbCrlf & susername


		 
		 'Cnt_MailServer''Const Cnt_MailServer = "172.16.1.39"
		'sResult = SndMail(Cnt_MailServer, sTo, sFrom, sSubject, sBody, "")
		 ''Response.Write "Result= " & sResult & "<br>"
		 ''Response.Write "sBody= " & sBody & "<br>"
		 If sResult=True Then
		 	Response.Write "送信が完了しました。"
		 else
		 	Response.Write "送信が失敗しました。"
		 End if
		

%>
 
		<!--------------- Footer start ----------------->
				<div class="footer-content">
					<div class="vertical">
						<div class="title">リス株式会社</div>
						<div class="footer-item"><a href="./Company.html">会社情報</a></div>
						<div class="footer-item"><a href="Tokyo-branch.html">支社情報</a></div>
						<!--<div class="footer-item"><a href="./Topics.html">トピックス</a></div>-->
						<div class="footer-item"><a href="./Privacy-policy.html">個人情報保護</a></div>
						<div class="footer-item"><a href="./recruit-sale.html">採用情報</a></div>
						<div class="footer-item"><a href="./inquiry.html">お問い合せ</a></div>
					</div>
					<div class="vertical">
						<div class="title">転職サポート</div>
						<div class="footer-item"><a href="./Temporary-staffing.html">人材派遣</a></div>
						<div class="footer-item"><a href="./Recruitment.html">人材紹介</a></div>
						<div class="footer-item"><a href="./Introduction.html">紹介予定派遣</a></div>
						<div class="footer-item"><a href="./q&a.html">Ｑ＆Ａ</a></div>
					</div>
					<div class="vertical">
						<div class="title">人材サービス</div>
						<div class="footer-item"><a href="./Dispatch.html">派遣</a></div>
						<div class="footer-item"><a href="./Prelusion.html">紹介</a></div>
						<div class="footer-item"><a href="./Schedule.html">紹介予定派遣</a></div>
					</div>
				</div>

				<div class="copyright">Copyright(c) 2024 LIS co.,Ltd. All rights Reserved.</div>
			</div>
			<div class="footer-map">
				<iframe title="myFrame"
					src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d6481.103588769505!2d139.69093757604216!3d35.68803667258506!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x60188cd3741a9df7%3A0x4fb5f8fb9f0a0195!2sShinjuku%20NS%20Building!5e0!3m2!1sen!2sjp!4v1715059757728!5m2!1sen!2sjp"
					width="600" height="450" style="border:0;" loading="lazy"
					referrerpolicy="no-referrer-when-downgrade"></iframe>

			</div>

		</footer>
		<!--------------- Footer end ----------------->
		<!--------------- footer-mobile ----------------->
		<div class="footer-mobile">
			<h2>リス株式会社</h2>
			<div class="copyright">Copyright(c) 2024 LIS co.,Ltd. All rights Reserved.</div>
		</div>
		<!--------------- Footer end ----------------->
	</div>  
	</body>
</html>

