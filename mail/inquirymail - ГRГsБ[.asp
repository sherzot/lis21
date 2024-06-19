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
<html lang="jpn">
	<head>
		<meta charset="sjis">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>問い合わせメールテスト</title>
	</head>
<body>
		<!-------------- navbar -------------------->
		<div class="navbar">
			<div class="top_nav">
				<div class="logo">
					<a href="./index.html">
						<img src="images/logo.svg" alt="logo">
					</a>
				</div>
				<div class="left_link">
					<a href="/Job seekers.html">お仕事をお探しの求職者様</a>
					<div class="line_1"></div>
					<a href="/human resurs.html">人材をお探しの採用担当者様</a>
				</div>
			</div>
		
			<!-- navigation start -->
			<div class="navigation">
				<div class="nav_left">
					<a href="/index.html" class="home">
						<img src="/images/home.svg" alt="home">
						<div class="home_text">ホーム</div>
					</a>
		
					<a href="/company.html">会社情報</a>
					<a href="/Tokyo-branch.html">支社情報</a>
					<a href="/Topics.html">トピックス</a>
					<a href="/inquiry.asp">お問い合わせ</a>
				</div>
				<div class="nav_right">
					<a href="./desktop-20.html">個人情報保護方針・取り扱い</a>
					<img src="images/P-mark.svg" alt="">
				</div>
			</div>
			<!-- navigation END -->
		
			<!----- submenu  ----->
			<nav class="submenu">
				<ul class="submenu-links">
				  <li class="submenu-dropdown">
					<a href="#">転職サポート</a>
					<div class="dropdown">    
						<a href="/Temporary-staffing.html">人材派遣</a>
						<a href="/Labor-regulations.html">就業規則</a>
						<a href="/Qualification.html">資格支援取得制度</a>
						<a href="/Recruitment.html">人材紹介</a>
						<a href="/Introduction.html">紹介予定派遣</a>
						<a href="/q&a.html">Ｑ＆Ａ</a>
					</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">人材サービス</a>
					<div class="dropdown">
						<a href="/Dispatch.html">派遣</a>
						<a href="/Prelusion.html">紹介</a>
						<a href="/Schedule.html">紹介予定派遣</a>
                	</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">求人広告の掲載</a>
					<div class="dropdown">
						<a href="/Service-contents.html">サービス内容</a>
						<a href="/Job-applicant.html">求職者情報</a>
						<a href="/Prices.html">プラン・料金</a>
						<a href="/Questions.html">Ｑ＆Ａ</a>
					</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">採用情報</a>
					<div class="dropdown">
					  <a href="/recruit.html">新卒採用情報</a>
					  <a href="./recruit.html">中途採用情報</a>
					</div>
				  </li>
				</ul>
			</nav>
			<!----- submenu end ----->  
			<!--- navbar END -------------->
		</div>
<%
Const Cnt_MailServer = "172.16.1.39"
Dim sTo
Dim sFrom
Dim sBody
Dim sSubject
dim sResult
Dim input1
Dim input2
Dim input3
Dim input4
Dim input5
Dim input6
Dim doui
input1 = Request.Form("input1")
input2 = Request.Form("input2")
input3 = Request.Form("input3")
input4 = Request.Form("input4")
input5 = Request.Form("input5")
input6 = Request.Form("input6")
doui = Request.Form("doui")

Response.Write "username= " & input1 & "<br>"
Response.Write "company= " & input2 & "<br>"
Response.Write "pref= " & input3 & "<br>"
esponse.Write "mail= " & input4 & "<br>"
'Response.Write "mobileno= " & input5 & "<br>"
Response.Write "doukitext= " & input6 & "<br>"


		''susername = Request.Form("username")
		'メール送信処理
		sTo = "mkatayama@lis21.co.jp" ''sReceiverMailAddress	'送信先メールアドレス
		
		sFrom = "lis@lis21.co.jp" ''Cnt_NaviMailAddress	'送信元メールアドレス
		
		'タイトル
		sSubject = "リスHPテストメール" ''MAIL_FROM_STAFF_SUBJECT

		'本文
		'' sBody = "testmail: username= " & VbCrlf & susername
		sBody = " 氏名= " & input-1 & vbCrLf  ''& "会社名= " &  company & VBCrlf
		sBody = sBody &  " 会社名= " & input-2 & vbCrLf 
		sBody = sBody &  " 都道府県= " & input-3 & VbCrLf
		sBody = sBody &  " メール= " & input-4 & vbCrLf 
'		sBody = sBody &  "電話= "　& input-5 & vbCrLf 
		sBody = sBody &  "  動機== " & input-6 & vbCrLf 
		sBody = sBody & " 個人情報の同意 =" & doui  & vbCrLf
		 

		 
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

