<%@ Language=VBScript CodePage=932 %>
<% Option Explicit %>
<%
'******************************************************************************
'概　要：メール送信画面
'備　考：
'更　新：2024/05/17 LIS katayama 作成（旧シテムからの改修）
'　　　：2008/03/13 LIS K.Kokubo メール送信時, 
'　　　：2008/05/07 LIS K.Kokubo 求人票の閲覧可否をChkOrderDspで判定するように変更

'　　　：2011/01/05 LIS K.Kokubo Basp.SendMail → SndMail
'******************************************************************************
%>

<!-- #INCLUDE VIRTUAL="/include/commonfunc.asp" -->
<!-- #INCLUDE VIRTUAL="/config/personnel.asp" -->
<html lang="ja">
	<head>
		<meta charset="sjis">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta charset="SJIS">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<link href="https://fonts.googleapis.com/css2?family=Noto+Sans:ital,wght@0,100..900;1,100..900&display=swap"
		rel="stylesheet">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
	<link rel="stylesheet" href="/css/inquiry.css">
	<script defer src="/js/app.js"></script>
	<script defer src="/js/mobile.js"></script>
	<link rel="icon" href="images/logo-small.svg" type="/image/icon type">
	<title>採用応募・問い合わせメール</title>
    
	</head>
<body>

	<div class="container">
		<!-------------- navbar -------------------->
		<div class="navbar">
			<div class="top_nav">
				<div class="logo">
					<a href="/index.html">
						<img src="/images/Web/Logo/Web/Desctop.svg" alt="">
					</a>
				</div>
				<div class="left_link">
					<a href="./Job seekers.html">お仕事をお探しの求職者様</a>
					<div class="line_1"></div>
					<a href="./human resurs.html">人材をお探しの採用担当者様</a>
				</div>
			</div>
			<!-------------- mobile top_nav -------------------->
			<div class="mobile_container">
				<div class="topnav">
					<a href="/index.html" class="active">
						<img src="/images/Mobile/Web/Logo/Phone.svg" alt="">
					</a>
					<!-- Navigation links (hidden by default) -->
					<div id="myLinks">
						<a href="/Company.html">会社情報</a>
						<a href="/Tokyo-branch.html">支社情報</a>
						<a href="#">採用情報</a>
						<a href="#">転職サポート</a>
						<a href="/recruit-sale.html">人材サービス</a>
						<a href="/inquiry.html">お問い合わせ</a>
					</div>
					<!-- "Hamburger menu" / "Bar icon" to toggle the navigation links -->
					<a href="javascript:void(0);" class="icon" onclick="myFunction()">
						<img src="/images/Mobile/Web/Menu.svg" alt="">
					</a>
				</div>
			</div>
			<!-------------- mobile top_nav end ---------------->

			<!-- navigation -->
			<div class="navigation">
				<div class="nav_left">
					<a href="/Company.html">会社情報</a>
					<a href="/Tokyo-branch.html">支社情報</a>
					<!--<a href="/Topics.html">トピックス</a>-->
					<a href="/inquiry.html">お問い合わせ</a>
					<a href="/recruit-sale.html" class="home">
						<div class="home_text">採用情報</div>
					</a>
				</div>
				<div class="nav_right">
					<a href="./Privacy-policy.html">個人情報保護方針・取り扱い</a>
					<img src="images/Web/P-mark.svg" alt="">
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
				</ul>
			</nav>
			<!----- submenu end ----->
		</div>
		<!--- navbar END -------------->


<%
'' Const Cnt_MailServer = "172.16.1.39" ''personnelに書いた
Dim sTo
Dim sFrom
Dim sBody
Dim sSubject
dim sResult
Dim susername
Dim input1
Dim input2
Dim input3
Dim input4
Dim input5
Dim input6
Dim doui

	input1 = Request.Form("input1")
	input2 = Request.Form("input2")'
	input3 = Request.Form("input3")	
	input4 = Request.Form("input4")
	input5 = Request.Form("input5")	
	input6 = Request.Form("input6")
	''input6 = Replace(input6, vbCrLf, "<BR>")
	input6 = Server.HTMLEncode(input6)
	doui = Request.Form("vehicle1")

	input1 = Trim(input1)
	input2 = Trim(input2)
	input3 = Trim(input3)
	input4 = Trim(input4)
	input5 = Trim(input5)
	input6 = Trim(input6)
	doui = Trim(doui)
	
'	Response.Write "username= " & input1 & "<br>"
'	Response.Write "company= " & input2 & "<br>"
'	Response.Write "pref= " & input3 & "<br>"'
'	Response.Write "mail= " & input4 & "<br>"
''	Response.Write "応募職種= " & input5 & "<br>"
'	Response.Write "doukitext= " & input6 & "<br>"
'	Response.Write "doui= " & doui & "<br>"

		'メール送信処理
		sTo = "lis@lis21.co.jp" ''sReceiverMailAddress	'送信先メールアドレス
		
		sFrom = "lis@lis21.co.jp" ''Cnt_NaviMailAddress	'送信元メールアドレス
		
		'タイトル
		sSubject = "リスHP採用応募メール" ''MAIL_FROM_STAFF_SUBJECT

		'本文
		sBody = " 氏名= " & input1 & vbCrLf  '
		sBody = sBody &  " ふりがな= " & input2 & vbCrLf 
		sBody = sBody &  " メール= " & input3 & VbCrLf
		sBody = sBody &  " 電話= " & input4 & vbCrLf 
		sBody = sBody &  " 応募職種= " & input5 & vbCrLf 
		sBody = sBody &  " 志望動機= " & input6 & vbCrLf 
		sBody = sBody & " 個人情報の同意 = " & doui  & vbCrLf
		 		
		 sResult = SndMail(Cnt_MailServer, sTo, sFrom, sSubject, sBody, "")''メール送信処理
		' Response.Write "Result= " & sResult & "<br>"
		  ' Response.Write "body= " & sBody & "<br>"
		 
		 
		 Response.Write "<div class='main-container-3'>"
		 If sResult=True Then
		 
		 Response.Write "	<div class='submit'>"
		 	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;送信が完了しました。<br>"
		 	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;担当者から連絡を差し上げます。<br>"
		 	Response.Write "	</div>"
		 else
		 	Response.Write "	<div class='submit'>"
		 		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;送信は失敗しました。<br>"
		 	Response.Write "	</div>"
				Response.Write "	<div class='submit'>"
				Response.Write "	    <input type='button' value='戻って修正する' onclick='history.back();'  id='btnBack'>"
				Response.Write "	</div>"
		 	
		 End if
			Response.Write " </div>"
%>
 
 
 		<!---------------- Footer ------------------->
		<footer>

			<div class="footer-left">
				<div class="footer-header">
					<div class="footer-logo">
						<img src="/images/Logo/Web/Desctop1.svg" alt="company">
					</div>
					<!-- <div class="social">
						<a href="#"><img src="./images/x.svg" alt="x">
						</a>
						<a href="#">
							<img src="./images/facebook.svg" alt="facebook">
						</a>
						<a href="#">
							<img src="./images/instagram.svg" alt="instagram">
						</a>
					</div> -->
				</div>

				<div class="footer-content">
					<div class="vertical">
						<div class="title">リス株式会社</div>
						<div class="footer-item"><a href="/Company.html">会社情報</a></div>
						<div class="footer-item"><a href="Tokyo-branch.html">支社情報</a></div>
						<!--<div class="footer-item"><a href="/Topics.html">トピックス</a></div>-->
						<div class="footer-item"><a href="/Privacy-policy.html">個人情報保護</a></div>
						<div class="footer-item"><a href="/recruit-sale.html">採用情報</a></div>
						<div class="footer-item"><a href="/inquiry.html">お問い合せ</a></div>
					</div>
					<div class="vertical">
						<div class="title">転職サポート</div>
						<div class="footer-item"><a href="/Temporary-staffing.html">人材派遣</a></div>
						<div class="footer-item"><a href="/Recruitment.html">人材紹介</a></div>
						<div class="footer-item"><a href="/Introduction.html">紹介予定派遣</a></div>
						<div class="footer-item"><a href="/q&a.html">Ｑ＆Ａ</a></div>
					</div>
					<div class="vertical">
						<div class="title">人材サービス</div>
						<div class="footer-item"><a href="/Dispatch.html">派遣</a></div>
						<div class="footer-item"><a href="/Prelusion.html">紹介</a></div>
						<div class="footer-item"><a href="/Schedule.html">紹介予定派遣</a></div>
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

