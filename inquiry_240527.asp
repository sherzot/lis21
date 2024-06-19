<%@ Language=VBScript CodePage=932 %>
<!DOCTYPE html>
<html lang="en">

<head>
	<meta charset="SJIS">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<link href="https://fonts.googleapis.com/css2?family=Noto+Sans:ital,wght@0,100..900;1,100..900&display=swap"
		rel="stylesheet">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
	<link rel="stylesheet" href="css/inquiry.css">
	<script defer src="js/app.js"></script>
	<script defer src="js/mobile.js"></script>
	<link rel="icon" href="images/logo-small.svg" type="image/icon type">
	<title>会社情報</title>
	<title>群馬支社</title>
</head>

<body>
	<div class="container">
		<!-------------- navbar -------------------->
		<div class="navbar">
			<div class="top_nav">
				<div class="logo">
					<a href="./index.html">
						<img src="images/Web/Logo/Web/Desctop.svg" alt="">
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
					<a href="./index.html" class="active">
						<img src="images/Mobile/Web/Logo/Phone.svg" alt="">
					</a>
					<!-- Navigation links (hidden by default) -->
					<div id="myLinks">
						<a href="./Company.html">会社情報</a>
						<a href="./Tokyo-branch.html">支社情報</a>
						<a href="#">採用情報</a>
						<a href="#">転職サポート</a>
						<a href="./recruit-sale.html">人材サービス</a>
						<a href="./inquiry.html">お問い合わせ</a>
					</div>
					<!-- "Hamburger menu" / "Bar icon" to toggle the navigation links -->
					<a href="javascript:void(0);" class="icon" onclick="myFunction()">
						<img src="./images/Mobile/Web/Menu.svg" alt="">
					</a>
				</div>
			</div>
			<!-------------- mobile top_nav end ---------------->

			<!-- navigation -->
			<div class="navigation">
				<div class="nav_left">
					<a href="./Company.html">会社情報</a>
					<a href="./Tokyo-branch.html">支社情報</a>
					<!--<a href="./Topics.html">トピックス</a>-->
					<a href="./inquiry.html">お問い合わせ</a>
					<a href="./recruit-sale.html" class="home">
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
							<a href="./Temporary-staffing.html">人材派遣</a>
							<a href="./Recruitment.html">人材紹介</a>
							<a href="./Introduction.html">紹介予定派遣</a>
							<a href="./q&a.html">Ｑ＆Ａ</a>
						</div>
					</li>
					<li class="submenu-dropdown">
						<a href="#">人材サービス</a>
						<div class="dropdown">
							<a href="./Dispatch.html">派遣</a>
							<a href="./Prelusion.html">紹介</a>
							<a href="./Schedule.html">紹介予定派遣</a>
						</div>
					</li>
				</ul>
			</nav>
			<!----- submenu end ----->
		</div>
		<!--- navbar END -------------->

		<!------------ step -------------->
		<div class="step">
			<div class="item">ホーム<img src="images/chevron_right.svg" alt=""></div>
			<div class="item-active">お問い合わせ</div>
		</div>
		<!---------- step END ----------------->
		<!-- HEADER -->
		<div class="header">
			<!-- form-section -->
			<form class="form" action="./inquiry-check.asp" name="inquiry" method="POST">
				<div class="main-container">
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">お名前</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input1" type="text" aria-label="text" placeholder="例：山田 太郎" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">御社名</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input2" type="text" aria-label="text" placeholder="例：リス株式会社" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">都道府県</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input3" type="text" aria-label="text" placeholder="例：埼玉県" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">メールアドレス</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input4" type="email" aria-label="email"
								placeholder="例：lis@lis21.co.jp" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">電話番号（携帯電話）</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input5" type="number" aria-label="number" placeholder="例：07069505055" />
						</div>
					</div>


					<div class="main-container-3">
						<div class="left">
							<span class="motivation">内容</span><button class="badge"><span
									class="required">必須</span></button>
						</div>
						<textarea class="right" name="input6" id="" cols="30" rows="10"
							placeholder=""></textarea>
					</div>
					<div class="chek-box">
						<input type="checkbox" id="vehicle1" name="doui" value="Bike">
						<label for="vehicle1"> <a href="#">個人情報の取り扱い</a></label>
					</div>
					<div class="submit">
						<input name="submit-botton" type="submit" value="個人情報の取り扱いについて 同意して送信する">
					</div>
				</div>
			</form>
		</div>
		<!---------------- Footer ------------------->
		<footer>

			<div class="footer-left">
				<div class="footer-header">
					<div class="footer-logo">
						<img src="./images/Logo/Web/Desctop1.svg" alt="company">
					</div>
					<div class="social">
						<a href="#"><img src="./images/x.svg" alt="x">
						</a>
						<a href="#">
							<img src="./images/facebook.svg" alt="facebook">
						</a>
						<a href="#">
							<img src="./images/instagram.svg" alt="instagram">
						</a>
					</div>
				</div>

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