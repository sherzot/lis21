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
	<title>��Џ��</title>
	<title>�Q�n�x��</title>
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
					<a href="./Job seekers.html">���d�������T���̋��E�җl</a>
					<div class="line_1"></div>
					<a href="./human resurs.html">�l�ނ����T���̗̍p�S���җl</a>
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
						<a href="./Company.html">��Џ��</a>
						<a href="./Tokyo-branch.html">�x�Џ��</a>
						<a href="#">�̗p���</a>
						<a href="#">�]�E�T�|�[�g</a>
						<a href="./recruit-sale.html">�l�ރT�[�r�X</a>
						<a href="./inquiry.html">���₢���킹</a>
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
					<a href="./Company.html">��Џ��</a>
					<a href="./Tokyo-branch.html">�x�Џ��</a>
					<!--<a href="./Topics.html">�g�s�b�N�X</a>-->
					<a href="./inquiry.html">���₢���킹</a>
					<a href="./recruit-sale.html" class="home">
						<div class="home_text">�̗p���</div>
					</a>
				</div>
				<div class="nav_right">
					<a href="./Privacy-policy.html">�l���ی���j�E��舵��</a>
					<img src="images/Web/P-mark.svg" alt="">
				</div>
			</div>
			<!-- navigation END -->

			<!----- submenu  ----->
			<nav class="submenu">
				<ul class="submenu-links">
					<li class="submenu-dropdown">
						<a href="#">�]�E�T�|�[�g</a>
						<div class="dropdown">
							<a href="./Temporary-staffing.html">�l�ޔh��</a>
							<a href="./Recruitment.html">�l�ޏЉ�</a>
							<a href="./Introduction.html">�Љ�\��h��</a>
							<a href="./q&a.html">�p���`</a>
						</div>
					</li>
					<li class="submenu-dropdown">
						<a href="#">�l�ރT�[�r�X</a>
						<div class="dropdown">
							<a href="./Dispatch.html">�h��</a>
							<a href="./Prelusion.html">�Љ�</a>
							<a href="./Schedule.html">�Љ�\��h��</a>
						</div>
					</li>
				</ul>
			</nav>
			<!----- submenu end ----->
		</div>
		<!--- navbar END -------------->

		<!------------ step -------------->
		<div class="step">
			<div class="item">�z�[��<img src="images/chevron_right.svg" alt=""></div>
			<div class="item-active">���₢���킹</div>
		</div>
		<!---------- step END ----------------->
		<!-- HEADER -->
		<div class="header">
			<!-- form-section -->
			<form class="form" action="./inquiry-check.asp" name="inquiry" method="POST">
				<div class="main-container">
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">�����O</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input1" type="text" aria-label="text" placeholder="��F�R�c ���Y" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">��Ж�</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input2" type="text" aria-label="text" placeholder="��F���X�������" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">�s���{��</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input3" type="text" aria-label="text" placeholder="��F��ʌ�" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">���[���A�h���X</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input4" type="email" aria-label="email"
								placeholder="��Flis@lis21.co.jp" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">�d�b�ԍ��i�g�ѓd�b�j</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input class="name-input-8" name="input5" type="number" aria-label="number" placeholder="��F07069505055" />
						</div>
					</div>


					<div class="main-container-3">
						<div class="left">
							<span class="motivation">���e</span><button class="badge"><span
									class="required">�K�{</span></button>
						</div>
						<textarea class="right" name="input6" id="" cols="30" rows="10"
							placeholder=""></textarea>
					</div>
					<div class="chek-box">
						<input type="checkbox" id="vehicle1" name="doui" value="Bike">
						<label for="vehicle1"> <a href="#">�l���̎�舵��</a></label>
					</div>
					<div class="submit">
						<input name="submit-botton" type="submit" value="�l���̎�舵���ɂ��� ���ӂ��đ��M����">
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
						<div class="title">���X�������</div>
						<div class="footer-item"><a href="./Company.html">��Џ��</a></div>
						<div class="footer-item"><a href="Tokyo-branch.html">�x�Џ��</a></div>
						<!--<div class="footer-item"><a href="./Topics.html">�g�s�b�N�X</a></div>-->
						<div class="footer-item"><a href="./Privacy-policy.html">�l���ی�</a></div>
						<div class="footer-item"><a href="./recruit-sale.html">�̗p���</a></div>
						<div class="footer-item"><a href="./inquiry.html">���₢����</a></div>
					</div>
					<div class="vertical">
						<div class="title">�]�E�T�|�[�g</div>
						<div class="footer-item"><a href="./Temporary-staffing.html">�l�ޔh��</a></div>
						<div class="footer-item"><a href="./Recruitment.html">�l�ޏЉ�</a></div>
						<div class="footer-item"><a href="./Introduction.html">�Љ�\��h��</a></div>
						<div class="footer-item"><a href="./q&a.html">�p���`</a></div>
					</div>
					<div class="vertical">
						<div class="title">�l�ރT�[�r�X</div>
						<div class="footer-item"><a href="./Dispatch.html">�h��</a></div>
						<div class="footer-item"><a href="./Prelusion.html">�Љ�</a></div>
						<div class="footer-item"><a href="./Schedule.html">�Љ�\��h��</a></div>
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
			<h2>���X�������</h2>
			<div class="copyright">Copyright(c) 2024 LIS co.,Ltd. All rights Reserved.</div>
		</div>
		<!--------------- Footer end ----------------->
	</div>
</body>

</html>