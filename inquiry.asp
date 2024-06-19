
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
	<script  src="js/jquery-1.12.4.min.js"></script>
	<link rel="icon" href="images/logo-small.svg" type="image/icon type">
	<title>問い合わせ</title>
	<!-- <title>群馬支社</title> -->
</head>

<body ><!-- id="document"  -->
<script type="text/javascript">
	function checkValue(check) {
		var btn = document.getElementById("btn");
		// alert "checkValue";
		if (check.checked) {
			btn.value="個人情報の取り扱いについて 同意して送信する";
			btn.removeAttribute('disabled');
		} else {
			btn.value="個人情報の取り扱いの「同意する」チェックをつけてください。";
			btn.setAttribute('disabled','disabled');
		}//end if
	}//end function

</script>
	<div class="container">
		<!-------------- navbar -------------------->
		<!-- #INCLUDE VIRTUAL="/include/navibar_sjis.html" -->
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
							<!-- <div class="badge-5"><span class="required-6">必須</span></div> -->
						</div> 
						<div class="name-7">
							<input class="name-input-8" name="input2" type="text" aria-label="text" placeholder="例：リス株式会社（法人の方のみ）" />
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
							<input class="name-input-8" name="input5" type="number" aria-label="number" placeholder="例：080********" />
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
						<input type="checkbox" id="vehicle1" name="doui" value="Bike" onclick="checkValue(this)">
						<label for="vehicle1"> <a href="#">個人情報の取り扱い</a></label>
					</div>
					<div class="submit">
						<input name="submit-botton" id="btn" type="submit" disabed="disabled" value="個人情報の取り扱いの「同意する」チェックをつけてください。"><!-- 個人情報の取り扱いについて 同意して送信する  -->
					</div>
				</div>
			</form>
		</div>
		<!---------------- Footer ------------------->
		<!-- #INCLUDE VIRTUAL="/include/footer_sjis.html" -->
		<!--------------- Footer end ----------------->
	</div>
</body>

</html>