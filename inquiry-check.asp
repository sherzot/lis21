
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
	<title>問い合わせ確認</title>

</head>

<body>


	<div class="container">
		<!-------------- navbar -------------------->
		<!-- #INCLUDE VIRTUAL="/include/navibar_sjis.html" -->
		<!--- navbar END -------------->

		<!------------ step -------------->
		<div class="step">
			<div class="item">ホーム<img src="images/chevron_right.svg" alt=""></div>
			<div class="item-active">お問い合わせ確認</div>
		</div>
		<!---------- step END ----------------->
		<!-- HEADER -->
		<div class="header">
			<!-- form-section -->
			<form class="form" action="/mail/inquirymail.asp" name="inquiry" method="POST"><!-- inquirymail.asp -->
				<div class="main-container">
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">お名前</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input1" name="input1" type="text" aria-label="text" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">御社名</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input2" name="input2" type="text" aria-label="text" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">都道府県</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input3" name="input3" type="text" aria-label="text" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">メールアドレス</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input4" name="input4" type="email" aria-label="email" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">電話番号（携帯電話）</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input5" name="input5" type="number" aria-label="number"/>
						</div>
					</div>


					<div class="main-container-3">
						<div class="left">
							<span class="motivation">内容</span>
						</div>
						<!-- <textarea class="right" name="input6"  id="input6" cols="30" rows="10"  value="<% Response.Write input6 %>" ></textarea> -->
						<input class="name-input-8" id="input6" name="input6" type="text"  aria-label="text"  />
					</div>


					<div class="chek-box">
						<input type="checkbox" id="vehicle1" name="doui" value="同意" checked="on">
						<label for="vehicle1"> <a href="#">個人情報の取り扱い</a></label>
					</div>
					
					<div class="submit">
					    <input type="button" value="戻って修正する" onclick="history.back();"  id="btnBack">
					</div>
					
					<div class="submit">
						<input name="submit-botton" id="btn" type="submit" disabled="disabled" value="メールを送信する">
					</div>
				</div>
			</form>
		</div>
		<!---------------- Footer ------------------->
				<!-- #INCLUDE VIRTUAL="/include/footer_sjis.html" -->

		<!--------------- Footer end ----------------->
	</div>
</body>
<script type="text/javascript">

 	const btn = document.getElementById("btn");
 
    btn.addEventListener('mouseover', function() {

		var input1 = document.getElementById("input1");
		var input2 = document.getElementById("input2");
		var input3 = document.getElementById("input3");
		var input4 = document.getElementById("input4");
		var input5 = document.getElementById("input5");
		var input6 = document.getElementById("input6");
		
 	
		
		if (!(input1.value)) {
			 //btn.value="お名前が入力されていません。戻って修正してください。";
			 alert('お名前が入力されていません');
			btn.setAttribute('disabled','disabled');
			
		//} else if (!(input2.value)) {
		//	//btn.value="御社名が入力されていません。戻って修正してください。";
		//	 alert('御社名が入力されていません');
		//	btn.setAttribute('disabled','disabled');
		} else if (!(input3.value)) {

			 alert('都道府県が入力されていません');		
			btn.setAttribute('disabled','disabled');
		} else if (!input4.value) {
			//btn.value="メールアドレスが入力されていません。戻って修正してください。";
			 alert('メールアドレスが入力されていません');
			btn.setAttribute('disabled','disabled');		
			
		} else if (!input5.value) {
			//btn.value="電話番号が入力されていません。戻って修正してください。";
			 alert('電話番号が入力されていません');
			btn.setAttribute('disabled','disabled');		
			
		} else if (!(input6.value)) {
			//btn.value="内容が入力されていません。戻って修正してください。";
			 alert('内容が入力されていません。戻って修正してください。');
			btn.setAttribute('disabled','disabled');

		} else {
			btn.value="メールを送信する。";
			btn.removeAttribute('disabled');
		}//end if
	});//end function
</script>


</html>
