
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
	<title>�₢���킹</title>
	<!-- <title>�Q�n�x��</title> -->
</head>

<body ><!-- id="document"  -->
<script type="text/javascript">
	function checkValue(check) {
		var btn = document.getElementById("btn");
		// alert "checkValue";
		if (check.checked) {
			btn.value="�l���̎�舵���ɂ��� ���ӂ��đ��M����";
			btn.removeAttribute('disabled');
		} else {
			btn.value="�l���̎�舵���́u���ӂ���v�`�F�b�N�����Ă��������B";
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
							<!-- <div class="badge-5"><span class="required-6">�K�{</span></div> -->
						</div> 
						<div class="name-7">
							<input class="name-input-8" name="input2" type="text" aria-label="text" placeholder="��F���X������Ёi�@�l�̕��̂݁j" />
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
							<input class="name-input-8" name="input5" type="number" aria-label="number" placeholder="��F080********" />
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
						<input type="checkbox" id="vehicle1" name="doui" value="Bike" onclick="checkValue(this)">
						<label for="vehicle1"> <a href="#">�l���̎�舵��</a></label>
					</div>
					<div class="submit">
						<input name="submit-botton" id="btn" type="submit" disabed="disabled" value="�l���̎�舵���́u���ӂ���v�`�F�b�N�����Ă��������B"><!-- �l���̎�舵���ɂ��� ���ӂ��đ��M����  -->
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