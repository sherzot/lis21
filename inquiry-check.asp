
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
	<title>�₢���킹�m�F</title>

</head>

<body>


	<div class="container">
		<!-------------- navbar -------------------->
		<!-- #INCLUDE VIRTUAL="/include/navibar_sjis.html" -->
		<!--- navbar END -------------->

		<!------------ step -------------->
		<div class="step">
			<div class="item">�z�[��<img src="images/chevron_right.svg" alt=""></div>
			<div class="item-active">���₢���킹�m�F</div>
		</div>
		<!---------- step END ----------------->
		<!-- HEADER -->
		<div class="header">
			<!-- form-section -->
			<form class="form" action="/mail/inquirymail.asp" name="inquiry" method="POST"><!-- inquirymail.asp -->
				<div class="main-container">
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">�����O</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input1" name="input1" type="text" aria-label="text" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">��Ж�</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input2" name="input2" type="text" aria-label="text" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">�s���{��</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input3" name="input3" type="text" aria-label="text" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">���[���A�h���X</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input4" name="input4" type="email" aria-label="email" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">�d�b�ԍ��i�g�ѓd�b�j</span>
						</div>
						<div class="name-7">
							<input class="name-input-8" id="input5" name="input5" type="number" aria-label="number"/>
						</div>
					</div>


					<div class="main-container-3">
						<div class="left">
							<span class="motivation">���e</span>
						</div>
						<!-- <textarea class="right" name="input6"  id="input6" cols="30" rows="10"  value="<% Response.Write input6 %>" ></textarea> -->
						<input class="name-input-8" id="input6" name="input6" type="text"  aria-label="text"  />
					</div>


					<div class="chek-box">
						<input type="checkbox" id="vehicle1" name="doui" value="����" checked="on">
						<label for="vehicle1"> <a href="#">�l���̎�舵��</a></label>
					</div>
					
					<div class="submit">
					    <input type="button" value="�߂��ďC������" onclick="history.back();"  id="btnBack">
					</div>
					
					<div class="submit">
						<input name="submit-botton" id="btn" type="submit" disabled="disabled" value="���[���𑗐M����">
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
			 //btn.value="�����O�����͂���Ă��܂���B�߂��ďC�����Ă��������B";
			 alert('�����O�����͂���Ă��܂���');
			btn.setAttribute('disabled','disabled');
			
		//} else if (!(input2.value)) {
		//	//btn.value="��Ж������͂���Ă��܂���B�߂��ďC�����Ă��������B";
		//	 alert('��Ж������͂���Ă��܂���');
		//	btn.setAttribute('disabled','disabled');
		} else if (!(input3.value)) {

			 alert('�s���{�������͂���Ă��܂���');		
			btn.setAttribute('disabled','disabled');
		} else if (!input4.value) {
			//btn.value="���[���A�h���X�����͂���Ă��܂���B�߂��ďC�����Ă��������B";
			 alert('���[���A�h���X�����͂���Ă��܂���');
			btn.setAttribute('disabled','disabled');		
			
		} else if (!input5.value) {
			//btn.value="�d�b�ԍ������͂���Ă��܂���B�߂��ďC�����Ă��������B";
			 alert('�d�b�ԍ������͂���Ă��܂���');
			btn.setAttribute('disabled','disabled');		
			
		} else if (!(input6.value)) {
			//btn.value="���e�����͂���Ă��܂���B�߂��ďC�����Ă��������B";
			 alert('���e�����͂���Ă��܂���B�߂��ďC�����Ă��������B');
			btn.setAttribute('disabled','disabled');

		} else {
			btn.value="���[���𑗐M����B";
			btn.removeAttribute('disabled');
		}//end if
	});//end function
</script>


</html>
