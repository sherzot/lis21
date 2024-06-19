
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="SJIS">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans:ital,wght@0,100..900;1,100..900&display=swap"
        rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
    <link rel="stylesheet" href="css/recruit-sale.css">
    <link rel="stylesheet" href="css/inquiry.css">
    <script src="js/app.js"></script>
    <script defer src="js/mobile.js"></script>
    <link rel="icon" href="images/logo-small.svg" type="image/icon type">
    <title>�̗p���</title>
</head>

<body>
    <div class="container">
        <!-------------- navbar -------------------->
        <!-- #INCLUDE VIRTUAL="/include/navibar_sjis.html" -->

        <!------------ step -------------->
        <div class="step">
            <div class="item">�z�[��<img src="images/Web/chevron_right.svg" alt=""></div>
            <div class="item-active">�̗p���</div>
        </div>

        <!---------- step END ----------------->

        <!-- HEADER -->

        <div class="header">
            <!-- list-menu-left -->
            <div class="list-menu">
                <div class="list-menu-title">��W�E��</div>
                <ul>
                    <li><a href="./recruit-sale.html">�Љ�R���T���i�S���V���_�j</a></li>
                    <li><a href="./recruit-system.html">�o�b�N�G���h�G���W�j�A�i�{�Ёj</a></li>
                    <li><a href="./recruit-designer.html">�t�����g�G���h�G���W�j�A�i�{�Ёj</a></li>
                </ul>
            </div>

                <form class="form" action="/mail/recruit-salemail.asp" method="POST" >
                    <div class="main-container">
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">����</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8"id="input1" name="input1" type="text" aria-label="text"  />
                                <!-- <span class="example-name">��F�R�c ���Y</span> -->
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">�ӂ肪��</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" id="input2" name="input2" type="text" aria-label="text"   />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">���[���A�h���X</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" type="email" id="input3" name="input3" aria-label="email"/>
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">�d�b�ԍ��i�g�ѓd�b�j</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" type="number" id="input4" name="input4" aria-label="number"/>
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">�L�����A</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>

                           <div class="name-7">
                                <input class="name-input-8" id="input5" name="input5" type="text" aria-label="text"/>
                            </div>

                             < -- <div class="radio"> -->
                                <!-- <input type="radio" id="input5" name="input5" checked="checked" value="<% Response.Write input5 %>" > --><!-- name="fav_language"  -->
                                  <!-- <label for="html"><% Response.Write input5 %></label><br> --><!-- �{�Љc�ƕ��i�Љ�R���T���j  -->
                             </div> -
                        </div>

                        <div class="main-container-3">
                            <div class="left">
                         <div class="main-container-3">
                            <div class="left">
                                <span class="motivation">�X�L��
                                
                                </span><button class="badge"><span
                                        class="required">�K�{</span></button>
                            </div>
                           <div class="name-7">
                                <input class="name-input-8" id="input6" name="input6" type="text" aria-label="text"/>
                            </div>
                                
                            <!-- <textarea class="right"  id="input6" name="input6"   cols="30" rows="10"
                                 value="<% Response.Write input6 %>" "><% Response.Write input6 %></textarea> --><!-- name="�����哮�@�������͂�������"  -->
                        	<input type="hidden" name="input7"   value="�{�Љc�ƕ��i�Ǘ��E�E��E�o���ҁj">
                        </div>
                        <div class="chek-box">
                            <input type="checkbox" id="vehicle1" name="vehicle1"  value="����"  checked="on" ><!-- <% Response.Write check %> -->
                            <!--  value="Bike"  -->
                            <label for="vehicle1"> <a href="#">�l���̎�舵��</a></label>
                        </div>
                        
						<div class="submit">
						  <input type="button" value="�߂��ďC������" onclick="history.back();"  id="btnBack">
						</div>                        
                        <div class="submit">
                            <input type="submit" id="btn" value="���[���𑗐M����">
                        </div>
                    </div>
                </form>
            </div>
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
			
		} else if (!(input2.value)) {
		//	//btn.value="��Ж������͂���Ă��܂���B�߂��ďC�����Ă��������B";
			 alert('�ӂ肪�Ȃ����͂���Ă��܂���');
			btn.setAttribute('disabled','disabled');
		} else if (!(input3.value)) {

			 alert('���[���A�h���X�����͂���Ă��܂���');		
			btn.setAttribute('disabled','disabled');
		} else if (!input4.value) {
			//btn.value="�d�b�ԍ������͂���Ă��܂���B�߂��ďC�����Ă��������B";
			 alert('�d�b�ԍ������͂���Ă��܂���');
			btn.setAttribute('disabled','disabled');		
			
		} else if (!input5.value) {
			//btn.value="����E�킪���͂���Ă��܂���B�߂��ďC�����Ă��������B";
			 alert('�L�����A�����͂���Ă��܂���');
			btn.setAttribute('disabled','disabled');		
			
		} else if (!(input6.value)) {
			
			 alert('�X�L�������͂���Ă��܂���B�߂��ďC�����Ă��������B');
			btn.setAttribute('disabled','disabled');

		} else {
			btn.value="���[���𑗐M����B";
			btn.removeAttribute('disabled');
		}//end if
	});//end function
</script>

</html>