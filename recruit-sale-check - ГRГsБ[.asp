<%@ Language=VBScript CodePage=932 %>
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
                        <a href="./recruit-sale.html">�̗p���</a>
                        <a href="./Privacy-policy.html">�l���ی���j</a>
                        <a href="./Temporary-staffing.html">�l�ޔh��</a>
                        <a href="./Recruitment.html">�l�ޏЉ�</a>
                        <a href="./Introduction.html">�Љ�\��h��</a>
                        <a href="./Trainer.html">�g���[�i�[�̏Љ�</a>
                        <a href="./q&a.html">�p���`</a>
                        <a href="./Dispatch.html">�h��</a>
                        <a href="./Prelusion.html">�Љ�</a>
                        <a href="./Schedule.html">�Љ�\��h��</a>
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
                            <a href="./Trainer.html">�g���[�i�[�̏Љ�</a>
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


                
<%
Dim input1
Dim input2
Dim input3
Dim input4
Dim input5
Dim input6
Dim check
input1 = Request.Form("input1")
input2 = Request.Form("input2")
input3 = Request.Form("input3")
input4 = Request.Form("input4")
input5 = Request.Form("input5")
input6 = Request.Form("input6")
check = Request.Form("vehicle1")


	'	Response.Write "username= " & input1 & "<br>"
	'	Response.Write "company= " & input2 & "<br>"
	'	Response.Write "pref= " & input3 & "<br>"
	'	Response.Write "mail= " & input4 & "<br>"
	'	Response.Write "mobileno= " & input5 & "<br>"
	'	Response.Write "doukitext= " & input6 & "<br>"
	'	Response.Write "check= " & check & "<br>"

%>
                <form class="form" action="/mail/recruit-salemail.asp" method="POST" >
                    <div class="main-container">
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">����</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8"id="input1" name="input1" type="text" aria-label="text"  value="<% Response.Write input1 %>"  />
                                <!-- <span class="example-name">��F�R�c ���Y</span> -->
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">�ӂ肪��</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" id="input2" name="input2" type="text" aria-label="text"  value="<% Response.Write input2 %>"  />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">���[���A�h���X</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" type="email" id="input3" name="input3" aria-label="email"
                                     value="<% Response.Write input3 %>"  />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">�d�b�ԍ��i�g�ѓd�b�j</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" type="number" id="input4" name="input4" aria-label="number"
                                     value="<% Response.Write input4 %>"  />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">�L�����A</span>
                                <div class="badge-5"><span class="required-6">�K�{</span></div>
                            </div>
                            <div class="radio">
                                <input type="radio" id="input5" name="input5" checked="checked" value="<% Response.Write input5 %>" ><!-- name="fav_language"  -->
                                  <label for="html"><% Response.Write input5 %></label><br><!-- �{�Љc�ƕ��i�Љ�R���T���j  -->
                            </div>
                        </div>

                        <div class="main-container-3">
                            <div class="left">
                                <span class="motivation">�X�L��
                                
                                </span><button class="badge"><span
                                        class="required">�K�{</span></button>
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
        <footer>

            <div class="footer-left">
                <div class="footer-header">
                    <div class="footer-logo">
                        <img src="images/Company.svg" alt="company">
                    </div>
                    <!-- <div class="social">
                        <a href="#"><img src="images/x.svg" alt="x">
                        </a>
                        <a href="#">
                            <img src="images/facebook.svg" alt="facebook">
                        </a>
                        <a href="#">
                            <img src="images/instagram.svg" alt="instagram">
                        </a>
                    </div> -->
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
                        <div class="footer-item"><a href="./Trainer.html">�g���[�i�[�̏Љ�</a></div>
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