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
    <title>採用情報</title>
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
                        <a href="./recruit-sale.html">採用情報</a>
                        <a href="./Privacy-policy.html">個人情報保護方針</a>
                        <a href="./Temporary-staffing.html">人材派遣</a>
                        <a href="./Recruitment.html">人材紹介</a>
                        <a href="./Introduction.html">紹介予定派遣</a>
                        <a href="./Trainer.html">トレーナーの紹介</a>
                        <a href="./q&a.html">Ｑ＆Ａ</a>
                        <a href="./Dispatch.html">派遣</a>
                        <a href="./Prelusion.html">紹介</a>
                        <a href="./Schedule.html">紹介予定派遣</a>
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
                            <a href="./Trainer.html">トレーナーの紹介</a>
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

        <!------------ step -------------->
        <div class="step">
            <div class="item">ホーム<img src="images/Web/chevron_right.svg" alt=""></div>
            <div class="item-active">採用情報</div>
        </div>

        <!---------- step END ----------------->

        <!-- HEADER -->

        <div class="header">
            <!-- list-menu-left -->
            <div class="list-menu">
                <div class="list-menu-title">募集職種</div>
                <ul>
                    <li><a href="./recruit-sale.html">紹介コンサル（全国７拠点）</a></li>
                    <li><a href="./recruit-system.html">バックエンドエンジニア（本社）</a></li>
                    <li><a href="./recruit-designer.html">フロントエンドエンジニア（本社）</a></li>
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
                                <span class="furigana">氏名</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8"id="input1" name="input1" type="text" aria-label="text"  value="<% Response.Write input1 %>"  />
                                <!-- <span class="example-name">例：山田 太郎</span> -->
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">ふりがな</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" id="input2" name="input2" type="text" aria-label="text"  value="<% Response.Write input2 %>"  />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">メールアドレス</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" type="email" id="input3" name="input3" aria-label="email"
                                     value="<% Response.Write input3 %>"  />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">電話番号（携帯電話）</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" type="number" id="input4" name="input4" aria-label="number"
                                     value="<% Response.Write input4 %>"  />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">キャリア</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="radio">
                                <input type="radio" id="input5" name="input5" checked="checked" value="<% Response.Write input5 %>" ><!-- name="fav_language"  -->
                                  <label for="html"><% Response.Write input5 %></label><br><!-- 本社営業部（紹介コンサル）  -->
                            </div>
                        </div>

                        <div class="main-container-3">
                            <div class="left">
                                <span class="motivation">スキル
                                
                                </span><button class="badge"><span
                                        class="required">必須</span></button>
                            </div>
                            <!-- <textarea class="right"  id="input6" name="input6"   cols="30" rows="10"
                                 value="<% Response.Write input6 %>" "><% Response.Write input6 %></textarea> --><!-- name="ご応募動機をご入力ください"  -->
                        	<input type="hidden" name="input7"   value="本社営業部（管理職・役職経験者）">
                        </div>
                        <div class="chek-box">
                            <input type="checkbox" id="vehicle1" name="vehicle1"  value="同意"  checked="on" ><!-- <% Response.Write check %> -->
                            <!--  value="Bike"  -->
                            <label for="vehicle1"> <a href="#">個人情報の取り扱い</a></label>
                        </div>
                        
						<div class="submit">
						  <input type="button" value="戻って修正する" onclick="history.back();"  id="btnBack">
						</div>                        
                        <div class="submit">
                            <input type="submit" id="btn" value="メールを送信する">
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
                        <div class="footer-item"><a href="./Trainer.html">トレーナーの紹介</a></div>
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
			
		} else if (!(input2.value)) {
		//	//btn.value="御社名が入力されていません。戻って修正してください。";
			 alert('ふりがなが入力されていません');
			btn.setAttribute('disabled','disabled');
		} else if (!(input3.value)) {

			 alert('メールアドレスが入力されていません');		
			btn.setAttribute('disabled','disabled');
		} else if (!input4.value) {
			//btn.value="電話番号が入力されていません。戻って修正してください。";
			 alert('電話番号が入力されていません');
			btn.setAttribute('disabled','disabled');		
			
		} else if (!input5.value) {
			//btn.value="応募職種が入力されていません。戻って修正してください。";
			 alert('キャリアが入力されていません');
			btn.setAttribute('disabled','disabled');		
			
		} else if (!(input6.value)) {
			
			 alert('スキルが入力されていません。戻って修正してください。');
			btn.setAttribute('disabled','disabled');

		} else {
			btn.value="メールを送信する。";
			btn.removeAttribute('disabled');
		}//end if
	});//end function
</script>

</html>