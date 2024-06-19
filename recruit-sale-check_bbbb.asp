
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
        <!-- #INCLUDE VIRTUAL="/include/navibar_sjis.html" -->

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

                <form class="form" action="/mail/recruit-salemail.asp" method="POST" >
                    <div class="main-container">
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">氏名</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8"id="input1" name="input1" type="text" aria-label="text"  />
                                <!-- <span class="example-name">例：山田 太郎</span> -->
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">ふりがな</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" id="input2" name="input2" type="text" aria-label="text"   />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">メールアドレス</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" type="email" id="input3" name="input3" aria-label="email"/>
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">電話番号（携帯電話）</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" type="number" id="input4" name="input4" aria-label="number"/>
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">キャリア</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>

                           <div class="name-7">
                                <input class="name-input-8" id="input5" name="input5" type="text" aria-label="text"/>
                            </div>

                             < -- <div class="radio"> -->
                                <!-- <input type="radio" id="input5" name="input5" checked="checked" value="<% Response.Write input5 %>" > --><!-- name="fav_language"  -->
                                  <!-- <label for="html"><% Response.Write input5 %></label><br> --><!-- 本社営業部（紹介コンサル）  -->
                             </div> -
                        </div>

                        <div class="main-container-3">
                            <div class="left">
                         <div class="main-container-3">
                            <div class="left">
                                <span class="motivation">スキル
                                
                                </span><button class="badge"><span
                                        class="required">必須</span></button>
                            </div>
                           <div class="name-7">
                                <input class="name-input-8" id="input6" name="input6" type="text" aria-label="text"/>
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