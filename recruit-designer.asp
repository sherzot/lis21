<%@ Language=VBScript CodePage=932 %>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="SJIS">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans:ital,wght@0,100..900;1,100..900&display=swap"
        rel="stylesheet">
	<script  src="js/jquery-1.12.4.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
    <link rel="stylesheet" href="css/recruit-sale.css">
    <script src="js/app.js"></script>
    <script defer src="js/mobile.js"></script>
	<script  src="js/jquery-1.12.4.min.js"></script>
    <link rel="icon" href="images/logo-small.svg" type="image/icon type">
    <title>採用情報</title>
</head>

<body>
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

            <div class="right-content">
                <div class="right-content-title">フロントエンドエンジニア（本社）</div>
                <div class="job-text">
                    <h3>
                        仕事内容
                    </h3>
                    <p>
                        フロントエンドエンジニア経験者を募集。
                        <br>フロントエンドエンドエンジニアの経験者を募集しています
                        <br>Webデザイナー、バックエンドエンジニアと協力してWebサイトを仕上げるお仕事です。
                        <br>HTML,CSS(sass,
                        scss),Javascriptの経験者を求めます。レスポンシブなデザインを実現するためにBootstrapが必要です。Figmaでデザインを共有して仕事ができる方を求めています。詳細はご相談ください。
                        <br>ふるってご応募ください。
                    </p><br>
                    <p>
                        応募資格・条件
                        <br>学歴不問。
                    </p><br>
                    <h3>
                        【必須スキル＆経験】
                    </h3><br>
                    <p>
                        フロントエンドエンジニアとしての経験を重視。
                        <br>HTML,CSS(sass, scss),Javascriptの経験をお持ちの方。
                    </p><br>
                    <h3>
                        【歓迎スキル＆経験】
                    </h3><br>
                    <p>
                        ※下記のいずれかの経験をお持ちの方大歓迎です！
                        <br>・スマホ対応の開発経験。レスポンシブなデザインをできる方。
                    </p><br>
                    <h3>
                        勤務地
                    </h3>
                    <p>
                        
                        <br>リス株式会社
                        <br>〒163-0825 東京都新宿区西新宿２丁目４番１号 新宿ＮＳビル２５階 (企画推進室)
                        <br>◇ 転勤なし
                        <br>◇[最寄駅]都庁前駅(徒歩5分)/新宿駅(徒歩13分)<!--  駅から徒歩13分以内  -->
				        <br>　[沿線]都営大江戸線,京王線 
                    </p><br>
                    <h3>
                        時間
                    </h3>
                    <p>
                        9:00 ～ 18:00
                        <br>※実働8時間／休憩1時間
                    </p><br>
                    <h3>
                        正社員<!-- ■   -->
                    </h3>
                    <p>
                        月給　300,000円　～　600,000円　
                    </p><br>
                    <p>
                        <br>※上記額にはみなし残業代（月40時間）を含みます。超過分は全額支給します。
                    </p><br>
                    <h3>
                        休日
                    </h3><br>
                    <p>
                        休日：完全週休２日
                        <br>休日備考：土・日・祝（夏季休暇；年末年始；特別休暇）（誕生日休暇3日）
                       
                    </p><br>
                    <h3>
                        選考手順：
                    </h3>
                    <p>
                        <br>・ステップ1 書類選考
                        <br>・ステップ2 １次面接
                        <br>・ステップ3 最終面接
                    </p>
                     <p>
                        <br><b>担当部署：管理部：採用担当</b>
                        <br>連絡先：※お問い合わせの際、「リスホームページを見た」と言っていただくとスムーズです。
                    </p>
   
                </div>

                <form class="form" action="./recruit-check.asp" method="POST">
                    <div class="main-container">
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">氏名</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" name="input1" type="text" aria-label="text" placeholder="例：山田 太郎" />
                                <!-- <span class="example-name">例：山田 太郎</span> -->
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">ふりがな</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" name="input2" type="text" aria-label="text" placeholder="例：やまだ たろう" />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">メールアドレス</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" name="input3" type="email" aria-label="email"
                                    placeholder="例：lis@lis21.co.jp" />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">電話番号（携帯電話）</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="name-7">
                                <input class="name-input-8" name="input4" type="number" aria-label="number"
                                    placeholder="例：080******** 数字のみ" />
                            </div>
                        </div>
                        <div class="frame-1">
                            <div class="left-4">
                                <span class="furigana">応募職種</span>
                                <div class="badge-5"><span class="required-6">必須</span></div>
                            </div>
                            <div class="radio">
                                <input type="radio" id="html"  name="input5" value=フロントサイドエンジニア"
                                 <label for="html">フロントサイドエンジニア</label><br>
                            </div>
                        </div>

                        <div class="main-container-3">
                            <div class="left">
                                <span class="motivation">ご応募動機</span><button class="badge"><span
                                        class="required">必須</span></button>
                            </div>
                            <textarea class="right"  name="input6" id="" cols="30" rows="10"
                                placeholder="ご応募動機をご入力ください"></textarea>
                        </div>
                        <div class="chek-box">
                            <input type="checkbox" id="vehicle1" name="vehicle1" value="同意" onclick="checkValue(this)">
                            <label for="vehicle1"> <a href="#">個人情報の取り扱い</a></label>
                        </div>
                        <div class="submit">
                            <input type="submit" id="btn" disabled="disabled" value="個人情報の取り扱いの「同意する」チェックをつけてください。"><!--  個人情報の取り扱いについて 同意して送信する  -->
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

</html>