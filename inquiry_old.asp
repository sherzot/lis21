<!-- #INCLUDE VIRTUAL="/config/personnel.asp" -->
<!-- #INCLUDE VIRTUAL="/config/constant.asp" -->
<!-- #INCLUDE VIRTUAL="/include/commonfunc.asp" -->
<!-- #INCLUDE VIRTUAL="/include/connect.asp" -->

<!DOCTYPE html>
<html lang="en">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
		<link href="https://fonts.googleapis.com/css2?family=Noto+Sans:ital,wght@0,100..900;1,100..900&display=swap" rel="stylesheet">
		<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
		<link rel="stylesheet" href="/css/inquiry.css">
		<script defer src="js/app.js"></script>
		<title>人材派遣・人材紹介リスホームページ　お問い合せ</title>

<%
Dim rc
Dim mailfrom
Dim subj
Dim body
Dim test
%>
	</head>
<body>
    <div class="container">
		<!-------------- navbar -------------------->
		<div class="navbar">
			<div class="top_nav">
				<div class="logo">
					<a href="./index.html">
						<img src="images/logo.svg" alt="logo">
					</a>
				</div>
				<div class="left_link">
					<a href="/Job seekers.html">お仕事をお探しの求職者様</a>
					<div class="line_1"></div>
					<a href="/human resurs.html">人材をお探しの採用担当者様</a>
				</div>
			</div>
		
			<!-- navigation start -->
			<div class="navigation">
				<div class="nav_left">
					<a href="/index.html" class="home">
						<img src="/images/home.svg" alt="home">
						<div class="home_text">ホーム</div>
					</a>
		
					<a href="/company.html">会社情報</a>
					<a href="/Tokyo-branch.html">支社情報</a>
					<a href="/Topics.html">トピックス</a>
					<a href="/inquiry.asp">お問い合わせ</a>
				</div>
				<div class="nav_right">
					<a href="./desktop-20.html">個人情報保護方針・取り扱い</a>
					<img src="images/P-mark.svg" alt="">
				</div>
			</div>
			<!-- navigation END -->
		
			<!----- submenu  ----->
			<nav class="submenu">
				<ul class="submenu-links">
				  <li class="submenu-dropdown">
					<a href="#">転職サポート</a>
					<div class="dropdown">    
						<a href="/Temporary-staffing.html">人材派遣</a>
						<a href="/Labor-regulations.html">就業規則</a>
						<a href="/Qualification.html">資格支援取得制度</a>
						<a href="/Recruitment.html">人材紹介</a>
						<a href="/Introduction.html">紹介予定派遣</a>
						<a href="/q&a.html">Ｑ＆Ａ</a>
					</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">人材サービス</a>
					<div class="dropdown">
						<a href="/Dispatch.html">派遣</a>
						<a href="/Prelusion.html">紹介</a>
						<a href="/Schedule.html">紹介予定派遣</a>
                	</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">求人広告の掲載</a>
					<div class="dropdown">
						<a href="/Service-contents.html">サービス内容</a>
						<a href="/Job-applicant.html">求職者情報</a>
						<a href="/Prices.html">プラン・料金</a>
						<a href="/Questions.html">Ｑ＆Ａ</a>
					</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">採用情報</a>
					<div class="dropdown">
					  <a href="/recruit.html">新卒採用情報</a>
					  <a href="./recruit.html">中途採用情報</a>
					</div>
				  </li>
				</ul>
			</nav>
			<!----- submenu end ----->  
			<!--- navbar END -------------->
		</div>
		
		<!------------ step -------------->
		<div class="step">
            <div class="item">ホーム<img src="images/chevron_right.svg" alt=""></div>
            <div class="item-active">お問い合わせ</div>
         </div>
		<!---------- step END ----------------->

<% If Request.QueryString("mail_flag") = "" Then %>
		<!-- HEADER -->
		<div class="header">
			<!-- form-section -->
			<form class="form" action="">
				<div class="main-container">
					<div class="frame-1">
						<div class="left-4">

							<span class="furigana">お名前</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input name="name" class="name-input-8" type="text" aria-label="text" placeholder="例：山田 太郎" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">御社名</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input name="company" class="name-input-8" type="text" aria-label="text" placeholder="例：リス株式会社"/>
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">都道府県</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input name="prefecture" class="name-input-8" type="text" aria-label="text" placeholder="例：埼玉県" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">メールアドレス</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input name="mail" class="name-input-8" type="email" aria-label="email" placeholder="例：lis@lis21.co.jp" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">電話番号（携帯電話）</span>
							<div class="badge-5"><span class="required-6">必須</span></div>
						</div>
						<div class="name-7">
							<input name="tel" class="name-input-8" type="number" aria-label="number" placeholder="例：07069505055" />
						</div>
					</div>
					
                   
					<div class="main-container-3">
						<div class="left">
						  <span class="motivation">内容</span
						  ><button class="badge"><span class="required">必須</span></button>
						</div>
						<textarea class="right" name="body" cols="30" rows="10" placeholder=""></textarea>
					  </div>
					  	<div class="chek-box">
							<input type="checkbox" id="vehicle1" name="vehicle1" value="Bike">
							<label for="vehicle1"> <a href="#">個人情報の取り扱い</a></label>
						</div>
						<div class="submit">
							<input type="submit" value="個人情報の取り扱いについて 同意して送信する">
							<input type="hidden" name="mail_flag" value="1">
						</div>
				</div>
			</form>
		</div>
<% Else %>
<!-- HEADER -->
<div class="header">
<%
'======================== send.asp ========================
'  メールを送信します
'    パラメータ
'      subj : サブジェクト
'      body : 本文
'=============================================================

	mailfrom = Request.QueryString("mail")
	subj = Request.QueryString("subject")

	subj = "■リスHP■　お問い合せ"
	body = "【名前】" & Request.QueryString("name")
	body = body & vbCrLf & "【会社名】" & Request.QueryString("company")
	body = body & vbCrLf & "【所在地】" & Request.QueryString("prefecture")
	body = body & vbCrLf & "【電話番号】" & Request.QueryString("tel")
	body = body & vbCrLf & "---------------------■　内容　■---------------------"
	body = body & vbCrLf & Request.QueryString("body")

	rc = SndMail("smtp.office365.com","lis@lis21.co.jp", mailfrom, subj, body, "")

	If rc = True Then
		Response.Write "<div style=""padding-top:20px; height:150px; text-align:center;"">お問合せ有難うございました。<BR>後日、弊社営業よりご連絡させて頂きます。</div>"
	Else
		Response.Write "<div style=""padding-top:20px; height:150px; text-align:center;""><font color=red>メール送信に失敗しました。</font><BR>記入されたメールアドレスが正しいかご確認ください。<BR>ブラウザの「戻る」ボタンでお戻りください。<P><font size=1>" & rc & ";" & bc &"</font><P></div>"
	End If
%>
</div>
<% End if %>


			<!---------------- Footer ------------------->
			<footer>
					
				<div class="footer-left">
					<div class="footer-header">
						<div class="footer-logo">
							<img src="images/logo.svg" alt="company">
						</div>
						<div class="social">
							<a href="#"><img src="images/x.svg" alt="x">
							</a>
							<a href="#">
								<img src="images/facebook.svg" alt="facebook">
							</a>
							<a href="#">
								<img src="images/instagram.svg" alt="instagram">
							</a>
						</div>
					</div>
			
					<div class="footer-content">
						<div class="vertical">
							<div class="title">リス株式会社</div>
							<div class="footer-item"><a href="./company.html">会社情報</a></div>
							<div class="footer-item"><a href="#">支社情報</a></div>
							<div class="footer-item"><a href="./Topics.html">トピックス</a></div>
							<div class="footer-item"><a href="./desktop-20.html">個人情報保護</a></div>
							<div class="footer-item"><a href="#">お問い合せ</a></div>
							<div class="footer-item"><a href="#">サイトマップ</a></div>
						</div>
						<div class="vertical">
							<div class="title">転職サポート</div>
							<div class="footer-item"><a href="./desktop-7.html">人材派遣</a></div>
							<div class="footer-item"><a href="./desktop-8.html">就業規則</a></div>
							<div class="footer-item"><a href="./point.html">資格支援取得制度</a></div>
							<div class="footer-item"><a href="./desktop-10.html">人材紹介</a></div>
							<div class="footer-item"><a href="./desktop-11.html">紹介予定派遣</a></div>
							<div class="footer-item"><a href="./q&a.html">Ｑ＆Ａ</a></div>
						</div>
						<div class="vertical">
							<div class="title">人材サービス</div>
							<div class="footer-item"><a href="./desktop-13.html">派遣</a></div>
							<div class="footer-item"><a href="./desktop-14.html">紹介</a></div>
							<div class="footer-item"><a href="./desktop-15.html">紹介予定派遣</a></div>
						</div>
						<div class="vertical">
							<div class="title">求人広告の掲載</div>
							<div class="footer-item"><a href="./desktop-16.html">サービス内容</a></div>
							<div class="footer-item"><a href="./desktop-17.html">求職者情報</a></div>
							<div class="footer-item"><a href="./desktop-18.html">プラン・料金</a></div>
							<div class="footer-item"><a href="./desktop-19.html">Ｑ＆Ａ</a></div>
						</div>
						<div class="vertical">
							<div class="title">採用情報</div>
							<div class="footer-item"><a href="./recruit.html">新卒採用情報</a></div>
							<div class="footer-item"><a href="./recruit.html">中途採用情報</a></div>
							<div class="footer-item"><a href="./recruit.html">アルバイト</a></div>
						</div>
					</div>
			
					<div class="copyright">Copyright(c) 2024 LIS co.,Ltd. All rights Reserved.</div>
				</div>
				<div class="footer-map">
					<iframe title="myFrame"
						src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d6481.103588769505!2d139.69093757604216!3d35.68803667258506!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x60188cd3741a9df7%3A0x4fb5f8fb9f0a0195!2sShinjuku%20NS%20Building!5e0!3m2!1sen!2sjp!4v1715059757728!5m2!1sen!2sjp"
						width="600" height="450" style="border:0;" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe>
			
				</div>
			
			</footer>
			
			
			<!--------------- Footer end ----------------->
	</div>
</body>
</html>