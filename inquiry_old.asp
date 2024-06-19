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
		<title>�l�ޔh���E�l�ޏЉ�X�z�[���y�[�W�@���₢����</title>

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
					<a href="/Job seekers.html">���d�������T���̋��E�җl</a>
					<div class="line_1"></div>
					<a href="/human resurs.html">�l�ނ����T���̗̍p�S���җl</a>
				</div>
			</div>
		
			<!-- navigation start -->
			<div class="navigation">
				<div class="nav_left">
					<a href="/index.html" class="home">
						<img src="/images/home.svg" alt="home">
						<div class="home_text">�z�[��</div>
					</a>
		
					<a href="/company.html">��Џ��</a>
					<a href="/Tokyo-branch.html">�x�Џ��</a>
					<a href="/Topics.html">�g�s�b�N�X</a>
					<a href="/inquiry.asp">���₢���킹</a>
				</div>
				<div class="nav_right">
					<a href="./desktop-20.html">�l���ی���j�E��舵��</a>
					<img src="images/P-mark.svg" alt="">
				</div>
			</div>
			<!-- navigation END -->
		
			<!----- submenu  ----->
			<nav class="submenu">
				<ul class="submenu-links">
				  <li class="submenu-dropdown">
					<a href="#">�]�E�T�|�[�g</a>
					<div class="dropdown">    
						<a href="/Temporary-staffing.html">�l�ޔh��</a>
						<a href="/Labor-regulations.html">�A�ƋK��</a>
						<a href="/Qualification.html">���i�x���擾���x</a>
						<a href="/Recruitment.html">�l�ޏЉ�</a>
						<a href="/Introduction.html">�Љ�\��h��</a>
						<a href="/q&a.html">�p���`</a>
					</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">�l�ރT�[�r�X</a>
					<div class="dropdown">
						<a href="/Dispatch.html">�h��</a>
						<a href="/Prelusion.html">�Љ�</a>
						<a href="/Schedule.html">�Љ�\��h��</a>
                	</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">���l�L���̌f��</a>
					<div class="dropdown">
						<a href="/Service-contents.html">�T�[�r�X���e</a>
						<a href="/Job-applicant.html">���E�ҏ��</a>
						<a href="/Prices.html">�v�����E����</a>
						<a href="/Questions.html">�p���`</a>
					</div>
				  </li>
				  <li class="submenu-dropdown">
					<a href="#">�̗p���</a>
					<div class="dropdown">
					  <a href="/recruit.html">�V���̗p���</a>
					  <a href="./recruit.html">���r�̗p���</a>
					</div>
				  </li>
				</ul>
			</nav>
			<!----- submenu end ----->  
			<!--- navbar END -------------->
		</div>
		
		<!------------ step -------------->
		<div class="step">
            <div class="item">�z�[��<img src="images/chevron_right.svg" alt=""></div>
            <div class="item-active">���₢���킹</div>
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

							<span class="furigana">�����O</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input name="name" class="name-input-8" type="text" aria-label="text" placeholder="��F�R�c ���Y" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">��Ж�</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input name="company" class="name-input-8" type="text" aria-label="text" placeholder="��F���X�������"/>
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">�s���{��</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input name="prefecture" class="name-input-8" type="text" aria-label="text" placeholder="��F��ʌ�" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">���[���A�h���X</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input name="mail" class="name-input-8" type="email" aria-label="email" placeholder="��Flis@lis21.co.jp" />
						</div>
					</div>
					<div class="frame-1">
						<div class="left-4">
							<span class="furigana">�d�b�ԍ��i�g�ѓd�b�j</span>
							<div class="badge-5"><span class="required-6">�K�{</span></div>
						</div>
						<div class="name-7">
							<input name="tel" class="name-input-8" type="number" aria-label="number" placeholder="��F07069505055" />
						</div>
					</div>
					
                   
					<div class="main-container-3">
						<div class="left">
						  <span class="motivation">���e</span
						  ><button class="badge"><span class="required">�K�{</span></button>
						</div>
						<textarea class="right" name="body" cols="30" rows="10" placeholder=""></textarea>
					  </div>
					  	<div class="chek-box">
							<input type="checkbox" id="vehicle1" name="vehicle1" value="Bike">
							<label for="vehicle1"> <a href="#">�l���̎�舵��</a></label>
						</div>
						<div class="submit">
							<input type="submit" value="�l���̎�舵���ɂ��� ���ӂ��đ��M����">
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
'  ���[���𑗐M���܂�
'    �p�����[�^
'      subj : �T�u�W�F�N�g
'      body : �{��
'=============================================================

	mailfrom = Request.QueryString("mail")
	subj = Request.QueryString("subject")

	subj = "�����XHP���@���₢����"
	body = "�y���O�z" & Request.QueryString("name")
	body = body & vbCrLf & "�y��Ж��z" & Request.QueryString("company")
	body = body & vbCrLf & "�y���ݒn�z" & Request.QueryString("prefecture")
	body = body & vbCrLf & "�y�d�b�ԍ��z" & Request.QueryString("tel")
	body = body & vbCrLf & "---------------------���@���e�@��---------------------"
	body = body & vbCrLf & Request.QueryString("body")

	rc = SndMail("smtp.office365.com","lis@lis21.co.jp", mailfrom, subj, body, "")

	If rc = True Then
		Response.Write "<div style=""padding-top:20px; height:150px; text-align:center;"">���⍇���L��������܂����B<BR>����A���Љc�Ƃ�育�A�������Ē����܂��B</div>"
	Else
		Response.Write "<div style=""padding-top:20px; height:150px; text-align:center;""><font color=red>���[�����M�Ɏ��s���܂����B</font><BR>�L�����ꂽ���[���A�h���X�������������m�F���������B<BR>�u���E�U�́u�߂�v�{�^���ł��߂肭�������B<P><font size=1>" & rc & ";" & bc &"</font><P></div>"
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
							<div class="title">���X�������</div>
							<div class="footer-item"><a href="./company.html">��Џ��</a></div>
							<div class="footer-item"><a href="#">�x�Џ��</a></div>
							<div class="footer-item"><a href="./Topics.html">�g�s�b�N�X</a></div>
							<div class="footer-item"><a href="./desktop-20.html">�l���ی�</a></div>
							<div class="footer-item"><a href="#">���₢����</a></div>
							<div class="footer-item"><a href="#">�T�C�g�}�b�v</a></div>
						</div>
						<div class="vertical">
							<div class="title">�]�E�T�|�[�g</div>
							<div class="footer-item"><a href="./desktop-7.html">�l�ޔh��</a></div>
							<div class="footer-item"><a href="./desktop-8.html">�A�ƋK��</a></div>
							<div class="footer-item"><a href="./point.html">���i�x���擾���x</a></div>
							<div class="footer-item"><a href="./desktop-10.html">�l�ޏЉ�</a></div>
							<div class="footer-item"><a href="./desktop-11.html">�Љ�\��h��</a></div>
							<div class="footer-item"><a href="./q&a.html">�p���`</a></div>
						</div>
						<div class="vertical">
							<div class="title">�l�ރT�[�r�X</div>
							<div class="footer-item"><a href="./desktop-13.html">�h��</a></div>
							<div class="footer-item"><a href="./desktop-14.html">�Љ�</a></div>
							<div class="footer-item"><a href="./desktop-15.html">�Љ�\��h��</a></div>
						</div>
						<div class="vertical">
							<div class="title">���l�L���̌f��</div>
							<div class="footer-item"><a href="./desktop-16.html">�T�[�r�X���e</a></div>
							<div class="footer-item"><a href="./desktop-17.html">���E�ҏ��</a></div>
							<div class="footer-item"><a href="./desktop-18.html">�v�����E����</a></div>
							<div class="footer-item"><a href="./desktop-19.html">�p���`</a></div>
						</div>
						<div class="vertical">
							<div class="title">�̗p���</div>
							<div class="footer-item"><a href="./recruit.html">�V���̗p���</a></div>
							<div class="footer-item"><a href="./recruit.html">���r�̗p���</a></div>
							<div class="footer-item"><a href="./recruit.html">�A���o�C�g</a></div>
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