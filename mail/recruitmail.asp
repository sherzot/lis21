<%@ Language=VBScript CodePage=932 %>
<% Option Explicit %>
<%
'******************************************************************************
'�T�@�v�F���[�����M���
'���@�l�F
'�X�@�V�F2024/05/17 LIS katayama �쐬�i���V�e������̉��C�j
'�@�@�@�F2008/03/13 LIS K.Kokubo ���[�����M��, 
'�@�@�@�F2008/05/07 LIS K.Kokubo ���l�[�̉{���ۂ�ChkOrderDsp�Ŕ��肷��悤�ɕύX

'�@�@�@�F2011/01/05 LIS K.Kokubo Basp.SendMail �� SndMail
'******************************************************************************
%>

<!-- #INCLUDE VIRTUAL="/include/commonfunc.asp" -->
<!-- #INCLUDE VIRTUAL="/config/personnel.asp" -->
<html lang="ja">
	<head>
		<meta charset="sjis">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta charset="SJIS">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<link href="https://fonts.googleapis.com/css2?family=Noto+Sans:ital,wght@0,100..900;1,100..900&display=swap"
		rel="stylesheet">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
	<link rel="stylesheet" href="/css/inquiry.css">
	<script defer src="/js/app.js"></script>
	<script defer src="/js/mobile.js"></script>
	<link rel="icon" href="images/logo-small.svg" type="/image/icon type">
	<title>�̗p����E�₢���킹���[��</title>
    
	</head>
<body>

	<div class="container">
		<!-------------- navbar -------------------->
		<div class="navbar">
			<div class="top_nav">
				<div class="logo">
					<a href="/index.html">
						<img src="/images/Web/Logo/Web/Desctop.svg" alt="">
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
					<a href="/index.html" class="active">
						<img src="/images/Mobile/Web/Logo/Phone.svg" alt="">
					</a>
					<!-- Navigation links (hidden by default) -->
					<div id="myLinks">
						<a href="/Company.html">��Џ��</a>
						<a href="/Tokyo-branch.html">�x�Џ��</a>
						<a href="#">�̗p���</a>
						<a href="#">�]�E�T�|�[�g</a>
						<a href="/recruit-sale.html">�l�ރT�[�r�X</a>
						<a href="/inquiry.html">���₢���킹</a>
					</div>
					<!-- "Hamburger menu" / "Bar icon" to toggle the navigation links -->
					<a href="javascript:void(0);" class="icon" onclick="myFunction()">
						<img src="/images/Mobile/Web/Menu.svg" alt="">
					</a>
				</div>
			</div>
			<!-------------- mobile top_nav end ---------------->

			<!-- navigation -->
			<div class="navigation">
				<div class="nav_left">
					<a href="/Company.html">��Џ��</a>
					<a href="/Tokyo-branch.html">�x�Џ��</a>
					<!--<a href="/Topics.html">�g�s�b�N�X</a>-->
					<a href="/inquiry.html">���₢���킹</a>
					<a href="/recruit-sale.html" class="home">
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
							<a href="/Temporary-staffing.html">�l�ޔh��</a>
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
				</ul>
			</nav>
			<!----- submenu end ----->
		</div>
		<!--- navbar END -------------->


<%
'' Const Cnt_MailServer = "172.16.1.39" ''personnel�ɏ�����
Dim sTo
Dim sFrom
Dim sBody
Dim sSubject
dim sResult
Dim susername
Dim input1
Dim input2
Dim input3
Dim input4
Dim input5
Dim input6
Dim doui

	input1 = Request.Form("input1")
	input2 = Request.Form("input2")'
	input3 = Request.Form("input3")	
	input4 = Request.Form("input4")
	input5 = Request.Form("input5")	
	input6 = Request.Form("input6")
	''input6 = Replace(input6, vbCrLf, "<BR>")
	input6 = Server.HTMLEncode(input6)
	doui = Request.Form("vehicle1")

	input1 = Trim(input1)
	input2 = Trim(input2)
	input3 = Trim(input3)
	input4 = Trim(input4)
	input5 = Trim(input5)
	input6 = Trim(input6)
	doui = Trim(doui)
	
'	Response.Write "username= " & input1 & "<br>"
'	Response.Write "company= " & input2 & "<br>"
'	Response.Write "pref= " & input3 & "<br>"'
'	Response.Write "mail= " & input4 & "<br>"
''	Response.Write "����E��= " & input5 & "<br>"
'	Response.Write "doukitext= " & input6 & "<br>"
'	Response.Write "doui= " & doui & "<br>"

		'���[�����M����
		sTo = "lis@lis21.co.jp" ''sReceiverMailAddress	'���M�惁�[���A�h���X
		
		sFrom = "lis@lis21.co.jp" ''Cnt_NaviMailAddress	'���M�����[���A�h���X
		
		'�^�C�g��
		sSubject = "���XHP�̗p���僁�[��" ''MAIL_FROM_STAFF_SUBJECT

		'�{��
		sBody = " ����= " & input1 & vbCrLf  '
		sBody = sBody &  " �ӂ肪��= " & input2 & vbCrLf 
		sBody = sBody &  " ���[��= " & input3 & VbCrLf
		sBody = sBody &  " �d�b= " & input4 & vbCrLf 
		sBody = sBody &  " ����E��= " & input5 & vbCrLf 
		sBody = sBody &  " �u�]���@= " & input6 & vbCrLf 
		sBody = sBody & " �l���̓��� = " & doui  & vbCrLf
		 		
		 sResult = SndMail(Cnt_MailServer, sTo, sFrom, sSubject, sBody, "")''���[�����M����
		' Response.Write "Result= " & sResult & "<br>"
		  ' Response.Write "body= " & sBody & "<br>"
		 
		 
		 Response.Write "<div class='main-container-3'>"
		 If sResult=True Then
		 
		 Response.Write "	<div class='submit'>"
		 	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;���M���������܂����B<br>"
		 	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�S���҂���A���������グ�܂��B<br>"
		 	Response.Write "	</div>"
		 else
		 	Response.Write "	<div class='submit'>"
		 		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;���M�͎��s���܂����B<br>"
		 	Response.Write "	</div>"
				Response.Write "	<div class='submit'>"
				Response.Write "	    <input type='button' value='�߂��ďC������' onclick='history.back();'  id='btnBack'>"
				Response.Write "	</div>"
		 	
		 End if
			Response.Write " </div>"
%>
 
 
 		<!---------------- Footer ------------------->
		<footer>

			<div class="footer-left">
				<div class="footer-header">
					<div class="footer-logo">
						<img src="/images/Logo/Web/Desctop1.svg" alt="company">
					</div>
					<!-- <div class="social">
						<a href="#"><img src="./images/x.svg" alt="x">
						</a>
						<a href="#">
							<img src="./images/facebook.svg" alt="facebook">
						</a>
						<a href="#">
							<img src="./images/instagram.svg" alt="instagram">
						</a>
					</div> -->
				</div>

				<div class="footer-content">
					<div class="vertical">
						<div class="title">���X�������</div>
						<div class="footer-item"><a href="/Company.html">��Џ��</a></div>
						<div class="footer-item"><a href="Tokyo-branch.html">�x�Џ��</a></div>
						<!--<div class="footer-item"><a href="/Topics.html">�g�s�b�N�X</a></div>-->
						<div class="footer-item"><a href="/Privacy-policy.html">�l���ی�</a></div>
						<div class="footer-item"><a href="/recruit-sale.html">�̗p���</a></div>
						<div class="footer-item"><a href="/inquiry.html">���₢����</a></div>
					</div>
					<div class="vertical">
						<div class="title">�]�E�T�|�[�g</div>
						<div class="footer-item"><a href="/Temporary-staffing.html">�l�ޔh��</a></div>
						<div class="footer-item"><a href="/Recruitment.html">�l�ޏЉ�</a></div>
						<div class="footer-item"><a href="/Introduction.html">�Љ�\��h��</a></div>
						<div class="footer-item"><a href="/q&a.html">�p���`</a></div>
					</div>
					<div class="vertical">
						<div class="title">�l�ރT�[�r�X</div>
						<div class="footer-item"><a href="/Dispatch.html">�h��</a></div>
						<div class="footer-item"><a href="/Prelusion.html">�Љ�</a></div>
						<div class="footer-item"><a href="/Schedule.html">�Љ�\��h��</a></div>
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
</html>

