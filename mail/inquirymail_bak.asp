<%@ Language=VBScript CodePage=932 %>
<% Option Explicit %>
<%
'******************************************************************************
'�T�@�v�F���[�����M���
'���@�l�F
'�X�@�V�F2024/05/17 LIS katayama �쐬�i���V�e������̉��C�j
'�@�@�@�F2008/03/13 LIS K.katayama ���[�����M��, 
'�@�@�@�F2008/05/07 LIS K.katayama �₢���킹���[��

'�@�@�@�F2011/01/05 LIS K.Kokubo Basp.SendMail �� SndMail
'******************************************************************************
%>

<!-- #INCLUDE VIRTUAL="/include/commonfunc.asp" -->
<html lang="ja">
<head>
		<meta charset="sjis">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>�₢���킹���[���e�X�g</title>
	</head>
<body>


<%

		''susername = Request.Form("username")
		'���[�����M����
		sTo = "mkatayama@lis21.co.jp" ''sReceiverMailAddress	'���M�惁�[���A�h���X
		
		sFrom = "lis@lis21.co.jp" ''Cnt_NaviMailAddress	'���M�����[���A�h���X
		
		'�^�C�g��
		sSubject = "���XHP�e�X�g���[��" ''MAIL_FROM_STAFF_SUBJECT

		'�{��
		'' sBody = "testmail: username= " & VbCrlf & susername


		 
		 'Cnt_MailServer''Const Cnt_MailServer = "172.16.1.39"
		'sResult = SndMail(Cnt_MailServer, sTo, sFrom, sSubject, sBody, "")
		 ''Response.Write "Result= " & sResult & "<br>"
		 ''Response.Write "sBody= " & sBody & "<br>"
		 If sResult=True Then
		 	Response.Write "���M���������܂����B"
		 else
		 	Response.Write "���M�����s���܂����B"
		 End if
		

%>
 
		<!--------------- Footer start ----------------->
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
</html>

