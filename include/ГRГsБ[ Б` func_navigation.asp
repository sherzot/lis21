<%
'******************************************************************************
'�T�@�v�F�w�b�_�[
'���@���FHeadType	0�y�g�b�v�z1�y���E�ҁz2�y��Ɓz3�y���p�z
'�쐬�ҁFLis Niina
'�쐬���F2008/02/07
'���@�l�F
'�g�p���F
'******************************************************************************
Function NaviHeader(HeadType)
	Dim sHeadcmt
	Dim sLinkurl
	Dim sLinkalt
	Dim sLinktext

	Dim sContents

	If HeadType = 0 Then '�g�b�v
		'sHeadcmt = "�@�]�E�����ɕK�v�ȏ���(��������)�̍쐬�E���d�����E�v���ɂ��M���ɓK�����]�E�T�|�[�g�����񋟂��Ă��܂��I"
		sHeadcmt = "<div style=""padding-left:8px; color:#ffffff;"">���l,��W���͂������A�������E�E���o�������̎����쐬�A�v���ɂ��M���ɓK�����]�E�T�|�[�g�����񋟂��Ă��܂��I</div>"
		sLinkurl = "/company/index.asp"
		sLinkalt = "���l�L�������ƃi�r"
		sLinktext = "�̗p�S���i���l��Ɓj�l�͂�����"
	ElseIf HeadType = 1 Then '���E��
		sHeadcmt = "<div style=""padding-left:8px;"">�]�E�����E���E�����̕��X�ɍœK�ȋ��l���Ɨ������c�[����񋟂��Ă��܂�</div>"
		sLinkurl = "/company/index.asp"
		sLinkalt = "���l�L�������ƃi�r"
		sLinktext = "�̗p�S���i���l��Ɓj�l�͂�����"
	ElseIf HeadType = 2 Then '���
		sHeadcmt = "<div style=""padding-left:8px;"">��Ƃ̐l�ތٗp�𕝍L���T�|�[�g���Ă���܂��B�i���l�L���A�l�ޔh���A�l�ޏЉ�j</div>"
		sLinkurl = "/"
		sLinkalt = "�]�E�E���l�T�C�g�����ƃi�r"
		sLinktext = "���d�������T���̕��͂�����"
	ElseIf HeadType = 3 Then '���p
		sHeadcmt = "<div style=""padding-left:8px;"">���E�����̕��X�ɍœK�ȋ��l���Ɨ������c�[����񋟂��Ă��܂�</div>"
		sLinkurl = "/company/index.asp"
		sLinkalt = "���l�L�������ƃi�r"
		sLinktext = "�̗p�S���i���l��Ɓj�l�͂�����"
	End If

	Response.Write "<div id=""wrap"" align=""center"">"
	Response.Write "<div id=""wrapw"">"
	Response.Write "<div id=""head"" align=""left"">"
	Response.Write "<table>"
	Response.Write "<tr>"

	'<�w�b�_�[���F�����ƃi�r���S>
	Response.Write "<td align=""left"" style=""height:42px; width:141px;"">"
	'�N���X�}�X
	If Month(Now) = 12 and (Day(now) > 9 and Day(now) < 26) Then
		Response.Write "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""https://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0"" width=""137"" height=""40"" align=""middle"">"
		Response.Write "<param name=""allowScriptAccess"" value=""sameDomain"">"
		Response.Write "<param name=""movie"" value=""/img/xmaslogo.swf"">"
		Response.Write "<param name=""Flashvars"" value="""">"
		Response.Write "<param name=""quality"" value=""high"">"
		Response.Write "<param name=""menu"" value=""false"">"
		Response.Write "<param name=""wmode"" value=""opaque"">"
		Response.Write "<embed src=""/img/xmaslogo.swf"" Flashvars="""" menu=""false"" quality=""high"" bgcolor=""#ffffff"" width=""137"" height=""40"" name=""stationmap"" align=""middle"" allowScriptAccess=""sameDomain"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"">"
		Response.Write "</object>"
	Else
		Response.Write "<a class=""decnone"" href=""/"" title=""�]�E�E���l�T�C�g�u�����ƃi�r�v""><img src=""/img/top/shigotonavi_logo.gif"" alt=""�����ƃi�r"" border=""0"" align=""left"" style=""margin-left:4px;""></a>"
	End If

	Response.Write "</td>"
	'</�w�b�_�[���F�����ƃi�r���S>

	'<�w�b�_�[�E>
	Response.Write "<td align=""right"" style=""font-size:11px;"">"
	Response.Write "�@<a href=""/staff/access.asp""><img src=""/img/top/head_icon.gif"" alt=""���⍇��"" border=""0"" style=""vertical-align:millde;"">���⍇��</a>"
	Response.Write "�@<a href=""" & HTTP_CURRENTURL & "shigotonavi/sitemap.asp""><img src=""/img/top/head_icon.gif"" alt=""�T�C�g�}�b�v"" border=""0"" style=""vertical-align:middle;"">�T�C�g�}�b�v</a>"

	'<Google�̃T�C�g������>
	If Request.ServerVariables("HTTPS") <> "on" Then
		Response.Write "<form action=""/search.asp"" id=""cse-search-box"" style=""margin-left:5px;padding:0px;display:inline"">"
		Response.Write "<div style=""display:inline;"">"
		Response.Write "<img src=""/img/top/head_icon.gif"" alt="""" border=""0"" style=""vertical-align:millde;"">"
		Response.Write "<label>"
		Response.Write "<span>�T�C�g������&nbsp;</span>"
		Response.Write "<input type=""hidden"" name=""cx"" value=""partner-pub-2905051069986345:lub5li-izzy"">"
		Response.Write "<input type=""hidden"" name=""cof"" value=""FORID:10"">"
		Response.Write "<input type=""hidden"" name=""ie"" value=""Shift_JIS"">"
		Response.Write "<input type=""text"" name=""q"" size=""20"">"
		Response.Write "</label>"
		Response.Write "<input type=""submit"" name=""sa"" value=""&#x691c;&#x7d22;"">"
		Response.Write "</div>"
		Response.Write "</form>"
		Response.Write "<script type=""text/javascript"" src=""http://www.google.co.jp/coop/cse/brand?form=cse-search-box&amp;lang=ja""></script>"
	End If
	'</Google�̃T�C�g������>

	Response.Write "�@<a href=""" & sLinkurl & """ title=""" & sLinkalt & """ style=""font-size:14px;""><img src=""/img/top/head_icon.gif"" alt=""" & sLinkalt & """ border=""0"" style=""vertical-align:middle;"">" & sLinktext & "</a>"
	'<!-- #INCLUDE FILE="../ad_banner_control/ad_banner.asp" -->
	Response.Write "</td>"
	'<�w�b�_�[�E>

	Response.Write "</tr>"

	'<�w�b�_�[�����F�w�i�΂̂��>
	Response.Write "<tr style=""background-image:url(/img/top/headtext_background.gif);"">"
	Response.Write "<td colspan=""2"" align=""left"" style=""margin:0px;padding:0px;color:#ffffff; height:20px;border-top:solid 1px #ffffff; border-bottom:solid 1px #ffffff;"">"
	Response.Write sHeadcmt
	Response.Write "</td>"
	Response.Write "</tr>"
	'</�w�b�_�[�����F�w�i�΂̂��>

	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
	Response.Write "<div align=""left"" style=""width:790px;float:left;"">" '�y�[�W�S�̂̕��ifooter�ŉ����ŕ�
	Response.Write "<div class=""moji912"" style=""padding-left:3px;width:615px;float:right"">" & vbCrLf '���C���R���e���c���isidemenu�ŏ㕔�ŕ߁j
End Function

'******************************************************************************
'�T�@�v�F�T�C�h���j���[
'���@���FSidemenuType	0�y�g�b�v�z1�y���E�ҁz2�y��Ɓz3�y���p�z
'�쐬�ҁFLis Niina
'�쐬���F2008/02/07
'���@�l�F
'�g�p���F
'���@���F
' 08/05/20 Lis�� �����ƃi�rFC�ǉ�
'******************************************************************************
Function NaviSidemenu(SidemenuType)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	If SidemenuType = 0 Then '�g�b�v�y�[�W
		Response.Write "</div>"'���C���R���e���c�̕��w��div�̕߁i�J�n��header�ŉ����j
		Response.Write "<div style=""width:170px; float:left; margin:0px;padding:0px;"">"

		'���g�b�v�y�[�W�����O�C�����ɂ���ăT�C�h���j���[��؂�ւ���
		If session("usertype") = "staff" Then '���E�҃��O�C�����Ă���ꍇ
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_staff_big"">My Menu �i<a title=""���O�A�E�g"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">���O�A�E�g����</a>�j</li>"
			Response.Write "<li class=""sidemenu_mypage""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_job""><a title=""�W���u�E�R���V�F���W��"" href=""" & HTTP_CURRENTURL & "staff/jobcon/"">�W���u�E�R���V�F���W��</a></li>"
			Response.Write "<li class=""sidemenu_job""><a title=""���d������"" href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">���d������</a></li>"
			Response.Write "<li class=""sidemenu_mail""><a title=""���[���Ǘ�"" href=""" & HTTPS_CURRENTURL & "staff/mailhistory_person.asp"">���[���Ǘ�"

			sSQL = "SELECT COUNT(*) AS Cnt FROM MailHistory WITH(NOLOCK) WHERE ReceiverCode ='" & Session("userid") & "' AND OpenDay IS NULL AND ReceiverDelFlag = '0'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If oRS.Collect("Cnt") = 0 Then
					Response.Write "(<img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">����" & oRS.Collect("Cnt") & "��)"
				Else
					Response.Write "(<span style=""color:#ff0000; font-weight:bold;""><img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">����" & oRS.Collect("Cnt") & "��</span>)"
				End If
			End If

			Response.Write "</a></li>"
			Response.Write "<li class=""sidemenu_detail""><a title=""�o�^���e�C��"" href=""" & HTTPS_CURRENTURL & "staff/person_detail.asp"">�o�^���e�C��</a></li>"
			Response.Write "<li class=""sidemenu_print""><a title=""�������E�E���o�����@���"" href=""" & HTTP_CURRENTURL & "staff/resume_print.asp"">�������E�E���o�����@�o��</a></li>"
			Response.Write "<li class=""sidemenu_wacth""><a title=""�E�H�b�`���X�g"" href=""" & HTTP_CURRENTURL & "staff/watchlist.asp"">�E�H�b�`���X�g</a></li>"
			Response.Write "<li class=""sidemenu_picture""><a title=""�������ʐ^�o�^"" href=""" & HTTP_CURRENTURL & "staff/resume_picture.asp"">�������ʐ^�o�^</a></li>"
			Response.Write "<li class=""sidemenu_pass""><a title=""�p�X���[�h�̕ύX"" href=""" & HTTPS_CURRENTURL & "staff/changepassword.asp"">�p�X���[�h�ύX</a></li>"
			Response.Write "<li class=""sidemenu_staff_bottom""></li>"
			Response.Write "</ul>"
		ElseIf ( Session("usertype") = "company" Or Session("usertype") = "dispatch") And G_USEFLAG <> "0" Then '��ƃ��O�C�����Ă���ꍇ
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">My Menu �i<a title=""���O�A�E�g"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">���O�A�E�g����</a>�j</li>"
			Response.Write "<li class=""sidemenu_company""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���E�҂̌���"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">���E�҂̌���</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""�E�H�b�`���X�g"" href=""" & HTTP_CURRENTURL & "company/watchlist.asp"">�E�H�b�`���X�g</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���[������"" href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">���[������</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���l�[�̏C��"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">���l�[�̏C��</a></li>"

			If Session("usertype") = "company" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "company/company_reg1.asp"">���Џ����X�V</a></li>"
				If G_IMAGELIMIT > 0 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">��Ǝʐ^�摜�f��</a></li>"
				End If

				If G_IMAGELIMIT > 1 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">���l�[�p�摜�X�g�b�N</a></li>"
				End If
			ElseIf Session("usertype") = "dispatch" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "dispatch/company_reg1.asp"">���Џ����X�V</a></li>"
				If G_IMAGELIMIT > 0 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">��Ǝʐ^�摜�f��</a></li>"
				End If

				If G_IMAGELIMIT > 1 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">���l�[�p�摜�X�g�b�N</a></li>"
				End If
			End If

			Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "mailtemplate/manager.asp"">���[���e���v���[�g�Ǘ�</a></li>"
			If G_PLANTYPE <> "mail" then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "company/costperformance/"">�̗p���P��߰ļ���<img src=""/img/new.gif"" border=""0""></a></li>"
			End If

			'Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "license/license_manager.asp"">���C�Z���X�Ǘ�</a></li>"
			Response.Write "<li class=""sidemenu_company_end""><a href=""" & HTTPS_CURRENTURL & "company/changepassword.asp"">�p�X���[�h�ύX</a></li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
			Response.Write "</ul>"
		ElseIf (Session("usertype") = "company" Or Session("usertype") = "dispatch") And G_USEFLAG = "0"  Then '��ƃ��O�C�����Ă��邪���C�Z���X���؂�Ă���ꍇ
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">My Menu</li>"
			Response.Write "<li class=""sidemenu_company""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page �i<a title=""���O�A�E�g"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">���O�A�E�g����</a>�j</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���E�҂̌���"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">���E�҂̌���</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""�E�H�b�`���X�g"" href=""" & HTTP_CURRENTURL & "company/watchlist.asp"">�E�H�b�`���X�g</a></li>"

			If G_MAILREADFLAG = "1" Then
				Response.Write "<li class=""sidemenu_company""><a title=""���[������"" href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">���[������</a></li>"
			End If

			Response.Write "<li class=""sidemenu_company""><a title=""���Ћ��l�[�ꗗ"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">���Ћ��l�[�ꗗ</a></li>"
			'Response.Write "<li class=""sidemenu_company""><a title=""���C�Z���X�Ǘ�"" href=""" & HTTP_CURRENTURL & "license/license_manager.asp"">���C�Z���X�Ǘ�</a></li>"
			Response.Write "<li class=""sidemenu_company_end""><a title=""�p�X���[�h�ύX"" href=""" & HTTPS_CURRENTURL & "company/changepassword.asp"">�p�X���[�h�ύX</a></li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
			Response.Write "</ul>"

			'Response.Write "<ul>"
			'Response.Write "<li class=""sidemenu_big"">���O�C��<span style=""font-size:10px;"">�i���Ƀ��O�C���ς݂ł��j</span></li>"
			'Response.Write "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "login_menu.asp"" title=""�]�E���O�C��"">My Page��</a></li>"
			'Response.Write "<li class=""sidemenu_bottom""></li>"
			'Response.Write "</ul>"
			'���g�b�v�y�[�W�����O�C�����ɂ���ăT�C�h���j���[��؂�ւ���
		Else
			Response.Write "<div align=""center"">"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp"" title=""�]�E,�V�K����o�^""><img src=""/img/common/reg1_button.jpg"" border=""0"" alt=""�V�K����o�^"" style=""margin-top:3px;""></a>"
			Response.Write "<script type=""text/javascript""><!-- document.forms[0].UserID.focus(); // --></script>"
			Response.Write "</div>"
			Response.Write "<form id=""frmlogin"" method=""post"" action=""" & HTTPS_CURRENTURL & "login_check.asp"">"

			Dim sName
			If LCase(Request.QueryString("JUMP_URL_FLAG")) = "true" Then
				For Each sName In Request.QueryString
					Response.Write "<input type=""hidden"" name=""" & sName & """ value=""" & Request.QueryString(sName) & """>"
				Next
			End If

			Dim si
			si = GetForm("si","2")

			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_big"">���O�C��</li>"
			Response.Write "<li style=""border-right:solid 1px #cccccc; border-left:solid 1px #cccccc;"">"
			Response.Write "<div style=""font-size:11px; padding-top:0px; padding-right:3px;"">"
			Response.Write "<div align=""right"">"

			If G_SSLFLAG = False Then
				Response.Write "<a href=""" & HTTPS_CURRENTURL & """ style=""color:#0045f9;""><img src=""/img/common/security_key.gif"" border=""0"" height=""12"" alt="""">�r�r�k���n�m�ɂ��� ������</a><br>"
			Else
				Response.Write "<a href=""" & HTTP_CURRENTURL & """ style=""color:#0045f9;"">�r�r�k���n�e�e�ɂ���</a><br>"
			End If
			Response.Write "</div>"

			If si <> "" Then
				Response.Write "<p class=""m0"" style=""float:right;"">�@<input type=""text"" name=""CONF_UserID"" value=""" & si & """ style=""width:80px;""></p>"
			Else
				Response.Write "<p class=""m0"" style=""float:right;"">�@<input type=""text"" name=""CONF_UserID"" value=""" & Request.Cookies("id_memory") & """ style=""width:80px;""></p>"
			End If

			Response.Write "<p class=""m0"" style=""font-size:10px;color:#666666;float:right;""><b>I�@D</b></p>"
			Response.Write "<br clear=""all"">"
			Response.Write "<p class=""m0"" style=""float:right;"">�@<input type=""password"" name=""CONF_Password"" value="""" style=""width:80px;""></p>"
			Response.Write "<p class=""m0"" style=""font-size:10px;color:#666666;float:right;""><b>�p�X���[�h</b></p>"
			Response.Write "<br clear=""all"">"
			Response.Write "<div align=""right"">"
			Response.Write "<label><input type=""checkbox"" name=""frmautologinflag"" value=""1"">����۸޲�</label>[<span style=""color:#0045f9; cursor:pointer;"" onclick=""window.open('/infomation/autologin.asp','autologin','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=220');""><u>�H</u></span>]"
			Response.Write "<input type=""submit"" value=""���O�C��"" onclick=""DataCheckIdreg(); return false""><br>"
			Response.Write "<a href=""staff/qa.asp#003"" style=""font-size:10px;"" title=""�]�E,���O�C���ł��Ȃ���"">۸޲݂ł��Ȃ�</a>�@"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "staff/passwordreminder.asp"" style=""font-size:10px;"" title=""�]�E,�p�X���[�h��Y�ꂽ��"">ID�E�߽ܰ�ނ�Y�ꂽ</a><br>"
			Response.Write "</div>"
			Response.Write "</div>"
			Response.Write "</li>"
			Response.Write "<li class=""sidemenu_bottom""></li>"
			Response.Write "</ul>"
%><!-- #INCLUDE FILE="../error/errHandle.asp" --><%
			Response.Write "</form>"
		End If

		'�g�b�v�y�[�W��
		'Response.Write "<div style=""width:170px;height:50px;margin-bottom:5px;"">"
		'Response.Write "<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=J0051817"" title=""SOHO�L���㗝�X"">"
		'Response.Write "<img src=""/img/top/soho_banner.gif"" alt=""SOHO�L���㗝�X"" border=""0""><br>"
		'Response.Write "</a>"
		'Response.Write "</div>"

		Response.Write "<div style=""width:170px;height:50px;margin-bottom:5px;"">"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "staff/jobcon/introduction.asp"" title=""�W���u�E�R���V�F���W��"">"
		Response.Write "<img src=""/img/staff/jobcon/top_mini_banner.gif"" alt=""�]�E�x���W���u�E�R���V�����W��"" border=""0""><br>"
		Response.Write "</a>"
		Response.Write "</div>"
		Response.Write "<div style=""width:170px;height:135px;background-image:url(/img/sidemenu/navicafe_banner_all.jpg);margin-bottom:5px;"">"
		Response.Write "<a href=""/cafe/cafe_list.asp"" title=""�i�r�J�t�F""><img src=""/img/sidemenu/navicafe_banner_top.jpg"" alt=""�i�r�J�t�F"" border=""0"" style=""margin:0px;padding:0px;""></a>"
		Response.Write "<div style=""margin-top:14px;padding:0px 6px 0px 8px;font-size:10px;line-height:15px;"">"

		'** TOP 08/11/05 Lis�� ADD
		'���݌f�ڒ���TOP3�̃g�s
		sSQL = "EXEC up_GetData_NC_Topic '','','','1','3';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			Response.Write "<a href='/cafe/cafe_detail.asp?t=" & oRS.Collect("TopicID")
			Response.Write "' title='" & oRS.Collect("Title") & "'>�E"
			If Len(oRS.Collect("Title")) > 14 Then
				Response.Write Left(oRS.Collect("Title"),14) & "..."
			Else
				Response.Write oRS.Collect("Title")
			End If
			Response.Write "</a><br>"
			oRS.MoveNext
		Loop
		Call RSClose(oRS)
		'** BTM 08/11/05 Lis�� ADD

		Response.Write "</div>"
		Response.Write "</div>"

		Response.Write "<div style=""width:170px;height:135px;background-image:url(/img/sidemenu/warmreception_banner_all.jpg);margin-bottom:5px;"">"
		Response.Write "<a href=""/s_contents/warmreception/"" title=""�����ƃi�r����D��""><img src=""/img/sidemenu/warmreception_banner_top.jpg"" alt=""�����ƃi�r����D��"" border=""0""></a>"
		Response.Write "<div style=""margin-top:14px;padding:0px 6px 0px 8px;font-size:10px;line-height:15px;"">"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "s_contents/warmreception/detail.asp?category=license&id=0102"">�E�r�W�l�X�E�L�����A���莎���i�Q���j</a><br>"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "s_contents/warmreception/detail.asp?category=license&id=0103"">�E�r�W�l�X�E�L�����A���莎���i�R���j</a><br>"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "s_contents/warmreception/detail.asp?category=skillup&id=0101"">�E��w�X�L�� �p��b�i�}���c�[�}���j</a><br>"
		Response.Write "</div>"
		Response.Write "</div>"

		Response.Write "<ul>"
		Response.Write "<li class=""sidemenu_big"">�֗��c�[��</li>"
		Response.Write "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; border-bottom:solid 1px #dddddd; line-height:17px;""><a href=""/staff/s_resume.asp"" title=""�������̎����쐬"" style=""display:block; background-image:url(/img/sidemenu/resume_banner.jpg); width:154px; height:73px; font-size:10px; padding:54px 0px 0px 14px; color:#444444; text-decoration:none;"">�K�v�ȍ��ڂ���͂��邾���Ŋ����I<br>" & Left(sAll,2) & "���l���g�����S�̃T�[�r�X�I<br>�����ɍ�����������������I</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_resume_kakikata.asp"" title=""�������̏�����"">�������̏�����</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_careersheet.asp"" title=""�E���o�����̎����쐬"">�E���o�����̎����쐬</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_careersheet_kakikata_1.asp"" title=""�E���o�����̏�����"">�E���o�����̏�����</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/s_contents/motive_index.asp"" title=""�u�]���@���[�J�["">�u�]���@���[�J�[</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/s_contents/s_jikopr.asp"" title=""����PR���[�J�["">����PR���[�J�[</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""/s_contents/s_taishokunegai.asp"" title=""�ސE��̏�����"">�ސE��̏�����</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"
		Response.Write "</ul>"

		Response.Write "<ul>"
		Response.Write "<li class=""sidemenu_big"">�T�|�[�g</li>"
		Response.Write "<li class=""sidemenu""><a href=""/s_contents/navistep_index.asp"" title=""���߂Ă̓]�E����"">���߂Ă̓]�E����</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_aboutnavi.asp"" title=""�����p�K�C�h"">�����p�K�C�h</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/qa.asp"" title=""�p���`"">�p���`</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_searchexplanation.asp"" title=""���d���������@"">���d���������@</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_kiyaku.asp"" title=""���p�K��"">���p�K��</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/shigotonavi/sitemap.asp"" title=""�T�C�g�}�b�v"">�T�C�g�}�b�v</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/link.asp"" title=""�����N�|���V�["">�����N�|���V�[</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/link_collection.asp"" title=""���𗧂����I�����N�W"">���𗧂����I�����N�W</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/s_contents/s_books.asp"" title=""�]�E�ɖ𗧂{"">�]�E�ɖ𗧂{</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/company/index.asp"" title=""��ƌ����̗p�R���e���c"">��ƌ������l�L���ɂ���</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/lis/lis.asp"" title=""�^�c���"">�^�c���</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "staff/access.asp"" title=""���₢���킹"">���₢���킹</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"
		Response.Write "</ul>"

		Response.Write "<div align=""center"" style=""width:100%;"">"
		Response.Write "<div class=""sidemenu_big"" style=""text-align:left;"">���E�ҏ��</div>"
		Response.Write "<div style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; background-image:url(/img/sidemenu/jinzaidata_background.gif);"" align=""center"">"
		Response.Write "<table style=""width:155px; font-size:10px; text-align:left;"">"

		Dim rank(2)
		Dim rankcount(2)
		Dim idx
		idx = 0

		sSQL = "SELECT top 3 Subitem,Number FROM Person_Statistics where item = '�s���{����' order by convert(int,Number) desc"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			rank(idx) = Replace(Replace(Replace(oRS.Collect("SubItem"),"�s",""),"�{",""),"��","")
			rankcount(idx) = oRS.Collect("Number")
			idx = idx + 1
			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		Response.Write "<tr>"
		Response.Write "<td>�s���{����</td>"
		Response.Write "<td>1��:" & rank(0) & "</td>"
		Response.Write "<td align=""right"">" & rankcount(0) & "��</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td></td>"
		Response.Write "<td>2��:" & rank(1) & "</td>"
		Response.Write "<td align=""right"">" & rankcount(1) & "��</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td></td>"
		Response.Write "<td>3��:" & rank(2) & "</td>"
		Response.Write "<td align=""right"">" & rankcount(2) & "��</td>"
		Response.Write "</tr>"

		idx = 0
		sSQL = "SELECT top 3 item,subitem, Number FROM Person_Statistics where item = '10�Α�' or item = '20�Α�' or item = '30�Α�' or item = '40�Α�' or item = '50�Α�' or item = '60�Έȏ�' order by convert(int,Number) desc"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			rank(idx) = Replace(oRS.Fields("Item").Value,"��","") & oRS.Fields("SubItem").Value
			rankcount(idx) = oRS.Fields("Number").Value
			idx = idx + 1
			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		Response.Write "<tr>"
		Response.Write "<td>�N���</td>"
		Response.Write "<td>1��:" & rank(0) & "</td>"
		Response.Write "<td align=""right"">" & rankcount(0) & "��</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td></td>"
		Response.Write "<td>2��:" & rank(1) & "</td>"
		Response.Write "<td align=""right"">" & rankcount(1) & "��</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td></td>"
		Response.Write "<td>3��:" & rank(2) & "</td>"
		Response.Write "<td align=""right"">" & rankcount(2) & "��</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td colspan=""3"" align=""right""><a href=""/company/c_staffdata.asp""><img src=""/img/sidemenu/kuwashiku_min.jpg"" alt=""�ڂ����͂�����"" border=""0""></a>"
		Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "<div class=""sidemenu_bottom"" style=""clear:both;""></div>"
		Response.Write "<br>"
		Response.Write "</div>"

		Response.Write "<div align=""center"" style=""width:100%;"">"
		Response.Write "<a href=""/lis/blog_kimura.asp"">"
		Response.Write "<img src=""/img/top/top_blogBanner.gif"" border=""0"" alt=""�ؑ����Y�̃q�g�r�W�l�X��Â�"">"
		Response.Write "</a>"
		Response.Write "</div>"

		'Response.Write "<div style=""text-align:center; font-size:11px;width:100%;padding:0px 15px;"">"
		'Response.Write "<img src=""/img/spacer.gif"" width=""3"" height=""10"" alt=""�]�E""><br>"
		'Response.Write "<div style=""float:left;"">"
		'Response.Write "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""�v���C�o�V�[�}�[�N"" border=""0"" width=""45""></a><br><a href=""/privacy/privacy.asp"">�l���ی�</a></div><div>"
		'Response.Write "<a href=""https://secure.secom.ne.jp/webp/db/1116062419.html"" target=""_blank""><img src=""img/secom/B0474507/B0474507_s.gif"" border=""0"" alt="""" height=""43""><br>SSL�Í���<br></a>"
		'Response.Write "</div>"
		'Response.Write "</div>"

		Response.Write "<div style=""text-align:center""></div>"
		Response.Write "</div>"
	ElseIf SidemenuType = 1 Then '���E��
		Response.Write "</div>" '���C���R���e���c�̕��w��div�̕߁i�J�n��header�ŉ����j
		Response.Write "<div id=""idNavigation"" style=""width:170px; float:left;"">"

		Response.Write "<!-- MENU START -->"
		'�������O�C�����̋��E�ҍ���
		If Session("usertype") = "staff" Then
			Response.Write "<div style=""clear:both; margin-bottom:5px;""></div>"
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_staff_big"">My Menu �i<a title=""���O�A�E�g"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">���O�A�E�g����</a>�j</li>"
			Response.Write "<li class=""sidemenu_mypage""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_job""><a title=""�W���u�E�R���V�F���W��"" href=""" & HTTP_CURRENTURL & "staff/jobcon/"">�W���u�E�R���V�F���W��</a></li>"
			Response.Write "<li class=""sidemenu_job""><a title=""���d������"" href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">���d������</a></li>"
			Response.Write "<li class=""sidemenu_mail""><a title=""���[���Ǘ�"" href=""" & HTTPS_CURRENTURL & "staff/mailhistory_person.asp"">���[���Ǘ�"

			sSQL = "SELECT COUNT(*) AS Cnt FROM MailHistory WITH(NOLOCK) WHERE ReceiverCode ='" & Session("userid") & "' AND OpenDay IS NULL AND ReceiverDelFlag = '0'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

			If GetRSState(oRS) = True Then
				If oRS.Collect("Cnt") = 0 Then
					Response.Write "(<img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">����" & oRS.Collect("Cnt") & "��)"
				Else
					Response.Write "(<span style=""color:#ff0000; font-weight:bold;""><img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">����" & oRS.Collect("Cnt") & "��</span>)"
				End If
			End If

			Response.Write "</a></li>"
			Response.Write "<li class=""sidemenu_detail""><a title=""�o�^���e�C��"" href=""" & HTTPS_CURRENTURL & "staff/person_detail.asp"">�o�^���e�C��</a></li>"
			Response.Write "<li class=""sidemenu_print""><a title=""�������E�E���o�����@���"" href=""" & HTTP_CURRENTURL & "staff/resume_print.asp"">�������E�E���o�����@�o��</a></li>"
			Response.Write "<li class=""sidemenu_wacth""><a title=""�E�H�b�`���X�g"" href=""" & HTTP_CURRENTURL & "staff/watchlist.asp"">�E�H�b�`���X�g</a></li>"
			Response.Write "<li class=""sidemenu_footprint""><a title=""�C�ɂȃ��X�g"" href=""" & HTTP_CURRENTURL & "staff/footprint.asp"">�C�ɂȃ��X�g</a></li>"
			Response.Write "<li class=""sidemenu_picture""><a title=""�������ʐ^�o�^"" href=""" & HTTP_CURRENTURL & "staff/resume_picture.asp"">�������ʐ^�o�^</a></li>"
			Response.Write "<li class=""sidemenu_pass""><a title=""�p�X���[�h�̕ύX"" href=""" & HTTPS_CURRENTURL & "staff/changepassword.asp"">�p�X���[�h�ύX</a></li>"
			Response.Write "<li class=""sidemenu_staff_bottom""></li>"
		Else
			'�������O�C�����Ă��Ȃ����̋��E�ҍ���
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp""><img src=""/img/common/reg1_button.jpg"" alt=""�����ƃi�r����o�^"" border=""0"" style=""margin:3px 0px 2px 0px;""></a><br>"
			Response.Write "<div align=""right"" style=""font-size:11px; margin-bottom:5px;"">"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "login_menu.asp"">����o�^�����ς݂̕��͂�����</a>"
			Response.Write "</div>"
			Response.Write "<a title=""���d������"" href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp""><img src=""/img/sidemenu/jobsearch_button.jpg"" alt=""���d������"" border=""0"" style=""margin-top:3px;""></a>"
			Response.Write "<div style=""width:170px;height:50px;margin:5px 0px;"">"
			Response.Write "<a href=""" & HTTP_CURRENTURL & "staff/jobcon/introduction.asp"" title=""�W���u�E�R���V�F���W��"">"
			Response.Write "<img src=""/img/staff/jobcon/top_mini_banner.gif"" alt=""�]�E�x���W���u�E�R���V�����W��"" border=""0""><br>"
			Response.Write "</a>"
			Response.Write "</div>"

			Response.Write "<ul>"
		End If

		'�������E�ҍ�������
		Response.Write "<li class=""sidemenu_big"">�R�~���j�e�B</li>"
		Response.Write "<li class=""sidemenu""><a title=""�i�r�J�t�F"" href=""" & HTTP_CURRENTURL & "cafe/cafe_list.asp"">�i�r�J�t�F</a></li>"
		Response.Write "<li class=""sidemenu_end""><a title=""�����ƃi�r�A���P�[�g"" href=""" & HTTP_CURRENTURL & "s_contents/enquete.asp"">�����ƃi�r�A���P�[�g</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		Response.Write "<li class=""sidemenu_big"">���ލ쐬�x��</li>"
		Response.Write "<li class=""sidemenu""><a title=""�������̎����쐬"" href=""" & HTTP_CURRENTURL & "staff/s_resume.asp"">�������̎����쐬</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�������̏�����"" href=""" & HTTP_CURRENTURL & "staff/s_resume_kakikata.asp"">�������̏�����</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�������p���`"" href=""" & HTTP_CURRENTURL & "staff/s_resume_qa.asp"">�������p���`</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�E���o�����̎����쐬"" href=""" & HTTP_CURRENTURL & "staff/s_careersheet.asp"">�E���o�����̎����쐬</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�E���o�����̏�����"" href=""" & HTTP_CURRENTURL & "staff/s_careersheet_kakikata_1.asp"">�E���o�����̏�����</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�u�]���@���[�J�["" href=""" & HTTP_CURRENTURL & "s_contents/motive_index.asp"">�u�]���@���[�J�[</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""����PR���[�J�["" href=""" & HTTP_CURRENTURL & "s_contents/s_jikopr.asp"">����PR���[�J�[</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�ސE��̏�����"" href=""" & HTTP_CURRENTURL & "s_contents/s_taishokunegai.asp"">�ސE��̏�����</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�w���v�Z�E����a����\"" href=""" & HTTP_CURRENTURL & "s_contents/s_year_calculation.asp"">�w���v�Z�E����a����\</a></li>"
		Response.Write "<li class=""sidemenu_end""><a title=""Conpri - �R���v��"" href=""" & HTTP_CURRENTURL & "conpri/"">�R���r�j���</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		Response.Write "<li class=""sidemenu_big"">�]�E�x���c�[��</li>"
		Response.Write "<li class=""sidemenu""><a title=""���߂Ă̓]�E����"" href=""" & HTTP_CURRENTURL & "s_contents/navistep_index.asp"">���߂Ă̓]�E����</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�����ƃi�r�]�E�R����"" href=""" & HTTP_CURRENTURL & "column/column_index.asp"">�����ƃi�r�]�E�R����</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�]�E�̐S�\��"" href=""" & HTTP_CURRENTURL & "s_contents/s_ready.asp"">�]�E�̐S�\��</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�]�E�ɕK�v�Ȏ葱��"" href=""" & HTTP_CURRENTURL & "s_contents/s_proce.asp"">�]�E�ɕK�v�Ȏ葱��</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""���i��UP�]�E�}�j���A��"" href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"">���i��UP�]�E�}�j���A��</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�j�[�g����̒E�o"" href=""" & HTTP_CURRENTURL & "s_contents/s_neet.asp"">�j�[�g����̒E�o</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�Љ�\��h���Ƃ�"" href=""" & HTTP_CURRENTURL & "s_contents/s_temptoperm.asp"">�Љ�\��h���Ƃ�</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""���Ȃ��̋��^����"" href=""" & HTTP_CURRENTURL & "s_contents/s_kyuuyomeisai.asp"">���Ȃ��̋��^����</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�K�E�f�f����Ԃ�i�r�"" href=""" & HTTP_CURRENTURL & "s_contents/s_mynavi.asp"">�K�E�f�f����Ԃ�i�r�</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""�X�J�E�g���[������������󂯂�ɂ́I�H"" href=""" & HTTP_CURRENTURL & "s_contents/labo/scoutlabo.asp"
		If G_USERTYPE = "staff" Then Response.Write "?staffcode=" & G_USERID & "&amp;linkno=3"
		Response.Write """>�X�J�E�g���{</a></li>"

		If G_USERTYPE = "staff" Then
			Response.Write "<li class=""sidemenu_end""><a title=""�����ƃi�r����D��"" href=""" & HTTP_CURRENTURL & "s_contents/warmreception/"">�i�r����D��</a></li>"
		End If

		Response.Write "<li class=""sidemenu_bottom""></li>"


		Response.Write "</ul>"
		Response.Write "<br clear=""all"">"
		Response.Write "<div align=""center"" style=""clear:both; margin-top:20px;"">"
		Response.Write "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""�v���C�o�V�[�}�[�N"" border=""0""></a><br>"
		Response.Write "<a href=""/privacy/privacy.asp"">�l���ی�</a>"
		Response.Write "</div>"
		Response.Write "<div style=""text-align:center""></div>"
		Response.Write "<!-- MENU END -->"
		Response.Write "</div>"
	ElseIf SidemenuType = 2 Then '���
		Response.Write "<script type=""text/javascript"">"
		Response.Write "function LoginCheckIdreg(){"
		Response.Write "var ofrm = document.forms.frmlogin;"
		Response.Write "if(!navigator.cookieEnabled) {"
		Response.Write "alert('cookie�i�N�b�L�[�j�̗��p���ł��Ȃ��ݒ�ɂȂ��Ă��܂��B\n�u���E�U��Z�L�����e�B�[�\�t�g��cookie�ݒ�����m�F�������B');"
		Response.Write "return false;"
		Response.Write "}"
		Response.Write "if(!ChkInput(ofrm.CONF_UserID, 'string', '1', '�F��ID����͂��Ă��������B')) return false;"
		Response.Write "if(!ChkInput(ofrm.CONF_Password,'string', '1', '�p�X���[�h����͂��Ă��������B')) return false;"
		Response.Write "if(!ChkLength(ofrm.CONF_Password, 3, 20, '�p�X���[�h�͂R�����ȏ�A�Q�O�����ȉ��œ��͂��Ă��������B'))return false;"
		Response.Write "ofrm.submit();"
		Response.Write "}"
		Response.Write "</script>"
		Response.Write "</div>"'���C���R���e���c�̕��w��div�̕߁i�J�n��header�ŉ����j

		Response.Write "<div id=""idNavigation"" style=""width: 170px; float: left;"">"
		If G_USEFLAG = "0" Then
			'���C�Z���X�؂�̊��
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">My Menu �i<a title=""���O�A�E�g"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">���O�A�E�g����</a>�j</li>"
			Response.Write "<li class=""sidemenu_company""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���E�҂̌���"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">���E�҂̌���</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""�E�H�b�`���X�g"" href=""" & HTTP_CURRENTURL & "company/watchlist.asp"">�E�H�b�`���X�g</a></li>"
			If G_MAILREADFLAG = "1" Then
				Response.Write "<li class=""sidemenu_company""><a title=""���[������"" href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">���[������</a></li>"
			End If
			Response.Write "<li class=""sidemenu_company""><a title=""���Ћ��l�[�ꗗ"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">���Ћ��l�[�ꗗ</a></li>"
			'Response.Write "<li class=""sidemenu_company""><a title=""���C�Z���X�Ǘ�"" href=""" & HTTP_CURRENTURL & "license/license_manager.asp"">���C�Z���X�Ǘ�</a></li>"
			Response.Write "<li class=""sidemenu_company_end""><a title=""�p�X���[�h�ύX"" href=""" & HTTPS_CURRENTURL & "company/changepassword.asp"">�p�X���[�h�ύX</a></li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
		ElseIf Session("usertype") = "company" Or Session("usertype") = "dispatch" Then
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">My Menu �i<a title=""���O�A�E�g"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">���O�A�E�g����</a>�j</li>"
			Response.Write "<li class=""sidemenu_company""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���E�҂̌���"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">���E�҂̌���</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���E�Ҍ��������Ǘ�"" href=""" & HTTP_CURRENTURL & "company/searchstaffcondition/list.asp"">���E�Ҍ��������Ǘ�</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""�E�H�b�`���X�g"" href=""" & HTTP_CURRENTURL & "company/watchlist.asp"">�E�H�b�`���X�g</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""�C�ɂȃ��X�g"" href=""" & HTTP_CURRENTURL & "company/report/footprint.asp"">�C�ɂȃ��X�g</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���[������"" href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">���[������</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""�ꊇ���[���Ǘ�"" href=""" & HTTPS_CURRENTURL & "company/lumpmail/list.asp"">�ꊇ���[���Ǘ�</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""���l�[�̏C��"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">���l�[�̏C��</a></li>"

			If Session("usertype") = "company" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "company/company_reg1.asp"">���Џ����X�V</a></li>"
				If G_IMAGELIMIT > 0 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">��Ǝʐ^�摜�f��</a></li>"
				End If

				If G_IMAGELIMIT > 1 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">���l�[�p�摜�X�g�b�N</a></li>"
				End If
			ElseIf Session("usertype") = "dispatch" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "dispatch/company_reg1.asp"">���Џ����X�V</a></li>"
				If G_IMAGELIMIT > 0 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">��Ǝʐ^�摜�f��</a></li>"
				End If

				If G_IMAGELIMIT > 1 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">���l�[�p�摜�X�g�b�N</a></li>"
				End If
			End If

			Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "mailtemplate/manager.asp"">���[���e���v���[�g�Ǘ�</a></li>"
			'Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "license/license_manager.asp"">���C�Z���X�Ǘ�</a></li>
			If G_PLANTYPE = "mail" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/point/"">�|�C���g�Ǘ�</a></li>"
			End If

			If G_PLANTYPE <> "mail" then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "company/costperformance/"">�̗p���P��߰ļ���<img src=""/img/new.gif"" border=""0""></a></li>"
			End If

			Response.Write "<li class=""sidemenu_company_end""><a href=""" & HTTPS_CURRENTURL & "company/changepassword.asp"">�p�X���[�h�ύX</a></li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
		Else

			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">���O�C��</li>"
			Response.Write "<li>"
			Response.Write "<form id=""frmlogin"" method=""post"" action=""" & HTTPS_CURRENTURL & "login_check.asp"">"
			Response.Write "<div style=""line-height:22px; color:#6666cc; font-size:11px; border-right:solid 1px #9999ff; border-left:solid 1px #9999ff; padding-right:3px;"" align=""right"">"

			If G_SSLFLAG = False Then
				Response.Write "<a href=""" & G_URLS & """ style=""color:#0045f9;""><img src=""/img/common/security_key.gif"" border=""0"" height=""12"" alt="""">�r�r�k���n�m�ɂ��� (����)</a><br>"
			Else
				Response.Write "<a href=""" & G_URL & """ style=""color:#0045f9;"">�r�r�k���n�e�e�ɂ���</a><br>"
			End If

			Response.Write "<font size=""1"" style=""width:50px; text-align:right; font-weight:bold;"">I�@D</font>"
			Response.Write "<input type=""text"" name=""CONF_UserID"" value=""" & Request.Cookies("id_memory") & """ style=""width:100px;""><br>"
			Response.Write "<font size=""1"" style=""width:50px; text-align:right; font-weight:bold;"">�p�X���[�h</font>"
			Response.Write "<input type=""password"" name=""CONF_Password"" size=""11"" value="""" style=""width:100px;""><br>"
			Response.Write "<div style=""text-align:right;"">"

			If Request.QueryString("JUMP_URL_FLAG") = "True" Then
				For Each name In Request.QueryString
					Response.Write "<input type=""hidden"" name=""" & name & """ value=""" & Request.QueryString(name) & """>"
				Next
			End If

			Response.Write "<label><input type=""checkbox"" name=""frmautologinflag"" value=""1"">����۸޲�</label>[<span style=""color:#0045f9; cursor:pointer;"" onclick=""window.open('/infomation/autologin.asp','autologin','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=220');""><u>�H</u></span>]"
			Response.Write "<input type=""Submit"" name=""Login"" value=""���O�C��"" onclick=""LoginCheckIdreg(); return false"" style=""font-size:12px; margin-right:1px;""><br>"
			Response.Write "</div>"
			Response.Write "</div>"
			Response.Write "<script type=""text/javascript""><!-- document.forms[0].UserID.focus(); // --></script>"
%><!-- #INCLUDE FILE="../error/errhandle.asp" --><%
			Response.Write "</form>"
			Response.Write "</li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
		End If

		Response.Write "<li class=""sidemenu_big"">���l�L��</li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/index.asp"" title=""�����ƃi�r���l�L���Ƃ�"">�����ƃi�r���l�L��</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_function.asp"" title=""�T�[�r�X�T�v"">�T�[�r�X�T�v</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_voice.asp"" title=""�����p��Ɨl�̐�"">�����p��Ɨl�̐�</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/research.asp"" title=""�l�ލ̗p���@�f�f"">�l�ލ̗p���@�f�f</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp"" title=""�l�ނc������"">�l�ނc������</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "jinzaisearch/index.asp"" title=""�l�ނ���������"">�l�ނ���������</a></li> "
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/charge.asp"" title=""�����V�X�e��"">�����V�X�e��</a></li>"

		If G_USERTYPE = "" Then
			Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/costperformance/"" title=""�̗p���P��߰ļ���"">�̗p���P��߰ļ���<img src=""/img/new.gif"" border=""0""></a></li>"
		End If

		'Response.Write "<li class=""sidemenu""><a href=""http://jinzai.shigotonavi.co.jp/joboffer/make_advertisement.asp"" target=""blank_"" title=""���l�L���쐬�ɂ���"">���l�L���쐬�ɂ���</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "company/request01.asp"" title=""���\������"">���\������</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		Response.Write "<li class=""sidemenu_big"">�l�ރT�[�r�X</li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"" title=""�l�ޏЉ�"">�l�ޏЉ�</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"" title=""�Љ�\��h��"">�Љ�\��h��</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"" title=""�l�ޔh��"">�l�ޔh��</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		'TOP 08/05/20 Lis�� �e�b ADD �� 08/09/04 Lis�� DEL
		'Response.Write "<ul class=""sidemenulink"">"
		'Response.Write "<li>&nbsp;<a href=""" & HTTP_CURRENTURL & "company/fc_index.asp"" title=""�l�ރT�[�r�X�t�����`���C�Y:�����ƃi�rFC"">�����ƃi�rFC</a></li>"
		'Response.Write "</ul><br>"
		'BTM 08/05/20 Lis�� �e�b ADD �� 08/09/04 Lis�� DEL

		Response.Write "<li class=""sidemenu_big"">�T�|�[�g</li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_successpoint.asp"" title=""�����ƃi�r���p�u�b�N"">�����ƃi�r���p�u�b�N</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_scout3point.asp"">�X�J�E�g���[���쐬�̃R�c</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/qa.asp"">�p���`</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_kiyaku.asp"">���p�K��</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "company/access.asp"">���⍇��</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		Response.Write "<li class=""sidemenu_big"">��ЊT�v</li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "lis/lis_annai.asp"" title=""��Јē�"">��Јē�</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "lis/service-development.asp"" title=""�����l�ރT�[�r�X�W�J"">�����l�ރT�[�r�X�W�J</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "lis/lis_saiyou.asp"" title=""�̗p���"">�̗p���</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"
		Response.Write "</ul>"

		Response.Write "<div align=""center"" style=""width:100%;padding-top:20px;"">"
		Response.Write "<a href=""/lis/blog_kimura.asp"">"
		Response.Write "<img src=""/img/top/top_blogBanner.gif"" border=""0"" alt=""�ؑ����Y�̃q�g�r�W�l�X��Â�"">"
		Response.Write "</a>"
		Response.Write "</div>"

		Response.Write "<div align=""center"" style=""margin-top:10px;"">"
		Response.Write "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""�v���C�o�V�[�}�[�N"" border=""0""></a><br>"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "privacy/privacy.asp"">�l���ی�ɂ���</a>"
		Response.Write "</div>"
		Response.Write "<div style=""text-align:center""></div>"

		Response.Write "</div>"
	ElseIf SidemenuType = 3 Then '���p
		If Session("usertype") = "staff" Then '���E�҃��O�C�����Ă���ꍇ
			Call NaviSidemenu(1)
		ElseIf Session("usertype") = "company" Or Session("usertype") = "dispatch" Then '��ƃ��O�C�����Ă���ꍇ
			Call NaviSidemenu(2)
		Else
			Response.Write "</div>" '���C���R���e���c�̕��w��div�̕߁i�J�n��header�ŉ����j
			Response.Write "<div id=""idNavigation"" style=""width: 170px; float: left;"">"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp""><img src=""/img/common/reg1_button.jpg"" alt=""�����ƃi�r����o�^"" border=""0"" style=""margin:3px 0px 2px 0px;""></a><br>"
			Response.Write "<div align=""right"" style=""font-size:11px; margin-bottom:5px;"">"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "login_menu.asp"">����o�^�����ς݂̕��͂�����</a>"
			Response.Write "</div>"

			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_big"">���d�������T���̕�</li>"
			Response.Write "<li class=""sidemenu""><a title=""���d������"" href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">���d������</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""�����p�K�C�h"" href=""" & HTTP_CURRENTURL & "staff/s_aboutnavi.asp"">�����p�K�C�h</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""�p���`"" href=""" & HTTP_CURRENTURL & "staff/qa.asp"">�p���`</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""���p�K��"" href=""" & HTTP_CURRENTURL & "staff/s_kiyaku.asp"">���p�K��</a></li>"
			Response.Write "<li class=""sidemenu_end""><a title=""���⍇��(���E�Ґ�p)"" href=""" & HTTPS_CURRENTURL & "staff/access.asp"">���⍇��(���E�Ґ�p)</a></li>"
			Response.Write "<li class=""sidemenu_bottom""></li>"
			Response.Write "</ul>"

			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_big"">�l�ނ����T���̊�Ɨl</li>"
			Response.Write "<li class=""sidemenu""><a title=""���O�C��"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">���O�C��</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""���l�L���ɂ���"" href=""" & HTTP_CURRENTURL & "company/c_hajime.asp"">���l�L���ɂ���</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""�l�ޏЉ�ɂ���"" href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"">�l�ޏЉ�ɂ���</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""�Љ�\��h���ɂ���"" href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"">�Љ�\��h���ɂ���</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""�l�ޔh���ɂ���"" href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"">�l�ޔh���ɂ���</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""���⍇��(���l��Ɨl��p)"" href=""" & HTTPS_CURRENTURL & "company/access.asp"">���⍇��(���l��Ɨl��p)</a></li>"
			Response.Write "<li class=""sidemenu_end""><a title=""��Јē�"" href=""" & HTTP_CURRENTURL & "lis/lis_annai.asp"">��Јē�</a></li>"
			Response.Write "<li class=""sidemenu_bottom""></li>"
			Response.Write "</ul>"

			Response.Write "<!-- SIDE-MENU END -->"
			Response.Write "<br>"
			Response.Write "<center>"
			Response.Write "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""�v���C�o�V�[�}�[�N"" border=""0""></a><br>"
			Response.Write "<a href=""" & HTTP_CURRENTURL & "privacy/privacy.asp"">�l���ی�ɂ���</a>"
			Response.Write "</center>"

			Response.Write "</div>"
		End If
	End If
End Function

'******************************************************************************
'�T�@�v�F�t�b�^�[
'���@���F
'�쐬�ҁFLis Niina
'�쐬���F2008/02/07
'���@�l�F
'�g�p���F
'���@���F2008/05/20 Lis�� �����ƃi�rFC�ǉ�
'******************************************************************************
Function NaviFooter()
	Response.Write "<div style=""clear:both;""></div>"
	Response.Write "</div>"
	If 1 = 2 Then
		Response.Write "<div style=""width:200px;float:right;margin-top:0px;"">"
		If Request.ServerVariables("URL") <> "/search.asp" Then
			Call NaviSidemenuRight()
		End If
		Response.Write "</div>"
	End If
	Response.Write "</div>"
	Response.Write "<br clear=""all"">"
%><!-- #INCLUDE VIRTUAL="/include/ads/navifooter.asp" --><%
	Response.Write "<br>"
	Response.Write "<div style=""text-align:left;height:55px; width:785px;"">"
	Response.Write "<ul class=""footer"" style=""float:left;padding-left:5px;"">"
	Response.Write "<li style=""float:left;""><a href=""" & HTTP_CURRENTURL & """ title=""�]�E�E���l�T�C�g�����ƃi�r"" class=""topdecnone"">�g�n�l�d</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "staff/Ranking.asp"" title=""���E�҃����L���O"" class=""topdecnone"">���E�҃����L���O</a></li>"
	'Response.Write "<li style=""float:left;"">�b<a href=""/company/c_hajime.asp"" class=""topdecnone"">���l�L��</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "infomation/info.asp"" title=""�L���q��"" class=""topdecnone"">�L���q��</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "lis/lis.asp"" title=""�^�c��ЁE���Ѝ̗p���"" class=""topdecnone"">�^�c��ЁE���Ѝ̗p���</a></li>"
	'Response.Write "<li style=""float:left;"">�b<a href=""/staff/s_aboutnavi.asp"" title=""�����p�K�C�h"" class=""topdecnone"">�����p�K�C�h</a></li>"
	'Response.Write "<li style=""float:left;"">�b<a href=""/staff/qa.asp"" title=""�p���`"" class=""topdecnone"">�p���`</a></li>"
	'Response.Write "<li style=""float:left;"">�b<a href=""/staff/s_kiyaku.asp"" title=""���p�K��"" class=""topdecnone"">���p�K��</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTPS_CURRENTURL & "staff/access.asp"" title=""���⍇��"" class=""topdecnone"">���⍇��&lt;�]�E��]�̕�����&gt;</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "s_contents/s_books.asp"" title=""�]�E�ɖ𗧂{"" style=""margin-left:5px;"" class=""topdecnone"">�]�E�ɖ𗧂{</a></li>"
	'Response.Write "<li style=""float:left;"">�b<a href=""/link.asp"" title=""�����N�|���V�["" class=""topdecnone"">�����N�|���V�[</a></li>"
	'Response.Write "<li style=""float:left;"">�b<a href=""/link_collection.asp"" title=""���𗧂����I�����N�W"" class=""topdecnone"">���𗧂����I�����N�W</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "shigotonavi/sitemap.asp"" class=""topdecnone"" title=""�T�C�g�}�b�v"">�T�C�g�}�b�v</a></li>"
	Response.Write "</ul>"
	Response.Write "<br clear=""all"">"

	Response.Write "<div style=""width:100%; height:5px; margin:0px; padding:0px; background-image:url(/img/footer/footer_1.gif); background-repeat:repeat-x;"">"
	Response.Write "</div>"
	Response.Write "<div style=""float:left;width:580px;padding-left:5px;"">"
	Response.Write "<ul class=""footer"">"

	Response.Write "<li style=""float:left;""><a href=""" & HTTP_CURRENTURL & "company/index.asp"" title=""��ƌ����R���e���c"" class=""topdecnone"">��ƌ����R���e���c</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "company/c_hajime.asp"" title=""���l�L��"" class=""topdecnone"">���l�L��</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"" title=""�l�ޔh��"" class=""topdecnone"">�l�ޔh��</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"" title=""�l�ޏЉ�"" class=""topdecnone"">�l�ޏЉ�</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"" title=""�Љ�\��h��"" class=""topdecnone"">�Љ�\��h��</a>�b</li>"
	'Response.Write "<li style="float:left;">�b<a href=""" & HTTPS_CURRENTURL & "company/fc_index.asp" title="�l�ރT�[�r�X�t�����`���C�Y,�����ƃi�rFC" class="topdecnone">�����ƃi�rFC</a>�b</li>"
	Response.Write "</ul><br>"
	Response.Write "<ul class=""footer"">"
	Response.Write "<li style=""float:left;""><a href=""" & HTTPS_CURRENTURL & "company/access.asp"" title=""���⍇��&lt;��Ɨl����&gt;"" class=""topdecnone"">���⍇��&lt;��Ɨl����&gt;</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp"" title=""�����ƃi�r���E�҂ƌf�ڊ�ƃf�[�^"" class=""topdecnone"">�����ƃi�r���E�҂ƌf�ڊ�ƃf�[�^</a></li>"
	Response.Write "<li style=""float:left;"">�b<a href=""" & HTTPS_CURRENTURL & "company/access.asp"" title=""���⍇��&lt;��Ɨl����&gt;"" class=""topdecnone"">�L���㗝�X�̕��̂��⍇��</a></li>"
	Response.Write "</ul>"
	Response.Write "</div>"
	Response.Write "<div style=""float:right;width:180px;"">"
	Response.Write "<a href=""" & HTTP_LIS_CURRENTURL & """ title=""�]�E�T�C�g�����ƃi�r�̉^�c���-���X�������-"" target=""_blank""><img src=""/img/footer/footer_lis_logo_1.gif"" alt=""�]�E�T�C�g������ƃi�r��^�c-���X�������-"" border=""0""></a>"
	Response.Write "</div>"
	Response.Write "<br clear=""all"">"
	Response.Write "</div>"
	Response.Write "</div>"
	Response.Write "</div>" & vbCrLf

	'�y�[�W�S�̂̕��߁i�J�n��header�ŏ㕔�j
	If Request.ServerVariables("SERVER_NAME") = "www.shigotonavi.co.jp" And InStr(Request.ServerVariables("REMOTE_HOST"),"192.168.") = 0 Then
%>
<script src="<%
	if Request.ServerVariables("HTTPS") = "off" then
		Response.write "http://www.google-analytics.com/urchin.js"
	else
		Response.write "https://ssl.google-analytics.com/urchin.js"
	end if
%>" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-2265459-3";
urchinTracker();
</script>
<%
	End If

	If IsObject(dbconn) = True Then
		If dbconn.State > 0 Then dbconn.Close
	End If
End Function


'******************************************************************************
'�T�@�v�F�E�T�C�h
'���@���F
'�쐬�ҁFLis Niina
'�쐬���F2008/02/07
'���@�l�F
'�g�p���F
'���@���F
' 08/05/20 Lis�� �����ƃi�rFC�ǉ�
'******************************************************************************
Function NaviSidemenuRight()
	Dim oRSnsr,sSQLnsr,sErrornsr,flgQEnsr

	Response.Write "<div style=""width:200px;height:135px;background-image:url(/img/rightmenu/navicafe_banner_all.jpg);margin-bottom:5px;"">"
	Response.Write "<a href=""" & HTTP_CURRENTURL & "cafe/cafe_list.asp"" title=""�i�r�J�t�F""><img src=""/img/rightmenu/navicafe_banner_top.jpg"" alt=""�i�r�J�t�F"" border=""0"" style=""margin:0px;padding:0px;""></a>"
	Response.Write "<div style=""margin-top:0px;padding:14px 6px 0px 8px;font-size:10px;line-height:15px;"">"

	'** TOP 08/11/05 Lis�� ADD
	'���݌f�ڒ���TOP3�̃g�s
	sSQLnsr = "up_GetData_NC_Topic '','','','1','3'"
	flgQEnsr = QUERYEXE(dbconn, oRSnsr, sSQLnsr, sErrornsr)
	Do While GetRSState(oRSnsr) = True
		Response.Write "<a href='" & HTTP_CURRENTURL & "cafe/cafe_detail.asp?t=" & oRSnsr.Collect("TopicID")
		Response.Write "' title='" & oRSnsr.Collect("Title") & "'>�E"
		If Len(oRSnsr.Collect("Title")) > 14 Then
			Response.Write Left(oRSnsr.Collect("Title"),14) & "..."
		Else
			Response.Write oRSnsr.Collect("Title")
		End If
		Response.Write "</a><br>"
		oRSnsr.MoveNext
	Loop
	Call RSClose(sSQLnsr)
	'** BTM 08/11/05 Lis�� ADD

	Response.Write "</div>"
	Response.Write "</div>"

	If Session("usertype") = "staff" Then '���E�҃��O�C�����Ă���ꍇ	
		Response.Write "<ul>"
		Response.Write "<li class=""rightmenu_big"">�T�|�[�g</li>"
		Response.Write "<li class=""rightmenu""><a title=""�����p�K�C�h"" href=""" & HTTP_CURRENTURL & "staff/s_aboutnavi.asp"">�����p�K�C�h</a></li>"
		Response.Write "<li class=""rightmenu""><a title=""�p���`"" href=""" & HTTP_CURRENTURL & "staff/qa.asp"">�p���`</a></li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/s_searchexplanation.asp"" title=""���d���������@"">���d���������@</a></li>"
		Response.Write "<li class=""rightmenu""><a title=""���p�K��"" href=""" & HTTP_CURRENTURL & "staff/s_kiyaku.asp"">���p�K��</a></li>"
		Response.Write "<li class=""rightmenu_end""><a title=""���⍇��(���E�Ґ�p)"" href=""" & HTTPS_CURRENTURL & "staff/access.asp"">���⍇��(���E�Ґ�p)</a></li>"
		Response.Write "<li class=""rightmenu_bottom""></li>"
		Response.Write "</ul>"
	End If

	Response.Write "<ul>"
	Response.Write "<li class=""rightmenu_big"">�P�[�^�C�ł������ƃi�r</li>"
	Response.Write "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc;""><a href=""" & HTTP_CURRENTURL & "promotion/mobilepromotion.asp"" style=""display:block;text-align:center;""><img src=""/img/sidemenu/mobile_banner.jpg"" alt=""�����ƃi�r���o�C��"" border=""0""></a></li>"
	Response.Write "<li class=""rightmenu_bottom"" style=""clear:both;""></li>"
	Response.Write "</ul>"

	Response.Write "<ul>"
	Response.Write "<li class=""rightmenu_big"">�b�����o�����i�R���v���j</li>"
	Response.Write "<li style=""height:51px; border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; border-bottom:solid 1px #eeeeee;""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp"" style=""display:block;text-align:center;""><img src=""/img/rightmenu/conpri_banner1.jpg"" alt=""�R���v��"" border=""0""></a></li>"
	Response.Write "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; padding:2px 3px; font-size:10px;"">�p�\�R���A�܂��͌g�т���쐬�������������R���r�j�ň���ł������I�T�[�r�X�I�ؖ��ʐ^����荞�߂�I</li>"
	Response.Write "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc;""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp"" style=""display:block;text-align:center;""><img src=""/img/rightmenu/conpri_banner2.jpg"" alt=""�ڂ����͂�����"" border=""0""></a></li>"
	Response.Write "<li class=""rightmenu_bottom"" style=""clear:both;""></li>"
	Response.Write "</ul>"
%><!-- #include VIRTUAL="/include/ads/navirighttext.asp" --><%
	Response.Write "<ul>"
	Response.Write "<li class=""rightmenu_big"">�R����</li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_mensetsu_index.asp"" title=""�ʐڑ΍�"">�ʐڑ΍�</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "column/column_1.asp"" title=""�h���Ј�-�����̌��̓v���ӎ�"">�h���Ј�<span style=""font-size:10px;"">-�����̌��̓v���ӎ�</span></a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_kyuuyomeisai.asp"" title=""���Ȃ��̋��^����"">���Ȃ��̋��^����</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_ready.asp"" title=""�]�E�̐S�\��"">�]�E�̐S�\��</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_proce.asp"" title=""�]�E�ɕK�v�Ȏ葱��"">�]�E�ɕK�v�Ȏ葱��</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"" title=""���i���t�o�}�j���A��"">���i���t�o�}�j���A��</a></li>"
	Response.Write "<li class=""rightmenu_bottom""></li>"
	Response.Write "</ul>"
%><!-- #include VIRTUAL="/include/ads/navirightview.asp" --><%
	Response.Write "<br>"
	Response.Write "<div align=""center"" style=""width:100%;"">"
	Response.Write "<div class=""rightmenu_big"" style=""text-align:left;"">���E�ҏ��</div>"
	Response.Write "<div style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; background-image:url(/img/sidemenu/jinzaidata_background.gif);"" align=""center"">"
	Response.Write "<table style=""width:155px; font-size:10px; text-align:left;"">"

	Dim rank(2)
	Dim rankcount(2)
	Dim idx
	idx = 0

	sSQLnsr = "SELECT top 3 Subitem,Number FROM Person_Statistics where item = '�s���{����' order by convert(int,Number) desc"
	flgQEnsr = QUERYEXE(dbconn, oRSnsr, sSQLnsr, sErrornsr)
	Do While GetRSState(oRSnsr) = True
		rank(idx) = Replace(Replace(Replace(oRSnsr.Collect("SubItem"),"�s",""),"�{",""),"��","")
		rankcount(idx) = oRSnsr.Collect("Number")
		idx = idx + 1
		oRSnsr.MoveNext
	Loop
	Call RSClose(oRSnsr)

	Response.Write "<tr>"
	Response.Write "<td>�s���{����</td>"
	Response.Write "<td>1��:" & rank(0) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(0) & "��</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td></td>"
	Response.Write "<td>2��:" & rank(1) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(1) & "��</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td></td>"
	Response.Write "<td>3��:" & rank(2) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(2) & "��</td>"
	Response.Write "</tr>"

	idx = 0

	sSQLnsr = "SELECT top 3 item,subitem, Number FROM Person_Statistics where item = '10�Α�' or item = '20�Α�' or item = '30�Α�' or item = '40�Α�' or item = '50�Α�' or item = '60�Έȏ�' order by convert(int,Number) desc"
	flgQEnsr = QUERYEXE(dbconn, oRSnsr, sSQLnsr, sErrornsr)
	Do While GetRSState(oRSnsr) = True
		rank(idx) = Replace(oRSnsr.Collect("Item"),"��","") & oRSnsr.Collect("SubItem")
		rankcount(idx) = oRSnsr.Collect("Number")
		idx = idx + 1
		oRSnsr.MoveNext
	Loop
	Call RSClose(oRSnsr)

	Response.Write "<tr>"
	Response.Write "<td>�N���</td>"
	Response.Write "<td>1��:" & rank(0) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(0) & "��</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td></td>"
	Response.Write "<td>2��:" & rank(1) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(1) & "��</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td></td>"
	Response.Write "<td>3��" & rank(2) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(2) & "��</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td colspan=""3"" align=""right""><a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp""><img src=""/img/sidemenu/kuwashiku_min.jpg"" alt=""�ڂ����͂�����"" border=""0""></a>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""rightmenu_bottom"" style=""clear:both;""></div>"
	Response.Write "<br>"

	Response.Write "</div>"

	If Session("usertype") = "staff" Then '���E�҃��O�C�����Ă���ꍇ	
	ElseIf Session("usertype") = "company" Or Session("usertype") = "dispatch" Then '��ƃ��O�C�����Ă���ꍇ
	Else
	End If
End Function
%>
