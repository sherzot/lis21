                                                                                                                                                                                                                                                                    <%
'******************************************************************************
'�T�@�v�F�T�C�h���j���[
'���@���FSidemenuType	0�y�g�b�v�z1�y���E�ҁz2�y��Ɓz3�y���p�z4�y�㗝�X�z
'���@�l�F
'�g�p���F
'���@���F2008/02/07 LIS K.Niina �쐬
'�@�@�@�F2008/05/20 LIS M.Hayashi �����ƃi�rFC�ǉ�
'�@�@�@�F2011/02/16 LIS K.Kokubo �X�p���I��title�����폜,�����ƃi�r�c�C�b�^�[�o�i�[�폜
'      �F2015/11/20 LIS K.Kimura �T�C�h���j���[��ύX
'******************************************************************************
Function NaviSidemenu(SidemenuType)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim iTabIndexType
	Dim sHTML10	'���E�҃��O�C������
	Dim sHTML11	'��ƃ��O�C������
	Dim sHTML12	'���E��,���My Menu
	Dim sHTML13	'�A�N�Z�X�ۋ�
	Dim sHTML20	'TOP�y�[�W�̃o�i�[
	Dim sHTML21	'���E�ҏ��
	Dim sHTML30	'�В��u���O�o�i�[
	Dim sHTML31	'P�}�[�N
	Dim sHTML40	'�В��i�r�o�i�[(���E�Ҍ���)
	Dim sHTML41	'�В��i�r�o�i�[(��ƌ���)
	Dim sHTML60	'�c�C�b�^�[
	Dim sHTML61	'�l�ޏЉ�c�C�b�^�[
	Dim sHTML62	'���k�n�������m���n�k�̉e���ɂ���
	Dim sHTML63	'�X�}�z
	Dim sHTML64	'Facebook�y�[�W�o�i�[
	Dim sHTML80	'���X�Ј���W
	Dim sHTML90	'�h������h���X�^�b�t�A���P�[�g�o�i�[
	Dim sHTML91	'�Г��Č��}��o�i�[
	Dim sHTML100 '�w�Ԃ̃��c
	Dim sHTML	'�^�uIndex���̃i�r�Q�[�V��������

	Dim sScript	'���O�C���`�F�b�N�X�N���v�g

	Dim sParamName
	Dim si
	Dim sMidoku
	Dim rank(2)
	Dim rankcount(2)
	Dim idx
	Dim iCollectionCount
	
		Dim cnt
	Dim iAll		'���E�Ґ�
	Dim iOrderCnt	'�f�ڒ����l��
	Dim iCompanyCnt	'�f�ڒ���Ɛ�

	iAll = 0
	iOrderCnt = 0
	iCompanyCnt = 0

	sSQL = ""
	sSQL = sSQL & "/* �����ƃi�r �s�n�o�y�[�W�p�̏o�̓f�[�^�擾 */"
	sSQL = sSQL & "EXEC up_DtlTopStatus;"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		iAll = oRS.Collect("StaffCnt")
		iOrderCnt = oRS.Collect("OrderCnt")
		iCompanyCnt = oRS.Collect("CompanyCnt")
	End If
	Call RSClose(oRS)

	sHTML10 = ""
	sHTML11 = ""
	sHTML20 = ""
	sHTML21 = ""
	sHTML30 = ""
	sHTML31 = ""
	sHTML40 = ""
	sHTML60 = ""
	sHTML61 = ""
	sHTML62 = ""
	sHTML63 = ""
	sHTML64 = ""
	sHTML80 = ""
	sHTML90 = ""
	sHTML = ""

    G_SSLFLAG = True

	iTabIndexType = getTabIndexType(Request.ServerVariables("URL"))

	If G_USERID = "" Then
		si = GetForm("si","2")

		'<���O�C���`�F�b�N�X�N���v�g>
		sScript = ""
		sScript = sScript & "<script type=""text/javascript"">"
		sScript = sScript & "function LoginCheckIdreg(){"
		sScript = sScript & "var ofrm = document.forms.frmlogin;"
		sScript = sScript & "if(!navigator.cookieEnabled) {"
		sScript = sScript & "alert('cookie�i�N�b�L�[�j�̗��p���ł��Ȃ��ݒ�ɂȂ��Ă��܂��B\n�u���E�U��Z�L�����e�B�[�\�t�g��cookie�ݒ�����m�F�������B');"
		sScript = sScript & "return false;"
		sScript = sScript & "}"
		sScript = sScript & "if(ofrm.CONF_UserID.value.length === 0){alert('�F��ID����͂��Ă��������B');return false;}"
		sScript = sScript & "if(ofrm.CONF_Password.value.length === 0){alert('�p�X���[�h����͂��Ă��������B');return false;}"
		sScript = sScript & "if(ofrm.CONF_Password.value.length < 3 || ofrm.CONF_Password.value.length > 20){alert('�p�X���[�h�͂R�`�Q�O�����œ��͂��Ă��������B');return false;}"
		sScript = sScript & "ofrm.submit();"
		sScript = sScript & "}"
		sScript = sScript & "</script>"
		'</���O�C���`�F�b�N�X�N���v�g>

        '<����o�i�[>
'            sHTML10 = sHTML10 & "<a href=""https://www.youtube.com/watch?v=T3n06VU8T-Q" & HTTPS_CURRENTURL & "valueoffer/""TARGET="_blank"><img src=""/img/common/tutrial_banner01.png"" alt=""�o�����[�I�t�@�["" style=""margin-top:3px;""></a>"
            sHTML10 = sHTML10 & "<a href=""https://www.youtube.com/watch?v=T3n06VU8T-Q" & HTTPS_CURRENTURL & "valueoffer/""target=""_blank""><img src=""/img/common/tutrial_banner01.png"" alt=""�o�����[�I�t�@�["" style=""margin-top:3px;""></a>"
        '�|�C���g�o�q�o�i�[����
            'sHTML10 = sHTML10 & "<a href=""https://youtu.be/9FlpxFA6TYc""target=""_blank""><img src=""/img/common/conpri_banner1.png"" alt=""�o�����[�I�t�@�["" style=""margin-top:3px;""></a>"
        '</����o�i�[>
        sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "pr/pushpoint.asp""><img src=""/img/common/how_to_use2.png""></a>"

	'�ȉ��̕�������X�N���[���ɒǏ]����T�C�h���j���[
	sHTML10 = sHTML10 & "<div class=""floatingmenu"" id=""moveside""><div align=""center"">"
        if GetForm("ordercode", 2) <> "" then
			if IsRE(Trim(Replace(Server.HTMLEncode(GetForm("ordercode", 2)), "'", "�f")), "^J\d\d\d\d\d\d\d$", True) = True then
				sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp?ordercode=" & GetForm("ordercode", 2) & """><img src=""/img/common/reg1_button_big_3.png"" border=""0"" alt=""����o�^(����)"" style=""margin-top:3px;""></a>"
			else
				sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp""><img src=""/img/common/reg1_button_big_3.png"" border=""0"" alt=""����o�^(����)"" style=""margin-top:3px;""></a>"
			end if
		else
			sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp""><img src=""/img/common/reg1_button_big_3.png"" border=""0"" alt=""����o�^(����)"" style=""margin-top:3px;""></a>"
		end if

		sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "/point/pr/""target=""_blank""><img src=""/img/neo/point_present.png""></a>"

		sHTML10 = sHTML10 & "<script type=""text/javascript""><!-- document.forms[0].UserID.focus(); // --></script>"
		sHTML10 = sHTML10 & "</div>"

		'<�R���T���Љ�>
		'2016/04/11 �ؑ��ǉ�
		'2016/04/21 3�l�������Ȃ��̂Ŕ�\��
		'sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "consultant/consultantbranch.asp""><img src=""/img/common/con_int.png"" alt=""�ݐЃR���T���^���g�Љ�"" style=""margin-top:3px;border:1px solid #000;""></a>"
		'</�R���T���Љ�>

        '<�o�����[�I�t�@�[����>
        'If Request.ServerVariables("PATH_INFO") = "/staff/s_resume_kakikata.asp" Then
            'sHTML10 = sHTML10 & "<div align=""center"" style=""margin: 9px 0px 5px;"">"
            'sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "valueoffer/persona.asp""><div><img src=""/img/C_K_NAVI.GIF"" height=""50"">�o�����[�I�t�@�[����</div></a>"
            'sHTML10 = sHTML10 & "</div>"
            'sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "valueoffer/persona02.asp""><img src=""/img/common/persona_banner02.png"" alt=""�o�����[�I�t�@�["" style=""margin-top:3px;""></a>"

        'End If
        '</�o�����[�I�t�@�[����>

		
		'sHTML10 = sHTML10 & "<form id=""mailReg""><input type=""text"" value=""mail""><br><input type=""button"" value=""�����}�K�o�^"" onClick=""location.href='/staff/mailReg.asp'""></form>"
		'sHTML10 = sHTML10 & "<a href=""http://www.shigotonavi.co.jp/iphone/index.html"" target=""_blank""><img src=""/img/link/iphone_banner.png"" style=""width: 210px;""></a>"
		'sHTML10 = sHTML10 & "<a href=""http://www.a-rirekisyo.jp/"" target=""_blank""><img src=""/img/link/a-resume_banner.gif"" style=""width: 210px;""></a>"
		'sHTML10 = sHTML10 & "<a href=""/recruit/se/""><img src=""/recruit/img/banner.png""></a>"

		'<��ƃ��O�C���t�H�[��>

		'</��ƃ��O�C���t�H�[��>
	End If

	'<���E��My Menu>
    '2015/08/19 �ؑ����C�F�i�g�݃��j���[
	If G_USERTYPE = "staff" Then
		'���ǃ��[������
		sSQL = "SELECT COUNT(*) AS Cnt FROM MailHistory WITH(NOLOCK) WHERE ReceiverCode ='" & G_USERID & "' AND OpenDay IS NULL AND ReceiverDelFlag = '0'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If oRS.Collect("Cnt") = 0 Then
				sMidoku = "(<img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">����" & oRS.Collect("Cnt") & "��)"
			Else
				sMidoku = "(<span style=""color:#ff0000; font-weight:bold;""><img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">����" & oRS.Collect("Cnt") & "��</span>)"
			End If
		End If

		sHTML12 = sHTML12 & "<div id=""moveside""><ul class=""smartSidenone"">"

         '<����o�i�[>
'            sHTML10 = sHTML10 & "<a href=""https://www.youtube.com/watch?v=T3n06VU8T-Q" & HTTPS_CURRENTURL & "valueoffer/""><img src=""/img/common/tutrial_banner01.png"" alt=""�o�����[�I�t�@�["" style=""margin-top:3px;""></a>"
            sHTML10 = sHTML10 & "<a href=""https://www.youtube.com/watch?v=T3n06VU8T-Q" & HTTPS_CURRENTURL & "valueoffer/""target=""_blank""><img src=""/img/common/tutrial_banner01.png"" alt=""�o�����[�I�t�@�["" style=""margin-top:3px;""></a>"
        '�|�C���g�o�q�o�i�[����
             sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "/point/pr""target=""_blank""><img src=""/img/neo/point_present.png""style=""margin-top:3px;""></a>"
        '</����o�i�[>


        '2015/08/28 �Ȃ�
        'If G_SSLFLAG = False Then
		'sHTML12 = sHTML12 & "<li class=""sidemenu_staff_big"">My&nbsp;���j���[&nbsp;(<a href=""" & HTTP_CURRENTURL & "logout.asp"">���O�A�E�g</a>)</li>"
        'Else
		sHTML12 = sHTML12 & "<li class=""sidemenu_staff_big"">My&nbsp;���j���[&nbsp;(<a href=""" & HTTPS_CURRENTURL & "logout.asp"">���O�A�E�g</a>)</li>"
        'End IF
        '����
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/s_login.asp"">My&nbsp;�y�[�W</a></li>"
        sHTML12 = sHTML12 & "<li class=""sidemenu"" style=""border-bottom:none;""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/person_detail.asp"">�v���t�B�[���Ǘ�</a></li>"

		'�o�����[�I�t�@�[
		'2015/03/02 �r�c���C
        Dim sMikaitou
		'���ǃ��[������
		sSQL = "EXEC up_ExistsOfferStep2c '" & G_USERID & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If oRS.RecordCount = 0 Then
				sMikaitou = "(<img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">����" & oRS.RecordCount & "��)"
			Else
				sMikaitou = "(<span style=""color:#ff0000; font-weight:bold;""><img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">����" & oRS.RecordCount & "��</span>)"
			End If
		End If
		Call RSClose(oRS)

        '�|�C���g�\��
		'sHTML12 = sHTML12 & "<li class=""sidetitle"">GPoint�\��<img src=""/img/c_new.gif"" alt="""" border=""0"">"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/apply_GPoint.asp?PointType=login"">���O�C���|�C���g�i1��1��j</a></li>"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/apply_GPoint.asp?PointType=DRegist"">�o�^�|�C���g</a></li>"

		sHTML12 = sHTML12 & "<li class=""sidetitle"">�o�����[�I�t�@�[</li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/step2a.asp"">��]��������</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/step2c.asp?offer_ques=true"">���Ȃ��ɋ�����������<br>��Ƃ���̎���" & sMikaitou & "</a></li>"

        '����E��
        sHTML12 = sHTML12 & "<li class=""sidetitle"">�]�E�T�|�[�g</li>"
        sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/my_footprint.asp"">�{������</a></li>"
        sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/watchlist.asp"">���C�ɓ��胊�X�g</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/edit_list.asp"">����ꗗ</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/mailhistory_person.asp"">���[���Ǘ�" & sMidoku & "</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/schedule/"">�X�P�W���[���Ǘ�</a></li>"


        '2015/09/01�@�v���C�A���󂠂܂�Ӗ����Ȃ��y�[�W
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/jobcon/"">�W���u�E�R���V�F���W��</a></li>"
        
        '<�X�e�b�v6�Ή�>
		'sHTML12 = sHTML12 & "<li class=""sidetitle"">�o�����[�I�t�@�[</li>"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/step2a.asp"">��]��������</a></li>"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/step2c.asp"">��Ƃ���̎���" & sMikaitou & "</a></li>"
		'</�X�e�b�v6�Ή�>

        '������
        sHTML12 = sHTML12 & "<li class=""sidetitle"">�������E�E���o�����̈��</li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/resume_print.asp"">�������E�E���o�������</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/resume_picture.asp"">�������p�ʐ^�o�^</a></li>"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/resumemanual.pdf"" target=""_blank"">�������쐬�}�j���A��</a></li>"
        '��~�މ�
        sHTML12 = sHTML12 & "<li class=""sidetitle"">�e��ݒ�</li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/person_edit6.asp"">��]�����i���[���z�M�����j</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/notification_mail_service.asp"">�X�P�W���[���ʒm</a></li>"

        '2015/09/01�@�v���C�A���󂠂܂�Ӗ����Ȃ��y�[�W
        'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/searchordercondition/"">���������Ǘ�</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/changepassword.asp"">�p�X���[�h�̕ύX</a></li>"
        sHTML12 = sHTML12 & "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "suspension/questionnarie.asp"">�x�~�E�މ�</a></li>"

		sHTML12 = sHTML12 & "</ul></div><!--/#moveside-->"
'		sHTML12 = sHTML12 & "<a href=""/neo/oiwai/"" target=""_blank"" id=""oiwai_page"">�|�C���g�\��</a>"
		sHTML12 = sHTML12 & "<a href=""/point/pr/"" target=""_blank""><img src=""/img/neo/point_present.png""></a>"
		'sHTML12 = sHTML12 & "<a href=""/recruit/se/""><img src=""/recruit/img/banner.png""></a>"
	End If
	'</���E��My Menu>

	'<���My Menu>
	If G_USERTYPE = "company" Then
		'<���l���擾>
		iCollectionCount = 0
		sSQL = "EXEC sp_GetCollectionCount '" & G_USERID & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			iCollectionCount = oRS.Collect("Cnt")
		End If
		Call RSClose(oRS)
		'</���l���擾>

		sHTML12 = sHTML12 & "<ul class=""smartSidenone"">"
		sHTML12 = sHTML12 & "<li class=""sidemenu_company_big"">My&nbsp;���j���[&nbsp;(<a href=""" & HTTP_CURRENTURL & "logout.asp"">���O�A�E�g</a>)</li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "management/index.asp"" target=""_blank"">�Ǘ���ʃy�[�W</a></li>"
		sHTML12 = sHTML12 & "</ul>"

	End If
	'</���My Menu>

	'<TOP�y�[�W�̃i�r�Q�[�V�����o�i�[�L��>
	If SideMenuType = 0 Then

		'<�����������쐬>
		sHTML20 = sHTML20 & "<ul>"
		sHTML20 = sHTML20 & "<li class=""sidemenu_big"">�֗��c�[��</li>"
		sHTML20 = sHTML20 & "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; line-height:17px;"">"
		sHTML20 = sHTML20 & "<a href=""/staff/s_resume.asp"" style=""display:block; background-image:url(/img/sidemenu/resume_banner.jpg); width:154px; height:73px; font-size:10px; padding:54px 0px 0px 14px; color:#444444; text-decoration:none;"">"
		sHTML20 = sHTML20 & "�K�v�ȍ��ڂ���͂��邾���Ŋ����I<br>55���l���g�����S�̃T�[�r�X�I<br>�����ɍ�����������������I"
		sHTML20 = sHTML20 & "</a>"
		sHTML20 = sHTML20 & "</li>"
		sHTML20 = sHTML20 & "</ul>"


	End If
	'</TOP�y�[�W�̃i�r�Q�[�V�����o�i�[�L��>

	'<���E�ҏ��>
	If iTabIndexType = 0 Then
		idx = 0
		sSQL = "SELECT TOP 3 Subitem,Number FROM Person_Statistics WHERE item = '�s���{����' ORDER BY CONVERT(INT,Number) DESC;"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		Do While GetRSState(oRS) = True
			rank(idx) = Replace(Replace(Replace(oRS.Collect("SubItem"),"�s",""),"�{",""),"��","")
			rankcount(idx) = oRS.Collect("Number")
			idx = idx + 1
			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		sHTML21 = sHTML21 & "<div align=""center"" style=""width:100%;"">"
		sHTML21 = sHTML21 & "<div class=""sidemenu_big"" style=""text-align:left;"">���E�ҏ��</div>"
		sHTML21 = sHTML21 & "<div style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; background-image:url(/img/sidemenu/jinzaidata_background.gif);"" align=""center"">"
		sHTML21 = sHTML21 & "<table style=""width:155px; font-size:10px; text-align:left;"">"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td>�s���{����</td>"
		sHTML21 = sHTML21 & "<td>1��:" & rank(0) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(0) & "��</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td></td>"
		sHTML21 = sHTML21 & "<td>2��:" & rank(1) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(1) & "��</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td></td>"
		sHTML21 = sHTML21 & "<td>3��:" & rank(2) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(2) & "��</td>"
		sHTML21 = sHTML21 & "</tr>"

		idx = 0
		sSQL = "SELECT TOP 3 item,subitem, Number FROM Person_Statistics where item = '10�Α�' or item = '20�Α�' or item = '30�Α�' or item = '40�Α�' or item = '50�Α�' or item = '60�Έȏ�' order by convert(int,Number) desc"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			rank(idx) = Replace(oRS.Fields("Item").Value,"��","") & oRS.Fields("SubItem").Value
			rankcount(idx) = oRS.Fields("Number").Value
			idx = idx + 1
			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td>�N���</td>"
		sHTML21 = sHTML21 & "<td>1��:" & rank(0) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(0) & "��</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td></td>"
		sHTML21 = sHTML21 & "<td>2��:" & rank(1) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(1) & "��</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td></td>"
		sHTML21 = sHTML21 & "<td>3��:" & rank(2) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(2) & "��</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td colspan=""3"" align=""right""><a href=""/company/c_staffdata.asp""><img src=""/img/sidemenu/kuwashiku_min.jpg"" alt=""�ڂ����͂�����"" border=""0""></a>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "</table>"
		sHTML21 = sHTML21 & "</div>"
		sHTML21 = sHTML21 & "<br style=""clear:both;"">"
		sHTML21 = sHTML21 & "</div>"
	End If
	'</���E�ҏ��>

	'<P�}�[�N>
	sHTML31 = sHTML31 & "<div align=""center"" style=""margin:10px 0 5px 0;"">"
	sHTML31 = sHTML31 & "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""�v���C�o�V�[�}�[�N"" border=""0""></a><br>"
	sHTML31 = sHTML31 & "<a href=""" & HTTP_CURRENTURL & "privacy/privacy.asp"">�l���ی�ɂ���</a>"
	sHTML31 = sHTML31 & "</div>"
	sHTML31 = sHTML31 & "<div class=""center""></div>"
	'</P�}�[�N>



	'<�Г��Č��}��o�i�[>
	If Date <= "2011/04/20" Then
		sHTML91 = sHTML91 & "<div style=""width:170px;margin-bottom:10px;"">"
		sHTML91 = sHTML91 & "<img src=""/img/banner/gu0001.jpg"" alt=""�H�i�����̕��͌o����,�Ζ��n�͌Q�n�����c�s,�}��"" border=""0"" style=""cursor:pointer;"" onclick=""location.href='/ad_banner_control/c_r.asp?origin=gu0001';"">"
		sHTML91 = sHTML91 & "</div>"
	End If
	'</�Г��Č��}��o�i�[>

	If iTabIndexType = 0 Then
		'<�͂��߂Ă̕�>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">�͂��߂Ă̕�</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp"">����o�^�i�������o�^�j</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/qa.asp"">�����ƃi�rQ&A</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">���l��񌟍�</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/passwordreminder.asp"">ID�E�p�X���[�h�̍Ď擾</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/mobilepromotion.asp"">�����ƃi�r���o�C��</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/conpri_riyou.asp"">�R���r�j�v�����g(�Z�u���]�C���u��)�̗��p���@</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp"">����������i���[�\���E�t�@�~���[�}�[�g�E�T�[�N��K�E�T���N�X�j</a></li>"
        sHTML = sHTML & "<li class=""sidetitle"">�������E�E���o�����̈��</li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_711.png""></a></li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_familymart.png""></a></li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_lawson.png""></a></li>"
		sHTML = sHTML & "</ul>"
        sHTML = sHTML & "</ul>"
		'</�͂��߂Ă̕�>
	ElseIf iTabIndexType = 1 Then
		'<���l��T��>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">���l��T��</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">���l����T��</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "order/order_list_accesscount.asp"">�l�C�̋��l���g�b�v10</a></li>"
		sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "railway/railway_search1.asp"">��������</a></li>"
		sHTML = sHTML & "</ul>"
		'</���l��T��>
	ElseIf iTabIndexType = 2 Then
		'<�֗��c�[��>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">�֗��c�[��</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_resume.asp"">�������̎����쐬</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_resume_kakikata.asp"">�������̏�����</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_resume_qa.asp"">�������p���`</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_careersheet.asp"">�E���o�����̎����쐬/�t�H�[�}�b�g�̃_�E�����[�h</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_careersheet_kakikata_1.asp"">�E���o�����̏�����</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/motive_index.asp"">�u�]���@���[�J�[</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_jikopr.asp"">����PR���[�J�[</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_taishokunegai.asp"">�ސE��̏�����</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_year_calculation.asp"">����E�a��/�w���v�Z</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp"">����������i�Z�u���]�C���u���j</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp"">����������i���[�\���E�t�@�~���[�}�[�g�E�T�[�N��K�E�T���N�X�j</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "conpri/"">���ވ���T�[�r�X</a></li>"
        sHTML = sHTML & "<li class=""sidetitle"">�������E�E���o�����̈��</li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_711.png""></a></li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_familymart.png""></a></li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_lawson.png""></a></li>"
		sHTML = sHTML & "</ul>"
		'</�֗��c�[��>
	ElseIf iTabIndexType = 3 Then
		'<�]�E�T�|�[�g>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">�]�E�T�|�[�g</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/jobcon/introduction.asp"">�W���u�E�R���V�F���W��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/jobcon/careeranalyzer/"">���ȕ��̓c�[��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/jobcon/searchadvice/"">���������⏕�c�[��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/jobcon/interviewsimulator/"">�ʐڑ΍�</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/notification_mail_service.asp"">�X�P�W���[���ʒm�T�[�r�X</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_ready.asp"">�]�E�̐S�\��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_proce.asp"">�]�E�ɕK�v�Ȏ葱��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"">�ʐڑ΍�}�j���A��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_kyuuyomeisai.asp"">���^���ׂɂ���</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/navistep_index.asp"">���߂Ă̓]�E����</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "column/column_index.asp"">�]�E�E�A�E�R����</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_mynavi.asp"">�K�E�f�f�u���Ԃ�i�r�v</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/businesscolumns/"">�r�W�l�X�R����</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_introduce.asp"">�l�ޏЉ�</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_temporary.asp"">�l�ޔh��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_temptoperm.asp"">�Љ�\��h��</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "staff/jobcon/careerconsultation/"">�L�����A���k</a></li>"
		sHTML = sHTML & "</ul>"
		'</�]�E�T�|�[�g>
	ElseIf iTabIndexType = 4 Then
		'<�R�~���j�e�B>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">�R�~���j�e�B</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "cafe/cafe_list.asp"">�����ƃi�r�J�t�F</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_introduce_swf.asp"">�l�ޏЉ</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_mynavi.asp"">�K�E�f�f�u���Ԃ�i�r�v</a></li>"
		sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "s_contents/enquete.asp"">�����ƃi�r�A���P�[�g</a></li>"
		sHTML = sHTML & "</ul>"
		'</�R�~���j�e�B>
	ElseIf iTabIndexType = 5 Then
		'<�̗p���S����>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">�̗p���S����</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/about.asp"">�����ƃi�r�̓��F</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_function.asp"">�T�[�r�X�T�v</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "company/costperformance/"">�̗p���P�T�|�[�g</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "company/request01.asp"">���l�L���f�ڐ\����</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/kiyaku.asp"">�����p�K��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "company/access.asp"">���⍇��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"">�l�ޏЉ�</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"">�l�ޔh��</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"">�Љ�\��h��</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_successpoint.asp"">�����p�ɂ���</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_scout3point.asp"">�X�J�E�g���[���̃|�C���g</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/qa.asp"">�̗pQ&A</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "jinzaisearch/index.asp"">��񂨎�������</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp"">���E�ҏW�v�f�[�^</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/research.asp"">�̗p���@�f�f</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/charge.asp"">�����p�v����</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "company/c_voice.asp"">�����p��Ƃ��܂̐�</a></li>"
		sHTML = sHTML & "</ul>"
		'</�̗p���S����>
	ElseIf iTabIndexType = 6 Then
		'<My Page(���E��)>
		'</My Page(���E��)>
	ElseIf iTabIndexType = 7 Then
		'<My Page(���)>
		'</My Page(���)>
	ElseIf iTabIndexType = 8 Then	
	
		sHTML = sHTML &""
		

		
	End If


	Response.Write "</div>"'���C���R���e���c�̕��w��div�̕߁i�J�n��header�ŉ����j

	If SidemenuType <> 9 Then

		Response.Write "<nav id=""side"">"

		If iTabIndexType = 0 Then
			Response.Write sScript
			Response.Write sHTML10
			Response.Write sHTML12
			Response.Write sHTML91
			Response.Write sHTML80 '���X�Ј���W
			Response.Write sHTML
			Response.Write sHTML40
			Response.Write sHTML60
			Response.Write sHTML90
			Response.Write sHTML31
		ElseIf iTabIndexType = 1 Then
			Response.Write sHTML10
			Response.Write sHTML12		
			
				%>
			
       <ul>
		<li class="sidemenu_big">�]�E�T�|�[�g</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume.asp">�������̎����쐬/�t�H�[�}�b�g�̃_�E�����[�h</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_kakikata.asp">�������̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_qa.asp">������Q��A</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet.asp">�E���o�����̎����쐬/�t�H�[�}�b�g�̃_�E�����[�h</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet_kakikata_1.asp">�E���o�����̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_taishokunegai.asp">�ސE��̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_year_calculation.asp">����E�a��/�w���v�Z</a></li>
        
        <li class="sidetitle">�������E�E���o�����̈��</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/conpripromotion.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_711.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_familymart.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_lawson.png"></a></li>
        
        <li class="sidetitle">�]�E�T�|�[�g�ē�</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_introduce.asp">�l�ޏЉ�</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temporary.asp">�l�ޔh��</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temptoperm.asp">�Љ�\��h��</a></li>
		<!--<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/careerconsultation/">�L�����A���k</a></li>-->
        
        <li class="sidetitle">�}�b�v</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>type_map.asp">�E��E�Ǝ�ʃ}�b�v</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>area_map.asp">�n��ʃ}�b�v</a></li>	
        <li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>keyword_map.asp">�L�[���[�h�}�b�v</a></li>	
		</ul>
		<!-- <a href="/point/pr/" target="_blank"><img src="/img/neo/point_present.png"></a> -->

        <!--<div id="side_pickup">
			������PickUp
        </div>-->
			<!--<a href="http://www.shigotonavi.co.jp/order/order_detail.asp?OrderCode=J0066098" target="_self"><img src="/img/banner/jisya/tokyo_20120822.gif"></a>-->
			<%	
		ElseIf iTabIndexType = 2 Then
			Response.Write sScript
			Response.Write sHTML10
			Response.Write sHTML12
			Response.Write sHTML91
			Response.Write sHTML80 '���X�Ј���W
			Response.Write sHTML
			Response.Write sHTML40
			Response.Write sHTML60
			Response.Write sHTML90
			Response.Write sHTML31
		ElseIf iTabIndexType = 3 Then
			Response.Write sScript
			Response.Write sHTML10
			Response.Write sHTML12
			Response.Write sHTML91
			Response.Write sHTML80 '���X�Ј���W
			Response.Write sHTML
			Response.Write sHTML40
			Response.Write sHTML60
			Response.Write sHTML90
			Response.Write sHTML31
		ElseIf iTabIndexType = 4 Then
			Response.Write sScript
			Response.Write sHTML10
			Response.Write sHTML12
			Response.Write sHTML91
			Response.Write sHTML80 '���X�Ј���W
			Response.Write sHTML
			Response.Write sHTML40
			Response.Write sHTML60
			Response.Write sHTML90
			Response.Write sHTML31
		ElseIf iTabIndexType = 5 Then
			Response.Write sScript
			Response.Write sHTML11
			Response.Write sHTML12
			Response.Write sHTML41
			Response.Write sHTML
			Response.Write sHTML21
			Response.Write sHTML61

			Response.Write sHTML31
			
		ElseIf iTabIndexType = 6 Then
			Response.Write sHTML10
			Response.Write sHTML12
			Response.Write sHTML91
			Response.Write sHTML80 '���X�Ј���W
			Response.Write sHTML
			Response.Write sHTML40
			Response.Write sHTML60
			Response.Write sHTML90
			Response.Write sHTML31
		ElseIf iTabIndexType = 7 Then
			Response.Write sHTML11
			Response.Write sHTML12
			Response.Write sHTML41
			Response.Write sHTML
			Response.Write sHTML61
			Response.Write sHTML31
			
		ElseIf iTabIndexType = 8 Then '�w��
			Response.Write sHTML10
			Response.Write sHTML12
			%>
			
        <ul>
		<li class="sidemenu_big">���ȕ���</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/introduction.asp">�W���u�E�R���V�F���W��</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/careeranalyzer/">���ȕ��̓c�[��</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/searchadvice/">���������⏕�c�[��</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_kyuuyomeisai.asp">���^���ׂɂ���</a></li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>staff/notification_mail_service.asp">�X�P�W���[���ʒm</a></li>
		</ul>
        
        <ul>
		<li class="sidemenu_big">�X�L���A�b�v</li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/interviewsimulator/">�ʐڑ΍�</a></li>
		</ul>
			
        <ul>
		<li class="sidemenu_big">�]�E�m�E�n�E</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_ready.asp">�]�E�̐S�\��</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_proce.asp">�]�E�ɕK�v�Ȏ葱��</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_goukaku.asp">�ʐڑ΍�}�j���A��</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/navistep_index.asp">���߂Ă̓]�E����</a></li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>column/column_index.asp">�]�E�E�A�E�R����</a></li>
		</ul>	

       <!-- <ul>
		<li class="sidemenu_big">�n�������̃y�[�W</li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>s_contents/s_localgoverment.asp">�n�������̃y�[�W</a></li>
		</ul>	-->

        <ul>
		<li class="sidemenu_big">�r�W�l�X�R����</li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>s_contents/businesscolumns/">�r�W�l�X�R����</a></li>
		</ul>
    		
			<%
			
		ElseIf iTabIndexType = 10 Then 'TOP

	If G_USERID = "" Then
		si = GetForm("si","2")

%>
<div style="width:975px; margin:10px 0 0 -773px;">
<div class="left">
<p id="shigotonavi_member">
<%
	'<���l���A��Ɛ��A���E�Ґ�>

	Response.Write "<img src=""/img/top/countericon_order.gif"" alt=""���l��"" border=""0"" style=""margin:0px 2px;"">���l<span class=""cnt"">" & iOrderCnt & "</span>��&nbsp;"
	Response.Write "<img src=""/img/top/countericon_company.gif"" alt=""��Ɛ�"" border=""0"" style=""margin:0px 2px;"">���<span class=""cnt"">" & iCompanyCnt & "</span>��&nbsp;"
	Response.Write "<img src=""/img/top/countericon_staff.gif"" alt=""���E�Ґ�"" border=""0"" style=""margin:0px 2px;"">���E��<span class=""cnt"">" & iAll & "</span>�l&nbsp;"
	Response.Write "" & MonthName(Month(Now)) & Day(Now) & "��(" & Left(WeekdayName(Weekday(Now)),1) & ")" & "�X�V"
	'</���l���A��Ɛ��A���E�Ґ�>

%>
</p>
<a href="<%= HTTPS_CURRENTURL %>pr/pushpoint.asp"><img src="/img/common/how_to_use2.png"></a>
<a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp"><img src="/img/common/reg1_button_big_3.png" border="0" alt="����o�^(����)" style="margin:0 0 0 15px;"></a>
</div>


</div>
<script type="text/javascript"><!-- document.forms[0].UserID.focus(); // --></script>

<form id="frmlogin" method="post" action="<%= HTTPS_CURRENTURL %>login_check.asp">

<%		If LCase(GetForm("JUMP_URL_FLAG",2)) = "true" Then
			For Each sParamName In Request.QueryString
			%><input type="hidden" name="<%= sParamName %>" value="<%= GetForm(sParamName,2) %>">
<%			Next
		End If
%>

<div class="right" style="width:515px; border:2px solid #ff9739; border-radius:8px; padding:0 10px 0 0;">
<div style="width:120px; float:left;font-size: 15px; font-weight: bold; color: chocolate; text-align:center; line-height: 90px; border-radius:8px 0 0 8px;border-right: 2px dashed #ffd2a9;">���O�C��</div>

<div class="right" style="float: left!important;margin: 11px 0 5px 25px;">

	<div class="left center" style="font-size: 14px;">
		ID
	<% If si <> "" Then %>
			<input type="text" name="CONF_UserID" value="<%= si %>" style="margin: 0 5px;width:120px; height: 25px; border-radius: 4px;">
	<%	Else %>
			<input type="text" name="CONF_UserID" value="<%= Request.Cookies("id_memory") %>" style="margin: 0 5px;width:120px;height: 25px;border-radius: 4px;">
	<%	End If %>
		�p�X���[�h<input type="password" name="CONF_Password" value="" style="margin: 0 0 0 5px;width:120px;height: 25px;border-radius: 4px;">
	</div>
<br clear="all">
	<div align="right" style="margin:0 0 5px 0;">
		<label><input type="checkbox" name="frmautologinflag" value="1">����۸޲�</label>[<span style="color:#0045f9; cursor:pointer;" onclick="window.open('/infomation/autologin.asp','autologin','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=220');"><u>�H</u></span>]
        
		<a href="<%= HTTPS_CURRENTURL %>staff/qa.asp#003" style="font-size:10px;">۸޲݂ł��Ȃ�</a>
		<a href="<%= HTTPS_CURRENTURL %>staff/passwordreminder.asp" style="font-size:10px;">ID�E�߽ܰ�ނ�Y�ꂽ</a> 
		<input type="submit" value="���O�C��" onclick="DataCheckIdreg(); return false" style="background: #FFA500;
    color: #fff;
    font-weight: bold;
    border: none;
    border-radius: 5px;
    padding: 2px 10px;
    font-size: 14px;">
<br>
		</div>
		</div>
        </div></div>
        

</form>

<%
	End If		
			'Response.Write sHTML12

		ElseIf iTabIndexType = 11 Then '��
			Response.Write sHTML10
			Response.Write sHTML12	
			
			
		ElseIf iTabIndexType = 12 Then '�����N
			Response.Write sHTML10
			Response.Write sHTML12
		
		%>	
		<ul>
		<li class="sidemenu_big">�R���e���c</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_introduce_swf.asp">�l�ޏЉ</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_mynavi.asp">�K�E�f�f�u���Ԃ�i�r�v</a></li>
        <li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>s_contents/enquete.asp">�����ƃi�r�A���P�[�g</a></li>

		</ul>
		<%
        		
		ElseIf iTabIndexType = 13 Then '�T��
			
			Response.Write sHTML12		
			
				%>
			
       <ul>
		<li class="sidemenu_big">�]�E�T�|�[�g</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume.asp">�������̎����쐬/�t�H�[�}�b�g�̃_�E�����[�h</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_kakikata.asp">�������̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_qa.asp">������Q��A</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet.asp">�E���o�����̎����쐬/�t�H�[�}�b�g�̃_�E�����[�h</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet_kakikata_1.asp">�E���o�����̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_taishokunegai.asp">�ސE��̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_year_calculation.asp">����E�a��/�w���v�Z</a></li>
        <li class="sidetitle">�������E�E���o�����̈��</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/conpripromotion.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_711.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_familymart.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_lawson.png"></a></li>
        <li class="sidetitle">�]�E�T�|�[�g�ē�</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_introduce.asp">�l�ޏЉ�</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temporary.asp">�l�ޔh��</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temptoperm.asp">�Љ�\��h��</a></li>
		<!--<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/careerconsultation/">�L�����A���k</a></li>-->
	
		</ul>
        
        <%
ElseIf iTabIndexType = 14 Then
			Response.Write sHTML10
			Response.Write sHTML12		
			
				%>
			
       <ul>
		<li class="sidemenu_big">�]�E�T�|�[�g</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume.asp">�������̎����쐬/�t�H�[�}�b�g�̃_�E�����[�h</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_kakikata.asp">�������̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_qa.asp">������Q��A</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet.asp">�E���o�����̎����쐬/�t�H�[�}�b�g�̃_�E�����[�h</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet_kakikata_1.asp">�E���o�����̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_taishokunegai.asp">�ސE��̏�����</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_year_calculation.asp">����E�a��/�w���v�Z</a></li>
        <li class="sidetitle">�������E�E���o�����̈��</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/conpripromotion.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_711.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_familymart.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_lawson.png"></a></li>
        <li class="sidetitle">�]�E�T�|�[�g�ē�</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_introduce.asp">�l�ޏЉ�</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temporary.asp">�l�ޔh��</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temptoperm.asp">�Љ�\��h��</a></li>
		<!--<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/careerconsultation/">�L�����A���k</a></li>-->
        <li class="sidetitle">�}�b�v</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>type_map.asp">�E��E�Ǝ�ʃ}�b�v</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>area_map.asp">�n��ʃ}�b�v</a></li>	
        <li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>keyword_map.asp">�L�[���[�h�}�b�v</a></li>	
		</ul>

			<%	
			
				
		End If

		Response.Write "</nav>"
	End If
End Function
%>

