<%
'******************************************************************************
'�T�@�v�F�w�b�_�[
'���@���FHeadType	0�y�g�b�v�z1�y���E�ҁz2�y��Ɓz3�y���p�z4�y�㗝�X�z
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
	Dim flgQE,oRS,sSQL,sError

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

	Response.Write "<a name=""pagetop""></a>" & vbCrLf

	sHeadcmt = getHeaderText(HeadType,G_URL)

	If HeadType = 0 Then '�g�b�v
		sLinkurl = "/company/index.asp"
		sLinkalt = "���l�L�������ƃi�r"
		sLinktext = "�̗p�S���i���l��Ɓj�l�͂�����"
	ElseIf HeadType = 1 Then '���E��
		sLinkurl = "/company/index.asp"
		sLinkalt = "���l�L�������ƃi�r"
		sLinktext = "�̗p�S���i���l��Ɓj�l�͂�����"
	ElseIf HeadType = 2 Or HeadType = 4 Then '���
		sLinkurl = "/"
		sLinkalt = "�]�E�E���l�T�C�g�����ƃi�r"
		sLinktext = "���d�������T���̕��͂�����"
	ElseIf HeadType = 3 Then '���p
		sLinkurl = "/company/index.asp"
		sLinkalt = "���l�L�������ƃi�r"
		sLinktext = "�̗p�S���i���l��Ɓj�l�͂�����"
	End If



	'<�X�}�[�g�t�H�����[�U�����̂����ƃi�r���o�C���ւ̗U���o�i�[�\��>
	If chkSmartPhone(G_USERAGENT) = True Then
		'Response.Write "<a href=""" & HTTPS_NAVI_MOBILE & "?an=spbanner""><img src=""/img/banner/smartphone_banner.png"" alt=""�X�}�[�g�t�H���̕��̓R�R���^�b�`�I�����ƃi�r���o�C��"" border=""0""></a>"
        Response.Write "<div style=""padding:15px;line-height:2em;font-size:xx-large;"">"
        Response.Write "<a href=""http://sp.shigotonavi.jp/"" border=""0""><img src=""/img/switch_btn_01.gif"" border=""0""></a>"
        Response.Write "<img src=""/img/switch_btn_02.gif"" border=""0"">"
        'Response.Write "PC | <a href=""http://sp.shigotonavi.jp/"">�X�}�[�g�t�H��</a>"
        Response.Write "</div>"

	End If
	'</�X�}�[�g�t�H�����[�U�����̂����ƃi�r���o�C���ւ̗U���o�i�[�\��>

%>
<div id="waku">
<header>

<div class="hblk1"></div>
<div class="lt">
<h1>�E�T�C�g�u�����ƃi�r�v�B���Ј��E�h���̋��l���͂������A�v���ɂ��M���ɓK�����]�E�T�|�[�g�����񋟂��Ă��܂��I</h1>
</div>
<div class="rt">
<a href="/staff/access.asp" class="stext"><img src="/img/top/head_icon.gif" height="10" alt="���⍇��" border="0">���⍇��</a>
<a href="/shigotonavi/sitemap.asp" class="stext">
<img src="/img/top/head_icon.gif" height="10" alt="�T�C�g�}�b�v" border="0">�T�C�g�}�b�v</a>
</div>
<br clear="all">
<div class="line1"></div>
	
<table>
<tr>


<td align="left" style="height:42px; width:141px;">
<%
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
		Response.Write "<a class=""decnone"" href=""/""><img src=""/img/top/shigotonavi_logo.gif"" alt=""�����ƃi�r"" border=""0"" align=""left"" style=""margin-left:4px;""></a>"
	End If
%>

</td>
	<!--/�w�b�_�[���F�����ƃi�r���S-->

	<!--�w�b�_�[�E-->
<td align="right" valign="bottom" style="font-size:11px;" class="topstatus">

	
    <%
	
	'<Google�̃T�C�g������>
	If Request.ServerVariables("HTTPS") <> "on" Then
		Response.Write "<form action=""/search.asp"" id=""cse-search-box"" style=""margin-left:5px;padding:0px;display:inline;"">"
		Response.Write "<div style=""display:inline;"">"
		Response.Write "<img src=""/img/top/head_icon.gif"" alt="""" border=""0"" style=""vertical-align:millde;"">"
		Response.Write "<label>"
		Response.write "<span>�T�C�g������&nbsp;&nbsp;</span>"
		Response.Write "<input type=""hidden"" name=""cx"" value=""partner-pub-2905051069986345:lub5li-izzy"">"
		Response.Write "<input type=""hidden"" name=""cof"" value=""FORID:10"">"
		Response.Write "<input type=""hidden"" name=""ie"" value=""Shift_JIS"">"
		Response.Write "<input type=""text"" name=""q"" size=""20"">"
		Response.Write "</label>"
		Response.Write "<input type=""submit"" name=""sa"" value=""&#x691c;&#x7d22;"">"
		Response.Write "</div>"
		Response.Write "</form>"
		Response.Write "<script type=""text/javascript"" src=""http://www.google.co.jp/coop/cse/brand?form=cse-search-box&amp;lang=ja""></script><br>"
	End If
	'</Google�̃T�C�g������>


	'<���l���A��Ɛ��A���E�Ґ�>
	Response.Write "<img src=""/img/top/countericon_order.gif"" alt=""���l��"" border=""0"" style=""margin:0px 2px;"">���l<span class=""cnt"">" & iOrderCnt & "</span>��&nbsp;"
	Response.Write "<img src=""/img/top/countericon_company.gif"" alt=""��Ɛ�"" border=""0"" style=""margin:0px 2px;"">���<span class=""cnt"">" & iCompanyCnt & "</span>��&nbsp;"
	Response.Write "<img src=""/img/top/countericon_staff.gif"" alt=""���E�Ґ�"" border=""0"" style=""margin:0px 2px;"">���E��<span class=""cnt"">" & iAll & "</span>�l&nbsp;"
	Response.Write "�i" & MonthName(Month(Now)) & Day(Now) & "��(" & Left(WeekdayName(Weekday(Now)),1) & ")" & "�X�V�j"
	'</���l���A��Ɛ��A���E�Ґ�>

	'�̗p�S���җl
	'Response.Write "�@<a href=""" & sLinkurl & """ style=""font-size:14px;""><img src=""/img/top/head_icon.gif"" alt=""" & sLinkalt & """ border=""0"" style=""vertical-align:middle;"">" & sLinktext & "</a>"
	'<!-- #INCLUDE FILE="../ad_banner_control/ad_banner.asp" -->
	Response.Write "</td>"
	'<�w�b�_�[�E>

	Response.Write "</tr>"

	'<�w�b�_�[�����F�w�i�΂̂��>
'	Response.Write "<tr style=""background-image:url(/img/top/headtext_background.gif);"">"
'	Response.Write "<td colspan=""2"" align=""left"" style=""margin:0px;padding:0px;color:#ffffff; height:20px;border-top:solid 1px #ffffff; border-bottom:solid 1px #ffffff;"">"
'	Response.Write sHeadcmt
'	Response.Write "</td>"
'	Response.Write "</tr>"
	'</�w�b�_�[�����F�w�i�΂̂��>

	Response.Write "</table>"
	Response.Write htmlTabIndex(Request.ServerVariables("URL"),G_USERTYPE,sHeadcmt)
	Response.Write "</header>"

	If HeadType = 9 Then
		'<�T�C�h���j���[����ver>
		Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
		Response.Write "<div align=""left"" style=""width:990px;foat:left;"">"
		Response.Write "<div class=""moji912"" style=""padding:3px 0px 0px 3px;float:left;"">" & vbCrLf
		'</�T�C�h���j���[����ver>
	Else
		Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
		Response.Write "<div align=""left"" style=""width:990px;float:left;"">" '�y�[�W�S�̂̕��ifooter�ŉ����ŕ�
		Response.Write "<div class=""moji912"" id=""main"">" & vbCrLf '���C���R���e���c���isidemenu�ŏ㕔�ŕ߁j
	End If
End Function

%><!-- #INCLUDE FILE="func/htmlNaviSideMenu.asp" --><%

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

	Response.Write "<p class=""m0"" style=""margin-top:15px;text-align:right;""><a href=""#pagetop"" class=""stext"">���y�[�WTOP��</a></p>"

	'<google�A�h�Z���X>
	'2011/06/15�`2011/06/21�̊��Ԃ̓A�h�Z���X���e�X�g�I�ɒ�~����
	'If Date < "2011/06/15" Or Date >= "2011/06/22" Then
	'2011/07/08�` �A�h�Z���X���~����
	If Date < "2011/07/08" Then
		Response.Write "<div style=""margin-bottom:10px;"">"
%><!-- #INCLUDE VIRTUAL="/include/ads/navifooter.asp" --><%
		Response.Write "</div>"
	End If
	'</google�A�h�Z���X>


	'�����ƃi�r���o�C���̏Љ�i�g�т̃A�h���X�o�^�҂̂݁j
	Server.Execute("/include/mobilesiteinfo.asp")

	Response.Write "<div id=""footer"">"

	Response.Write "<ul>"
	Response.Write "<li class=""ttl"">�]�E�T�C�g�u�����ƃi�r�v�ɂ���</li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & """ class=""topdecnone"">�]�E�T�C�g�u�����ƃi�r�vHOME</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "infomation/info.asp"" class=""topdecnone"">�L���q��</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "lis/lis.asp"" class=""topdecnone"">�^�c��ЁE���Ѝ̗p���</a></li>"
'	Response.Write "<li><a href=""/staff/s_aboutnavi.asp"" class=""topdecnone"">�����p�K�C�h</a></li>"
'	Response.Write "<li><a href=""/staff/qa.asp"" class=""topdecnone"">�p���`</a></li>"
'	Response.Write "<li><a href=""/staff/s_kiyaku.asp"" class=""topdecnone"">���p�K��</a></li>"
'	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/s_books.asp"" class=""topdecnone"">�]�E�ɖ𗧂{</a></li>"
'	Response.Write "<li><a href=""/link.asp"" class=""topdecnone"">�����N�|���V�[</a></li>"
'	Response.Write "<li><a href=""/link_collection.asp"" class=""topdecnone"">���𗧂����I�����N�W</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "shigotonavi/sitemap.asp"" class=""topdecnone"">�T�C�g�}�b�v</a></li>"
'	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/Ranking.asp"" class=""topdecnone"">���E�҃����L���O</a></li>"
	Response.Write "</ul>"

	Response.Write "<ul>"
	Response.Write "<li class=""ttl"">�]�E�����l���̋��E�җl</li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"" class=""topdecnone"">���l��T��</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/s_resume.asp"" class=""topdecnone"">�������̎����쐬�c�[��</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/s_resume_kakikata.asp"" class=""topdecnone"">�������̏�����</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/s_jikopr.asp"" class=""topdecnone"">���Ȃo�q���[�J�[</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/motive_index.asp"" class=""topdecnone"">�u�]���@���[�J�[</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/s_careersheet.asp"" class=""topdecnone"">�E���o�����̎����쐬�c�[��</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/s_careersheet_kakikata_1.asp"" class=""topdecnone"">�E���o�����̏�����</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/s_mynavi.asp"" class=""topdecnone"">�K�E�f�f�u���Ԃ�i�r�v</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/s_temporary.asp"" class=""topdecnone"">�l�ޔh��</a>�b<a href=""" & HTTP_CURRENTURL & "s_contents/s_introduce.asp"" class=""topdecnone"">�l�ޏЉ�</a>�b<a href=""" & HTTP_CURRENTURL & "s_contents/s_temptoperm.asp"" class=""topdecnone"">�Љ�\��h��</a></li>"
	Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "staff/access.asp"" class=""topdecnone"">���⍇��</a></li>"
	Response.Write "</ul>"

	Response.Write "<ul>"
	Response.Write "<li class=""ttl"">���W</li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/ad/0001/"" class=""topdecnone"">SE�]�E</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/tg/0004/"" class=""topdecnone"">�Տ������Z�t�̋��l</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/tg/0005/"" class=""topdecnone"">�p����������Ĕh���œ���</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/sz/0001/"" class=""topdecnone"">�É��œ]�E!!</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/ng/0002/"" class=""topdecnone"">���É��̔h��</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/or/0001/"" class=""topdecnone"">DTP�I�y���[�^�[�E�f�U�C�i�[���l</a></li>"
	If Now <= "2011/09/15 12:00:00" Then
		'<�L�����y�[��>
		Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "campaign/2011090101/"" target=""_blank"" class=""topdecnone"" style=""font-size:95%;"">���R����!�c�ƐE�̓]�E�x����������߰�</a></li>"
		'</�L�����y�[��>
	Else
		Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/oy/0001/"" class=""topdecnone"">���R�̋��l</a></li>"
	End if
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/hr/0001/"" class=""topdecnone"">�L���œ]�E!!</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/license/1700101.asp"" class=""topdecnone"">��n���������C�� ���l</a></li>"
	Response.Write "</ul>"

	Response.Write "<ul style=""margin-right:0px;"">"
	Response.Write "<li class=""ttl"">���l��Ɨl</li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "tab/index5.asp"" class=""topdecnone"">�̗p���S���җl</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "company/c_hajime.asp"" class=""topdecnone"">���l�L��</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp"" class=""topdecnone"">�����ƃi�r���E�҂ƌf�ڊ�ƃf�[�^</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"" class=""topdecnone"">�l�ޔh��</a>�b<a href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"" class=""topdecnone"">�l�ޏЉ�</a>�b<a href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"" class=""topdecnone"">�Љ�\��h��</a></li>"
	Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "company/access.asp"" class=""topdecnone"">���⍇��</a></li>"
'	Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "company/access.asp"" class=""topdecnone"">�L���㗝�X�̕��̂��⍇��</a></li>"
'	Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "company/fc_index.asp" class="topdecnone">�����ƃi�rFC</a></li>"
	Response.Write "</ul>"

	Response.Write "<br clear=""all"">"

	Response.Write "<div style=""text-align:center;"">"
	Response.Write "<a href=""" & HTTP_LIS_CURRENTURL & """ target=""_blank""><img src=""/img/footer/footer_lis_logo_1.gif"" alt=""�]�E�T�C�g������ƃi�r��^�c-���X�������-"" border=""0""></a>"
	Response.Write "</div>"

	Response.Write "</div>"
	Response.Write "</div>"
	Response.Write "</div>" & vbCrLf

	'<Twitter�o�b�W>
'	Select Case getTabIndexType(Request.ServerVariables("URL"))
'		Case 0,1,2,3,4,6: Response.Write scrTwitterFollowBadge()
'		Case 5,7: Response.Write scrIntroTwitterFollowBadge()
'	End Select
	'</Twitter�o�b�W>

	'<analytics>
	If Request.ServerVariables("SERVER_NAME") = "www.shigotonavi.co.jp" And InStr(Request.ServerVariables("REMOTE_HOST"),"192.168.") = 0 Then
		Response.Write "<script src="""
		If Request.ServerVariables("HTTPS") = "off" Then
			Response.Write "http://www.google-analytics.com/urchin.js"
		Else
			Response.Write "https://ssl.google-analytics.com/urchin.js"
		End If
		Response.Write """ type=""text/javascript""></script>"
		Response.Write "<script type=""text/javascript"">"
		Response.Write "_uacct = ""UA-2265459-3"";"
		Response.Write "urchinTracker();"
		Response.Write "</script>" & vbCrLf
	End If
	'</analytics>

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
	Response.Write "<a href=""" & HTTP_CURRENTURL & "cafe/cafe_list.asp""><img src=""/img/rightmenu/navicafe_banner_top.jpg"" alt=""�i�r�J�t�F"" border=""0"" style=""margin:0px;padding:0px;""></a>"
	Response.Write "<div style=""margin-top:0px;padding:14px 6px 0px 8px;font-size:10px;line-height:15px;"">"

	'** TOP 08/11/05 Lis�� ADD
	'���݌f�ڒ���TOP3�̃g�s
	sSQLnsr = "up_GetData_NC_Topic '','','','1','3'"
	flgQEnsr = QUERYEXE(dbconn, oRSnsr, sSQLnsr, sErrornsr)
	Do While GetRSState(oRSnsr) = True
		Response.Write "<a href=""" & HTTP_CURRENTURL & "cafe/cafe_detail.asp?t=" & oRSnsr.Collect("TopicID")
		Response.Write """>�E"
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
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/s_aboutnavi.asp"">�����p�K�C�h</a></li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/qa.asp"">�p���`</a></li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/s_searchexplanation.asp"">���d���������@</a></li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/s_kiyaku.asp"">���p�K��</a></li>"
		Response.Write "<li class=""rightmenu_end""><a href=""" & HTTPS_CURRENTURL & "staff/access.asp"">���⍇��(���E�Ґ�p)</a></li>"
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
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_mensetsu_index.asp"">�ʐڑ΍�</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "column/column_1.asp"">�h���Ј�<span class=""stext"">-�����̌��̓v���ӎ�</span></a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_kyuuyomeisai.asp"">���Ȃ��̋��^����</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_ready.asp"">�]�E�̐S�\��</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_proce.asp"">�]�E�ɕK�v�Ȏ葱��</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"">���i���t�o�}�j���A��</a></li>"
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
End Function
%>
