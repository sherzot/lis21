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




%>
<!-- Google Tag Manager -->
<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
})(window,document,'script','dataLayer','GTM-PG92H5L');</script>
<!-- End Google Tag Manager -->

<div id="smartMenu" style="display:none;">
	<a id="smartLogo" href="/search/"><img src="/img/smart/smartLogo.png" alt="�����ƃi�r" border="0"></a>
	<a href="/order/order_search_detail.asp" id="smartSearch">�T��</a>
    <div id="smartButton">
    </div>


</div><!--smartMenu-->
<div id="smartPhoneNavi" style="display:none;">
<% If G_USERTYPE = "staff" Then %>
	<h3>My���j���[</h3>
    <ul>
    	<li class="topNavi"><a href="/staff/s_login.asp">My�y�[�W</a></li>
		<li class="topNavi"><a href="/staff/person_detail.asp">�v���t�B�[���Ǘ�</a></li>
	<h3>�]�E�T�|�[�g</h3>
        <li class="topNavi"><a href="/staff/my_footprint.asp">�{������</a></li>
        <li class="topNavi"><a href="/staff/watchlist.asp">���C�ɓ��胊�X�g</a></li>
		<li class="topNavi"><a href="/staff/edit_list.asp">����ꗗ</a></li>
        <li class="topNavi"><a href="/staff/mailhistory_person.asp">���[���Ǘ�</a></li>
        <li class="topNavi"><a href="/staff/schedule/">�X�P�W���[���Ǘ�</a></li>
        
	<h3>�o�����[�I�t�@�[</h3>
        <li class="topNavi"><a href="/staff/step2a.asp">��]��������</a></li>
        <li class="topNavi"><a href="/staff/step2a.asp">��Ƃ���̎���</a></li>
	<h3>�������E�E���o�����̍쐬</h3>
		<li class="topNavi"><a href="/staff/resume_print.asp">�������E�E���o�������</a></li>
		<li class="topNavi"><a href="/staff/resume_picture.asp">�������p�ʐ^�o�^</a></li>
	<h3>�e��ݒ�</h3>
        <%'<li class="topNavi"><a href="/staff/searchordercondition/">���������Ǘ�</a></li>%>
		<li class="topNavi"><a href="/staff/notification_mail_service.asp">�X�P�W���[���ʒm</a></li>
		<li class="topNavi"><a href="/staff/changepassword.asp">�p�X���[�h�̕ύX</a></li>
		<li class="topNavi"><a href="/suspension/questionnarie.asp">�x�~�E�މ�</a></li>


        <% '�Ȃɂ���if��
        If G_SSLFLAG = False Then %>
        <li class="topNavi"><a href="/logout.asp">���O�A�E�g</a></li>
        <% ELSE %>
        <li class="topNavi"><a href="/logout.asp">���O�A�E�g</a></li>
        <% END IF %>	
    	
    </ul>
    
<% End If %>
	<h3>���C�����j���[</h3>
	<nav>
    	<ul>
            <li class="topNavi"><a href="/search/">�����Ƃ�T��</a></li>
            <li class="topNavi"><a href="/koryu/">��</a></li>
            <li class="topNavi"><a href="/manabu/">�w��</a></li>
            <li class="topNavi"><a href="/link/">�����N</a></li>
        </ul>
    </nav>
    <h3>���߂Ă̕���</h3>
    <nav>
    	<ul>
            <li class="topNavi"><a href="/tab/index1.asp">���߂Ă̕���</a></li>
            <li class="topNavi"><a href="/valueoffer/">�]�E�̐V�X�^�C���u�o�����[�I�t�@�[�v</a></li>
            <!--<li class="topNavi"><a href="/valueoffer/persona.asp">�u�o�����[�I�t�@�[����`�����D�� �ҁ`�v</a></li>-->
            <%'<li class="topNavi"><a href="/neo/howabout/">�]�E�̐V�X�^�C���u�G�[�W�F���gNEO�v</a></li>%>
        </ul>
    </nav>
    
    <br clear="both">
    <div id="smartNaviClose">
    	�~CLOSE
    </div>
</div><!--/smartPhoneNavi-->

<div id="fb-root"></div>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = "//connect.facebook.net/ja_JP/sdk.js#xfbml=1&version=v2.3";
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));</script>


<%  '2015/08/27 �X�}�z�����O�C���ς݂̎��w�b�_�[�̈ꕔ���\��
    '���F�X�ȃX�N���v�g�̓���ɗ��ނ��ߒ��~
    'If chkSmartPhone(G_USERAGENT) = True and G_USERTYPE = "staff" Then
    'Response.Write "<div style=""display:none;"">" else %>
<div id="header_waku">
    <%' end if %>
	
    <div id="maku">    
    </div>
<header id="pagetop">


<div class="lt" id="top">
<a class="decnone" href="/"><img src="/img/top/logo.gif" alt="�����ƃi�r" border="0" align="left" style="margin-left:4px;"></a>
<br>
<p>�͂��炭�l�̃\�[�V�����R�~���j�e�B�[</p>
</div>

<div id="neoBanner">
    <!--<a href="/valueoffer/" id="toC">���E�җl</a>-->
    <a href="/lis/lis_group.asp" id="toA">�p�[�g�i�[��Ɨl</a>
 

</div>

<div class="rt">

<%

If G_USERTYPE = "" Then
	if GetForm("ordercode", 2) <> "" then
		if IsRE(Trim(Replace(Server.HTMLEncode(GetForm("ordercode", 2)), "'", "�f")), "^J\d\d\d\d\d\d\d$", True) = True then
			response.write "<a href=""/staff/person_reg1.asp?ordercode=" & GetForm("ordercode", 2) & """target=""_self"" id=""reg_new"">����o�^</a>"
		else
			response.write "<a href=""/staff/person_reg1.asp"" target=""_self"" id=""reg_new"">����o�^</a>"
		end if
	else
		response.write "<a href=""/staff/person_reg1.asp"" target=""_self"" id=""reg_new"">����o�^</a>"
	end if


	    response.write "<a href=""/login_menu.asp"" target=""_self"" id=""login"">���O�C��</a>"

ElseIf G_USERTYPE = "staff" Then
	response.write "<a href=""/staff/s_login.asp"" target=""_self"" id=""s_mypage"">My �߰��</a>"

Else

End If
%>
<!--<a href="/staff/access.asp" class="stext"><img src="/img/top/head_icon.gif" height="10" alt="���⍇��" border="0">���⍇��</a>
<a href="/shigotonavi/sitemap.asp" class="stext">
<img src="/img/top/head_icon.gif" height="10" alt="�T�C�g�}�b�v" border="0">�T�C�g�}�b�v</a>
--></div>

<br clear="all">
<div style="position:absolute; right:25px; top:37px;">

<!-- #include file="../../caution.html" -->
</div>






<div id="number">
<%


	'<���l���A��Ɛ��A���E�Ґ�>

	Response.Write "���l<span class=""cnt"">" & iOrderCnt & "</span>��&nbsp;"
	Response.Write "���<span class=""cnt"">" & iCompanyCnt & "</span>��&nbsp;"
	Response.Write "���E��<span class=""cnt"">" & iAll & "</span>�l&nbsp;"
	Response.Write "�i" & MonthName(Month(Now)) & Day(Now) & "��(" & Left(WeekdayName(Weekday(Now)),1) & ")" & "�X�V�j</div>"
	'</���l���A��Ɛ��A���E�Ґ�>
	    
%>

<BR>
<div class="campaign" style="text-align:center;display:block;">
<strong style="color:#009900;font-size:250%;line-height:1.1em;text-align:center;display:inline;">���������܂Łw�����ƃi�r�x�o�^�҂��P�O�O���l�˔j���܂����B</strong>  
</div>



<div class="notice">
�y�����ƃi�r����̂��m�点�z�u�V�^�R���i�E�C���X�v�����g��̗\�h�΍�Ƃ��āA���d�b�܂���WEB�ʒk(Zoom��Skype��)�𗘗p���Ă̔�Ζʎ��ł́A<br> �]�E�̑��k�����Ή��\�ł��B���C�y�ɂ��₢���킹���������B

</div>


<%
	
If HeadType = 0 Then	
%>
<div id="img_map">

    <div id="comment_sagasu">
        <h4 class="center">�u�����Ƃ�T���v�Ƃ�</h4>
        ���l��񌟍��◚�����̎����쐬<br>
        �Ȃǂ��ł���A�K�E�ɏA�����߂̊�{�I�ȓ]�E���ł��B
    
    </div>
    
    <div id="comment_koryu">
        <h4 class="center">�u�𗬁v�Ƃ�</h4>
        SNS�Ŏd����]�E�Ɋւ������m���A�A�h�o�C�X�Ȃǂ������郆�[�U�[�Q���^�̌𗬍L��ł��B
        
    
    </div>
    
    <div id="comment_manabu">
        <h4 class="center">�u�w�ԁv�Ƃ�</h4>
        �X�L���A�b�v��r�W�l�X�R�����A<br>
        ���ȕ��́A�]�E�m�E�n�E�Ȃǂ��w�ׂ�R�[�i�[�ł��B
    
    
    </div>
    
    <div id="comment_link">
        <h4 class="center">�u�����N�v�Ƃ�</h4>
        ���Ђ��^�c����֘A�T�C�g��A<br>
        �֘A��񂪏[���̂��𗧂��R�[�i�[�ł��B
    
    
    </div>
    
    <div id="comment_bes">
        <h4 class="center">����o�^�i�����j �������</h4>
        �������E�o�����̎����쐬�⎩�ȕ��̓c�[���Ȃ�
        �]�E�ɖ𗧂��܂��܂ȃR���e���c���g���܂��B
        ��Ƃ���̃X�J�E�g���[�����󂯎�邱�Ƃ��ł��܂��I
    
    
    </div>


</div><!--/img_map-->





<%

End If

	Response.Write "</header>"
	Response.Write "</div><!--/#header_waku-->"
	
	Response.Write htmlTabIndex(Request.ServerVariables("URL"),G_USERTYPE,sHeadcmt)
	
	If HeadType = 0 Then
	
	
		
	%>


<!--
�V�o�[�W����
<div id="top_contents_waku">
	<div class="samune" id="sa1"><a href="/search/index.asp"><img src="/img/top/top_samune_search.png"></a></div>
    <div class="samune" id="sa5"><a href="/iphone/index.html" target="_blank"><img src="/img/top/top_samune_iphone.png"></a></div>
    <div class="samune" id="sa3"><a href="/neo/oiwai/index.asp"><img src="/img/top/top_samune_oiwai.png"></a></div>
    
    <div class="samune" id="sa4"><a href="/company/access.asp"><img src="/img/top/top_samune_contact.png"></a></div>

<div id="top_contents">

<div id="topKokoku">
	<div id="kokokuWaku">
        <div>
            <img src="https://www-b1.shigotonavi.co.jp/company/imgdsp.asp?companycode=C0018268&optionno=10">
            <p>������Ё@NEO</p>
            <p>���o�C���[���E�u���[�h�o���h�T�[�r�X�̒�āE�̔�</p>
            <p>�����s</p>
        </div>
        
        <div>
            <img src="https://www-b1.shigotonavi.co.jp/company/imgdsp.asp?companycode=C0018268&optionno=10">
            <p>������Ё@NEO</p>
            <p>���o�C���[���E�u���[�h�o���h�T�[�r�X�̒�āE�̔�</p>
            <p>�����s</p>
        </div>
        
        <div>
            <img src="https://www-b1.shigotonavi.co.jp/company/imgdsp.asp?companycode=C0018268&optionno=10">
            <p>������Ё@NEO</p>
            <p>���o�C���[���E�u���[�h�o���h�T�[�r�X�̒�āE�̔�</p>
            <p>�����s</p>
        </div>
        
        <div>
            <img src="https://www-b1.shigotonavi.co.jp/company/imgdsp.asp?companycode=C0018268&optionno=10">
            <p>������Ё@NEO</p>
            <p>���o�C���[���E�u���[�h�o���h�T�[�r�X�̒�āE�̔�</p>
            <p>�����s</p>
        </div>
	</div>
</div>

-->

<div id="top_contents_waku">
    <div class="samune" id="sa1"><a href="https://www.shigotonavi.co.jp/order/order_detail.asp?OrderCode=J0110872"><img src="/img/top/top_samune_shokairecruit.png"></a></div>
    <div class="samune" id="sa2"><a href="/search/index.asp"><img src="/img/top/top_samune_search.png"></a></div>
    <div class="samune" id="sa3"><a href="https://www.shigotonavi.co.jp/order/order_detail.asp?OrderCode=J0111745"><img src="/img/top/top_samune_SErecruit.png"></a></div>
    <div class="samune" id="sa4"><a href="/point/pr/"><img src="/img/top/top_samune_oiwai.png"></a></div>


<div id="top_contents">
	<a href="https://www.shigotonavi.co.jp/order/order_detail.asp?OrderCode=J0110872"><img src="/img/top_contents/shokairecruit.png"></a>
</div>
	
</div>


<br>

<div style="width:990px;margin:0 auto 20px;padding:20px 0 3px 0;border-top:1px solid #3e3e3e;box-sizing: border-box;" class="smartNone">


	<!--<a href="/promotion/s_conpri_riyou.asp"><img src="/img/top/shConpri2.png"></a><br>-->
    <p style="margin:0 auto;text-align:center;font-size:27px;vertical-align:middle;line-height:32px;font-weight:bold;border:0px solid #000;">
        �����ƃi�r�̗������R���r�j�v�����g.
    </p><br />
    
    <div style="text-align:center;margin:0 auto;background:#fff;">
    <!--<hr style="width:20px;border:1px solid #000;">-->
        <div style="padding:10px 0;background:#000;">
        <a href="/promotion/conpri_riyou.asp">  <img style="width:150px;border:0px solid #000;" src="/img/top/clogo_711.png"></a>
        <a href="/promotion/s_conpri_riyou.asp"><img style="width:150px;border:0px solid #000;" src="/img/top/clogo_familymart.png"></a>
        <a href="/promotion/s_conpri_riyou.asp"><img style="width:150px;border:0px solid #000;" src="/img/top/clogo_lawson.png"></a>
        </div>

    <p style="font-weight:bold;margin-top:12px;">�T�[�r�X�̂����p���\�ȃR���r�j�͂�����i�܏\�����j</p>
    </div>

</div>


    <!-- #INCLUDE VIRTUAL="/attention.html" -->
   <!-- <div id="mailReg">
    	<form>
        	<input type="text" value="������mail">
            <input type="button" value="�o�^" onClick="location.href='/staff/mailReg.asp'">
        </form>
    </div>-->
    <%
	End If
	



	response.write "<section id=""waku"">"


	If HeadType = 9 Then
		'<�T�C�h���j���[����ver>
		'Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
		'Response.Write "<div align=""left"" style=""width:990px;foat:left;"">"
		Response.Write "<div class=""moji912"" style=""padding:3px 0px 0px 3px;float:left;"">" & vbCrLf
		'</�T�C�h���j���[����ver>
	Else
		'Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
		'Response.Write "<div align=""left"" style=""width:990px;float:left;"">" '�y�[�W�S�̂̕��ifooter�ŉ����ŕ�
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
	'Response.Write "</div>"
	If 1 = 2 Then
		Response.Write "<div style=""width:200px;float:right;margin-top:0px;"">"
		If Request.ServerVariables("URL") <> "/search.asp" Then
			Call NaviSidemenuRight()
		End If
		Response.Write "</div>"
	End If
	'Response.Write "</div>"
	Response.Write "<br clear=""all"">"
	
	Response.Write "<p class=""m0"" style=""margin-top:15px;text-align:right;""><a href=""#pagetop"" class=""stext_bottom"">���y�[�WTOP��</a></p>"
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
	'Server.Execute("/include/mobilesiteinfo.asp")
%>
	
</section>
<footer>

<!-- Google Tag Manager (noscript) -->
<noscript><iframe src="https://www.googletagmanager.com/ns.html?id=GTM-PG92H5L"
height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<!-- End Google Tag Manager (noscript) -->

<div id="foot_child">
	<ul>
	<li class="ttl">�u�����ƃi�r�v�ɂ���</li>
	<li><a href="<%= HTTP_CURRENTURL %> " class="topdecnone">�����ƃi�rHOME</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>tab/index1.asp" class="topdecnone">�͂��߂Ă̕���</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>search/" class="topdecnone">�����Ƃ�T���G���A</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>koryu/" class="topdecnone">�𗬃G���A</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>manabu/" class="topdecnone">�w�ԃG���A</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>link/" class="topdecnone">�����N�G���A</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>support/" class="topdecnone">�]�E�T�|�[�g</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>staff/ranking_index.asp" class="topdecnone">�����ƃi�r�����L���O</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>/staff/s_aboutnavi.asp" class="topdecnone">�����p�K�C�h</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>lis/lis.asp" class="topdecnone">�^�c��Ђɂ���</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>recruit/" class="topdecnone">�̗p���</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>shigotonavi/sitemap.asp" class="topdecnone">�T�C�g�}�b�v</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>privacy/privacymark.asp" class="topdecnone">P�}�[�N�ɂ���</a></li>
	</ul>

	<ul>
	<li class="ttl">���E�җl</li>
	<li><a href="<%= HTTP_CURRENTURL %>order/order_search_detail.asp" class="topdecnone">���l��T��</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>staff/s_resume.asp" class="topdecnone">�������̎����쐬/�t�H�[�}�b�g��<br>&nbsp�_�E�����[�h</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>staff/s_resume_kakikata.asp" class="topdecnone">�������̏�����</a></li>
    	<li><a href="<%= HTTPS_CURRENTURL %>column/column_index.asp" class="topdecnone">�]�E�E�A�E�R����</a></li>
    	<li><a href="<%= HTTPS_CURRENTURL %>type_map.asp" class="topdecnone">�E��Ǝ�ʃ}�b�v</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>s_contents/s_jikopr.asp" class="topdecnone">����PR���[�J�[</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>s_contents/motive_index.asp" class="topdecnone">�u�]���@���[�J�[</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet.asp" class="topdecnone">�E���o�����̎����쐬/�t�H�[�}�b�g��<br>&nbsp�_�E�����[�h</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet_kakikata_1.asp" class="topdecnone">�E���o�����̏�����</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>s_contents/s_mynavi.asp" class="topdecnone">�K�E�f�f�u���Ԃ�i�r�v</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>s_contents/s_temporary.asp" class="topdecnone">�l�ޔh��</a>�b<a href="<%= HTTPS_CURRENTURL %>s_contents/s_introduce.asp" class="topdecnone">�l�ޏЉ�</a>�b<a href="<%= HTTPS_CURRENTURL %>s_contents/s_temptoperm.asp" class="topdecnone">�Љ�\��h��</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>staff/access.asp" class="topdecnone">���⍇��</a></li>
	</ul>

	<ul>
	<li class="ttl">���W</li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/ad/0001/" class="topdecnone">SE�̓]�E���W</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/tg/0004/" class="topdecnone">�Տ������Z�t�̋��l</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/tg/0005/" class="topdecnone">�p����������Ĕh���œ���</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/or/0001/" class="topdecnone">DTP�E�f�U�C�i�[�̋��l</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>s_contents/license/1700101.asp" class="topdecnone">��n���������C�҂̋��l</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/tg/0006/index.asp" class="topdecnone">�N��1000���~�N���X�̓]�E</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/tk/0001/index.asp" class="topdecnone">���z�E�s���Y�ƊE���W</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/ng000001.asp" class="topdecnone">��ƊŌ�t�̋��l���W</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/tokyo.asp" class="topdecnone">�����̓]�E�E�A�E</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/sz/0001/" class="topdecnone">�É��̓]�E</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/ng/0002/" class="topdecnone">���É��̓]�E</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/oy/0001/" class="topdecnone">���R�̓]�E</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/hr/0001/" class="topdecnone">�L���̓]�E</a></li>
	</ul>

	<ul style="margin-right:0px;">
	<li class="ttl">�̗p��Ɨl</li>
    <li><a href="<%= HTTP_CURRENTURL %>neo/shoukai/" class="topdecnone">�̗p��ƃg�b�v</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>company/" class="topdecnone">�u�G�[�W�F���gNEO�v�ɂ���</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>company/about.asp" class="topdecnone">�����ƃi�r�̓��F</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>company/c_staffdata.asp" class="topdecnone">���E�ҏW�v�f�[�^</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>company/c_dispatch.asp" class="topdecnone">�l�ޔh��</a>�b<a href="<%= HTTP_CURRENTURL %>company/c_introduce.asp" class="topdecnone">�l�ޏЉ�</a>�b<a href="<%= HTTP_CURRENTURL %>company/c_temptoperm.asp" class="topdecnone">�Љ�\��h��</a></li>
    <!--<li><a href="<%= HTTPS_CURRENTURL %>neo/kokoku/advertisement.asp" class="topdecnone">���l�L���̂��\����</a></li>-->
    <li><a href="<%= HTTPS_CURRENTURL %>neo/shoukai/index.asp" class="topdecnone">�l�ޏЉ�̂��\����</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>company/research.asp" class="topdecnone">�̗p���@�f�f</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>company/access.asp" class="topdecnone">���⍇��</a></li>

    <li><a href="<%= HTTPS_CURRENTURL %>neo/TempRegist/TempRegistEdit_AD.asp" class="topdecnone">���l�L���T�[�r�X�ɂ���</a></li>
	</ul>

	<br clear="all">

	<div style="text-align:center;">
	<a href="http://tekiseika.jp/job-offering/" target="_blank"><img src="/img/tekiseika_job-offering2.jpg" alt="���l�҂̊F����" border="0" style="margin-top:3px;"></a>
	<a href="http://tekiseika.jp/job-applicant/" target="_blank"><img src="/img/tekiseika_job-applicant2.jpg" alt="���E�҂̊F����" border="0" style="margin-top:3px;"></a>
	</div>
	
	<div style="text-align:center;">
	<a href="<%= HTTP_LIS_CURRENTURL %>" target="_blank"><img src="/img/footer/footer_lis_logo_1.gif" alt="�]�E�T�C�g������ƃi�r��^�c-���X�������-" border="0"></a>
	</div>
	
    <div id="smartFooter" style="display:none;">
    
    
    	<p>CopyRights(c)LIS co.,ltd.</p>
    </div>
</div>
	</footer>

<%    
    	'<�X�}�[�g�t�H�����[�U�����̂����ƃi�r���o�C���ւ̗U���o�i�[�\��>
'	If chkSmartPhone(G_USERAGENT) = True Then
'		'Response.Write "<a href=""" & HTTPS_NAVI_MOBILE & "?an=spbanner""><img src=""/img/banner/smartphone_banner.png"" alt=""�X�}�[�g�t�H���̕��̓R�R���^�b�`�I�����ƃi�r���o�C��"" border=""0""></a>"
'        Response.Write "<div style=""padding:15px;line-height:2em;font-size:xx-large;"">"
'        Response.Write "<a href=""http://sp.shigotonavi.jp/"" border=""0""><img src=""/img/switch_btn_01.gif"" border=""0""></a>"
'        Response.Write "<img src=""/img/switch_btn_02.gif"" border=""0"">"
'        'Response.Write "PC | <a href=""http://sp.shigotonavi.jp/"">�X�}�[�g�t�H��</a>"
'        Response.Write "</div>"
'
'	End If
	'</�X�}�[�g�t�H�����[�U�����̂����ƃi�r���o�C���ւ̗U���o�i�[�\��>
%>
<% If Request.ServerVariables("SERVER_NAME") = "www.shigotonavi.co.jp" And InStr(Request.ServerVariables("REMOTE_HOST"),"192.168.") = 0 Then %>
<script>
    (function (i, s, o, g, r, a, m) {
        i['GoogleAnalyticsObject'] = r; i[r] = i[r] || function () {
            (i[r].q = i[r].q || []).push(arguments)
        }, i[r].l = 1 * new Date(); a = s.createElement(o),
        m = s.getElementsByTagName(o)[0]; a.async = 1; a.src = g; m.parentNode.insertBefore(a, m)
    })(window, document, 'script', '//www.google-analytics.com/analytics.js', 'ga');

    ga('create', 'UA-2265459-3', 'auto');
    if (location.href.substring(1, 5) == 'https') {
        ga('set', 'forceSSL', true);
    }
    ga('require', 'displayfeatures');
    if (location.href.indexOf('person_registed.asp') != -1) {
        var StaffCode = '<%= Session("userid") %>';
        ga('set', 'dimension1', StaffCode);
    }
    if (location.href.indexOf('s_login.asp') != -1) {
        var StaffCode = '<%= Session("userid") %>';
        ga('set', 'dimension2', StaffCode);
    }
    ga('send', 'pageview');
</script>
<script type="text/javascript">
/* <![CDATA[ */
var google_conversion_id = 1070319369;
var google_custom_params = window.google_tag_params;
var google_remarketing_only = true;
/* ]]> */
</script>
<script type="text/javascript" src="//www.googleadservices.com/pagead/conversion.js">
</script>
<noscript>
<div style="display:inline;">
<img height="1" width="1" style="border-style:none;" alt="" src="//googleads.g.doubleclick.net/pagead/viewthroughconversion/1070319369/?value=0&amp;guid=ON&amp;script=0"/>
</div>
</noscript>
<script type="text/javascript" language="javascript">
/* <![CDATA[ */
var yahoo_retargeting_id = 'ZDIA65ITG8';
var yahoo_retargeting_label = '';
/* ]]> */
</script>
<script type="text/javascript" language="javascript" src="//b92.yahoo.co.jp/js/s_retargeting.js"></script>

<% '2017/08/22 YSS�p���}�[�P�e�B���O�^�O�ǉ� %>
<!-- Yahoo Code for your Target List -->
<script type="text/javascript">
/* <![CDATA[ */
var yahoo_ss_retargeting_id = 1000012858;
var yahoo_sstag_custom_params = window.yahoo_sstag_params;
var yahoo_ss_retargeting = true;
/* ]]> */
</script>
<script type="text/javascript" src="https://s.yimg.jp/images/listing/tool/cv/conversion.js">
</script>
<noscript>
<div style="display:inline;">
<img height="1" width="1" style="border-style:none;" alt="" src="https://b97.yahoo.co.jp/pagead/conversion/1000012858/?guid=ON&script=0&disvt=false"/>
</div>
</noscript>
<% '2017/08/22 YSS�p���}�[�P�e�B���O�^�O�ǉ� %>

<% End If %>

<!--<div id="footer_border"></div>-->

<!--logicad-->
<script type="text/javascript">var smnAdvertiserId = '00000517';</script>
<script type="text/javascript" src="//cd-ladsp-com.s3.amazonaws.com/script/conv.js"></script>

<script type="text/javascript">var smnAdvertiserId = '00000517';</script>
<script type="text/javascript" src="//cd-ladsp-com.s3.amazonaws.com/script/pixel.js"></script>


<!--/logicad-->

<%

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
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"">���i��UP�}�j���A��</a></li>"
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


End Function
%>
